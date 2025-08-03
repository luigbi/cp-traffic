Attribute VB_Name = "EngrRecGet"
'
' Release: 1.0
'
' Description:
'   This file contains the Get Record Modules

Option Explicit
Type LISTINFO
    lCode As Long
    sCurrent As String * 1
    lOrigCode As Long
    iVersion As Integer
    sState As String * 1
End Type
Public tgListInfo() As LISTINFO


Public Function gGetRec_SEE_ScheduleEvent(llCode As Long, slForm_Module As String, tlSEE As SEE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM SEE_Schedule_Events WHERE seeCode = " & llCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlSEE.lCode = rst!seeCode
        tlSEE.lSheCode = rst!seeshecode
        tlSEE.sAction = rst!seeAction
        tlSEE.lDeeCode = rst!seeDeeCode
        tlSEE.iBdeCode = rst!seeBdeCode
        tlSEE.iBusCceCode = rst!seeBusCceCode
        tlSEE.sSchdType = rst!seeSchdType
        tlSEE.iEteCode = rst!seeEteCode
        tlSEE.lTime = rst!seeTime
        tlSEE.iStartTteCode = rst!seeStartTteCode
        tlSEE.sFixedTime = rst!seeFixedTime
        tlSEE.iEndTteCode = rst!seeEndTteCode
        tlSEE.lDuration = rst!seeDuration
        tlSEE.iMteCode = rst!seeMteCode
        tlSEE.iAudioAseCode = rst!seeAudioAseCode
        tlSEE.sAudioItemID = rst!seeAudioItemID
        tlSEE.sAudioItemIDChk = rst!seeAudioItemIDChk
        tlSEE.sAudioISCI = rst!seeAudioISCI
        tlSEE.iAudioCceCode = rst!seeAudioCceCode
        tlSEE.iBkupAneCode = rst!seeBkupAneCode
        tlSEE.iBkupCceCode = rst!seeBkupCceCode
        tlSEE.iProtAneCode = rst!seeProtAneCode
        tlSEE.sProtItemID = rst!seeProtItemID
        tlSEE.sProtItemIDChk = rst!seeProtItemIDChk
        tlSEE.sProtISCI = rst!seeProtISCI
        tlSEE.iProtCceCode = rst!seeProtCceCode
        tlSEE.i1RneCode = rst!see1RneCode
        tlSEE.i2RneCode = rst!see2RneCode
        tlSEE.iFneCode = rst!seeFneCode
        tlSEE.lSilenceTime = rst!seeSilenceTime
        tlSEE.i1SceCode = rst!see1SceCode
        tlSEE.i2SceCode = rst!see2SceCode
        tlSEE.i3SceCode = rst!see3SceCode
        tlSEE.i4SceCode = rst!see4SceCode
        tlSEE.iStartNneCode = rst!seeStartNneCode
        tlSEE.iEndNneCode = rst!seeEndNneCode
        tlSEE.l1CteCode = rst!see1CteCode
        tlSEE.l2CteCode = rst!see2CteCode
        tlSEE.lAreCode = rst!seeAreCode
        tlSEE.lSpotTime = rst!seeSpotTime
        tlSEE.lEventID = rst!seeEventID
        tlSEE.sAsAirStatus = rst!seeAsAirStatus
        tlSEE.sSentStatus = rst!seeSentStatus
        tlSEE.sSentDate = Format$(rst!seeSentDate, sgShowDateForm)
        tlSEE.sIgnoreConflicts = rst!seeIgnoreConflicts
        tlSEE.lDheCode = rst!seeDheCode
        tlSEE.lOrigDHECode = rst!seeOrigDHECode
        tlSEE.sInsertFlag = "N"     'Temporary flag used in Schedule Definition
        tlSEE.sABCFormat = rst!seeABCFormat
        tlSEE.sABCPgmCode = rst!seeABCPgmCode
        tlSEE.sABCXDSMode = rst!seeABCXDSMode
        tlSEE.sABCRecordItem = rst!seeABCRecordItem
        tlSEE.sUnused = ""
        'Extra field not part of record
        tlSEE.lAvailLength = tlSEE.lDuration
        rst.Close
        gGetRec_SEE_ScheduleEvent = True
        Exit Function
    Else
        rst.Close
        gGetRec_SEE_ScheduleEvent = False
        Exit Function
    End If
    
ErrHand:
    rst.Close
    gShowErrorMsg slForm_Module
    gGetRec_SEE_ScheduleEvent = False
    Exit Function
End Function
Public Function gGetRec_EPE_EventProperties(ilCode As Integer, slForm_Module As String, tlEPE As EPE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM EPE_Event_Properties WHERE EPECode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlEPE.iCode = rst!epeCode
        tlEPE.iEteCode = rst!epeEteCode
        tlEPE.sType = rst!epeType
        tlEPE.sBus = rst!epeBus
        tlEPE.sBusControl = rst!epeBusControl
        tlEPE.sTime = rst!epeTime
        tlEPE.sStartType = rst!epeStartType
        tlEPE.sFixedTime = rst!epeFixedTime
        tlEPE.sEndType = rst!epeEndType
        tlEPE.sDuration = rst!epeDuration
        tlEPE.sMaterialType = rst!epeMaterialType
        tlEPE.sAudioName = rst!epeAudioName
        tlEPE.sAudioItemID = rst!epeAudioItemID
        tlEPE.sAudioISCI = rst!epeAudioISCI
        tlEPE.sAudioControl = rst!epeAudioControl
        tlEPE.sBkupAudioName = rst!epeBkupAudioName
        tlEPE.sBkupAudioControl = rst!epeBkupAudioControl
        tlEPE.sProtAudioName = rst!epeProtAudioName
        tlEPE.sProtAudioItemID = rst!epeProtAudioItemID
        tlEPE.sProtAudioISCI = rst!epeProtAudioISCI
        tlEPE.sProtAudioControl = rst!epeProtAudioControl
        tlEPE.sRelay1 = rst!epeRelay1
        tlEPE.sRelay2 = rst!epeRelay2
        tlEPE.sFollow = rst!epeFollow
        tlEPE.sSilenceTime = rst!epeSilenceTime
        tlEPE.sSilence1 = rst!epeSilence1
        tlEPE.sSilence2 = rst!epeSilence2
        tlEPE.sSilence3 = rst!epeSilence3
        tlEPE.sSilence4 = rst!epeSilence4
        tlEPE.sStartNetcue = rst!epeStartNetcue
        tlEPE.sStopNetcue = rst!epeStopNetcue
        tlEPE.sTitle1 = rst!epeTitle1
        tlEPE.sTitle2 = rst!epeTitle2
        tlEPE.sABCFormat = rst!epeABCFormat
        tlEPE.sABCPgmCode = rst!epeABCPgmCode
        tlEPE.sABCXDSMode = rst!epeABCXDSMode
        tlEPE.sABCRecordItem = rst!epeABCRecordItem
        tlEPE.sUnused = rst!epeUnused
        rst.Close
        gGetRec_EPE_EventProperties = True
        Exit Function
    Else
        rst.Close
        gGetRec_EPE_EventProperties = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_EPE_EventProperties = False
    Exit Function
End Function
Public Function gGetRec_APE_AutoPath(ilCode As Integer, slForm_Module As String, tlAPE As APE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM APE_Auto_Path WHERE APECode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlAPE.iCode = rst!apeCode
        tlAPE.iAeeCode = rst!apeAeeCode
        tlAPE.sType = rst!apeType
        tlAPE.sSubType = rst!apeSubType
        If (tlAPE.sSubType <> "P") And (tlAPE.sSubType <> "T") Then
            tlAPE.sSubType = "P"
        End If
        tlAPE.sNewFileName = rst!apeNewFileName
        tlAPE.sChgFileName = rst!apeChgFileName
        tlAPE.sDelFileName = rst!apeDelFileName
        tlAPE.sNewFileExt = rst!apeNewFileExt
        tlAPE.sChgFileExt = rst!apeChgFileExt
        tlAPE.sDelFileExt = rst!apeDelFileExt
        tlAPE.sPath = rst!apePath
        tlAPE.sDateFormat = rst!apeDateFormat
        tlAPE.sTimeFormat = rst!apeTimeFormat
        tlAPE.sUnused = rst!apeUnused
        rst.Close
        gGetRec_APE_AutoPath = True
        Exit Function
    Else
        rst.Close
        gGetRec_APE_AutoPath = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_APE_AutoPath = False
    Exit Function
End Function

Public Function gGetRec_AAE_As_Aired(llCode As Long, slForm_Module As String, tlAAE As AAE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM AAE_As_Aired WHERE aaeCode = " & llCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlAAE.lCode = rst!aaeCode
        tlAAE.lSheCode = rst!aaeSheCode
        tlAAE.lSeeCode = rst!aaeSeeCode
        tlAAE.sAirDate = Format$(rst!aaeAirDate, sgShowDateForm)
        tlAAE.lAirTime = rst!aaeAirTime
        tlAAE.sAutoOff = rst!aaeAutoOff
        tlAAE.sData = rst!aaeData
        tlAAE.sSchedule = rst!aaeSchedule
        tlAAE.sTrueTime = rst!aaeTrueTime
        tlAAE.sSourceConflict = rst!aaeSourceConflict
        tlAAE.sSourceUnavail = rst!aaeSourceUnavail
        tlAAE.sSourceItem = rst!aaeSourceItem
        tlAAE.sBkupSrceUnavail = rst!aaeBkupSrceUnavail
        tlAAE.sBkupSrceItem = rst!aaeBkupSrceItem
        tlAAE.sProtSrceUnavail = rst!aaeProtSrceUnavail
        tlAAE.sProtSrceItem = rst!aaeProtSrceItem
        tlAAE.sDate = rst!aaeDate
        tlAAE.lEventID = rst!aaeEventID
        tlAAE.sBusName = rst!aaeBusName
        tlAAE.sBusControl = rst!aaeBusControl
        tlAAE.sEventType = rst!aaeEventType
        tlAAE.sStartTime = rst!aaeStartTime
        tlAAE.sStartType = rst!aaeStartType
        tlAAE.sFixedTime = rst!aaeFixedTime
        tlAAE.sEndType = rst!aaeEndType
        tlAAE.sDuration = rst!aaeDuration
        tlAAE.sOutTime = rst!aaeOutTime
        tlAAE.sMaterialType = rst!aaeMaterialType
        tlAAE.sAudioName = rst!aaeAudioName
        tlAAE.sAudioItemID = rst!aaeAudioItemID
        tlAAE.sAudioISCI = rst!aaeAudioISCI
        tlAAE.sAudioCrtlChar = rst!aaeAudioCrtlChar
        tlAAE.sBkupAudioName = rst!aaeBkupAudioName
        tlAAE.sBkupCtrlChar = rst!aaeBkupCtrlChar
        tlAAE.sProtAudioName = rst!aaeProtAudioName
        tlAAE.sProtItemID = rst!aaeProtItemID
        tlAAE.sProtISCI = rst!aaeProtISCI
        tlAAE.sProtCtrlChar = rst!aaeProtCtrlChar
        tlAAE.sRelay1 = rst!aaeRelay1
        tlAAE.sRelay2 = rst!aaeRelay2
        tlAAE.sFollow = rst!aaeFollow
        tlAAE.sSilenceTime = rst!aaeSilenceTime
        tlAAE.sSilence1 = rst!aaeSilence1
        tlAAE.sSilence2 = rst!aaeSilence2
        tlAAE.sSilence3 = rst!aaeSilence3
        tlAAE.sSilence4 = rst!aaeSilence4
        tlAAE.sNetcueStart = rst!aaeNetcueStart
        tlAAE.sNetcueEnd = rst!aaeNetcueEnd
        tlAAE.sTitle1 = rst!aaeTitle1
        tlAAE.sTitle2 = rst!aaeTitle2
        tlAAE.sTitle2 = rst!aaeABCFormat
        tlAAE.sTitle2 = rst!aaeABCPgmCode
        tlAAE.sTitle2 = rst!aaeABCXDSMode
        tlAAE.sTitle2 = rst!aaeABCRecordItem
        tlAAE.sEnteredDate = Format$(rst!aaeEnteredDate, sgShowDateForm)
        tlAAE.sEnteredTime = Format$(rst!aaeEnteredTime, sgShowTimeWSecForm)
        tlAAE.sUnused = rst!aaeUnused
        rst.Close
        gGetRec_AAE_As_Aired = True
        Exit Function
    Else
        rst.Close
        gGetRec_AAE_As_Aired = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_AAE_As_Aired = False
    Exit Function
End Function

Public Function gGetRec_AEE_AutoEquip(ilCode As Integer, slForm_Module As String, tlAEE As AEE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM AEE_Auto_Equip WHERE aeeCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlAEE.iCode = rst!aeeCode
        tlAEE.sName = rst!aeeName
        tlAEE.sDescription = rst!aeeDescription
        tlAEE.sManufacture = rst!aeeManufacture
        tlAEE.sFixedTimeChar = rst!aeeFixedTimeChar
        tlAEE.lAlertSchdDelay = rst!aeeAlertSchdDelay
        tlAEE.sState = rst!aeeState
        tlAEE.sUsedFlag = rst!aeeUsedFlag
        tlAEE.iVersion = rst!aeeVersion
        tlAEE.iOrigAeeCode = rst!aeeOrigAeeCode
        tlAEE.sCurrent = rst!aeeCurrent
        tlAEE.sEnteredDate = Format$(rst!aeeEnteredDate, sgShowDateForm)
        tlAEE.sEnteredTime = Format$(rst!aeeEnteredTime, sgShowTimeWSecForm)
        tlAEE.iUieCode = rst!aeeUieCode
        tlAEE.sUnused = rst!aeeUnused
        rst.Close
        gGetRec_AEE_AutoEquip = True
        Exit Function
    Else
        rst.Close
        gGetRec_AEE_AutoEquip = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_AEE_AutoEquip = False
    Exit Function
End Function
Public Function gGetRec_ANE_AudioName(ilCode As Integer, slForm_Module As String, tlANE As ANE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM ANE_Audio_Name WHERE aneCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlANE.iCode = rst!aneCode
        tlANE.sName = rst!aneName
        tlANE.sDescription = rst!aneDescription
        tlANE.iCceCode = rst!aneCceCode
        tlANE.iAteCode = rst!aneAteCode
        tlANE.sState = rst!aneState
        tlANE.sUsedFlag = rst!aneUsedFlag
        tlANE.iVersion = rst!aneVersion
        tlANE.iOrigAneCode = rst!aneOrigAneCode
        tlANE.sCurrent = rst!aneCurrent
        tlANE.sEnteredDate = Format$(rst!aneEnteredDate, sgShowDateForm)
        tlANE.sEnteredTime = Format$(rst!aneEnteredTime, sgShowTimeWSecForm)
        tlANE.iUieCode = rst!aneUieCode
        tlANE.sCheckConflicts = rst!aneCheckConflicts
        tlANE.sUnused = rst!aneUnused
        rst.Close
        gGetRec_ANE_AudioName = True
        Exit Function
    Else
        rst.Close
        gGetRec_ANE_AudioName = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_ANE_AudioName = False
    Exit Function
End Function

Public Function gGetRec_ARE_AdvertiserRefer(llCode As Long, slForm_Module As String, tlARE As ARE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    If llCode <= 0 Then
        tlARE.lCode = 0
        tlARE.sName = ""
        tlARE.sUnusued = ""
        rst.Close
        gGetRec_ARE_AdvertiserRefer = True
        Exit Function
    End If
    sgSQLQuery = "SELECT * FROM ARE_Advertiser_Refer WHERE areCode = " & llCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlARE.lCode = rst!areCode
        tlARE.sName = rst!areName
        tlARE.sUnusued = rst!areUnusued
        rst.Close
        gGetRec_ARE_AdvertiserRefer = True
        Exit Function
    Else
        rst.Close
        gGetRec_ARE_AdvertiserRefer = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_ARE_AdvertiserRefer = False
    Exit Function
End Function

Public Function gGetRec_ASE_AudioSource(ilCode As Integer, slForm_Module As String, tlASE As ASE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM ASE_Audio_Source WHERE aseCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlASE.iCode = rst!aseCode
        tlASE.iPriAneCode = rst!asePriAneCode
        tlASE.iPriCceCode = rst!asePriCceCode
        tlASE.sDescription = rst!aseDescription
        tlASE.iBkupAneCode = rst!aseBkupAneCode
        tlASE.iBkupCceCode = rst!aseBkupCceCode
        tlASE.iProtAneCode = rst!aseProtAneCode
        tlASE.iProtCceCode = rst!aseProtCceCode
        tlASE.sState = rst!aseState
        tlASE.sUsedFlag = rst!aseUsedFlag
        tlASE.iVersion = rst!aseVersion
        tlASE.iOrigAseCode = rst!aseOrigAseCode
        tlASE.sCurrent = rst!aseCurrent
        tlASE.sEnteredDate = Format$(rst!aseEnteredDate, sgShowDateForm)
        tlASE.sEnteredTime = Format$(rst!aseEnteredTime, sgShowTimeWSecForm)
        tlASE.iUieCode = rst!aseUieCode
        tlASE.sUnused = rst!aseUnused
        rst.Close
        gGetRec_ASE_AudioSource = True
        Exit Function
    Else
        rst.Close
        gGetRec_ASE_AudioSource = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_ASE_AudioSource = False
    Exit Function
End Function


Public Function gGetRec_ATE_AudioType(ilCode As Integer, slForm_Module As String, tlATE As ATE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM ATE_Audio_Type WHERE ateCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlATE.iCode = rst!ateCode
        tlATE.sName = rst!ateName
        tlATE.sDescription = rst!ateDescription
        tlATE.sState = rst!ateState
        tlATE.sTestItemID = rst!ateTestItemID
        tlATE.lPreBufferTime = rst!atePreBufferTime
        tlATE.lPostBufferTime = rst!atePostBufferTime
        tlATE.sUsedFlag = rst!ateUsedFlag
        tlATE.iVersion = rst!ateVersion
        tlATE.iOrigAteCode = rst!ateOrigAteCode
        tlATE.sCurrent = rst!ateCurrent
        tlATE.sEnteredDate = Format$(rst!ateEnteredDate, sgShowDateForm)
        tlATE.sEnteredTime = Format$(rst!ateEnteredTime, sgShowTimeWSecForm)
        tlATE.iUieCode = rst!ateUieCode
        tlATE.sUnused = rst!ateUnused
        rst.Close
        gGetRec_ATE_AudioType = True
        Exit Function
    Else
        rst.Close
        gGetRec_ATE_AudioType = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_ATE_AudioType = False
    Exit Function
End Function


Public Function gGetRec_BDE_BusDefinition(ilCode As Integer, slForm_Module As String, tlBDE As BDE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM BDE_Bus_Definition WHERE bdeCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlBDE.iCode = rst!bdeCode
        tlBDE.sName = rst!bdeName
        tlBDE.sDescription = rst!bdeDescription
        tlBDE.sChannel = rst!bdeChannel
        tlBDE.iAseCode = rst!bdeAseCode
        tlBDE.sState = rst!bdeState
        tlBDE.iCceCode = rst!bdeCceCode
        tlBDE.sUsedFlag = rst!bdeUsedFlag
        tlBDE.iVersion = rst!bdeVersion
        tlBDE.iOrigBdeCode = rst!bdeOrigBdeCode
        tlBDE.sCurrent = rst!bdeCurrent
        tlBDE.sEnteredDate = Format$(rst!bdeEnteredDate, sgShowDateForm)
        tlBDE.sEnteredTime = Format$(rst!bdeEnteredTime, sgShowTimeWSecForm)
        tlBDE.iUieCode = rst!bdeUieCode
        tlBDE.sUnused = rst!bdeUnused
        rst.Close
        gGetRec_BDE_BusDefinition = True
        Exit Function
    Else
        rst.Close
        gGetRec_BDE_BusDefinition = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_BDE_BusDefinition = False
    Exit Function
End Function


Public Function gGetRec_BGE_BusGroup(ilCode As Integer, slForm_Module As String, tlBGE As BGE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM BGE_Bus_Group WHERE bgeCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlBGE.iCode = rst!bgeCode
        tlBGE.sName = rst!bgeName
        tlBGE.sDescription = rst!bgeDescription
        tlBGE.sState = rst!bgeState
        tlBGE.sUsedFlag = rst!bgeUsedFlag
        tlBGE.iVersion = rst!bgeVersion
        tlBGE.iOrigBgeCode = rst!bgeOrigBgeCode
        tlBGE.sCurrent = rst!bgeCurrent
        tlBGE.sEnteredDate = Format$(rst!bgeEnteredDate, sgShowDateForm)
        tlBGE.sEnteredTime = Format$(rst!bgeEnteredTime, sgShowTimeWSecForm)
        tlBGE.iUieCode = rst!bgeUieCode
        tlBGE.sUnused = rst!bgeUnused
        rst.Close
        gGetRec_BGE_BusGroup = True
        Exit Function
    Else
        rst.Close
        gGetRec_BGE_BusGroup = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_BGE_BusGroup = False
    Exit Function
End Function


Public Function gGetRec_CCE_ControlChar(ilCode As Integer, slForm_Module As String, tlCCE As CCE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM CCE_Control_Char WHERE cceCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlCCE.iCode = rst!cceCode
        tlCCE.sType = rst!cceType
        tlCCE.sAutoChar = rst!cceAutoChar
        tlCCE.sDescription = rst!cceDescription
        tlCCE.sState = rst!cceState
        tlCCE.sUsedFlag = rst!cceUsedFlag
        tlCCE.iVersion = rst!cceVersion
        tlCCE.iOrigCceCode = rst!cceOrigCceCode
        tlCCE.sCurrent = rst!cceCurrent
        tlCCE.sEnteredDate = Format$(rst!cceEnteredDate, sgShowDateForm)
        tlCCE.sEnteredTime = Format$(rst!cceEnteredTime, sgShowTimeWSecForm)
        tlCCE.iUieCode = rst!cceUieCode
        tlCCE.sUnused = rst!cceUnused
        rst.Close
        gGetRec_CCE_ControlChar = True
        Exit Function
    Else
        rst.Close
        gGetRec_CCE_ControlChar = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_CCE_ControlChar = False
    Exit Function
End Function

Public Function gGetRec_CTE_CommtsTitle(llCode As Long, slForm_Module As String, tlCTE As CTE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM CTE_Commts_And_Title WHERE cteCode = " & llCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlCTE.lCode = rst!cteCode
        tlCTE.sType = rst!cteType
        tlCTE.sComment = rst!cteComment
        tlCTE.sState = rst!cteState
        tlCTE.sUsedFlag = rst!cteUsedFlag
        tlCTE.iVersion = rst!cteVersion
        tlCTE.lOrigCteCode = rst!cteOrigCteCode
        tlCTE.sCurrent = rst!cteCurrent
        tlCTE.sEnteredDate = Format$(rst!cteEnteredDate, sgShowDateForm)
        tlCTE.sEnteredTime = Format$(rst!cteEnteredTime, sgShowTimeWSecForm)
        tlCTE.iUieCode = rst!cteUieCode
        tlCTE.sUnused = rst!cteUnused
        rst.Close
        gGetRec_CTE_CommtsTitle = True
        Exit Function
    Else
        tlCTE.lCode = 0
        tlCTE.sType = ""
        tlCTE.sComment = ""
        tlCTE.sState = ""
        tlCTE.sUsedFlag = ""
        tlCTE.iVersion = 0
        tlCTE.lOrigCteCode = 0
        tlCTE.sCurrent = ""
        tlCTE.sEnteredDate = ""
        tlCTE.sEnteredTime = ""
        tlCTE.iUieCode = 0
        tlCTE.sUnused = ""
        rst.Close
        gGetRec_CTE_CommtsTitle = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_CTE_CommtsTitle = False
    Exit Function
End Function

Public Function gGetRec_CTE_CommtsTitleAPI(hlCTE As Integer, llCode As Long, slForm_Module As String, tlCTE As CTE) As Integer
    Dim tlCTESrchKey As LONGKEY0
    Dim ilCTERecLen As Integer
    Dim ilRet As Integer
    Dim tlCTEAPI As CTEAPI
    
    On Error GoTo ErrHand
    ilCTERecLen = Len(tlCTEAPI)
    tlCTESrchKey.lCode = llCode
    ilRet = btrGetEqual(hlCTE, tlCTEAPI, ilCTERecLen, tlCTESrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        tlCTE.lCode = tlCTEAPI.lCode
        tlCTE.sType = tlCTEAPI.sType
        tlCTE.sComment = tlCTEAPI.sComment
        tlCTE.sState = tlCTEAPI.sState
        tlCTE.sUsedFlag = tlCTEAPI.sUsedFlag
        tlCTE.iVersion = tlCTEAPI.iVersion
        tlCTE.lOrigCteCode = tlCTEAPI.lOrigCteCode
        tlCTE.sCurrent = tlCTEAPI.sCurrent
        gUnpackDate tlCTEAPI.iEneteredDate(0), tlCTEAPI.iEneteredDate(1), tlCTE.sEnteredDate
        tlCTE.sEnteredDate = Format$(gAdjYear(tlCTE.sEnteredDate), sgShowDateForm)
        gUnpackTime tlCTEAPI.iEnteredTime(0), tlCTEAPI.iEnteredTime(1), "A", "1", tlCTE.sEnteredTime
        tlCTE.sEnteredTime = Format$(tlCTE.sEnteredTime, sgShowTimeWSecForm)
        'tlCTE.sEnteredTime = ""
        tlCTE.iUieCode = tlCTEAPI.iUieCode
        tlCTE.sUnused = tlCTEAPI.sUnused
        gGetRec_CTE_CommtsTitleAPI = True
        Exit Function
    Else
        tlCTE.lCode = 0
        tlCTE.sType = ""
        tlCTE.sComment = ""
        tlCTE.sState = ""
        tlCTE.sUsedFlag = ""
        tlCTE.iVersion = 0
        tlCTE.lOrigCteCode = 0
        tlCTE.sCurrent = ""
        tlCTE.sEnteredDate = ""
        tlCTE.sEnteredTime = ""
        tlCTE.iUieCode = 0
        tlCTE.sUnused = ""
        gGetRec_CTE_CommtsTitleAPI = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    gGetRec_CTE_CommtsTitleAPI = False
    Exit Function
End Function

Public Function gGetRec_DEE_DayEvent(llDeeCode As Long, slForm_Module As String, tlDEE As DEE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM DEE_Day_Event_Info Where deeCode = " & llDeeCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlDEE.lCode = rst!deeCode
        tlDEE.lDheCode = rst!deeDheCode
        tlDEE.iCceCode = rst!deeCceCode
        tlDEE.iEteCode = rst!deeEteCode
        tlDEE.lTime = rst!deeTime
        tlDEE.iStartTteCode = rst!deeStartTteCode
        tlDEE.sFixedTime = rst!deeFixedTime
        tlDEE.iEndTteCode = rst!deeEndTteCode
        tlDEE.lDuration = rst!deeDuration
        tlDEE.sHours = rst!deeHours
        tlDEE.sDays = rst!deeDays
        tlDEE.iMteCode = rst!deeMteCode
        tlDEE.iAudioAseCode = rst!deeAudioAseCode
        tlDEE.sAudioItemID = rst!deeAudioItemID
        tlDEE.sAudioISCI = rst!deeAudioISCI
        tlDEE.iAudioCceCode = rst!deeAudioCceCode
        tlDEE.iBkupAneCode = rst!deeBkupAneCode
        tlDEE.iBkupCceCode = rst!deeBkupCceCode
        tlDEE.iProtAneCode = rst!deeProtAneCode
        tlDEE.sProtItemID = rst!deeProtItemID
        tlDEE.sProtISCI = rst!deeProtISCI
        tlDEE.iProtCceCode = rst!deeProtCceCode
        tlDEE.i1RneCode = rst!dee1RneCode
        tlDEE.i2RneCode = rst!dee2RneCode
        tlDEE.iFneCode = rst!deeFneCode
        tlDEE.lSilenceTime = rst!deeSilenceTime
        tlDEE.i1SceCode = rst!dee1SceCode
        tlDEE.i2SceCode = rst!dee2SceCode
        tlDEE.i3SceCode = rst!dee3SceCode
        tlDEE.i4SceCode = rst!dee4SceCode
        tlDEE.iStartNneCode = rst!deeStartNneCode
        tlDEE.iEndNneCode = rst!deeEndNneCode
        tlDEE.l1CteCode = rst!dee1CteCode
        tlDEE.l2CteCode = rst!dee2CteCode
        tlDEE.lEventID = rst!deeEventID
        tlDEE.sIgnoreConflicts = rst!deeIgnoreConflicts
        tlDEE.sABCFormat = rst!deeABCFormat
        tlDEE.sABCPgmCode = rst!deeABCPgmCode
        tlDEE.sABCXDSMode = rst!deeABCXDSMode
        tlDEE.sABCRecordItem = rst!deeABCRecordItem
        tlDEE.sUnused = rst!deeUnused
        rst.Close
        gGetRec_DEE_DayEvent = True
        Exit Function
    Else
        rst.Close
        gGetRec_DEE_DayEvent = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_DEE_DayEvent = False
    Exit Function

End Function

Public Function gGetRec_DHE_DayHeaderInfo(llCode As Long, slForm_Module As String, tlDHE As DHE) As Integer
    Dim tlDHESrchKey As LONGKEY0
    Dim ilDHERecLen As Integer
    Dim ilRet As Integer
    Dim tlDHEAPI As DHEAPI
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info WHERE dheCode = " & llCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlDHE.lCode = rst!dheCode
        tlDHE.sType = rst!dheType
        tlDHE.lDneCode = rst!dheDneCode
        tlDHE.lDseCode = rst!dheDseCode
        tlDHE.sStartTime = Format$(rst!dheStartTime, sgShowTimeWSecForm)
        tlDHE.lLength = rst!dheLength
        tlDHE.sHours = rst!dheHours
        tlDHE.sStartDate = Format$(rst!dheStartDate, sgShowDateForm)
        tlDHE.sEndDate = Format$(rst!dheEndDate, sgShowDateForm)
        tlDHE.sDays = rst!dheDays
        tlDHE.lCteCode = rst!dheCteCode
        tlDHE.sState = rst!dheState
        tlDHE.sUsedFlag = rst!dheUsedFlag
        tlDHE.iVersion = rst!dheVersion
        tlDHE.lOrigDHECode = rst!dheOrigDheCode
        tlDHE.sCurrent = rst!dheCurrent
        tlDHE.sEnteredDate = Format$(rst!dheEnteredDate, sgShowDateForm)
        tlDHE.sEnteredTime = Format$(rst!dheEnteredTime, sgShowTimeWSecForm)
        tlDHE.iUieCode = rst!dheUieCode
        tlDHE.sIgnoreConflicts = rst!dheIgnoreConflicts
        tlDHE.sBusNames = rst!dheBusNames
        tlDHE.sUnused = rst!dheUnused
        rst.Close
        gGetRec_DHE_DayHeaderInfo = True
        Exit Function
    Else
        rst.Close
        gGetRec_DHE_DayHeaderInfo = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_DHE_DayHeaderInfo = False
    Exit Function
End Function

Public Function gGetRec_DHE_DayHeaderInfoAPI(hlDHE As Integer, llCode As Long, slForm_Module As String, tlDHE As DHE) As Integer
    Dim tlDHESrchKey As LONGKEY0
    Dim ilDHERecLen As Integer
    Dim ilRet As Integer
    Dim tlDHEAPI As DHEAPI
    
    On Error GoTo ErrHand
    ilDHERecLen = Len(tlDHEAPI)
    tlDHESrchKey.lCode = llCode
    ilRet = btrGetEqual(hlDHE, tlDHEAPI, ilDHERecLen, tlDHESrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        tlDHE.lCode = tlDHEAPI.lCode
        tlDHE.sType = tlDHEAPI.sType
        tlDHE.lDneCode = tlDHEAPI.lDneCode
        tlDHE.lDseCode = tlDHEAPI.lDseCode
        gUnpackTime tlDHEAPI.iStartTime(0), tlDHEAPI.iStartTime(1), "A", "1", tlDHE.sStartTime
        tlDHE.sStartTime = Format$(tlDHE.sStartTime, sgShowTimeWSecForm)
        tlDHE.lLength = tlDHEAPI.lLength
        tlDHE.sHours = tlDHEAPI.sHours
        gUnpackDate tlDHEAPI.iStartDate(0), tlDHEAPI.iStartDate(1), tlDHE.sStartDate
        tlDHE.sStartDate = Format$(gAdjYear(tlDHE.sStartDate), sgShowDateForm)
        gUnpackDate tlDHEAPI.iEndDate(0), tlDHEAPI.iEndDate(1), tlDHE.sEndDate
        tlDHE.sEndDate = Format$(gAdjYear(tlDHE.sEndDate), sgShowDateForm)
        tlDHE.sDays = tlDHEAPI.sDays
        tlDHE.lCteCode = tlDHEAPI.lCteCode
        tlDHE.sState = tlDHEAPI.sState
        tlDHE.sUsedFlag = tlDHEAPI.sUsedFlag
        tlDHE.iVersion = tlDHEAPI.iVersion
        tlDHE.lOrigDHECode = tlDHEAPI.lOrigDHECode
        tlDHE.sCurrent = tlDHEAPI.sCurrent
        gUnpackDate tlDHEAPI.iEnteredDate(0), tlDHEAPI.iEnteredDate(1), tlDHE.sEnteredDate
        tlDHE.sEnteredDate = Format$(gAdjYear(tlDHE.sEnteredDate), sgShowDateForm)
        gUnpackTime tlDHEAPI.iEnteredTime(0), tlDHEAPI.iEnteredTime(1), "A", "1", tlDHE.sEnteredTime
        tlDHE.sEnteredTime = Format$(tlDHE.sEnteredTime, sgShowTimeWSecForm)
        tlDHE.iUieCode = tlDHEAPI.iUieCode
        tlDHE.sIgnoreConflicts = tlDHEAPI.sIgnoreConflicts
        tlDHE.sBusNames = tlDHEAPI.sBusNames
        tlDHE.sUnused = tlDHEAPI.sUnused
        gGetRec_DHE_DayHeaderInfoAPI = True
        Exit Function
    Else
        gGetRec_DHE_DayHeaderInfoAPI = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    'rst.Close
    gGetRec_DHE_DayHeaderInfoAPI = False
    Exit Function
End Function

Public Function gGetRec_DNE_DayName(llCode As Long, slForm_Module As String, tlDNE As DNE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM DNE_Day_Name WHERE dneCode = " & llCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlDNE.lCode = rst!dneCode
        tlDNE.sType = rst!dneType
        tlDNE.sName = rst!dneName
        tlDNE.sDescription = rst!dneDescription
        tlDNE.sState = rst!dneState
        tlDNE.sUsedFlag = rst!dneUsedFlag
        tlDNE.iVersion = rst!dneVersion
        tlDNE.lOrigDneCode = rst!dneOrigDneCode
        tlDNE.sCurrent = rst!dneCurrent
        tlDNE.sEnteredDate = Format$(rst!dneEnteredDate, sgShowDateForm)
        tlDNE.sEnteredTime = Format$(rst!dneEnteredTime, sgShowTimeWSecForm)
        tlDNE.iUieCode = rst!dneUieCode
        tlDNE.sUnused = rst!dneUnused
        rst.Close
        gGetRec_DNE_DayName = True
        Exit Function
    Else
        rst.Close
        gGetRec_DNE_DayName = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_DNE_DayName = False
    Exit Function
End Function

Public Function gGetRec_DSE_DaySubName(llCode As Long, slForm_Module As String, tlDSE As DSE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM DSE_Day_SubName WHERE dseCode = " & llCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlDSE.lCode = rst!dseCode
        tlDSE.sName = rst!dseName
        tlDSE.sDescription = rst!dseDescription
        tlDSE.sState = rst!dseState
        tlDSE.sUsedFlag = rst!dseUsedFlag
        tlDSE.iVersion = rst!dseVersion
        tlDSE.lOrigDseCode = rst!dseOrigDseCode
        tlDSE.sCurrent = rst!dseCurrent
        tlDSE.sEnteredDate = Format$(rst!dseEnteredDate, sgShowDateForm)
        tlDSE.sEnteredTime = Format$(rst!dseEnteredTime, sgShowTimeWSecForm)
        tlDSE.iUieCode = rst!dseUieCode
        tlDSE.sUnused = rst!dseUnused
        rst.Close
        gGetRec_DSE_DaySubName = True
        Exit Function
    Else
        rst.Close
        gGetRec_DSE_DaySubName = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_DSE_DaySubName = False
    Exit Function
End Function

Public Function gGetRec_ETE_EventType(ilCode As Integer, slForm_Module As String, tlETE As ETE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM ETE_Event_Type WHERE eteCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlETE.iCode = rst!eteCode
        tlETE.sName = rst!eteName
        tlETE.sDescription = rst!eteDescription
        tlETE.sCategory = rst!eteCategory
        tlETE.sAutoCodeChar = rst!eteAutoCodeChar
        tlETE.sState = rst!eteState
        tlETE.sUsedFlag = rst!eteUsedFlag
        tlETE.iVersion = rst!eteVersion
        tlETE.iOrigEteCode = rst!eteOrigEteCode
        tlETE.sCurrent = rst!eteCurrent
        tlETE.sEnteredDate = Format$(rst!eteEnteredDate, sgShowDateForm)
        tlETE.sEnteredTime = Format$(rst!eteEnteredTime, sgShowTimeWSecForm)
        tlETE.iUieCode = rst!eteUieCode
        tlETE.sUnused = rst!eteUnused
        rst.Close
        gGetRec_ETE_EventType = True
        Exit Function
    Else
        rst.Close
        gGetRec_ETE_EventType = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_ETE_EventType = False
    Exit Function
End Function

Public Function gGetRec_FNE_FollowName(ilCode As Integer, slForm_Module As String, tlFNE As FNE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM FNE_Follow_Name WHERE fneCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlFNE.iCode = rst!fneCode
        tlFNE.sName = rst!fneName
        tlFNE.sDescription = rst!fneDescription
        tlFNE.sState = rst!fneState
        tlFNE.sUsedFlag = rst!fneUsedFlag
        tlFNE.iVersion = rst!fneVersion
        tlFNE.iOrigFneCode = rst!fneOrigFneCode
        tlFNE.sCurrent = rst!fneCurrent
        tlFNE.sEnteredDate = Format$(rst!fneEnteredDate, sgShowDateForm)
        tlFNE.sEnteredTime = Format$(rst!fneEnteredTime, sgShowTimeWSecForm)
        tlFNE.iUieCode = rst!fneUieCode
        tlFNE.sUnused = rst!fneUnused
        rst.Close
        gGetRec_FNE_FollowName = True
        Exit Function
    Else
        rst.Close
        gGetRec_FNE_FollowName = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_FNE_FollowName = False
    Exit Function
End Function

Public Function gGetRec_MTE_MaterialType(ilCode As Integer, slForm_Module As String, tlMTE As MTE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM MTE_Material_Type WHERE mteCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlMTE.iCode = rst!mteCode
        tlMTE.sName = rst!mteName
        tlMTE.sDescription = rst!mteDescription
        tlMTE.sState = rst!mteState
        tlMTE.sUsedFlag = rst!mteUsedFlag
        tlMTE.iVersion = rst!mteVersion
        tlMTE.iOrigMteCode = rst!mteOrigmteCode
        tlMTE.sCurrent = rst!mteCurrent
        tlMTE.sEnteredDate = Format$(rst!mteEnteredDate, sgShowDateForm)
        tlMTE.sEnteredTime = Format$(rst!mteEnteredTime, sgShowTimeWSecForm)
        tlMTE.iUieCode = rst!mteUieCode
        tlMTE.sUnused = rst!mteUnused
        rst.Close
        gGetRec_MTE_MaterialType = True
        Exit Function
    Else
        rst.Close
        gGetRec_MTE_MaterialType = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_MTE_MaterialType = False
    Exit Function
End Function

Public Function gGetRec_NNE_NetcueName(ilCode As Integer, slForm_Module As String, tlNNE As NNE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM NNE_Netcue_Name WHERE nneCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlNNE.iCode = rst!nneCode
        tlNNE.sName = rst!nneName
        tlNNE.sDescription = rst!nneDescription
        tlNNE.lDneCode = rst!nneDneCode
        tlNNE.sState = rst!nneState
        tlNNE.sUsedFlag = rst!nneUsedFlag
        tlNNE.iVersion = rst!nneVersion
        tlNNE.iOrigNneCode = rst!nneOrigNneCode
        tlNNE.sCurrent = rst!nneCurrent
        tlNNE.sEnteredDate = Format$(rst!nneEnteredDate, sgShowDateForm)
        tlNNE.sEnteredTime = Format$(rst!nneEnteredTime, sgShowTimeWSecForm)
        tlNNE.iUieCode = rst!nneUieCode
        tlNNE.sUnused = rst!nneUnused
        rst.Close
        gGetRec_NNE_NetcueName = True
        Exit Function
    Else
        rst.Close
        gGetRec_NNE_NetcueName = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_NNE_NetcueName = False
    Exit Function
End Function

Public Function gGetRec_RNE_RelayName(ilCode As Integer, slForm_Module As String, tlRNE As RNE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM RNE_Relay_Name WHERE rneCode = " & ilCode
    'Set rst = cnn.Execute(sgSQLQuery)
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlRNE.iCode = rst!rneCode
        tlRNE.sName = rst!rneName
        tlRNE.sDescription = rst!rneDescription
        tlRNE.sState = rst!rneState
        tlRNE.sUsedFlag = rst!rneUsedFlag
        tlRNE.iVersion = rst!rneVersion
        tlRNE.iOrigRneCode = rst!rneOrigRneCode
        tlRNE.sCurrent = rst!rneCurrent
        tlRNE.sEnteredDate = Format$(rst!rneEnteredDate, sgShowDateForm)
        tlRNE.sEnteredTime = Format$(rst!rneEnteredTime, sgShowTimeWSecForm)
        tlRNE.iUieCode = rst!rneUieCode
        tlRNE.sUnused = rst!rneUnused
        rst.Close
        gGetRec_RNE_RelayName = True
        Exit Function
    Else
        rst.Close
        gGetRec_RNE_RelayName = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_RNE_RelayName = False
    Exit Function
End Function

Public Function gGetRec_SCE_SilenceChar(ilCode As Integer, slForm_Module As String, tlSCE As SCE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM SCE_Silence_Char WHERE sceCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlSCE.iCode = rst!sceCode
        tlSCE.sAutoChar = rst!sceAutoChar
        tlSCE.sDescription = rst!sceDescription
        tlSCE.sState = rst!sceState
        tlSCE.sUsedFlag = rst!sceUsedFlag
        tlSCE.iVersion = rst!sceVersion
        tlSCE.iOrigSceCode = rst!sceOrigSceCode
        tlSCE.sCurrent = rst!sceCurrent
        tlSCE.sEnteredDate = Format$(rst!sceEnteredDate, sgShowDateForm)
        tlSCE.sEnteredTime = Format$(rst!sceEnteredTime, sgShowTimeWSecForm)
        tlSCE.iUieCode = rst!sceUieCode
        tlSCE.sUnused = rst!sceUnused
        rst.Close
        gGetRec_SCE_SilenceChar = True
        Exit Function
    Else
        rst.Close
        gGetRec_SCE_SilenceChar = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_SCE_SilenceChar = False
    Exit Function
End Function

Public Function gGetRec_SHE_ScheduleHeader(llCode As Long, slForm_Module As String, tlSHE As SHE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM SHE_Schedule_Header WHERE sheCode = " & llCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlSHE.lCode = rst!sheCode
        tlSHE.iAeeCode = rst!sheAeeCode
        tlSHE.sAirDate = Format$(rst!sheAirDate, sgShowDateForm)
        tlSHE.sLoadedAutoStatus = rst!sheLoadedAutoStatus
        tlSHE.sLoadedAutoDate = Format$(rst!sheLoadedAutoDate, sgShowDateForm)
        tlSHE.iChgSeqNo = rst!sheChgSeqNo
        tlSHE.sAsAirStatus = rst!sheAsAirStatus
        tlSHE.sLoadedAsAirDate = Format$(rst!sheLoadedAsAirDate, sgShowDateForm)
        tlSHE.sLastDateItemChk = Format$(rst!sheLastDateItemChk, sgShowDateForm)
        tlSHE.sCreateLoad = rst!sheCreateLoad
        tlSHE.iVersion = rst!sheVersion
        tlSHE.lOrigSheCode = rst!sheOrigSheCode
        tlSHE.sCurrent = rst!sheCurrent
        tlSHE.sEnteredDate = Format$(rst!sheEnteredDate, sgShowDateForm)
        tlSHE.sEnteredTime = Format$(rst!sheEnteredTime, sgShowTimeWSecForm)
        tlSHE.iUieCode = rst!sheUieCode
        tlSHE.sConflictExist = rst!sheConflictExist
        tlSHE.sSpotMergeStatus = rst!sheSpotMergeStatus
        tlSHE.sLoadStatus = rst!sheLoadStatus
        tlSHE.sUnused = rst!sheUnused
        rst.Close
        gGetRec_SHE_ScheduleHeader = True
        Exit Function
    Else
        rst.Close
        gGetRec_SHE_ScheduleHeader = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_SHE_ScheduleHeader = False
    Exit Function
End Function


Public Function gGetRec_SHE_ScheduleHeaderByDate(slDate As String, slForm_Module As String, tlSHE As SHE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM SHE_Schedule_Header WHERE sheCurrent = 'Y' AND sheAirDate = '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "'"
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlSHE.lCode = rst!sheCode
        tlSHE.iAeeCode = rst!sheAeeCode
        tlSHE.sAirDate = Format$(rst!sheAirDate, sgShowDateForm)
        tlSHE.sLoadedAutoStatus = rst!sheLoadedAutoStatus
        tlSHE.sLoadedAutoDate = Format$(rst!sheLoadedAutoDate, sgShowDateForm)
        tlSHE.iChgSeqNo = rst!sheChgSeqNo
        tlSHE.sAsAirStatus = rst!sheAsAirStatus
        tlSHE.sLoadedAsAirDate = Format$(rst!sheLoadedAsAirDate, sgShowDateForm)
        tlSHE.sLastDateItemChk = Format$(rst!sheLastDateItemChk, sgShowDateForm)
        tlSHE.sCreateLoad = rst!sheCreateLoad
        tlSHE.iVersion = rst!sheVersion
        tlSHE.lOrigSheCode = rst!sheOrigSheCode
        tlSHE.sCurrent = rst!sheCurrent
        tlSHE.sEnteredDate = Format$(rst!sheEnteredDate, sgShowDateForm)
        tlSHE.sEnteredTime = Format$(rst!sheEnteredTime, sgShowTimeWSecForm)
        tlSHE.iUieCode = rst!sheUieCode
        tlSHE.sConflictExist = rst!sheConflictExist
        tlSHE.sSpotMergeStatus = rst!sheSpotMergeStatus
        tlSHE.sLoadStatus = rst!sheLoadStatus
        tlSHE.sUnused = rst!sheUnused
        rst.Close
        gGetRec_SHE_ScheduleHeaderByDate = True
        Exit Function
    Else
        rst.Close
        gGetRec_SHE_ScheduleHeaderByDate = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_SHE_ScheduleHeaderByDate = False
    Exit Function
End Function

Public Function gGetRec_TSE_TemplateSchd(llCode As Long, slForm_Module As String, tlTSE As TSE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM TSE_Template_Schd WHERE tseCode = " & llCode
    'Set rst = cnn.Execute(sgSQLQuery)
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlTSE.lCode = rst!tseCode
        tlTSE.lDheCode = rst!tseDheCode
        tlTSE.iBdeCode = rst!tseBdeCode
        tlTSE.sLogDate = Format$(rst!tseLogDate, sgShowDateForm)
        tlTSE.sStartTime = Format$(rst!tseStartTime, sgShowTimeWSecForm)
        tlTSE.sDescription = rst!tseDescription
        tlTSE.sState = rst!tseState
        tlTSE.lCteCode = rst!tseCteCode
        tlTSE.iVersion = rst!tseVersion
        tlTSE.lOrigTseCode = rst!tseOrigTseCode
        tlTSE.sCurrent = rst!tseCurrent
        tlTSE.sEnteredDate = Format$(rst!tseEnteredDate, sgShowDateForm)
        tlTSE.sEnteredTime = Format$(rst!tseEnteredTime, sgShowTimeWSecForm)
        tlTSE.iUieCode = rst!tseUieCode
        tlTSE.sUnused = rst!tseUnused
        rst.Close
        gGetRec_TSE_TemplateSchd = True
        Exit Function
    Else
        rst.Close
        gGetRec_TSE_TemplateSchd = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_TSE_TemplateSchd = False
    Exit Function
End Function


Public Function gGetRec_TSE_TemplateSchdByDHETSE(llDheCode As Long, llTSEOrigCode As Long, slForm_Module As String, tlTSE As TSE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM TSE_Template_Schd WHERE tseDHECode = " & llDheCode & " AND tseOrigTseCode = " & llTSEOrigCode
    'Set rst = cnn.Execute(sgSQLQuery)
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlTSE.iVersion = -1
        While Not rst.EOF
            If rst!tseVersion > tlTSE.iVersion Then
                tlTSE.lCode = rst!tseCode
                tlTSE.lDheCode = rst!tseDheCode
                tlTSE.iBdeCode = rst!tseBdeCode
                tlTSE.sLogDate = Format$(rst!tseLogDate, sgShowDateForm)
                tlTSE.sStartTime = Format$(rst!tseStartTime, sgShowTimeWSecForm)
                tlTSE.sDescription = rst!tseDescription
                tlTSE.sState = rst!tseState
                tlTSE.lCteCode = rst!tseCteCode
                tlTSE.iVersion = rst!tseVersion
                tlTSE.lOrigTseCode = rst!tseOrigTseCode
                tlTSE.sCurrent = rst!tseCurrent
                tlTSE.sEnteredDate = Format$(rst!tseEnteredDate, sgShowDateForm)
                tlTSE.sEnteredTime = Format$(rst!tseEnteredTime, sgShowTimeWSecForm)
                tlTSE.iUieCode = rst!tseUieCode
                tlTSE.sUnused = rst!tseUnused
            End If
            rst.MoveNext
        Wend
        rst.Close
        gGetRec_TSE_TemplateSchdByDHETSE = True
        Exit Function
    Else
        tlTSE.lCode = 0
        rst.Close
        gGetRec_TSE_TemplateSchdByDHETSE = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_TSE_TemplateSchdByDHETSE = False
    tlTSE.lCode = 0
    Exit Function
End Function

Public Function gGetRec_TTE_TimeType(ilCode As Integer, slForm_Module As String, tlTTE As TTE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM TTE_Time_Type WHERE tteCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlTTE.iCode = rst!tteCode
        tlTTE.sType = rst!tteType
        tlTTE.sName = rst!tteName
        tlTTE.sDescription = rst!tteDescription
        tlTTE.sState = rst!tteState
        tlTTE.sUsedFlag = rst!tteUsedFlag
        tlTTE.iVersion = rst!tteVersion
        tlTTE.iOrigTteCode = rst!tteOrigTteCode
        tlTTE.sCurrent = rst!tteCurrent
        tlTTE.sEnteredDate = Format$(rst!tteEnteredDate, sgShowDateForm)
        tlTTE.sEnteredTime = Format$(rst!tteEnteredTime, sgShowTimeWSecForm)
        tlTTE.iUieCode = rst!tteUieCode
        tlTTE.sUnused = rst!tteUnused
        rst.Close
        gGetRec_TTE_TimeType = True
        Exit Function
    Else
        rst.Close
        gGetRec_TTE_TimeType = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_TTE_TimeType = False
    Exit Function
End Function

Public Function gGetTypeOfRecs_SOE_SiteOption(slGetType As String, slSOEStamp As String, slForm_Module As String, tlSOE() As SOE) As Integer
'
'   slGetType(I)- C=Current; H=History, B=Both
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim tlCurrSOE As SOE
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "soe.eng") & slGetType
    
    ilRet = 0
    On Error GoTo ErrHand
    If (slSOEStamp <> "") Then
        If slStamp = slSOEStamp Then
            gGetTypeOfRecs_SOE_SiteOption = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM SOE_Site_Option"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM SOE_Site_Option WHERE soeCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM SOE_Site_Option WHERE soeCurrent = 'Y'"
    End If
    Set rst = cnn.Execute(sgSQLQuery)
    If (rst.EOF) And (slGetType = "C") Then
        ReDim tlSOE(0 To 0) As SOE
        ilUpper = 0
        slNowDate = Format(Now, sgShowDateForm)  'Format(gNow(), sgShowDateForm)
        slNowTime = Format(Now, sgShowTimeWSecForm)  'Format(gNow(), sgShowTimeWSecForm)
        tlSOE(ilUpper).iCode = 0
        tlSOE(ilUpper).sClientName = "Client Name"
        tlSOE(ilUpper).sAddr1 = ""
        tlSOE(ilUpper).sAddr2 = ""
        tlSOE(ilUpper).sAddr3 = ""
        tlSOE(ilUpper).sPhone = ""
        tlSOE(ilUpper).sFax = ""
        tlSOE(ilUpper).iDaysRetainAsAir = 0
        tlSOE(ilUpper).lChgInterval = 0
        tlSOE(ilUpper).sMergeDateFormat = ""
        tlSOE(ilUpper).sMergeTimeFormat = ""
        tlSOE(ilUpper).sMergeFileFormat = ""
        tlSOE(ilUpper).sMergeFileExt = ""
        tlSOE(ilUpper).sMergeStartTime = Format("12am", sgShowTimeWSecForm)
        tlSOE(ilUpper).sMergeEndTime = Format("12am", sgShowTimeWSecForm)
        tlSOE(ilUpper).iMergeChkInterval = 0
        tlSOE(ilUpper).sMergeStopFlag = "N"
        tlSOE(ilUpper).iAlertInterval = 0
        tlSOE(ilUpper).sSchAutoGenSeq = "I"
        tlSOE(ilUpper).lMinEventID = 0
        tlSOE(ilUpper).lMaxEventID = 99999
        tlSOE(ilUpper).lCurrEventID = 10000
        tlSOE(ilUpper).iNoDaysRetainPW = 0
        tlSOE(ilUpper).iVersion = 0
        tlSOE(ilUpper).iOrigSoeCode = 0
        tlSOE(ilUpper).sCurrent = "Y"
        tlSOE(ilUpper).sEnteredDate = slNowDate
        tlSOE(ilUpper).sEnteredTime = slNowTime
        tlSOE(ilUpper).iUieCode = tgUIE.iCode
        tlSOE(ilUpper).iSpotItemIDWindow = 1000
        tlSOE(ilUpper).lTimeTolerance = 0
        tlSOE(ilUpper).lLengthTolerance = 0
        tlSOE(ilUpper).sMatchATNotB = "Y"
        tlSOE(ilUpper).sMatchATBNotI = "Y"
        tlSOE(ilUpper).sMatchANotT = "Y"
        tlSOE(ilUpper).sMatchBNotT = "Y"
        tlSOE(ilUpper).sSchAutoGenSeqTst = "I"
        tlSOE(ilUpper).sMergeStopFlagTst = "Y"
'        tlSOE(ilUpper).sUnused = ""
'        ilRet = gPutInsert_SOE_SiteOption(tlCurrSOE, slForm_Module)
'        If Not ilRet Then
'            gGetTypeOfRecs_SOE_SiteOption = False
'            Exit Function
'        End If
'        sgSQLQuery = "SELECT * FROM SOE_Site_Option WHERE soeCurrent = 'Y'"
'        Set rst = cnn.Execute(sgSQLQuery)
        rst.Close
        gGetTypeOfRecs_SOE_SiteOption = True
        Exit Function
    End If
    ReDim tlSOE(0 To 1) As SOE
    ilUpper = 0
    If Not rst.EOF Then
        tlSOE(ilUpper).iCode = rst!soeCode
        tlSOE(ilUpper).sClientName = rst!soeClientName
        tlSOE(ilUpper).sAddr1 = rst!soeAddr1
        tlSOE(ilUpper).sAddr2 = rst!soeAddr2
        tlSOE(ilUpper).sAddr3 = rst!soeAddr3
        tlSOE(ilUpper).sPhone = rst!soePhone
        tlSOE(ilUpper).sFax = rst!soeFax
        tlSOE(ilUpper).iDaysRetainAsAir = rst!soeDaysRetainAsAir
        tlSOE(ilUpper).lChgInterval = rst!soeChgInterval
        tlSOE(ilUpper).sMergeDateFormat = rst!soeMergeDateFormat
        tlSOE(ilUpper).sMergeTimeFormat = rst!soeMergeTimeFormat
        tlSOE(ilUpper).sMergeFileFormat = rst!soeMergeFileFormat
        tlSOE(ilUpper).sMergeFileExt = rst!soeMergeFileExt
        tlSOE(ilUpper).sMergeStartTime = Format$(rst!soeMergeStartTime, sgShowTimeWSecForm)
        tlSOE(ilUpper).sMergeEndTime = Format$(rst!soeMergeEndTime, sgShowTimeWSecForm)
        tlSOE(ilUpper).iMergeChkInterval = rst!soeMergeChkInterval
        tlSOE(ilUpper).sMergeStopFlag = rst!soeMergeStopFlag
        tlSOE(ilUpper).iAlertInterval = rst!soeAlertInterval
        tlSOE(ilUpper).sSchAutoGenSeq = rst!soeSchAutoGenSeq
        tlSOE(ilUpper).lMinEventID = rst!soeMinEventID
        tlSOE(ilUpper).lMaxEventID = rst!soeMaxEventID
        tlSOE(ilUpper).lCurrEventID = rst!soeCurrEventID
        tlSOE(ilUpper).iNoDaysRetainPW = rst!soeNoDaysRetainPW
        tlSOE(ilUpper).iVersion = rst!soeVersion
        tlSOE(ilUpper).iOrigSoeCode = rst!soeOrigSOECode
        tlSOE(ilUpper).sCurrent = rst!soeCurrent
        tlSOE(ilUpper).sEnteredDate = Format$(rst!soeEnteredDate, sgShowDateForm)
        tlSOE(ilUpper).sEnteredTime = Format$(rst!soeEnteredTime, sgShowTimeWSecForm)
        tlSOE(ilUpper).iUieCode = rst!soeUieCode
        tlSOE(ilUpper).iSpotItemIDWindow = rst!soeSpotItemIDWindow
        tlSOE(ilUpper).lTimeTolerance = rst!soeTimeTolerance
        tlSOE(ilUpper).lLengthTolerance = rst!soeLengthTolerance
        tlSOE(ilUpper).sMatchATNotB = rst!soeMatchATNotB
        tlSOE(ilUpper).sMatchATBNotI = rst!soeMatchATBNotI
        tlSOE(ilUpper).sMatchANotT = rst!soeMatchANotT
        tlSOE(ilUpper).sMatchBNotT = rst!soeMatchBNotT
        tlSOE(ilUpper).sSchAutoGenSeqTst = rst!soeSchAutoGenSeqTst
        tlSOE(ilUpper).sMergeStopFlagTst = rst!soeMergeStopFlagTst
        tlSOE(ilUpper).sUnused = ""
    End If
    slSOEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_SOE_SiteOption = True
    Exit Function
    
gGetTypeOfRecs_SOE_SiteOptionErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_SOE_SiteOption = False
    Exit Function

End Function

Public Function gGetAll_TNE_TaskName(slForm_Module As String) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "tne.eng")
    
    On Error GoTo gGetAll_TNE_TaskNameErr
    ilRet = 0
    ilLowLimit = LBound(tgCurrTNE)
    If ilRet <> 0 Then
        sgCurrUIEStamp = ""
    End If
    On Error GoTo ErrHand
    If (sgCurrTNEStamp <> "") Then
        sgSQLQuery = "SELECT Count(tneCode) FROM TNE_Task_Name"
        Set rst = cnn.Execute(sgSQLQuery)
        If rst(0).Value = UBound(tgCurrTNE) Then
            rst.Close
            gGetAll_TNE_TaskName = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM TNE_Task_Name ORDER BY tneType"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tgCurrTNE(0 To 0) As TNE
    ilUpper = 0
    While Not rst.EOF
        tgCurrTNE(ilUpper).iCode = rst!tneCode
        tgCurrTNE(ilUpper).sType = rst!tneType
        tgCurrTNE(ilUpper).sName = rst!tneName
        tgCurrTNE(ilUpper).sUnused = ""
        ilUpper = ilUpper + 1
        ReDim Preserve tgCurrTNE(0 To ilUpper) As TNE
        rst.MoveNext
    Wend
    sgCurrTNEStamp = FileDateTime(sgDBPath & "tne.eng")
    rst.Close
    gGetAll_TNE_TaskName = True
    Exit Function
    
gGetAll_TNE_TaskNameErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetAll_TNE_TaskName = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_UIE_UserInfo(slGetType As String, slUIEStamp As String, slForm_Module As String, tlUie() As UIE) As Integer
'
'   slGetType(I)- C=Current; H=History, B=Both
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim tlCurrUIE As UIE
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "uie.eng") & slGetType
    slStamp = slGetType
    
    On Error GoTo gGetTypeOfRecs_UIE_UserInfoErr
    ilRet = 0
    ilLowLimit = LBound(tlUie)
    If ilRet <> 0 Then
        slUIEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slUIEStamp <> "") Then
        If slGetType = "B" Then
            sgSQLQuery = "SELECT Count(uieCode) FROM UIE_User_Info"
        ElseIf slGetType = "H" Then
            sgSQLQuery = "SELECT Count(uieCode) FROM UIE_User_Info WHERE uieCurrent = 'N'"
        Else
            sgSQLQuery = "SELECT Count(uieCode) FROM UIE_User_Info WHERE uieCurrent = 'Y'"
        End If
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlUie)) And (slStamp = slUIEStamp) Then
            rst.Close
            gGetTypeOfRecs_UIE_UserInfo = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM UIE_User_Info"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM UIE_User_Info WHERE uieCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM UIE_User_Info WHERE uieCurrent = 'Y'"
    End If
    Set rst = cnn.Execute(sgSQLQuery)
    If (rst.EOF) And (slGetType = "C") Then
        slNowDate = Format(Now, sgShowDateForm)  'Format(gNow(), sgShowDateForm)
        slNowTime = Format(Now, sgShowTimeWSecForm)  'Format(gNow(), sgShowTimeWSecForm)
        tlCurrUIE.iCode = 1
        tlCurrUIE.sSignOnName = "Guide"
        tlCurrUIE.sPassword = "Guide"
        tlCurrUIE.sLastDatePWSet = slNowDate
        tlCurrUIE.sShowName = "Guide"
        tlCurrUIE.sState = "A"
        tlCurrUIE.sEMail = ""
        tlCurrUIE.sLastSignOnDate = slNowDate
        tlCurrUIE.sLastSignOnTime = slNowTime
        tlCurrUIE.sUsedFlag = "Y"
        tlCurrUIE.iVersion = 0
        tlCurrUIE.iOrigUieCode = 1
        tlCurrUIE.sCurrent = "Y"
        tlCurrUIE.sEnteredDate = slNowDate
        tlCurrUIE.sEnteredTime = slNowTime
        tlCurrUIE.iUieCode = 1
        tlCurrUIE.sUnused = ""
        ilRet = gPutInsert_UIE_UserInfo(0, tlCurrUIE, slForm_Module)
        If Not ilRet Then
            gGetTypeOfRecs_UIE_UserInfo = False
            Exit Function
        End If
        sgSQLQuery = "SELECT * FROM UIE_User_Info WHERE uieCurrent = 'Y'"
        Set rst = cnn.Execute(sgSQLQuery)
    End If
    ReDim tlUie(0 To 0) As UIE
    ilUpper = 0
    While Not rst.EOF
        tlUie(ilUpper).iCode = rst!uieCode
        tlUie(ilUpper).sSignOnName = rst!uieSignOnName
        tlUie(ilUpper).sPassword = rst!uiePassword
        tlUie(ilUpper).sLastDatePWSet = Format$(rst!uieLastDatePWSet, sgShowDateForm)
        tlUie(ilUpper).sShowName = rst!uieShowName
        tlUie(ilUpper).sState = rst!uieState
        tlUie(ilUpper).sEMail = rst!uieEmail
        tlUie(ilUpper).sLastSignOnDate = Format$(rst!uieLastSignOnDate, sgShowDateForm)
        tlUie(ilUpper).sLastSignOnTime = Format$(rst!uieLastSignOnTime, sgShowTimeWSecForm)
        tlUie(ilUpper).sUsedFlag = rst!uieUsedFlag
        tlUie(ilUpper).iVersion = rst!uieVersion
        tlUie(ilUpper).iOrigUieCode = rst!uieOrigUieCode
        tlUie(ilUpper).sCurrent = rst!uieCurrent
        tlUie(ilUpper).sEnteredDate = Format$(rst!uieEnteredDate, sgShowDateForm)
        tlUie(ilUpper).sEnteredTime = Format$(rst!uieEnteredTime, sgShowTimeWSecForm)
        tlUie(ilUpper).iUieCode = rst!uieUieCode
        tlUie(ilUpper).sUnused = ""
        ilUpper = ilUpper + 1
        ReDim Preserve tlUie(0 To ilUpper) As UIE
        rst.MoveNext
    Wend
    slUIEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_UIE_UserInfo = True
    Exit Function
    
gGetTypeOfRecs_UIE_UserInfoErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_UIE_UserInfo = False
    Exit Function

End Function


Public Function gGetRecs_ITE_ItemTest(slITEStamp As String, ilSoeCode As Integer, slForm_Module As String, tlITE() As ITE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "ite.eng") & ilSoeCode
    
    On Error GoTo gGetRecs_ITE_ItemTestErr
    ilRet = 0
    ilLowLimit = LBound(tlITE)
    If ilRet <> 0 Then
        slITEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slITEStamp <> "") Then
        If slStamp = slITEStamp Then
            gGetRecs_ITE_ItemTest = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM ITE_Item_Test Where iteSoeCode = " & ilSoeCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlITE(0 To 0) As ITE
    ilUpper = 0
    While Not rst.EOF
        tlITE(ilUpper).iCode = rst!iteCode
        tlITE(ilUpper).iSoeCode = rst!iteSoeCode
        tlITE(ilUpper).sType = rst!iteType
        tlITE(ilUpper).sName = rst!iteName
        tlITE(ilUpper).iDataBits = rst!iteDataBits
        tlITE(ilUpper).sParity = rst!iteParity
        tlITE(ilUpper).sStopBit = rst!iteStopBit
        tlITE(ilUpper).iBaud = rst!iteBaud
        tlITE(ilUpper).sMachineID = rst!iteMachineID
        tlITE(ilUpper).sStartCode = rst!iteStartCode
        tlITE(ilUpper).sReplyCode = rst!iteReplyCode
        tlITE(ilUpper).iMinMgsID = rst!iteMinMgsID
        tlITE(ilUpper).iMaxMgsID = rst!iteMaxMgsID
        tlITE(ilUpper).iCurrMgsID = rst!iteCurrMgsID
        tlITE(ilUpper).sMgsType = rst!iteMgsType
        tlITE(ilUpper).sCheckSum = rst!iteCheckSum
        tlITE(ilUpper).sCmmdSeq = rst!iteCmmdSeq
        tlITE(ilUpper).sMgsEndCode = rst!iteMgsEndCode
        tlITE(ilUpper).sTitleID = rst!iteTitleID
        tlITE(ilUpper).sLengthID = rst!iteLengthID
        tlITE(ilUpper).sConnectSeq = rst!iteConnectSeq
        tlITE(ilUpper).sMgsErrType = rst!iteMgsErrType
        tlITE(ilUpper).sUnused = ""
        ilUpper = ilUpper + 1
        ReDim Preserve tlITE(0 To ilUpper) As ITE
        rst.MoveNext
    Wend
    slITEStamp = slStamp
    rst.Close
    gGetRecs_ITE_ItemTest = True
    Exit Function
    
gGetRecs_ITE_ItemTestErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_ITE_ItemTest = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_ACE_AutoContact(slGetType As String, slACEStamp As String, slForm_Module As String, tlACE() As ACE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "ace.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_ACE_AutoContactErr
    ilRet = 0
    ilLowLimit = LBound(tlACE)
    If ilRet <> 0 Then
        slACEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slACEStamp <> "") Then
        If slStamp = slACEStamp Then
            gGetTypeOfRecs_ACE_AutoContact = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM ACE_Auto_Contact, AEE_Auto_Equip ORDER BY aceAeeCode"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM ACE_Auto_Contact, AEE_Auto_Equip WHERE aceAeeCode = aeeCode and aeeCurrent = 'N' ORDER BY aceAeeCode"
    Else
        sgSQLQuery = "SELECT * FROM ACE_Auto_Contact, AEE_Auto_Equip WHERE aceAeeCode = aeeCode and aeeCurrent = 'Y' ORDER BY aceAeeCode"
    End If
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlACE(0 To 0) As ACE
    ilUpper = 0
    While Not rst.EOF
        tlACE(ilUpper).iCode = rst!aceCode
        tlACE(ilUpper).iAeeCode = rst!aceAeeCode
        tlACE(ilUpper).sType = rst!aceType
        tlACE(ilUpper).sContact = rst!aceContact
        tlACE(ilUpper).sPhone = rst!acePhone
        tlACE(ilUpper).sFax = rst!aceFax
        tlACE(ilUpper).sEMail = rst!aceEMail
        tlACE(ilUpper).sUnused = rst!aceUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlACE(0 To ilUpper) As ACE
        rst.MoveNext
    Wend
    slACEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_ACE_AutoContact = True
    Exit Function
    
gGetTypeOfRecs_ACE_AutoContactErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_ACE_AutoContact = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_ADE_AutoDataFlags(slGetType As String, slADEStamp As String, slForm_Module As String, tlADE() As ADE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "ade.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_ADE_AutoDataFlagsErr
    ilRet = 0
    ilLowLimit = LBound(tlADE)
    If ilRet <> 0 Then
        slADEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slADEStamp <> "") Then
        If slStamp = slADEStamp Then
            gGetTypeOfRecs_ADE_AutoDataFlags = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM ADE_Auto_Data_Flags, AEE_Auto_Equip ORDER BY adeAeeCode"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM ADE_Auto_Data_Flags, AEE_Auto_Equip WHERE adeAeeCode = aeeCode and aeeCurrent = 'N' ORDER BY adeAeeCode"
    Else
        sgSQLQuery = "SELECT * FROM ADE_Auto_Data_Flags, AEE_Auto_Equip WHERE adeAeeCode = aeeCode and aeeCurrent = 'Y' ORDER BY adeAeeCode"
    End If
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlADE(0 To 0) As ADE
    ilUpper = 0
    While Not rst.EOF
        tlADE(ilUpper).iCode = rst!adeCode
        tlADE(ilUpper).iAeeCode = rst!adeAeeCode
        tlADE(ilUpper).iScheduleData = rst!adeScheduleData
        tlADE(ilUpper).iDate = rst!adeDate
        tlADE(ilUpper).iDateNoChar = rst!adeDateNoChar
        tlADE(ilUpper).iTime = rst!adeTime
        tlADE(ilUpper).iTimeNoChar = rst!adeTimeNoChar
        tlADE(ilUpper).iAutoOff = rst!adeAutoOff
        tlADE(ilUpper).iData = rst!adeData
        tlADE(ilUpper).iSchedule = rst!adeSchedule
        tlADE(ilUpper).iTrueTime = rst!adeTrueTime
        tlADE(ilUpper).iSourceConflict = rst!adeSourceConflict
        tlADE(ilUpper).iSourceUnavail = rst!adeSourceUnavail
        tlADE(ilUpper).iSourceItem = rst!adeSourceItem
        tlADE(ilUpper).iBkupSrceUnavail = rst!adeBkupSrceUnavail
        tlADE(ilUpper).iBkupSrceItem = rst!adeBkupSrceItem
        tlADE(ilUpper).iProtSrceUnavail = rst!adeProtSrceUnavail
        tlADE(ilUpper).iProtSrceItem = rst!adeProtSrceItem
        tlADE(ilUpper).sUnused = rst!adeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlADE(0 To ilUpper) As ADE
        rst.MoveNext
    Wend
    slADEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_ADE_AutoDataFlags = True
    Exit Function
    
gGetTypeOfRecs_ADE_AutoDataFlagsErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    Screen.MousePointer = vbDefault
    rst.Close
    gGetTypeOfRecs_ADE_AutoDataFlags = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_AEE_AutoEquip(slGetType As String, slAEEStamp As String, slForm_Module As String, tlAEE() As AEE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "aee.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_AEE_AutoEquipErr
    ilRet = 0
    ilLowLimit = LBound(tlAEE)
    If ilRet <> 0 Then
        slAEEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slAEEStamp <> "") Then
        If slStamp = slAEEStamp Then
            gGetTypeOfRecs_AEE_AutoEquip = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM AEE_Auto_Equip"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM AEE_Auto_Equip WHERE aeeCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM AEE_Auto_Equip WHERE aeeCurrent = 'Y'"
    End If
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlAEE(0 To 0) As AEE
    ilUpper = 0
    While Not rst.EOF
        tlAEE(ilUpper).iCode = rst!aeeCode
        tlAEE(ilUpper).sName = rst!aeeName
        tlAEE(ilUpper).sDescription = rst!aeeDescription
        tlAEE(ilUpper).sManufacture = rst!aeeManufacture
        tlAEE(ilUpper).sFixedTimeChar = rst!aeeFixedTimeChar
        tlAEE(ilUpper).lAlertSchdDelay = rst!aeeAlertSchdDelay
        tlAEE(ilUpper).sState = rst!aeeState
        tlAEE(ilUpper).sUsedFlag = rst!aeeUsedFlag
        tlAEE(ilUpper).iVersion = rst!aeeVersion
        tlAEE(ilUpper).iOrigAeeCode = rst!aeeOrigAeeCode
        tlAEE(ilUpper).sCurrent = rst!aeeCurrent
        tlAEE(ilUpper).sEnteredDate = Format$(rst!aeeEnteredDate, sgShowDateForm)
        tlAEE(ilUpper).sEnteredTime = Format$(rst!aeeEnteredTime, sgShowTimeWSecForm)
        tlAEE(ilUpper).iUieCode = rst!aeeUieCode
        tlAEE(ilUpper).sUnused = rst!aeeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlAEE(0 To ilUpper) As AEE
        rst.MoveNext
    Wend
    slAEEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_AEE_AutoEquip = True
    Exit Function
    
gGetTypeOfRecs_AEE_AutoEquipErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_AEE_AutoEquip = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_AFE_AutoFormat(slGetType As String, slAFEStamp As String, slForm_Module As String, tlAFE() As AFE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "afe.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_AFE_AutoFormatErr
    ilRet = 0
    ilLowLimit = LBound(tlAFE)
    If ilRet <> 0 Then
        slAFEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slAFEStamp <> "") Then
        If slStamp = slAFEStamp Then
            gGetTypeOfRecs_AFE_AutoFormat = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM AFE_Auto_Format, AEE_Auto_Equip ORDER BY afeAeeCode"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM AFE_Auto_Format, AEE_Auto_Equip WHERE afeAeeCode = aeeCode and aeeCurrent = 'N' ORDER BY afeAeeCode"
    Else
        sgSQLQuery = "SELECT * FROM AFE_Auto_Format, AEE_Auto_Equip WHERE afeAeeCode = aeeCode and aeeCurrent = 'Y' ORDER BY afeAeeCode"
    End If
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlAFE(0 To 0) As AFE
    ilUpper = 0
    While Not rst.EOF
        tlAFE(ilUpper).iCode = rst!afeCode
        tlAFE(ilUpper).iAeeCode = rst!afeAeeCode
        tlAFE(ilUpper).sType = rst!afeType
        tlAFE(ilUpper).sSubType = rst!afeSubType
        tlAFE(ilUpper).iBus = rst!afeBus
        tlAFE(ilUpper).iBusControl = rst!afeBusControl
        tlAFE(ilUpper).iEventType = rst!afeEventType
        tlAFE(ilUpper).iTime = rst!afeTime
        tlAFE(ilUpper).iStartType = rst!afeStartType
        tlAFE(ilUpper).iFixedTime = rst!afeFixedTime
        tlAFE(ilUpper).iEndType = rst!afeEndType
        tlAFE(ilUpper).iDuration = rst!afeDuration
        tlAFE(ilUpper).iEndTime = rst!afeEndTime
        tlAFE(ilUpper).iMaterialType = rst!afeMaterialType
        tlAFE(ilUpper).iAudioName = rst!afeAudioName
        tlAFE(ilUpper).iAudioItemID = rst!afeAudioItemID
        tlAFE(ilUpper).iAudioISCI = rst!afeAudioISCI
        tlAFE(ilUpper).iAudioControl = rst!afeAudioControl
        tlAFE(ilUpper).iBkupAudioName = rst!afeBkupAudioName
        tlAFE(ilUpper).iBkupAudioControl = rst!afeBkupAudioControl
        tlAFE(ilUpper).iProtAudioName = rst!afeProtAudioName
        tlAFE(ilUpper).iProtItemID = rst!afeProtItemID
        tlAFE(ilUpper).iProtISCI = rst!afeProtISCI
        tlAFE(ilUpper).iProtAudioControl = rst!afeProtAudioControl
        tlAFE(ilUpper).iRelay1 = rst!afeRelay1
        tlAFE(ilUpper).iRelay2 = rst!afeRelay2
        tlAFE(ilUpper).iFollow = rst!afeFollow
        tlAFE(ilUpper).iSilenceTime = rst!afeSilenceTime
        tlAFE(ilUpper).iSilence1 = rst!afeSilence1
        tlAFE(ilUpper).iSilence2 = rst!afeSilence2
        tlAFE(ilUpper).iSilence3 = rst!afeSilence3
        tlAFE(ilUpper).iSilence4 = rst!afeSilence4
        tlAFE(ilUpper).iStartNetcue = rst!afeStartNetcue
        tlAFE(ilUpper).iStopNetcue = rst!afeStopNetcue
        tlAFE(ilUpper).iTitle1 = rst!afeTitle1
        tlAFE(ilUpper).iTitle2 = rst!afeTitle2
        tlAFE(ilUpper).iEventID = rst!afeEventID
        tlAFE(ilUpper).iDate = rst!afeDate
        tlAFE(ilUpper).iABCFormat = rst!afeABCFormat
        tlAFE(ilUpper).iABCPgmCode = rst!afeABCPgmCode
        tlAFE(ilUpper).iABCXDSMode = rst!afeABCXDSMode
        tlAFE(ilUpper).iABCRecordItem = rst!afeABCRecordItem
        tlAFE(ilUpper).sUnused = rst!afeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlAFE(0 To ilUpper) As AFE
        rst.MoveNext
    Wend
    slAFEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_AFE_AutoFormat = True
    Exit Function
    
gGetTypeOfRecs_AFE_AutoFormatErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_AFE_AutoFormat = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_AIE_ActiveInfo(slRefFileName As String, slStartDate As String, slEndDate As String, slAIEStamp As String, slForm_Module As String, tlAIE() As AIE) As Integer
'
'   slRefFileName(I)- Referenced file name
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "aie.eng") & slRefFileName & slStartDate & slEndDate
    
    On Error GoTo gGetTypeOfRecs_AIE_ActiveInfoErr
    ilRet = 0
    ilLowLimit = LBound(tlAIE)
    If ilRet <> 0 Then
        slAIEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slAIEStamp <> "") Then
        If slStamp = slAIEStamp Then
            gGetTypeOfRecs_AIE_ActiveInfo = True
            Exit Function
        End If
    End If
    If slRefFileName <> "" Then
        sgSQLQuery = "SELECT * FROM AIE_Active_Info WHERE aieRefFileName = '" & slRefFileName & "'"
        If Trim$(slStartDate) <> "" Then
            sgSQLQuery = sgSQLQuery & " AND aieEnteredDate >= '" & Format$(gAdjYear(slStartDate), sgSQLDateForm) & "'"
        End If
        If Trim$(slEndDate) <> "" Then
            sgSQLQuery = sgSQLQuery & " AND aieEnteredDate <= '" & Format$(gAdjYear(slEndDate), sgSQLDateForm) & "'"
        End If
    Else
        sgSQLQuery = "SELECT * FROM AIE_Active_Info"
        If Trim$(slStartDate) <> "" Then
            sgSQLQuery = sgSQLQuery & " WHERE aieEnteredDate >= '" & Format$(gAdjYear(slStartDate), sgSQLDateForm) & "'"
            If Trim$(slEndDate) <> "" Then
                sgSQLQuery = sgSQLQuery & " AND aieEnteredDate <= '" & Format$(gAdjYear(slEndDate), sgSQLDateForm) & "'"
            End If
        Else
            If Trim$(slEndDate) <> "" Then
                sgSQLQuery = sgSQLQuery & " WHERE aieEnteredDate <= '" & Format$(gAdjYear(slEndDate), sgSQLDateForm) & "'"
            End If
        End If
    End If
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlAIE(0 To 0) As AIE
    ilUpper = 0
    While Not rst.EOF
        tlAIE(ilUpper).lCode = rst!aieCode
        tlAIE(ilUpper).iUieCode = rst!aieUieCode
        tlAIE(ilUpper).sEnteredDate = Format$(rst!aieEnteredDate, sgShowDateForm)
        tlAIE(ilUpper).sEnteredTime = Format$(rst!aieEnteredTime, sgShowTimeWSecForm)
        tlAIE(ilUpper).sRefFileName = rst!aieRefFileName
        tlAIE(ilUpper).lToFileCode = rst!aieToFileCode
        tlAIE(ilUpper).lFromFileCode = rst!aieFromFileCode
        tlAIE(ilUpper).lOrigFileCode = rst!aieOrigFileCode
        tlAIE(ilUpper).sUnused = rst!aieUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlAIE(0 To ilUpper) As AIE
        rst.MoveNext
    Wend
    slAIEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_AIE_ActiveInfo = True
    Exit Function
    
gGetTypeOfRecs_AIE_ActiveInfoErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_AIE_ActiveInfo = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_ANE_AudioName(slGetType As String, slANEStamp As String, slForm_Module As String, tlANE() As ANE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim ilLoop As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "ane.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_ANE_AudioNameErr
    ilRet = 0
    ilLowLimit = LBound(tlANE)
    If ilRet <> 0 Then
        slANEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slANEStamp <> "") Then
        If slStamp = slANEStamp Then
            gGetTypeOfRecs_ANE_AudioName = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM ANE_Audio_Name"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM ANE_Audio_Name WHERE aneCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM ANE_Audio_Name WHERE aneCurrent = 'Y'"
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY aneCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlANE(0 To 0) As ANE
    ilUpper = 0
    While Not rst.EOF
        tlANE(ilUpper).iCode = rst!aneCode
        tlANE(ilUpper).sName = rst!aneName
        tlANE(ilUpper).sDescription = rst!aneDescription
        tlANE(ilUpper).iCceCode = rst!aneCceCode
        tlANE(ilUpper).iAteCode = rst!aneAteCode
        tlANE(ilUpper).sState = rst!aneState
        tlANE(ilUpper).sUsedFlag = rst!aneUsedFlag
        tlANE(ilUpper).iVersion = rst!aneVersion
        tlANE(ilUpper).iOrigAneCode = rst!aneOrigAneCode
        tlANE(ilUpper).sCurrent = rst!aneCurrent
        tlANE(ilUpper).sEnteredDate = Format$(rst!aneEnteredDate, sgShowDateForm)
        tlANE(ilUpper).sEnteredTime = Format$(rst!aneEnteredTime, sgShowTimeWSecForm)
        tlANE(ilUpper).iUieCode = rst!aneUieCode
        tlANE(ilUpper).sCheckConflicts = rst!aneCheckConflicts
        tlANE(ilUpper).sUnused = rst!aneUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlANE(0 To ilUpper) As ANE
        rst.MoveNext
    Wend
    ReDim tgCurrANE_Name(0 To ilUpper) As NAMESORT
    For ilLoop = 0 To ilUpper Step 1
        tgCurrANE_Name(ilLoop).sKey = tlANE(ilLoop).sName
        tgCurrANE_Name(ilLoop).iCode = tlANE(ilLoop).iCode
    Next ilLoop
    If ilUpper > 0 Then
        ArraySortTyp fnAV(tgCurrANE_Name(), 0), ilUpper, 0, LenB(tgCurrANE_Name(0)), 0, LenB(tgCurrANE_Name(0).sKey), 0
    End If
    slANEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_ANE_AudioName = True
    Exit Function
    
gGetTypeOfRecs_ANE_AudioNameErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_ANE_AudioName = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_APE_AutoPath(slGetType As String, slAPEStamp As String, slForm_Module As String, tlAPE() As APE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "ape.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_APE_AutoPathErr
    ilRet = 0
    ilLowLimit = LBound(tlAPE)
    If ilRet <> 0 Then
        slAPEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slAPEStamp <> "") Then
        If slStamp = slAPEStamp Then
            gGetTypeOfRecs_APE_AutoPath = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM APE_Auto_Path, AEE_Auto_Equip ORDER BY apceAeeCode"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM APE_Auto_Path, AEE_Auto_Equip WHERE apeAeeCode = aeeCode and aeeCurrent = 'N' ORDER BY apeAeeCode"
    Else
        sgSQLQuery = "SELECT * FROM APE_Auto_Path, AEE_Auto_Equip WHERE apeAeeCode = aeeCode and aeeCurrent = 'Y' ORDER BY apeAeeCode"
    End If
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlAPE(0 To 0) As APE
    ilUpper = 0
    While Not rst.EOF
        tlAPE(ilUpper).iCode = rst!apeCode
        tlAPE(ilUpper).iAeeCode = rst!apeAeeCode
        tlAPE(ilUpper).sType = rst!apeType
        tlAPE(ilUpper).sSubType = rst!apeSubType
        If (tlAPE(ilUpper).sSubType <> "P") And (tlAPE(ilUpper).sSubType <> "T") Then
            tlAPE(ilUpper).sSubType = "P"
        End If
        tlAPE(ilUpper).sNewFileName = rst!apeNewFileName
        tlAPE(ilUpper).sChgFileName = rst!apeChgFileName
        tlAPE(ilUpper).sDelFileName = rst!apeDelFileName
        tlAPE(ilUpper).sNewFileExt = rst!apeNewFileExt
        tlAPE(ilUpper).sChgFileExt = rst!apeChgFileExt
        tlAPE(ilUpper).sDelFileExt = rst!apeDelFileExt
        tlAPE(ilUpper).sPath = rst!apePath
        tlAPE(ilUpper).sDateFormat = rst!apeDateFormat
        tlAPE(ilUpper).sTimeFormat = rst!apeTimeFormat
        tlAPE(ilUpper).sUnused = rst!apeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlAPE(0 To ilUpper) As APE
        rst.MoveNext
    Wend
    slAPEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_APE_AutoPath = True
    Exit Function
    
gGetTypeOfRecs_APE_AutoPathErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_APE_AutoPath = False
    Exit Function
End Function

Public Function gGetTypeOfRecs_ASE_AudioSource(slGetType As String, slASEStamp As String, slForm_Module As String, tlASE() As ASE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "ase.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_ASE_AudioSourceErr
    ilRet = 0
    ilLowLimit = LBound(tlASE)
    If ilRet <> 0 Then
        slASEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slASEStamp <> "") Then
        If slStamp = slASEStamp Then
            gGetTypeOfRecs_ASE_AudioSource = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM ASE_Audio_Source"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM ASE_Audio_Source WHERE aseCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM ASE_Audio_Source WHERE aseCurrent = 'Y'"
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY aseCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlASE(0 To 0) As ASE
    ilUpper = 0
    While Not rst.EOF
        tlASE(ilUpper).iCode = rst!aseCode
        tlASE(ilUpper).iPriAneCode = rst!asePriAneCode
        tlASE(ilUpper).iPriCceCode = rst!asePriCceCode
        tlASE(ilUpper).sDescription = rst!aseDescription
        tlASE(ilUpper).iBkupAneCode = rst!aseBkupAneCode
        tlASE(ilUpper).iBkupCceCode = rst!aseBkupCceCode
        tlASE(ilUpper).iProtAneCode = rst!aseProtAneCode
        tlASE(ilUpper).iProtCceCode = rst!aseProtCceCode
        tlASE(ilUpper).sState = rst!aseState
        tlASE(ilUpper).sUsedFlag = rst!aseUsedFlag
        tlASE(ilUpper).iVersion = rst!aseVersion
        tlASE(ilUpper).iOrigAseCode = rst!aseOrigAseCode
        tlASE(ilUpper).sCurrent = rst!aseCurrent
        tlASE(ilUpper).sEnteredDate = Format$(rst!aseEnteredDate, sgShowDateForm)
        tlASE(ilUpper).sEnteredTime = Format$(rst!aseEnteredTime, sgShowTimeWSecForm)
        tlASE(ilUpper).iUieCode = rst!aseUieCode
        tlASE(ilUpper).sUnused = rst!aseUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlASE(0 To ilUpper) As ASE
        rst.MoveNext
    Wend
    slASEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_ASE_AudioSource = True
    Exit Function
    
gGetTypeOfRecs_ASE_AudioSourceErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_ASE_AudioSource = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_ATE_AudioType(slGetType As String, slATEStamp As String, slForm_Module As String, tlATE() As ATE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "ate.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_ATE_AudioTypeErr
    ilRet = 0
    ilLowLimit = LBound(tlATE)
    If ilRet <> 0 Then
        slATEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slATEStamp <> "") Then
        If slStamp = slATEStamp Then
            gGetTypeOfRecs_ATE_AudioType = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM ATE_Audio_Type"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM ATE_Audio_Type WHERE ateCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM ATE_Audio_Type WHERE ateCurrent = 'Y'"
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY ateCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlATE(0 To 0) As ATE
    ilUpper = 0
    While Not rst.EOF
        tlATE(ilUpper).iCode = rst!ateCode
        tlATE(ilUpper).sName = rst!ateName
        tlATE(ilUpper).sDescription = rst!ateDescription
        tlATE(ilUpper).sState = rst!ateState
        tlATE(ilUpper).sTestItemID = rst!ateTestItemID
        tlATE(ilUpper).lPreBufferTime = rst!atePreBufferTime
        tlATE(ilUpper).lPostBufferTime = rst!atePostBufferTime
        tlATE(ilUpper).sUsedFlag = rst!ateUsedFlag
        tlATE(ilUpper).iVersion = rst!ateVersion
        tlATE(ilUpper).iOrigAteCode = rst!ateOrigAteCode
        tlATE(ilUpper).sCurrent = rst!ateCurrent
        tlATE(ilUpper).sEnteredDate = Format$(rst!ateEnteredDate, sgShowDateForm)
        tlATE(ilUpper).sEnteredTime = Format$(rst!ateEnteredTime, sgShowTimeWSecForm)
        tlATE(ilUpper).iUieCode = rst!ateUieCode
        tlATE(ilUpper).sUnused = rst!ateUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlATE(0 To ilUpper) As ATE
        rst.MoveNext
    Wend
    slATEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_ATE_AudioType = True
    Exit Function
    
gGetTypeOfRecs_ATE_AudioTypeErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_ATE_AudioType = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_BDE_BusDefinition(slGetType As String, slBDEStamp As String, slForm_Module As String, tlBDE() As BDE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim ilLoop As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "bde.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_BDE_BusDefinitionErr
    ilRet = 0
    ilLowLimit = LBound(tlBDE)
    If ilRet <> 0 Then
        slBDEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slBDEStamp <> "") Then
        If slStamp = slBDEStamp Then
            gGetTypeOfRecs_BDE_BusDefinition = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM BDE_Bus_Definition"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM BDE_Bus_Definition WHERE bdeCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM BDE_Bus_Definition WHERE bdeCurrent = 'Y'"
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY bdeCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlBDE(0 To 0) As BDE
    ilUpper = 0
    While Not rst.EOF
        tlBDE(ilUpper).iCode = rst!bdeCode
        tlBDE(ilUpper).sName = rst!bdeName
        tlBDE(ilUpper).sDescription = rst!bdeDescription
        tlBDE(ilUpper).sChannel = rst!bdeChannel
        tlBDE(ilUpper).iAseCode = rst!bdeAseCode
        tlBDE(ilUpper).sState = rst!bdeState
        tlBDE(ilUpper).iCceCode = rst!bdeCceCode
        tlBDE(ilUpper).sUsedFlag = rst!bdeUsedFlag
        tlBDE(ilUpper).iVersion = rst!bdeVersion
        tlBDE(ilUpper).iOrigBdeCode = rst!bdeOrigBdeCode
        tlBDE(ilUpper).sCurrent = rst!bdeCurrent
        tlBDE(ilUpper).sEnteredDate = Format$(rst!bdeEnteredDate, sgShowDateForm)
        tlBDE(ilUpper).sEnteredTime = Format$(rst!bdeEnteredTime, sgShowTimeWSecForm)
        tlBDE(ilUpper).iUieCode = rst!bdeUieCode
        tlBDE(ilUpper).sUnused = rst!bdeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlBDE(0 To ilUpper) As BDE
        rst.MoveNext
    Wend
    ReDim tgCurrBDE_Name(0 To ilUpper) As NAMESORT
    For ilLoop = 0 To ilUpper Step 1
        tgCurrBDE_Name(ilLoop).sKey = tlBDE(ilLoop).sName
        tgCurrBDE_Name(ilLoop).iCode = tlBDE(ilLoop).iCode
    Next ilLoop
    If ilUpper > 0 Then
        ArraySortTyp fnAV(tgCurrBDE_Name(), 0), ilUpper, 0, LenB(tgCurrBDE_Name(0)), 0, LenB(tgCurrBDE_Name(0).sKey), 0
    End If
    slBDEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_BDE_BusDefinition = True
    Exit Function
    
gGetTypeOfRecs_BDE_BusDefinitionErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_BDE_BusDefinition = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_BGE_BusGroup(slGetType As String, slBGEStamp As String, slForm_Module As String, tlBGE() As BGE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "bge.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_BGE_BusGroupErr
    ilRet = 0
    ilLowLimit = LBound(tlBGE)
    If ilRet <> 0 Then
        slBGEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slBGEStamp <> "") Then
        If slStamp = slBGEStamp Then
            gGetTypeOfRecs_BGE_BusGroup = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM BGE_Bus_Group"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM BGE_Bus_Group WHERE bgeCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM BGE_Bus_Group WHERE bgeCurrent = 'Y'"
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY bgeCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlBGE(0 To 0) As BGE
    ilUpper = 0
    While Not rst.EOF
        tlBGE(ilUpper).iCode = rst!bgeCode
        tlBGE(ilUpper).sName = rst!bgeName
        tlBGE(ilUpper).sDescription = rst!bgeDescription
        tlBGE(ilUpper).sState = rst!bgeState
        tlBGE(ilUpper).sUsedFlag = rst!bgeUsedFlag
        tlBGE(ilUpper).iVersion = rst!bgeVersion
        tlBGE(ilUpper).iOrigBgeCode = rst!bgeOrigBgeCode
        tlBGE(ilUpper).sCurrent = rst!bgeCurrent
        tlBGE(ilUpper).sEnteredDate = Format$(rst!bgeEnteredDate, sgShowDateForm)
        tlBGE(ilUpper).sEnteredTime = Format$(rst!bgeEnteredTime, sgShowTimeWSecForm)
        tlBGE(ilUpper).iUieCode = rst!bgeUieCode
        tlBGE(ilUpper).sUnused = rst!bgeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlBGE(0 To ilUpper) As BGE
        rst.MoveNext
    Wend
    slBGEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_BGE_BusGroup = True
    Exit Function
    
gGetTypeOfRecs_BGE_BusGroupErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_BGE_BusGroup = False
    Exit Function

End Function


Public Function gGetTypeOfRecs_CCE_ControlChar(slGetType As String, slControlType As String, slCCEStamp As String, slForm_Module As String, tlCCE() As CCE) As Integer
'
'   slGetType(I)- C=Current; H=History, B=Both
'   slControlType(I)- A=Audio; B=Bus
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "cce.eng") & slGetType & slControlType
    
    On Error GoTo gGetTypeOfRecs_CCE_ControlCharErr
    ilRet = 0
    ilLowLimit = LBound(tlCCE)
    If ilRet <> 0 Then
        slCCEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slCCEStamp <> "") Then
        If slStamp = slCCEStamp Then
            gGetTypeOfRecs_CCE_ControlChar = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM CCE_Control_Char"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM CCE_Control_Char WHERE cceCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM CCE_Control_Char WHERE cceCurrent = 'Y'"
    End If
    If slGetType = "B" Then
        If slControlType = "A" Or slControlType = "B" Then
            sgSQLQuery = sgSQLQuery & " Where cceType = '" & slControlType & "'"
        End If
    Else
        If slControlType = "A" Or slControlType = "B" Then
            sgSQLQuery = sgSQLQuery & " And cceType = '" & slControlType & "'"
        End If
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY cceCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlCCE(0 To 0) As CCE
    ilUpper = 0
    While Not rst.EOF
        tlCCE(ilUpper).iCode = rst!cceCode
        tlCCE(ilUpper).sType = rst!cceType
        tlCCE(ilUpper).sAutoChar = rst!cceAutoChar
        tlCCE(ilUpper).sDescription = rst!cceDescription
        tlCCE(ilUpper).sState = rst!cceState
        tlCCE(ilUpper).sUsedFlag = rst!cceUsedFlag
        tlCCE(ilUpper).iVersion = rst!cceVersion
        tlCCE(ilUpper).iOrigCceCode = rst!cceOrigCceCode
        tlCCE(ilUpper).sCurrent = rst!cceCurrent
        tlCCE(ilUpper).sEnteredDate = Format$(rst!cceEnteredDate, sgShowDateForm)
        tlCCE(ilUpper).sEnteredTime = Format$(rst!cceEnteredTime, sgShowTimeWSecForm)
        tlCCE(ilUpper).iUieCode = rst!cceUieCode
        tlCCE(ilUpper).sUnused = rst!cceUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlCCE(0 To ilUpper) As CCE
        rst.MoveNext
    Wend
    slCCEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_CCE_ControlChar = True
    Exit Function
    
gGetTypeOfRecs_CCE_ControlCharErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_CCE_ControlChar = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_CTE_CommtsTitle(slGetType As String, slType As String, slCTEStamp As String, slForm_Module As String, tlCTE() As CTE) As Integer
'
'   slGetType(I)- C=Current; H=History, B=Both
'   slType(I)- T1=Title 1; T2=Title 2; DH=Day Header Comment
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim llLoop As Long
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "cte.eng") & slGetType & slType
    
    On Error GoTo gGetTypeOfRecs_CTE_CommtsTitleErr
    ilRet = 0
    ilLowLimit = LBound(tlCTE)
    If ilRet <> 0 Then
        slCTEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slCTEStamp <> "") Then
        If slStamp = slCTEStamp Then
            gGetTypeOfRecs_CTE_CommtsTitle = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM CTE_Commts_And_Title"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM CTE_Commts_And_Title WHERE cteCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM CTE_Commts_And_Title WHERE cteCurrent = 'Y'"
    End If
    If slGetType = "B" Then
        If slType = "T1" Or slType = "T2" Or slType = "DH" Then
            sgSQLQuery = sgSQLQuery & " Where cteType = '" & slType & "'"
        End If
    Else
        If slType = "T1" Or slType = "T2" Or slType = "DH" Then
            sgSQLQuery = sgSQLQuery & " And cteType = '" & slType & "'"
        End If
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY cteCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlCTE(0 To 0) As CTE
    llUpper = 0
    While Not rst.EOF
        tlCTE(llUpper).lCode = rst!cteCode
        tlCTE(llUpper).sType = rst!cteType
        tlCTE(llUpper).sComment = rst!cteComment
        tlCTE(llUpper).sState = rst!cteState
        tlCTE(llUpper).sUsedFlag = rst!cteUsedFlag
        tlCTE(llUpper).iVersion = rst!cteVersion
        tlCTE(llUpper).lOrigCteCode = rst!cteOrigCteCode
        tlCTE(llUpper).sCurrent = rst!cteCurrent
        tlCTE(llUpper).sEnteredDate = Format$(rst!cteEnteredDate, sgShowDateForm)
        tlCTE(llUpper).sEnteredTime = Format$(rst!cteEnteredTime, sgShowTimeWSecForm)
        tlCTE(llUpper).iUieCode = rst!cteUieCode
        tlCTE(llUpper).sUnused = rst!cteUnused
        llUpper = llUpper + 1
        ReDim Preserve tlCTE(0 To llUpper) As CTE
        rst.MoveNext
    Wend
    '7/11/11: Make T2 work like T1
    'If slType = "T2" Then
    '    ReDim tgCurr2CTE_Name(0 To llUpper) As CTESORT
    '    For llLoop = 0 To llUpper Step 1
    '        tgCurr2CTE_Name(llLoop).sKey = tlCTE(llLoop).sName
    '        tgCurr2CTE_Name(llLoop).lCode = tlCTE(llLoop).lCode
    '    Next llLoop
    '    If llUpper > 0 Then
    '        ArraySortTyp fnAV(tgCurr2CTE_Name(), 0), llUpper, 0, LenB(tgCurr2CTE_Name(0)), 0, LenB(tgCurr2CTE_Name(0).sKey), 0
    '    End If
    'End If
    slCTEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_CTE_CommtsTitle = True
    Exit Function
    
gGetTypeOfRecs_CTE_CommtsTitleErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_CTE_CommtsTitle = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_DEECTE_ForDHE(llDheCode As Long, slForm_Module As String, tlDeeCte() As DEECTE) As Integer
'
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim llLoop As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT deeCode, cteCode, cteComment FROM DEE_Day_Event_Info INNER JOIN CTE_Commts_And_Title On dee1CTECode = cteCode WHERE deeDheCode = " & llDheCode & " ORDER BY cteComment, cteCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlDeeCte(0 To 0) As DEECTE
    llUpper = 0
    While Not rst.EOF
        tlDeeCte(llUpper).lDeeCode = rst!deeCode
        tlDeeCte(llUpper).lCteCode = rst!cteCode
        tlDeeCte(llUpper).sComment = rst!cteComment
        llUpper = llUpper + 1
        ReDim Preserve tlDeeCte(0 To llUpper) As DEECTE
        rst.MoveNext
    Wend
    rst.Close
    gGetTypeOfRecs_DEECTE_ForDHE = True
    Exit Function
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_DEECTE_ForDHE = False
    Exit Function

End Function
Public Function gGetTypeOfRecs_DHE_DayHeaderInfo(slGetType As String, slType As String, slDHEStamp As String, slForm_Module As String, tlDHE() As DHE) As Integer
'
'   slGetType(I)- C=Current; H=History, B=Both
'   slType(I)- L=Library; T=Template
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "dhe.eng") & slGetType & slType
    
    On Error GoTo gGetTypeOfRecs_DHE_DayHeaderInfoErr
    ilRet = 0
    ilLowLimit = LBound(tlDHE)
    If ilRet <> 0 Then
        slDHEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slDHEStamp <> "") Then
        If slStamp = slDHEStamp Then
            gGetTypeOfRecs_DHE_DayHeaderInfo = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info WHERE dheCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info WHERE dheCurrent = 'Y'"
    End If
    If slGetType = "B" Then
        If slType = "L" Or slType = "T" Then
            sgSQLQuery = sgSQLQuery & " Where dheType = '" & slType & "'" & " ORDER BY dheOrigDheCode, dheVersion"
        End If
    Else
        If slType = "L" Or slType = "T" Then
            sgSQLQuery = sgSQLQuery & " And dheType = '" & slType & "'"
        End If
    End If
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlDHE(0 To 0) As DHE
    llUpper = 0
    While Not rst.EOF
        tlDHE(llUpper).lCode = rst!dheCode
        tlDHE(llUpper).sType = rst!dheType
        tlDHE(llUpper).lDneCode = rst!dheDneCode
        tlDHE(llUpper).lDseCode = rst!dheDseCode
        tlDHE(llUpper).sStartTime = Format$(rst!dheStartTime, sgShowTimeWSecForm)
        tlDHE(llUpper).lLength = rst!dheLength
        tlDHE(llUpper).sHours = rst!dheHours
        tlDHE(llUpper).sStartDate = Format$(rst!dheStartDate, sgShowDateForm)
        tlDHE(llUpper).sEndDate = Format$(rst!dheEndDate, sgShowDateForm)
        tlDHE(llUpper).sDays = rst!dheDays
        tlDHE(llUpper).lCteCode = rst!dheCteCode
        tlDHE(llUpper).sState = rst!dheState
        tlDHE(llUpper).sUsedFlag = rst!dheUsedFlag
        tlDHE(llUpper).iVersion = rst!dheVersion
        tlDHE(llUpper).lOrigDHECode = rst!dheOrigDheCode
        tlDHE(llUpper).sCurrent = rst!dheCurrent
        tlDHE(llUpper).sEnteredDate = Format$(rst!dheEnteredDate, sgShowDateForm)
        tlDHE(llUpper).sEnteredTime = Format$(rst!dheEnteredTime, sgShowTimeWSecForm)
        tlDHE(llUpper).iUieCode = rst!dheUieCode
        tlDHE(llUpper).sIgnoreConflicts = rst!dheIgnoreConflicts
        tlDHE(llUpper).sBusNames = rst!dheBusNames
        tlDHE(llUpper).sUnused = rst!dheUnused
        llUpper = llUpper + 1
        ReDim Preserve tlDHE(0 To llUpper) As DHE
        rst.MoveNext
    Wend
    slDHEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_DHE_DayHeaderInfo = True
    Exit Function
    
gGetTypeOfRecs_DHE_DayHeaderInfoErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_DHE_DayHeaderInfo = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_DHE_DayHeaderInfoForLibByDate(slGetType As String, slState As String, slDate As String, slDHEStamp As String, slForm_Module As String, tlDHE() As DHE) As Integer
'
'   slGetType(I)- C=Current; H=History, B=Both
'   slState(I)- A=Active; D=Dormant; L=Limbo; AL=Active and Limbo; ALD=Active; Limbo and Dormant
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "dhe.eng") & slGetType & slDate
    
    On Error GoTo gGetTypeOfRecs_DHE_DayHeaderInfoForLibByDateErr
    ilRet = 0
    ilLowLimit = LBound(tlDHE)
    If ilRet <> 0 Then
        slDHEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slDHEStamp <> "") Then
        If slStamp = slDHEStamp Then
            gGetTypeOfRecs_DHE_DayHeaderInfoForLibByDate = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info WHERE dheCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info WHERE dheCurrent = 'Y'"
    End If
    If slGetType = "B" Then
        sgSQLQuery = sgSQLQuery & " Where dheType = '" & "L" & "'"
    Else
        sgSQLQuery = sgSQLQuery & " And dheType = '" & "L" & "'"
    End If
    'If slState = "D" Then
    '    sgSQLQuery = sgSQLQuery & " And dheState = '" & "D" & "'"
    'ElseIf slState <> "B" Then
    '    sgSQLQuery = sgSQLQuery & " And dheState <> '" & "D" & "'"
    'End If
    If InStr(1, slState, "A", vbTextCompare) > 0 Then
        If InStr(1, slState, "L", vbTextCompare) > 0 Then
            If InStr(1, slState, "D", vbTextCompare) <= 0 Then
                sgSQLQuery = sgSQLQuery & " And dheState <> '" & "D" & "'"
            End If
        Else
            sgSQLQuery = sgSQLQuery & " And dheState <> '" & "L" & "'"
            If InStr(1, slState, "D", vbTextCompare) <= 0 Then
                sgSQLQuery = sgSQLQuery & " And dheState <> '" & "D" & "'"
            End If
        End If
    Else
        sgSQLQuery = sgSQLQuery & " And dheState <> '" & "A" & "'"
        If InStr(1, slState, "L", vbTextCompare) > 0 Then
            If InStr(1, slState, "D", vbTextCompare) <= 0 Then
                sgSQLQuery = sgSQLQuery & " And dheState <> '" & "D" & "'"
            End If
        Else
            sgSQLQuery = sgSQLQuery & " And dheState <> '" & "L" & "'"
            If InStr(1, slState, "D", vbTextCompare) <= 0 Then
                sgSQLQuery = sgSQLQuery & " And dheState <> '" & "D" & "'"
            End If
        End If
    End If
    sgSQLQuery = sgSQLQuery & " AND dheEndDate >= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "'"
    sgSQLQuery = sgSQLQuery & " AND dheStartDate <= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "'"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlDHE(0 To 0) As DHE
    llUpper = 0
    While Not rst.EOF
        tlDHE(llUpper).lCode = rst!dheCode
        tlDHE(llUpper).sType = rst!dheType
        tlDHE(llUpper).lDneCode = rst!dheDneCode
        tlDHE(llUpper).lDseCode = rst!dheDseCode
        tlDHE(llUpper).sStartTime = Format$(rst!dheStartTime, sgShowTimeWSecForm)
        tlDHE(llUpper).lLength = rst!dheLength
        tlDHE(llUpper).sHours = rst!dheHours
        tlDHE(llUpper).sStartDate = Format$(rst!dheStartDate, sgShowDateForm)
        tlDHE(llUpper).sEndDate = Format$(rst!dheEndDate, sgShowDateForm)
        tlDHE(llUpper).sDays = rst!dheDays
        tlDHE(llUpper).lCteCode = rst!dheCteCode
        tlDHE(llUpper).sState = rst!dheState
        tlDHE(llUpper).sUsedFlag = rst!dheUsedFlag
        tlDHE(llUpper).iVersion = rst!dheVersion
        tlDHE(llUpper).lOrigDHECode = rst!dheOrigDheCode
        tlDHE(llUpper).sCurrent = rst!dheCurrent
        tlDHE(llUpper).sEnteredDate = Format$(rst!dheEnteredDate, sgShowDateForm)
        tlDHE(llUpper).sEnteredTime = Format$(rst!dheEnteredTime, sgShowTimeWSecForm)
        tlDHE(llUpper).iUieCode = rst!dheUieCode
        tlDHE(llUpper).sIgnoreConflicts = rst!dheIgnoreConflicts
        tlDHE(llUpper).sBusNames = rst!dheBusNames
        tlDHE(llUpper).sUnused = rst!dheUnused
        llUpper = llUpper + 1
        ReDim Preserve tlDHE(0 To llUpper) As DHE
        rst.MoveNext
    Wend
    slDHEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_DHE_DayHeaderInfoForLibByDate = True
    Exit Function
    
gGetTypeOfRecs_DHE_DayHeaderInfoForLibByDateErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_DHE_DayHeaderInfoForLibByDate = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_DHE_DayHeaderInfoForLibByRange(slGetType As String, slState As String, slStartDate As String, slEndDate As String, slDHEStamp As String, slForm_Module As String, tlDHE() As DHE) As Integer
'
'   slGetType(I)- C=Current; H=History, B=Both
'   slState(I)- A=Active; D=Dormant; L=Limbo; AL=Active and Limbo; ALD=Active; Limbo and Dormant
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "dhe.eng") & slGetType & slStartDate & slEndDate
    
    On Error GoTo gGetTypeOfRecs_DHE_DayHeaderInfoForLibByRangeErr
    ilRet = 0
    ilLowLimit = LBound(tlDHE)
    If ilRet <> 0 Then
        slDHEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slDHEStamp <> "") Then
        If slStamp = slDHEStamp Then
            gGetTypeOfRecs_DHE_DayHeaderInfoForLibByRange = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info WHERE dheCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info WHERE dheCurrent = 'Y'"
    End If
    If slGetType = "B" Then
        sgSQLQuery = sgSQLQuery & " Where dheType = '" & "L" & "'"
    Else
        sgSQLQuery = sgSQLQuery & " And dheType = '" & "L" & "'"
    End If
    'If slState = "D" Then
    '    sgSQLQuery = sgSQLQuery & " And dheState = '" & "D" & "'"
    'ElseIf slState <> "B" Then
    '    sgSQLQuery = sgSQLQuery & " And dheState <> '" & "D" & "'"
    'End If
    If InStr(1, slState, "A", vbTextCompare) > 0 Then
        If InStr(1, slState, "L", vbTextCompare) > 0 Then
            If InStr(1, slState, "D", vbTextCompare) <= 0 Then
                sgSQLQuery = sgSQLQuery & " And dheState <> '" & "D" & "'"
            End If
        Else
            sgSQLQuery = sgSQLQuery & " And dheState <> '" & "L" & "'"
            If InStr(1, slState, "D", vbTextCompare) <= 0 Then
                sgSQLQuery = sgSQLQuery & " And dheState <> '" & "D" & "'"
            End If
        End If
    Else
        sgSQLQuery = sgSQLQuery & " And dheState <> '" & "A" & "'"
        If InStr(1, slState, "L", vbTextCompare) > 0 Then
            If InStr(1, slState, "D", vbTextCompare) <= 0 Then
                sgSQLQuery = sgSQLQuery & " And dheState <> '" & "D" & "'"
            End If
        Else
            sgSQLQuery = sgSQLQuery & " And dheState <> '" & "L" & "'"
            If InStr(1, slState, "D", vbTextCompare) <= 0 Then
                sgSQLQuery = sgSQLQuery & " And dheState <> '" & "D" & "'"
            End If
        End If
    End If
    sgSQLQuery = sgSQLQuery & " AND dheEndDate >= '" & Format$(gAdjYear(slStartDate), sgSQLDateForm) & "'"
    sgSQLQuery = sgSQLQuery & " AND dheStartDate <= '" & Format$(gAdjYear(slEndDate), sgSQLDateForm) & "'"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlDHE(0 To 0) As DHE
    llUpper = 0
    While Not rst.EOF
        tlDHE(llUpper).lCode = rst!dheCode
        tlDHE(llUpper).sType = rst!dheType
        tlDHE(llUpper).lDneCode = rst!dheDneCode
        tlDHE(llUpper).lDseCode = rst!dheDseCode
        tlDHE(llUpper).sStartTime = Format$(rst!dheStartTime, sgShowTimeWSecForm)
        tlDHE(llUpper).lLength = rst!dheLength
        tlDHE(llUpper).sHours = rst!dheHours
        tlDHE(llUpper).sStartDate = Format$(rst!dheStartDate, sgShowDateForm)
        tlDHE(llUpper).sEndDate = Format$(rst!dheEndDate, sgShowDateForm)
        tlDHE(llUpper).sDays = rst!dheDays
        tlDHE(llUpper).lCteCode = rst!dheCteCode
        tlDHE(llUpper).sState = rst!dheState
        tlDHE(llUpper).sUsedFlag = rst!dheUsedFlag
        tlDHE(llUpper).iVersion = rst!dheVersion
        tlDHE(llUpper).lOrigDHECode = rst!dheOrigDheCode
        tlDHE(llUpper).sCurrent = rst!dheCurrent
        tlDHE(llUpper).sEnteredDate = Format$(rst!dheEnteredDate, sgShowDateForm)
        tlDHE(llUpper).sEnteredTime = Format$(rst!dheEnteredTime, sgShowTimeWSecForm)
        tlDHE(llUpper).iUieCode = rst!dheUieCode
        tlDHE(llUpper).sIgnoreConflicts = rst!dheIgnoreConflicts
        tlDHE(llUpper).sBusNames = rst!dheBusNames
        tlDHE(llUpper).sUnused = rst!dheUnused
        llUpper = llUpper + 1
        ReDim Preserve tlDHE(0 To llUpper) As DHE
        rst.MoveNext
    Wend
    slDHEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_DHE_DayHeaderInfoForLibByRange = True
    Exit Function
    
gGetTypeOfRecs_DHE_DayHeaderInfoForLibByRangeErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    rst.Close
    gGetTypeOfRecs_DHE_DayHeaderInfoForLibByRange = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_DHETSE_DayHeaderInfoForTempByRange(slGetType As String, slStartDate As String, slEndDate As String, slDHEStamp As String, slForm_Module As String, tlDHETSE() As DHETSE) As Integer
'
'   slGetType(I)- C=Current; H=History, B=Both
'
'Note:  Start and End Date set to Template Log Date.  Days is also set based on Log Date
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim ilHour As Integer
    Dim slHours As String
    Dim ilLoop As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "dhe.eng") & slGetType & slStartDate & slEndDate
    
    On Error GoTo gGetTypeOfRecs_DHETSE_DayHeaderInfoForTempByRangeErr
    ilRet = 0
    ilLowLimit = LBound(tlDHETSE)
    If ilRet <> 0 Then
        slDHEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slDHEStamp <> "") Then
        If slStamp = slDHEStamp Then
            gGetTypeOfRecs_DHETSE_DayHeaderInfoForTempByRange = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info, TSE_Template_Schd WHERE dheCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info, TSE_Template_Schd WHERE dheCurrent = 'Y'"
    End If
    If slGetType = "B" Then
        sgSQLQuery = sgSQLQuery & " Where dheType = '" & "T" & "'"
    Else
        sgSQLQuery = sgSQLQuery & " And dheType = '" & "T" & "'"
    End If
    sgSQLQuery = sgSQLQuery & " AND dheCode = tseDheCode"
    sgSQLQuery = sgSQLQuery & " AND tseLogDate >= '" & Format$(gAdjYear(slStartDate), sgSQLDateForm) & "'"
    sgSQLQuery = sgSQLQuery & " AND tseLogDate <= '" & Format$(gAdjYear(slEndDate), sgSQLDateForm) & "'"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlDHETSE(0 To 0) As DHETSE
    llUpper = 0
    While Not rst.EOF
        tlDHETSE(llUpper).tDHE.lCode = rst!dheCode
        tlDHETSE(llUpper).tDHE.sType = rst!dheType
        tlDHETSE(llUpper).tDHE.lDneCode = rst!dheDneCode
        tlDHETSE(llUpper).tDHE.lDseCode = rst!dheDseCode
        tlDHETSE(llUpper).tDHE.sStartTime = Format$(rst!dheStartTime, sgShowTimeWSecForm)
        tlDHETSE(llUpper).tDHE.lLength = rst!dheLength
        tlDHETSE(llUpper).tDHE.sHours = rst!dheHours
'        'Displace hours
'        ilHour = Hour(Format$(rst!tseStartTime, sgShowTimeWSecForm))
'        If ilHour = 0 Then
'            tlDHETSE(llUpper).sHours = rst!dheHours
'        Else
'            slHours = rst!dheHours
'            tlDHETSE(llUpper).sHours = String(24, "N")
'            For ilLoop = 0 To 23 Step 1
'                Mid$(tlDHETSE(llUpper).sHours, ilHour, 1) = Mid$(slHours, ilLoop, 1)
'                ilHour = ilHour + 1
'                If ilHour > 23 Then
'                    Exit For
'                End If
'            Next ilLoop
'        End If
        tlDHETSE(llUpper).tDHE.sStartDate = Format$(rst!tseLogDate, sgShowDateForm)   'Format$(rst!dheStartDate, sgShowDateForm)
        tlDHETSE(llUpper).tDHE.sEndDate = Format$(rst!tseLogDate, sgShowDateForm)   'Format$(rst!dheEndDate, sgShowDateForm)
'        tlDHETSE(llUpper).sDays = String(7, "N")
'        Select Case Weekday(tlDHETSE(llUpper).sStartDate)
'            Case vbMonday
'                Mid(tlDHETSE(llUpper).sDays, 1, 1) = "Y"
'            Case vbTuesday
'                Mid(tlDHETSE(llUpper).sDays, 2, 1) = "Y"
'            Case vbWednesday
'                Mid(tlDHETSE(llUpper).sDays, 3, 1) = "Y"
'            Case vbThursday
'                Mid(tlDHETSE(llUpper).sDays, 4, 1) = "Y"
'            Case vbFriday
'                Mid(tlDHETSE(llUpper).sDays, 5, 1) = "Y"
'            Case vbSaturday
'                Mid(tlDHETSE(llUpper).sDays, 6, 1) = "Y"
'            Case vbSunday
'                Mid(tlDHETSE(llUpper).sDays, 7, 1) = "Y"
'        End Select
        tlDHETSE(llUpper).tDHE.sDays = rst!dheDays
        tlDHETSE(llUpper).tDHE.lCteCode = rst!dheCteCode
'        If rst!dheState <> "D" Then
'            tlDHETSE(llUpper).sState = rst!tseState    'rst!dheState
'        Else
'            tlDHETSE(llUpper).sState = rst!dheState
'        End If
        tlDHETSE(llUpper).tDHE.sState = rst!dheState
        tlDHETSE(llUpper).tDHE.sUsedFlag = rst!dheUsedFlag
        tlDHETSE(llUpper).tDHE.iVersion = rst!dheVersion
        tlDHETSE(llUpper).tDHE.lOrigDHECode = rst!dheOrigDheCode
        tlDHETSE(llUpper).tDHE.sCurrent = rst!dheCurrent
        tlDHETSE(llUpper).tDHE.sEnteredDate = Format$(rst!dheEnteredDate, sgShowDateForm)
        tlDHETSE(llUpper).tDHE.sEnteredTime = Format$(rst!dheEnteredTime, sgShowTimeWSecForm)
        tlDHETSE(llUpper).tDHE.iUieCode = rst!dheUieCode
        tlDHETSE(llUpper).tDHE.sIgnoreConflicts = rst!dheIgnoreConflicts
        tlDHETSE(llUpper).tDHE.sBusNames = rst!dheBusNames
        tlDHETSE(llUpper).tDHE.sUnused = rst!dheUnused
        tlDHETSE(llUpper).tTSE.lCode = rst!tseCode
        tlDHETSE(llUpper).tTSE.lDheCode = rst!tseDheCode
        tlDHETSE(llUpper).tTSE.iBdeCode = rst!tseBdeCode
        tlDHETSE(llUpper).tTSE.sLogDate = Format$(rst!tseLogDate, sgShowDateForm)
        tlDHETSE(llUpper).tTSE.sStartTime = Format$(rst!tseStartTime, sgShowTimeWSecForm)
        tlDHETSE(llUpper).tTSE.sDescription = rst!tseDescription
        tlDHETSE(llUpper).tTSE.sState = rst!tseState
        tlDHETSE(llUpper).tTSE.lCteCode = rst!tseCteCode
        tlDHETSE(llUpper).tTSE.iVersion = rst!tseVersion
        tlDHETSE(llUpper).tTSE.lOrigTseCode = rst!tseOrigTseCode
        tlDHETSE(llUpper).tTSE.sCurrent = rst!tseCurrent
        tlDHETSE(llUpper).tTSE.sEnteredDate = Format$(rst!tseEnteredDate, sgShowDateForm)
        tlDHETSE(llUpper).tTSE.sEnteredTime = Format$(rst!tseEnteredTime, sgShowTimeWSecForm)
        tlDHETSE(llUpper).tTSE.iUieCode = rst!tseUieCode
        tlDHETSE(llUpper).tTSE.sUnused = rst!tseUnused
        llUpper = llUpper + 1
        ReDim Preserve tlDHETSE(0 To llUpper) As DHETSE
        rst.MoveNext
    Wend
    slDHEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_DHETSE_DayHeaderInfoForTempByRange = True
    Exit Function
    
gGetTypeOfRecs_DHETSE_DayHeaderInfoForTempByRangeErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    rst.Close
    gGetTypeOfRecs_DHETSE_DayHeaderInfoForTempByRange = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_DNE_DayName(slGetType As String, slType As String, slDNEStamp As String, slForm_Module As String, tlDNE() As DNE) As Integer
'
'   slGetType(I)- C=Current; H=History, B=Both
'   slType(I)- L=Library; T=Template
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "dne.eng") & slGetType & slType
    
    On Error GoTo gGetTypeOfRecs_DNE_DayNameErr
    ilRet = 0
    ilLowLimit = LBound(tlDNE)
    If ilRet <> 0 Then
        slDNEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slDNEStamp <> "") Then
        If slStamp = slDNEStamp Then
            gGetTypeOfRecs_DNE_DayName = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM DNE_Day_Name"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM DNE_Day_Name WHERE dneCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM DNE_Day_Name WHERE dneCurrent = 'Y'"
    End If
    If slGetType = "B" Then
        If slType = "L" Or slType = "T" Then
            sgSQLQuery = sgSQLQuery & " Where dneType = '" & slType & "'"
        End If
    Else
        If slType = "L" Or slType = "T" Then
            sgSQLQuery = sgSQLQuery & " And dneType = '" & slType & "'"
        End If
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY dneCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlDNE(0 To 0) As DNE
    llUpper = 0
    While Not rst.EOF
        tlDNE(llUpper).lCode = rst!dneCode
        tlDNE(llUpper).sType = rst!dneType
        tlDNE(llUpper).sName = rst!dneName
        tlDNE(llUpper).sDescription = rst!dneDescription
        tlDNE(llUpper).sState = rst!dneState
        tlDNE(llUpper).sUsedFlag = rst!dneUsedFlag
        tlDNE(llUpper).iVersion = rst!dneVersion
        tlDNE(llUpper).lOrigDneCode = rst!dneOrigDneCode
        tlDNE(llUpper).sCurrent = rst!dneCurrent
        tlDNE(llUpper).sEnteredDate = Format$(rst!dneEnteredDate, sgShowDateForm)
        tlDNE(llUpper).sEnteredTime = Format$(rst!dneEnteredTime, sgShowTimeWSecForm)
        tlDNE(llUpper).iUieCode = rst!dneUieCode
        tlDNE(llUpper).sUnused = rst!dneUnused
        llUpper = llUpper + 1
        ReDim Preserve tlDNE(0 To llUpper) As DNE
        rst.MoveNext
    Wend
    slDNEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_DNE_DayName = True
    Exit Function
    
gGetTypeOfRecs_DNE_DayNameErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_DNE_DayName = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_DSE_DaySubName(slGetType As String, slDSEStamp As String, slForm_Module As String, tlDSE() As DSE) As Integer
'
'   slGetType(I)- C=Current; H=History, B=Both
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "dse.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_DSE_DaySubNameErr
    ilRet = 0
    ilLowLimit = LBound(tlDSE)
    If ilRet <> 0 Then
        slDSEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slDSEStamp <> "") Then
        If slStamp = slDSEStamp Then
            gGetTypeOfRecs_DSE_DaySubName = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM DSE_Day_SubName"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM DSE_Day_SubName WHERE dseCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM DSE_Day_SubName WHERE dseCurrent = 'Y'"
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY dseCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlDSE(0 To 0) As DSE
    llUpper = 0
    While Not rst.EOF
        tlDSE(llUpper).lCode = rst!dseCode
        tlDSE(llUpper).sName = rst!dseName
        tlDSE(llUpper).sDescription = rst!dseDescription
        tlDSE(llUpper).sState = rst!dseState
        tlDSE(llUpper).sUsedFlag = rst!dseUsedFlag
        tlDSE(llUpper).iVersion = rst!dseVersion
        tlDSE(llUpper).lOrigDseCode = rst!dseOrigDseCode
        tlDSE(llUpper).sCurrent = rst!dseCurrent
        tlDSE(llUpper).sEnteredDate = Format$(rst!dseEnteredDate, sgShowDateForm)
        tlDSE(llUpper).sEnteredTime = Format$(rst!dseEnteredTime, sgShowTimeWSecForm)
        tlDSE(llUpper).iUieCode = rst!dseUieCode
        tlDSE(llUpper).sUnused = rst!dseUnused
        llUpper = llUpper + 1
        ReDim Preserve tlDSE(0 To llUpper) As DSE
        rst.MoveNext
    Wend
    slDSEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_DSE_DaySubName = True
    Exit Function
    
gGetTypeOfRecs_DSE_DaySubNameErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_DSE_DaySubName = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_ETE_EventType(slGetType As String, slETEStamp As String, slForm_Module As String, tlETE() As ETE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "ete.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_ETE_EventTypeErr
    ilRet = 0
    ilLowLimit = LBound(tlETE)
    If ilRet <> 0 Then
        slETEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slETEStamp <> "") Then
        If slStamp = slETEStamp Then
            gGetTypeOfRecs_ETE_EventType = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM ETE_Event_Type"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM ETE_Event_Type WHERE eteCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM ETE_Event_Type WHERE eteCurrent = 'Y'"
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY eteCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlETE(0 To 0) As ETE
    ilUpper = 0
    While Not rst.EOF
        tlETE(ilUpper).iCode = rst!eteCode
        tlETE(ilUpper).sName = rst!eteName
        tlETE(ilUpper).sDescription = rst!eteDescription
        tlETE(ilUpper).sCategory = rst!eteCategory
        tlETE(ilUpper).sAutoCodeChar = rst!eteAutoCodeChar
        tlETE(ilUpper).sState = rst!eteState
        tlETE(ilUpper).sUsedFlag = rst!eteUsedFlag
        tlETE(ilUpper).iVersion = rst!eteVersion
        tlETE(ilUpper).iOrigEteCode = rst!eteOrigEteCode
        tlETE(ilUpper).sCurrent = rst!eteCurrent
        tlETE(ilUpper).sEnteredDate = Format$(rst!eteEnteredDate, sgShowDateForm)
        tlETE(ilUpper).sEnteredTime = Format$(rst!eteEnteredTime, sgShowTimeWSecForm)
        tlETE(ilUpper).iUieCode = rst!eteUieCode
        tlETE(ilUpper).sUnused = rst!eteUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlETE(0 To ilUpper) As ETE
        rst.MoveNext
    Wend
    slETEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_ETE_EventType = True
    Exit Function
    
gGetTypeOfRecs_ETE_EventTypeErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_ETE_EventType = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_EPE_EventProperties(slGetType As String, slEPEStamp As String, slForm_Module As String, tlEPE() As EPE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "epe.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_EPE_EventPropertiesErr
    ilRet = 0
    ilLowLimit = LBound(tlEPE)
    If ilRet <> 0 Then
        slEPEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slEPEStamp <> "") Then
        If slStamp = slEPEStamp Then
            gGetTypeOfRecs_EPE_EventProperties = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM EPE_Event_Properties, ETE_Event_Type ORDER BY epeEteCode"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM EPE_Event_Properties, ETE_Event_Type WHERE epeEteCode = eteCode and eteCurrent = 'N' ORDER BY epeEteCode"
    Else
        sgSQLQuery = "SELECT * FROM EPE_Event_Properties, ETE_Event_Type WHERE epeEteCode = eteCode and eteCurrent = 'Y' ORDER BY epeEteCode"
    End If
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlEPE(0 To 0) As EPE
    ilUpper = 0
    While Not rst.EOF
        tlEPE(ilUpper).iCode = rst!epeCode
        tlEPE(ilUpper).iEteCode = rst!epeEteCode
        tlEPE(ilUpper).sType = rst!epeType
        tlEPE(ilUpper).sBus = rst!epeBus
        tlEPE(ilUpper).sBusControl = rst!epeBusControl
        tlEPE(ilUpper).sTime = rst!epeTime
        tlEPE(ilUpper).sStartType = rst!epeStartType
        tlEPE(ilUpper).sFixedTime = rst!epeFixedTime
        tlEPE(ilUpper).sEndType = rst!epeEndType
        tlEPE(ilUpper).sDuration = rst!epeDuration
        tlEPE(ilUpper).sMaterialType = rst!epeMaterialType
        tlEPE(ilUpper).sAudioName = rst!epeAudioName
        tlEPE(ilUpper).sAudioItemID = rst!epeAudioItemID
        tlEPE(ilUpper).sAudioISCI = rst!epeAudioISCI
        tlEPE(ilUpper).sAudioControl = rst!epeAudioControl
        tlEPE(ilUpper).sBkupAudioName = rst!epeBkupAudioName
        tlEPE(ilUpper).sBkupAudioControl = rst!epeBkupAudioControl
        tlEPE(ilUpper).sProtAudioName = rst!epeProtAudioName
        tlEPE(ilUpper).sProtAudioItemID = rst!epeProtAudioItemID
        tlEPE(ilUpper).sProtAudioISCI = rst!epeProtAudioISCI
        tlEPE(ilUpper).sProtAudioControl = rst!epeProtAudioControl
        tlEPE(ilUpper).sRelay1 = rst!epeRelay1
        tlEPE(ilUpper).sRelay2 = rst!epeRelay2
        tlEPE(ilUpper).sFollow = rst!epeFollow
        tlEPE(ilUpper).sSilenceTime = rst!epeSilenceTime
        tlEPE(ilUpper).sSilence1 = rst!epeSilence1
        tlEPE(ilUpper).sSilence2 = rst!epeSilence2
        tlEPE(ilUpper).sSilence3 = rst!epeSilence3
        tlEPE(ilUpper).sSilence4 = rst!epeSilence4
        tlEPE(ilUpper).sStartNetcue = rst!epeStartNetcue
        tlEPE(ilUpper).sStopNetcue = rst!epeStopNetcue
        tlEPE(ilUpper).sTitle1 = rst!epeTitle1
        tlEPE(ilUpper).sTitle2 = rst!epeTitle2
        tlEPE(ilUpper).sABCFormat = rst!epeABCFormat
        tlEPE(ilUpper).sABCPgmCode = rst!epeABCPgmCode
        tlEPE(ilUpper).sABCXDSMode = rst!epeABCXDSMode
        tlEPE(ilUpper).sABCRecordItem = rst!epeABCRecordItem
        tlEPE(ilUpper).sUnused = rst!epeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlEPE(0 To ilUpper) As EPE
        rst.MoveNext
    Wend
    slEPEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_EPE_EventProperties = True
    Exit Function
    
gGetTypeOfRecs_EPE_EventPropertiesErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_EPE_EventProperties = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_FNE_FollowName(slGetType As String, slFNEStamp As String, slForm_Module As String, tlFNE() As FNE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "fne.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_FNE_FollowNameErr
    ilRet = 0
    ilLowLimit = LBound(tlFNE)
    If ilRet <> 0 Then
        slFNEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slFNEStamp <> "") Then
        If slStamp = slFNEStamp Then
            gGetTypeOfRecs_FNE_FollowName = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM FNE_Follow_Name"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM FNE_Follow_Name WHERE fneCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM FNE_Follow_Name WHERE fneCurrent = 'Y'"
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY fneCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlFNE(0 To 0) As FNE
    ilUpper = 0
    While Not rst.EOF
        tlFNE(ilUpper).iCode = rst!fneCode
        tlFNE(ilUpper).sName = rst!fneName
        tlFNE(ilUpper).sDescription = rst!fneDescription
        tlFNE(ilUpper).sState = rst!fneState
        tlFNE(ilUpper).sUsedFlag = rst!fneUsedFlag
        tlFNE(ilUpper).iVersion = rst!fneVersion
        tlFNE(ilUpper).iOrigFneCode = rst!fneOrigFneCode
        tlFNE(ilUpper).sCurrent = rst!fneCurrent
        tlFNE(ilUpper).sEnteredDate = Format$(rst!fneEnteredDate, sgShowDateForm)
        tlFNE(ilUpper).sEnteredTime = Format$(rst!fneEnteredTime, sgShowTimeWSecForm)
        tlFNE(ilUpper).iUieCode = rst!fneUieCode
        tlFNE(ilUpper).sUnused = rst!fneUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlFNE(0 To ilUpper) As FNE
        rst.MoveNext
    Wend
    slFNEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_FNE_FollowName = True
    Exit Function
    
gGetTypeOfRecs_FNE_FollowNameErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_FNE_FollowName = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_MTE_MaterialType(slGetType As String, slMTEStamp As String, slForm_Module As String, tlMTE() As MTE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "mte.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_MTE_MaterialTypeErr
    ilRet = 0
    ilLowLimit = LBound(tlMTE)
    If ilRet <> 0 Then
        slMTEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slMTEStamp <> "") Then
        If slStamp = slMTEStamp Then
            gGetTypeOfRecs_MTE_MaterialType = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM MTE_Material_Type"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM MTE_Material_Type WHERE mteCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM MTE_Material_Type WHERE mteCurrent = 'Y'"
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY mteCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlMTE(0 To 0) As MTE
    ilUpper = 0
    While Not rst.EOF
        tlMTE(ilUpper).iCode = rst!mteCode
        tlMTE(ilUpper).sName = rst!mteName
        tlMTE(ilUpper).sDescription = rst!mteDescription
        tlMTE(ilUpper).sState = rst!mteState
        tlMTE(ilUpper).sUsedFlag = rst!mteUsedFlag
        tlMTE(ilUpper).iVersion = rst!mteVersion
        tlMTE(ilUpper).iOrigMteCode = rst!mteOrigmteCode
        tlMTE(ilUpper).sCurrent = rst!mteCurrent
        tlMTE(ilUpper).sEnteredDate = Format$(rst!mteEnteredDate, sgShowDateForm)
        tlMTE(ilUpper).sEnteredTime = Format$(rst!mteEnteredTime, sgShowTimeWSecForm)
        tlMTE(ilUpper).iUieCode = rst!mteUieCode
        tlMTE(ilUpper).sUnused = rst!mteUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlMTE(0 To ilUpper) As MTE
        rst.MoveNext
    Wend
    slMTEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_MTE_MaterialType = True
    Exit Function
    
gGetTypeOfRecs_MTE_MaterialTypeErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_MTE_MaterialType = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_NNE_NetcueName(slGetType As String, slNNEStamp As String, slForm_Module As String, tlNNE() As NNE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim ilLoop As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "nne.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_NNE_NetcueNameErr
    ilRet = 0
    ilLowLimit = LBound(tlNNE)
    If ilRet <> 0 Then
        slNNEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slNNEStamp <> "") Then
        If slStamp = slNNEStamp Then
            gGetTypeOfRecs_NNE_NetcueName = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM NNE_Netcue_Name"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM NNE_Netcue_Name WHERE nneCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM NNE_Netcue_Name WHERE nneCurrent = 'Y'"
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY nneCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlNNE(0 To 0) As NNE
    ilUpper = 0
    While Not rst.EOF
        tlNNE(ilUpper).iCode = rst!nneCode
        tlNNE(ilUpper).sName = rst!nneName
        tlNNE(ilUpper).sDescription = rst!nneDescription
        tlNNE(ilUpper).lDneCode = rst!nneDneCode
        tlNNE(ilUpper).sState = rst!nneState
        tlNNE(ilUpper).sUsedFlag = rst!nneUsedFlag
        tlNNE(ilUpper).iVersion = rst!nneVersion
        tlNNE(ilUpper).iOrigNneCode = rst!nneOrigNneCode
        tlNNE(ilUpper).sCurrent = rst!nneCurrent
        tlNNE(ilUpper).sEnteredDate = Format$(rst!nneEnteredDate, sgShowDateForm)
        tlNNE(ilUpper).sEnteredTime = Format$(rst!nneEnteredTime, sgShowTimeWSecForm)
        tlNNE(ilUpper).iUieCode = rst!nneUieCode
        tlNNE(ilUpper).sUnused = rst!nneUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlNNE(0 To ilUpper) As NNE
        rst.MoveNext
    Wend
    ReDim tgCurrNNE_Name(0 To ilUpper) As NAMESORT
    For ilLoop = 0 To ilUpper Step 1
        tgCurrNNE_Name(ilLoop).sKey = tlNNE(ilLoop).sName
        tgCurrNNE_Name(ilLoop).iCode = tlNNE(ilLoop).iCode
    Next ilLoop
    If ilUpper > 0 Then
        ArraySortTyp fnAV(tgCurrNNE_Name(), 0), ilUpper, 0, LenB(tgCurrNNE_Name(0)), 0, LenB(tgCurrNNE_Name(0).sKey), 0
    End If
    slNNEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_NNE_NetcueName = True
    Exit Function
    
gGetTypeOfRecs_NNE_NetcueNameErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_NNE_NetcueName = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_RNE_RelayName(slGetType As String, slRNEStamp As String, slForm_Module As String, tlRNE() As RNE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim ilLoop As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "rne.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_RNE_RelayNameErr
    ilRet = 0
    ilLowLimit = LBound(tlRNE)
    If ilRet <> 0 Then
        slRNEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slRNEStamp <> "") Then
        If slStamp = slRNEStamp Then
            gGetTypeOfRecs_RNE_RelayName = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM RNE_Relay_Name"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM RNE_Relay_Name WHERE rneCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM RNE_Relay_Name WHERE rneCurrent = 'Y'"
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY rneCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlRNE(0 To 0) As RNE
    ilUpper = 0
    While Not rst.EOF
        tlRNE(ilUpper).iCode = rst!rneCode
        tlRNE(ilUpper).sName = rst!rneName
        tlRNE(ilUpper).sDescription = rst!rneDescription
        tlRNE(ilUpper).sState = rst!rneState
        tlRNE(ilUpper).sUsedFlag = rst!rneUsedFlag
        tlRNE(ilUpper).iVersion = rst!rneVersion
        tlRNE(ilUpper).iOrigRneCode = rst!rneOrigRneCode
        tlRNE(ilUpper).sCurrent = rst!rneCurrent
        tlRNE(ilUpper).sEnteredDate = Format$(rst!rneEnteredDate, sgShowDateForm)
        tlRNE(ilUpper).sEnteredTime = Format$(rst!rneEnteredTime, sgShowTimeWSecForm)
        tlRNE(ilUpper).iUieCode = rst!rneUieCode
        tlRNE(ilUpper).sUnused = rst!rneUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlRNE(0 To ilUpper) As RNE
        rst.MoveNext
    Wend
    ReDim tgCurrRNE_Name(0 To ilUpper) As NAMESORT
    For ilLoop = 0 To ilUpper Step 1
        tgCurrRNE_Name(ilLoop).sKey = tlRNE(ilLoop).sName
        tgCurrRNE_Name(ilLoop).iCode = tlRNE(ilLoop).iCode
    Next ilLoop
    If ilUpper > 0 Then
        ArraySortTyp fnAV(tgCurrRNE_Name(), 0), ilUpper, 0, LenB(tgCurrRNE_Name(0)), 0, LenB(tgCurrRNE_Name(0).sKey), 0
    End If
    slRNEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_RNE_RelayName = True
    Exit Function
    
gGetTypeOfRecs_RNE_RelayNameErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_RNE_RelayName = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_SCE_SilenceChar(slGetType As String, slSCEStamp As String, slForm_Module As String, tlSCE() As SCE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "sce.eng") & slGetType
    
    On Error GoTo gGetTypeOfRecs_SCE_SilenceCharErr
    ilRet = 0
    ilLowLimit = LBound(tlSCE)
    If ilRet <> 0 Then
        slSCEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slSCEStamp <> "") Then
        If slStamp = slSCEStamp Then
            gGetTypeOfRecs_SCE_SilenceChar = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM SCE_Silence_Char"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM SCE_Silence_Char WHERE sceCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM SCE_Silence_Char WHERE sceCurrent = 'Y'"
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY sceCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlSCE(0 To 0) As SCE
    ilUpper = 0
    While Not rst.EOF
        tlSCE(ilUpper).iCode = rst!sceCode
        tlSCE(ilUpper).sAutoChar = rst!sceAutoChar
        tlSCE(ilUpper).sDescription = rst!sceDescription
        tlSCE(ilUpper).sState = rst!sceState
        tlSCE(ilUpper).sUsedFlag = rst!sceUsedFlag
        tlSCE(ilUpper).iVersion = rst!sceVersion
        tlSCE(ilUpper).iOrigSceCode = rst!sceOrigSceCode
        tlSCE(ilUpper).sCurrent = rst!sceCurrent
        tlSCE(ilUpper).sEnteredDate = Format$(rst!sceEnteredDate, sgShowDateForm)
        tlSCE(ilUpper).sEnteredTime = Format$(rst!sceEnteredTime, sgShowTimeWSecForm)
        tlSCE(ilUpper).iUieCode = rst!sceUieCode
        tlSCE(ilUpper).sUnused = rst!sceUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlSCE(0 To ilUpper) As SCE
        rst.MoveNext
    Wend
    slSCEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_SCE_SilenceChar = True
    Exit Function
    
gGetTypeOfRecs_SCE_SilenceCharErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_SCE_SilenceChar = False
    Exit Function

End Function

Public Function gGetTypeOfRecs_SHE_ScheduleHeaderByLoadStatusAndDate(slLoadStatus As String, slDate As String, slForm_Module As String, tlSHE() As SHE) As Integer
    Dim ilUpper As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM SHE_Schedule_Header WHERE sheCurrent = 'Y' AND sheCreateLoad = '" & slLoadStatus & "' AND sheAirDate >= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "'"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlSHE(0 To 0) As SHE
    ilUpper = 0
    While Not rst.EOF
        tlSHE(ilUpper).lCode = rst!sheCode
        tlSHE(ilUpper).iAeeCode = rst!sheAeeCode
        tlSHE(ilUpper).sAirDate = Format$(rst!sheAirDate, sgShowDateForm)
        tlSHE(ilUpper).sLoadedAutoStatus = rst!sheLoadedAutoStatus
        tlSHE(ilUpper).sLoadedAutoDate = Format$(rst!sheLoadedAutoDate, sgShowDateForm)
        tlSHE(ilUpper).iChgSeqNo = rst!sheChgSeqNo
        tlSHE(ilUpper).sAsAirStatus = rst!sheAsAirStatus
        tlSHE(ilUpper).sLoadedAsAirDate = Format$(rst!sheLoadedAsAirDate, sgShowDateForm)
        tlSHE(ilUpper).sLastDateItemChk = Format$(rst!sheLastDateItemChk, sgShowDateForm)
        tlSHE(ilUpper).sCreateLoad = rst!sheCreateLoad
        tlSHE(ilUpper).iVersion = rst!sheVersion
        tlSHE(ilUpper).lOrigSheCode = rst!sheOrigSheCode
        tlSHE(ilUpper).sCurrent = rst!sheCurrent
        tlSHE(ilUpper).sEnteredDate = Format$(rst!sheEnteredDate, sgShowDateForm)
        tlSHE(ilUpper).sEnteredTime = Format$(rst!sheEnteredTime, sgShowTimeWSecForm)
        tlSHE(ilUpper).iUieCode = rst!sheUieCode
        tlSHE(ilUpper).sConflictExist = rst!sheConflictExist
        tlSHE(ilUpper).sSpotMergeStatus = rst!sheSpotMergeStatus
        tlSHE(ilUpper).sLoadStatus = rst!sheLoadStatus
        tlSHE(ilUpper).sUnused = rst!sheUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlSHE(0 To ilUpper) As SHE
        rst.MoveNext
    Wend
    rst.Close
    gGetTypeOfRecs_SHE_ScheduleHeaderByLoadStatusAndDate = True
    Exit Function
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_SHE_ScheduleHeaderByLoadStatusAndDate = False
    Exit Function
End Function

Public Function gGetTypeOfRecs_SHE_ScheduleHeaderByDate(slDate As String, slForm_Module As String, tlSHE() As SHE) As Integer
    Dim ilUpper As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM SHE_Schedule_Header WHERE sheCurrent = 'Y' AND sheAirDate >= '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "'"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlSHE(0 To 0) As SHE
    ilUpper = 0
    While Not rst.EOF
        tlSHE(ilUpper).lCode = rst!sheCode
        tlSHE(ilUpper).iAeeCode = rst!sheAeeCode
        tlSHE(ilUpper).sAirDate = Format$(rst!sheAirDate, sgShowDateForm)
        tlSHE(ilUpper).sLoadedAutoStatus = rst!sheLoadedAutoStatus
        tlSHE(ilUpper).sLoadedAutoDate = Format$(rst!sheLoadedAutoDate, sgShowDateForm)
        tlSHE(ilUpper).iChgSeqNo = rst!sheChgSeqNo
        tlSHE(ilUpper).sAsAirStatus = rst!sheAsAirStatus
        tlSHE(ilUpper).sLoadedAsAirDate = Format$(rst!sheLoadedAsAirDate, sgShowDateForm)
        tlSHE(ilUpper).sLastDateItemChk = Format$(rst!sheLastDateItemChk, sgShowDateForm)
        tlSHE(ilUpper).sCreateLoad = rst!sheCreateLoad
        tlSHE(ilUpper).iVersion = rst!sheVersion
        tlSHE(ilUpper).lOrigSheCode = rst!sheOrigSheCode
        tlSHE(ilUpper).sCurrent = rst!sheCurrent
        tlSHE(ilUpper).sEnteredDate = Format$(rst!sheEnteredDate, sgShowDateForm)
        tlSHE(ilUpper).sEnteredTime = Format$(rst!sheEnteredTime, sgShowTimeWSecForm)
        tlSHE(ilUpper).iUieCode = rst!sheUieCode
        tlSHE(ilUpper).sConflictExist = rst!sheConflictExist
        tlSHE(ilUpper).sSpotMergeStatus = rst!sheSpotMergeStatus
        tlSHE(ilUpper).sLoadStatus = rst!sheLoadStatus
        tlSHE(ilUpper).sUnused = rst!sheUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlSHE(0 To ilUpper) As SHE
        rst.MoveNext
    Wend
    rst.Close
    gGetTypeOfRecs_SHE_ScheduleHeaderByDate = True
    Exit Function
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_SHE_ScheduleHeaderByDate = False
    Exit Function
End Function

Public Function gGetRecs_SGE_SiteGenSchd(slSGEStamp As String, ilSoeCode As Integer, slForm_Module As String, tlSGE() As SGE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "sge.eng") & ilSoeCode
    
    On Error GoTo gGetRecs_SGE_SiteGenSchdErr
    ilRet = 0
    ilLowLimit = LBound(tlSGE)
    If ilRet <> 0 Then
        slSGEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slSGEStamp <> "") Then
        If slStamp = slSGEStamp Then
            gGetRecs_SGE_SiteGenSchd = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM SGE_Site_Gen_Schd Where sgeSoeCode = " & ilSoeCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlSGE(0 To 0) As SGE
    ilUpper = 0
    While Not rst.EOF
        tlSGE(ilUpper).iCode = rst!sgeCode
        tlSGE(ilUpper).iSoeCode = rst!sgeSoeCode
        tlSGE(ilUpper).sType = rst!sgeType
        tlSGE(ilUpper).sSubType = rst!sgeSubType
        If (tlSGE(ilUpper).sSubType <> "P") And (tlSGE(ilUpper).sSubType <> "T") Then
            tlSGE(ilUpper).sSubType = "P"
        End If
        tlSGE(ilUpper).iGenMo = rst!sgeGenMo
        tlSGE(ilUpper).iGenTu = rst!sgeGenTu
        tlSGE(ilUpper).iGenWe = rst!sgeGenWe
        tlSGE(ilUpper).iGenTh = rst!sgeGenTh
        tlSGE(ilUpper).iGenFr = rst!sgeGenFr
        tlSGE(ilUpper).iGenSa = rst!sgeGenSa
        tlSGE(ilUpper).iGenSu = rst!sgeGenSu
        tlSGE(ilUpper).sGenTime = Format$(rst!sgeGenTime, sgShowTimeWSecForm)
        tlSGE(ilUpper).sPurgeAfterGen = rst!sgePurgeAfterGen
        tlSGE(ilUpper).sPurgeTime = Format$(rst!sgePurgeTime, sgShowTimeWSecForm)
        tlSGE(ilUpper).lAlertInterval = rst!sgeAlertInterval
        tlSGE(ilUpper).sUnused = ""
        ilUpper = ilUpper + 1
        ReDim Preserve tlSGE(0 To ilUpper) As SGE
        rst.MoveNext
    Wend
    slSGEStamp = slStamp
    rst.Close
    gGetRecs_SGE_SiteGenSchd = True
    Exit Function
    
gGetRecs_SGE_SiteGenSchdErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_SGE_SiteGenSchd = False
    Exit Function

End Function



Public Function gGetRec_SOE_SiteOption(ilCode As Integer, slForm_Module As String, tlSOE As SOE) As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM SOE_Site_Option WHERE soeCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlSOE.iCode = rst!soeCode
        tlSOE.sClientName = rst!soeClientName
        tlSOE.sAddr1 = rst!soeAddr1
        tlSOE.sAddr2 = rst!soeAddr2
        tlSOE.sAddr3 = rst!soeAddr3
        tlSOE.sPhone = rst!soePhone
        tlSOE.sFax = rst!soeFax
        tlSOE.iDaysRetainAsAir = rst!soeDaysRetainAsAir
        tlSOE.lChgInterval = rst!soeChgInterval
        tlSOE.sMergeDateFormat = rst!soeMergeDateFormat
        tlSOE.sMergeTimeFormat = rst!soeMergeTimeFormat
        tlSOE.sMergeFileFormat = rst!soeMergeFileFormat
        tlSOE.sMergeFileExt = rst!soeMergeFileExt
        tlSOE.sMergeStartTime = Format$(rst!soeMergeStartTime, sgShowTimeWSecForm)
        tlSOE.sMergeEndTime = Format$(rst!soeMergeEndTime, sgShowTimeWSecForm)
        tlSOE.iMergeChkInterval = rst!soeMergeChkInterval
        tlSOE.sMergeStopFlag = rst!soeMergeStopFlag
        tlSOE.iAlertInterval = rst!soeAlertInterval
        tlSOE.sSchAutoGenSeq = rst!soeSchAutoGenSeq
        tlSOE.lMinEventID = rst!soeMinEventID
        tlSOE.lMaxEventID = rst!soeMaxEventID
        tlSOE.lCurrEventID = rst!soeCurrEventID
        tlSOE.iNoDaysRetainPW = rst!soeNoDaysRetainPW
        tlSOE.iVersion = rst!soeVersion
        tlSOE.iOrigSoeCode = rst!soeOrigSOECode
        tlSOE.sCurrent = rst!soeCurrent
        tlSOE.sEnteredDate = Format$(rst!soeEnteredDate, sgShowDateForm)
        tlSOE.sEnteredTime = Format$(rst!soeEnteredTime, sgShowTimeWSecForm)
        tlSOE.iUieCode = rst!soeUieCode
        tlSOE.iSpotItemIDWindow = rst!soeSpotItemIDWindow
        tlSOE.lTimeTolerance = rst!soeTimeTolerance
        tlSOE.lLengthTolerance = rst!soeLengthTolerance
        tlSOE.sMatchATNotB = rst!soeMatchATNotB
        tlSOE.sMatchATBNotI = rst!soeMatchATBNotI
        tlSOE.sMatchANotT = rst!soeMatchANotT
        tlSOE.sMatchBNotT = rst!soeMatchBNotT
        tlSOE.sSchAutoGenSeqTst = rst!soeSchAutoGenSeqTst
        tlSOE.sMergeStopFlagTst = rst!soeMergeStopFlagTst
        tlSOE.sUnused = ""
        rst.Close
        gGetRec_SOE_SiteOption = True
        Exit Function
    Else
        slNowDate = Format(Now, sgShowDateForm)  'Format(gNow(), sgShowDateForm)
        slNowTime = Format(Now, sgShowTimeWSecForm)  'Format(gNow(), sgShowTimeWSecForm)
        tlSOE.iCode = 0
        tlSOE.sClientName = "Client Name"
        tlSOE.sAddr1 = ""
        tlSOE.sAddr2 = ""
        tlSOE.sAddr3 = ""
        tlSOE.sPhone = ""
        tlSOE.sFax = ""
        tlSOE.iDaysRetainAsAir = 0
        tlSOE.lChgInterval = 0
        tlSOE.sMergeDateFormat = ""
        tlSOE.sMergeTimeFormat = ""
        tlSOE.sMergeFileFormat = ""
        tlSOE.sMergeFileExt = ""
        tlSOE.sMergeStartTime = Format("12am", sgShowTimeWSecForm)
        tlSOE.sMergeEndTime = Format("12am", sgShowTimeWSecForm)
        tlSOE.iMergeChkInterval = 0
        tlSOE.sMergeStopFlag = "N"
        tlSOE.iAlertInterval = 0
        tlSOE.sSchAutoGenSeq = "I"
        tlSOE.lMinEventID = 0
        tlSOE.lMaxEventID = 99999
        tlSOE.lCurrEventID = 10000
        tlSOE.iVersion = 0
        tlSOE.iOrigSoeCode = 0
        tlSOE.sCurrent = "Y"
        tlSOE.sEnteredDate = slNowDate
        tlSOE.sEnteredTime = slNowTime
        tlSOE.iUieCode = tgUIE.iCode
        tlSOE.iSpotItemIDWindow = 1000
        tlSOE.lTimeTolerance = 0
        tlSOE.lLengthTolerance = 0
        tlSOE.sMatchATNotB = "Y"
        tlSOE.sMatchATBNotI = "Y"
        tlSOE.sMatchANotT = "Y"
        tlSOE.sMatchBNotT = "Y"
        tlSOE.sSchAutoGenSeqTst = "I"
        tlSOE.sMergeStopFlagTst = "Y"
        rst.Close
        gGetRec_SOE_SiteOption = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_SOE_SiteOption = False
    Exit Function
End Function

Public Function gGetRecs_SPE_SitePath(slSPEStamp As String, ilSoeCode As Integer, slForm_Module As String, tlSPE() As SPE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "spe.eng") & ilSoeCode
    
    On Error GoTo gGetRecs_SPE_SitePathErr
    ilRet = 0
    ilLowLimit = LBound(tlSPE)
    If ilRet <> 0 Then
        slSPEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slSPEStamp <> "") Then
        If slStamp = slSPEStamp Then
            gGetRecs_SPE_SitePath = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM SPE_Site_Path Where speSoeCode = " & ilSoeCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlSPE(0 To 0) As SPE
    ilUpper = 0
    While Not rst.EOF
        tlSPE(ilUpper).iCode = rst!speCode
        tlSPE(ilUpper).iSoeCode = rst!speSoeCode
        tlSPE(ilUpper).sType = rst!speType
        tlSPE(ilUpper).sSubType = rst!speSubType
        If (tlSPE(ilUpper).sSubType <> "P") And (tlSPE(ilUpper).sSubType <> "T") Then
            tlSPE(ilUpper).sSubType = "P"
        End If
        tlSPE(ilUpper).sPath = rst!spePath
        tlSPE(ilUpper).sUnused = ""
        ilUpper = ilUpper + 1
        ReDim Preserve tlSPE(0 To ilUpper) As SPE
        rst.MoveNext
    Wend
    slSPEStamp = slStamp
    rst.Close
    gGetRecs_SPE_SitePath = True
    Exit Function
    
gGetRecs_SPE_SitePathErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_SPE_SitePath = False
    Exit Function

End Function
Public Function gGetRec_SSE_Site_SMTP_Info(ilSoeCode As Integer, slForm_Module As String, tlSSE As SSE) As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM SSE_Site_SMTP_Info WHERE sseSoeCode = " & ilSoeCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlSSE.iCode = rst!sseCode
        tlSSE.iSoeCode = rst!sseSoeCode
        tlSSE.sEMailHost = rst!sseEmailHost
        tlSSE.iEMailPort = rst!sseEmailPort
        tlSSE.sEMailAcctName = rst!sseEmailAcctName
        tlSSE.sEMailPassword = rst!sseEmailPassword
        tlSSE.sEMailTLS = rst!sseEmailTLS
        tlSSE.sUnused = rst!sseUnused

        rst.Close
        gGetRec_SSE_Site_SMTP_Info = True
        Exit Function
    Else
        tlSSE.iCode = 0
        tlSSE.iSoeCode = ilSoeCode
        tlSSE.sEMailHost = ""
        tlSSE.iEMailPort = 0
        tlSSE.sEMailAcctName = ""
        tlSSE.sEMailPassword = ""
        tlSSE.sEMailTLS = ""
        tlSSE.sUnused = ""
        rst.Close
        gGetRec_SSE_Site_SMTP_Info = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_SSE_Site_SMTP_Info = False
    Exit Function
End Function


Public Function gGetRecs_TSE_TemplateSchd(slTSEStamp As String, llDheCode As Long, slForm_Module As String, tlTSE() As TSE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "TSE.eng") & llDheCode
    slStamp = Trim$(Str$(llDheCode))
    
    On Error GoTo gGetRecs_TSE_TemplateSchdErr
    ilRet = 0
    ilLowLimit = LBound(tlTSE)
    If ilRet <> 0 Then
        slTSEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slTSEStamp <> "") Then
        sgSQLQuery = "SELECT Count(tseCode) FROM TSE_Template_Schd"
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlTSE)) And (slStamp = slTSEStamp) Then
            gGetRecs_TSE_TemplateSchd = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM TSE_Template_Schd Where tseDheCode = " & llDheCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlTSE(0 To 0) As TSE
    llUpper = 0
    While Not rst.EOF
        tlTSE(llUpper).lCode = rst!tseCode
        tlTSE(llUpper).lDheCode = rst!tseDheCode
        tlTSE(llUpper).iBdeCode = rst!tseBdeCode
        tlTSE(llUpper).sLogDate = Format$(rst!tseLogDate, sgShowDateForm)
        tlTSE(llUpper).sStartTime = Format$(rst!tseStartTime, sgShowTimeWSecForm)
        tlTSE(llUpper).sDescription = rst!tseDescription
        tlTSE(llUpper).sState = rst!tseState
        tlTSE(llUpper).lCteCode = rst!tseCteCode
        tlTSE(llUpper).iVersion = rst!tseVersion
        tlTSE(llUpper).lOrigTseCode = rst!tseOrigTseCode
        tlTSE(llUpper).sCurrent = rst!tseCurrent
        tlTSE(llUpper).sEnteredDate = Format$(rst!tseEnteredDate, sgShowDateForm)
        tlTSE(llUpper).sEnteredTime = Format$(rst!tseEnteredTime, sgShowTimeWSecForm)
        tlTSE(llUpper).iUieCode = rst!tseUieCode
        tlTSE(llUpper).sUnused = rst!tseUnused
        llUpper = llUpper + 1
        ReDim Preserve tlTSE(0 To llUpper) As TSE
        rst.MoveNext
    Wend
    slTSEStamp = slStamp
    rst.Close
    gGetRecs_TSE_TemplateSchd = True
    Exit Function
    
gGetRecs_TSE_TemplateSchdErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_TSE_TemplateSchd = False
    Exit Function

End Function

Public Function gGetRec_UIE_UserInfo(ilCode As Integer, slForm_Module As String, tlUie As UIE) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM UIE_User_Info WHERE uieCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlUie.iCode = rst!uieCode
        tlUie.sSignOnName = rst!uieSignOnName
        tlUie.sPassword = rst!uiePassword
        tlUie.sLastDatePWSet = Format$(rst!uieLastDatePWSet, sgShowDateForm)
        tlUie.sShowName = rst!uieShowName
        tlUie.sState = rst!uieState
        tlUie.sEMail = rst!uieEmail
        tlUie.sLastSignOnDate = Format$(rst!uieLastSignOnDate, sgShowDateForm)
        tlUie.sLastSignOnTime = Format$(rst!uieLastSignOnTime, sgShowTimeWSecForm)
        tlUie.sUsedFlag = rst!uieUsedFlag
        tlUie.iVersion = rst!uieVersion
        tlUie.iOrigUieCode = rst!uieOrigUieCode
        tlUie.sCurrent = rst!uieCurrent
        tlUie.sEnteredDate = Format$(rst!uieEnteredDate, sgShowDateForm)
        tlUie.sEnteredTime = Format$(rst!uieEnteredTime, sgShowTimeWSecForm)
        tlUie.iUieCode = rst!uieUieCode
        tlUie.sUnused = ""
        rst.Close
        gGetRec_UIE_UserInfo = True
        Exit Function
    Else
        rst.Close
        gGetRec_UIE_UserInfo = False
        Exit Function
    End If
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRec_UIE_UserInfo = False
    Exit Function
End Function

Public Function gGetTypeOfRecs_TTE_TimeType(slGetType As String, slTimeType As String, slTTEStamp As String, slForm_Module As String, tlTTE() As TTE) As Integer
'
'   slGetType(I)- C=Current; H=History, B=Both
'   slTimeType(I)- S=Start; E=End Time Type
'
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    slStamp = FileDateTime(sgDBPath & "tte.eng") & slGetType & slTimeType
    
    On Error GoTo gGetTypeOfRecs_TTE_TimeTypeErr
    ilRet = 0
    ilLowLimit = LBound(tlTTE)
    If ilRet <> 0 Then
        slTTEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slTTEStamp <> "") Then
        If slStamp = slTTEStamp Then
            gGetTypeOfRecs_TTE_TimeType = True
            Exit Function
        End If
    End If
    If slGetType = "B" Then
        sgSQLQuery = "SELECT * FROM TTE_Time_Type"
    ElseIf slGetType = "H" Then
        sgSQLQuery = "SELECT * FROM TTE_Time_Type WHERE tteCurrent = 'N'"
    Else
        sgSQLQuery = "SELECT * FROM TTE_Time_Type WHERE tteCurrent = 'Y'"
    End If
    If slGetType = "B" Then
        If slTimeType = "S" Or slTimeType = "E" Then
            sgSQLQuery = sgSQLQuery & " Where tteType = '" & slTimeType & "'"
        End If
    Else
        If slTimeType = "S" Or slTimeType = "E" Then
            sgSQLQuery = sgSQLQuery & " And tteType = '" & slTimeType & "'"
        End If
    End If
    sgSQLQuery = sgSQLQuery & " ORDER BY tteCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlTTE(0 To 0) As TTE
    ilUpper = 0
    While Not rst.EOF
        tlTTE(ilUpper).iCode = rst!tteCode
        tlTTE(ilUpper).sType = rst!tteType
        tlTTE(ilUpper).sName = rst!tteName
        tlTTE(ilUpper).sDescription = rst!tteDescription
        tlTTE(ilUpper).sState = rst!tteState
        tlTTE(ilUpper).sUsedFlag = rst!tteUsedFlag
        tlTTE(ilUpper).iVersion = rst!tteVersion
        tlTTE(ilUpper).iOrigTteCode = rst!tteOrigTteCode
        tlTTE(ilUpper).sCurrent = rst!tteCurrent
        tlTTE(ilUpper).sEnteredDate = Format$(rst!tteEnteredDate, sgShowDateForm)
        tlTTE(ilUpper).sEnteredTime = Format$(rst!tteEnteredTime, sgShowTimeWSecForm)
        tlTTE(ilUpper).iUieCode = rst!tteUieCode
        tlTTE(ilUpper).sUnused = rst!tteUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlTTE(0 To ilUpper) As TTE
        rst.MoveNext
    Wend
    slTTEStamp = slStamp
    rst.Close
    gGetTypeOfRecs_TTE_TimeType = True
    Exit Function
    
gGetTypeOfRecs_TTE_TimeTypeErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetTypeOfRecs_TTE_TimeType = False
    Exit Function

End Function

Public Function gGetRecs_AAE_As_Aired(slAAEStamp As String, llSheCode As Long, slForm_Module As String, tlAAE() As AAE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "AAE.eng") & ilAeeCode
    slStamp = Trim$(Str$(llSheCode))
    
    On Error GoTo gGetRecs_AAE_As_AiredErr
    ilRet = 0
    ilLowLimit = LBound(tlAAE)
    If ilRet <> 0 Then
        slAAEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slAAEStamp <> "") Then
        sgSQLQuery = "SELECT Count(aaeCode) FROM AAE_As_Aired WHERE aaeSheCode = " & llSheCode
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlAAE)) And (slStamp = slAAEStamp) Then
            gGetRecs_AAE_As_Aired = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM AAE_As_Aired WHERE aaeSheCode = " & llSheCode & "ORDER BY aaeEventID"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlAAE(0 To 30000) As AAE
    llUpper = 0
    While Not rst.EOF
        tlAAE(llUpper).lCode = rst!aaeCode
        tlAAE(llUpper).lSheCode = rst!aaeSheCode
        tlAAE(llUpper).lSeeCode = rst!aaeSeeCode
        tlAAE(llUpper).sAirDate = Format$(rst!aaeAirDate, sgShowDateForm)
        tlAAE(llUpper).lAirTime = rst!aaeAirTime
        tlAAE(llUpper).sAutoOff = rst!aaeAutoOff
        tlAAE(llUpper).sData = rst!aaeData
        tlAAE(llUpper).sSchedule = rst!aaeSchedule
        tlAAE(llUpper).sTrueTime = rst!aaeTrueTime
        tlAAE(llUpper).sSourceConflict = rst!aaeSourceConflict
        tlAAE(llUpper).sSourceUnavail = rst!aaeSourceUnavail
        tlAAE(llUpper).sSourceItem = rst!aaeSourceItem
        tlAAE(llUpper).sBkupSrceUnavail = rst!aaeBkupSrceUnavail
        tlAAE(llUpper).sBkupSrceItem = rst!aaeBkupSrceItem
        tlAAE(llUpper).sProtSrceUnavail = rst!aaeProtSrceUnavail
        tlAAE(llUpper).sProtSrceItem = rst!aaeProtSrceItem
        tlAAE(llUpper).sDate = rst!aaeDate
        tlAAE(llUpper).lEventID = rst!aaeEventID
        tlAAE(llUpper).sBusName = rst!aaeBusName
        tlAAE(llUpper).sBusControl = rst!aaeBusControl
        tlAAE(llUpper).sEventType = rst!aaeEventType
        tlAAE(llUpper).sStartTime = rst!aaeStartTime
        tlAAE(llUpper).sStartType = rst!aaeStartType
        tlAAE(llUpper).sFixedTime = rst!aaeFixedTime
        tlAAE(llUpper).sEndType = rst!aaeEndType
        tlAAE(llUpper).sDuration = rst!aaeDuration
        tlAAE(llUpper).sOutTime = rst!aaeOutTime
        tlAAE(llUpper).sMaterialType = rst!aaeMaterialType
        tlAAE(llUpper).sAudioName = rst!aaeAudioName
        tlAAE(llUpper).sAudioItemID = rst!aaeAudioItemID
        tlAAE(llUpper).sAudioISCI = rst!aaeAudioISCI
        tlAAE(llUpper).sAudioCrtlChar = rst!aaeAudioCrtlChar
        tlAAE(llUpper).sBkupAudioName = rst!aaeBkupAudioName
        tlAAE(llUpper).sBkupCtrlChar = rst!aaeBkupCtrlChar
        tlAAE(llUpper).sProtAudioName = rst!aaeProtAudioName
        tlAAE(llUpper).sProtItemID = rst!aaeProtItemID
        tlAAE(llUpper).sProtISCI = rst!aaeProtISCI
        tlAAE(llUpper).sProtCtrlChar = rst!aaeProtCtrlChar
        tlAAE(llUpper).sRelay1 = rst!aaeRelay1
        tlAAE(llUpper).sRelay2 = rst!aaeRelay2
        tlAAE(llUpper).sFollow = rst!aaeFollow
        tlAAE(llUpper).sSilenceTime = rst!aaeSilenceTime
        tlAAE(llUpper).sSilence1 = rst!aaeSilence1
        tlAAE(llUpper).sSilence2 = rst!aaeSilence2
        tlAAE(llUpper).sSilence3 = rst!aaeSilence3
        tlAAE(llUpper).sSilence4 = rst!aaeSilence4
        tlAAE(llUpper).sNetcueStart = rst!aaeNetcueStart
        tlAAE(llUpper).sNetcueEnd = rst!aaeNetcueEnd
        tlAAE(llUpper).sTitle1 = rst!aaeTitle1
        tlAAE(llUpper).sTitle2 = rst!aaeTitle2
        tlAAE(llUpper).sABCFormat = rst!aaeABCFormat
        tlAAE(llUpper).sABCPgmCode = rst!aaeABCPgmCode
        tlAAE(llUpper).sABCXDSMode = rst!aaeABCXDSMode
        tlAAE(llUpper).sABCRecordItem = rst!aaeABCRecordItem
        tlAAE(llUpper).sEnteredDate = Format$(rst!aaeEnteredDate, sgShowDateForm)
        tlAAE(llUpper).sEnteredTime = Format$(rst!aaeEnteredTime, sgShowTimeWSecForm)
        tlAAE(llUpper).sUnused = rst!aaeUnused
        llUpper = llUpper + 1
        If llUpper > UBound(tlAAE) Then
            ReDim Preserve tlAAE(0 To llUpper + 1000) As AAE
        End If
        rst.MoveNext
    Wend
    ReDim Preserve tlAAE(0 To llUpper) As AAE
    slAAEStamp = slStamp
    rst.Close
    gGetRecs_AAE_As_Aired = True
    Exit Function
    
gGetRecs_AAE_As_AiredErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_AAE_As_Aired = False
    Exit Function

End Function

Public Function gGetRecs_ACE_AutoContact(slACEStamp As String, ilAeeCode As Integer, slForm_Module As String, tlACE() As ACE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "ace.eng") & ilAeeCode
    slStamp = Trim$(Str$(ilAeeCode))
    
    On Error GoTo gGetRecs_ACE_AutoContactErr
    ilRet = 0
    ilLowLimit = LBound(tlACE)
    If ilRet <> 0 Then
        slACEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slACEStamp <> "") Then
        sgSQLQuery = "SELECT Count(aceCode) FROM ACE_Auto_Contact WHERE aceAeeCode = " & ilAeeCode
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlACE)) And (slStamp = slACEStamp) Then
            rst.Close
            gGetRecs_ACE_AutoContact = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM ACE_Auto_Contact WHERE aceAeeCode = " & ilAeeCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlACE(0 To 0) As ACE
    ilUpper = 0
    While Not rst.EOF
        tlACE(ilUpper).iCode = rst!aceCode
        tlACE(ilUpper).iAeeCode = rst!aceAeeCode
        tlACE(ilUpper).sType = rst!aceType
        tlACE(ilUpper).sContact = rst!aceContact
        tlACE(ilUpper).sPhone = rst!acePhone
        tlACE(ilUpper).sFax = rst!aceFax
        tlACE(ilUpper).sEMail = rst!aceEMail
        tlACE(ilUpper).sUnused = rst!aceUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlACE(0 To ilUpper) As ACE
        rst.MoveNext
    Wend
    slACEStamp = slStamp
    rst.Close
    gGetRecs_ACE_AutoContact = True
    Exit Function
    
gGetRecs_ACE_AutoContactErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_ACE_AutoContact = False
    Exit Function

End Function

Public Function gGetRecs_ADE_AutoDataFlags(slADEStamp As String, ilAeeCode As Integer, slForm_Module As String, tlADE() As ADE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "ade.eng") & ilAeeCode
    slStamp = Trim$(Str$(ilAeeCode))
    
    On Error GoTo gGetRecs_ADE_AutoDataFlagsErr
    ilRet = 0
    ilLowLimit = LBound(tlADE)
    If ilRet <> 0 Then
        slADEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slADEStamp <> "") Then
        sgSQLQuery = "SELECT Count(adeCode) FROM ADE_Auto_Data_Flags WHERE adeAeeCode = " & ilAeeCode
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlADE)) And (slStamp = slADEStamp) Then
            rst.Close
            gGetRecs_ADE_AutoDataFlags = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM ADE_Auto_Data_Flags WHERE adeAeeCode = " & ilAeeCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlADE(0 To 0) As ADE
    ilUpper = 0
    While Not rst.EOF
        tlADE(ilUpper).iCode = rst!adeCode
        tlADE(ilUpper).iAeeCode = rst!adeAeeCode
        tlADE(ilUpper).iScheduleData = rst!adeScheduleData
        tlADE(ilUpper).iDate = rst!adeDate
        tlADE(ilUpper).iDateNoChar = rst!adeDateNoChar
        tlADE(ilUpper).iTime = rst!adeTime
        tlADE(ilUpper).iTimeNoChar = rst!adeTimeNoChar
        tlADE(ilUpper).iAutoOff = rst!adeAutoOff
        tlADE(ilUpper).iData = rst!adeData
        tlADE(ilUpper).iSchedule = rst!adeSchedule
        tlADE(ilUpper).iTrueTime = rst!adeTrueTime
        tlADE(ilUpper).iSourceConflict = rst!adeSourceConflict
        tlADE(ilUpper).iSourceUnavail = rst!adeSourceUnavail
        tlADE(ilUpper).iSourceItem = rst!adeSourceItem
        tlADE(ilUpper).iBkupSrceUnavail = rst!adeBkupSrceUnavail
        tlADE(ilUpper).iBkupSrceItem = rst!adeBkupSrceItem
        tlADE(ilUpper).iProtSrceUnavail = rst!adeProtSrceUnavail
        tlADE(ilUpper).iProtSrceItem = rst!adeProtSrceItem
        tlADE(ilUpper).sUnused = rst!adeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlADE(0 To ilUpper) As ADE
        rst.MoveNext
    Wend
    slADEStamp = slStamp
    rst.Close
    gGetRecs_ADE_AutoDataFlags = True
    Exit Function
    
gGetRecs_ADE_AutoDataFlagsErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_ADE_AutoDataFlags = False
    Exit Function

End Function

Public Function gGetRecs_AFE_AutoFormat(slAFEStamp As String, ilAeeCode As Integer, slForm_Module As String, tlAFE() As AFE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "efe.eng") & ilAeeCode
    slStamp = Trim$(Str$(ilAeeCode))
    
    On Error GoTo gGetRecs_AFE_AutoFormatErr
    ilRet = 0
    ilLowLimit = LBound(tlAFE)
    If ilRet <> 0 Then
        slAFEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slAFEStamp <> "") Then
        sgSQLQuery = "SELECT Count(afeCode) FROM AFE_Auto_Format WHERE afeAeeCode = " & ilAeeCode
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlAFE)) And (slStamp = slAFEStamp) Then
            rst.Close
            gGetRecs_AFE_AutoFormat = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM AFE_Auto_Format WHERE afeAeeCode = " & ilAeeCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlAFE(0 To 0) As AFE
    ilUpper = 0
    While Not rst.EOF
        tlAFE(ilUpper).iCode = rst!afeCode
        tlAFE(ilUpper).iAeeCode = rst!afeAeeCode
        tlAFE(ilUpper).sType = rst!afeType
        tlAFE(ilUpper).sSubType = rst!afeSubType
        tlAFE(ilUpper).iBus = rst!afeBus
        tlAFE(ilUpper).iBusControl = rst!afeBusControl
        tlAFE(ilUpper).iEventType = rst!afeEventType
        tlAFE(ilUpper).iTime = rst!afeTime
        tlAFE(ilUpper).iStartType = rst!afeStartType
        tlAFE(ilUpper).iFixedTime = rst!afeFixedTime
        tlAFE(ilUpper).iEndType = rst!afeEndType
        tlAFE(ilUpper).iDuration = rst!afeDuration
        tlAFE(ilUpper).iEndTime = rst!afeEndTime
        tlAFE(ilUpper).iMaterialType = rst!afeMaterialType
        tlAFE(ilUpper).iAudioName = rst!afeAudioName
        tlAFE(ilUpper).iAudioItemID = rst!afeAudioItemID
        tlAFE(ilUpper).iAudioISCI = rst!afeAudioISCI
        tlAFE(ilUpper).iAudioControl = rst!afeAudioControl
        tlAFE(ilUpper).iBkupAudioName = rst!afeBkupAudioName
        tlAFE(ilUpper).iBkupAudioControl = rst!afeBkupAudioControl
        tlAFE(ilUpper).iProtAudioName = rst!afeProtAudioName
        tlAFE(ilUpper).iProtItemID = rst!afeProtItemID
        tlAFE(ilUpper).iProtISCI = rst!afeProtISCI
        tlAFE(ilUpper).iProtAudioControl = rst!afeProtAudioControl
        tlAFE(ilUpper).iRelay1 = rst!afeRelay1
        tlAFE(ilUpper).iRelay2 = rst!afeRelay2
        tlAFE(ilUpper).iFollow = rst!afeFollow
        tlAFE(ilUpper).iSilenceTime = rst!afeSilenceTime
        tlAFE(ilUpper).iSilence1 = rst!afeSilence1
        tlAFE(ilUpper).iSilence2 = rst!afeSilence2
        tlAFE(ilUpper).iSilence3 = rst!afeSilence3
        tlAFE(ilUpper).iSilence4 = rst!afeSilence4
        tlAFE(ilUpper).iStartNetcue = rst!afeStartNetcue
        tlAFE(ilUpper).iStopNetcue = rst!afeStopNetcue
        tlAFE(ilUpper).iTitle1 = rst!afeTitle1
        tlAFE(ilUpper).iTitle2 = rst!afeTitle2
        tlAFE(ilUpper).iEventID = rst!afeEventID
        tlAFE(ilUpper).iDate = rst!afeDate
        tlAFE(ilUpper).iABCFormat = rst!afeABCFormat
        tlAFE(ilUpper).iABCPgmCode = rst!afeABCPgmCode
        tlAFE(ilUpper).iABCXDSMode = rst!afeABCXDSMode
        tlAFE(ilUpper).iABCRecordItem = rst!afeABCRecordItem
        tlAFE(ilUpper).sUnused = rst!afeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlAFE(0 To ilUpper) As AFE
        rst.MoveNext
    Wend
    slAFEStamp = slStamp
    rst.Close
    gGetRecs_AFE_AutoFormat = True
    Exit Function
    
gGetRecs_AFE_AutoFormatErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_AFE_AutoFormat = False
    Exit Function

End Function


Public Function gGetRecs_APE_AutoPath(slAPEStamp As String, ilAeeCode As Integer, slForm_Module As String, tlAPE() As APE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "epe.eng") & ilAeeCode
    slStamp = Trim$(Str$(ilAeeCode))
    
    On Error GoTo gGetRecs_APE_AutoPathErr
    ilRet = 0
    ilLowLimit = LBound(tlAPE)
    If ilRet <> 0 Then
        slAPEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slAPEStamp <> "") Then
        sgSQLQuery = "SELECT Count(apeCode) FROM APE_Auto_Path WHERE apeAeeCode = " & ilAeeCode
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlAPE)) And (slStamp = slAPEStamp) Then
            rst.Close
            gGetRecs_APE_AutoPath = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM APE_AutoPath WHERE apeAeeCode = " & ilAeeCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlAPE(0 To 0) As APE
    ilUpper = 0
    While Not rst.EOF
        tlAPE(ilUpper).iCode = rst!apeCode
        tlAPE(ilUpper).iAeeCode = rst!apeAeeCode
        tlAPE(ilUpper).sType = rst!apeType
        tlAPE(ilUpper).sSubType = rst!apeSubType
        If (tlAPE(ilUpper).sSubType <> "P") And (tlAPE(ilUpper).sSubType <> "T") Then
            tlAPE(ilUpper).sSubType = "P"
        End If
        tlAPE(ilUpper).sNewFileName = rst!apeNewFileName
        tlAPE(ilUpper).sChgFileName = rst!apeChgFileName
        tlAPE(ilUpper).sDelFileName = rst!apeDelFileName
        tlAPE(ilUpper).sNewFileExt = rst!apeNewFileExt
        tlAPE(ilUpper).sChgFileExt = rst!apeChgFileExt
        tlAPE(ilUpper).sDelFileExt = rst!apeDelFileExt
        tlAPE(ilUpper).sPath = rst!apePath
        tlAPE(ilUpper).sDateFormat = rst!apeDateFormat
        tlAPE(ilUpper).sTimeFormat = rst!apeTimeFormat
        tlAPE(ilUpper).sUnused = rst!apeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlAPE(0 To ilUpper) As APE
        rst.MoveNext
    Wend
    slAPEStamp = slStamp
    rst.Close
    gGetRecs_APE_AutoPath = True
    Exit Function
    
gGetRecs_APE_AutoPathErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_APE_AutoPath = False
    Exit Function

End Function

Public Function gGetRecs_ARE_AdvertiserRefer(slAREStamp As String, slForm_Module As String, tlARE() As ARE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo gGetRecs_ARE_AdvertiserReferErr
    ilRet = 0
    ilLowLimit = LBound(tlARE)
    If ilRet <> 0 Then
        slAREStamp = ""
    End If
    On Error GoTo ErrHand
    If (slAREStamp <> "") Then
        sgSQLQuery = "SELECT Count(areCode) FROM ARE_Advertiser_Refer"
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlARE)) Then
            rst.Close
            gGetRecs_ARE_AdvertiserRefer = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM ARE_Advertiser_Refer"
    sgSQLQuery = sgSQLQuery & " ORDER BY areCode"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlARE(0 To 0) As ARE
    ilUpper = 0
    While Not rst.EOF
        tlARE(ilUpper).lCode = rst!areCode
        tlARE(ilUpper).sName = rst!areName
        tlARE(ilUpper).sUnusued = rst!areUnusued
        ilUpper = ilUpper + 1
        ReDim Preserve tlARE(0 To ilUpper) As ARE
        rst.MoveNext
    Wend
    slAREStamp = UBound(tlARE)
    rst.Close
    gGetRecs_ARE_AdvertiserRefer = True
    Exit Function
    
gGetRecs_ARE_AdvertiserReferErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_ARE_AdvertiserRefer = False
    Exit Function

End Function

Public Function gGetRecs_BSE_BusSelGroup(slType As String, slBSEStamp As String, ilBdeCode As Integer, slForm_Module As String, tlBSE() As BSE) As Integer
    'slType(I)- G=By Group; B=By Buses
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "bse.eng") & ilBdeCode & slType
    slStamp = Trim$(Str$(ilBdeCode)) & slType
    
    On Error GoTo gGetRecs_BSE_BusSelGroupErr
    ilRet = 0
    ilLowLimit = LBound(tlBSE)
    If ilRet <> 0 Then
        slBSEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slBSEStamp <> "") Then
        sgSQLQuery = "SELECT Count(bseCode) FROM BSE_Bus_Sel_Group"
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlBSE)) And (slStamp = slBSEStamp) Then
            rst.Close
            gGetRecs_BSE_BusSelGroup = True
            Exit Function
        End If
    End If
    If slType = "B" Then
        sgSQLQuery = "SELECT * FROM BSE_Bus_Sel_Group Where bseBdeCode = " & ilBdeCode
    Else
        sgSQLQuery = "SELECT * FROM BSE_Bus_Sel_Group Where bseBgeCode = " & ilBdeCode
    End If
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlBSE(0 To 0) As BSE
    ilUpper = 0
    While Not rst.EOF
        tlBSE(ilUpper).iCode = rst!bseCode
        tlBSE(ilUpper).iBdeCode = rst!bseBdeCode
        tlBSE(ilUpper).iBgeCode = rst!bseBgeCode
        tlBSE(ilUpper).sUnused = rst!bseUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlBSE(0 To ilUpper) As BSE
        rst.MoveNext
    Wend
    slBSEStamp = slStamp
    rst.Close
    gGetRecs_BSE_BusSelGroup = True
    Exit Function
    
gGetRecs_BSE_BusSelGroupErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_BSE_BusSelGroup = False
    Exit Function

End Function

Public Function gGetRecs_DEE_DayEvent(slDEEStamp As String, llDheCode As Long, slForm_Module As String, tlDEE() As DEE) As Integer
    Dim tlDEESrchKey As LONGKEY0
    Dim ilDEERecLen As Integer
    Dim ilRet As Integer
    Dim tlDEEAPI As DEEAPI
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "dee.eng") & llDheCode
    slStamp = Trim$(Str$(llDheCode))
    
    On Error GoTo gGetRecs_DEE_DayEventErr
    ilRet = 0
    ilLowLimit = LBound(tlDEE)
    If ilRet <> 0 Then
        slDEEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slDEEStamp <> "") Then
        sgSQLQuery = "SELECT Count(deeCode) FROM DEE_Day_Event_Info"
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlDEE)) And (slStamp = slDEEStamp) Then
            rst.Close
            gGetRecs_DEE_DayEvent = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM DEE_Day_Event_Info Where deeDheCode = " & llDheCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlDEE(0 To 0) As DEE
    llUpper = 0
    While Not rst.EOF
        tlDEE(llUpper).lCode = rst!deeCode
        tlDEE(llUpper).lDheCode = rst!deeDheCode
        tlDEE(llUpper).iCceCode = rst!deeCceCode
        tlDEE(llUpper).iEteCode = rst!deeEteCode
        tlDEE(llUpper).lTime = rst!deeTime
        tlDEE(llUpper).iStartTteCode = rst!deeStartTteCode
        tlDEE(llUpper).sFixedTime = rst!deeFixedTime
        tlDEE(llUpper).iEndTteCode = rst!deeEndTteCode
        tlDEE(llUpper).lDuration = rst!deeDuration
        tlDEE(llUpper).sHours = rst!deeHours
        tlDEE(llUpper).sDays = rst!deeDays
        tlDEE(llUpper).iMteCode = rst!deeMteCode
        tlDEE(llUpper).iAudioAseCode = rst!deeAudioAseCode
        tlDEE(llUpper).sAudioItemID = rst!deeAudioItemID
        tlDEE(llUpper).sAudioISCI = rst!deeAudioISCI
        tlDEE(llUpper).iAudioCceCode = rst!deeAudioCceCode
        tlDEE(llUpper).iBkupAneCode = rst!deeBkupAneCode
        tlDEE(llUpper).iBkupCceCode = rst!deeBkupCceCode
        tlDEE(llUpper).iProtAneCode = rst!deeProtAneCode
        tlDEE(llUpper).sProtItemID = rst!deeProtItemID
        tlDEE(llUpper).sProtISCI = rst!deeProtISCI
        tlDEE(llUpper).iProtCceCode = rst!deeProtCceCode
        tlDEE(llUpper).i1RneCode = rst!dee1RneCode
        tlDEE(llUpper).i2RneCode = rst!dee2RneCode
        tlDEE(llUpper).iFneCode = rst!deeFneCode
        tlDEE(llUpper).lSilenceTime = rst!deeSilenceTime
        tlDEE(llUpper).i1SceCode = rst!dee1SceCode
        tlDEE(llUpper).i2SceCode = rst!dee2SceCode
        tlDEE(llUpper).i3SceCode = rst!dee3SceCode
        tlDEE(llUpper).i4SceCode = rst!dee4SceCode
        tlDEE(llUpper).iStartNneCode = rst!deeStartNneCode
        tlDEE(llUpper).iEndNneCode = rst!deeEndNneCode
        tlDEE(llUpper).l1CteCode = rst!dee1CteCode
        tlDEE(llUpper).l2CteCode = rst!dee2CteCode
        tlDEE(llUpper).lEventID = rst!deeEventID
        tlDEE(llUpper).sIgnoreConflicts = rst!deeIgnoreConflicts
        tlDEE(llUpper).sABCFormat = rst!deeABCFormat
        tlDEE(llUpper).sABCPgmCode = rst!deeABCPgmCode
        tlDEE(llUpper).sABCXDSMode = rst!deeABCXDSMode
        tlDEE(llUpper).sABCRecordItem = rst!deeABCRecordItem
        tlDEE(llUpper).sUnused = rst!deeUnused
        llUpper = llUpper + 1
        ReDim Preserve tlDEE(0 To llUpper) As DEE
        rst.MoveNext
    Wend
    slDEEStamp = slStamp
    rst.Close
    gGetRecs_DEE_DayEvent = True
    Exit Function
    
gGetRecs_DEE_DayEventErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_DEE_DayEvent = False
    Exit Function

End Function


Public Function gGetRecs_DEE_DayEventAPI(hlDEE As Integer, slDEEStamp As String, llDheCode As Long, slForm_Module As String, tlDEE() As DEE) As Integer
    Dim tlDEESrchKey As LONGKEY0
    Dim ilDEERecLen As Integer
    Dim ilRet As Integer
    Dim tlDEEAPI As DEEAPI
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "dee.eng") & llDheCode
    slStamp = Trim$(Str$(llDheCode))
    
    On Error GoTo gGetRecs_DEE_DayEventErr
    ilRet = 0
    ilLowLimit = LBound(tlDEE)
    If ilRet <> 0 Then
        slDEEStamp = ""
    End If
    On Error GoTo ErrHand
    'If (slDEEStamp <> "") Then
    '    sgSQLQuery = "SELECT Count(deeCode) FROM DEE_Day_Event_Info"
    '    Set rst = cnn.Execute(sgSQLQuery)
    '    If (rst(0).Value = UBound(tlDEE)) And (slStamp = slDEEStamp) Then
    '        gGetRecs_DEE_DayEvent = True
    '        Exit Function
    '    End If
    'End If
    'sgSQLQuery = "SELECT * FROM DEE_Day_Event_Info Where deeDheCode = " & llDheCode
    'Set rst = cnn.Execute(sgSQLQuery)
    'ReDim tlDEE(0 To 0) As DEE
    'llUpper = 0
    'While Not rst.EOF
    '    tlDEE(llUpper).lCode = rst!deeCode
    '    tlDEE(llUpper).lDheCode = rst!deeDheCode
    '    tlDEE(llUpper).iCceCode = rst!deeCceCode
    '    tlDEE(llUpper).iEteCode = rst!deeEteCode
    '    tlDEE(llUpper).lTime = rst!deeTime
    '    tlDEE(llUpper).iStartTteCode = rst!deeStartTteCode
    '    tlDEE(llUpper).sFixedTime = rst!deeFixedTime
    '    tlDEE(llUpper).iEndTteCode = rst!deeEndTteCode
    '    tlDEE(llUpper).lDuration = rst!deeDuration
    '    tlDEE(llUpper).sHours = rst!deeHours
    '    tlDEE(llUpper).sDays = rst!deeDays
    '    tlDEE(llUpper).iMteCode = rst!deeMteCode
    '    tlDEE(llUpper).iAudioAseCode = rst!deeAudioAseCode
    '    tlDEE(llUpper).sAudioItemID = rst!deeAudioItemID
    '    tlDEE(llUpper).sAudioISCI = rst!deeAudioISCI
    '    tlDEE(llUpper).iAudioCceCode = rst!deeAudioCceCode
    '    tlDEE(llUpper).iBkupAneCode = rst!deeBkupAneCode
    '    tlDEE(llUpper).iBkupCceCode = rst!deeBkupCceCode
    '    tlDEE(llUpper).iProtAneCode = rst!deeProtAneCode
    '    tlDEE(llUpper).sProtItemID = rst!deeProtItemID
    '    tlDEE(llUpper).sProtISCI = rst!deeProtISCI
    '    tlDEE(llUpper).iProtCceCode = rst!deeProtCceCode
    '    tlDEE(llUpper).i1RneCode = rst!dee1RneCode
    '    tlDEE(llUpper).i2RneCode = rst!dee2RneCode
    '    tlDEE(llUpper).iFneCode = rst!deeFneCode
    '    tlDEE(llUpper).lSilenceTime = rst!deeSilenceTime
    '    tlDEE(llUpper).i1SceCode = rst!dee1SceCode
    '    tlDEE(llUpper).i2SceCode = rst!dee2SceCode
    '    tlDEE(llUpper).i3SceCode = rst!dee3SceCode
    '    tlDEE(llUpper).i4SceCode = rst!dee4SceCode
    '    tlDEE(llUpper).iStartNneCode = rst!deeStartNneCode
    '    tlDEE(llUpper).iEndNneCode = rst!deeEndNneCode
    '    tlDEE(llUpper).l1CteCode = rst!dee1CteCode
    '    tlDEE(llUpper).l2CteCode = rst!dee2CteCode
    '    tlDEE(llUpper).lEventID = rst!deeEventID
    '    tlDEE(llUpper).sIgnoreConflicts = rst!deeIgnoreConflicts
    '    tlDEE(llUpper).sUnused = rst!deeUnused
    '    llUpper = llUpper + 1
    '    ReDim Preserve tlDEE(0 To llUpper) As DEE
    '    rst.MoveNext
    'Wend
    'slDEEStamp = slStamp
    'rst.Close
    ReDim tlDEE(0 To 0) As DEE
    llUpper = 0
    ilDEERecLen = Len(tlDEEAPI)
    tlDEESrchKey.lCode = llDheCode
    ilRet = btrGetEqual(hlDEE, tlDEEAPI, ilDEERecLen, tlDEESrchKey, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do While (ilRet = BTRV_ERR_NONE) And (tlDEEAPI.lDheCode = llDheCode)
        tlDEE(llUpper).lCode = tlDEEAPI.lCode
        tlDEE(llUpper).lDheCode = tlDEEAPI.lDheCode
        tlDEE(llUpper).iCceCode = tlDEEAPI.iCceCode
        tlDEE(llUpper).iEteCode = tlDEEAPI.iEteCode
        tlDEE(llUpper).lTime = tlDEEAPI.lTime
        tlDEE(llUpper).iStartTteCode = tlDEEAPI.iStartTteCode
        tlDEE(llUpper).sFixedTime = tlDEEAPI.sFixedTime
        tlDEE(llUpper).iEndTteCode = tlDEEAPI.iEndTteCode
        tlDEE(llUpper).lDuration = tlDEEAPI.lDuration
        tlDEE(llUpper).sHours = tlDEEAPI.sHours
        tlDEE(llUpper).sDays = tlDEEAPI.sDays
        tlDEE(llUpper).iMteCode = tlDEEAPI.iMteCode
        tlDEE(llUpper).iAudioAseCode = tlDEEAPI.iAudioAseCode
        tlDEE(llUpper).sAudioItemID = tlDEEAPI.sAudioItemID
        tlDEE(llUpper).sAudioISCI = tlDEEAPI.sAudioISCI
        tlDEE(llUpper).iAudioCceCode = tlDEEAPI.iAudioCceCode
        tlDEE(llUpper).iBkupAneCode = tlDEEAPI.iBkupAneCode
        tlDEE(llUpper).iBkupCceCode = tlDEEAPI.iBkupCceCode
        tlDEE(llUpper).iProtAneCode = tlDEEAPI.iProtAneCode
        tlDEE(llUpper).sProtItemID = tlDEEAPI.sProtItemID
        tlDEE(llUpper).sProtISCI = tlDEEAPI.sProtISCI
        tlDEE(llUpper).iProtCceCode = tlDEEAPI.iProtCceCode
        tlDEE(llUpper).i1RneCode = tlDEEAPI.i1RneCode
        tlDEE(llUpper).i2RneCode = tlDEEAPI.i2RneCode
        tlDEE(llUpper).iFneCode = tlDEEAPI.iFneCode
        tlDEE(llUpper).lSilenceTime = tlDEEAPI.lSilenceTime
        tlDEE(llUpper).i1SceCode = tlDEEAPI.i1SceCode
        tlDEE(llUpper).i2SceCode = tlDEEAPI.i2SceCode
        tlDEE(llUpper).i3SceCode = tlDEEAPI.i3SceCode
        tlDEE(llUpper).i4SceCode = tlDEEAPI.i4SceCode
        tlDEE(llUpper).iStartNneCode = tlDEEAPI.iStartNneCode
        tlDEE(llUpper).iEndNneCode = tlDEEAPI.iEndNneCode
        tlDEE(llUpper).l1CteCode = tlDEEAPI.l1CteCode
        tlDEE(llUpper).l2CteCode = tlDEEAPI.l2CteCode
        tlDEE(llUpper).lEventID = tlDEEAPI.lEventID
        tlDEE(llUpper).sIgnoreConflicts = tlDEEAPI.sIgnoreConflicts
        tlDEE(llUpper).sABCFormat = tlDEEAPI.sABCFormat
        tlDEE(llUpper).sABCPgmCode = tlDEEAPI.sABCPgmCode
        tlDEE(llUpper).sABCXDSMode = tlDEEAPI.sABCXDSMode
        tlDEE(llUpper).sABCRecordItem = tlDEEAPI.sABCRecordItem
        tlDEE(llUpper).sUnused = tlDEEAPI.sUnused
        llUpper = llUpper + 1
        ReDim Preserve tlDEE(0 To llUpper) As DEE
        ilRet = btrGetNext(hlDEE, tlDEEAPI, ilDEERecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    gGetRecs_DEE_DayEventAPI = True
    Exit Function
    
gGetRecs_DEE_DayEventErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    'rst.Close
    gGetRecs_DEE_DayEventAPI = False
    Exit Function

End Function


Public Function gGetRecs_DHE_DayHeaderInfoByLibrary(slDHEStamp As String, llDNECode As Long, slForm_Module As String, tlDHE() As DHE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "dhe.eng") & llDneCode
    slStamp = Trim$(Str$(llDNECode))
    
    On Error GoTo gGetRecs_DHE_DayHeaderInfoByLibraryErr
    ilRet = 0
    ilLowLimit = LBound(tlDHE)
    If ilRet <> 0 Then
        slDHEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slDHEStamp <> "") Then
        sgSQLQuery = "SELECT Count(dheCode) FROM DHE_Day_Header_Info Where dheDneCode = " & llDNECode
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlDHE)) And (slStamp = slDHEStamp) Then
            rst.Close
            gGetRecs_DHE_DayHeaderInfoByLibrary = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM DHE_Day_Header_Info Where dheDneCode = " & llDNECode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlDHE(0 To 0) As DHE
    llUpper = 0
    While Not rst.EOF
        tlDHE(llUpper).lCode = rst!dheCode
        tlDHE(llUpper).sType = rst!dheType
        tlDHE(llUpper).lDneCode = rst!dheDneCode
        tlDHE(llUpper).lDseCode = rst!dheDseCode
        tlDHE(llUpper).sStartTime = Format$(rst!dheStartTime, sgShowTimeWSecForm)
        tlDHE(llUpper).lLength = rst!dheLength
        tlDHE(llUpper).sHours = rst!dheHours
        tlDHE(llUpper).sStartDate = Format$(rst!dheStartDate, sgShowDateForm)
        tlDHE(llUpper).sEndDate = Format$(rst!dheEndDate, sgShowDateForm)
        tlDHE(llUpper).sDays = rst!dheDays
        tlDHE(llUpper).lCteCode = rst!dheCteCode
        tlDHE(llUpper).sState = rst!dheState
        tlDHE(llUpper).sUsedFlag = rst!dheUsedFlag
        tlDHE(llUpper).iVersion = rst!dheVersion
        tlDHE(llUpper).lOrigDHECode = rst!dheOrigDheCode
        tlDHE(llUpper).sCurrent = rst!dheCurrent
        tlDHE(llUpper).sEnteredDate = Format$(rst!dheEnteredDate, sgShowDateForm)
        tlDHE(llUpper).sEnteredTime = Format$(rst!dheEnteredTime, sgShowTimeWSecForm)
        tlDHE(llUpper).iUieCode = rst!dheUieCode
        tlDHE(llUpper).sIgnoreConflicts = rst!dheIgnoreConflicts
        tlDHE(llUpper).sBusNames = rst!dheBusNames
        tlDHE(llUpper).sUnused = rst!dheUnused
        llUpper = llUpper + 1
        ReDim Preserve tlDHE(0 To llUpper) As DHE
        rst.MoveNext
    Wend
    slDHEStamp = slStamp
    rst.Close
    gGetRecs_DHE_DayHeaderInfoByLibrary = True
    Exit Function
    
gGetRecs_DHE_DayHeaderInfoByLibraryErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_DHE_DayHeaderInfoByLibrary = False
    Exit Function

End Function

Public Function gGetRecs_UTE_UserTasks(slUTEStamp As String, ilUieCode As Integer, slForm_Module As String, tlUte() As UTE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "ute.eng") & ilUieCode
    slStamp = Trim$(Str$(ilUieCode))
    
    On Error GoTo gGetRecs_UTE_UserTasksErr
    ilRet = 0
    ilLowLimit = LBound(tlUte)
    If ilRet <> 0 Then
        slUTEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slUTEStamp <> "") Then
        sgSQLQuery = "SELECT Count(uteCode) FROM UTE_User_Tasks"
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlUte)) And (slStamp = slUTEStamp) Then
            rst.Close
            gGetRecs_UTE_UserTasks = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM UTE_User_Tasks Where uteUieCode = " & ilUieCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlUte(0 To 0) As UTE
    ilUpper = 0
    While Not rst.EOF
        tlUte(ilUpper).iCode = rst!uteCode
        tlUte(ilUpper).iUieCode = rst!uteUieCode
        tlUte(ilUpper).iTneCode = rst!uteTneCode
        tlUte(ilUpper).sTaskStatus = rst!uteTaskStatus
        tlUte(ilUpper).sUnused = ""
        ilUpper = ilUpper + 1
        ReDim Preserve tlUte(0 To ilUpper) As UTE
        rst.MoveNext
    Wend
    slUTEStamp = slStamp
    rst.Close
    gGetRecs_UTE_UserTasks = True
    Exit Function
    
gGetRecs_UTE_UserTasksErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_UTE_UserTasks = False
    Exit Function

End Function


Public Function gGetRecs_EPE_EventProperties(slEPEStamp As String, ilEteCode As Integer, slForm_Module As String, tlEPE() As EPE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "epe.eng") & ilEteCode
    slStamp = Trim$(Str$(ilEteCode))
    
    On Error GoTo gGetRecs_EPE_EventPropertiesErr
    ilRet = 0
    ilLowLimit = LBound(tlEPE)
    If ilRet <> 0 Then
        slEPEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slEPEStamp <> "") Then
        sgSQLQuery = "SELECT Count(epeCode) FROM EPE_Event_Properties WHERE epeEteCode = " & ilEteCode
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlEPE)) And (slStamp = slEPEStamp) Then
            rst.Close
            gGetRecs_EPE_EventProperties = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM EPE_Event_Properties WHERE epeEteCode = " & ilEteCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlEPE(0 To 0) As EPE
    ilUpper = 0
    While Not rst.EOF
        tlEPE(ilUpper).iCode = rst!epeCode
        tlEPE(ilUpper).iEteCode = rst!epeEteCode
        tlEPE(ilUpper).sType = rst!epeType
        tlEPE(ilUpper).sBus = rst!epeBus
        tlEPE(ilUpper).sBusControl = rst!epeBusControl
        tlEPE(ilUpper).sTime = rst!epeTime
        tlEPE(ilUpper).sStartType = rst!epeStartType
        tlEPE(ilUpper).sFixedTime = rst!epeFixedTime
        tlEPE(ilUpper).sEndType = rst!epeEndType
        tlEPE(ilUpper).sDuration = rst!epeDuration
        tlEPE(ilUpper).sMaterialType = rst!epeMaterialType
        tlEPE(ilUpper).sAudioName = rst!epeAudioName
        tlEPE(ilUpper).sAudioItemID = rst!epeAudioItemID
        tlEPE(ilUpper).sAudioISCI = rst!epeAudioISCI
        tlEPE(ilUpper).sAudioControl = rst!epeAudioControl
        tlEPE(ilUpper).sBkupAudioName = rst!epeBkupAudioName
        tlEPE(ilUpper).sBkupAudioControl = rst!epeBkupAudioControl
        tlEPE(ilUpper).sProtAudioName = rst!epeProtAudioName
        tlEPE(ilUpper).sProtAudioItemID = rst!epeProtAudioItemID
        tlEPE(ilUpper).sProtAudioISCI = rst!epeProtAudioISCI
        tlEPE(ilUpper).sProtAudioControl = rst!epeProtAudioControl
        tlEPE(ilUpper).sRelay1 = rst!epeRelay1
        tlEPE(ilUpper).sRelay2 = rst!epeRelay2
        tlEPE(ilUpper).sFollow = rst!epeFollow
        tlEPE(ilUpper).sSilenceTime = rst!epeSilenceTime
        tlEPE(ilUpper).sSilence1 = rst!epeSilence1
        tlEPE(ilUpper).sSilence2 = rst!epeSilence2
        tlEPE(ilUpper).sSilence3 = rst!epeSilence3
        tlEPE(ilUpper).sSilence4 = rst!epeSilence4
        tlEPE(ilUpper).sStartNetcue = rst!epeStartNetcue
        tlEPE(ilUpper).sStopNetcue = rst!epeStopNetcue
        tlEPE(ilUpper).sTitle1 = rst!epeTitle1
        tlEPE(ilUpper).sTitle2 = rst!epeTitle2
        tlEPE(ilUpper).sABCFormat = rst!epeABCFormat
        tlEPE(ilUpper).sABCPgmCode = rst!epeABCPgmCode
        tlEPE(ilUpper).sABCXDSMode = rst!epeABCXDSMode
        tlEPE(ilUpper).sABCRecordItem = rst!epeABCRecordItem
        tlEPE(ilUpper).sUnused = rst!epeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlEPE(0 To ilUpper) As EPE
        rst.MoveNext
    Wend
    slEPEStamp = slStamp
    rst.Close
    gGetRecs_EPE_EventProperties = True
    Exit Function
    
gGetRecs_EPE_EventPropertiesErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_EPE_EventProperties = False
    Exit Function

End Function

Public Function gGetRecs_DBE_DayBusSel(slDBEStamp As String, llDheCode As Long, slForm_Module As String, tlDBE() As DBE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "dbe.eng") & llDheCode
    slStamp = Trim$(Str$(llDheCode))
    
    On Error GoTo gGetRecs_DBE_DayBusSelErr
    ilRet = 0
    ilLowLimit = LBound(tlDBE)
    If ilRet <> 0 Then
        slDBEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slDBEStamp <> "") Then
        sgSQLQuery = "SELECT Count(dbeCode) FROM DBE_Day_Bus_Sel WHERE dbeDheCode = " & llDheCode
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlDBE)) And (slStamp = slDBEStamp) Then
            rst.Close
            gGetRecs_DBE_DayBusSel = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM DBE_Day_Bus_Sel WHERE dbeDheCode = " & llDheCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlDBE(0 To 0) As DBE
    ilUpper = 0
    While Not rst.EOF
    
        tlDBE(ilUpper).lCode = rst!dbeCode
        tlDBE(ilUpper).sType = rst!dbeType
        tlDBE(ilUpper).lDheCode = rst!dbeDheCode
        tlDBE(ilUpper).iBdeCode = rst!dbeBdeCode
        tlDBE(ilUpper).iBgeCode = rst!dbeBgeCode
        tlDBE(ilUpper).sUnused = rst!dbeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlDBE(0 To ilUpper) As DBE
        rst.MoveNext
    Wend
    slDBEStamp = slStamp
    rst.Close
    gGetRecs_DBE_DayBusSel = True
    Exit Function
    
gGetRecs_DBE_DayBusSelErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_DBE_DayBusSel = False
    Exit Function

End Function

Public Function gGetRecs_EBE_EventBusSel(slEBEStamp As String, llDeeCode As Long, slForm_Module As String, tlEBE() As EBE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim ilUpper As Integer
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "ebe.eng") & llDeeCode
    slStamp = Trim$(Str$(llDeeCode))
    
    On Error GoTo gGetRecs_EBE_EventBusSelErr
    ilRet = 0
    ilLowLimit = LBound(tlEBE)
    If ilRet <> 0 Then
        slEBEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slEBEStamp <> "") Then
        sgSQLQuery = "SELECT Count(ebeCode) FROM EBE_Event_Bus_Sel WHERE ebeDeeCode = " & llDeeCode
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlEBE)) And (slStamp = slEBEStamp) Then
            rst.Close
            gGetRecs_EBE_EventBusSel = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM EBE_Event_Bus_Sel WHERE ebeDeeCode = " & llDeeCode
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlEBE(0 To 0) As EBE
    ilUpper = 0
    While Not rst.EOF
    
        tlEBE(ilUpper).lCode = rst!ebeCode
        tlEBE(ilUpper).lDeeCode = rst!ebeDeeCode
        tlEBE(ilUpper).iBdeCode = rst!ebeBdeCode
        tlEBE(ilUpper).sUnused = rst!ebeUnused
        ilUpper = ilUpper + 1
        ReDim Preserve tlEBE(0 To ilUpper) As EBE
        rst.MoveNext
    Wend
    slEBEStamp = slStamp
    rst.Close
    gGetRecs_EBE_EventBusSel = True
    Exit Function
    
gGetRecs_EBE_EventBusSelErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_EBE_EventBusSel = False
    Exit Function

End Function

Public Function gGetRecs_SEE_ScheduleEvents(slSEEStamp As String, llSheCode As Long, slForm_Module As String, tlSEE() As SEE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "dee.eng") & llDheCode
    slStamp = Trim$(Str$(llSheCode))
    
    On Error GoTo gGetRecs_SEE_ScheduleEventsErr
    ilRet = 0
    ilLowLimit = LBound(tlSEE)
    If ilRet <> 0 Then
        slSEEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slSEEStamp <> "") Then
        sgSQLQuery = "SELECT Count(seeCode) FROM SEE_Schedule_Events Where seeAction <> 'R' and seeSheCode = " & llSheCode
        Set rst = cnn.Execute(sgSQLQuery)
        If (rst(0).Value = UBound(tlSEE)) And (slStamp = slSEEStamp) Then
            rst.Close
            gGetRecs_SEE_ScheduleEvents = True
            Exit Function
        End If
    End If
    sgSQLQuery = "SELECT * FROM SEE_Schedule_Events Where seeAction <> 'R' and seeSheCode = " & llSheCode & " ORDER BY seeTime, seeBDECode, seeSpotTime"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlSEE(0 To 30000) As SEE
    llUpper = 0
    While Not rst.EOF
        tlSEE(llUpper).lCode = rst!seeCode
        tlSEE(llUpper).lSheCode = rst!seeshecode
        tlSEE(llUpper).sAction = rst!seeAction
        tlSEE(llUpper).lDeeCode = rst!seeDeeCode
        tlSEE(llUpper).iBdeCode = rst!seeBdeCode
        tlSEE(llUpper).iBusCceCode = rst!seeBusCceCode
        tlSEE(llUpper).sSchdType = rst!seeSchdType
        tlSEE(llUpper).iEteCode = rst!seeEteCode
        tlSEE(llUpper).lTime = rst!seeTime
        tlSEE(llUpper).iStartTteCode = rst!seeStartTteCode
        tlSEE(llUpper).sFixedTime = rst!seeFixedTime
        tlSEE(llUpper).iEndTteCode = rst!seeEndTteCode
        tlSEE(llUpper).lDuration = rst!seeDuration
        tlSEE(llUpper).iMteCode = rst!seeMteCode
        tlSEE(llUpper).iAudioAseCode = rst!seeAudioAseCode
        tlSEE(llUpper).sAudioItemID = rst!seeAudioItemID
        tlSEE(llUpper).sAudioItemIDChk = rst!seeAudioItemIDChk
        tlSEE(llUpper).sAudioISCI = rst!seeAudioISCI
        tlSEE(llUpper).iAudioCceCode = rst!seeAudioCceCode
        tlSEE(llUpper).iBkupAneCode = rst!seeBkupAneCode
        tlSEE(llUpper).iBkupCceCode = rst!seeBkupCceCode
        tlSEE(llUpper).iProtAneCode = rst!seeProtAneCode
        tlSEE(llUpper).sProtItemID = rst!seeProtItemID
        tlSEE(llUpper).sProtItemIDChk = rst!seeProtItemIDChk
        tlSEE(llUpper).sProtISCI = rst!seeProtISCI
        tlSEE(llUpper).iProtCceCode = rst!seeProtCceCode
        tlSEE(llUpper).i1RneCode = rst!see1RneCode
        tlSEE(llUpper).i2RneCode = rst!see2RneCode
        tlSEE(llUpper).iFneCode = rst!seeFneCode
        tlSEE(llUpper).lSilenceTime = rst!seeSilenceTime
        tlSEE(llUpper).i1SceCode = rst!see1SceCode
        tlSEE(llUpper).i2SceCode = rst!see2SceCode
        tlSEE(llUpper).i3SceCode = rst!see3SceCode
        tlSEE(llUpper).i4SceCode = rst!see4SceCode
        tlSEE(llUpper).iStartNneCode = rst!seeStartNneCode
        tlSEE(llUpper).iEndNneCode = rst!seeEndNneCode
        tlSEE(llUpper).l1CteCode = rst!see1CteCode
        tlSEE(llUpper).l2CteCode = rst!see2CteCode
        tlSEE(llUpper).lAreCode = rst!seeAreCode
        tlSEE(llUpper).lSpotTime = rst!seeSpotTime
        tlSEE(llUpper).lEventID = rst!seeEventID
        tlSEE(llUpper).sAsAirStatus = rst!seeAsAirStatus
        tlSEE(llUpper).sSentStatus = rst!seeSentStatus
        tlSEE(llUpper).sSentDate = Format$(rst!seeSentDate, sgShowDateForm)
        tlSEE(llUpper).sIgnoreConflicts = rst!seeIgnoreConflicts
        tlSEE(llUpper).lDheCode = rst!seeDheCode
        tlSEE(llUpper).lOrigDHECode = rst!seeOrigDHECode
        tlSEE(llUpper).sInsertFlag = "N"    'Temporary flag used in Schedule Definition only
        tlSEE(llUpper).sABCFormat = rst!seeABCFormat
        tlSEE(llUpper).sABCPgmCode = rst!seeABCPgmCode
        tlSEE(llUpper).sABCXDSMode = rst!seeABCXDSMode
        tlSEE(llUpper).sABCRecordItem = rst!seeABCRecordItem
        tlSEE(llUpper).sUnused = ""
        'Field not part of record
        tlSEE(llUpper).lAvailLength = tlSEE(llUpper).lDuration
        llUpper = llUpper + 1
        If llUpper > UBound(tlSEE) Then
            ReDim Preserve tlSEE(0 To UBound(tlSEE) + 1000) As SEE
        End If
        rst.MoveNext
    Wend
    ReDim Preserve tlSEE(0 To llUpper) As SEE
    slSEEStamp = slStamp
    rst.Close
    gGetRecs_SEE_ScheduleEvents = True
    Exit Function
    
gGetRecs_SEE_ScheduleEventsErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_SEE_ScheduleEvents = False
    Exit Function

End Function

Public Function gGetRecs_SEE_ScheduleEventsAPI(hlSEE As Integer, slSEEStamp As String, llDheCode As Long, llSheCode As Long, slForm_Module As String, tlSEE() As SEE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim slTime As String
    Dim slBDECode As String
    Dim slSpotTime As String
    Dim ilSEERecLen As Integer
    Dim tlSEESrchKey As LONGKEY0
    Dim llLoop As Long
    Dim ilOffset As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "dee.eng") & llDheCode
    slStamp = Trim$(Str$(llSheCode))
    
    On Error GoTo gGetRecs_SEE_ScheduleEventsAPIErr
    ilRet = 0
    ilLowLimit = LBound(tlSEE)
    If ilRet <> 0 Then
        slSEEStamp = ""
    End If
    On Error GoTo ErrHand
    'If (slSEEStamp <> "") Then
    '    sgSQLQuery = "SELECT Count(seeCode) FROM SEE_Schedule_Events Where seeAction <> 'R' and seeSheCode = " & llSheCode
    '    Set rst = cnn.Execute(sgSQLQuery)
    '    If (rst(0).Value = UBound(tlSEE)) And (slStamp = slSEEStamp) Then
    '        gGetRecs_SEE_ScheduleEventsAPI = True
    '        Exit Function
    '    End If
    'End If
    'sgSQLQuery = "SELECT * FROM SEE_Schedule_Events Where seeAction <> 'R' and seeSheCode = " & llSheCode & " ORDER BY seeTime, seeBDECode, seeSpotTime"
    'Set rst = cnn.Execute(sgSQLQuery)
    'ReDim tlSEE(0 To 30000) As SEE
    'llUpper = 0
    'While Not rst.EOF
    '    tlSEE(llUpper).lCode = rst!seeCode
    '    tlSEE(llUpper).lSheCode = rst!seeshecode
    '    tlSEE(llUpper).sAction = rst!seeAction
    '    tlSEE(llUpper).lDeeCode = rst!seeDeeCode
    '    tlSEE(llUpper).iBdeCode = rst!seeBdeCode
    '    tlSEE(llUpper).iBusCceCode = rst!seeBusCceCode
    '    tlSEE(llUpper).sSchdType = rst!seeSchdType
    '    tlSEE(llUpper).iEteCode = rst!seeEteCode
    '    tlSEE(llUpper).lTime = rst!seeTime
    '    tlSEE(llUpper).iStartTteCode = rst!seeStartTteCode
    '    tlSEE(llUpper).sFixedTime = rst!seeFixedTime
    '    tlSEE(llUpper).iEndTteCode = rst!seeEndTteCode
    '    tlSEE(llUpper).lDuration = rst!seeDuration
    '    tlSEE(llUpper).iMteCode = rst!seeMteCode
    '    tlSEE(llUpper).iAudioAseCode = rst!seeAudioAseCode
    '    tlSEE(llUpper).sAudioItemID = rst!seeAudioItemID
    '    tlSEE(llUpper).sAudioItemIDChk = rst!seeAudioItemIDChk
    '    tlSEE(llUpper).sAudioISCI = rst!seeAudioISCI
    '    tlSEE(llUpper).iAudioCceCode = rst!seeAudioCceCode
    '    tlSEE(llUpper).iBkupAneCode = rst!seeBkupAneCode
    '    tlSEE(llUpper).iBkupCceCode = rst!seeBkupCceCode
    '    tlSEE(llUpper).iProtAneCode = rst!seeProtAneCode
    '    tlSEE(llUpper).sProtItemID = rst!seeProtItemID
    '    tlSEE(llUpper).sProtItemIDChk = rst!seeProtItemIDChk
    '    tlSEE(llUpper).sProtISCI = rst!seeProtISCI
    '    tlSEE(llUpper).iProtCceCode = rst!seeProtCceCode
    '    tlSEE(llUpper).i1RneCode = rst!see1RneCode
    '    tlSEE(llUpper).i2RneCode = rst!see2RneCode
    '    tlSEE(llUpper).iFneCode = rst!seeFneCode
    '    tlSEE(llUpper).lSilenceTime = rst!seeSilenceTime
    '    tlSEE(llUpper).i1SceCode = rst!see1SceCode
    '    tlSEE(llUpper).i2SceCode = rst!see2SceCode
    '    tlSEE(llUpper).i3SceCode = rst!see3SceCode
    '    tlSEE(llUpper).i4SceCode = rst!see4SceCode
    '    tlSEE(llUpper).iStartNneCode = rst!seeStartNneCode
    '    tlSEE(llUpper).iEndNneCode = rst!seeEndNneCode
    '    tlSEE(llUpper).l1CteCode = rst!see1CteCode
    '    tlSEE(llUpper).l2CteCode = rst!see2CteCode
    '    tlSEE(llUpper).lAreCode = rst!seeAreCode
    '    tlSEE(llUpper).lSpotTime = rst!seeSpotTime
    '    tlSEE(llUpper).lEventID = rst!seeEventID
     '   tlSEE(llUpper).sAsAirStatus = rst!seeAsAirStatus
    '    tlSEE(llUpper).sSentStatus = rst!seeSentStatus
    '    tlSEE(llUpper).sSentDate = Format$(rst!seeSentDate, sgShowDateForm)
    '    tlSEE(llUpper).sIgnoreConflicts = rst!seeIgnoreConflicts
    '    tlSEE(llUpper).lDheCode = rst!seeDheCode
    '    tlSEE(llUpper).lOrigDheCode = rst!seeOrigDHECode
    '    tlSEE(llUpper).sInsertFlag = "N"    'Temporary flag used in Schedule Definition only
    '    tlSEE(llUpper).sUnused = ""
    '    'Field not part of record
    '    tlSEE(llUpper).lAvailLength = tlSEE(llUpper).lDuration
    '    llUpper = llUpper + 1
    '    If llUpper > UBound(tlSEE) Then
    '        ReDim Preserve tlSEE(0 To UBound(tlSEE) + 1000) As SEE
    '    End If
    '    rst.MoveNext
    'Wend
    ReDim tlSEESort(0 To 30000) As SEESORT
    llUpper = 0
    ilSEERecLen = Len(tlSEESort(0).tSEEAPI)
    tlSEESrchKey.lCode = llSheCode
    ilRet = btrGetEqual(hlSEE, tlSEESort(llUpper).tSEEAPI, ilSEERecLen, tlSEESrchKey, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    If (ilRet = BTRV_ERR_NONE) And (tlSEESort(llUpper).tSEEAPI.lSheCode = llSheCode) Then
        llNoRec = gExtNoRec(ilSEERecLen) 'btrRecords(hlAnf) 'Obtain number of records
        btrExtClear hlSEE   'Clear any previous extend operation
        Call btrExtSetBounds(hlSEE, llNoRec, -1, "UC", "SEE", "") 'Set extract limits (all records)
        If llDheCode > 0 Then
            ilOffset = gFieldOffset("SEE", "seeDHECode")
            ilRet = btrExtAddLogicConst(hlSEE, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, llDheCode, 4)
        End If
        ilOffset = gFieldOffset("SEE", "seeSHECode")
        ilRet = btrExtAddLogicConst(hlSEE, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, llSheCode, 4)
        ilOffset = 0
        ilRet = btrExtAddField(hlSEE, ilOffset, ilSEERecLen)  'Extract iCode field
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrExtGetNext(hlSEE, tlSEESort(llUpper).tSEEAPI, ilSEERecLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                    ilSEERecLen = Len(tlSEESort(0).tSEEAPI)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlSEE, tlSEESort(llUpper).tSEEAPI, ilSEERecLen, llRecPos)
                    Loop
                    Do While ilRet = BTRV_ERR_NONE
                        If tlSEESort(llUpper).tSEEAPI.sAction <> "R" Then
                            slTime = Trim$(Str$(tlSEESort(llUpper).tSEEAPI.lTime))
                            Do While Len(slTime) < 10
                                slTime = "0" & slTime
                            Loop
                            slBDECode = Trim$(Str$(tlSEESort(llUpper).tSEEAPI.iBdeCode))
                            Do While Len(slBDECode) < 5
                                slBDECode = "0" & slBDECode
                            Loop
                            slSpotTime = Trim$(Str$(tlSEESort(llUpper).tSEEAPI.lSpotTime))
                            Do While Len(slSpotTime) < 10
                                slSpotTime = "0" & slSpotTime
                            Loop
                            tlSEESort(llUpper).sKey = slTime & slBDECode & slSpotTime
                            llUpper = llUpper + 1
                            If llUpper > UBound(tlSEESort) Then
                                ReDim Preserve tlSEESort(0 To UBound(tlSEESort) + 1000) As SEESORT
                            End If
                        End If
                        ilRet = btrExtGetNext(hlSEE, tlSEESort(llUpper).tSEEAPI, ilSEERecLen, llRecPos)
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hlSEE, tlSEESort(llUpper).tSEEAPI, ilSEERecLen, llRecPos)
                        Loop
                    Loop
                End If
            End If
        End If
        
    End If
    ReDim Preserve tlSEESort(0 To llUpper) As SEESORT
    'Sort by Time
    If UBound(tlSEESort) - 1 > 0 Then
        ArraySortTyp fnAV(tlSEESort(), 0), UBound(tlSEESort), 0, LenB(tlSEESort(0)), 0, LenB(tlSEESort(0).sKey), 0
    End If
    ReDim tlSEE(0 To UBound(tlSEESort)) As SEE
    For llLoop = 0 To UBound(tlSEESort) - 1 Step 1
        LSet tlSEE(llLoop) = tlSEESort(llLoop).tSEEAPI
        tlSEE(llLoop).sSentDate = Format$(tlSEE(llLoop).sSentDate, sgShowDateForm)
        If llDheCode > 0 Then
            tlSEE(llLoop).sInsertFlag = ""
        Else
            tlSEE(llLoop).sInsertFlag = "N"
        End If
        tlSEE(llLoop).lAvailLength = tlSEE(llLoop).lDuration
    Next llLoop
    Erase tlSEESort
    slSEEStamp = slStamp
    gGetRecs_SEE_ScheduleEventsAPI = True
    Exit Function
    
gGetRecs_SEE_ScheduleEventsAPIErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    gGetRecs_SEE_ScheduleEventsAPI = False
    Exit Function

End Function

Public Function gGetRecs_SEE_ScheduleEventsAPIWithFilter(hlSEE As Integer, slSEEStamp As String, llDheCode As Long, llSheCode As Long, slForm_Module As String, tlSEE() As SEE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim slTime As String
    Dim slBDECode As String
    Dim slSpotTime As String
    Dim ilSEERecLen As Integer
    Dim tlSEESrchKey As LONGKEY0
    Dim tlSEESrchKey4 As SEEKEY4
    Dim tlSEESrchKey5 As SEEKEY5
    Dim llLoop As Long
    Dim ilOffset As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilFilter As Integer
    Dim ilField As Integer
    Dim ilCountB1 As Integer
    Dim ilCountB2 As Integer
    Dim ilCountB3 As Integer
    Dim ilKey As Integer
    Dim ilBdeCode As Integer
    Dim llSTime As Long
    Dim llETime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim blLastTermAdded As Boolean
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "dee.eng") & llDheCode
    slStamp = Trim$(Str$(llSheCode))
    
    On Error GoTo gGetRecs_SEE_ScheduleEventsAPIWithFilterErr
    ilRet = 0
    ilLowLimit = LBound(tlSEE)
    If ilRet <> 0 Then
        slSEEStamp = ""
    End If
    On Error GoTo ErrHand
    'If (slSEEStamp <> "") Then
    '    sgSQLQuery = "SELECT Count(seeCode) FROM SEE_Schedule_Events Where seeAction <> 'R' and seeSheCode = " & llSheCode
    '    Set rst = cnn.Execute(sgSQLQuery)
    '    If (rst(0).Value = UBound(tlSEE)) And (slStamp = slSEEStamp) Then
    '        gGetRecs_SEE_ScheduleEventsAPIWithFilter = True
    '        Exit Function
    '    End If
    'End If
    'sgSQLQuery = "SELECT * FROM SEE_Schedule_Events Where seeAction <> 'R' and seeSheCode = " & llSheCode & " ORDER BY seeTime, seeBDECode, seeSpotTime"
    'Set rst = cnn.Execute(sgSQLQuery)
    'ReDim tlSEE(0 To 30000) As SEE
    'llUpper = 0
    'While Not rst.EOF
    '    tlSEE(llUpper).lCode = rst!seeCode
    '    tlSEE(llUpper).lSheCode = rst!seeshecode
    '    tlSEE(llUpper).sAction = rst!seeAction
    '    tlSEE(llUpper).lDeeCode = rst!seeDeeCode
    '    tlSEE(llUpper).iBdeCode = rst!seeBdeCode
    '    tlSEE(llUpper).iBusCceCode = rst!seeBusCceCode
    '    tlSEE(llUpper).sSchdType = rst!seeSchdType
    '    tlSEE(llUpper).iEteCode = rst!seeEteCode
    '    tlSEE(llUpper).lTime = rst!seeTime
    '    tlSEE(llUpper).iStartTteCode = rst!seeStartTteCode
    '    tlSEE(llUpper).sFixedTime = rst!seeFixedTime
    '    tlSEE(llUpper).iEndTteCode = rst!seeEndTteCode
    '    tlSEE(llUpper).lDuration = rst!seeDuration
    '    tlSEE(llUpper).iMteCode = rst!seeMteCode
    '    tlSEE(llUpper).iAudioAseCode = rst!seeAudioAseCode
    '    tlSEE(llUpper).sAudioItemID = rst!seeAudioItemID
    '    tlSEE(llUpper).sAudioItemIDChk = rst!seeAudioItemIDChk
    '    tlSEE(llUpper).sAudioISCI = rst!seeAudioISCI
    '    tlSEE(llUpper).iAudioCceCode = rst!seeAudioCceCode
    '    tlSEE(llUpper).iBkupAneCode = rst!seeBkupAneCode
    '    tlSEE(llUpper).iBkupCceCode = rst!seeBkupCceCode
    '    tlSEE(llUpper).iProtAneCode = rst!seeProtAneCode
    '    tlSEE(llUpper).sProtItemID = rst!seeProtItemID
    '    tlSEE(llUpper).sProtItemIDChk = rst!seeProtItemIDChk
    '    tlSEE(llUpper).sProtISCI = rst!seeProtISCI
    '    tlSEE(llUpper).iProtCceCode = rst!seeProtCceCode
    '    tlSEE(llUpper).i1RneCode = rst!see1RneCode
    '    tlSEE(llUpper).i2RneCode = rst!see2RneCode
    '    tlSEE(llUpper).iFneCode = rst!seeFneCode
    '    tlSEE(llUpper).lSilenceTime = rst!seeSilenceTime
    '    tlSEE(llUpper).i1SceCode = rst!see1SceCode
    '    tlSEE(llUpper).i2SceCode = rst!see2SceCode
    '    tlSEE(llUpper).i3SceCode = rst!see3SceCode
    '    tlSEE(llUpper).i4SceCode = rst!see4SceCode
    '    tlSEE(llUpper).iStartNneCode = rst!seeStartNneCode
    '    tlSEE(llUpper).iEndNneCode = rst!seeEndNneCode
    '    tlSEE(llUpper).l1CteCode = rst!see1CteCode
    '    tlSEE(llUpper).l2CteCode = rst!see2CteCode
    '    tlSEE(llUpper).lAreCode = rst!seeAreCode
    '    tlSEE(llUpper).lSpotTime = rst!seeSpotTime
    '    tlSEE(llUpper).lEventID = rst!seeEventID
     '   tlSEE(llUpper).sAsAirStatus = rst!seeAsAirStatus
    '    tlSEE(llUpper).sSentStatus = rst!seeSentStatus
    '    tlSEE(llUpper).sSentDate = Format$(rst!seeSentDate, sgShowDateForm)
    '    tlSEE(llUpper).sIgnoreConflicts = rst!seeIgnoreConflicts
    '    tlSEE(llUpper).lDheCode = rst!seeDheCode
    '    tlSEE(llUpper).lOrigDheCode = rst!seeOrigDHECode
    '    tlSEE(llUpper).sInsertFlag = "N"    'Temporary flag used in Schedule Definition only
    '    tlSEE(llUpper).sUnused = ""
    '    'Field not part of record
    '    tlSEE(llUpper).lAvailLength = tlSEE(llUpper).lDuration
    '    llUpper = llUpper + 1
    '    If llUpper > UBound(tlSEE) Then
    '        ReDim Preserve tlSEE(0 To UBound(tlSEE) + 1000) As SEE
    '    End If
    '    rst.MoveNext
    'Wend
    ReDim tlSEESort(0 To 30000) As SEESORT
    ilKey = INDEXKEY1
    
    llStartTime = -1
    llEndTime = -1
    For ilFilter = LBound(tgFilterValues) To UBound(tgFilterValues) - 1 Step 1
        For ilField = LBound(tgFilterFields) To UBound(tgFilterFields) - 1 Step 1
            If tgFilterFields(ilField).sFieldName = tgFilterValues(ilFilter).sFieldName Then
                If tgFilterFields(ilField).iFieldType = 6 Then  'Time
                    llSTime = gStrTimeInTenthToLong(tgFilterValues(ilFilter).sValue, False)
                    llETime = gStrTimeInTenthToLong(tgFilterValues(ilFilter).sValue, True)
                    If tgFilterValues(ilFilter).iOperator = 1 Then   'Equal Match
                        If (llStartTime = -1) Or (llSTime < llStartTime) Then
                            llStartTime = llSTime
                        End If
                        If (llEndTime = -1) Or (llETime > llEndTime) Then
                            llEndTime = llETime
                        End If
                    ElseIf tgFilterValues(ilFilter).iOperator = 3 Then   'GT
                        If (llStartTime = -1) Or (llSTime < llStartTime) Then
                            llStartTime = llSTime
                        End If
                    ElseIf tgFilterValues(ilFilter).iOperator = 4 Then   'LT
                        If (llEndTime = -1) Or (llETime > llEndTime) Then
                            llEndTime = llETime
                        End If
                    ElseIf tgFilterValues(ilFilter).iOperator = 5 Then   'GTE
                        If (llStartTime = -1) Or (llSTime < llStartTime) Then
                            llStartTime = llSTime
                        End If
                    ElseIf tgFilterValues(ilFilter).iOperator = 6 Then   'LTE
                        If (llEndTime = -1) Or (llETime > llEndTime) Then
                            llEndTime = llETime
                        End If
                    End If
                End If
            End If
        Next ilField
    Next ilFilter
    
    ilCountB1 = 0
    ilCountB2 = 0
    For ilFilter = LBound(tgFilterValues) To UBound(tgFilterValues) - 1 Step 1
        For ilField = LBound(tgFilterFields) To UBound(tgFilterFields) - 1 Step 1
            If tgFilterFields(ilField).sFieldName = tgFilterValues(ilFilter).sFieldName Then
                If tgFilterFields(ilField).iFieldType = 5 Then  'List
                    If Trim$(tgFilterFields(ilField).sListFile) = "BDE" Then
                        If tgFilterValues(ilFilter).iOperator = 1 Then   'Equal Match
                            ilCountB1 = ilCountB1 + 1
                            ilBdeCode = CInt(tgFilterValues(ilFilter).lCode)
                        Else
                            ilCountB2 = ilCountB2 + 1
                        End If
                    End If
                End If
            End If
        Next ilField
    Next ilFilter
    If ((ilCountB1 = 1) And (ilCountB2 = 0)) Then
        ilKey = INDEXKEY5
        tlSEESrchKey5.lSheCode = llSheCode
        tlSEESrchKey5.iBdeCode = ilBdeCode
    ElseIf ((ilCountB1 = 0) And (ilCountB2 = 1)) Then
        ilKey = INDEXKEY1
        tlSEESrchKey.lCode = llSheCode
    ElseIf llStartTime <> -1 Then
        ilKey = INDEXKEY4
        tlSEESrchKey4.lSheCode = llSheCode
        tlSEESrchKey4.lTime = llStartTime
    Else
        ilKey = INDEXKEY1
        tlSEESrchKey.lCode = llSheCode
    End If
    
    llUpper = 0
    ilSEERecLen = Len(tlSEESort(0).tSEEAPI)
    If ilKey = INDEXKEY5 Then
        ilRet = btrGetEqual(hlSEE, tlSEESort(llUpper).tSEEAPI, ilSEERecLen, tlSEESrchKey5, ilKey, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    ElseIf ilKey = INDEXKEY4 Then
        ilRet = btrGetEqual(hlSEE, tlSEESort(llUpper).tSEEAPI, ilSEERecLen, tlSEESrchKey4, ilKey, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Else
        ilRet = btrGetEqual(hlSEE, tlSEESort(llUpper).tSEEAPI, ilSEERecLen, tlSEESrchKey, ilKey, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    End If
    If (ilRet = BTRV_ERR_NONE) And (tlSEESort(llUpper).tSEEAPI.lSheCode = llSheCode) Then
        llNoRec = gExtNoRec(ilSEERecLen) 'btrRecords(hlAnf) 'Obtain number of records
        btrExtClear hlSEE   'Clear any previous extend operation
        Call btrExtSetBounds(hlSEE, llNoRec, -1, "UC", "SEE", "") 'Set extract limits (all records)
        'Using GTE and LTE because the filter within SchDef will handle the GT and LT case
        If llStartTime <> -1 Then
            ilOffset = gFieldOffset("SEE", "seeTime")
            ilRet = btrExtAddLogicConst(hlSEE, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_AND, llStartTime, 4)
        End If
        If llEndTime <> -1 Then
            ilOffset = gFieldOffset("SEE", "seeTime")
            ilRet = btrExtAddLogicConst(hlSEE, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_AND, llEndTime, 4)
        End If
        blLastTermAdded = False
        If (ilCountB1 >= 1) And (ilCountB2 = 0) Then
            blLastTermAdded = True
            If llDheCode > 0 Then
                ilOffset = gFieldOffset("SEE", "seeDHECode")
                ilRet = btrExtAddLogicConst(hlSEE, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, llDheCode, 4)
            End If
            ilOffset = gFieldOffset("SEE", "seeSHECode")
            ilRet = btrExtAddLogicConst(hlSEE, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, llSheCode, 4)
            ilCountB3 = 0
            For ilFilter = LBound(tgFilterValues) To UBound(tgFilterValues) - 1 Step 1
                For ilField = LBound(tgFilterFields) To UBound(tgFilterFields) - 1 Step 1
                    If tgFilterFields(ilField).sFieldName = tgFilterValues(ilFilter).sFieldName Then
                        If tgFilterFields(ilField).iFieldType = 5 Then  'List
                            If Trim$(tgFilterFields(ilField).sListFile) = "BDE" Then
                                ilOffset = gFieldOffset("SEE", "seeBDECode")
                                ilCountB3 = ilCountB3 + 1
                                If ilCountB1 = ilCountB3 Then
                                    ilRet = btrExtAddLogicConst(hlSEE, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, CInt(tgFilterValues(ilFilter).lCode), 2)
                                Else
                                    ilRet = btrExtAddLogicConst(hlSEE, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_OR, CInt(tgFilterValues(ilFilter).lCode), 2)
                                End If
                            End If
                        End If
                    End If
                Next ilField
            Next ilFilter
        ElseIf (ilCountB1 = 0) And (ilCountB2 >= 1) Then
            For ilFilter = LBound(tgFilterValues) To UBound(tgFilterValues) - 1 Step 1
                For ilField = LBound(tgFilterFields) To UBound(tgFilterFields) - 1 Step 1
                    If tgFilterFields(ilField).sFieldName = tgFilterValues(ilFilter).sFieldName Then
                        If tgFilterFields(ilField).iFieldType = 5 Then  'List
                            If Trim$(tgFilterFields(ilField).sListFile) = "BDE" Then
                                ilOffset = gFieldOffset("SEE", "seeBDECode")
                                ilRet = btrExtAddLogicConst(hlSEE, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, CInt(tgFilterValues(ilFilter).lCode), 2)
                            End If
                        End If
                    End If
                Next ilField
            Next ilFilter
        End If

        If Not blLastTermAdded Then
            If llDheCode > 0 Then
                ilOffset = gFieldOffset("SEE", "seeDHECode")
                ilRet = btrExtAddLogicConst(hlSEE, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, llDheCode, 4)
            End If
            ilOffset = gFieldOffset("SEE", "seeSHECode")
            ilRet = btrExtAddLogicConst(hlSEE, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, llSheCode, 4)
        End If
        ilOffset = 0
        ilRet = btrExtAddField(hlSEE, ilOffset, ilSEERecLen)  'Extract iCode field
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrExtGetNext(hlSEE, tlSEESort(llUpper).tSEEAPI, ilSEERecLen, llRecPos)
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                If (ilRet = BTRV_ERR_NONE) Or (ilRet = BTRV_ERR_REJECT_COUNT) Then
                    ilSEERecLen = Len(tlSEESort(0).tSEEAPI)
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlSEE, tlSEESort(llUpper).tSEEAPI, ilSEERecLen, llRecPos)
                    Loop
                    Do While ilRet = BTRV_ERR_NONE
                        If tlSEESort(llUpper).tSEEAPI.sAction <> "R" Then
                            slTime = Trim$(Str$(tlSEESort(llUpper).tSEEAPI.lTime))
                            Do While Len(slTime) < 10
                                slTime = "0" & slTime
                            Loop
                            slBDECode = Trim$(Str$(tlSEESort(llUpper).tSEEAPI.iBdeCode))
                            Do While Len(slBDECode) < 5
                                slBDECode = "0" & slBDECode
                            Loop
                            slSpotTime = Trim$(Str$(tlSEESort(llUpper).tSEEAPI.lSpotTime))
                            Do While Len(slSpotTime) < 10
                                slSpotTime = "0" & slSpotTime
                            Loop
                            tlSEESort(llUpper).sKey = slTime & slBDECode & slSpotTime
                            llUpper = llUpper + 1
                            If llUpper > UBound(tlSEESort) Then
                                ReDim Preserve tlSEESort(0 To UBound(tlSEESort) + 1000) As SEESORT
                            End If
                        End If
                        ilRet = btrExtGetNext(hlSEE, tlSEESort(llUpper).tSEEAPI, ilSEERecLen, llRecPos)
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hlSEE, tlSEESort(llUpper).tSEEAPI, ilSEERecLen, llRecPos)
                        Loop
                    Loop
                End If
            End If
        End If
        
    End If
    ReDim Preserve tlSEESort(0 To llUpper) As SEESORT
    'Sort by Time
    If UBound(tlSEESort) - 1 > 0 Then
        ArraySortTyp fnAV(tlSEESort(), 0), UBound(tlSEESort), 0, LenB(tlSEESort(0)), 0, LenB(tlSEESort(0).sKey), 0
    End If
    ReDim tlSEE(0 To UBound(tlSEESort)) As SEE
    For llLoop = 0 To UBound(tlSEESort) - 1 Step 1
        LSet tlSEE(llLoop) = tlSEESort(llLoop).tSEEAPI
        tlSEE(llLoop).sSentDate = Format$(tlSEE(llLoop).sSentDate, sgShowDateForm)
        If llDheCode > 0 Then
            tlSEE(llLoop).sInsertFlag = ""
        Else
            tlSEE(llLoop).sInsertFlag = "N"
        End If
        tlSEE(llLoop).lAvailLength = tlSEE(llLoop).lDuration
    Next llLoop
    Erase tlSEESort
    slSEEStamp = slStamp
    gGetRecs_SEE_ScheduleEventsAPIWithFilter = True
    Exit Function
    
gGetRecs_SEE_ScheduleEventsAPIWithFilterErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    gGetRecs_SEE_ScheduleEventsAPIWithFilter = False
    Exit Function

End Function
Public Function gGetRecs_SEE_ScheduleEventsByDHEandSHE(slSEEStamp As String, llDheCode As Long, llSheCode As Long, slForm_Module As String, tlSEE() As SEE) As Integer
    Dim ilRet As Integer
    Dim slStamp As String
    Dim llUpper As Long
    Dim ilLowLimit As Integer
    Dim rst As ADODB.Recordset
    
    'slStamp = FileDateTime(sgDBPath & "dee.eng") & llDheCode
    slStamp = Trim$(Str$(llDheCode))
    
    On Error GoTo gGetRecs_SEE_ScheduleEventsByDHEErr
    ilRet = 0
    ilLowLimit = LBound(tlSEE)
    If ilRet <> 0 Then
        slSEEStamp = ""
    End If
    On Error GoTo ErrHand
    If (slSEEStamp <> "") Then
        'sgSQLQuery = "SELECT Count(seeCode) FROM SEE_Schedule_Events"
        'Set rst = cnn.Execute(sgSQLQuery)
        'If (rst(0).Value = UBound(tlSEE)) And (slStamp = slSEEStamp) Then
        '    gGetRecs_SEE_ScheduleEventsByDHEandSHE = True
        '    Exit Function
        'End If
    End If
    'sgSQLQuery = "SELECT * FROM SEE_Schedule_Events, SHE_Schedule_Header WHERE seeshecode = shecode and sheCode = " & llSheCode & " and seeDeeCode IN (SELECT seeDeeCode FROM SEE_Schedule_Events, DEE_Day_Event_Info, DHE_Day_Header_Info WHERE seeDeecode = deeCode and seeAction <> 'R' and deeDhecode = dheCode and dheCode = " & llDheCode & ") ORDER BY seeTime, seeBDECode, seeSpotTime"
    sgSQLQuery = "SELECT * FROM SEE_Schedule_Events WHERE seeshecode = " & llSheCode & " and seeDheCode = " & llDheCode & " and seeAction <> 'R' ORDER BY seeTime, seeBDECode, seeSpotTime"
    Set rst = cnn.Execute(sgSQLQuery)
    ReDim tlSEE(0 To 0) As SEE
    llUpper = 0
    While Not rst.EOF
        tlSEE(llUpper).lCode = rst!seeCode
        tlSEE(llUpper).lSheCode = rst!seeshecode
        tlSEE(llUpper).sAction = rst!seeAction
        tlSEE(llUpper).lDeeCode = rst!seeDeeCode
        tlSEE(llUpper).iBdeCode = rst!seeBdeCode
        tlSEE(llUpper).iBusCceCode = rst!seeBusCceCode
        tlSEE(llUpper).sSchdType = rst!seeSchdType
        tlSEE(llUpper).iEteCode = rst!seeEteCode
        tlSEE(llUpper).lTime = rst!seeTime
        tlSEE(llUpper).iStartTteCode = rst!seeStartTteCode
        tlSEE(llUpper).sFixedTime = rst!seeFixedTime
        tlSEE(llUpper).iEndTteCode = rst!seeEndTteCode
        tlSEE(llUpper).lDuration = rst!seeDuration
        tlSEE(llUpper).iMteCode = rst!seeMteCode
        tlSEE(llUpper).iAudioAseCode = rst!seeAudioAseCode
        tlSEE(llUpper).sAudioItemID = rst!seeAudioItemID
        tlSEE(llUpper).sAudioItemIDChk = rst!seeAudioItemIDChk
        tlSEE(llUpper).sAudioISCI = rst!seeAudioISCI
        tlSEE(llUpper).iAudioCceCode = rst!seeAudioCceCode
        tlSEE(llUpper).iBkupAneCode = rst!seeBkupAneCode
        tlSEE(llUpper).iBkupCceCode = rst!seeBkupCceCode
        tlSEE(llUpper).iProtAneCode = rst!seeProtAneCode
        tlSEE(llUpper).sProtItemID = rst!seeProtItemID
        tlSEE(llUpper).sProtItemIDChk = rst!seeProtItemIDChk
        tlSEE(llUpper).sProtISCI = rst!seeProtISCI
        tlSEE(llUpper).iProtCceCode = rst!seeProtCceCode
        tlSEE(llUpper).i1RneCode = rst!see1RneCode
        tlSEE(llUpper).i2RneCode = rst!see2RneCode
        tlSEE(llUpper).iFneCode = rst!seeFneCode
        tlSEE(llUpper).lSilenceTime = rst!seeSilenceTime
        tlSEE(llUpper).i1SceCode = rst!see1SceCode
        tlSEE(llUpper).i2SceCode = rst!see2SceCode
        tlSEE(llUpper).i3SceCode = rst!see3SceCode
        tlSEE(llUpper).i4SceCode = rst!see4SceCode
        tlSEE(llUpper).iStartNneCode = rst!seeStartNneCode
        tlSEE(llUpper).iEndNneCode = rst!seeEndNneCode
        tlSEE(llUpper).l1CteCode = rst!see1CteCode
        tlSEE(llUpper).l2CteCode = rst!see2CteCode
        tlSEE(llUpper).lAreCode = rst!seeAreCode
        tlSEE(llUpper).lSpotTime = rst!seeSpotTime
        tlSEE(llUpper).lEventID = rst!seeEventID
        tlSEE(llUpper).sAsAirStatus = rst!seeAsAirStatus
        tlSEE(llUpper).sSentStatus = rst!seeSentStatus
        tlSEE(llUpper).sSentDate = Format$(rst!seeSentDate, sgShowDateForm)
        tlSEE(llUpper).sIgnoreConflicts = rst!seeIgnoreConflicts
        tlSEE(llUpper).lDheCode = rst!seeDheCode
        tlSEE(llUpper).lOrigDHECode = rst!seeOrigDHECode
        tlSEE(llUpper).sInsertFlag = ""
        tlSEE(llUpper).sABCFormat = rst!seeABCFormat
        tlSEE(llUpper).sABCPgmCode = rst!seeABCPgmCode
        tlSEE(llUpper).sABCXDSMode = rst!seeABCXDSMode
        tlSEE(llUpper).sABCRecordItem = rst!seeABCRecordItem
        tlSEE(llUpper).sUnused = ""
        'Field not part of record
        tlSEE(llUpper).lAvailLength = tlSEE(llUpper).lDuration
        llUpper = llUpper + 1
        ReDim Preserve tlSEE(0 To llUpper) As SEE
        rst.MoveNext
    Wend
    slSEEStamp = slStamp
    rst.Close
    gGetRecs_SEE_ScheduleEventsByDHEandSHE = True
    Exit Function
    
gGetRecs_SEE_ScheduleEventsByDHEErr:
    ilRet = 1
    Resume Next
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetRecs_SEE_ScheduleEventsByDHEandSHE = False
    Exit Function

End Function


Public Sub gGetEPEManSummary()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slInitSetting As String
    Dim slCategory As String
    Dim ilETE As Integer
    
    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrEventType-mPopulate", tgCurrEPE())
    If UBound(tgCurrEPE) <= LBound(tgCurrEPE) Then
        slInitSetting = "N"
    Else
        slInitSetting = "N"
    End If
    tgManSumEPE.sBus = slInitSetting
    tgManSumEPE.sBusControl = slInitSetting
    tgManSumEPE.sTime = slInitSetting
    tgManSumEPE.sStartType = slInitSetting
    tgManSumEPE.sFixedTime = slInitSetting
    tgManSumEPE.sEndType = slInitSetting
    tgManSumEPE.sDuration = slInitSetting
    tgManSumEPE.sMaterialType = slInitSetting
    tgManSumEPE.sAudioName = slInitSetting
    tgManSumEPE.sAudioItemID = slInitSetting
    tgManSumEPE.sAudioISCI = slInitSetting
    tgManSumEPE.sAudioControl = slInitSetting
    tgManSumEPE.sBkupAudioName = slInitSetting
    tgManSumEPE.sBkupAudioControl = slInitSetting
    tgManSumEPE.sProtAudioName = slInitSetting
    tgManSumEPE.sProtAudioItemID = slInitSetting
    tgManSumEPE.sProtAudioISCI = slInitSetting
    tgManSumEPE.sProtAudioControl = slInitSetting
    tgManSumEPE.sRelay1 = slInitSetting
    tgManSumEPE.sRelay2 = slInitSetting
    tgManSumEPE.sFollow = slInitSetting
    tgManSumEPE.sSilenceTime = slInitSetting
    tgManSumEPE.sSilence1 = slInitSetting
    tgManSumEPE.sSilence2 = slInitSetting
    tgManSumEPE.sSilence3 = slInitSetting
    tgManSumEPE.sSilence4 = slInitSetting
    tgManSumEPE.sStartNetcue = slInitSetting
    tgManSumEPE.sStopNetcue = slInitSetting
    tgManSumEPE.sTitle1 = slInitSetting
    tgManSumEPE.sTitle2 = slInitSetting
    tgManSumEPE.sABCFormat = slInitSetting
    tgManSumEPE.sABCPgmCode = slInitSetting
    tgManSumEPE.sABCXDSMode = slInitSetting
    tgManSumEPE.sABCRecordItem = slInitSetting
    LSet tgSchManSumEPE = tgManSumEPE
    For ilLoop = 0 To UBound(tgCurrEPE) - 1 Step 1
        slCategory = ""
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tgCurrETE(ilETE).iCode = tgCurrEPE(ilLoop).iEteCode Then
                slCategory = tgCurrETE(ilETE).sCategory
                Exit For
            End If
        Next ilETE
        If tgCurrEPE(ilLoop).sType = "M" Then
            If tgCurrEPE(ilLoop).sBus = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sBus = "Y"
                End If
                tgSchManSumEPE.sBus = "Y"
            End If
            If tgCurrEPE(ilLoop).sBusControl = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sBusControl = "Y"
                End If
                tgSchManSumEPE.sBusControl = "Y"
            End If
            If tgCurrEPE(ilLoop).sTime = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sTime = "Y"
                End If
                tgSchManSumEPE.sTime = "Y"
            End If
            If tgCurrEPE(ilLoop).sStartType = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sStartType = "Y"
                End If
                tgSchManSumEPE.sStartType = "Y"
            End If
            If tgCurrEPE(ilLoop).sFixedTime = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sFixedTime = "Y"
                End If
                tgSchManSumEPE.sFixedTime = "Y"
            End If
            If tgCurrEPE(ilLoop).sEndType = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sEndType = "Y"
                End If
                tgSchManSumEPE.sEndType = "Y"
            End If
            If tgCurrEPE(ilLoop).sDuration = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sDuration = "Y"
                End If
                tgSchManSumEPE.sDuration = "Y"
            End If
            If tgCurrEPE(ilLoop).sMaterialType = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sMaterialType = "Y"
                End If
                tgSchManSumEPE.sMaterialType = "Y"
            End If
            If tgCurrEPE(ilLoop).sAudioName = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sAudioName = "Y"
                End If
                tgSchManSumEPE.sAudioName = "Y"
            End If
            If tgCurrEPE(ilLoop).sAudioItemID = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sAudioItemID = "Y"
                End If
                tgSchManSumEPE.sAudioItemID = "Y"
            End If
            If tgCurrEPE(ilLoop).sAudioISCI = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sAudioISCI = "Y"
                End If
                tgSchManSumEPE.sAudioISCI = "Y"
            End If
            If tgCurrEPE(ilLoop).sAudioControl = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sAudioControl = "Y"
                End If
                tgSchManSumEPE.sAudioControl = "Y"
            End If
            If tgCurrEPE(ilLoop).sBkupAudioName = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sBkupAudioName = "Y"
                End If
                tgSchManSumEPE.sBkupAudioName = "Y"
            End If
            If tgCurrEPE(ilLoop).sBkupAudioControl = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sBkupAudioControl = "Y"
                End If
                tgSchManSumEPE.sBkupAudioControl = "Y"
            End If
            If tgCurrEPE(ilLoop).sProtAudioName = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sProtAudioName = "Y"
                End If
                tgSchManSumEPE.sProtAudioName = "Y"
            End If
            If tgCurrEPE(ilLoop).sProtAudioItemID = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sProtAudioItemID = "Y"
                End If
                tgSchManSumEPE.sProtAudioItemID = "Y"
            End If
            If tgCurrEPE(ilLoop).sProtAudioISCI = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sProtAudioISCI = "Y"
                End If
                tgSchManSumEPE.sProtAudioISCI = "Y"
            End If
            If tgCurrEPE(ilLoop).sProtAudioControl = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sProtAudioControl = "Y"
                End If
                tgSchManSumEPE.sProtAudioControl = "Y"
            End If
            If tgCurrEPE(ilLoop).sRelay1 = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sRelay1 = "Y"
                End If
                tgSchManSumEPE.sRelay1 = "Y"
            End If
            If tgCurrEPE(ilLoop).sRelay2 = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sRelay2 = "Y"
                End If
                tgSchManSumEPE.sRelay2 = "Y"
            End If
            If tgCurrEPE(ilLoop).sFollow = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sFollow = "Y"
                End If
                tgSchManSumEPE.sFollow = "Y"
            End If
            If tgCurrEPE(ilLoop).sSilenceTime = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sSilenceTime = "Y"
                End If
                tgSchManSumEPE.sSilenceTime = "Y"
            End If
            If tgCurrEPE(ilLoop).sSilence1 = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sSilence1 = "Y"
                End If
                tgSchManSumEPE.sSilence1 = "Y"
            End If
            If tgCurrEPE(ilLoop).sSilence2 = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sSilence2 = "Y"
                End If
                tgSchManSumEPE.sSilence2 = "Y"
            End If
            If tgCurrEPE(ilLoop).sSilence3 = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sSilence3 = "Y"
                End If
                tgSchManSumEPE.sSilence3 = "Y"
            End If
            If tgCurrEPE(ilLoop).sSilence4 = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sSilence4 = "Y"
                End If
                tgSchManSumEPE.sSilence4 = "Y"
            End If
            If tgCurrEPE(ilLoop).sStartNetcue = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sStartNetcue = "Y"
                End If
                tgSchManSumEPE.sStartNetcue = "Y"
            End If
            If tgCurrEPE(ilLoop).sStopNetcue = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sStopNetcue = "Y"
                End If
                tgSchManSumEPE.sStopNetcue = "Y"
            End If
            If tgCurrEPE(ilLoop).sTitle1 = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sTitle1 = "Y"
                End If
                tgSchManSumEPE.sTitle1 = "Y"
            End If
            If tgCurrEPE(ilLoop).sTitle2 = "Y" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sTitle2 = "Y"
                End If
                tgSchManSumEPE.sTitle2 = "Y"
            End If
            If sgClientFields = "A" Then
                If slCategory <> "S" Then
                    tgManSumEPE.sABCFormat = "Y"
                End If
                tgSchManSumEPE.sABCFormat = "Y"
                If slCategory <> "S" Then
                    tgManSumEPE.sABCPgmCode = "Y"
                End If
                tgSchManSumEPE.sABCPgmCode = "Y"
                If slCategory <> "S" Then
                    tgManSumEPE.sABCXDSMode = "Y"
                End If
                tgSchManSumEPE.sABCXDSMode = "Y"
                If slCategory <> "S" Then
                    tgManSumEPE.sABCRecordItem = "Y"
                End If
                tgSchManSumEPE.sABCRecordItem = "Y"
            End If
        End If
    Next ilLoop
    If tgCurrEPE(ilLoop).sBus = "" Then
        tgManSumEPE.sBus = "N"
    End If
    If tgCurrEPE(ilLoop).sBusControl = "" Then
        tgManSumEPE.sBusControl = "N"
    End If
    If tgCurrEPE(ilLoop).sTime = "" Then
        tgManSumEPE.sTime = "N"
    End If
    If tgCurrEPE(ilLoop).sStartType = "" Then
        tgManSumEPE.sStartType = "N"
    End If
    If tgCurrEPE(ilLoop).sFixedTime = "" Then
        tgManSumEPE.sFixedTime = "N"
    End If
    If tgCurrEPE(ilLoop).sEndType = "" Then
        tgManSumEPE.sEndType = "N"
    End If
    If tgCurrEPE(ilLoop).sDuration = "" Then
        tgManSumEPE.sDuration = "N"
    End If
    If tgCurrEPE(ilLoop).sMaterialType = "" Then
        tgManSumEPE.sMaterialType = "N"
    End If
    If tgCurrEPE(ilLoop).sAudioName = "" Then
        tgManSumEPE.sAudioName = "N"
    End If
    If tgCurrEPE(ilLoop).sAudioItemID = "" Then
        tgManSumEPE.sAudioItemID = "N"
    End If
    If tgCurrEPE(ilLoop).sAudioISCI = "" Then
        tgManSumEPE.sAudioISCI = "N"
    End If
    If tgCurrEPE(ilLoop).sAudioControl = "" Then
        tgManSumEPE.sAudioControl = "N"
    End If
    If tgCurrEPE(ilLoop).sBkupAudioName = "" Then
        tgManSumEPE.sBkupAudioName = "N"
    End If
    If tgCurrEPE(ilLoop).sBkupAudioControl = "" Then
        tgManSumEPE.sBkupAudioControl = "N"
    End If
    If tgCurrEPE(ilLoop).sProtAudioName = "" Then
        tgManSumEPE.sProtAudioName = "N"
    End If
    If tgCurrEPE(ilLoop).sProtAudioItemID = "" Then
        tgManSumEPE.sProtAudioItemID = "N"
    End If
    If tgCurrEPE(ilLoop).sProtAudioISCI = "" Then
        tgManSumEPE.sProtAudioISCI = "N"
    End If
    If tgCurrEPE(ilLoop).sProtAudioControl = "" Then
        tgManSumEPE.sProtAudioControl = "N"
    End If
    If tgCurrEPE(ilLoop).sRelay1 = "" Then
        tgManSumEPE.sRelay1 = "N"
    End If
    If tgCurrEPE(ilLoop).sRelay2 = "" Then
        tgManSumEPE.sRelay2 = "N"
    End If
    If tgCurrEPE(ilLoop).sFollow = "" Then
        tgManSumEPE.sFollow = "N"
    End If
    If tgCurrEPE(ilLoop).sSilenceTime = "" Then
        tgManSumEPE.sSilenceTime = "N"
    End If
    If tgCurrEPE(ilLoop).sSilence1 = "" Then
        tgManSumEPE.sSilence1 = "N"
    End If
    If tgCurrEPE(ilLoop).sSilence2 = "" Then
        tgManSumEPE.sSilence2 = "N"
    End If
    If tgCurrEPE(ilLoop).sSilence3 = "" Then
        tgManSumEPE.sSilence3 = "N"
    End If
    If tgCurrEPE(ilLoop).sSilence4 = "" Then
        tgManSumEPE.sSilence4 = "N"
    End If
    If tgCurrEPE(ilLoop).sStartNetcue = "" Then
        tgManSumEPE.sStartNetcue = "N"
    End If
    If tgCurrEPE(ilLoop).sStopNetcue = "" Then
        tgManSumEPE.sStopNetcue = "N"
    End If
    If tgCurrEPE(ilLoop).sTitle1 = "" Then
        tgManSumEPE.sTitle1 = "N"
    End If
    If tgCurrEPE(ilLoop).sTitle2 = "" Then
        tgManSumEPE.sTitle2 = "N"
    End If
    If sgClientFields = "A" Then
        If tgCurrEPE(ilLoop).sABCFormat = "" Then
            tgManSumEPE.sABCFormat = "N"
        End If
        If tgCurrEPE(ilLoop).sABCPgmCode = "" Then
            tgManSumEPE.sABCPgmCode = "N"
        End If
        If tgCurrEPE(ilLoop).sABCXDSMode = "" Then
            tgManSumEPE.sABCXDSMode = "N"
        End If
        If tgCurrEPE(ilLoop).sABCRecordItem = "" Then
            tgManSumEPE.sABCRecordItem = "N"
        End If
    Else
        tgManSumEPE.sABCFormat = "N"
        tgManSumEPE.sABCPgmCode = "N"
        tgManSumEPE.sABCXDSMode = "N"
        tgManSumEPE.sABCRecordItem = "N"
    End If
End Sub
Public Sub gGetEPEUsedSummary()
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slInitSetting As String
    Dim slCategory As String
    Dim ilETE As Integer
    
    ilRet = gGetTypeOfRecs_EPE_EventProperties("C", sgCurrEPEStamp, "EngrEventType-mPopulate", tgCurrEPE())
    If UBound(tgCurrEPE) <= LBound(tgCurrEPE) Then
        slInitSetting = "Y"
    Else
        slInitSetting = "N"
    End If
    tgUsedSumEPE.sBus = slInitSetting
    tgUsedSumEPE.sBusControl = slInitSetting
    tgUsedSumEPE.sTime = slInitSetting
    tgUsedSumEPE.sStartType = slInitSetting
    tgUsedSumEPE.sFixedTime = slInitSetting
    tgUsedSumEPE.sEndType = slInitSetting
    tgUsedSumEPE.sDuration = slInitSetting
    tgUsedSumEPE.sMaterialType = slInitSetting
    tgUsedSumEPE.sAudioName = slInitSetting
    tgUsedSumEPE.sAudioItemID = slInitSetting
    tgUsedSumEPE.sAudioISCI = slInitSetting
    tgUsedSumEPE.sAudioControl = slInitSetting
    tgUsedSumEPE.sBkupAudioName = slInitSetting
    tgUsedSumEPE.sBkupAudioControl = slInitSetting
    tgUsedSumEPE.sProtAudioName = slInitSetting
    tgUsedSumEPE.sProtAudioItemID = slInitSetting
    tgUsedSumEPE.sProtAudioISCI = slInitSetting
    tgUsedSumEPE.sProtAudioControl = slInitSetting
    tgUsedSumEPE.sRelay1 = slInitSetting
    tgUsedSumEPE.sRelay2 = slInitSetting
    tgUsedSumEPE.sFollow = slInitSetting
    tgUsedSumEPE.sSilenceTime = slInitSetting
    tgUsedSumEPE.sSilence1 = slInitSetting
    tgUsedSumEPE.sSilence2 = slInitSetting
    tgUsedSumEPE.sSilence3 = slInitSetting
    tgUsedSumEPE.sSilence4 = slInitSetting
    tgUsedSumEPE.sStartNetcue = slInitSetting
    tgUsedSumEPE.sStopNetcue = slInitSetting
    tgUsedSumEPE.sTitle1 = slInitSetting
    tgUsedSumEPE.sTitle2 = slInitSetting
    tgUsedSumEPE.sABCFormat = slInitSetting
    tgUsedSumEPE.sABCPgmCode = slInitSetting
    tgUsedSumEPE.sABCXDSMode = slInitSetting
    tgUsedSumEPE.sABCRecordItem = slInitSetting
    LSet tgSchUsedSumEPE = tgUsedSumEPE
    For ilLoop = 0 To UBound(tgCurrEPE) - 1 Step 1
        slCategory = ""
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tgCurrETE(ilETE).iCode = tgCurrEPE(ilLoop).iEteCode Then
                slCategory = tgCurrETE(ilETE).sCategory
                Exit For
            End If
        Next ilETE
        If tgCurrEPE(ilLoop).sType = "U" Then
            If tgCurrEPE(ilLoop).sBus = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sBus = "Y"
                End If
                tgSchUsedSumEPE.sBus = "Y"
            End If
            If tgCurrEPE(ilLoop).sBusControl = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sBusControl = "Y"
                End If
                tgSchUsedSumEPE.sBusControl = "Y"
            End If
            If tgCurrEPE(ilLoop).sTime = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sTime = "Y"
                End If
                tgSchUsedSumEPE.sTime = "Y"
            End If
            If tgCurrEPE(ilLoop).sStartType = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sStartType = "Y"
                End If
                tgSchUsedSumEPE.sStartType = "Y"
            End If
            If tgCurrEPE(ilLoop).sFixedTime = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sFixedTime = "Y"
                End If
                tgSchUsedSumEPE.sFixedTime = "Y"
            End If
            If tgCurrEPE(ilLoop).sEndType = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sEndType = "Y"
                End If
                tgSchUsedSumEPE.sEndType = "Y"
            End If
            If tgCurrEPE(ilLoop).sDuration = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sDuration = "Y"
                End If
                tgSchUsedSumEPE.sDuration = "Y"
            End If
            If tgCurrEPE(ilLoop).sMaterialType = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sMaterialType = "Y"
                End If
                tgSchUsedSumEPE.sMaterialType = "Y"
            End If
            If tgCurrEPE(ilLoop).sAudioName = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sAudioName = "Y"
                End If
                tgSchUsedSumEPE.sAudioName = "Y"
            End If
            If tgCurrEPE(ilLoop).sAudioItemID = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sAudioItemID = "Y"
                End If
                tgSchUsedSumEPE.sAudioItemID = "Y"
            End If
            If tgCurrEPE(ilLoop).sAudioISCI = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sAudioISCI = "Y"
                End If
                tgSchUsedSumEPE.sAudioISCI = "Y"
            End If
            If tgCurrEPE(ilLoop).sAudioControl = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sAudioControl = "Y"
                End If
                tgSchUsedSumEPE.sAudioControl = "Y"
            End If
            If tgCurrEPE(ilLoop).sBkupAudioName = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sBkupAudioName = "Y"
                End If
                tgSchUsedSumEPE.sBkupAudioName = "Y"
            End If
            If tgCurrEPE(ilLoop).sBkupAudioControl = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sBkupAudioControl = "Y"
                End If
                tgSchUsedSumEPE.sBkupAudioControl = "Y"
            End If
            If tgCurrEPE(ilLoop).sProtAudioName = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sProtAudioName = "Y"
                End If
                tgSchUsedSumEPE.sProtAudioName = "Y"
            End If
            If tgCurrEPE(ilLoop).sProtAudioItemID = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sProtAudioItemID = "Y"
                End If
                tgSchUsedSumEPE.sProtAudioItemID = "Y"
            End If
            If tgCurrEPE(ilLoop).sProtAudioISCI = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sProtAudioISCI = "Y"
                End If
                tgSchUsedSumEPE.sProtAudioISCI = "Y"
            End If
            If tgCurrEPE(ilLoop).sProtAudioControl = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sProtAudioControl = "Y"
                End If
                tgSchUsedSumEPE.sProtAudioControl = "Y"
            End If
            If tgCurrEPE(ilLoop).sRelay1 = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sRelay1 = "Y"
                End If
                tgSchUsedSumEPE.sRelay1 = "Y"
            End If
            If tgCurrEPE(ilLoop).sRelay2 = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sRelay2 = "Y"
                End If
                tgSchUsedSumEPE.sRelay2 = "Y"
            End If
            If tgCurrEPE(ilLoop).sFollow = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sFollow = "Y"
                End If
                tgSchUsedSumEPE.sFollow = "Y"
            End If
            If tgCurrEPE(ilLoop).sSilenceTime = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sSilenceTime = "Y"
                End If
                tgSchUsedSumEPE.sSilenceTime = "Y"
            End If
            If tgCurrEPE(ilLoop).sSilence1 = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sSilence1 = "Y"
                End If
                tgSchUsedSumEPE.sSilence1 = "Y"
            End If
            If tgCurrEPE(ilLoop).sSilence2 = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sSilence2 = "Y"
                End If
                tgSchUsedSumEPE.sSilence2 = "Y"
            End If
            If tgCurrEPE(ilLoop).sSilence3 = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sSilence3 = "Y"
                End If
                tgSchUsedSumEPE.sSilence3 = "Y"
            End If
            If tgCurrEPE(ilLoop).sSilence4 = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sSilence4 = "Y"
                End If
                tgSchUsedSumEPE.sSilence4 = "Y"
            End If
            If tgCurrEPE(ilLoop).sStartNetcue = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sStartNetcue = "Y"
                End If
                tgSchUsedSumEPE.sStartNetcue = "Y"
            End If
            If tgCurrEPE(ilLoop).sStopNetcue = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sStopNetcue = "Y"
                End If
                tgSchUsedSumEPE.sStopNetcue = "Y"
            End If
            If tgCurrEPE(ilLoop).sTitle1 = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sTitle1 = "Y"
                End If
                tgSchUsedSumEPE.sTitle1 = "Y"
            End If
            If tgCurrEPE(ilLoop).sTitle2 = "Y" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sTitle2 = "Y"
                End If
                tgSchUsedSumEPE.sTitle2 = "Y"
            End If
            If sgClientFields = "A" Then
                If slCategory <> "S" Then
                    tgUsedSumEPE.sABCFormat = "Y"
                End If
                tgSchUsedSumEPE.sABCFormat = "Y"
                If slCategory <> "S" Then
                    tgUsedSumEPE.sABCPgmCode = "Y"
                End If
                tgSchUsedSumEPE.sABCPgmCode = "Y"
                If slCategory <> "S" Then
                    tgUsedSumEPE.sABCXDSMode = "Y"
                End If
                tgSchUsedSumEPE.sABCXDSMode = "Y"
                If slCategory <> "S" Then
                    tgUsedSumEPE.sABCRecordItem = "Y"
                End If
                tgSchUsedSumEPE.sABCRecordItem = "Y"
            End If
        End If
    Next ilLoop
    If tgCurrEPE(ilLoop).sBus = "" Then
        tgUsedSumEPE.sBus = "N"
    End If
    If tgCurrEPE(ilLoop).sBusControl = "" Then
        tgUsedSumEPE.sBusControl = "N"
    End If
    If tgCurrEPE(ilLoop).sTime = "" Then
        tgUsedSumEPE.sTime = "N"
    End If
    If tgCurrEPE(ilLoop).sStartType = "" Then
        tgUsedSumEPE.sStartType = "N"
    End If
    If tgCurrEPE(ilLoop).sFixedTime = "" Then
        tgUsedSumEPE.sFixedTime = "N"
    End If
    If tgCurrEPE(ilLoop).sEndType = "" Then
        tgUsedSumEPE.sEndType = "N"
    End If
    If tgCurrEPE(ilLoop).sDuration = "" Then
        tgUsedSumEPE.sDuration = "N"
    End If
    If tgCurrEPE(ilLoop).sMaterialType = "" Then
        tgUsedSumEPE.sMaterialType = "N"
    End If
    If tgCurrEPE(ilLoop).sAudioName = "" Then
        tgUsedSumEPE.sAudioName = "N"
    End If
    If tgCurrEPE(ilLoop).sAudioItemID = "" Then
        tgUsedSumEPE.sAudioItemID = "N"
    End If
    If tgCurrEPE(ilLoop).sAudioISCI = "" Then
        tgUsedSumEPE.sAudioISCI = "N"
    End If
    If tgCurrEPE(ilLoop).sAudioControl = "" Then
        tgUsedSumEPE.sAudioControl = "N"
    End If
    If tgCurrEPE(ilLoop).sBkupAudioName = "" Then
        tgUsedSumEPE.sBkupAudioName = "N"
    End If
    If tgCurrEPE(ilLoop).sBkupAudioControl = "" Then
        tgUsedSumEPE.sBkupAudioControl = "N"
    End If
    If tgCurrEPE(ilLoop).sProtAudioName = "" Then
        tgUsedSumEPE.sProtAudioName = "N"
    End If
    If tgCurrEPE(ilLoop).sProtAudioItemID = "" Then
        tgUsedSumEPE.sProtAudioItemID = "N"
    End If
    If tgCurrEPE(ilLoop).sProtAudioISCI = "" Then
        tgUsedSumEPE.sProtAudioISCI = "N"
    End If
    If tgCurrEPE(ilLoop).sProtAudioControl = "" Then
        tgUsedSumEPE.sProtAudioControl = "N"
    End If
    If tgCurrEPE(ilLoop).sRelay1 = "" Then
        tgUsedSumEPE.sRelay1 = "N"
    End If
    If tgCurrEPE(ilLoop).sRelay2 = "" Then
        tgUsedSumEPE.sRelay2 = "N"
    End If
    If tgCurrEPE(ilLoop).sFollow = "" Then
        tgUsedSumEPE.sFollow = "N"
    End If
    If tgCurrEPE(ilLoop).sSilenceTime = "" Then
        tgUsedSumEPE.sSilenceTime = "N"
    End If
    If tgCurrEPE(ilLoop).sSilence1 = "" Then
        tgUsedSumEPE.sSilence1 = "N"
    End If
    If tgCurrEPE(ilLoop).sSilence2 = "" Then
        tgUsedSumEPE.sSilence2 = "N"
    End If
    If tgCurrEPE(ilLoop).sSilence3 = "" Then
        tgUsedSumEPE.sSilence3 = "N"
    End If
    If tgCurrEPE(ilLoop).sSilence4 = "" Then
        tgUsedSumEPE.sSilence4 = "N"
    End If
    If tgCurrEPE(ilLoop).sStartNetcue = "" Then
        tgUsedSumEPE.sStartNetcue = "N"
    End If
    If tgCurrEPE(ilLoop).sStopNetcue = "" Then
        tgUsedSumEPE.sStopNetcue = "N"
    End If
    If tgCurrEPE(ilLoop).sTitle1 = "" Then
        tgUsedSumEPE.sTitle1 = "N"
    End If
    If tgCurrEPE(ilLoop).sTitle2 = "" Then
        tgUsedSumEPE.sTitle2 = "N"
    End If
    If sgClientFields = "A" Then
        If tgCurrEPE(ilLoop).sABCFormat = "" Then
            tgUsedSumEPE.sABCFormat = "N"
        End If
        If tgCurrEPE(ilLoop).sABCPgmCode = "" Then
            tgUsedSumEPE.sABCPgmCode = "N"
        End If
        If tgCurrEPE(ilLoop).sABCXDSMode = "" Then
            tgUsedSumEPE.sABCXDSMode = "N"
        End If
        If tgCurrEPE(ilLoop).sABCRecordItem = "" Then
            tgUsedSumEPE.sABCRecordItem = "N"
        End If
    Else
        tgUsedSumEPE.sABCFormat = "N"
        tgUsedSumEPE.sABCPgmCode = "N"
        tgUsedSumEPE.sABCXDSMode = "N"
        tgUsedSumEPE.sABCRecordItem = "N"
    End If
End Sub

Public Sub gInitMaxAFE()
    tgNoCharAFE.iBus = 8
    tgNoCharAFE.iBusControl = 1
    tgNoCharAFE.iEventType = 1
    tgNoCharAFE.iTime = 10
    tgNoCharAFE.iStartType = 3
    tgNoCharAFE.iFixedTime = 1
    tgNoCharAFE.iEndType = 3
    tgNoCharAFE.iDuration = 10
    tgNoCharAFE.iEndTime = 10
    tgNoCharAFE.iMaterialType = 3
    tgNoCharAFE.iAudioName = 8
    tgNoCharAFE.iAudioItemID = 32
    tgNoCharAFE.iAudioISCI = 20
    tgNoCharAFE.iAudioControl = 1
    tgNoCharAFE.iBkupAudioName = 8
    tgNoCharAFE.iBkupAudioControl = 1
    tgNoCharAFE.iProtAudioName = 8
    tgNoCharAFE.iProtItemID = 32
    tgNoCharAFE.iProtISCI = 20
    tgNoCharAFE.iProtAudioControl = 1
    tgNoCharAFE.iRelay1 = 8
    tgNoCharAFE.iRelay2 = 8
    tgNoCharAFE.iFollow = 19
    tgNoCharAFE.iSilenceTime = 5
    tgNoCharAFE.iSilence1 = 1
    tgNoCharAFE.iSilence2 = 1
    tgNoCharAFE.iSilence3 = 1
    tgNoCharAFE.iSilence4 = 1
    tgNoCharAFE.iStartNetcue = 3
    tgNoCharAFE.iStopNetcue = 3
    tgNoCharAFE.iTitle1 = 66
    tgNoCharAFE.iTitle2 = 90
    tgNoCharAFE.iEventID = 20
    tgNoCharAFE.iDate = 8

End Sub

Public Function gGetLatestVersion_DHE(llOrigDHECode As Long, slForm_Module As String) As Integer
    Dim rst As ADODB.Recordset
    
    gGetLatestVersion_DHE = -1
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT Max(dheVersion) FROM DHE_Day_Header_Info WHERE dheOrigDheCode = " & llOrigDHECode
    Set rst = cnn.Execute(sgSQLQuery)
    If IsNull(rst(0).Value) Then
    Else
        If rst(0).Value >= 0 Then
            gGetLatestVersion_DHE = rst(0).Value
        Else
        End If
    End If
    rst.Close
    Exit Function
    
ErrHand:
    gShowErrorMsg slForm_Module
    Exit Function
        
End Function

Public Function gGetItemMsgID(ilIteCode As Integer, slForm_Module As String, ilMsgID As Integer) As Integer
    Dim rst As ADODB.Recordset
    
    gGetItemMsgID = False
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM ITE_Item_Test Where iteCode = " & ilIteCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        ilMsgID = rst!iteCurrMgsID + 1
        If ilMsgID > rst!iteMaxMgsID Then
            ilMsgID = rst!iteMinMgsID
        End If
        sgSQLQuery = "UPDATE ITE_Item_Test SET "
        sgSQLQuery = sgSQLQuery & "iteCurrMgsID = " & ilMsgID
        sgSQLQuery = sgSQLQuery & " WHERE iteCode = " & ilIteCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gGetItemMsgID = True
    Else
        gGetItemMsgID = False
    End If
    rst.Close
    Exit Function
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    Exit Function

End Function

Public Sub gShowErrorMsg(slForm_Module As String)
    Dim slMsg As String
    
    Screen.MousePointer = vbDefault
    slMsg = ""
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            slMsg = "A SQL error has occured in " & slForm_Module & ": "
            If igOperationMode = 1 Then
                gLogMsg slMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, "EngrServiceErrors.Txt", False
            Else
                gLogMsg slMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, "EngrErrors.Txt", False
                MsgBox slMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError, vbCritical
            End If
        End If
    Next gErrSQL
    If (Err.Number <> 0) And (slMsg = "") Then
        slMsg = "A general error has occured in " & slForm_Module & ": "
        If igOperationMode = 1 Then
            gLogMsg slMsg & Err.Description & "; Error #" & Err.Number, "EngrServiceErrors.Txt", False
        Else
            gLogMsg slMsg & Err.Description & "; Error #" & Err.Number, "EngrErrors.Txt", False
            MsgBox slMsg & Err.Description & "; Error #" & Err.Number, vbCritical
        End If
    End If

End Sub

Public Function gGetMergeStatus(slForm_Module As String, slMergeStatus As String) As Integer
    Dim rst As ADODB.Recordset
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT soeMergeStopFlag FROM SOE_Site_Option WHERE soeCurrent = 'Y'"
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        slMergeStatus = rst!soeMergeStopFlag
        gGetMergeStatus = True
    Else
        gGetMergeStatus = False
    End If
    rst.Close
    Exit Function
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetMergeStatus = False
    Exit Function
End Function

Public Function gBinarySearchBDE(ilCode As Integer, tlBDE() As BDE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchBDE = -1
        Exit Function
    End If
    llMin = LBound(tlBDE)
    llMax = UBound(tlBDE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlBDE(llMiddle).iCode Then
            'found the match
            gBinarySearchBDE = llMiddle
            Exit Function
        ElseIf ilCode < tlBDE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchBDE = -1
End Function

Public Function gBinarySearchCCE(ilCode As Integer, tlCCE() As CCE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchCCE = -1
        Exit Function
    End If
    llMin = LBound(tlCCE)
    llMax = UBound(tlCCE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlCCE(llMiddle).iCode Then
            'found the match
            gBinarySearchCCE = llMiddle
            Exit Function
        ElseIf ilCode < tlCCE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchCCE = -1
End Function



Public Function gBinarySearchETE(ilCode As Integer, tlETE() As ETE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchETE = -1
        Exit Function
    End If
    llMin = LBound(tlETE)
    llMax = UBound(tlETE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlETE(llMiddle).iCode Then
            'found the match
            gBinarySearchETE = llMiddle
            Exit Function
        ElseIf ilCode < tlETE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchETE = -1
End Function

Public Function gBinarySearchTTE(ilCode As Integer, tlTTE() As TTE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchTTE = -1
        Exit Function
    End If
    llMin = LBound(tlTTE)
    llMax = UBound(tlTTE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlTTE(llMiddle).iCode Then
            'found the match
            gBinarySearchTTE = llMiddle
            Exit Function
        ElseIf ilCode < tlTTE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchTTE = -1
End Function


Public Function gBinarySearchMTE(ilCode As Integer, tlMTE() As MTE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchMTE = -1
        Exit Function
    End If
    llMin = LBound(tlMTE)
    llMax = UBound(tlMTE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlMTE(llMiddle).iCode Then
            'found the match
            gBinarySearchMTE = llMiddle
            Exit Function
        ElseIf ilCode < tlMTE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchMTE = -1
End Function

Public Function gBinarySearchASE(ilCode As Integer, tlASE() As ASE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchASE = -1
        Exit Function
    End If
    llMin = LBound(tlASE)
    llMax = UBound(tlASE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlASE(llMiddle).iCode Then
            'found the match
            gBinarySearchASE = llMiddle
            Exit Function
        ElseIf ilCode < tlASE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchASE = -1
End Function

Public Function gBinarySearchANE(ilCode As Integer, tlANE() As ANE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchANE = -1
        Exit Function
    End If
    llMin = LBound(tlANE)
    llMax = UBound(tlANE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlANE(llMiddle).iCode Then
            'found the match
            gBinarySearchANE = llMiddle
            Exit Function
        ElseIf ilCode < tlANE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchANE = -1
End Function

Public Function gBinarySearchRNE(ilCode As Integer, tlRNE() As RNE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchRNE = -1
        Exit Function
    End If
    llMin = LBound(tlRNE)
    llMax = UBound(tlRNE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlRNE(llMiddle).iCode Then
            'found the match
            gBinarySearchRNE = llMiddle
            Exit Function
        ElseIf ilCode < tlRNE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchRNE = -1
End Function

Public Function gBinarySearchFNE(ilCode As Integer, tlFNE() As FNE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchFNE = -1
        Exit Function
    End If
    llMin = LBound(tlFNE)
    llMax = UBound(tlFNE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlFNE(llMiddle).iCode Then
            'found the match
            gBinarySearchFNE = llMiddle
            Exit Function
        ElseIf ilCode < tlFNE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchFNE = -1
End Function

Public Function gBinarySearchSCE(ilCode As Integer, tlSCE() As SCE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchSCE = -1
        Exit Function
    End If
    llMin = LBound(tlSCE)
    llMax = UBound(tlSCE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlSCE(llMiddle).iCode Then
            'found the match
            gBinarySearchSCE = llMiddle
            Exit Function
        ElseIf ilCode < tlSCE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchSCE = -1
End Function

Public Function gBinarySearchNNE(ilCode As Integer, tlNNE() As NNE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchNNE = -1
        Exit Function
    End If
    llMin = LBound(tlNNE)
    llMax = UBound(tlNNE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlNNE(llMiddle).iCode Then
            'found the match
            gBinarySearchNNE = llMiddle
            Exit Function
        ElseIf ilCode < tlNNE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchNNE = -1
End Function

Public Function gBinarySearchATE(ilCode As Integer, tlATE() As ATE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchATE = -1
        Exit Function
    End If
    llMin = LBound(tlATE)
    llMax = UBound(tlATE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlATE(llMiddle).iCode Then
            'found the match
            gBinarySearchATE = llMiddle
            Exit Function
        ElseIf ilCode < tlATE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchATE = -1
End Function

Public Function gBinarySearchBGE(ilCode As Integer, tlBGE() As BGE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If ilCode <= 0 Then
        gBinarySearchBGE = -1
        Exit Function
    End If
    llMin = LBound(tlBGE)
    llMax = UBound(tlBGE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilCode = tlBGE(llMiddle).iCode Then
            'found the match
            gBinarySearchBGE = llMiddle
            Exit Function
        ElseIf ilCode < tlBGE(llMiddle).iCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchBGE = -1
End Function

Public Function gBinarySearchCTE(llCode As Long, tlCTE() As CTE) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If llCode <= 0 Then
        gBinarySearchCTE = -1
        Exit Function
    End If
    llMin = LBound(tlCTE)
    llMax = UBound(tlCTE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tlCTE(llMiddle).lCode Then
            'found the match
            gBinarySearchCTE = llMiddle
            Exit Function
        ElseIf llCode < tlCTE(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchCTE = -1
End Function
Public Function gBinarySearchName(slInName As String, tlName() As NAMESORT) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim slName As String
    Dim slKey As String
    
    slName = Trim$(UCase$(slInName))
    If slName = "" Then
        gBinarySearchName = -1
        Exit Function
    End If
    llMin = LBound(tlName)
    llMax = UBound(tlName) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        slKey = Trim$(UCase$(tlName(llMiddle).sKey))
        If StrComp(slName, slKey, vbBinaryCompare) = 0 Then
            'found the match
            gBinarySearchName = llMiddle
            Exit Function
        ElseIf StrComp(slName, slKey, vbBinaryCompare) < 0 Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchName = -1
End Function

Public Function gBinarySearchCTEName(slInName As String, tlName() As CTESORT) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim slName As String
    Dim slKey As String
    
    slName = Trim$(UCase$(slInName))
    If slName = "" Then
        gBinarySearchCTEName = -1
        Exit Function
    End If
    llMin = LBound(tlName)
    llMax = UBound(tlName) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        slKey = Trim$(UCase$(tlName(llMiddle).sKey))
        If StrComp(slName, slKey, vbBinaryCompare) = 0 Then
            'found the match
            gBinarySearchCTEName = llMiddle
            Exit Function
        ElseIf StrComp(slName, slKey, vbBinaryCompare) < 0 Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchCTEName = -1
End Function

Public Function gBinarySearchDNE(llCode As Long, tlDNE() As DNE) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If llCode <= 0 Then
        gBinarySearchDNE = -1
        Exit Function
    End If
    llMin = LBound(tlDNE)
    llMax = UBound(tlDNE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tlDNE(llMiddle).lCode Then
            'found the match
            gBinarySearchDNE = llMiddle
            Exit Function
        ElseIf llCode < tlDNE(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchDNE = -1
End Function

Public Function gBinarySearchDSE(llCode As Long, tlDSE() As DSE) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If llCode <= 0 Then
        gBinarySearchDSE = -1
        Exit Function
    End If
    llMin = LBound(tlDSE)
    llMax = UBound(tlDSE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tlDSE(llMiddle).lCode Then
            'found the match
            gBinarySearchDSE = llMiddle
            Exit Function
        ElseIf llCode < tlDSE(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchDSE = -1
End Function



Public Function gGetConflicts_CEE_CME(llGenDate As Long, llGenTime As Long, slType As String, slForm_Module As String, tlConflictResults() As CONFLICTRESULTS) As Integer
    Dim llUpper As Long
    Dim ilPass As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    ReDim tlConflictResults(0 To 0) As CONFLICTRESULTS
    For ilPass = 0 To 1 Step 1
        sgSQLQuery = "SELECT * FROM CEE_Conflict_Events, CME_Conflict_Master Where ceeGenDate = " & llGenDate & " AND ceeGenTime = " & llGenTime
        If ilPass = 0 Then
            sgSQLQuery = sgSQLQuery & " AND ceeEvtType = 'B'"
            sgSQLQuery = sgSQLQuery & " AND cmeBDECode = ceeBDECode"
            sgSQLQuery = sgSQLQuery & " AND cmeEvtType = 'B'"
        Else
            sgSQLQuery = sgSQLQuery & " AND ceeEvtType = 'A'"
            sgSQLQuery = sgSQLQuery & " AND cmeANECode = ceeANECode"
            sgSQLQuery = sgSQLQuery & " AND cmeEvtType = 'A'"
        End If
        sgSQLQuery = sgSQLQuery & " AND cmeday = ceeDay"
        sgSQLQuery = sgSQLQuery & " AND cmeEndTime > ceeStartTime"      'End and start times can overlap
        sgSQLQuery = sgSQLQuery & " AND cmeStartTime < ceeEndTime"
        sgSQLQuery = sgSQLQuery & " AND cmeEndDate >= ceeStartDate"
        sgSQLQuery = sgSQLQuery & " AND cmeStartDate <= ceeEndDate"
        If slType = "S" Then
            sgSQLQuery = sgSQLQuery & " AND cmeXMidNight = 'Y'"
        End If
        Set rst = cnn.Execute(sgSQLQuery)
        llUpper = UBound(tlConflictResults)
        While Not rst.EOF
            tlConflictResults(llUpper).tCEE.lGridEventRow = rst!ceeGridEventRow
            tlConflictResults(llUpper).tCEE.iGridEventCol = rst!ceeGridEventCol
            tlConflictResults(llUpper).tCEE.iBdeCode = rst!ceeBDECode
            tlConflictResults(llUpper).tCEE.iANECode = rst!ceeANECode
            tlConflictResults(llUpper).tCEE.lStartTime = rst!ceeStartTime
            tlConflictResults(llUpper).tCEE.lEndTime = rst!ceeEndTime
            tlConflictResults(llUpper).tCME.lCode = rst!cmeCode
            tlConflictResults(llUpper).tCME.sSource = rst!cmeSource
            tlConflictResults(llUpper).tCME.lSHEDHECode = rst!cmeSHEDHECode
            tlConflictResults(llUpper).tCME.lDseCode = rst!cmeDSECode
            tlConflictResults(llUpper).tCME.lDeeCode = rst!cmeDEECode
            tlConflictResults(llUpper).tCME.lSeeCode = rst!cmeSEECode
            tlConflictResults(llUpper).tCME.sEvtType = rst!cmeEvtType
            tlConflictResults(llUpper).tCME.iBdeCode = rst!cmeBDECode
            tlConflictResults(llUpper).tCME.iANECode = rst!cmeANECode
            tlConflictResults(llUpper).tCME.lStartDate = rst!cmeStartDate
            tlConflictResults(llUpper).tCME.lEndDate = rst!cmeEndDate
            tlConflictResults(llUpper).tCME.sDay = rst!cmeDay
            tlConflictResults(llUpper).tCME.lStartTime = rst!cmeStartTime
            tlConflictResults(llUpper).tCME.lEndTime = rst!cmeEndTime
            tlConflictResults(llUpper).tCME.sItemID = rst!cmeItemID
            tlConflictResults(llUpper).tCME.sXMidNight = rst!cmeXMidNight
            tlConflictResults(llUpper).tCME.sUnused = rst!cmeUnused
            llUpper = llUpper + 1
            ReDim Preserve tlConflictResults(0 To llUpper) As CONFLICTRESULTS
            rst.MoveNext
        Wend
    Next ilPass
    rst.Close
    gGetConflicts_CEE_CME = True
    Exit Function
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetConflicts_CEE_CME = False
    Exit Function
End Function

Public Function gGetConflicts_CME(tlCEE As CEE, slType As String, slForm_Module As String, tlConflictResults() As CONFLICTRESULTS) As Integer
    Dim llUpper As Long
    Dim ilPass As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "SELECT * FROM CME_Conflict_Master "
    If tlCEE.sEvtType = "B" Then
        sgSQLQuery = sgSQLQuery & " WHERE cmeBDECode = " & tlCEE.iBdeCode
    Else
        sgSQLQuery = sgSQLQuery & " WHERE cmeANECode = " & tlCEE.iANECode
    End If
    sgSQLQuery = sgSQLQuery & " AND cmeEvtType = '" & tlCEE.sEvtType & "'"
    sgSQLQuery = sgSQLQuery & " AND cmeday = '" & tlCEE.sDay & "'"
    sgSQLQuery = sgSQLQuery & " AND cmeEndTime > " & tlCEE.lStartTime      'End and start times can overlap
    sgSQLQuery = sgSQLQuery & " AND cmeStartTime < " & tlCEE.lEndTime
    sgSQLQuery = sgSQLQuery & " AND cmeEndDate >= " & tlCEE.lStartDate
    sgSQLQuery = sgSQLQuery & " AND cmeStartDate <= " & tlCEE.lEndDate
    If slType = "S" Then
        sgSQLQuery = sgSQLQuery & " AND cmeXMidNight = 'Y'"
    End If
    Set rst = cnn.Execute(sgSQLQuery)
    llUpper = UBound(tlConflictResults)
    While Not rst.EOF
        tlConflictResults(llUpper).tCEE.lGridEventRow = tlCEE.lGridEventRow 'rst!ceeGridEventRow
        tlConflictResults(llUpper).tCEE.iGridEventCol = tlCEE.iGridEventCol 'rst!ceeGridEventCol
        tlConflictResults(llUpper).tCEE.iBdeCode = tlCEE.iBdeCode     'rst!ceeBDECode
        tlConflictResults(llUpper).tCEE.iANECode = tlCEE.iANECode     'rst!ceeANECode
        tlConflictResults(llUpper).tCEE.lStartTime = tlCEE.lStartTime   'rst!ceeStartTime
        tlConflictResults(llUpper).tCEE.lEndTime = tlCEE.lEndTime 'rst!ceeEndTime
        tlConflictResults(llUpper).tCME.lCode = rst!cmeCode
        tlConflictResults(llUpper).tCME.sSource = rst!cmeSource
        tlConflictResults(llUpper).tCME.lSHEDHECode = rst!cmeSHEDHECode
        tlConflictResults(llUpper).tCME.lDseCode = rst!cmeDSECode
        tlConflictResults(llUpper).tCME.lDeeCode = rst!cmeDEECode
        tlConflictResults(llUpper).tCME.lSeeCode = rst!cmeSEECode
        tlConflictResults(llUpper).tCME.sEvtType = rst!cmeEvtType
        tlConflictResults(llUpper).tCME.iBdeCode = rst!cmeBDECode
        tlConflictResults(llUpper).tCME.iANECode = rst!cmeANECode
        tlConflictResults(llUpper).tCME.lStartDate = rst!cmeStartDate
        tlConflictResults(llUpper).tCME.lEndDate = rst!cmeEndDate
        tlConflictResults(llUpper).tCME.sDay = rst!cmeDay
        tlConflictResults(llUpper).tCME.lStartTime = rst!cmeStartTime
        tlConflictResults(llUpper).tCME.lEndTime = rst!cmeEndTime
        tlConflictResults(llUpper).tCME.sItemID = rst!cmeItemID
        tlConflictResults(llUpper).tCME.sXMidNight = rst!cmeXMidNight
        tlConflictResults(llUpper).tCME.sUnused = rst!cmeUnused
        llUpper = llUpper + 1
        ReDim Preserve tlConflictResults(0 To llUpper) As CONFLICTRESULTS
        rst.MoveNext
    Wend
    rst.Close
    gGetConflicts_CME = True
    Exit Function
    
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetConflicts_CME = False
    Exit Function
End Function

Public Function gBinarySearchARE(llCode As Long, tlARE() As ARE) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    If llCode <= 0 Then
        gBinarySearchARE = -1
        Exit Function
    End If
    llMin = LBound(tlARE)
    llMax = UBound(tlARE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tlARE(llMiddle).lCode Then
            'found the match
            gBinarySearchARE = llMiddle
            Exit Function
        ElseIf llCode < tlARE(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchARE = -1
End Function

Public Function gBinarySearchAAEbyEventID(llEventID As Long, tlAAE() As AAE) As Integer
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    llMin = LBound(tlAAE)
    llMax = UBound(tlAAE) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llEventID = tlAAE(llMiddle).lEventID Then
            'found the match
            gBinarySearchAAEbyEventID = llMiddle
            Exit Function
        ElseIf llEventID < tlAAE(llMiddle).lEventID Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    gBinarySearchAAEbyEventID = -1
End Function

Public Function gGetCount_AAE_As_Aired(llSheCode As Long, slForm_Module As String) As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    gGetCount_AAE_As_Aired = 0
    sgSQLQuery = "SELECT Count(aaeCode) FROM AAE_As_Aired WHERE aaeSheCode = " & llSheCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        If Not IsNull(rst(0).Value) Then
            If rst(0).Value > 0 Then
                gGetCount_AAE_As_Aired = rst(0).Value
            End If
        End If
    End If
    rst.Close
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetCount_AAE_As_Aired = 0
    Exit Function
End Function

Public Function gGetCount(slSQLCommand As String, slForm_Module As String) As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    gGetCount = 0
    sgSQLQuery = slSQLCommand
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        If Not IsNull(rst(0).Value) Then
            If rst(0).Value > 0 Then
                gGetCount = rst(0).Value
            End If
        End If
    End If
    rst.Close
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetCount = 0
    Exit Function
End Function


Public Function gStartPervasive() As Integer
    On Error GoTo gStartPervasiveErr
    
    Set cnn = New ADODB.Connection
    cnn.Open "DSN=" & sgDSN
    Set rst = New ADODB.Recordset
    
    If igTimeOut >= 0 Then
        cnn.CommandTimeout = igTimeOut
    End If
    
    hgDB = CBtrvMngrInit(0, "", "", sgDBPath, 0, sgDBPath) 'Use 0 as 1 gets a GPF. 1=Initialize Btrieve only if not initialized
    gStartPervasive = True
    Exit Function
gStartPervasiveErr:
    gStartPervasive = False
    Exit Function
End Function

Public Sub gCheckIfDisconnected()
    Dim ilRet As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM UIE_User_Info WHERE uieCode = " & 1
    Set rst = cnn.Execute(sgSQLQuery)
    rst.Close
    Exit Sub
ErrHand:
    For Each gErrSQL In cnn.Errors
        If gErrSQL.NativeError <> 0 Then 'SQLSetConnectAttr vs. SQLSetOpenConnection
            If gErrSQL.NativeError = 9305 Then
                Err.Clear
                rst.Close
                cnn.Close
                Set rst = Nothing
                Set cnn = Nothing
                btrStopAppl
                ilRet = gStartPervasive()
                Exit Sub
            End If
        End If
    Next gErrSQL
End Sub

Public Function gGetCheckStatus() As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT Count(sheCode) FROM SHE_Schedule_Header WHERE sheCurrent = 'Y' AND sheAirDate >= '" & Format$(gNow(), sgSQLDateForm) & "' AND ((sheConflictExist = 'Y') or (sheSpotMergeStatus = 'E') or (sheLoadStatus = 'E'))"
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        gGetCheckStatus = rst(0).Value
    Else
        gGetCheckStatus = 0
    End If
    rst.Close
    Exit Function
ErrHand:
    rst.Close
    gGetCheckStatus = -1
    Exit Function
End Function

Public Function gGetConflictStatusByDate(slDate As String) As String
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT sheConflictExist FROM SHE_Schedule_Header WHERE sheCurrent = 'Y'  AND sheAirDate = '" & Format$(gAdjYear(slDate), sgSQLDateForm) & "'"
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        gGetConflictStatusByDate = rst!sheConflictExist
    Else
        gGetConflictStatusByDate = ""
    End If
    rst.Close
    Exit Function
ErrHand:
    rst.Close
    gGetConflictStatusByDate = ""
    Exit Function
End Function
Public Function gGetLatestSchdDate(blSetNowDateIfBlank As Boolean) As String
    '5/31/11: Disallow changes in the schedule area
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT Max(sheAirDate) FROM SHE_Schedule_Header WHERE sheCurrent = 'Y' "
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        If Not IsNull(rst(0).Value) Then
            gGetLatestSchdDate = Format$(rst(0).Value, sgShowDateForm)
        Else
            If blSetNowDateIfBlank Then
                gGetLatestSchdDate = Format(gNow(), sgShowDateForm)
            Else
                gGetLatestSchdDate = ""
            End If
        End If
    Else
        If blSetNowDateIfBlank Then
            gGetLatestSchdDate = Format(gNow(), sgShowDateForm)
        Else
            gGetLatestSchdDate = ""
        End If
    End If
    rst.Close
    Exit Function
ErrHand:
    rst.Close
    If blSetNowDateIfBlank Then
        gGetLatestSchdDate = Format(gNow(), sgShowDateForm)
    Else
        gGetLatestSchdDate = ""
    End If
    Exit Function
End Function
Public Function gGetEarlestSchdDate(blSetNowDateIfBlank As Boolean) As String
    '5/31/11: Disallow changes in the schedule area
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT Min(sheAirDate) FROM SHE_Schedule_Header WHERE sheCurrent = 'Y' "
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        If Not IsNull(rst(0).Value) Then
            gGetEarlestSchdDate = Format$(rst(0).Value, sgShowDateForm)
        Else
            If blSetNowDateIfBlank Then
                gGetEarlestSchdDate = Format(gNow(), sgShowDateForm)
            Else
                gGetEarlestSchdDate = ""
            End If
        End If
    Else
        If blSetNowDateIfBlank Then
            gGetEarlestSchdDate = Format(gNow(), sgShowDateForm)
        Else
            gGetEarlestSchdDate = ""
        End If
    End If
    On Error Resume Next
    rst.Close
    Exit Function
ErrHand:
    rst.Close
    If blSetNowDateIfBlank Then
        gGetEarlestSchdDate = Format(gNow(), sgShowDateForm)
    Else
        gGetEarlestSchdDate = ""
    End If
    Exit Function
End Function
Public Function gGetLatestLoadDate(blSetNowDateIfBlank As Boolean) As String
    '5/31/11: Disallow changes in the schedule area
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT Max(sheAirDate) FROM SHE_Schedule_Header WHERE sheCurrent = 'Y' AND sheLoadedAutoStatus = 'L'"
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        If Not IsNull(rst(0).Value) Then
            gGetLatestLoadDate = Format$(rst(0).Value, sgShowDateForm)
        Else
            If blSetNowDateIfBlank Then
                gGetLatestLoadDate = Format(gNow(), sgShowDateForm)
            Else
                gGetLatestLoadDate = ""
            End If
        End If
    Else
        If blSetNowDateIfBlank Then
            gGetLatestLoadDate = Format(gNow(), sgShowDateForm)
        Else
            gGetLatestLoadDate = ""
        End If
    End If
    rst.Close
    Exit Function
ErrHand:
    rst.Close
    If blSetNowDateIfBlank Then
        gGetLatestLoadDate = Format(gNow(), sgShowDateForm)
    Else
        gGetLatestLoadDate = ""
    End If
    Exit Function
End Function
Public Function gGetCount_SEE_For_DHE(llDheCode As Long, slForm_Module As String) As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    gGetCount_SEE_For_DHE = 0
    sgSQLQuery = "SELECT Count(seeCode) FROM SEE_Schedule_Events WHERE seeDheCode = " & llDheCode
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        If Not IsNull(rst(0).Value) Then
            If rst(0).Value > 0 Then
                gGetCount_SEE_For_DHE = rst(0).Value
            End If
        End If
    End If
    rst.Close
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetCount_SEE_For_DHE = 0
    Exit Function
End Function

Public Function gGetServiceStatus_MIE_MessageInfo(slForm_Module As String, tlMie As MIE) As Integer
'
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim rst As ADODB.Recordset
    
    
    On Error GoTo ErrHand
    sgSQLQuery = "SELECT * FROM MIE_Message_Info WHERE mieType = 'T' AND mieID = -1"
    Set rst = cnn.Execute(sgSQLQuery)
    If (rst.EOF) Then
        slNowDate = Format(Now, sgShowDateForm)  'Format(gNow(), sgShowDateForm)
        slNowTime = Format(Now, sgShowTimeWSecForm)  'Format(gNow(), sgShowTimeWSecForm)
        tlMie.lCode = 0
        tlMie.sType = "T"
        tlMie.lID = -1
        tlMie.sMessage = "Service Status"
        tlMie.sEnteredDate = Format$(slNowDate, sgShowDateForm)
        tlMie.sEnteredTime = Format$(slNowTime, sgShowTimeWSecForm)
        tlMie.iUieCode = tgUIE.iCode
        tlMie.sUnused = ""
        ilRet = gPutInsert_MIE_MessageInfo(tlMie, slForm_Module)
        If Not ilRet Then
            tlMie.lCode = 0
            gGetServiceStatus_MIE_MessageInfo = False
            Exit Function
        End If
        sgSQLQuery = "SELECT * FROM MIE_Message_Info WHERE mieType = 'T' AND mieID = -1"
        Set rst = cnn.Execute(sgSQLQuery)
    End If
    tlMie.lCode = rst!mieCode
    tlMie.sType = rst!mieType
    tlMie.lID = rst!mieID
    tlMie.sMessage = rst!mieMessage
    tlMie.sEnteredDate = Format$(rst!mieEnteredDate, sgShowDateForm)
    tlMie.sEnteredTime = Format$(rst!mieEnteredTime, sgShowTimeWSecForm)
    tlMie.iUieCode = rst!mieUieCode
    tlMie.sUnused = rst!mieUnused
    rst.Close
    gGetServiceStatus_MIE_MessageInfo = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    rst.Close
    gGetServiceStatus_MIE_MessageInfo = False
    tlMie.lCode = 0
    Exit Function

End Function

Public Function gGetListInfo(slSQLCall As String)
    Dim ilRet As Integer
    Dim ilUpper As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    Set rst = cnn.Execute(slSQLCall)
    ReDim tgListInfo(0 To 0) As LISTINFO
    ilUpper = 0
    While Not rst.EOF
        tgListInfo(ilUpper).lCode = rst(0).Value
        tgListInfo(ilUpper).sCurrent = rst(1).Value
        tgListInfo(ilUpper).lOrigCode = rst(2).Value
        tgListInfo(ilUpper).iVersion = rst(3).Value
        tgListInfo(ilUpper).sState = rst(4).Value
        ilUpper = ilUpper + 1
        ReDim Preserve tgListInfo(0 To ilUpper) As LISTINFO
        rst.MoveNext
    Wend
    rst.Close
    gGetListInfo = True
    Exit Function
ErrHand:
    rst.Close
    gGetListInfo = False
    Exit Function
End Function
