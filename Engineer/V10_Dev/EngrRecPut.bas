Attribute VB_Name = "EngrRecPut"
'
' Release: 1.0
'
' Description:
'   This file contains the Get Record Modules

Option Explicit
Dim tmAIE As AIE
Private smCurrLibEBEStamp As String
Private tmCurrLibEBE() As EBE
Private smCurrLibDEEStamp As String
Private tmCurrLibDEE() As DEE




Public Function gPutInsert_AAE_As_Aired(tlAAE As AAE, slForm_Module As String) As Integer
'
'   tlAAE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    llLastCode = 0
    llInitCode = tlAAE.lCode
    Do
        If tlAAE.lCode <= 0 Then
            sgSQLQuery = "Select MAX(AAECode) from AAE_AS_Aired"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlAAE.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlAAE.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlAAE.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlAAE.lCode
        sgSQLQuery = "Insert Into AAE_As_Aired ( "
        sgSQLQuery = sgSQLQuery & "aaeCode, "
        sgSQLQuery = sgSQLQuery & "aaeSheCode, "
        sgSQLQuery = sgSQLQuery & "aaeSeeCode, "
        sgSQLQuery = sgSQLQuery & "aaeAirDate, "
        sgSQLQuery = sgSQLQuery & "aaeAirTime, "
        sgSQLQuery = sgSQLQuery & "aaeAutoOff, "
        sgSQLQuery = sgSQLQuery & "aaeData, "
        sgSQLQuery = sgSQLQuery & "aaeSchedule, "
        sgSQLQuery = sgSQLQuery & "aaeTrueTime, "
        sgSQLQuery = sgSQLQuery & "aaeSourceConflict, "
        sgSQLQuery = sgSQLQuery & "aaeSourceUnavail, "
        sgSQLQuery = sgSQLQuery & "aaeSourceItem, "
        sgSQLQuery = sgSQLQuery & "aaeBkupSrceUnavail, "
        sgSQLQuery = sgSQLQuery & "aaeBkupSrceItem, "
        sgSQLQuery = sgSQLQuery & "aaeProtSrceUnavail, "
        sgSQLQuery = sgSQLQuery & "aaeProtSrceItem, "
        sgSQLQuery = sgSQLQuery & "aaeDate, "
        sgSQLQuery = sgSQLQuery & "aaeEventID, "
        sgSQLQuery = sgSQLQuery & "aaeBusName, "
        sgSQLQuery = sgSQLQuery & "aaeBusControl, "
        sgSQLQuery = sgSQLQuery & "aaeEventType, "
        sgSQLQuery = sgSQLQuery & "aaeStartTime, "
        sgSQLQuery = sgSQLQuery & "aaeStartType, "
        sgSQLQuery = sgSQLQuery & "aaeFixedTime, "
        sgSQLQuery = sgSQLQuery & "aaeEndType, "
        sgSQLQuery = sgSQLQuery & "aaeDuration, "
        sgSQLQuery = sgSQLQuery & "aaeOutTime, "
        sgSQLQuery = sgSQLQuery & "aaeMaterialType, "
        sgSQLQuery = sgSQLQuery & "aaeAudioName, "
        sgSQLQuery = sgSQLQuery & "aaeAudioItemID, "
        sgSQLQuery = sgSQLQuery & "aaeAudioISCI, "
        sgSQLQuery = sgSQLQuery & "aaeAudioCrtlChar, "
        sgSQLQuery = sgSQLQuery & "aaeBkupAudioName, "
        sgSQLQuery = sgSQLQuery & "aaeBkupCtrlChar, "
        sgSQLQuery = sgSQLQuery & "aaeProtAudioName, "
        sgSQLQuery = sgSQLQuery & "aaeProtItemID, "
        sgSQLQuery = sgSQLQuery & "aaeProtISCI, "
        sgSQLQuery = sgSQLQuery & "aaeProtCtrlChar, "
        sgSQLQuery = sgSQLQuery & "aaeRelay1, "
        sgSQLQuery = sgSQLQuery & "aaeRelay2, "
        sgSQLQuery = sgSQLQuery & "aaeFollow, "
        sgSQLQuery = sgSQLQuery & "aaeSilenceTime, "
        sgSQLQuery = sgSQLQuery & "aaeSilence1, "
        sgSQLQuery = sgSQLQuery & "aaeSilence2, "
        sgSQLQuery = sgSQLQuery & "aaeSilence3, "
        sgSQLQuery = sgSQLQuery & "aaeSilence4, "
        sgSQLQuery = sgSQLQuery & "aaeNetcueStart, "
        sgSQLQuery = sgSQLQuery & "aaeNetcueEnd, "
        sgSQLQuery = sgSQLQuery & "aaeTitle1, "
        sgSQLQuery = sgSQLQuery & "aaeTitle2, "
        sgSQLQuery = sgSQLQuery & "aaeABCFormat, "
        sgSQLQuery = sgSQLQuery & "aaeABCPgmCode, "
        sgSQLQuery = sgSQLQuery & "aaeABCXDSMode, "
        sgSQLQuery = sgSQLQuery & "aaeABCRecordItem, "
        sgSQLQuery = sgSQLQuery & "aaeEnteredDate, "
        sgSQLQuery = sgSQLQuery & "aaeEnteredTime, "
        sgSQLQuery = sgSQLQuery & "aaeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlAAE.lCode & ", "
        sgSQLQuery = sgSQLQuery & tlAAE.lSheCode & ", "
        sgSQLQuery = sgSQLQuery & tlAAE.lSeeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlAAE.sAirDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & tlAAE.lAirTime & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sAutoOff) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sData) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sSchedule) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sTrueTime) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sSourceConflict) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sSourceUnavail) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sSourceItem) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sBkupSrceUnavail) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sBkupSrceItem) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sProtSrceUnavail) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sProtSrceItem) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sDate) & "', "
        sgSQLQuery = sgSQLQuery & tlAAE.lEventID & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sBusName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sBusControl) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sEventType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sStartTime) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sStartType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sFixedTime) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sEndType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sDuration) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sOutTime) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sMaterialType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sAudioName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sAudioItemID) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sAudioISCI) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sAudioCrtlChar) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sBkupAudioName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sBkupCtrlChar) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sProtAudioName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sProtItemID) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sProtISCI) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sProtCtrlChar) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sRelay1) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sRelay2) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sFollow) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sSilenceTime) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sSilence1) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sSilence2) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sSilence3) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sSilence4) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sNetcueStart) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sNetcueEnd) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sTitle1) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sTitle2) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sABCFormat) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sABCPgmCode) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sABCXDSMode) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sABCRecordItem) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlAAE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlAAE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAAE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery  ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_AAE_As_Aired = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlAAE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_AAE_As_Aired = False
    Exit Function
End Function
Public Function gPutInsert_ACE_AutoContact(tlACE As ACE, slForm_Module As String) As Integer
'
'   tlACE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlACE.iCode
    Do
        If tlACE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(aceCode) from ACE_Auto_Contact"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlACE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlACE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlACE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlACE.iCode
        sgSQLQuery = "INSERT INTO ACE_Auto_Contact ("
        sgSQLQuery = sgSQLQuery & "aceCode, "
        sgSQLQuery = sgSQLQuery & "aceAeeCode, "
        sgSQLQuery = sgSQLQuery & "aceType, "
        sgSQLQuery = sgSQLQuery & "aceContact, "
        sgSQLQuery = sgSQLQuery & "acePhone, "
        sgSQLQuery = sgSQLQuery & "aceFax, "
        sgSQLQuery = sgSQLQuery & "aceEMail, "
        sgSQLQuery = sgSQLQuery & "aceUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlACE.iCode & ", "
        sgSQLQuery = sgSQLQuery & tlACE.iAeeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlACE.sType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlACE.sContact) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlACE.sPhone) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlACE.sFax) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlACE.sEMail) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlACE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery  ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_ACE_AutoContact = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlACE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_ACE_AutoContact = False
    Exit Function
End Function

Public Function gPutInsert_ADE_AutoDataFlags(tlADE As ADE, slForm_Module As String) As Integer
'
'   tlADE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlADE.iCode
    Do
        If tlADE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(adeCode) from ADE_Auto_Data_Flags"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlADE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlADE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlADE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlADE.iCode
        sgSQLQuery = "INSERT INTO ADE_Auto_Data_Flags ("
        sgSQLQuery = sgSQLQuery & "adeCode, "
        sgSQLQuery = sgSQLQuery & "adeAeeCode, "
        sgSQLQuery = sgSQLQuery & "adeScheduleData, "
        sgSQLQuery = sgSQLQuery & "adeDate, "
        sgSQLQuery = sgSQLQuery & "adeDateNoChar, "
        sgSQLQuery = sgSQLQuery & "adeTime, "
        sgSQLQuery = sgSQLQuery & "adeTimeNoChar, "
        sgSQLQuery = sgSQLQuery & "adeAutoOff, "
        sgSQLQuery = sgSQLQuery & "adeData, "
        sgSQLQuery = sgSQLQuery & "adeSchedule, "
        sgSQLQuery = sgSQLQuery & "adeTrueTime, "
        sgSQLQuery = sgSQLQuery & "adeSourceConflict, "
        sgSQLQuery = sgSQLQuery & "adeSourceUnavail, "
        sgSQLQuery = sgSQLQuery & "adeSourceItem, "
        sgSQLQuery = sgSQLQuery & "adeBkupSrceUnavail, "
        sgSQLQuery = sgSQLQuery & "adeBkupSrceItem, "
        sgSQLQuery = sgSQLQuery & "adeProtSrceUnavail, "
        sgSQLQuery = sgSQLQuery & "adeProtSrceItem, "
        sgSQLQuery = sgSQLQuery & "adeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlADE.iCode & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iAeeCode & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iScheduleData & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iDate & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iDateNoChar & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iTime & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iTimeNoChar & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iAutoOff & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iData & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iSchedule & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iTrueTime & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iSourceConflict & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iSourceUnavail & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iSourceItem & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iBkupSrceUnavail & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iBkupSrceItem & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iProtSrceUnavail & ", "
        sgSQLQuery = sgSQLQuery & tlADE.iProtSrceItem & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlADE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery  ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_ADE_AutoDataFlags = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlADE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_ADE_AutoDataFlags = False
    Exit Function
End Function

Public Function gPutInsert_AEE_AutoEquip(ilInsertType As Integer, tlAEE As AEE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigAeeCode); 1=From Update (retain OrigAeeCode)
'   tlAEE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlAEE.iCode
    Do
        If tlAEE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(aeeCode) from AEE_Auto_Equip"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlAEE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlAEE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlAEE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlAEE.iCode
        sgSQLQuery = "INSERT INTO AEE_Auto_Equip ("
        sgSQLQuery = sgSQLQuery & "aeeCode, "
        sgSQLQuery = sgSQLQuery & "aeeName, "
        sgSQLQuery = sgSQLQuery & "aeeDescription, "
        sgSQLQuery = sgSQLQuery & "aeeManufacture, "
        sgSQLQuery = sgSQLQuery & "aeeFixedTimeChar, "
        sgSQLQuery = sgSQLQuery & "aeeAlertSchdDelay, "
        sgSQLQuery = sgSQLQuery & "aeeState, "
        sgSQLQuery = sgSQLQuery & "aeeUsedFlag, "
        sgSQLQuery = sgSQLQuery & "aeeVersion, "
        sgSQLQuery = sgSQLQuery & "aeeOrigAeeCode, "
        sgSQLQuery = sgSQLQuery & "aeeCurrent, "
        sgSQLQuery = sgSQLQuery & "aeeEnteredDate, "
        sgSQLQuery = sgSQLQuery & "aeeEnteredTime, "
        sgSQLQuery = sgSQLQuery & "aeeUieCode, "
        sgSQLQuery = sgSQLQuery & "aeeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlAEE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAEE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAEE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAEE.sManufacture) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAEE.sFixedTimeChar) & "', "
        sgSQLQuery = sgSQLQuery & tlAEE.lAlertSchdDelay & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAEE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAEE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlAEE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlAEE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlAEE.iOrigAeeCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAEE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlAEE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlAEE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlAEE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAEE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ")"
        cnn.BeginTrans
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    cnn.CommitTrans
    gPutInsert_AEE_AutoEquip = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlAEE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_AEE_AutoEquip = False
    Exit Function
End Function

Public Function gPutInsert_AFE_AutoFormat(tlAFE As AFE, slForm_Module As String) As Integer
'
'   tlAFE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlAFE.iCode
    Do
        If tlAFE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(afeCode) from AFE_Auto_Format"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlAFE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlAFE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlAFE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlAFE.iCode
        sgSQLQuery = "INSERT INTO AFE_Auto_Format ("
        sgSQLQuery = sgSQLQuery & "afeCode, "
        sgSQLQuery = sgSQLQuery & "afeAeeCode, "
        sgSQLQuery = sgSQLQuery & "afeType, "
        sgSQLQuery = sgSQLQuery & "afeSubType, "
        sgSQLQuery = sgSQLQuery & "afeBus, "
        sgSQLQuery = sgSQLQuery & "afeBusControl, "
        sgSQLQuery = sgSQLQuery & "afeEventType, "
        sgSQLQuery = sgSQLQuery & "afeTime, "
        sgSQLQuery = sgSQLQuery & "afeStartType, "
        sgSQLQuery = sgSQLQuery & "afeFixedTime, "
        sgSQLQuery = sgSQLQuery & "afeEndType, "
        sgSQLQuery = sgSQLQuery & "afeDuration, "
        sgSQLQuery = sgSQLQuery & "afeEndTime, "
        sgSQLQuery = sgSQLQuery & "afeMaterialType, "
        sgSQLQuery = sgSQLQuery & "afeAudioName, "
        sgSQLQuery = sgSQLQuery & "afeAudioItemID, "
        sgSQLQuery = sgSQLQuery & "afeAudioISCI, "
        sgSQLQuery = sgSQLQuery & "afeAudioControl, "
        sgSQLQuery = sgSQLQuery & "afeBkupAudioName, "
        sgSQLQuery = sgSQLQuery & "afeBkupAudioControl, "
        sgSQLQuery = sgSQLQuery & "afeProtAudioName, "
        sgSQLQuery = sgSQLQuery & "afeProtItemID, "
        sgSQLQuery = sgSQLQuery & "afeProtISCI, "
        sgSQLQuery = sgSQLQuery & "afeProtAudioControl, "
        sgSQLQuery = sgSQLQuery & "afeRelay1, "
        sgSQLQuery = sgSQLQuery & "afeRelay2, "
        sgSQLQuery = sgSQLQuery & "afeFollow, "
        sgSQLQuery = sgSQLQuery & "afeSilenceTime, "
        sgSQLQuery = sgSQLQuery & "afeSilence1, "
        sgSQLQuery = sgSQLQuery & "afeSilence2, "
        sgSQLQuery = sgSQLQuery & "afeSilence3, "
        sgSQLQuery = sgSQLQuery & "afeSilence4, "
        sgSQLQuery = sgSQLQuery & "afeStartNetcue, "
        sgSQLQuery = sgSQLQuery & "afeStopNetcue, "
        sgSQLQuery = sgSQLQuery & "afeTitle1, "
        sgSQLQuery = sgSQLQuery & "afeTitle2, "
        sgSQLQuery = sgSQLQuery & "afeEventID, "
        sgSQLQuery = sgSQLQuery & "afeDate, "
        sgSQLQuery = sgSQLQuery & "afeABCFormat, "
        sgSQLQuery = sgSQLQuery & "afeABCPgmCode, "
        sgSQLQuery = sgSQLQuery & "afeABCXDSMode, "
        sgSQLQuery = sgSQLQuery & "afeABCRecordItem, "
        sgSQLQuery = sgSQLQuery & "afeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlAFE.iCode & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iAeeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAFE.sType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAFE.sSubType) & "', "
        sgSQLQuery = sgSQLQuery & tlAFE.iBus & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iBusControl & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iEventType & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iTime & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iStartType & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iFixedTime & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iEndType & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iDuration & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iEndTime & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iMaterialType & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iAudioName & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iAudioItemID & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iAudioISCI & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iAudioControl & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iBkupAudioName & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iBkupAudioControl & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iProtAudioName & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iProtItemID & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iProtISCI & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iProtAudioControl & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iRelay1 & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iRelay2 & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iFollow & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iSilenceTime & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iSilence1 & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iSilence2 & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iSilence3 & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iSilence4 & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iStartNetcue & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iStopNetcue & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iTitle1 & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iTitle2 & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iEventID & ", "
        sgSQLQuery = sgSQLQuery & tlAFE.iDate & ", "
        If sgClientFields = "A" Then
            sgSQLQuery = sgSQLQuery & tlAFE.iABCFormat & ", "
            sgSQLQuery = sgSQLQuery & tlAFE.iABCPgmCode & ", "
            sgSQLQuery = sgSQLQuery & tlAFE.iABCXDSMode & ", "
            sgSQLQuery = sgSQLQuery & tlAFE.iABCRecordItem & ", "
        Else
            sgSQLQuery = sgSQLQuery & "0" & ", "
            sgSQLQuery = sgSQLQuery & "0" & ", "
            sgSQLQuery = sgSQLQuery & "0" & ", "
            sgSQLQuery = sgSQLQuery & "0" & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAFE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_AFE_AutoFormat = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlAFE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_AFE_AutoFormat = False
    Exit Function
End Function

Public Function gPutInsert_AIE_ActiveInfo(tlAIE As AIE, slForm_Module As String) As Integer
'
'   tlUIE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    llLastCode = 0
    llInitCode = tlAIE.lCode
    Do
        If tlAIE.lCode <= 0 Then
            sgSQLQuery = "Select MAX(aieCode) from AIE_Active_Info"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlAIE.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlAIE.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlAIE.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlAIE.lCode
        sgSQLQuery = "INSERT INTO AIE_Active_Info ("
        sgSQLQuery = sgSQLQuery & "aieCode, "
        sgSQLQuery = sgSQLQuery & "aieUieCode, "
        sgSQLQuery = sgSQLQuery & "aieEnteredDate, "
        sgSQLQuery = sgSQLQuery & "aieEnteredTime, "
        sgSQLQuery = sgSQLQuery & "aieRefFileName, "
        sgSQLQuery = sgSQLQuery & "aieToFileCode, "
        sgSQLQuery = sgSQLQuery & "aieFromFileCode, "
        sgSQLQuery = sgSQLQuery & "aieOrigFileCode, "
        sgSQLQuery = sgSQLQuery & "aieUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlAIE.lCode & ", "
        sgSQLQuery = sgSQLQuery & tlAIE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlAIE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlAIE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAIE.sRefFileName) & "', "
        sgSQLQuery = sgSQLQuery & tlAIE.lToFileCode & ", "
        sgSQLQuery = sgSQLQuery & tlAIE.lFromFileCode & ", "
        sgSQLQuery = sgSQLQuery & tlAIE.lOrigFileCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAIE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_AIE_ActiveInfo = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlAIE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_AIE_ActiveInfo = False
    Exit Function
End Function

Public Function gPutInsert_ANE_AudioName(ilInsertType As Integer, tlANE As ANE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigAneCode); 1=From Update (retain OrigAneCode)
'   tlANE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlANE.iCode
    Do
        If tlANE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(aneCode) from ANE_Audio_Name"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlANE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlANE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlANE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlANE.iCode
        sgSQLQuery = "Insert Into ANE_Audio_Name ( "
        sgSQLQuery = sgSQLQuery & "aneCode, "
        sgSQLQuery = sgSQLQuery & "aneName, "
        sgSQLQuery = sgSQLQuery & "aneDescription, "
        sgSQLQuery = sgSQLQuery & "aneCCECode, "
        sgSQLQuery = sgSQLQuery & "aneAteCode, "
        sgSQLQuery = sgSQLQuery & "aneState, "
        sgSQLQuery = sgSQLQuery & "aneUsedFlag, "
        sgSQLQuery = sgSQLQuery & "aneVersion, "
        sgSQLQuery = sgSQLQuery & "aneOrigAneCode, "
        sgSQLQuery = sgSQLQuery & "aneCurrent, "
        sgSQLQuery = sgSQLQuery & "aneEnteredDate, "
        sgSQLQuery = sgSQLQuery & "aneEnteredTime, "
        sgSQLQuery = sgSQLQuery & "aneUieCode, "
        sgSQLQuery = sgSQLQuery & "aneCheckConflicts, "
        sgSQLQuery = sgSQLQuery & "aneUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlANE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlANE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlANE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & tlANE.iCceCode & ", "
        sgSQLQuery = sgSQLQuery & tlANE.iAteCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlANE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlANE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlANE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlANE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlANE.iOrigAneCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlANE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlANE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlANE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlANE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlANE.sCheckConflicts) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlANE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_ANE_AudioName = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlANE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_ANE_AudioName = False
    Exit Function
End Function


Public Function gPutInsert_APE_AutoPath(tlAPE As APE, slForm_Module As String) As Integer
'
'   tlAPE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlAPE.iCode
    Do
        If tlAPE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(apeCode) from APE_Auto_Path"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlAPE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlAPE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlAPE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlAPE.iCode
        sgSQLQuery = "INSERT INTO APE_Auto_Path ("
        sgSQLQuery = sgSQLQuery & "apeCode, "
        sgSQLQuery = sgSQLQuery & "apeAeeCode, "
        sgSQLQuery = sgSQLQuery & "apeType, "
        sgSQLQuery = sgSQLQuery & "apeSubType, "
        sgSQLQuery = sgSQLQuery & "apeNewFileName, "
        sgSQLQuery = sgSQLQuery & "apeChgFileName, "
        sgSQLQuery = sgSQLQuery & "apeDelFileName, "
        sgSQLQuery = sgSQLQuery & "apeNewFileExt, "
        sgSQLQuery = sgSQLQuery & "apeChgFileExt, "
        sgSQLQuery = sgSQLQuery & "apeDelFileExt, "
        sgSQLQuery = sgSQLQuery & "apePath, "
        sgSQLQuery = sgSQLQuery & "apeDateFormat, "
        sgSQLQuery = sgSQLQuery & "apeTimeFormat, "
        sgSQLQuery = sgSQLQuery & "apeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlAPE.iCode & ", "
        sgSQLQuery = sgSQLQuery & tlAPE.iAeeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAPE.sType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAPE.sSubType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAPE.sNewFileName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAPE.sChgFileName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAPE.sDelFileName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAPE.sNewFileExt) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAPE.sChgFileExt) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAPE.sDelFileExt) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAPE.sPath) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAPE.sDateFormat) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAPE.sTimeFormat) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlAPE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_APE_AutoPath = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlAPE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_APE_AutoPath = False
    Exit Function
End Function

Public Function gPutInsert_ARE_AdvertiserRefer(tlARE As ARE, slForm_Module As String) As Integer
'
'   tlARE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    sgSQLQuery = "Select * from ARE_Advertiser_Refer WHERE areName = '" & gFixQuote(Trim$(tlARE.sName)) & "'"
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        tlARE.lCode = rst!areCode
        rst.Close
        gPutInsert_ARE_AdvertiserRefer = True
        Exit Function
    End If
    rst.Close
    llLastCode = 0
    llInitCode = tlARE.lCode
    Do
        If tlARE.lCode <= 0 Then
            sgSQLQuery = "Select MAX(areCode) from ARE_Advertiser_Refer"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlARE.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlARE.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlARE.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlARE.lCode
        sgSQLQuery = "INSERT INTO ARE_Advertiser_Refer ("
        sgSQLQuery = sgSQLQuery & "areCode, "
        sgSQLQuery = sgSQLQuery & "areName, "
        sgSQLQuery = sgSQLQuery & "areUnusued "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlARE.lCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlARE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlARE.sUnusued) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_ARE_AdvertiserRefer = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlARE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_ARE_AdvertiserRefer = False
    Exit Function
End Function


Public Function gPutInsert_ASE_AudioSource(ilInsertType As Integer, tlASE As ASE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigAseCode); 1=From Update (retain OrigAseCode)
'   tlASE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlASE.iCode
    Do
        If tlASE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(aseCode) from ASE_Audio_Source"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlASE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlASE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlASE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlASE.iCode
        sgSQLQuery = "Insert Into ASE_Audio_Source ( "
        sgSQLQuery = sgSQLQuery & "aseCode, "
        sgSQLQuery = sgSQLQuery & "asePriAneCode, "
        sgSQLQuery = sgSQLQuery & "asePriCceCode, "
        sgSQLQuery = sgSQLQuery & "aseDescription, "
        sgSQLQuery = sgSQLQuery & "aseBkupAneCode, "
        sgSQLQuery = sgSQLQuery & "aseBkupCceCode, "
        sgSQLQuery = sgSQLQuery & "aseProtAneCode, "
        sgSQLQuery = sgSQLQuery & "aseProtCceCode, "
        sgSQLQuery = sgSQLQuery & "aseState, "
        sgSQLQuery = sgSQLQuery & "aseUsedFlag, "
        sgSQLQuery = sgSQLQuery & "aseVersion, "
        sgSQLQuery = sgSQLQuery & "aseOrigAseCode, "
        sgSQLQuery = sgSQLQuery & "aseCurrent, "
        sgSQLQuery = sgSQLQuery & "aseEnteredDate, "
        sgSQLQuery = sgSQLQuery & "aseEnteredTime, "
        sgSQLQuery = sgSQLQuery & "aseUieCode, "
        sgSQLQuery = sgSQLQuery & "aseUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlASE.iCode & ", "
        sgSQLQuery = sgSQLQuery & tlASE.iPriAneCode & ", "
        sgSQLQuery = sgSQLQuery & tlASE.iPriCceCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlASE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & tlASE.iBkupAneCode & ", "
        sgSQLQuery = sgSQLQuery & tlASE.iBkupCceCode & ", "
        sgSQLQuery = sgSQLQuery & tlASE.iProtAneCode & ", "
        sgSQLQuery = sgSQLQuery & tlASE.iProtCceCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlASE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlASE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlASE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlASE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlASE.iOrigAseCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlASE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlASE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlASE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlASE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlASE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_ASE_AudioSource = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlASE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_ASE_AudioSource = False
    Exit Function
End Function

Public Function gPutInsert_ATE_AudioType(ilInsertType As Integer, tlATE As ATE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigAteCode); 1=From Update (retain OrigAteCode)
'   tlATE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlATE.iCode
    Do
        If tlATE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(ateCode) from ATE_Audio_Type"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlATE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlATE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlATE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlATE.iCode
        sgSQLQuery = "Insert Into ATE_Audio_Type ( "
        sgSQLQuery = sgSQLQuery & "ateCode, "
        sgSQLQuery = sgSQLQuery & "ateName, "
        sgSQLQuery = sgSQLQuery & "ateDescription, "
        sgSQLQuery = sgSQLQuery & "ateState, "
        sgSQLQuery = sgSQLQuery & "ateTestItemID, "
        sgSQLQuery = sgSQLQuery & "atePreBufferTime, "
        sgSQLQuery = sgSQLQuery & "atePostBufferTime, "
        sgSQLQuery = sgSQLQuery & "ateUsedFlag, "
        sgSQLQuery = sgSQLQuery & "ateVersion, "
        sgSQLQuery = sgSQLQuery & "ateOrigAteCode, "
        sgSQLQuery = sgSQLQuery & "ateCurrent, "
        sgSQLQuery = sgSQLQuery & "ateEnteredDate, "
        sgSQLQuery = sgSQLQuery & "ateEnteredTime, "
        sgSQLQuery = sgSQLQuery & "ateUieCode, "
        sgSQLQuery = sgSQLQuery & "ateUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlATE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlATE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlATE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlATE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlATE.sTestItemID) & "', "
        sgSQLQuery = sgSQLQuery & tlATE.lPreBufferTime & ", "
        sgSQLQuery = sgSQLQuery & tlATE.lPostBufferTime & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlATE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlATE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlATE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlATE.iOrigAteCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlATE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlATE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlATE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlATE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlATE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_ATE_AudioType = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlATE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_ATE_AudioType = False
    Exit Function
End Function

Public Function gPutInsert_BDE_BusDefinition(ilInsertType As Integer, tlBDE As BDE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigBdeCode); 1=From Update (retain OrigBdeCode)
'   tlBDE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlBDE.iCode
    Do
        If tlBDE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(bdeCode) from BDE_Bus_Definition"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlBDE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlBDE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlBDE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlBDE.iCode
        sgSQLQuery = "Insert Into BDE_Bus_Definition ( "
        sgSQLQuery = sgSQLQuery & "bdeCode, "
        sgSQLQuery = sgSQLQuery & "bdeName, "
        sgSQLQuery = sgSQLQuery & "bdeDescription, "
        sgSQLQuery = sgSQLQuery & "bdeChannel, "
        sgSQLQuery = sgSQLQuery & "bdeAseCode, "
        sgSQLQuery = sgSQLQuery & "bdeState, "
        sgSQLQuery = sgSQLQuery & "bdeCceCode, "
        sgSQLQuery = sgSQLQuery & "bdeUsedFlag, "
        sgSQLQuery = sgSQLQuery & "bdeVersion, "
        sgSQLQuery = sgSQLQuery & "bdeOrigBdeCode, "
        sgSQLQuery = sgSQLQuery & "bdeCurrent, "
        sgSQLQuery = sgSQLQuery & "bdeEnteredDate, "
        sgSQLQuery = sgSQLQuery & "bdeEnteredTime, "
        sgSQLQuery = sgSQLQuery & "bdeUieCode, "
        sgSQLQuery = sgSQLQuery & "bdeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlBDE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBDE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBDE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBDE.sChannel) & "', "
        sgSQLQuery = sgSQLQuery & tlBDE.iAseCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBDE.sState) & "', "
        sgSQLQuery = sgSQLQuery & tlBDE.iCceCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBDE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlBDE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlBDE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlBDE.iOrigBdeCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBDE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlBDE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlBDE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlBDE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBDE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_BDE_BusDefinition = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlBDE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_BDE_BusDefinition = False
    Exit Function
End Function

Public Function gPutInsert_BGE_BusGroup(ilInsertType As Integer, tlBGE As BGE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigBgeCode); 1=From Update (retain OrigBgeCode)
'   tlBGE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlBGE.iCode
    Do
        If tlBGE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(bgeCode) from BGE_Bus_Group"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlBGE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlBGE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlBGE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlBGE.iCode
        sgSQLQuery = "Insert Into BGE_Bus_Group ( "
        sgSQLQuery = sgSQLQuery & "bgeCode, "
        sgSQLQuery = sgSQLQuery & "bgeName, "
        sgSQLQuery = sgSQLQuery & "bgeDescription, "
        sgSQLQuery = sgSQLQuery & "bgeState, "
        sgSQLQuery = sgSQLQuery & "bgeUsedFlag, "
        sgSQLQuery = sgSQLQuery & "bgeVersion, "
        sgSQLQuery = sgSQLQuery & "bgeOrigBgeCode, "
        sgSQLQuery = sgSQLQuery & "bgeCurrent, "
        sgSQLQuery = sgSQLQuery & "bgeEnteredDate, "
        sgSQLQuery = sgSQLQuery & "bgeEnteredTime, "
        sgSQLQuery = sgSQLQuery & "bgeUieCode, "
        sgSQLQuery = sgSQLQuery & "bgeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlBGE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBGE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBGE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBGE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBGE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlBGE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlBGE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlBGE.iOrigBgeCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBGE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlBGE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlBGE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlBGE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBGE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_BGE_BusGroup = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlBGE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_BGE_BusGroup = False
    Exit Function
End Function

Public Function gPutInsert_BSE_BusSelGroup(tlBSE As BSE, slForm_Module As String) As Integer
'
'   tlBSE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlBSE.iCode
    Do
        If tlBSE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(bseCode) from BSE_Bus_Sel_Group"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlBSE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlBSE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlBSE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlBSE.iCode
        sgSQLQuery = "Insert Into BSE_Bus_Sel_Group ( "
        sgSQLQuery = sgSQLQuery & "bseCode, "
        sgSQLQuery = sgSQLQuery & "bseBdeCode, "
        sgSQLQuery = sgSQLQuery & "bseBgeCode, "
        sgSQLQuery = sgSQLQuery & "bseUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlBSE.iCode & ", "
        sgSQLQuery = sgSQLQuery & tlBSE.iBdeCode & ", "
        sgSQLQuery = sgSQLQuery & tlBSE.iBgeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlBSE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "

        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_BSE_BusSelGroup = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlBSE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_BSE_BusSelGroup = False
    Exit Function
End Function

Public Function gPutInsert_CCE_ControlChar(ilInsertType As Integer, tlCCE As CCE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigCceCode); 1=From Update (retain OrigCceCode)
'   tlCCE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlCCE.iCode
    Do
        If tlCCE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(cceCode) from CCE_Control_Char"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlCCE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlCCE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlCCE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlCCE.iCode
        sgSQLQuery = "Insert Into CCE_Control_Char ( "
        sgSQLQuery = sgSQLQuery & "cceCode, "
        sgSQLQuery = sgSQLQuery & "cceType, "
        sgSQLQuery = sgSQLQuery & "cceAutoChar, "
        sgSQLQuery = sgSQLQuery & "cceDescription, "
        sgSQLQuery = sgSQLQuery & "cceState, "
        sgSQLQuery = sgSQLQuery & "cceUsedFlag, "
        sgSQLQuery = sgSQLQuery & "cceVersion, "
        sgSQLQuery = sgSQLQuery & "cceOrigCceCode, "
        sgSQLQuery = sgSQLQuery & "cceCurrent, "
        sgSQLQuery = sgSQLQuery & "cceEnteredDate, "
        sgSQLQuery = sgSQLQuery & "cceEnteredTime, "
        sgSQLQuery = sgSQLQuery & "cceUieCode, "
        sgSQLQuery = sgSQLQuery & "cceUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlCCE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCCE.sType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCCE.sAutoChar) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCCE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCCE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCCE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlCCE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlCCE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlCCE.iOrigCceCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCCE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlCCE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlCCE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlCCE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCCE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_CCE_ControlChar = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlCCE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_CCE_ControlChar = False
    Exit Function
End Function

Public Function gPutInsert_CEE_Conflict_Events(tlCEE As CEE, slForm_Module As String) As Integer
'
'   tlCEE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    '12/11/09:  Remove to make Saving Libraries faster
    gPutInsert_CEE_Conflict_Events = True
    Exit Function
    
    
    On Error GoTo ErrHand
    llLastCode = 0
    llInitCode = tlCEE.lCode
    Do
        If tlCEE.lCode <= 0 Then
            sgSQLQuery = "Select MAX(ceeCode) from CEE_Conflict_Events"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlCEE.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlCEE.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlCEE.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlCEE.lCode
        sgSQLQuery = "Insert Into CEE_Conflict_Events ( "
        sgSQLQuery = sgSQLQuery & "ceeCode, "
        sgSQLQuery = sgSQLQuery & "ceeGenDate, "
        sgSQLQuery = sgSQLQuery & "ceeGenTime, "
        sgSQLQuery = sgSQLQuery & "ceeEvtType, "
        sgSQLQuery = sgSQLQuery & "ceeBDECode, "
        sgSQLQuery = sgSQLQuery & "ceeANECode, "
        sgSQLQuery = sgSQLQuery & "ceeStartDate, "
        sgSQLQuery = sgSQLQuery & "ceeEndDate, "
        sgSQLQuery = sgSQLQuery & "ceeDay, "
        sgSQLQuery = sgSQLQuery & "ceeStartTime, "
        sgSQLQuery = sgSQLQuery & "ceeEndTime, "
        sgSQLQuery = sgSQLQuery & "ceeGridEventRow, "
        sgSQLQuery = sgSQLQuery & "ceeGridEventCol, "
        sgSQLQuery = sgSQLQuery & "ceeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlCEE.lCode & ", "
        sgSQLQuery = sgSQLQuery & tlCEE.lGenDate & ", "
        sgSQLQuery = sgSQLQuery & tlCEE.lGenTime & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCEE.sEvtType) & "', "
        sgSQLQuery = sgSQLQuery & tlCEE.iBdeCode & ", "
        sgSQLQuery = sgSQLQuery & tlCEE.iANECode & ", "
        sgSQLQuery = sgSQLQuery & tlCEE.lStartDate & ", "
        sgSQLQuery = sgSQLQuery & tlCEE.lEndDate & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCEE.sDay) & "', "
        sgSQLQuery = sgSQLQuery & tlCEE.lStartTime & ", "
        sgSQLQuery = sgSQLQuery & tlCEE.lEndTime & ", "
        sgSQLQuery = sgSQLQuery & tlCEE.lGridEventRow & ", "
        sgSQLQuery = sgSQLQuery & tlCEE.iGridEventCol & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCEE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_CEE_Conflict_Events = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlCEE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_CEE_Conflict_Events = False
    Exit Function
End Function

Public Function gPutInsert_CME_Conflict_Master(tlCME As CME, slForm_Module As String, hlCME As Integer) As Integer
'
'   tlCME(I)- Record to be added to Database
'
    Dim ilRet As Integer
    Dim llLastCode As Long
    Dim llInitCode As Long
    
    '12/11/09:  Remove to make Saving Libraries faster
    gPutInsert_CME_Conflict_Master = True
    Exit Function
    
    
    On Error GoTo ErrHand
'    llLastCode = 0
'    llInitCode = tlCME.lCode
'    Do
'        If tlCME.lCode <= 0 Then
'            sgSQLQuery = "Select MAX(cmeCode) from CME_Conflict_Master"
'            Set rst = cnn.Execute(sgSQLQuery)
'            If IsNull(rst(0).Value) Then
'                tlCME.lCode = 1
'            Else
'                If rst(0).Value > 0 Then
'                    tlCME.lCode = rst(0).Value + 1
'                End If
'            End If
'            If llLastCode = tlCME.lCode Then
'                GoTo ErrHand1:
'            End If
'        End If
'        sgSQLQuery = "Insert Into CME_Conflict_Master ( "
'        sgSQLQuery = sgSQLQuery & "cmeCode, "
'        sgSQLQuery = sgSQLQuery & "cmeSource, "
'        sgSQLQuery = sgSQLQuery & "cmeSHEDHECode, "
'        sgSQLQuery = sgSQLQuery & "cmeDSECode, "
'        sgSQLQuery = sgSQLQuery & "cmeDEECode, "
'        sgSQLQuery = sgSQLQuery & "cmeSEECode, "
'        sgSQLQuery = sgSQLQuery & "cmeEvtType, "
'        sgSQLQuery = sgSQLQuery & "cmeBDECode, "
'        sgSQLQuery = sgSQLQuery & "cmeANECode, "
'        sgSQLQuery = sgSQLQuery & "cmeStartDate, "
'        sgSQLQuery = sgSQLQuery & "cmeEndDate, "
'        sgSQLQuery = sgSQLQuery & "cmeDay, "
'        sgSQLQuery = sgSQLQuery & "cmeStartTime, "
'        sgSQLQuery = sgSQLQuery & "cmeEndTime, "
'        sgSQLQuery = sgSQLQuery & "cmeItemID, "
'        sgSQLQuery = sgSQLQuery & "cmeXMidNight, "
'        sgSQLQuery = sgSQLQuery & "cmeUnused "
'        sgSQLQuery = sgSQLQuery & ") "
'        sgSQLQuery = sgSQLQuery & "Values ( "
'        sgSQLQuery = sgSQLQuery & tlCME.lCode & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCME.sSource) & "', "
'        sgSQLQuery = sgSQLQuery & tlCME.lSHEDHECode & ", "
'        sgSQLQuery = sgSQLQuery & tlCME.lDSECode & ", "
'        sgSQLQuery = sgSQLQuery & tlCME.lDeeCode & ", "
'        sgSQLQuery = sgSQLQuery & tlCME.lSEECode & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCME.sEvtType) & "', "
'        sgSQLQuery = sgSQLQuery & tlCME.iBDECode & ", "
'        sgSQLQuery = sgSQLQuery & tlCME.iANECode & ", "
'        sgSQLQuery = sgSQLQuery & tlCME.lStartDate & ", "
'        sgSQLQuery = sgSQLQuery & tlCME.lEndDate & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCME.sDay) & "', "
'        sgSQLQuery = sgSQLQuery & tlCME.lStartTime & ", "
'        sgSQLQuery = sgSQLQuery & tlCME.lEndTime & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCME.sItemID) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCME.sXMidNight) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCME.sUnused) & "' "
'        sgSQLQuery = sgSQLQuery & ") "
'        ilRet = 0
'        cnn.Execute sgSQLQuery    ', rdExecDirect
'    Loop While ilRet = BTRV_ERR_DUPLICATE_KEY
    Dim ilCMERecLen As Integer
    
    ilCMERecLen = Len(tlCME)
    ilRet = btrInsert(hlCME, tlCME, ilCMERecLen, 0)
    If ilRet <> BTRV_ERR_NONE Then
        gPutInsert_CME_Conflict_Master = False
        Exit Function
    End If
    gPutInsert_CME_Conflict_Master = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            ilRet = gErrSQL.NativeError
            If ilRet < 0 Then
                ilRet = ilRet + 4999
            End If
            If (ilRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlCME.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_CME_Conflict_Master = False
    Exit Function
End Function

Public Function gPutInsert_CTE_CommtsTitle(ilInsertType As Integer, tlCTE As CTE, slForm_Module As String, hlCTE As Integer) As Integer
'
'   ilInsertType(I) = 0 New (set OrigCteCode); 1=From Update (retain OrigCteCode)
'   tlCTE(I)- Record to be added to Database
'
    Dim ilRet As Integer
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim tlCTEAPI As CTEAPI
    
    On Error GoTo ErrHand
'    llLastCode = 0
    llInitCode = tlCTE.lCode
'    Do
'        If tlCTE.lCode <= 0 Then
'            sgSQLQuery = "Select MAX(cteCode) from CTE_Commts_And_Title"
'            Set rst = cnn.Execute(sgSQLQuery)
'            If IsNull(rst(0).Value) Then
'                tlCTE.lCode = 1
'            Else
'                If rst(0).Value > 0 Then
'                    tlCTE.lCode = rst(0).Value + 1
'                End If
'            End If
'            If llLastCode = tlCTE.lCode Then
'                GoTo ErrHand1:
'            End If
'        End If
'        sgSQLQuery = "Insert Into CTE_Commts_And_Title ( "
'        sgSQLQuery = sgSQLQuery & "cteCode, "
'        sgSQLQuery = sgSQLQuery & "cteType, "
'        sgSQLQuery = sgSQLQuery & "cteName, "
'        sgSQLQuery = sgSQLQuery & "cteComment, "
'        sgSQLQuery = sgSQLQuery & "cteState, "
'        sgSQLQuery = sgSQLQuery & "cteUsedFlag, "
'        sgSQLQuery = sgSQLQuery & "cteVersion, "
'        sgSQLQuery = sgSQLQuery & "cteOrigCteCode, "
'        sgSQLQuery = sgSQLQuery & "cteCurrent, "
'        sgSQLQuery = sgSQLQuery & "cteEnteredDate, "
'        sgSQLQuery = sgSQLQuery & "cteEnteredTime, "
'        sgSQLQuery = sgSQLQuery & "cteUieCode, "
'        sgSQLQuery = sgSQLQuery & "cteUnused "
'        sgSQLQuery = sgSQLQuery & ") "
'        sgSQLQuery = sgSQLQuery & "Values ( "
'        sgSQLQuery = sgSQLQuery & tlCTE.lCode & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCTE.sType) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCTE.sName) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCTE.sComment) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCTE.sState) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCTE.sUsedFlag) & "', "
'        sgSQLQuery = sgSQLQuery & tlCTE.iVersion & ", "
'        If (llInitCode <= 0) And (ilInsertType = 0) Then
'            sgSQLQuery = sgSQLQuery & tlCTE.lCode & ", "
'        Else
'            sgSQLQuery = sgSQLQuery & tlCTE.lOrigCteCode & ", "
'        End If
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCTE.sCurrent) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & Format$(tlCTE.sEnteredDate, sgSQLDateForm) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & Format$(tlCTE.sEnteredTime, sgSQLTimeForm) & "', "
'        sgSQLQuery = sgSQLQuery & tlCTE.iUieCode & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlCTE.sUnused) & "' "
'        sgSQLQuery = sgSQLQuery & ") "
'        ilRet = 0
'        cnn.Execute sgSQLQuery    ', rdExecDirect
'    Loop While ilRet = BTRV_ERR_DUPLICATE_KEY
    Dim ilCTERecLen As Integer
    ilCTERecLen = Len(tlCTEAPI)
    tlCTEAPI.lCode = tlCTE.lCode
    tlCTEAPI.sType = tlCTE.sType
    tlCTEAPI.sComment = tlCTE.sComment
    tlCTEAPI.sState = tlCTE.sState
    tlCTEAPI.sUsedFlag = tlCTE.sUsedFlag
    tlCTEAPI.iVersion = tlCTE.iVersion
    tlCTEAPI.lOrigCteCode = tlCTE.lOrigCteCode
    tlCTEAPI.sCurrent = tlCTE.sCurrent
    gPackDate gAdjYear(tlCTE.sEnteredDate), tlCTEAPI.iEneteredDate(0), tlCTEAPI.iEneteredDate(1)
    gPackTime tlCTE.sEnteredTime, tlCTEAPI.iEnteredTime(0), tlCTEAPI.iEnteredTime(1)
    tlCTEAPI.iUieCode = tlCTE.iUieCode
    tlCTEAPI.sUnused = tlCTE.sUnused
    ilRet = btrInsert(hlCTE, tlCTEAPI, ilCTERecLen, 0)
    If ilRet <> BTRV_ERR_NONE Then
        gPutInsert_CTE_CommtsTitle = False
        Exit Function
    End If
    tlCTE.lCode = tlCTEAPI.lCode
    If llInitCode <= 0 Then
        tlCTE.lOrigCteCode = tlCTEAPI.lCode
        tlCTEAPI.lOrigCteCode = tlCTEAPI.lCode
        ilRet = btrUpdate(hlCTE, tlCTEAPI, ilCTERecLen)
    End If
    gPutInsert_CTE_CommtsTitle = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            ilRet = gErrSQL.NativeError
            If ilRet < 0 Then
                ilRet = ilRet + 4999
            End If
            If (ilRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlCTE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_CTE_CommtsTitle = False
    Exit Function
End Function


Public Function gPutInsert_DEE_DayEvent(tlDEE As DEE, slForm_Module As String) As Integer
'
'   tlDEE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    llLastCode = 0
    llInitCode = tlDEE.lCode
    Do
        If tlDEE.lCode <= 0 Then
            sgSQLQuery = "Select MAX(deeCode) from DEE_Day_Event_Info"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlDEE.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlDEE.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlDEE.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlDEE.lCode
        sgSQLQuery = "Insert Into DEE_Day_Event_Info ( "
        sgSQLQuery = sgSQLQuery & "deeCode, "
        sgSQLQuery = sgSQLQuery & "deeDheCode, "
        sgSQLQuery = sgSQLQuery & "deeCceCode, "
        sgSQLQuery = sgSQLQuery & "deeEteCode, "
        sgSQLQuery = sgSQLQuery & "deeTime, "
        sgSQLQuery = sgSQLQuery & "deeStartTteCode, "
        sgSQLQuery = sgSQLQuery & "deeFixedTime, "
        sgSQLQuery = sgSQLQuery & "deeEndTteCode, "
        sgSQLQuery = sgSQLQuery & "deeDuration, "
        sgSQLQuery = sgSQLQuery & "deeHours, "
        sgSQLQuery = sgSQLQuery & "deeDays, "
        sgSQLQuery = sgSQLQuery & "deeMteCode, "
        sgSQLQuery = sgSQLQuery & "deeAudioAseCode, "
        sgSQLQuery = sgSQLQuery & "deeAudioItemID, "
        sgSQLQuery = sgSQLQuery & "deeAudioISCI, "
        sgSQLQuery = sgSQLQuery & "deeAudioCceCode, "
        sgSQLQuery = sgSQLQuery & "deeBkupAneCode, "
        sgSQLQuery = sgSQLQuery & "deeBkupCceCode, "
        sgSQLQuery = sgSQLQuery & "deeProtAneCode, "
        sgSQLQuery = sgSQLQuery & "deeProtItemID, "
        sgSQLQuery = sgSQLQuery & "deeProtISCI, "
        sgSQLQuery = sgSQLQuery & "deeProtCceCode, "
        sgSQLQuery = sgSQLQuery & "dee1RneCode, "
        sgSQLQuery = sgSQLQuery & "dee2RneCode, "
        sgSQLQuery = sgSQLQuery & "deeFneCode, "
        sgSQLQuery = sgSQLQuery & "deeSilenceTime, "
        sgSQLQuery = sgSQLQuery & "dee1SceCode, "
        sgSQLQuery = sgSQLQuery & "dee2SceCode, "
        sgSQLQuery = sgSQLQuery & "dee3SceCode, "
        sgSQLQuery = sgSQLQuery & "dee4SceCode, "
        sgSQLQuery = sgSQLQuery & "deeStartNneCode, "
        sgSQLQuery = sgSQLQuery & "deeEndNneCode, "
        sgSQLQuery = sgSQLQuery & "dee1CteCode, "
        sgSQLQuery = sgSQLQuery & "dee2CteCode, "
        sgSQLQuery = sgSQLQuery & "deeEventID, "
        sgSQLQuery = sgSQLQuery & "deeIgnoreConflicts, "
        sgSQLQuery = sgSQLQuery & "deeABCFormat, "
        sgSQLQuery = sgSQLQuery & "deeABCPgmCode, "
        sgSQLQuery = sgSQLQuery & "deeABCXDSMode, "
        sgSQLQuery = sgSQLQuery & "deeABCRecordItem, "
        sgSQLQuery = sgSQLQuery & "deeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlDEE.lCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.lDheCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.iCceCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.iEteCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.lTime & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.iStartTteCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sFixedTime) & "', "
        sgSQLQuery = sgSQLQuery & tlDEE.iEndTteCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.lDuration & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sHours) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sDays) & "', "
        sgSQLQuery = sgSQLQuery & tlDEE.iMteCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.iAudioAseCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sAudioItemID) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sAudioISCI) & "', "
        sgSQLQuery = sgSQLQuery & tlDEE.iAudioCceCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.iBkupAneCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.iBkupCceCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.iProtAneCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sProtItemID) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sProtISCI) & "', "
        sgSQLQuery = sgSQLQuery & tlDEE.iProtCceCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.i1RneCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.i2RneCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.iFneCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.lSilenceTime & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.i1SceCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.i2SceCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.i3SceCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.i4SceCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.iStartNneCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.iEndNneCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.l1CteCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.l2CteCode & ", "
        sgSQLQuery = sgSQLQuery & tlDEE.lEventID & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sIgnoreConflicts) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sABCFormat) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sABCPgmCode) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sABCXDSMode) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sABCRecordItem) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDEE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_DEE_DayEvent = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlDEE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_DEE_DayEvent = False
    Exit Function
End Function


Public Function gPutInsert_DHE_DayHeaderInfo(ilInsertType As Integer, tlDHE As DHE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigDheCode); 1=From Update (retain OrigDheCode)
'   tlDHE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    llLastCode = 0
    llInitCode = tlDHE.lCode
    Do
        If tlDHE.lCode <= 0 Then
            sgSQLQuery = "Select MAX(dheCode) from DHE_Day_Header_Info"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlDHE.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlDHE.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlDHE.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlDHE.lCode
        sgSQLQuery = "Insert Into DHE_Day_Header_Info ( "
        sgSQLQuery = sgSQLQuery & "dheCode, "
        sgSQLQuery = sgSQLQuery & "dheType, "
        sgSQLQuery = sgSQLQuery & "dheDneCode, "
        sgSQLQuery = sgSQLQuery & "dheDseCode, "
        sgSQLQuery = sgSQLQuery & "dheStartTime, "
        sgSQLQuery = sgSQLQuery & "dheLength, "
        sgSQLQuery = sgSQLQuery & "dheHours, "
        sgSQLQuery = sgSQLQuery & "dheStartDate, "
        sgSQLQuery = sgSQLQuery & "dheEndDate, "
        sgSQLQuery = sgSQLQuery & "dheDays, "
        sgSQLQuery = sgSQLQuery & "dheCteCode, "
        sgSQLQuery = sgSQLQuery & "dheState, "
        sgSQLQuery = sgSQLQuery & "dheUsedFlag, "
        sgSQLQuery = sgSQLQuery & "dheVersion, "
        sgSQLQuery = sgSQLQuery & "dheOrigDheCode, "
        sgSQLQuery = sgSQLQuery & "dheCurrent, "
        sgSQLQuery = sgSQLQuery & "dheEnteredDate, "
        sgSQLQuery = sgSQLQuery & "dheEnteredTime, "
        sgSQLQuery = sgSQLQuery & "dheUieCode, "
        sgSQLQuery = sgSQLQuery & "dheIgnoreConflicts, "
        sgSQLQuery = sgSQLQuery & "dheBusNames, "
        sgSQLQuery = sgSQLQuery & "dheUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlDHE.lCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDHE.sType) & "', "
        sgSQLQuery = sgSQLQuery & tlDHE.lDneCode & ", "
        sgSQLQuery = sgSQLQuery & tlDHE.lDseCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlDHE.sStartTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlDHE.lLength & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDHE.sHours) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlDHE.sStartDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlDHE.sEndDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDHE.sDays) & "', "
        sgSQLQuery = sgSQLQuery & tlDHE.lCteCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDHE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDHE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlDHE.iVersion & ", "
        If (llInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlDHE.lCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlDHE.lOrigDHECode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDHE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlDHE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlDHE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlDHE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDHE.sIgnoreConflicts) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDHE.sBusNames) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDHE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_DHE_DayHeaderInfo = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlDHE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_DHE_DayHeaderInfo = False
    Exit Function
End Function

Public Function gPutInsert_DNE_DayName(ilInsertType As Integer, tlDNE As DNE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigDneCode); 1=From Update (retain OrigDneCode)
'   tlDNE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    llLastCode = 0
    llInitCode = tlDNE.lCode
    Do
        If tlDNE.lCode <= 0 Then
            sgSQLQuery = "Select MAX(dneCode) from DNE_Day_Name"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlDNE.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlDNE.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlDNE.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlDNE.lCode
        sgSQLQuery = "Insert Into DNE_Day_Name ( "
        sgSQLQuery = sgSQLQuery & "dneCode, "
        sgSQLQuery = sgSQLQuery & "dneType, "
        sgSQLQuery = sgSQLQuery & "dneName, "
        sgSQLQuery = sgSQLQuery & "dneDescription, "
        sgSQLQuery = sgSQLQuery & "dneState, "
        sgSQLQuery = sgSQLQuery & "dneUsedFlag, "
        sgSQLQuery = sgSQLQuery & "dneVersion, "
        sgSQLQuery = sgSQLQuery & "dneOrigDneCode, "
        sgSQLQuery = sgSQLQuery & "dneCurrent, "
        sgSQLQuery = sgSQLQuery & "dneEnteredDate, "
        sgSQLQuery = sgSQLQuery & "dneEnteredTime, "
        sgSQLQuery = sgSQLQuery & "dneUieCode, "
        sgSQLQuery = sgSQLQuery & "dneUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlDNE.lCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDNE.sType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDNE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDNE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDNE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDNE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlDNE.iVersion & ", "
        If (llInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlDNE.lCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlDNE.lOrigDneCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDNE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlDNE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlDNE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlDNE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDNE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_DNE_DayName = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlDNE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_DNE_DayName = False
    Exit Function
End Function

Public Function gPutInsert_DSE_DaySubName(ilInsertType As Integer, tlDSE As DSE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigDseCode); 1=From Update (retain OrigDseCode)
'   tlDSE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    llLastCode = 0
    llInitCode = tlDSE.lCode
    Do
        If tlDSE.lCode <= 0 Then
            sgSQLQuery = "Select MAX(dseCode) from DSE_Day_SubName"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlDSE.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlDSE.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlDSE.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlDSE.lCode
        sgSQLQuery = "Insert Into DSE_Day_SubName ( "
        sgSQLQuery = sgSQLQuery & "dseCode, "
        sgSQLQuery = sgSQLQuery & "dseName, "
        sgSQLQuery = sgSQLQuery & "dseDescription, "
        sgSQLQuery = sgSQLQuery & "dseState, "
        sgSQLQuery = sgSQLQuery & "dseUsedFlag, "
        sgSQLQuery = sgSQLQuery & "dseVersion, "
        sgSQLQuery = sgSQLQuery & "dseOrigDseCode, "
        sgSQLQuery = sgSQLQuery & "dseCurrent, "
        sgSQLQuery = sgSQLQuery & "dseEnteredDate, "
        sgSQLQuery = sgSQLQuery & "dseEnteredTime, "
        sgSQLQuery = sgSQLQuery & "dseUieCode, "
        sgSQLQuery = sgSQLQuery & "dseUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlDSE.lCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDSE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDSE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDSE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDSE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlDSE.iVersion & ", "
        If (llInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlDSE.lCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlDSE.lOrigDseCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDSE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlDSE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlDSE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlDSE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDSE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_DSE_DaySubName = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlDSE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_DSE_DaySubName = False
    Exit Function
End Function

Public Function gPutInsert_EPE_EventProperties(tlEPE As EPE, slForm_Module As String) As Integer
'
'   tlEPTE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlEPE.iCode
    Do
        If tlEPE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(epeCode) from EPE_Event_Properties"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlEPE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlEPE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlEPE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlEPE.iCode
        sgSQLQuery = "Insert Into EPE_Event_Properties ( "
        sgSQLQuery = sgSQLQuery & "epeCode, "
        sgSQLQuery = sgSQLQuery & "epeEteCode, "
        sgSQLQuery = sgSQLQuery & "epeType, "
        sgSQLQuery = sgSQLQuery & "epeBus, "
        sgSQLQuery = sgSQLQuery & "epeBusControl, "
        sgSQLQuery = sgSQLQuery & "epeTime, "
        sgSQLQuery = sgSQLQuery & "epeStartType, "
        sgSQLQuery = sgSQLQuery & "epeFixedTime, "
        sgSQLQuery = sgSQLQuery & "epeEndType, "
        sgSQLQuery = sgSQLQuery & "epeDuration, "
        sgSQLQuery = sgSQLQuery & "epeMaterialType, "
        sgSQLQuery = sgSQLQuery & "epeAudioName, "
        sgSQLQuery = sgSQLQuery & "epeAudioItemID, "
        sgSQLQuery = sgSQLQuery & "epeAudioISCI, "
        sgSQLQuery = sgSQLQuery & "epeAudioControl, "
        sgSQLQuery = sgSQLQuery & "epeBkupAudioName, "
        sgSQLQuery = sgSQLQuery & "epeBkupAudioControl, "
        sgSQLQuery = sgSQLQuery & "epeProtAudioName, "
        sgSQLQuery = sgSQLQuery & "epeProtAudioItemID, "
        sgSQLQuery = sgSQLQuery & "epeProtAudioISCI, "
        sgSQLQuery = sgSQLQuery & "epeProtAudioControl, "
        sgSQLQuery = sgSQLQuery & "epeRelay1, "
        sgSQLQuery = sgSQLQuery & "epeRelay2, "
        sgSQLQuery = sgSQLQuery & "epeFollow, "
        sgSQLQuery = sgSQLQuery & "epeSilenceTime, "
        sgSQLQuery = sgSQLQuery & "epeSilence1, "
        sgSQLQuery = sgSQLQuery & "epeSilence2, "
        sgSQLQuery = sgSQLQuery & "epeSilence3, "
        sgSQLQuery = sgSQLQuery & "epeSilence4, "
        sgSQLQuery = sgSQLQuery & "epeStartNetcue, "
        sgSQLQuery = sgSQLQuery & "epeStopNetcue, "
        sgSQLQuery = sgSQLQuery & "epeTitle1, "
        sgSQLQuery = sgSQLQuery & "epeTitle2, "
        sgSQLQuery = sgSQLQuery & "epeABCFormat, "
        sgSQLQuery = sgSQLQuery & "epeABCPgmCode, "
        sgSQLQuery = sgSQLQuery & "epeABCXDSMode, "
        sgSQLQuery = sgSQLQuery & "epeABCRecordItem, "
        sgSQLQuery = sgSQLQuery & "epeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlEPE.iCode & ", "
        sgSQLQuery = sgSQLQuery & tlEPE.iEteCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sBus) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sBusControl) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sTime) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sStartType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sFixedTime) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sEndType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sDuration) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sMaterialType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sAudioName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sAudioItemID) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sAudioISCI) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sAudioControl) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sBkupAudioName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sBkupAudioControl) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sProtAudioName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sProtAudioItemID) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sProtAudioISCI) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sProtAudioControl) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sRelay1) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sRelay2) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sFollow) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sSilenceTime) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sSilence1) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sSilence2) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sSilence3) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sSilence4) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sStartNetcue) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sStopNetcue) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sTitle1) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sTitle2) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sABCFormat) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sABCPgmCode) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sABCXDSMode) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sABCRecordItem) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEPE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_EPE_EventProperties = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlEPE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_EPE_EventProperties = False
    Exit Function
End Function


Public Function gPutInsert_ETE_EventType(ilInsertType As Integer, tlETE As ETE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigEteCode); 1=From Update (retain OrigEteCode)
'   tlETE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlETE.iCode
    Do
        If tlETE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(eteCode) from ETE_Event_Type"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlETE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlETE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlETE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlETE.iCode
        sgSQLQuery = "Insert Into ETE_Event_Type ( "
        sgSQLQuery = sgSQLQuery & "eteCode, "
        sgSQLQuery = sgSQLQuery & "eteName, "
        sgSQLQuery = sgSQLQuery & "eteDescription, "
        sgSQLQuery = sgSQLQuery & "eteCategory, "
        sgSQLQuery = sgSQLQuery & "eteAutoCodeChar, "
        sgSQLQuery = sgSQLQuery & "eteState, "
        sgSQLQuery = sgSQLQuery & "eteUsedFlag, "
        sgSQLQuery = sgSQLQuery & "eteVersion, "
        sgSQLQuery = sgSQLQuery & "eteOrigEteCode, "
        sgSQLQuery = sgSQLQuery & "eteCurrent, "
        sgSQLQuery = sgSQLQuery & "eteEnteredDate, "
        sgSQLQuery = sgSQLQuery & "eteEnteredTime, "
        sgSQLQuery = sgSQLQuery & "eteUieCode, "
        sgSQLQuery = sgSQLQuery & "eteUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlETE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlETE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlETE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlETE.sCategory) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlETE.sAutoCodeChar) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlETE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlETE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlETE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlETE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlETE.iOrigEteCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlETE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlETE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlETE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlETE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlETE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_ETE_EventType = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlETE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_ETE_EventType = False
    Exit Function
End Function


Public Function gPutInsert_FNE_FollowName(ilInsertType As Integer, tlFNE As FNE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigFneCode); 1=From Update (retain OrigFneCode)
'   tlFNE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlFNE.iCode
    Do
        If tlFNE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(fneCode) from FNE_Follow_Name"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlFNE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlFNE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlFNE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlFNE.iCode
        sgSQLQuery = "Insert Into FNE_Follow_Name ( "
        sgSQLQuery = sgSQLQuery & "fneCode, "
        sgSQLQuery = sgSQLQuery & "fneName, "
        sgSQLQuery = sgSQLQuery & "fneDescription, "
        sgSQLQuery = sgSQLQuery & "fneState, "
        sgSQLQuery = sgSQLQuery & "fneUsedFlag, "
        sgSQLQuery = sgSQLQuery & "fneVersion, "
        sgSQLQuery = sgSQLQuery & "fneOrigFneCode, "
        sgSQLQuery = sgSQLQuery & "fneCurrent, "
        sgSQLQuery = sgSQLQuery & "fneEnteredDate, "
        sgSQLQuery = sgSQLQuery & "fneEnteredTime, "
        sgSQLQuery = sgSQLQuery & "fneUieCode, "
        sgSQLQuery = sgSQLQuery & "fneUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlFNE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlFNE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlFNE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlFNE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlFNE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlFNE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlFNE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlFNE.iOrigFneCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlFNE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlFNE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlFNE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlFNE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlFNE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_FNE_FollowName = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlFNE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_FNE_FollowName = False
    Exit Function
End Function


Public Function gPutInsert_ITE_ItemTest(tlITE As ITE, slForm_Module As String) As Integer
'
'   tlITE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlITE.iCode
    Do
        If tlITE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(iteCode) from ITE_Item_Test"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlITE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlITE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlITE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlITE.iCode
        sgSQLQuery = "INSERT INTO ITE_Item_Test ("
        sgSQLQuery = sgSQLQuery & "iteCode, "
        sgSQLQuery = sgSQLQuery & "iteSoeCode, "
        sgSQLQuery = sgSQLQuery & "iteType, "
        sgSQLQuery = sgSQLQuery & "iteName, "
        sgSQLQuery = sgSQLQuery & "iteDataBits, "
        sgSQLQuery = sgSQLQuery & "iteParity, "
        sgSQLQuery = sgSQLQuery & "iteStopBit, "
        sgSQLQuery = sgSQLQuery & "iteBaud, "
        sgSQLQuery = sgSQLQuery & "iteMachineID, "
        sgSQLQuery = sgSQLQuery & "iteStartCode, "
        sgSQLQuery = sgSQLQuery & "iteReplyCode, "
        sgSQLQuery = sgSQLQuery & "iteMinMgsID, "
        sgSQLQuery = sgSQLQuery & "iteMaxMgsID, "
        sgSQLQuery = sgSQLQuery & "iteCurrMgsID, "
        sgSQLQuery = sgSQLQuery & "iteMgsType, "
        sgSQLQuery = sgSQLQuery & "iteCheckSum, "
        sgSQLQuery = sgSQLQuery & "iteCmmdSeq, "
        sgSQLQuery = sgSQLQuery & "iteMgsEndCode, "
        sgSQLQuery = sgSQLQuery & "iteTitleID, "
        sgSQLQuery = sgSQLQuery & "iteLengthID, "
        sgSQLQuery = sgSQLQuery & "iteConnectSeq, "
        sgSQLQuery = sgSQLQuery & "iteMgsErrType, "
        sgSQLQuery = sgSQLQuery & "iteUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlITE.iCode & ", "
        sgSQLQuery = sgSQLQuery & tlITE.iSoeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sName) & "', "
        sgSQLQuery = sgSQLQuery & tlITE.iDataBits & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sParity) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sStopBit) & "', "
        sgSQLQuery = sgSQLQuery & tlITE.iBaud & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sMachineID) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sStartCode) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sReplyCode) & "', "
        sgSQLQuery = sgSQLQuery & tlITE.iMinMgsID & ", "
        sgSQLQuery = sgSQLQuery & tlITE.iMaxMgsID & ", "
        sgSQLQuery = sgSQLQuery & tlITE.iCurrMgsID & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sMgsType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sCheckSum) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sCmmdSeq) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sMgsEndCode) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sTitleID) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sLengthID) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sConnectSeq) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sMgsErrType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlITE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ")"
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_ITE_ItemTest = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlITE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_ITE_ItemTest = False
    Exit Function
End Function

Public Function gPutInsert_MTE_MaterialType(ilInsertType As Integer, tlMTE As MTE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigFneCode); 1=From Update (retain OrigFneCode)
'   tlMTE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlMTE.iCode
    Do
        If tlMTE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(mteCode) from MTE_Material_Type"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlMTE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlMTE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlMTE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlMTE.iCode
        sgSQLQuery = "Insert Into MTE_Material_Type ( "
        sgSQLQuery = sgSQLQuery & "mteCode, "
        sgSQLQuery = sgSQLQuery & "mteName, "
        sgSQLQuery = sgSQLQuery & "mteDescription, "
        sgSQLQuery = sgSQLQuery & "mteState, "
        sgSQLQuery = sgSQLQuery & "mteUsedFlag, "
        sgSQLQuery = sgSQLQuery & "mteVersion, "
        sgSQLQuery = sgSQLQuery & "mteOrigMteCode, "
        sgSQLQuery = sgSQLQuery & "mteCurrent, "
        sgSQLQuery = sgSQLQuery & "mteEnteredDate, "
        sgSQLQuery = sgSQLQuery & "mteEnteredTime, "
        sgSQLQuery = sgSQLQuery & "mteUieCode, "
        sgSQLQuery = sgSQLQuery & "mteUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlMTE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlMTE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlMTE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlMTE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlMTE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlMTE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlMTE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlMTE.iOrigMteCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlMTE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlMTE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlMTE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlMTE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlMTE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_MTE_MaterialType = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlMTE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_MTE_MaterialType = False
    Exit Function
End Function

Public Function gPutInsert_NNE_NetcueName(ilInsertType As Integer, tlNNE As NNE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigNneCode); 1=From Update (retain OrigNneCode)
'   tlNNE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlNNE.iCode
    Do
        If tlNNE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(nneCode) from NNE_Netcue_Name"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlNNE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlNNE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlNNE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlNNE.iCode
        sgSQLQuery = "Insert Into NNE_Netcue_Name ( "
        sgSQLQuery = sgSQLQuery & "nneCode, "
        sgSQLQuery = sgSQLQuery & "nneName, "
        sgSQLQuery = sgSQLQuery & "nneDescription, "
        sgSQLQuery = sgSQLQuery & "nneDneCode, "
        sgSQLQuery = sgSQLQuery & "nneState, "
        sgSQLQuery = sgSQLQuery & "nneUsedFlag, "
        sgSQLQuery = sgSQLQuery & "nneVersion, "
        sgSQLQuery = sgSQLQuery & "nneOrigNneCode, "
        sgSQLQuery = sgSQLQuery & "nneCurrent, "
        sgSQLQuery = sgSQLQuery & "nneEnteredDate, "
        sgSQLQuery = sgSQLQuery & "nneEnteredTime, "
        sgSQLQuery = sgSQLQuery & "nneUieCode, "
        sgSQLQuery = sgSQLQuery & "nneUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlNNE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlNNE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlNNE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & tlNNE.lDneCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlNNE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlNNE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlNNE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlNNE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlNNE.iOrigNneCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlNNE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlNNE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlNNE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlNNE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlNNE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_NNE_NetcueName = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlNNE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_NNE_NetcueName = False
    Exit Function
End Function

Public Function gPutInsert_RLE_Record_Locks(tlRLE As RLE, slForm_Module As String, hlRLE As Integer) As Integer
'
'   ilInsertType(I) = 0 New (set OrigRneCode); 1=From Update (retain OrigRneCode)
'   tlRLE(I)- Record to be added to Database
'
    Dim ilRet As Integer
'    Dim llLastCode As Long
'    Dim llInitCode As Long
    Dim tlRLEAPI As RLEAPI
    
    On Error GoTo ErrHand
'    llLastCode = 0
'    llInitCode = tlRLE.iCode
'    Do
'        If tlRLE.iCode <= 0 Then
'            sgSQLQuery = "Select MAX(rleCode) from RLE_Record_Locks"
'            Set rst = cnn.Execute(sgSQLQuery)
'            If IsNull(rst(0).Value) Then
'                tlRLE.iCode = 1
'            Else
'                If rst(0).Value > 0 Then
'                    tlRLE.iCode = rst(0).Value + 1
'                End If
'            End If
'            If llLastCode = tlRLE.iCode Then
'                GoTo ErrHand1:
'            End If
'        End If
'        llLastCode = tlRLE.iCode
'        sgSQLQuery = "Insert Into RLE_Record_Locks ( "
'        sgSQLQuery = sgSQLQuery & "rleCode, "
'        sgSQLQuery = sgSQLQuery & "rleUieCode, "
'        sgSQLQuery = sgSQLQuery & "rleFileName, "
'        sgSQLQuery = sgSQLQuery & "rleRecCode, "
'        sgSQLQuery = sgSQLQuery & "rleEnteredDate, "
'        sgSQLQuery = sgSQLQuery & "rleEnteredTime, "
'        sgSQLQuery = sgSQLQuery & "rleUnused "
'        sgSQLQuery = sgSQLQuery & ") "
'        sgSQLQuery = sgSQLQuery & "Values ( "
'        sgSQLQuery = sgSQLQuery & tlRLE.lCode & ", "
'        sgSQLQuery = sgSQLQuery & tlRLE.iUieCode & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlRLE.sFileName) & "', "
'        sgSQLQuery = sgSQLQuery & tlRLE.lRecCode & ", "
'        sgSQLQuery = sgSQLQuery & "'" & Format$(tlRLE.sEnteredDate, sgSQLDateForm) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & Format$(tlRLE.sEnteredTime, sgSQLTimeForm) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlRLE.sUnused) & "' "
'        sgSQLQuery = sgSQLQuery & ") "
'        llRet = 0
'        cnn.Execute sgSQLQuery    ', rdExecDirect
'    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    Dim ilRLERecLen As Integer
    ilRLERecLen = Len(tlRLEAPI)
    tlRLEAPI.lCode = tlRLE.lCode
    tlRLEAPI.iUieCode = tlRLE.iUieCode
    tlRLEAPI.sFileName = tlRLE.sFileName
    tlRLEAPI.lRecCode = tlRLE.lRecCode
    gPackDate gAdjYear(tlRLE.sEnteredDate), tlRLEAPI.iEnteredDate(0), tlRLEAPI.iEnteredDate(1)
    gPackTime tlRLE.sEnteredTime, tlRLEAPI.iEnteredTime(0), tlRLEAPI.iEnteredTime(1)
    tlRLEAPI.sUnused = tlRLE.sUnused
    ilRet = btrInsert(hlRLE, tlRLEAPI, ilRLERecLen, 0)
    If ilRet <> BTRV_ERR_NONE Then
        gPutInsert_RLE_Record_Locks = False
        Exit Function
    End If
    tlRLE.lCode = tlRLEAPI.lCode
    gPutInsert_RLE_Record_Locks = True
    Exit Function
ErrHand:
'    If (llInitCode = 0) Then
'        For Each gErrSQL In cnn.Errors
'            llRet = gErrSQL.NativeError
'            If llRet < 0 Then
'                llRet = llRet + 4999
'            End If
'            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
'                tlRLE.iCode = 0
'                Resume Next
'            End If
'        Next gErrSQL
'    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_RLE_Record_Locks = False
    Exit Function
End Function

Public Function gPutInsert_RNE_RelayName(ilInsertType As Integer, tlRNE As RNE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigRneCode); 1=From Update (retain OrigRneCode)
'   tlRNE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlRNE.iCode
    Do
        If tlRNE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(rneCode) from RNE_Relay_Name"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlRNE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlRNE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlRNE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlRNE.iCode
        sgSQLQuery = "Insert Into RNE_Relay_Name ( "
        sgSQLQuery = sgSQLQuery & "rneCode, "
        sgSQLQuery = sgSQLQuery & "rneName, "
        sgSQLQuery = sgSQLQuery & "rneDescription, "
        sgSQLQuery = sgSQLQuery & "rneState, "
        sgSQLQuery = sgSQLQuery & "rneUsedFlag, "
        sgSQLQuery = sgSQLQuery & "rneVersion, "
        sgSQLQuery = sgSQLQuery & "rneOrigRneCode, "
        sgSQLQuery = sgSQLQuery & "rneCurrent, "
        sgSQLQuery = sgSQLQuery & "rneEnteredDate, "
        sgSQLQuery = sgSQLQuery & "rneEnteredTime, "
        sgSQLQuery = sgSQLQuery & "rneUieCode, "
        sgSQLQuery = sgSQLQuery & "rneUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlRNE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlRNE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlRNE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlRNE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlRNE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlRNE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlRNE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlRNE.iOrigRneCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlRNE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlRNE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlRNE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlRNE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlRNE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_RNE_RelayName = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlRNE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_RNE_RelayName = False
    Exit Function
End Function

Public Function gPutInsert_SCE_SilenceChar(ilInsertType As Integer, tlSCE As SCE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigSceCode); 1=From Update (retain OrigSceCode)
'   tlSCE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlSCE.iCode
    Do
        If tlSCE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(sceCode) from SCE_Silence_Char"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlSCE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlSCE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlSCE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlSCE.iCode
        sgSQLQuery = "Insert Into SCE_Silence_Char ( "
        sgSQLQuery = sgSQLQuery & "sceCode, "
        sgSQLQuery = sgSQLQuery & "sceAutoChar, "
        sgSQLQuery = sgSQLQuery & "sceDescription, "
        sgSQLQuery = sgSQLQuery & "sceState, "
        sgSQLQuery = sgSQLQuery & "sceUsedFlag, "
        sgSQLQuery = sgSQLQuery & "sceVersion, "
        sgSQLQuery = sgSQLQuery & "sceOrigSceCode, "
        sgSQLQuery = sgSQLQuery & "sceCurrent, "
        sgSQLQuery = sgSQLQuery & "sceEnteredDate, "
        sgSQLQuery = sgSQLQuery & "sceEnteredTime, "
        sgSQLQuery = sgSQLQuery & "sceUieCode, "
        sgSQLQuery = sgSQLQuery & "sceUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlSCE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSCE.sAutoChar) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSCE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSCE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSCE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlSCE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlSCE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlSCE.iOrigSceCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSCE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSCE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSCE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlSCE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSCE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_SCE_SilenceChar = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlSCE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_SCE_SilenceChar = False
    Exit Function
End Function

Public Function gPutInsert_SEE_ScheduleEvents(tlSEE As SEE, slForm_Module As String, hlSEE As Integer, hlSOE As Integer) As Integer
'
'   tlSEE(I)- Record to be added to Database
'
    Dim ilRet As Integer
    Dim llLastCode As Long
    Dim llInitCode As Long
    
    On Error GoTo ErrHand
'    llLastCode = 0
'    llInitCode = tlSEE.lCode
'    Do
'        If tlSEE.lCode <= 0 Then
'            sgSQLQuery = "Select MAX(seeCode) from SEE_Schedule_Events"
'            Set rst = cnn.Execute(sgSQLQuery)
'            If IsNull(rst(0).Value) Then
'                tlSEE.lCode = 1
'            Else
'                If rst(0).Value > 0 Then
'                    tlSEE.lCode = rst(0).Value + 1
'                End If
'            End If
'            If llLastCode = tlSEE.lCode Then
'                GoTo ErrHand1:
'            End If
'        End If
'        Do
'            ilRet = gGetRec_SOE_SiteOption(tgSOE.iCode, "gPutInsert_SEE: Get Event ID from SOE", tgSOE)
'            tlSEE.lEventID = tgSOE.lCurrEventID + 1
'            If (tlSEE.lEventID < tgSOE.lMinEventID) Or (tlSEE.lEventID > tgSOE.lMaxEventID) Then
'                tlSEE.lEventID = tgSOE.lMinEventID
'            End If
'            tgSOE.lCurrEventID = tlSEE.lEventID
'            ilRet = gPutUpdate_SOE_SiteOption(1, tgSOE, "gPutInsert_See: Update EventID in SOE")
'            sgSQLQuery = "Select seeCode from SEE_Schedule_Events WHERE seeEventID = " & tlSEE.lEventID
'            Set rst = cnn.Execute(sgSQLQuery)
'        Loop While Not rst.EOF
'        Set rst = cnn.Execute(sgSQLQuery)
'        sgSQLQuery = "Insert Into SEE_Schedule_Events ( "
'        sgSQLQuery = sgSQLQuery & "seeCode, "
'        sgSQLQuery = sgSQLQuery & "seeSheCode, "
'        sgSQLQuery = sgSQLQuery & "seeAction, "
'        sgSQLQuery = sgSQLQuery & "seeDeeCode, "
'        sgSQLQuery = sgSQLQuery & "seeBdeCode, "
'        sgSQLQuery = sgSQLQuery & "seeBusCceCode, "
'        sgSQLQuery = sgSQLQuery & "seeSchdType, "
'        sgSQLQuery = sgSQLQuery & "seeEteCode, "
'        sgSQLQuery = sgSQLQuery & "seeTime, "
'        sgSQLQuery = sgSQLQuery & "seeStartTteCode, "
'        sgSQLQuery = sgSQLQuery & "seeFixedTime, "
'        sgSQLQuery = sgSQLQuery & "seeEndTteCode, "
'        sgSQLQuery = sgSQLQuery & "seeDuration, "
'        sgSQLQuery = sgSQLQuery & "seeMteCode, "
'        sgSQLQuery = sgSQLQuery & "seeAudioAseCode, "
'        sgSQLQuery = sgSQLQuery & "seeAudioItemID, "
'        sgSQLQuery = sgSQLQuery & "seeAudioItemIDChk, "
'        sgSQLQuery = sgSQLQuery & "seeAudioCceCode, "
'        sgSQLQuery = sgSQLQuery & "seeBkupAneCode, "
'        sgSQLQuery = sgSQLQuery & "seeBkupCceCode, "
'        sgSQLQuery = sgSQLQuery & "seeProtAneCode, "
'        sgSQLQuery = sgSQLQuery & "seeProtItemID, "
'        sgSQLQuery = sgSQLQuery & "seeProtItemIDChk, "
'        sgSQLQuery = sgSQLQuery & "seeProtCceCode, "
'        sgSQLQuery = sgSQLQuery & "see1RneCode, "
'        sgSQLQuery = sgSQLQuery & "see2RneCode, "
'        sgSQLQuery = sgSQLQuery & "seeFneCode, "
'        sgSQLQuery = sgSQLQuery & "seeSilenceTime, "
'        sgSQLQuery = sgSQLQuery & "see1SceCode, "
'        sgSQLQuery = sgSQLQuery & "see2SceCode, "
'        sgSQLQuery = sgSQLQuery & "see3SceCode, "
'        sgSQLQuery = sgSQLQuery & "see4SceCode, "
'        sgSQLQuery = sgSQLQuery & "seeStartNneCode, "
'        sgSQLQuery = sgSQLQuery & "seeEndNneCode, "
'        sgSQLQuery = sgSQLQuery & "see1CteCode, "
'        sgSQLQuery = sgSQLQuery & "see2CteCode, "
'        sgSQLQuery = sgSQLQuery & "seeAreCode, "
'        sgSQLQuery = sgSQLQuery & "seeSpotTime, "
'        sgSQLQuery = sgSQLQuery & "seeEventID, "
'        sgSQLQuery = sgSQLQuery & "seeAsAirStatus, "
'        sgSQLQuery = sgSQLQuery & "seeSentStatus, "
'        sgSQLQuery = sgSQLQuery & "seeSentDate, "
'        sgSQLQuery = sgSQLQuery & "seeIgnoreConflicts, "
'        sgSQLQuery = sgSQLQuery & "seeDHECode, "
'        sgSQLQuery = sgSQLQuery & "seeOrigDHECode, "
'        sgSQLQuery = sgSQLQuery & "seeInsertFlag, "
'        sgSQLQuery = sgSQLQuery & "seeUnused "
'        sgSQLQuery = sgSQLQuery & ") "
'        sgSQLQuery = sgSQLQuery & "Values ( "
'        sgSQLQuery = sgSQLQuery & tlSEE.lCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.lSheCode & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSEE.sAction) & "', "
'        sgSQLQuery = sgSQLQuery & tlSEE.lDeeCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.iBDECode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.iBusCceCode & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSEE.sSchdType) & "', "
'        sgSQLQuery = sgSQLQuery & tlSEE.iEteCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.lTime & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.iStartTteCode & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSEE.sFixedTime) & "', "
'        sgSQLQuery = sgSQLQuery & tlSEE.iEndTteCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.lDuration & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.iMteCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.iAudioAseCode & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSEE.sAudioItemID) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSEE.sAudioItemIDChk) & "', "
'        sgSQLQuery = sgSQLQuery & tlSEE.iAudioCceCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.iBkupAneCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.iBkupCceCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.iProtAneCode & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSEE.sProtItemID) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSEE.sProtItemIDChk) & "', "
'        sgSQLQuery = sgSQLQuery & tlSEE.iProtCceCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.i1RneCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.i2RneCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.iFneCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.lSilenceTime & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.i1SceCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.i2SceCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.i3SceCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.i4SceCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.iStartNneCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.iEndNneCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.l1CteCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.l2CteCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.lAreCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.lSpotTime & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.lEventID & ", "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSEE.sAsAirStatus) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSEE.sSentStatus) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSEE.sSentDate, sgSQLDateForm) & "', "
'        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSEE.sIgnoreConflicts) & "', "
'        sgSQLQuery = sgSQLQuery & tlSEE.lDheCode & ", "
'        sgSQLQuery = sgSQLQuery & tlSEE.lOrigDHECode & ", "
'        sgSQLQuery = sgSQLQuery & "'', "        'Insert Flag field, only temporarily used in Schedule definition
'        sgSQLQuery = sgSQLQuery & "'' "
'        sgSQLQuery = sgSQLQuery & ") "
'        ilRet = 0
'        cnn.Execute sgSQLQuery    ', rdExecDirect
'    Loop While ilRet = BTRV_ERR_DUPLICATE_KEY
    Dim ilSEERecLen As Integer
    Dim ilSOERecLen As Integer
    Dim tlSOESrchKey As INTKEY0
    ilSEERecLen = Len(tlSEE)
    ilSOERecLen = Len(tgSOE)
    Do
        tlSOESrchKey.iCode = tgSOE.iCode
        ilRet = btrGetEqual(hlSOE, tgSOE, ilSOERecLen, tlSOESrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        tlSEE.lEventID = tgSOE.lCurrEventID + 1
        If (tlSEE.lEventID < tgSOE.lMinEventID) Or (tlSEE.lEventID > tgSOE.lMaxEventID) Then
            tlSEE.lEventID = tgSOE.lMinEventID
        End If
        tgSOE.lCurrEventID = tlSEE.lEventID
        ilRet = btrUpdate(hlSOE, tgSOE, ilSOERecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    tlSEE.sSentDate = Format$(gAdjYear(tlSEE.sSentDate), sgSQLDateForm)
    ilRet = btrInsert(hlSEE, tlSEE, ilSEERecLen, 0)
    If ilRet <> BTRV_ERR_NONE Then
        gPutInsert_SEE_ScheduleEvents = False
        Exit Function
    End If
    gPutInsert_SEE_ScheduleEvents = True
    Exit Function
ErrHand:
'    If (llInitCode = 0) Then
'        For Each gErrSQL In cnn.Errors
'            ilRet = gErrSQL.NativeError
'            If ilRet < 0 Then
'                ilRet = ilRet + 4999
'            End If
'            If (ilRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
'                tlSEE.lCode = 0
'                Resume Next
'            End If
'        Next gErrSQL
'    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_SEE_ScheduleEvents = False
    Exit Function
End Function

Public Function gPutInsert_SGE_SiteGenSchd(tlSGE As SGE, slForm_Module As String) As Integer
'
'   tlSGE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlSGE.iCode
    Do
        If tlSGE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(sgeCode) from SGE_Site_Gen_Schd"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlSGE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlSGE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlSGE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlSGE.iCode
        sgSQLQuery = "Insert Into SGE_Site_Gen_Schd ( "
        sgSQLQuery = sgSQLQuery & "sgeCode, "
        sgSQLQuery = sgSQLQuery & "sgeSoeCode, "
        sgSQLQuery = sgSQLQuery & "sgeType, "
        sgSQLQuery = sgSQLQuery & "sgeSubType, "
        sgSQLQuery = sgSQLQuery & "sgeGenMo, "
        sgSQLQuery = sgSQLQuery & "sgeGenTu, "
        sgSQLQuery = sgSQLQuery & "sgeGenWe, "
        sgSQLQuery = sgSQLQuery & "sgeGenTh, "
        sgSQLQuery = sgSQLQuery & "sgeGenFr, "
        sgSQLQuery = sgSQLQuery & "sgeGenSa, "
        sgSQLQuery = sgSQLQuery & "sgeGenSu, "
        sgSQLQuery = sgSQLQuery & "sgeGenTime, "
        sgSQLQuery = sgSQLQuery & "sgePurgeAfterGen, "
        sgSQLQuery = sgSQLQuery & "sgePurgeTime, "
        sgSQLQuery = sgSQLQuery & "sgeAlertInterval, "
        sgSQLQuery = sgSQLQuery & "sgeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlSGE.iCode & ", "
        sgSQLQuery = sgSQLQuery & tlSGE.iSoeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSGE.sType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSGE.sSubType) & "', "
        sgSQLQuery = sgSQLQuery & tlSGE.iGenMo & ", "
        sgSQLQuery = sgSQLQuery & tlSGE.iGenTu & ", "
        sgSQLQuery = sgSQLQuery & tlSGE.iGenWe & ", "
        sgSQLQuery = sgSQLQuery & tlSGE.iGenTh & ", "
        sgSQLQuery = sgSQLQuery & tlSGE.iGenFr & ", "
        sgSQLQuery = sgSQLQuery & tlSGE.iGenSa & ", "
        sgSQLQuery = sgSQLQuery & tlSGE.iGenSu & ", "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSGE.sGenTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSGE.sPurgeAfterGen) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSGE.sPurgeTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlSGE.lAlertInterval & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSGE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_SGE_SiteGenSchd = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlSGE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_SGE_SiteGenSchd = False
    Exit Function
End Function

Public Function gPutInsert_SHE_ScheduleHeader(ilInsertType As Integer, tlSHE As SHE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigSheCode); 1=From Update (retain OrigSheCode)
'   tlSHE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    llLastCode = 0
    llInitCode = tlSHE.lCode
    Do
        If tlSHE.lCode <= 0 Then
            sgSQLQuery = "Select MAX(sheCode) from SHE_Schedule_Header"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlSHE.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlSHE.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlSHE.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlSHE.lCode
        sgSQLQuery = "Insert Into SHE_Schedule_Header ( "
        sgSQLQuery = sgSQLQuery & "sheCode, "
        sgSQLQuery = sgSQLQuery & "sheAeeCode, "
        sgSQLQuery = sgSQLQuery & "sheAirDate, "
        sgSQLQuery = sgSQLQuery & "sheLoadedAutoStatus, "
        sgSQLQuery = sgSQLQuery & "sheLoadedAutoDate, "
        sgSQLQuery = sgSQLQuery & "sheChgSeqNo, "
        sgSQLQuery = sgSQLQuery & "sheAsAirStatus, "
        sgSQLQuery = sgSQLQuery & "sheLoadedAsAirDate, "
        sgSQLQuery = sgSQLQuery & "sheLastDateItemChk, "
        sgSQLQuery = sgSQLQuery & "sheCreateLoad, "
        sgSQLQuery = sgSQLQuery & "sheVersion, "
        sgSQLQuery = sgSQLQuery & "sheOrigSheCode, "
        sgSQLQuery = sgSQLQuery & "sheCurrent, "
        sgSQLQuery = sgSQLQuery & "sheEnteredDate, "
        sgSQLQuery = sgSQLQuery & "sheEnteredTime, "
        sgSQLQuery = sgSQLQuery & "sheUieCode, "
        sgSQLQuery = sgSQLQuery & "sheConflictExist, "
        sgSQLQuery = sgSQLQuery & "sheSpotMergeStatus, "
        sgSQLQuery = sgSQLQuery & "sheLoadStatus, "
        sgSQLQuery = sgSQLQuery & "sheUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlSHE.lCode & ", "
        sgSQLQuery = sgSQLQuery & tlSHE.iAeeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSHE.sAirDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSHE.sLoadedAutoStatus) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSHE.sLoadedAutoDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & tlSHE.iChgSeqNo & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSHE.sAsAirStatus) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSHE.sLoadedAsAirDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSHE.sLastDateItemChk, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSHE.sCreateLoad) & "', "
        sgSQLQuery = sgSQLQuery & tlSHE.iVersion & ", "
        If (llInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlSHE.lCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlSHE.lOrigSheCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSHE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSHE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSHE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlSHE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSHE.sConflictExist) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSHE.sSpotMergeStatus) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSHE.sLoadStatus) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSHE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_SHE_ScheduleHeader = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlSHE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_SHE_ScheduleHeader = False
    Exit Function
End Function

Public Function gPutInsert_SPE_SitePath(tlSPE As SPE, slForm_Module As String) As Integer
'
'   tlSPE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlSPE.iCode
    Do
        If tlSPE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(speCode) from SPE_Site_Path"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlSPE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlSPE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlSPE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlSPE.iCode
        sgSQLQuery = "Insert Into SPE_Site_Path ( "
        sgSQLQuery = sgSQLQuery & "speCode, "
        sgSQLQuery = sgSQLQuery & "speSoeCode, "
        sgSQLQuery = sgSQLQuery & "speType, "
        sgSQLQuery = sgSQLQuery & "speSubType, "
        sgSQLQuery = sgSQLQuery & "spePath, "
        sgSQLQuery = sgSQLQuery & "speUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlSPE.iCode & ", "
        sgSQLQuery = sgSQLQuery & tlSPE.iSoeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSPE.sType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSPE.sSubType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSPE.sPath) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSPE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_SPE_SitePath = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlSPE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_SPE_SitePath = False
    Exit Function
End Function

Public Function gPutInsert_SOE_SiteOption(ilInsertType As Integer, tlSOE As SOE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigSoeCode); 1=From Update (retain OrigSoeCode)
'   tlSOE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlSOE.iCode
    Do
        If tlSOE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(SoeCode) from SOE_Site_Option"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlSOE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlSOE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlSOE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlSOE.iCode
        sgSQLQuery = "INSERT INTO SOE_Site_Option ("
        sgSQLQuery = sgSQLQuery & "soeCode, "
        sgSQLQuery = sgSQLQuery & "soeClientName, "
        sgSQLQuery = sgSQLQuery & "soeAddr1, "
        sgSQLQuery = sgSQLQuery & "soeAddr2, "
        sgSQLQuery = sgSQLQuery & "soeAddr3, "
        sgSQLQuery = sgSQLQuery & "soePhone, "
        sgSQLQuery = sgSQLQuery & "soeFax, "
        sgSQLQuery = sgSQLQuery & "soeDaysRetainAsAir, "
        sgSQLQuery = sgSQLQuery & "soeDaysRetainActive, "
        sgSQLQuery = sgSQLQuery & "soeChgInterval, "
        sgSQLQuery = sgSQLQuery & "soeMergeDateFormat, "
        sgSQLQuery = sgSQLQuery & "soeMergeTimeFormat, "
        sgSQLQuery = sgSQLQuery & "soeMergeFileFormat, "
        sgSQLQuery = sgSQLQuery & "soeMergeFileExt, "
        sgSQLQuery = sgSQLQuery & "soeMergeStartTime, "
        sgSQLQuery = sgSQLQuery & "soeMergeEndTime, "
        sgSQLQuery = sgSQLQuery & "soeMergeChkInterval, "
        sgSQLQuery = sgSQLQuery & "soeMergeStopFlag, "
        sgSQLQuery = sgSQLQuery & "soeAlertInterval, "
        sgSQLQuery = sgSQLQuery & "soeSchAutoGenSeq, "
        sgSQLQuery = sgSQLQuery & "soeMinEventID, "
        sgSQLQuery = sgSQLQuery & "soeMaxEventID, "
        sgSQLQuery = sgSQLQuery & "soeCurrEventID, "
        sgSQLQuery = sgSQLQuery & "soeNoDaysRetainPW, "
        sgSQLQuery = sgSQLQuery & "soeVersion, "
        sgSQLQuery = sgSQLQuery & "soeOrigSOECode, "
        sgSQLQuery = sgSQLQuery & "soeCurrent, "
        sgSQLQuery = sgSQLQuery & "soeEnteredDate, "
        sgSQLQuery = sgSQLQuery & "soeEnteredTime, "
        sgSQLQuery = sgSQLQuery & "soeUieCode, "
        sgSQLQuery = sgSQLQuery & "soeSpotItemIDWindow, "
        sgSQLQuery = sgSQLQuery & "soeTimeTolerance, "
        sgSQLQuery = sgSQLQuery & "soeLengthTolerance, "
        sgSQLQuery = sgSQLQuery & "soeMatchATNotB, "
        sgSQLQuery = sgSQLQuery & "soeMatchATBNotI, "
        sgSQLQuery = sgSQLQuery & "soeMatchANotT, "
        sgSQLQuery = sgSQLQuery & "soeMatchBNotT, "
        sgSQLQuery = sgSQLQuery & "soeSchAutoGenSeqTst, "
        sgSQLQuery = sgSQLQuery & "soeMergeStopFlagTst, "
        sgSQLQuery = sgSQLQuery & "soeUnused )"
        sgSQLQuery = sgSQLQuery & "VALUES ("
        sgSQLQuery = sgSQLQuery & tlSOE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sClientName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sAddr1) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sAddr2) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sAddr3) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sPhone) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sFax) & "', "
        sgSQLQuery = sgSQLQuery & tlSOE.iDaysRetainAsAir & ", "
        sgSQLQuery = sgSQLQuery & tlSOE.iDaysRetainActive & ", "
        sgSQLQuery = sgSQLQuery & tlSOE.lChgInterval & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sMergeDateFormat) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sMergeTimeFormat) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sMergeFileFormat) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sMergeFileExt) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSOE.sMergeStartTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSOE.sMergeEndTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlSOE.iMergeChkInterval & ", "
        sgSQLQuery = sgSQLQuery & "'" & tlSOE.sMergeStopFlag & "', "
        sgSQLQuery = sgSQLQuery & tlSOE.iAlertInterval & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sSchAutoGenSeq) & "', "
        sgSQLQuery = sgSQLQuery & tlSOE.lMinEventID & ", "
        sgSQLQuery = sgSQLQuery & tlSOE.lMaxEventID & ", "
        sgSQLQuery = sgSQLQuery & tlSOE.lCurrEventID & ", "
        sgSQLQuery = sgSQLQuery & tlSOE.iNoDaysRetainPW & ", "
        sgSQLQuery = sgSQLQuery & tlSOE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlSOE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlSOE.iOrigSoeCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSOE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlSOE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlSOE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & tlSOE.iSpotItemIDWindow & ", "
        sgSQLQuery = sgSQLQuery & tlSOE.lTimeTolerance & ", "
        sgSQLQuery = sgSQLQuery & tlSOE.lLengthTolerance & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sMatchATNotB) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sMatchATBNotI) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sMatchANotT) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sMatchBNotT) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sSchAutoGenSeqTst) & "', "
        sgSQLQuery = sgSQLQuery & "'" & tlSOE.sMergeStopFlagTst & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSOE.sUnused) & "'"
        sgSQLQuery = sgSQLQuery & ")"
        cnn.BeginTrans
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    cnn.CommitTrans
    gPutInsert_SOE_SiteOption = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlSOE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_SOE_SiteOption = False
    Exit Function
End Function

Public Function gPutInsert_SSE_Site_SMTP_Info(tlSSE As SSE, slForm_Module As String) As Integer
'
'   tlSSE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlSSE.iCode
    Do
        If tlSSE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(sseCode) from SSE_Site_SMTP_Info"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlSSE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlSSE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlSSE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlSSE.iCode
        sgSQLQuery = "Insert Into SSE_Site_SMTP_Info ( "
        sgSQLQuery = sgSQLQuery & "sseCode, "
        sgSQLQuery = sgSQLQuery & "sseSoeCode, "
        sgSQLQuery = sgSQLQuery & "sseEMailHost, "
        sgSQLQuery = sgSQLQuery & "sseEMailPort, "
        sgSQLQuery = sgSQLQuery & "sseEMailAcctName, "
        sgSQLQuery = sgSQLQuery & "sseEMailPassword, "
        sgSQLQuery = sgSQLQuery & "sseEMailTLS, "
        sgSQLQuery = sgSQLQuery & "sseUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlSSE.iCode & ", "
        sgSQLQuery = sgSQLQuery & tlSSE.iSoeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSSE.sEMailHost) & "', "
        sgSQLQuery = sgSQLQuery & tlSSE.iEMailPort & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSSE.sEMailAcctName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSSE.sEMailPassword) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSSE.sEMailTLS) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlSSE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_SSE_Site_SMTP_Info = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlSSE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_SSE_Site_SMTP_Info = False
    Exit Function
End Function


Public Function gPutInsert_TNE_TaskName(tlTne As TNE, slForm_Module As String) As Integer
'
'   tlTNE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlTne.iCode
    Do
        If tlTne.iCode <= 0 Then
            sgSQLQuery = "Select MAX(tneCode) from TNE_Task_Name"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlTne.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlTne.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlTne.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlTne.iCode
        sgSQLQuery = "INSERT INTO TNE_Task_Name ("
        sgSQLQuery = sgSQLQuery & "tneCode, "
        sgSQLQuery = sgSQLQuery & "tneType, "
        sgSQLQuery = sgSQLQuery & "tneName, "
        sgSQLQuery = sgSQLQuery & "tneUnused )"
        sgSQLQuery = sgSQLQuery & "VALUES ("
        sgSQLQuery = sgSQLQuery & tlTne.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & tlTne.sType & "', "
        sgSQLQuery = sgSQLQuery & "'" & tlTne.sName & "', "
        sgSQLQuery = sgSQLQuery & "'" & tlTne.sUnused & "'"
        sgSQLQuery = sgSQLQuery & ")"
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_TNE_TaskName = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlTne.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_TNE_TaskName = False
    Exit Function
End Function

Public Function gPutInsert_TSE_TemplateSchd(ilInsertType As Integer, tlTSE As TSE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigRneCode); 1=From Update (retain OrigRneCode)
'   tlTSE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    llLastCode = 0
    llInitCode = tlTSE.lCode
    Do
        If tlTSE.lCode <= 0 Then
            sgSQLQuery = "Select MAX(tseCode) from TSE_Template_Schd"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlTSE.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlTSE.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlTSE.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlTSE.lCode
        sgSQLQuery = "Insert Into TSE_Template_Schd ( "
        sgSQLQuery = sgSQLQuery & "tseCode, "
        sgSQLQuery = sgSQLQuery & "tseDheCode, "
        sgSQLQuery = sgSQLQuery & "tseBdeCode, "
        sgSQLQuery = sgSQLQuery & "tseLogDate, "
        sgSQLQuery = sgSQLQuery & "tseStartTime, "
        sgSQLQuery = sgSQLQuery & "tseDescription, "
        sgSQLQuery = sgSQLQuery & "tseState, "
        sgSQLQuery = sgSQLQuery & "tseCteCode, "
        sgSQLQuery = sgSQLQuery & "tseVersion, "
        sgSQLQuery = sgSQLQuery & "tseOrigTseCode, "
        sgSQLQuery = sgSQLQuery & "tseCurrent, "
        sgSQLQuery = sgSQLQuery & "tseEnteredDate, "
        sgSQLQuery = sgSQLQuery & "tseEnteredTime, "
        sgSQLQuery = sgSQLQuery & "tseUieCode, "
        sgSQLQuery = sgSQLQuery & "tseUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlTSE.lCode & ", "
        sgSQLQuery = sgSQLQuery & tlTSE.lDheCode & ", "
        sgSQLQuery = sgSQLQuery & tlTSE.iBdeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlTSE.sLogDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlTSE.sStartTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlTSE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlTSE.sState) & "', "
        sgSQLQuery = sgSQLQuery & tlTSE.lCteCode & ", "
        sgSQLQuery = sgSQLQuery & tlTSE.iVersion & ", "
        If (llInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlTSE.lCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlTSE.lOrigTseCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlTSE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlTSE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlTSE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlTSE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlTSE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_TSE_TemplateSchd = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlTSE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_TSE_TemplateSchd = False
    Exit Function
End Function

Public Function gPutInsert_TTE_TimeType(ilInsertType As Integer, tlTTE As TTE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigTteCode); 1=From Update (retain OrigTteCode)
'   tlTTE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlTTE.iCode
    Do
        If tlTTE.iCode <= 0 Then
            sgSQLQuery = "Select MAX(tteCode) from TTE_Time_Type"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlTTE.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlTTE.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlTTE.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlTTE.iCode
        sgSQLQuery = "Insert Into TTE_Time_Type ( "
        sgSQLQuery = sgSQLQuery & "tteCode, "
        sgSQLQuery = sgSQLQuery & "tteType, "
        sgSQLQuery = sgSQLQuery & "tteName, "
        sgSQLQuery = sgSQLQuery & "tteDescription, "
        sgSQLQuery = sgSQLQuery & "tteState, "
        sgSQLQuery = sgSQLQuery & "tteUsedFlag, "
        sgSQLQuery = sgSQLQuery & "tteVersion, "
        sgSQLQuery = sgSQLQuery & "tteOrigTteCode, "
        sgSQLQuery = sgSQLQuery & "tteCurrent, "
        sgSQLQuery = sgSQLQuery & "tteEnteredDate, "
        sgSQLQuery = sgSQLQuery & "tteEnteredTime, "
        sgSQLQuery = sgSQLQuery & "tteUieCode, "
        sgSQLQuery = sgSQLQuery & "tteUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlTTE.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlTTE.sType) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlTTE.sName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlTTE.sDescription) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlTTE.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlTTE.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlTTE.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlTTE.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlTTE.iOrigTteCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlTTE.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlTTE.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlTTE.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlTTE.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlTTE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_TTE_TimeType = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlTTE.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_TTE_TimeType = False
    Exit Function
End Function


Public Function gPutInsert_UIE_UserInfo(ilInsertType As Integer, tlUie As UIE, slForm_Module As String) As Integer
'
'   ilInsertType(I) = 0 New (set OrigUieCode); 1=From Update (retain OrigUieCode)
'   tlUIE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlUie.iCode
    Do
        If tlUie.iCode <= 0 Then
            sgSQLQuery = "Select MAX(uieCode) from UIE_User_Info"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlUie.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlUie.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlUie.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlUie.iCode
        sgSQLQuery = "INSERT INTO UIE_User_Info ("
        sgSQLQuery = sgSQLQuery & "uieCode, "
        sgSQLQuery = sgSQLQuery & "uieSignOnName, "
        sgSQLQuery = sgSQLQuery & "uiePassword, "
        sgSQLQuery = sgSQLQuery & "uieLastDatePWSet, "
        sgSQLQuery = sgSQLQuery & "uieShowName, "
        sgSQLQuery = sgSQLQuery & "uieState, "
        sgSQLQuery = sgSQLQuery & "uieEMail, "
        sgSQLQuery = sgSQLQuery & "uieLastSignOnDate, "
        sgSQLQuery = sgSQLQuery & "uieLastSignOnTime, "
        sgSQLQuery = sgSQLQuery & "uieUsedFlag, "
        sgSQLQuery = sgSQLQuery & "uieVersion, "
        sgSQLQuery = sgSQLQuery & "uieOrigUieCode, "
        sgSQLQuery = sgSQLQuery & "uieCurrent, "
        sgSQLQuery = sgSQLQuery & "uieEnteredDate, "
        sgSQLQuery = sgSQLQuery & "uieEnteredTime, "
        sgSQLQuery = sgSQLQuery & "uieUieCode, "
        sgSQLQuery = sgSQLQuery & "uieUnused )"
        sgSQLQuery = sgSQLQuery & "VALUES ("
        sgSQLQuery = sgSQLQuery & tlUie.iCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlUie.sSignOnName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlUie.sPassword) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlUie.sLastDatePWSet, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlUie.sShowName) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlUie.sState) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlUie.sEMail) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlUie.sLastSignOnDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlUie.sLastSignOnTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlUie.sUsedFlag) & "', "
        sgSQLQuery = sgSQLQuery & tlUie.iVersion & ", "
        If (ilInitCode <= 0) And (ilInsertType = 0) Then
            sgSQLQuery = sgSQLQuery & tlUie.iCode & ", "
        Else
            sgSQLQuery = sgSQLQuery & tlUie.iOrigUieCode & ", "
        End If
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlUie.sCurrent) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlUie.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlUie.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlUie.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlUie.sUnused) & "'"
        sgSQLQuery = sgSQLQuery & ")"
        cnn.BeginTrans
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    cnn.CommitTrans
    gPutInsert_UIE_UserInfo = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlUie.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_UIE_UserInfo = False
    Exit Function
End Function

Public Function gPutInsert_UTE_UserTasks(tlUte As UTE, slForm_Module As String) As Integer
'
'   tlUTE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim ilLastCode As Integer
    Dim ilInitCode As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    ilLastCode = 0
    ilInitCode = tlUte.iCode
    Do
        If tlUte.iCode <= 0 Then
            sgSQLQuery = "Select MAX(uteCode) from UTE_User_Tasks"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlUte.iCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlUte.iCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If ilLastCode = tlUte.iCode Then
                GoTo ErrHand1:
            End If
        End If
        ilLastCode = tlUte.iCode
        sgSQLQuery = "INSERT INTO UTE_USER_TASKS ("
        sgSQLQuery = sgSQLQuery & "uteCode, "
        sgSQLQuery = sgSQLQuery & "uteUieCode, "
        sgSQLQuery = sgSQLQuery & "uteTneCode, "
        sgSQLQuery = sgSQLQuery & "uteTaskStatus, "
        sgSQLQuery = sgSQLQuery & "uteUnused )"
        sgSQLQuery = sgSQLQuery & "VALUES ("
        sgSQLQuery = sgSQLQuery & tlUte.iCode & ", "
        sgSQLQuery = sgSQLQuery & tlUte.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & tlUte.iTneCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & tlUte.sTaskStatus & "', "
        sgSQLQuery = sgSQLQuery & "'" & tlUte.sUnused & "'"
        sgSQLQuery = sgSQLQuery & ")"
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_UTE_UserTasks = True
    Exit Function
ErrHand:
    If (ilInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (ilInitCode = 0) Then
                tlUte.iCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_UTE_UserTasks = False
    Exit Function
End Function


Public Function gPutUpdate_AEE_AutoEquip(ilUpdateType As Integer, tlAEE As AEE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE
'   tlUIE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldAEE As AEE
    
    On Error GoTo ErrHand
    
    
    ilRet = gGetRec_AEE_AutoEquip(tlAEE.iCode, slForm_Module, tlOldAEE)
    If ilRet Then
        
        tlOldAEE.iCode = 0
        tlOldAEE.sCurrent = "N"
        ilRet = gPutInsert_AEE_AutoEquip(1, tlOldAEE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlAEE.iVersion, "AEE", CLng(tlOldAEE.iCode), CLng(tlAEE.iCode), CLng(tlAEE.iOrigAeeCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_AEE_AutoEquip = False
                Exit Function
            End If
        
            sgSQLQuery = "UPDATE ADE_Auto_Data_Flags SET "
            sgSQLQuery = sgSQLQuery & "adeAeeCode = " & tlOldAEE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE adeAeeCode = " & tlAEE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            
            sgSQLQuery = "UPDATE ACE_Auto_Contact SET "
            sgSQLQuery = sgSQLQuery & "aceAeeCode = " & tlOldAEE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE aceAeeCode = " & tlAEE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect

            sgSQLQuery = "UPDATE AFE_Auto_Format SET "
            sgSQLQuery = sgSQLQuery & "afeAeeCode = " & tlOldAEE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE afeAeeCode = " & tlAEE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect

            sgSQLQuery = "UPDATE APE_Auto_Path SET "
            sgSQLQuery = sgSQLQuery & "apeAeeCode = " & tlOldAEE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE apeAeeCode = " & tlAEE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "UPDATE AEE_Auto_Equip SET "
            sgSQLQuery = sgSQLQuery & "aeeCode = " & tlAEE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "aeeName = '" & gFixQuote(tlAEE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "aeeDescription = '" & gFixQuote(tlAEE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "aeeManufacture = '" & gFixQuote(tlAEE.sManufacture) & "', "
            sgSQLQuery = sgSQLQuery & "aeeFixedTimeChar = '" & gFixQuote(tlAEE.sFixedTimeChar) & "', "
            sgSQLQuery = sgSQLQuery & "aeeAlertSchdDelay = " & tlAEE.lAlertSchdDelay & ", "
            sgSQLQuery = sgSQLQuery & "aeeState = '" & gFixQuote(tlAEE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "aeeUsedFlag = '" & gFixQuote(tlAEE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "aeeVersion = " & tlAEE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "aeeOrigAeeCode = " & tlAEE.iOrigAeeCode & ", "
            sgSQLQuery = sgSQLQuery & "aeeCurrent = '" & gFixQuote(tlAEE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "aeeEnteredDate = '" & Format$(tlAEE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "aeeEnteredTime = '" & Format$(tlAEE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "aeeUieCode = " & tlAEE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "aeeUnused = '" & gFixQuote(tlAEE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE aeeCode = " & tlAEE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_AEE_AutoEquip = True
            Exit Function
        Else
            gPutUpdate_AEE_AutoEquip = False
            Exit Function
        End If
    Else
        gPutUpdate_AEE_AutoEquip = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_AEE_AutoEquip = False
    Exit Function

End Function

Public Function gPutUpdate_ANE_AudioName(ilUpdateType As Integer, tlANE As ANE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlANE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldANE As ANE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update ANE_Audio_Name Set "
        sgSQLQuery = sgSQLQuery & "aneUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE aneCode = " & tlANE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_ANE_AudioName = True
        Exit Function
    End If
    
    ilRet = gGetRec_ANE_AudioName(tlANE.iCode, slForm_Module, tlOldANE)
    If ilRet Then
        
        tlOldANE.iCode = 0
        tlOldANE.sCurrent = "N"
        ilRet = gPutInsert_ANE_AudioName(1, tlOldANE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlANE.iVersion, "ANE", CLng(tlOldANE.iCode), CLng(tlANE.iCode), CLng(tlANE.iOrigAneCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_ANE_AudioName = False
                Exit Function
            End If
            
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update ANE_Audio_Name Set "
            sgSQLQuery = sgSQLQuery & "aneCode = " & tlANE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "aneName = '" & gFixQuote(tlANE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "aneDescription = '" & gFixQuote(tlANE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "aneCCECode = " & tlANE.iCceCode & ", "
            sgSQLQuery = sgSQLQuery & "aneAteCode = " & tlANE.iAteCode & ", "
            sgSQLQuery = sgSQLQuery & "aneState = '" & gFixQuote(tlANE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "aneUsedFlag = '" & gFixQuote(tlANE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "aneVersion = " & tlANE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "aneOrigAneCode = " & tlANE.iOrigAneCode & ", "
            sgSQLQuery = sgSQLQuery & "aneCurrent = '" & gFixQuote(tlANE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "aneEnteredDate = '" & Format$(tlANE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "aneEnteredTime = '" & Format$(tlANE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "aneUieCode = " & tlANE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "aneCheckConflicts = '" & gFixQuote(tlANE.sCheckConflicts) & "', "
            sgSQLQuery = sgSQLQuery & "aneUnused = '" & gFixQuote(tlANE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE aneCode = " & tlANE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_ANE_AudioName = True
            Exit Function
        Else
            gPutUpdate_ANE_AudioName = False
            Exit Function
        End If
    Else
        gPutUpdate_ANE_AudioName = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_ANE_AudioName = False
    Exit Function

End Function

Public Function gPutUpdate_ASE_AudioSource(ilUpdateType As Integer, tlASE As ASE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlASE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldASE As ASE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update ASE_Audio_Source Set "
        sgSQLQuery = sgSQLQuery & "aseUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE aseCode = " & tlASE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_ASE_AudioSource = True
        Exit Function
    End If
    ilRet = gGetRec_ASE_AudioSource(tlASE.iCode, slForm_Module, tlOldASE)
    If ilRet Then
        
        tlOldASE.iCode = 0
        tlOldASE.sCurrent = "N"
        ilRet = gPutInsert_ASE_AudioSource(1, tlOldASE, slForm_Module)
        If ilRet Then
            ilRet = gUpdateAIE(ilUpdateType, tlASE.iVersion, "ASE", CLng(tlOldASE.iCode), CLng(tlASE.iCode), CLng(tlASE.iOrigAseCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_ASE_AudioSource = False
                Exit Function
            End If
            
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update ASE_Audio_Source Set "
            sgSQLQuery = sgSQLQuery & "aseCode = " & tlASE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "asePriAneCode = " & tlASE.iPriAneCode & ", "
            sgSQLQuery = sgSQLQuery & "asePriCceCode = " & tlASE.iPriCceCode & ", "
            sgSQLQuery = sgSQLQuery & "aseDescription = '" & gFixQuote(tlASE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "aseBkupAneCode = " & tlASE.iBkupAneCode & ", "
            sgSQLQuery = sgSQLQuery & "aseBkupCceCode = " & tlASE.iBkupCceCode & ", "
            sgSQLQuery = sgSQLQuery & "aseProtAneCode = " & tlASE.iProtAneCode & ", "
            sgSQLQuery = sgSQLQuery & "aseProtCceCode = " & tlASE.iProtCceCode & ", "
            sgSQLQuery = sgSQLQuery & "aseState = '" & gFixQuote(tlASE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "aseUsedFlag = '" & gFixQuote(tlASE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "aseVersion = " & tlASE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "aseOrigAseCode = " & tlASE.iOrigAseCode & ", "
            sgSQLQuery = sgSQLQuery & "aseCurrent = '" & gFixQuote(tlASE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "aseEnteredDate = '" & Format$(tlASE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "aseEnteredTime = '" & Format$(tlASE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "aseUieCode = " & tlASE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "aseUnused = '" & gFixQuote(tlASE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE aseCode = " & tlASE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_ASE_AudioSource = True
            Exit Function
        Else
            gPutUpdate_ASE_AudioSource = False
            Exit Function
        End If
    Else
        gPutUpdate_ASE_AudioSource = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_ASE_AudioSource = False
    Exit Function

End Function

Public Function gPutUpdate_ATE_AudioType(ilUpdateType As Integer, tlATE As ATE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlATE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldATE As ATE
    
    On Error GoTo ErrHand
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update ATE_Audio_Type Set "
        sgSQLQuery = sgSQLQuery & "ateUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE ateCode = " & tlATE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_ATE_AudioType = True
        Exit Function
    End If
    
    ilRet = gGetRec_ATE_AudioType(tlATE.iCode, slForm_Module, tlOldATE)
    If ilRet Then
        
        tlOldATE.iCode = 0
        tlOldATE.sCurrent = "N"
        ilRet = gPutInsert_ATE_AudioType(1, tlOldATE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlATE.iVersion, "ATE", CLng(tlOldATE.iCode), CLng(tlATE.iCode), CLng(tlATE.iOrigAteCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_ATE_AudioType = False
                Exit Function
            End If
        
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update ATE_Audio_Type Set "
            sgSQLQuery = sgSQLQuery & "ateCode = " & tlATE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "ateName = '" & gFixQuote(tlATE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "ateDescription = '" & gFixQuote(tlATE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "ateState = '" & gFixQuote(tlATE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "ateTestItemID = '" & gFixQuote(tlATE.sTestItemID) & "', "
            sgSQLQuery = sgSQLQuery & "atePreBufferTime = " & tlATE.lPreBufferTime & ", "
            sgSQLQuery = sgSQLQuery & "atePostBufferTime = " & tlATE.lPostBufferTime & ", "
            sgSQLQuery = sgSQLQuery & "ateUsedFlag = '" & gFixQuote(tlATE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "ateVersion = " & tlATE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "ateOrigAteCode = " & tlATE.iOrigAteCode & ", "
            sgSQLQuery = sgSQLQuery & "ateCurrent = '" & gFixQuote(tlATE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "ateEnteredDate = '" & Format$(tlATE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "ateEnteredTime = '" & Format$(tlATE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "ateUieCode = " & tlATE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "ateUnused = '" & gFixQuote(tlATE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE ateCode = " & tlATE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_ATE_AudioType = True
            Exit Function
        Else
            gPutUpdate_ATE_AudioType = False
            Exit Function
        End If
    Else
        gPutUpdate_ATE_AudioType = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_ATE_AudioType = False
    Exit Function

End Function

Public Function gPutUpdate_BDE_BusDefinition(ilUpdateType As Integer, tlBDE As BDE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlBDE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldBDE As BDE
    
    On Error GoTo ErrHand
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update BDE_Bus_Definition Set "
        sgSQLQuery = sgSQLQuery & "bdeUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE bdeCode = " & tlBDE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_BDE_BusDefinition = True
        Exit Function
    End If
    ilRet = gGetRec_BDE_BusDefinition(tlBDE.iCode, slForm_Module, tlOldBDE)
    If ilRet Then
        
        tlOldBDE.iCode = 0
        tlOldBDE.sCurrent = "N"
        ilRet = gPutInsert_BDE_BusDefinition(1, tlOldBDE, slForm_Module)
        If ilRet Then

            ilRet = gUpdateAIE(ilUpdateType, tlBDE.iVersion, "BDE", CLng(tlOldBDE.iCode), CLng(tlBDE.iCode), CLng(tlBDE.iOrigBdeCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_BDE_BusDefinition = False
                Exit Function
            End If

            sgSQLQuery = "UPDATE BSE_Bus_Sel_Group SET "
            sgSQLQuery = sgSQLQuery & "bseBdeCode = " & tlOldBDE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE bseBdeCode = " & tlBDE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
        
            sgSQLQuery = "UPDATE DBE_Day_Bus_Sel SET "
            sgSQLQuery = sgSQLQuery & "dbeBdeCode = " & tlOldBDE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE dbeBdeCode = " & tlBDE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
        
            sgSQLQuery = "UPDATE EBE_Event_Bus_Sel SET "
            sgSQLQuery = sgSQLQuery & "ebeBdeCode = " & tlOldBDE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE ebeBdeCode = " & tlBDE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
        
            sgSQLQuery = "UPDATE SEE_Schedule_Events SET "
            sgSQLQuery = sgSQLQuery & "seeBdeCode = " & tlOldBDE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE seeBdeCode = " & tlBDE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
        
            sgSQLQuery = "UPDATE TSE_Template_Schd SET "
            sgSQLQuery = sgSQLQuery & "tseBdeCode = " & tlOldBDE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE tseBdeCode = " & tlBDE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
        
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update BDE_Bus_Definition Set "
            sgSQLQuery = sgSQLQuery & "bdeCode = " & tlBDE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "bdeName = '" & gFixQuote(tlBDE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "bdeDescription = '" & gFixQuote(tlBDE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "bdeChannel = '" & gFixQuote(tlBDE.sChannel) & "', "
            sgSQLQuery = sgSQLQuery & "bdeAseCode = " & tlBDE.iAseCode & ", "
            sgSQLQuery = sgSQLQuery & "bdeState = '" & gFixQuote(tlBDE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "bdeCceCode = " & tlBDE.iCceCode & ", "
            sgSQLQuery = sgSQLQuery & "bdeUsedFlag = '" & gFixQuote(tlBDE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "bdeVersion = " & tlBDE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "bdeOrigBdeCode = " & tlBDE.iOrigBdeCode & ", "
            sgSQLQuery = sgSQLQuery & "bdeCurrent = '" & gFixQuote(tlBDE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "bdeEnteredDate = '" & Format$(tlBDE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "bdeEnteredTime = '" & Format$(tlBDE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "bdeUieCode = " & tlBDE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "bdeUnused = '" & gFixQuote(tlBDE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE bdeCode = " & tlBDE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_BDE_BusDefinition = True
            Exit Function
        Else
            gPutUpdate_BDE_BusDefinition = False
            Exit Function
        End If
    Else
        gPutUpdate_BDE_BusDefinition = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_BDE_BusDefinition = False
    Exit Function

End Function

Public Function gPutUpdate_BGE_BusGroup(ilUpdateType As Integer, tlBGE As BGE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlBGE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldBGE As BGE
    
    On Error GoTo ErrHand
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update BGE_Bus_Group Set "
        sgSQLQuery = sgSQLQuery & "bgeUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE bgeCode = " & tlBGE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_BGE_BusGroup = True
        Exit Function
    End If
    ilRet = gGetRec_BGE_BusGroup(tlBGE.iCode, slForm_Module, tlOldBGE)
    If ilRet Then
        
        tlOldBGE.iCode = 0
        tlOldBGE.sCurrent = "N"
        ilRet = gPutInsert_BGE_BusGroup(1, tlOldBGE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlBGE.iVersion, "BGE", CLng(tlOldBGE.iCode), CLng(tlBGE.iCode), CLng(tlBGE.iOrigBgeCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_BGE_BusGroup = False
                Exit Function
            End If
        
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update BGE_Bus_Group Set "
            sgSQLQuery = sgSQLQuery & "bgeCode = " & tlBGE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "bgeName = '" & gFixQuote(tlBGE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "bgeDescription = '" & gFixQuote(tlBGE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "bgeState = '" & gFixQuote(tlBGE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "bgeUsedFlag = '" & gFixQuote(tlBGE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "bgeVersion = " & tlBGE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "bgeOrigBgeCode = " & tlBGE.iOrigBgeCode & ", "
            sgSQLQuery = sgSQLQuery & "bgeCurrent = '" & gFixQuote(tlBGE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "bgeEnteredDate = '" & Format$(tlBGE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "bgeEnteredTime = '" & Format$(tlBGE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "bgeUieCode = " & tlBGE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "bgeUnused = '" & gFixQuote(tlBGE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE bgeCode = " & tlBGE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_BGE_BusGroup = True
            Exit Function
        Else
            gPutUpdate_BGE_BusGroup = False
            Exit Function
        End If
    Else
        gPutUpdate_BGE_BusGroup = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_BGE_BusGroup = False
    Exit Function

End Function

Public Function gPutUpdate_CCE_ControlChar(ilUpdateType As Integer, tlCCE As CCE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlCCE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldCCE As CCE
    
    On Error GoTo ErrHand
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update CCE_Control_Char Set "
        sgSQLQuery = sgSQLQuery & "cceUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE cceCode = " & tlCCE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_CCE_ControlChar = True
        Exit Function
    End If
    ilRet = gGetRec_CCE_ControlChar(tlCCE.iCode, slForm_Module, tlOldCCE)
    If ilRet Then
        
        tlOldCCE.iCode = 0
        tlOldCCE.sCurrent = "N"
        ilRet = gPutInsert_CCE_ControlChar(1, tlOldCCE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlCCE.iVersion, "CCE", CLng(tlOldCCE.iCode), CLng(tlCCE.iCode), CLng(tlCCE.iOrigCceCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_CCE_ControlChar = False
                Exit Function
            End If
            
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update CCE_Control_Char Set "
            sgSQLQuery = sgSQLQuery & "cceCode = " & tlCCE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "cceType = '" & gFixQuote(tlCCE.sType) & "', "
            sgSQLQuery = sgSQLQuery & "cceAutoChar = '" & gFixQuote(tlCCE.sAutoChar) & "', "
            sgSQLQuery = sgSQLQuery & "cceDescription = '" & gFixQuote(tlCCE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "cceState = '" & gFixQuote(tlCCE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "cceUsedFlag = '" & gFixQuote(tlCCE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "cceVersion = " & tlCCE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "cceOrigCceCode = " & tlCCE.iOrigCceCode & ", "
            sgSQLQuery = sgSQLQuery & "cceCurrent = '" & gFixQuote(tlCCE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "cceEnteredDate = '" & Format$(tlCCE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "cceEnteredTime = '" & Format$(tlCCE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "cceUieCode = " & tlCCE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "cceUnused = '" & gFixQuote(tlCCE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE cceCode = " & tlCCE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_CCE_ControlChar = True
            Exit Function
        Else
            gPutUpdate_CCE_ControlChar = False
            Exit Function
        End If
    Else
        gPutUpdate_CCE_ControlChar = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_CCE_ControlChar = False
    Exit Function

End Function

Public Function gPutUpdate_CTE_CommtsTitle(ilUpdateType As Integer, tlCTE As CTE, slForm_Module As String, hlCTE As Integer) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlCTE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slMsg As String
    Dim tlOldCTE As CTE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update CTE_Commts_And_Title Set "
        sgSQLQuery = sgSQLQuery & "cteUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE cteCode = " & tlCTE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_CTE_CommtsTitle = True
        Exit Function
    End If
    
    ilRet = gGetRec_CTE_CommtsTitle(tlCTE.lCode, slForm_Module, tlOldCTE)
    If ilRet Then
        
        tlOldCTE.lCode = 0
        tlOldCTE.sCurrent = "N"
        ilRet = gPutInsert_CTE_CommtsTitle(1, tlOldCTE, slForm_Module, hlCTE)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlCTE.iVersion, "CTE", tlOldCTE.lCode, tlCTE.lCode, tlCTE.lOrigCteCode, slForm_Module)
            If Not ilRet Then
                gPutUpdate_CTE_CommtsTitle = False
                Exit Function
            End If
            
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update CTE_Commts_And_Title Set "
            sgSQLQuery = sgSQLQuery & "cteCode = " & tlCTE.lCode & ", "
            sgSQLQuery = sgSQLQuery & "cteType = '" & gFixQuote(tlCTE.sType) & "', "
            sgSQLQuery = sgSQLQuery & "cteComment = '" & gFixQuote(tlCTE.sComment) & "', "
            sgSQLQuery = sgSQLQuery & "cteState = '" & gFixQuote(tlCTE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "cteUsedFlag = '" & gFixQuote(tlCTE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "cteVersion = " & tlCTE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "cteOrigCteCode = " & tlCTE.lOrigCteCode & ", "
            sgSQLQuery = sgSQLQuery & "cteCurrent = '" & gFixQuote(tlCTE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "cteEnteredDate = '" & Format$(tlCTE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "cteEnteredTime = '" & Format$(tlCTE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "cteUieCode = " & tlCTE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "cteUnused = '" & gFixQuote(tlCTE.sUnused) & "' "

            sgSQLQuery = sgSQLQuery & " WHERE cteCode = " & tlCTE.lCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_CTE_CommtsTitle = True
            Exit Function
        Else
            gPutUpdate_CTE_CommtsTitle = False
            Exit Function
        End If
    Else
        gPutUpdate_CTE_CommtsTitle = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_CTE_CommtsTitle = False
    Exit Function

End Function

Public Function gPutUpdate_CTE_UsedFlag(llCode As Long, tlCurrCTE() As CTE, hlCTE As Integer) As Integer
    Dim ilCTE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If llCode > 0 Then
        For ilCTE = 0 To UBound(tlCurrCTE) - 1 Step 1
            If llCode = tlCurrCTE(ilCTE).lCode Then
                If Trim$(tlCurrCTE(ilCTE).sUsedFlag) <> "Y" Then
                    tgCTE.lCode = llCode
                    ilRet = gPutUpdate_CTE_CommtsTitle(2, tgCTE, "Update Used Flag in CTE", hlCTE)
                    tlCurrCTE(ilCTE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilCTE
    End If
    gPutUpdate_CTE_UsedFlag = ilRet
End Function


Public Function gPutUpdate_DHE_DayHeaderInfo(ilUpdateType As Integer, tlDHE As DHE, slForm_Module As String, llNewAgedDHECode As Long) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag;
'                     3=Like 1 but don't update DEE or DBE; 4=State and Current; 5=Only end date; 6 = Bus Names; 7=Only start date
'   tlDHE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldDHE As DHE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update DHE_Day_Header_Info Set "
        sgSQLQuery = sgSQLQuery & "dheUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE dheCode = " & tlDHE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_DHE_DayHeaderInfo = True
        Exit Function
    End If
    
    '5/31/11: Disallow changes in the schedule area
    If ilUpdateType = 4 Then    'Set Sate
        sgSQLQuery = "Update DHE_Day_Header_Info Set "
        sgSQLQuery = sgSQLQuery & "dheState = '" & tlDHE.sState & "'"
        If tlDHE.sState = "D" Then
            sgSQLQuery = sgSQLQuery & ", dheCurrent = '" & tlDHE.sCurrent & "'"
        End If
        sgSQLQuery = sgSQLQuery & " WHERE dheCode = " & tlDHE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_DHE_DayHeaderInfo = True
        Exit Function
    End If
    
    '5/31/11: Disallow changes in the schedule area
    If ilUpdateType = 5 Then    'Set end date
        sgSQLQuery = "Update DHE_Day_Header_Info Set "
        sgSQLQuery = sgSQLQuery & "dheEndDate = '" & Format$(tlDHE.sEndDate, sgSQLDateForm) & "'"
        sgSQLQuery = sgSQLQuery & " WHERE dheCode = " & tlDHE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_DHE_DayHeaderInfo = True
        Exit Function
    End If
    
    '5/31/11: Disallow changes in the schedule area
    If ilUpdateType = 6 Then    'Set bus names
        sgSQLQuery = "Update DHE_Day_Header_Info Set "
        sgSQLQuery = sgSQLQuery & "dheBusNames = '" & gFixQuote(Trim$(tlDHE.sBusNames)) & "'"
        sgSQLQuery = sgSQLQuery & " WHERE dheCode = " & tlDHE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_DHE_DayHeaderInfo = True
        Exit Function
    End If
    
    If ilUpdateType = 7 Then    'Set end date
        sgSQLQuery = "Update DHE_Day_Header_Info Set "
        sgSQLQuery = sgSQLQuery & "dheStartDate = '" & Format$(tlDHE.sStartDate, sgSQLDateForm) & "'"
        sgSQLQuery = sgSQLQuery & " WHERE dheCode = " & tlDHE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_DHE_DayHeaderInfo = True
        Exit Function
    End If
    
    ilRet = gGetRec_DHE_DayHeaderInfo(tlDHE.lCode, slForm_Module, tlOldDHE)
    If ilRet Then
        
        tlOldDHE.lCode = 0
        tlOldDHE.sCurrent = "N"
        ilRet = gPutInsert_DHE_DayHeaderInfo(1, tlOldDHE, slForm_Module)
        If ilRet Then
            llNewAgedDHECode = tlOldDHE.lCode
            If ilUpdateType <> 3 Then
                ilRet = gUpdateAIE(ilUpdateType, tlDHE.iVersion, "DHE", tlOldDHE.lCode, tlDHE.lCode, tlDHE.lOrigDHECode, slForm_Module)
            Else
                ilRet = gUpdateAIE(1, tlDHE.iVersion, "DHE", tlOldDHE.lCode, tlDHE.lCode, tlDHE.lOrigDHECode, slForm_Module)
            End If
            If Not ilRet Then
                gPutUpdate_DHE_DayHeaderInfo = False
                Exit Function
            End If

            'EBE does not need to be changed as the deeCode is not changed

            If ilUpdateType <> 3 Then
                sgSQLQuery = "UPDATE DBE_Day_Bus_Sel SET "
                sgSQLQuery = sgSQLQuery & "dbeDheCode = " & tlOldDHE.lCode
                sgSQLQuery = sgSQLQuery & " WHERE dbeDheCode = " & tlDHE.lCode
                cnn.Execute sgSQLQuery    ', rdExecDirect
                
                sgSQLQuery = "UPDATE DEE_Day_Event_Info SET "
                sgSQLQuery = sgSQLQuery & "deeDheCode = " & tlOldDHE.lCode
                sgSQLQuery = sgSQLQuery & " WHERE deeDheCode = " & tlDHE.lCode
                cnn.Execute sgSQLQuery    ', rdExecDirect
'Moved to EngrTempDef so that Update could be used.  Need to be able to associate TSE with correct old versions
'                If Trim$(tlDHE.sType) = "T" Then
'                    sgSQLQuery = "UPDATE TSE_Template_Schd SET "
'                    sgSQLQuery = sgSQLQuery & "tseDheCode = " & tlOldDHE.lCode & ", "
'                    sgSQLQuery = sgSQLQuery & "tseCurrent = " & "'N'"
'                    sgSQLQuery = sgSQLQuery & " WHERE tseDheCode = " & tlDHE.lCode
'                    cnn.Execute sgSQLQuery    ', rdExecDirect
'                End If
            End If

            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update DHE_Day_Header_Info Set "
            sgSQLQuery = sgSQLQuery & "dheCode = " & tlDHE.lCode & ", "
            sgSQLQuery = sgSQLQuery & "dheType = '" & gFixQuote(tlDHE.sType) & "', "
            sgSQLQuery = sgSQLQuery & "dheDneCode = " & tlDHE.lDneCode & ", "
            sgSQLQuery = sgSQLQuery & "dheDseCode = " & tlDHE.lDseCode & ", "
            sgSQLQuery = sgSQLQuery & "dheStartTime = '" & Format$(tlDHE.sStartTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "dheLength = " & tlDHE.lLength & ", "
            sgSQLQuery = sgSQLQuery & "dheHours = '" & gFixQuote(tlDHE.sHours) & "', "
            sgSQLQuery = sgSQLQuery & "dheStartDate = '" & Format$(tlDHE.sStartDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "dheEndDate = '" & Format$(tlDHE.sEndDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "dheDays = '" & gFixQuote(tlDHE.sDays) & "', "
            sgSQLQuery = sgSQLQuery & "dheCteCode = " & tlDHE.lCteCode & ", "
            sgSQLQuery = sgSQLQuery & "dheState = '" & gFixQuote(tlDHE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "dheUsedFlag = '" & gFixQuote(tlDHE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "dheVersion = " & tlDHE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "dheOrigDheCode = " & tlDHE.lOrigDHECode & ", "
            sgSQLQuery = sgSQLQuery & "dheCurrent = '" & gFixQuote(tlDHE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "dheEnteredDate = '" & Format$(tlDHE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "dheEnteredTime = '" & Format$(tlDHE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "dheUieCode = " & tlDHE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "dheIgnoreConflicts = '" & gFixQuote(tlDHE.sIgnoreConflicts) & "', "
            sgSQLQuery = sgSQLQuery & "dheBusNames = '" & gFixQuote(tlDHE.sBusNames) & "', "
            sgSQLQuery = sgSQLQuery & "dheUnused = '" & gFixQuote(tlDHE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE dheCode = " & tlDHE.lCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_DHE_DayHeaderInfo = True
            Exit Function
        Else
            gPutUpdate_DHE_DayHeaderInfo = False
            Exit Function
        End If
    Else
        gPutUpdate_DHE_DayHeaderInfo = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_DHE_DayHeaderInfo = False
    Exit Function

End Function


Public Function gPutUpdate_DNE_DayName(ilUpdateType As Integer, tlDNE As DNE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlDNE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldDNE As DNE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update DNE_Day_Name Set "
        sgSQLQuery = sgSQLQuery & "dneUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE dneCode = " & tlDNE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_DNE_DayName = True
        Exit Function
    End If
        
    ilRet = gGetRec_DNE_DayName(tlDNE.lCode, slForm_Module, tlOldDNE)
    If ilRet Then
        
        tlOldDNE.lCode = 0
        tlOldDNE.sCurrent = "N"
        ilRet = gPutInsert_DNE_DayName(1, tlOldDNE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlDNE.iVersion, "DNE", tlOldDNE.lCode, tlDNE.lCode, tlDNE.lOrigDneCode, slForm_Module)
            If Not ilRet Then
                gPutUpdate_DNE_DayName = False
                Exit Function
            End If
            
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update DNE_Day_Name Set "
            sgSQLQuery = sgSQLQuery & "dneCode = " & tlDNE.lCode & ", "
            sgSQLQuery = sgSQLQuery & "dneType = '" & gFixQuote(tlDNE.sType) & "', "
            sgSQLQuery = sgSQLQuery & "dneName = '" & gFixQuote(tlDNE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "dneDescription = '" & gFixQuote(tlDNE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "dneState = '" & gFixQuote(tlDNE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "dneUsedFlag = '" & gFixQuote(tlDNE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "dneVersion = " & tlDNE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "dneOrigDneCode = " & tlDNE.lOrigDneCode & ", "
            sgSQLQuery = sgSQLQuery & "dneCurrent = '" & gFixQuote(tlDNE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "dneEnteredDate = '" & Format$(tlDNE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "dneEnteredTime = '" & Format$(tlDNE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "dneUieCode = " & tlDNE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "dneUnused = '" & gFixQuote(tlDNE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE dneCode = " & tlDNE.lCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_DNE_DayName = True
            Exit Function
        Else
            gPutUpdate_DNE_DayName = False
            Exit Function
        End If
    Else
        gPutUpdate_DNE_DayName = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_DNE_DayName = False
    Exit Function

End Function


Public Function gPutUpdate_DSE_DaySubName(ilUpdateType As Integer, tlDSE As DSE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlDSE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldDSE As DSE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update DSE_Day_SubName Set "
        sgSQLQuery = sgSQLQuery & "dseUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE dseCode = " & tlDSE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_DSE_DaySubName = True
        Exit Function
    End If
    
    ilRet = gGetRec_DSE_DaySubName(tlDSE.lCode, slForm_Module, tlOldDSE)
    If ilRet Then
        
        tlOldDSE.lCode = 0
        tlOldDSE.sCurrent = "N"
        ilRet = gPutInsert_DSE_DaySubName(1, tlOldDSE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlDSE.iVersion, "DSE", tlOldDSE.lCode, tlDSE.lCode, tlDSE.lOrigDseCode, slForm_Module)
            If Not ilRet Then
                gPutUpdate_DSE_DaySubName = False
                Exit Function
            End If
                    
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update DSE_Day_SubName Set "
            sgSQLQuery = sgSQLQuery & "dseCode = " & tlDSE.lCode & ", "
            sgSQLQuery = sgSQLQuery & "dseName = '" & gFixQuote(tlDSE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "dseDescription = '" & gFixQuote(tlDSE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "dseState = '" & gFixQuote(tlDSE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "dseUsedFlag = '" & gFixQuote(tlDSE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "dseVersion = " & tlDSE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "dseOrigDseCode = " & tlDSE.lOrigDseCode & ", "
            sgSQLQuery = sgSQLQuery & "dseCurrent = '" & gFixQuote(tlDSE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "dseEnteredDate = '" & Format$(tlDSE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "dseEnteredTime = '" & Format$(tlDSE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "dseUieCode = " & tlDSE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "dseUnused = '" & gFixQuote(tlDSE.sUnused) & "' "

            sgSQLQuery = sgSQLQuery & " WHERE dseCode = " & tlDSE.lCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_DSE_DaySubName = True
            Exit Function
        Else
            gPutUpdate_DSE_DaySubName = False
            Exit Function
        End If
    Else
        gPutUpdate_DSE_DaySubName = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_DSE_DaySubName = False
    Exit Function

End Function


Public Function gPutUpdate_ETE_EventType(ilUpdateType As Integer, tlETE As ETE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlETE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldETE As ETE
    
    On Error GoTo ErrHand
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update ETE_Event_Type Set "
        sgSQLQuery = sgSQLQuery & "eteUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE eteCode = " & tlETE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_ETE_EventType = True
        Exit Function
    End If
    ilRet = gGetRec_ETE_EventType(tlETE.iCode, slForm_Module, tlOldETE)
    If ilRet Then
        
        tlOldETE.iCode = 0
        tlOldETE.sCurrent = "N"
        ilRet = gPutInsert_ETE_EventType(1, tlOldETE, slForm_Module)
        If ilRet Then

            ilRet = gUpdateAIE(ilUpdateType, tlETE.iVersion, "ETE", CLng(tlOldETE.iCode), CLng(tlETE.iCode), CLng(tlETE.iOrigEteCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_ETE_EventType = False
                Exit Function
            End If

            sgSQLQuery = "UPDATE EPE_Event_Properties SET "
            sgSQLQuery = sgSQLQuery & "epeEteCode = " & tlOldETE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE epeEteCode = " & tlETE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
        
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update ETE_Event_Type Set "
            sgSQLQuery = sgSQLQuery & "eteCode = " & tlETE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "eteName = '" & gFixQuote(tlETE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "eteDescription = '" & gFixQuote(tlETE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "eteCategory = '" & gFixQuote(tlETE.sCategory) & "', "
            sgSQLQuery = sgSQLQuery & "eteAutoCodeChar = '" & gFixQuote(tlETE.sAutoCodeChar) & "', "
            sgSQLQuery = sgSQLQuery & "eteState = '" & gFixQuote(tlETE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "eteUsedFlag = '" & gFixQuote(tlETE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "eteVersion = " & tlETE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "eteOrigEteCode = " & tlETE.iOrigEteCode & ", "
            sgSQLQuery = sgSQLQuery & "eteCurrent = '" & gFixQuote(tlETE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "eteEnteredDate = '" & Format$(tlETE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "eteEnteredTime = '" & Format$(tlETE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "eteUieCode = " & tlETE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "eteUnused = '" & gFixQuote(tlETE.sUnused) & "' "

            sgSQLQuery = sgSQLQuery & " WHERE eteCode = " & tlETE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_ETE_EventType = True
            Exit Function
        Else
            gPutUpdate_ETE_EventType = False
            Exit Function
        End If
    Else
        gPutUpdate_ETE_EventType = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_ETE_EventType = False
    Exit Function

End Function
Public Function gPutUpdate_ETE_UsedFlag(ilCode As Integer, tlCurrETE() As ETE) As Integer
    Dim ilETE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilETE = 0 To UBound(tlCurrETE) - 1 Step 1
            If ilCode = tlCurrETE(ilETE).iCode Then
                If Trim$(tlCurrETE(ilETE).sUsedFlag) <> "Y" Then
                    tgETE.iCode = ilCode
                    ilRet = gPutUpdate_ETE_EventType(2, tgETE, "Update Used Flag in ETE")
                    tlCurrETE(ilETE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilETE
    End If
    gPutUpdate_ETE_UsedFlag = ilRet
End Function

Public Function gPutUpdate_FNE_FollowName(ilUpdateType As Integer, tlFNE As FNE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlFNE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldFNE As FNE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update FNE_Follow_Name Set "
        sgSQLQuery = sgSQLQuery & "fneUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE fneCode = " & tlFNE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_FNE_FollowName = True
        Exit Function
    End If
    
    ilRet = gGetRec_FNE_FollowName(tlFNE.iCode, slForm_Module, tlOldFNE)
    If ilRet Then
        
        tlOldFNE.iCode = 0
        tlOldFNE.sCurrent = "N"
        ilRet = gPutInsert_FNE_FollowName(1, tlOldFNE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlFNE.iVersion, "FNE", CLng(tlOldFNE.iCode), CLng(tlFNE.iCode), CLng(tlFNE.iOrigFneCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_FNE_FollowName = False
                Exit Function
            End If
            
        
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update FNE_Follow_Name Set "
            sgSQLQuery = sgSQLQuery & "fneCode = " & tlFNE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "fneName = '" & gFixQuote(tlFNE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "fneDescription = '" & gFixQuote(tlFNE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "fneState = '" & gFixQuote(tlFNE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "fneUsedFlag = '" & gFixQuote(tlFNE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "fneVersion = " & tlFNE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "fneOrigFneCode = " & tlFNE.iOrigFneCode & ", "
            sgSQLQuery = sgSQLQuery & "fneCurrent = '" & gFixQuote(tlFNE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "fneEnteredDate = '" & Format$(tlFNE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "fneEnteredTime = '" & Format$(tlFNE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "fneUieCode = " & tlFNE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "fneUnused = '" & gFixQuote(tlFNE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE fneCode = " & tlFNE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_FNE_FollowName = True
            Exit Function
        Else
            gPutUpdate_FNE_FollowName = False
            Exit Function
        End If
    Else
        gPutUpdate_FNE_FollowName = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_FNE_FollowName = False
    Exit Function

End Function
Public Function gPutUpdate_FNE_UsedFlag(ilCode As Integer, tlCurrFNE() As FNE) As Integer
    Dim ilFNE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilFNE = 0 To UBound(tlCurrFNE) - 1 Step 1
            If ilCode = tlCurrFNE(ilFNE).iCode Then
                If Trim$(tlCurrFNE(ilFNE).sUsedFlag) <> "Y" Then
                    tgFNE.iCode = ilCode
                    ilRet = gPutUpdate_FNE_FollowName(2, tgFNE, "Update Used Flag in FNE")
                    tlCurrFNE(ilFNE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilFNE
    End If
    gPutUpdate_FNE_UsedFlag = ilRet
End Function


Public Function gPutUpdate_MTE_MaterialType(ilUpdateType As Integer, tlMTE As MTE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlMTE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldMTE As MTE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update MTE_Material_Type Set "
        sgSQLQuery = sgSQLQuery & "mteUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE mteCode = " & tlMTE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_MTE_MaterialType = True
        Exit Function
    End If
        
    ilRet = gGetRec_MTE_MaterialType(tlMTE.iCode, slForm_Module, tlOldMTE)
    If ilRet Then
        
        tlOldMTE.iCode = 0
        tlOldMTE.sCurrent = "N"
        ilRet = gPutInsert_MTE_MaterialType(1, tlOldMTE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlMTE.iVersion, "MTE", CLng(tlOldMTE.iCode), CLng(tlMTE.iCode), CLng(tlMTE.iOrigMteCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_MTE_MaterialType = False
                Exit Function
            End If
                    
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update MTE_Material_Type Set "
            sgSQLQuery = sgSQLQuery & "mteCode = " & tlMTE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "mteName = '" & gFixQuote(tlMTE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "mteDescription = '" & gFixQuote(tlMTE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "mteState = '" & gFixQuote(tlMTE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "mteUsedFlag = '" & gFixQuote(tlMTE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "mteVersion = " & tlMTE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "mteOrigMteCode = " & tlMTE.iOrigMteCode & ", "
            sgSQLQuery = sgSQLQuery & "mteCurrent = '" & gFixQuote(tlMTE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "mteEnteredDate = '" & Format$(tlMTE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "mteEnteredTime = '" & Format$(tlMTE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "mteUieCode = " & tlMTE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "mteUnused = '" & gFixQuote(tlMTE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE mteCode = " & tlMTE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_MTE_MaterialType = True
            Exit Function
        Else
            gPutUpdate_MTE_MaterialType = False
            Exit Function
        End If
    Else
        gPutUpdate_MTE_MaterialType = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_MTE_MaterialType = False
    Exit Function

End Function
Public Function gPutUpdate_MTE_UsedFlag(ilCode As Integer, tlCurrMTE() As MTE) As Integer
    Dim ilMTE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilMTE = 0 To UBound(tlCurrMTE) - 1 Step 1
            If ilCode = tlCurrMTE(ilMTE).iCode Then
                If Trim$(tlCurrMTE(ilMTE).sUsedFlag) <> "Y" Then
                    tgMTE.iCode = ilCode
                    ilRet = gPutUpdate_MTE_MaterialType(2, tgMTE, "Update Used Flag in MTE")
                    tlCurrMTE(ilMTE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilMTE
    End If
    gPutUpdate_MTE_UsedFlag = ilRet
End Function


Public Function gPutUpdate_NNE_NetcueName(ilUpdateType As Integer, tlNNE As NNE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag; 3=Description only
'   tlNNE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldNNE As NNE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update NNE_Netcue_Name Set "
        sgSQLQuery = sgSQLQuery & "nneUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE nneCode = " & tlNNE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_NNE_NetcueName = True
        Exit Function
    End If
    If ilUpdateType = 3 Then
        sgSQLQuery = "Update NNE_Netcue_Name Set "
        sgSQLQuery = sgSQLQuery & "nneDescription = '" & gFixQuote(tlNNE.sDescription) & "'"
        sgSQLQuery = sgSQLQuery & " WHERE nneCode = " & tlNNE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_NNE_NetcueName = True
        Exit Function
    End If
        
    ilRet = gGetRec_NNE_NetcueName(tlNNE.iCode, slForm_Module, tlOldNNE)
    If ilRet Then
        
        tlOldNNE.iCode = 0
        tlOldNNE.sCurrent = "N"
        ilRet = gPutInsert_NNE_NetcueName(1, tlOldNNE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlNNE.iVersion, "NNE", CLng(tlOldNNE.iCode), CLng(tlNNE.iCode), CLng(tlNNE.iOrigNneCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_NNE_NetcueName = False
                Exit Function
            End If
            
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update NNE_Netcue_Name Set "
            sgSQLQuery = sgSQLQuery & "nneCode = " & tlNNE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "nneName = '" & gFixQuote(tlNNE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "nneDescription = '" & gFixQuote(tlNNE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "nneDneCode = " & tlNNE.lDneCode & ", "
            sgSQLQuery = sgSQLQuery & "nneState = '" & gFixQuote(tlNNE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "nneUsedFlag = '" & gFixQuote(tlNNE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "nneVersion = " & tlNNE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "nneOrigNneCode = " & tlNNE.iOrigNneCode & ", "
            sgSQLQuery = sgSQLQuery & "nneCurrent = '" & gFixQuote(tlNNE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "nneEnteredDate = '" & Format$(tlNNE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "nneEnteredTime = '" & Format$(tlNNE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "nneUieCode = " & tlNNE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "nneUnused = '" & gFixQuote(tlNNE.sUnused) & "' "

            sgSQLQuery = sgSQLQuery & " WHERE nneCode = " & tlNNE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_NNE_NetcueName = True
            Exit Function
        Else
            gPutUpdate_NNE_NetcueName = False
            Exit Function
        End If
    Else
        gPutUpdate_NNE_NetcueName = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_NNE_NetcueName = False
    Exit Function

End Function
Public Function gPutUpdate_NNE_UsedFlag(ilCode As Integer, tlCurrNNE() As NNE) As Integer
    Dim ilNNE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilNNE = 0 To UBound(tlCurrNNE) - 1 Step 1
            If ilCode = tlCurrNNE(ilNNE).iCode Then
                If Trim$(tlCurrNNE(ilNNE).sUsedFlag) <> "Y" Then
                    tgNNE.iCode = ilCode
                    ilRet = gPutUpdate_NNE_NetcueName(2, tgNNE, "Update Used Flag in NNE")
                    tlCurrNNE(ilNNE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilNNE
    End If
    gPutUpdate_NNE_UsedFlag = ilRet
End Function

Public Function gPutUpdate_RNE_RelayName(ilUpdateType As Integer, tlRNE As RNE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlRNE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldRNE As RNE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update RNE_Relay_Name Set "
        sgSQLQuery = sgSQLQuery & "rneUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE rneCode = " & tlRNE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_RNE_RelayName = True
        Exit Function
    End If
    
    ilRet = gGetRec_RNE_RelayName(tlRNE.iCode, slForm_Module, tlOldRNE)
    If ilRet Then
        
        tlOldRNE.iCode = 0
        tlOldRNE.sCurrent = "N"
        ilRet = gPutInsert_RNE_RelayName(1, tlOldRNE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlRNE.iVersion, "RNE", CLng(tlOldRNE.iCode), CLng(tlRNE.iCode), CLng(tlRNE.iOrigRneCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_RNE_RelayName = False
                Exit Function
            End If
                    
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update RNE_Relay_Name Set "
            sgSQLQuery = sgSQLQuery & "rneCode = " & tlRNE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "rneName = '" & gFixQuote(tlRNE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "rneDescription = '" & gFixQuote(tlRNE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "rneState = '" & gFixQuote(tlRNE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "rneUsedFlag = '" & gFixQuote(tlRNE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "rneVersion = " & tlRNE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "rneOrigRneCode = " & tlRNE.iOrigRneCode & ", "
            sgSQLQuery = sgSQLQuery & "rneCurrent = '" & gFixQuote(tlRNE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "rneEnteredDate = '" & Format$(tlRNE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "rneEnteredTime = '" & Format$(tlRNE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "rneUieCode = " & tlRNE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "rneUnused = '" & gFixQuote(tlRNE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE rneCode = " & tlRNE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_RNE_RelayName = True
            Exit Function
        Else
            gPutUpdate_RNE_RelayName = False
            Exit Function
        End If
    Else
        gPutUpdate_RNE_RelayName = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_RNE_RelayName = False
    Exit Function

End Function
Public Function gPutUpdate_RNE_UsedFlag(ilCode As Integer, tlCurrRNE() As RNE) As Integer
    Dim ilRNE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilRNE = 0 To UBound(tlCurrRNE) - 1 Step 1
            If ilCode = tlCurrRNE(ilRNE).iCode Then
                If Trim$(tlCurrRNE(ilRNE).sUsedFlag) <> "Y" Then
                    tgRNE.iCode = ilCode
                    ilRet = gPutUpdate_RNE_RelayName(2, tgRNE, "Update Used Flag in RNE")
                    tlCurrRNE(ilRNE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilRNE
    End If
    gPutUpdate_RNE_UsedFlag = ilRet
End Function

Public Function gPutUpdate_SCE_SilenceChar(ilUpdateType As Integer, tlSCE As SCE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlSCE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldSCE As SCE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update SCE_Silence_Char Set "
        sgSQLQuery = sgSQLQuery & "sceUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE sceCode = " & tlSCE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_SCE_SilenceChar = True
        Exit Function
    End If
        
    ilRet = gGetRec_SCE_SilenceChar(tlSCE.iCode, slForm_Module, tlOldSCE)
    If ilRet Then
        
        tlOldSCE.iCode = 0
        tlOldSCE.sCurrent = "N"
        ilRet = gPutInsert_SCE_SilenceChar(1, tlOldSCE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlSCE.iVersion, "SCE", CLng(tlOldSCE.iCode), CLng(tlSCE.iCode), CLng(tlSCE.iOrigSceCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_SCE_SilenceChar = False
                Exit Function
            End If
                    
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update SCE_Silence_Char Set "
            sgSQLQuery = sgSQLQuery & "sceCode = " & tlSCE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "sceAutoChar = '" & gFixQuote(tlSCE.sAutoChar) & "', "
            sgSQLQuery = sgSQLQuery & "sceDescription = '" & gFixQuote(tlSCE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "sceState = '" & gFixQuote(tlSCE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "sceUsedFlag = '" & gFixQuote(tlSCE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "sceVersion = " & tlSCE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "sceOrigSceCode = " & tlSCE.iOrigSceCode & ", "
            sgSQLQuery = sgSQLQuery & "sceCurrent = '" & gFixQuote(tlSCE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "sceEnteredDate = '" & Format$(tlSCE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "sceEnteredTime = '" & Format$(tlSCE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "sceUieCode = " & tlSCE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "sceUnused = '" & gFixQuote(tlSCE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE sceCode = " & tlSCE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_SCE_SilenceChar = True
            Exit Function
        Else
            gPutUpdate_SCE_SilenceChar = False
            Exit Function
        End If
    Else
        gPutUpdate_SCE_SilenceChar = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_SCE_SilenceChar = False
    Exit Function

End Function
Public Function gPutUpdate_SCE_UsedFlag(ilCode As Integer, tlCurrSCE() As SCE) As Integer
    Dim ilSCE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilSCE = 0 To UBound(tlCurrSCE) - 1 Step 1
            If ilCode = tlCurrSCE(ilSCE).iCode Then
                If Trim$(tlCurrSCE(ilSCE).sUsedFlag) <> "Y" Then
                    tgSCE.iCode = ilCode
                    ilRet = gPutUpdate_SCE_SilenceChar(2, tgSCE, "Update Used Flag in SCE")
                    tlCurrSCE(ilSCE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilSCE
    End If
    gPutUpdate_SCE_UsedFlag = ilRet
End Function


Public Function gPutUpdate_SHE_ScheduleHeader(ilUpdateType As Integer, tlSHE As SHE, slForm_Module As String, llNewAgedSHECode As Long) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag; 3=Like 1 but don't update DEE or DB; 4= Update sequence number only; 5= Update Conflict flag
'   tlSHE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldSHE As SHE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update SHE_Schedule_Header Set "
        sgSQLQuery = sgSQLQuery & "sheUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE sheCode = " & tlSHE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_SHE_ScheduleHeader = True
        Exit Function
    End If
    If ilUpdateType = 5 Then
        sgSQLQuery = "Update SHE_Schedule_Header Set "
        sgSQLQuery = sgSQLQuery & "sheConflictExist = '" & gFixQuote(tlSHE.sConflictExist) & "'"
        sgSQLQuery = sgSQLQuery & " WHERE sheCode = " & tlSHE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_SHE_ScheduleHeader = True
        Exit Function
    End If
    If ilUpdateType = 6 Then
        sgSQLQuery = "Update SHE_Schedule_Header Set "
        sgSQLQuery = sgSQLQuery & "sheSpotMergeStatus = '" & gFixQuote(tlSHE.sSpotMergeStatus) & "'"
        sgSQLQuery = sgSQLQuery & " WHERE sheCode = " & tlSHE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_SHE_ScheduleHeader = True
        Exit Function
    End If
    If ilUpdateType = 7 Then
        sgSQLQuery = "Update SHE_Schedule_Header Set "
        sgSQLQuery = sgSQLQuery & "sheLoadStatus = '" & gFixQuote(tlSHE.sLoadStatus) & "'"
        sgSQLQuery = sgSQLQuery & " WHERE sheCode = " & tlSHE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_SHE_ScheduleHeader = True
        Exit Function
    End If
    If ilUpdateType = 8 Then
        sgSQLQuery = "Update SHE_Schedule_Header Set "
        sgSQLQuery = sgSQLQuery & "sheConflictExist = '" & gFixQuote(tlSHE.sConflictExist) & "'" & ", "
        sgSQLQuery = sgSQLQuery & "sheSpotMergeStatus = '" & gFixQuote(tlSHE.sSpotMergeStatus) & "'"
        sgSQLQuery = sgSQLQuery & " WHERE sheCode = " & tlSHE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_SHE_ScheduleHeader = True
        Exit Function
    End If
    If ilUpdateType = 4 Then
        sgSQLQuery = "Update SHE_Schedule_Header Set "
        sgSQLQuery = sgSQLQuery & "sheLoadedAutoStatus = '" & gFixQuote(tlSHE.sLoadedAutoStatus) & "', "
        sgSQLQuery = sgSQLQuery & "sheLoadedAutoDate = '" & Format$(tlSHE.sLoadedAutoDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "sheChgSeqNo = " & tlSHE.iChgSeqNo & ", "
        sgSQLQuery = sgSQLQuery & "sheCreateLoad = '" & gFixQuote(tlSHE.sCreateLoad) & "'"
        sgSQLQuery = sgSQLQuery & " WHERE sheCode = " & tlSHE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_SHE_ScheduleHeader = True
        Exit Function
    End If
    
    ilRet = gGetRec_SHE_ScheduleHeader(tlSHE.lCode, slForm_Module, tlOldSHE)
    If ilRet Then
        
        tlOldSHE.lCode = 0
        tlOldSHE.sCurrent = "N"
        ilRet = gPutInsert_SHE_ScheduleHeader(1, tlOldSHE, slForm_Module)
        If ilRet Then
            llNewAgedSHECode = tlOldSHE.lCode
            If ilUpdateType <> 3 Then
                ilRet = gUpdateAIE(ilUpdateType, tlSHE.iVersion, "SHE", tlOldSHE.lCode, tlSHE.lCode, tlSHE.lOrigSheCode, slForm_Module)
            Else
                ilRet = gUpdateAIE(1, tlSHE.iVersion, "SHE", tlOldSHE.lCode, tlSHE.lCode, tlSHE.lOrigSheCode, slForm_Module)
            End If
            If Not ilRet Then
                gPutUpdate_SHE_ScheduleHeader = False
                Exit Function
            End If

            'EBE does not need to be changed as the deeCode is not changed

            If ilUpdateType <> 3 Then
                
                sgSQLQuery = "UPDATE SEE_Schedule_Events SET "
                sgSQLQuery = sgSQLQuery & "seeSheCode = " & tlOldSHE.lCode
                sgSQLQuery = sgSQLQuery & " WHERE seeSheCode = " & tlSHE.lCode
                cnn.Execute sgSQLQuery    ', rdExecDirect
            End If

            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update SHE_Schedule_Header Set "
            sgSQLQuery = sgSQLQuery & "sheCode = " & tlSHE.lCode & ", "
            sgSQLQuery = sgSQLQuery & "sheAeeCode = " & tlSHE.iAeeCode & ", "
            sgSQLQuery = sgSQLQuery & "sheAirDate = '" & Format$(tlSHE.sAirDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "sheLoadedAutoStatus = '" & gFixQuote(tlSHE.sLoadedAutoStatus) & "', "
            sgSQLQuery = sgSQLQuery & "sheLoadedAutoDate = '" & Format$(tlSHE.sLoadedAutoDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "sheChgSeqNo = " & tlSHE.iChgSeqNo & ", "
            sgSQLQuery = sgSQLQuery & "sheAsAirStatus = '" & gFixQuote(tlSHE.sAsAirStatus) & "', "
            sgSQLQuery = sgSQLQuery & "sheLoadedAsAirDate = '" & Format$(tlSHE.sLoadedAsAirDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "sheLastDateItemChk = '" & Format$(tlSHE.sLastDateItemChk, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "sheCreateLoad = '" & gFixQuote(tlSHE.sCreateLoad) & "', "
            sgSQLQuery = sgSQLQuery & "sheVersion = " & tlSHE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "sheOrigSheCode = " & tlSHE.lOrigSheCode & ", "
            sgSQLQuery = sgSQLQuery & "sheCurrent = '" & gFixQuote(tlSHE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "sheEnteredDate = '" & Format$(tlSHE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "sheEnteredTime = '" & Format$(tlSHE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "sheUieCode = " & tlSHE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "sheConflictExist = '" & gFixQuote(tlSHE.sConflictExist) & "', "
            sgSQLQuery = sgSQLQuery & "sheSpotMergeStatus = '" & gFixQuote(tlSHE.sSpotMergeStatus) & "', "
            sgSQLQuery = sgSQLQuery & "sheLoadStatus = '" & gFixQuote(tlSHE.sLoadStatus) & "', "
            sgSQLQuery = sgSQLQuery & "sheUnused = '" & gFixQuote(tlSHE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE sheCode = " & tlSHE.lCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_SHE_ScheduleHeader = True
            Exit Function
        Else
            gPutUpdate_SHE_ScheduleHeader = False
            Exit Function
        End If
    Else
        gPutUpdate_SHE_ScheduleHeader = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_SHE_ScheduleHeader = False
    Exit Function

End Function

Public Function gPutUpdate_SOE_SiteOption(ilUpdateType As Integer, tlSOE As SOE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Only Event ID
'   tlSOE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldSOE As SOE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 1 Then
        sgSQLQuery = "UPDATE SOE_Site_Option SET "
        sgSQLQuery = sgSQLQuery & "soeCurrEventID = " & tlSOE.lCurrEventID
        sgSQLQuery = sgSQLQuery & " WHERE soeCode = " & tlSOE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_SOE_SiteOption = True
        Exit Function
    End If
    
    ilRet = gGetRec_SOE_SiteOption(tlSOE.iCode, slForm_Module, tlOldSOE)
    If ilRet Then
        
        tlOldSOE.iCode = 0
        tlOldSOE.sCurrent = "N"
        ilRet = gPutInsert_SOE_SiteOption(1, tlOldSOE, slForm_Module)
        If ilRet Then
        
            sgSQLQuery = "UPDATE SPE_Site_Path SET "
            sgSQLQuery = sgSQLQuery & "speSoeCode = " & tlOldSOE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE speSoeCode = " & tlSOE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            
            sgSQLQuery = "UPDATE SGE_Site_Gen_Schd SET "
            sgSQLQuery = sgSQLQuery & "sgeSoeCode = " & tlOldSOE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE sgeSoeCode = " & tlSOE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
           
            sgSQLQuery = "UPDATE ITE_Item_Test SET "
            sgSQLQuery = sgSQLQuery & "iteSoeCode = " & tlOldSOE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE iteSoeCode = " & tlSOE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
           
            sgSQLQuery = "UPDATE SSE_Site_SMTP_Info SET "
            sgSQLQuery = sgSQLQuery & "sseSoeCode = " & tlOldSOE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE sseSoeCode = " & tlSOE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
           
            slNowDate = Format(Now, sgShowDateForm)
            slNowTime = Format(Now, sgShowTimeWSecForm)
            sgSQLQuery = "UPDATE SOE_Site_Option SET "
            sgSQLQuery = sgSQLQuery & "soeClientName= '" & gFixQuote(tlSOE.sClientName) & "', "
            sgSQLQuery = sgSQLQuery & "soeAddr1 = '" & gFixQuote(tlSOE.sAddr1) & "', "
            sgSQLQuery = sgSQLQuery & "soeAddr2 = '" & gFixQuote(tlSOE.sAddr2) & "', "
            sgSQLQuery = sgSQLQuery & "soeAddr3 = '" & gFixQuote(tlSOE.sAddr3) & "', "
            sgSQLQuery = sgSQLQuery & "soePhone = '" & gFixQuote(tlSOE.sPhone) & "', "
            sgSQLQuery = sgSQLQuery & "soeFax = '" & gFixQuote(tlSOE.sFax) & "', "
            sgSQLQuery = sgSQLQuery & "soeDaysRetainAsAir = " & tlSOE.iDaysRetainAsAir & ", "
            sgSQLQuery = sgSQLQuery & "soeDaysRetainActive = " & tlSOE.iDaysRetainActive & ", "
            sgSQLQuery = sgSQLQuery & "soeChgInterval = " & tlSOE.lChgInterval & ", "
            sgSQLQuery = sgSQLQuery & "soeMergeDateFormat = '" & gFixQuote(tlSOE.sMergeDateFormat) & "', "
            sgSQLQuery = sgSQLQuery & "soeMergeTimeFormat = '" & gFixQuote(tlSOE.sMergeTimeFormat) & "', "
            sgSQLQuery = sgSQLQuery & "soeMergeFileFormat = '" & gFixQuote(tlSOE.sMergeFileFormat) & "', "
            sgSQLQuery = sgSQLQuery & "soeMergeFileExt = '" & gFixQuote(tlSOE.sMergeFileExt) & "', "
            sgSQLQuery = sgSQLQuery & "soeMergeStartTime = '" & Format$(tlSOE.sMergeStartTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "soeMergeEndTime = '" & Format$(tlSOE.sMergeEndTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "soeMergeChkInterval = " & tlSOE.iMergeChkInterval & ", "
            sgSQLQuery = sgSQLQuery & "soeMergeStopFlag = '" & tlSOE.sMergeStopFlag & "', "
            sgSQLQuery = sgSQLQuery & "soeAlertInterval = " & tlSOE.iAlertInterval & ", "
            sgSQLQuery = sgSQLQuery & "soeSchAutoGenSeq = '" & gFixQuote(tlSOE.sSchAutoGenSeq) & "', "
            sgSQLQuery = sgSQLQuery & "soeMinEventID = " & tlSOE.lMinEventID & ", "
            sgSQLQuery = sgSQLQuery & "soeMaxEventID = " & tlSOE.lMaxEventID & ", "
            sgSQLQuery = sgSQLQuery & "soeCurrEventID = " & tlSOE.lCurrEventID & ", "
            sgSQLQuery = sgSQLQuery & "soeNoDaysRetainPW = " & tlSOE.iNoDaysRetainPW & ", "
            sgSQLQuery = sgSQLQuery & "soeVersion = " & tlOldSOE.iVersion + 1 & ", "
            sgSQLQuery = sgSQLQuery & "soeOrigSoeCode = " & tlSOE.iOrigSoeCode & ", "
            sgSQLQuery = sgSQLQuery & "soeCurrent = " & "'Y'" & ", "
            sgSQLQuery = sgSQLQuery & "soeEnteredDate = '" & Format$(slNowDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "soeEnteredTime = '" & Format$(slNowTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "soeUieCode = " & tgUIE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "soeSpotItemIDWindow = " & tlSOE.iSpotItemIDWindow & ", "
            sgSQLQuery = sgSQLQuery & "soeTimeTolerance = " & tlSOE.lTimeTolerance & ", "
            sgSQLQuery = sgSQLQuery & "soeLengthTolerance = " & tlSOE.lLengthTolerance & ", "
            sgSQLQuery = sgSQLQuery & "soeMatchATNotB = " & "'" & gFixQuote(tlSOE.sMatchATNotB) & "', "
            sgSQLQuery = sgSQLQuery & "soeMatchATBNotI = " & "'" & gFixQuote(tlSOE.sMatchATBNotI) & "', "
            sgSQLQuery = sgSQLQuery & "soeMatchANotT = " & "'" & gFixQuote(tlSOE.sMatchANotT) & "', "
            sgSQLQuery = sgSQLQuery & "soeMatchBNotT = " & "'" & gFixQuote(tlSOE.sMatchBNotT) & "', "
            sgSQLQuery = sgSQLQuery & "soeSchAutoGenSeqTst = '" & gFixQuote(tlSOE.sSchAutoGenSeqTst) & "', "
            sgSQLQuery = sgSQLQuery & "soeMergeStopFlagTst = '" & tlSOE.sMergeStopFlagTst & "', "
            sgSQLQuery = sgSQLQuery & "soeUnused = '" & "" & "'"
            sgSQLQuery = sgSQLQuery & " WHERE soeCode = " & tlSOE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_SOE_SiteOption = True
            Exit Function
        Else
            gPutUpdate_SOE_SiteOption = False
            Exit Function
        End If
    Else
        gPutUpdate_SOE_SiteOption = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_SOE_SiteOption = False
    Exit Function

End Function


Public Function gPutUpdate_TSE_TemplateSchd(ilUpdateType As Integer, llOldDheCode As Long, tlTSE As TSE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE
'   tlTSE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldTSE As TSE
    
    On Error GoTo ErrHand
    
    
    ilRet = gGetRec_TSE_TemplateSchd(tlTSE.lCode, slForm_Module, tlOldTSE)
    If ilRet Then
        
        tlOldTSE.lCode = 0
        tlOldTSE.sCurrent = "N"
        tlOldTSE.lDheCode = llOldDheCode
        ilRet = gPutInsert_TSE_TemplateSchd(1, tlOldTSE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlTSE.iVersion, "TSE", CLng(tlOldTSE.lCode), CLng(tlTSE.lCode), CLng(tlTSE.lOrigTseCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_TSE_TemplateSchd = False
                Exit Function
            End If
                    
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update TSE_Template_Schd Set "
            sgSQLQuery = sgSQLQuery & "tseCode = " & tlTSE.lCode & ", "
            sgSQLQuery = sgSQLQuery & "tseDheCode = " & tlTSE.lDheCode & ", "
            sgSQLQuery = sgSQLQuery & "tseBdeCode = " & tlTSE.iBdeCode & ", "
            sgSQLQuery = sgSQLQuery & "tseLogDate = '" & Format$(tlTSE.sLogDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "tseStartTime = '" & Format$(tlTSE.sStartTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "tseDescription = '" & gFixQuote(tlTSE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "tseState = '" & gFixQuote(tlTSE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "tseCteCode = " & tlTSE.lCteCode & ", "
            sgSQLQuery = sgSQLQuery & "tseVersion = " & tlTSE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "tseOrigTseCode = " & tlTSE.lOrigTseCode & ", "
            sgSQLQuery = sgSQLQuery & "tseCurrent = '" & gFixQuote(tlTSE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "tseEnteredDate = '" & Format$(tlTSE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "tseEnteredTime = '" & Format$(tlTSE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "tseUieCode = " & tlTSE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "tseUnused = '" & gFixQuote(tlTSE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE TSECode = " & tlTSE.lCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_TSE_TemplateSchd = True
            Exit Function
        Else
            gPutUpdate_TSE_TemplateSchd = False
            Exit Function
        End If
    Else
        gPutUpdate_TSE_TemplateSchd = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_TSE_TemplateSchd = False
    Exit Function

End Function
Public Function gPutUpdate_TTE_TimeType(ilUpdateType As Integer, tlTTE As TTE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Only Used Flag
'   tlTTE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldTTE As TTE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        sgSQLQuery = "Update TTE_Time_Type Set "
        sgSQLQuery = sgSQLQuery & "tteUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE tteCode = " & tlTTE.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_TTE_TimeType = True
        Exit Function
    End If
        
    ilRet = gGetRec_TTE_TimeType(tlTTE.iCode, slForm_Module, tlOldTTE)
    If ilRet Then
        
        tlOldTTE.iCode = 0
        tlOldTTE.sCurrent = "N"
        ilRet = gPutInsert_TTE_TimeType(1, tlOldTTE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlTTE.iVersion, "TTE", CLng(tlOldTTE.iCode), CLng(tlTTE.iCode), CLng(tlTTE.iOrigTteCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_TTE_TimeType = False
                Exit Function
            End If
            
        
            slNowDate = Format(gNow(), sgShowDateForm)
            slNowTime = Format(gNow(), sgShowTimeWSecForm)
            sgSQLQuery = "Update TTE_Time_Type Set "
            sgSQLQuery = sgSQLQuery & "tteCode = " & tlTTE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "tteType = '" & gFixQuote(tlTTE.sType) & "', "
            sgSQLQuery = sgSQLQuery & "tteName = '" & gFixQuote(tlTTE.sName) & "', "
            sgSQLQuery = sgSQLQuery & "tteDescription = '" & gFixQuote(tlTTE.sDescription) & "', "
            sgSQLQuery = sgSQLQuery & "tteState = '" & gFixQuote(tlTTE.sState) & "', "
            sgSQLQuery = sgSQLQuery & "tteUsedFlag = '" & gFixQuote(tlTTE.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "tteVersion = " & tlTTE.iVersion & ", "
            sgSQLQuery = sgSQLQuery & "tteOrigTteCode = " & tlTTE.iOrigTteCode & ", "
            sgSQLQuery = sgSQLQuery & "tteCurrent = '" & gFixQuote(tlTTE.sCurrent) & "', "
            sgSQLQuery = sgSQLQuery & "tteEnteredDate = '" & Format$(tlTTE.sEnteredDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "tteEnteredTime = '" & Format$(tlTTE.sEnteredTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "tteUieCode = " & tlTTE.iUieCode & ", "
            sgSQLQuery = sgSQLQuery & "tteUnused = '" & gFixQuote(tlTTE.sUnused) & "' "
            sgSQLQuery = sgSQLQuery & " WHERE tteCode = " & tlTTE.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_TTE_TimeType = True
            Exit Function
        Else
            gPutUpdate_TTE_TimeType = False
            Exit Function
        End If
    Else
        gPutUpdate_TTE_TimeType = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_TTE_TimeType = False
    Exit Function

End Function
Public Function gPutUpdate_TTE_UsedFlag(ilCode As Integer, tlCurrTTE() As TTE) As Integer
    Dim ilTTE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilTTE = 0 To UBound(tlCurrTTE) - 1 Step 1
            If ilCode = tlCurrTTE(ilTTE).iCode Then
                If Trim$(tlCurrTTE(ilTTE).sUsedFlag) <> "Y" Then
                    tgTTE.iCode = ilCode
                    ilRet = gPutUpdate_TTE_TimeType(2, tgTTE, "Update Used Flag in TTE")
                    tlCurrTTE(ilTTE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilTTE
    End If
    gPutUpdate_TTE_UsedFlag = ilRet
End Function

Public Function gPutUpdate_UIE_UserInfo(ilUpdateType As Integer, tlUie As UIE, slForm_Module As String) As Integer
'
'   ilUpdateType (I)- 0=Standard; 1=Standard plus AIE; 2=Used Flag and sign on Date
'   tlUIE(I)- Updated record to be placed into file
'
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim tlOldUIE As UIE
    
    On Error GoTo ErrHand
    
    If ilUpdateType = 2 Then
        slNowDate = Format(Now, sgShowDateForm)
        slNowTime = Format(Now, sgShowTimeWSecForm)
        sgSQLQuery = "UPDATE UIE_User_Info SET "
        sgSQLQuery = sgSQLQuery & "uieLastSignOnDate = '" & Format$(tlUie.sLastSignOnDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "uieLastSignOnTime = '" & Format$(tlUie.sLastSignOnTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & "uieUsedFlag = 'Y'"
        sgSQLQuery = sgSQLQuery & " WHERE uieCode = " & tlUie.iCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        gPutUpdate_UIE_UserInfo = True
        Exit Function
    End If
    
    ilRet = gGetRec_UIE_UserInfo(tlUie.iCode, slForm_Module, tlOldUIE)
    If ilRet Then
        
        tlOldUIE.iCode = 0
        tlOldUIE.sCurrent = "N"
        ilRet = gPutInsert_UIE_UserInfo(1, tlOldUIE, slForm_Module)
        If ilRet Then
        
            ilRet = gUpdateAIE(ilUpdateType, tlUie.iVersion, "UIE", CLng(tlOldUIE.iCode), CLng(tlUie.iCode), CLng(tlUie.iOrigUieCode), slForm_Module)
            If Not ilRet Then
                gPutUpdate_UIE_UserInfo = False
                Exit Function
            End If
        
            sgSQLQuery = "UPDATE UTE_User_Tasks SET "
            sgSQLQuery = sgSQLQuery & "uteUieCode = " & tlOldUIE.iCode
            sgSQLQuery = sgSQLQuery & " WHERE uteUieCode = " & tlUie.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            
            slNowDate = Format(Now, sgShowDateForm)
            slNowTime = Format(Now, sgShowTimeWSecForm)
            sgSQLQuery = "UPDATE UIE_User_Info SET "
            sgSQLQuery = sgSQLQuery & "uieSignOnName = '" & gFixQuote(tlUie.sSignOnName) & "', "
            sgSQLQuery = sgSQLQuery & "uiePassword = '" & gFixQuote(tlUie.sPassword) & "', "
            sgSQLQuery = sgSQLQuery & "uieLastDatePWSet = '" & Format$(tlUie.sLastDatePWSet, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "uieShowName = '" & gFixQuote(tlUie.sShowName) & "', "
            sgSQLQuery = sgSQLQuery & "uieState = '" & gFixQuote(tlUie.sState) & "', "
            sgSQLQuery = sgSQLQuery & "uieEMail = '" & gFixQuote(tlUie.sEMail) & "', "
            sgSQLQuery = sgSQLQuery & "uieLastSignOnDate = '" & Format$(tlUie.sLastSignOnDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "uieLastSignOnTime = '" & Format$(tlUie.sLastSignOnTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "uieUsedFlag = '" & gFixQuote(tlUie.sUsedFlag) & "', "
            sgSQLQuery = sgSQLQuery & "uieVersion = " & tlOldUIE.iVersion + 1 & ", "
            sgSQLQuery = sgSQLQuery & "uieOrigUieCode = " & tlUie.iOrigUieCode & ", "
            sgSQLQuery = sgSQLQuery & "uieCurrent = " & "'Y'" & ", "
            sgSQLQuery = sgSQLQuery & "uieEnteredDate = '" & Format$(slNowDate, sgSQLDateForm) & "', "
            sgSQLQuery = sgSQLQuery & "uieEnteredTime = '" & Format$(slNowTime, sgSQLTimeForm) & "', "
            sgSQLQuery = sgSQLQuery & "uieUieCode = " & tgUIE.iCode & ", "
            sgSQLQuery = sgSQLQuery & "uieUnused = '" & "" & "'"
            sgSQLQuery = sgSQLQuery & " WHERE uieCode = " & tlUie.iCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            gPutUpdate_UIE_UserInfo = True
            Exit Function
        Else
            gPutUpdate_UIE_UserInfo = False
            Exit Function
        End If
    Else
        gPutUpdate_UIE_UserInfo = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_UIE_UserInfo = False
    Exit Function

End Function



Public Function gUpdateAIE(ilUpdateType As Integer, ilVersion As Integer, slFileName As String, llFromCode As Long, llToCode As Long, llOrigCode As Long, slForm_Module As String) As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim ilRet As Integer
    
    '5/9/11: Remove retaining who altered the information
    gUpdateAIE = True
    Exit Function

    If ilUpdateType <> 1 Then
        gUpdateAIE = True
        Exit Function
    End If
    On Error GoTo ErrHand
    If ilVersion > 1 Then
        sgSQLQuery = "UPDATE AIE_Active_Info SET "
        sgSQLQuery = sgSQLQuery & "aieToFileCode = " & llFromCode
        sgSQLQuery = sgSQLQuery & " WHERE aieToFileCode = " & llToCode & " and " & "aieRefFileName = '" & slFileName & "'"
        cnn.Execute sgSQLQuery    ', rdExecDirect
    End If
    slNowDate = Format(Now, sgShowDateForm)
    slNowTime = Format(Now, sgShowTimeWSecForm)
    tmAIE.lCode = 0
    tmAIE.iUieCode = tgUIE.iCode
    tmAIE.sEnteredDate = Format(slNowDate, sgShowDateForm)
    tmAIE.sEnteredTime = Format(slNowTime, sgShowTimeWSecForm)
    tmAIE.sRefFileName = slFileName
    tmAIE.lToFileCode = llToCode
    tmAIE.lFromFileCode = llFromCode
    tmAIE.lOrigFileCode = llOrigCode
    tmAIE.sUnused = ""
    ilRet = gPutInsert_AIE_ActiveInfo(tmAIE, slForm_Module)
    gUpdateAIE = ilRet
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gUpdateAIE = False
    Exit Function
End Function

Public Function gPutUpdate_ANE_UsedFlag(ilCode As Integer, tlCurrANE() As ANE) As Integer
    Dim ilANE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilANE = 0 To UBound(tlCurrANE) - 1 Step 1
            If ilCode = tlCurrANE(ilANE).iCode Then
                If Trim$(tlCurrANE(ilANE).sUsedFlag) <> "Y" Then
                    tgANE.iCode = ilCode
                    ilRet = gPutUpdate_ANE_AudioName(2, tgANE, "Update Used Flag in ANE")
                    tlCurrANE(ilANE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilANE
    End If
    gPutUpdate_ANE_UsedFlag = ilRet
End Function
Public Function gPutUpdate_ASE_UsedFlag(ilCode As Integer, tlCurrASE() As ASE) As Integer
    Dim ilASE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilASE = 0 To UBound(tlCurrASE) - 1 Step 1
            If ilCode = tlCurrASE(ilASE).iCode Then
                If Trim$(tlCurrASE(ilASE).sUsedFlag) <> "Y" Then
                    tgASE.iCode = ilCode
                    ilRet = gPutUpdate_ASE_AudioSource(2, tgASE, "Update Used Flag in ASE")
                    tlCurrASE(ilASE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilASE
    End If
    gPutUpdate_ASE_UsedFlag = ilRet
End Function
Public Function gPutUpdate_ATE_UsedFlag(ilCode As Integer, tlCurrATE() As ATE) As Integer
    Dim ilATE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilATE = 0 To UBound(tlCurrATE) - 1 Step 1
            If ilCode = tlCurrATE(ilATE).iCode Then
                If Trim$(tlCurrATE(ilATE).sUsedFlag) <> "Y" Then
                    tgATE.iCode = ilCode
                    ilRet = gPutUpdate_ATE_AudioType(2, tgATE, "Update Used Flag in ATE")
                    tlCurrATE(ilATE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilATE
    End If
    gPutUpdate_ATE_UsedFlag = ilRet
End Function
Public Function gPutUpdate_BGE_UsedFlag(ilCode As Integer, tlCurrBGE() As BGE) As Integer
    Dim ilBGE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilBGE = 0 To UBound(tlCurrBGE) - 1 Step 1
            If ilCode = tlCurrBGE(ilBGE).iCode Then
                If Trim$(tlCurrBGE(ilBGE).sUsedFlag) <> "Y" Then
                    tgBGE.iCode = ilCode
                    ilRet = gPutUpdate_BGE_BusGroup(2, tgBGE, "Update Used Flag in BGE")
                    tlCurrBGE(ilBGE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilBGE
    End If
    gPutUpdate_BGE_UsedFlag = ilRet
End Function
Public Function gPutUpdate_BDE_UsedFlag(ilCode As Integer, tlCurrBDE() As BDE) As Integer
    Dim ilBDE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilBDE = 0 To UBound(tlCurrBDE) - 1 Step 1
            If ilCode = tlCurrBDE(ilBDE).iCode Then
                If Trim$(tlCurrBDE(ilBDE).sUsedFlag) <> "Y" Then
                    tgBDE.iCode = ilCode
                    ilRet = gPutUpdate_BDE_BusDefinition(2, tgBDE, "Update Used Flag in BDE")
                    tlCurrBDE(ilBDE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilBDE
    End If
    gPutUpdate_BDE_UsedFlag = ilRet
End Function

Public Function gPutUpdate_CCE_UsedFlag(ilCode As Integer, tlCurrCCE() As CCE) As Integer
    Dim ilCCE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If ilCode > 0 Then
        For ilCCE = 0 To UBound(tlCurrCCE) - 1 Step 1
            If ilCode = tlCurrCCE(ilCCE).iCode Then
                If Trim$(tlCurrCCE(ilCCE).sUsedFlag) <> "Y" Then
                    tgCCE.iCode = ilCode
                    ilRet = gPutUpdate_CCE_ControlChar(2, tgCCE, "Update Used Flag in CCE")
                    tlCurrCCE(ilCCE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilCCE
    End If
    gPutUpdate_CCE_UsedFlag = ilRet
End Function

Public Function gPutUpdate_DNE_UsedFlag(llCode As Long, tlCurrDNE() As DNE) As Integer
    Dim ilDNE As Integer
    Dim ilRet As Integer
    
    ilRet = True
    If llCode > 0 Then
        For ilDNE = 0 To UBound(tlCurrDNE) - 1 Step 1
            If llCode = tlCurrDNE(ilDNE).lCode Then
                If Trim$(tlCurrDNE(ilDNE).sUsedFlag) <> "Y" Then
                    tgDNE.lCode = llCode
                    ilRet = gPutUpdate_DNE_DayName(2, tgDNE, "Update Used Flag in DNE")
                    tlCurrDNE(ilDNE).sUsedFlag = "Y"
                End If
                Exit For
            End If
        Next ilDNE
    End If
    gPutUpdate_DNE_UsedFlag = ilRet
End Function


Public Function gPutDelete_ASE_AudioSource(ilCode As Integer, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'ASE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM ASE_Audio_Source WHERE aseOrigAseCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_ASE_AudioSource = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_ASE_AudioSource = False
    Exit Function
    
End Function

Public Function gPutDelete_AAE_As_Aired(llSheCode As Long, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AAE_As_Aired Where aaeSheCode = " & llSheCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_AAE_As_Aired = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_AAE_As_Aired = False
    Exit Function
    
End Function

Public Function gPutDelete_AAE_As_AiredByBus(llSheCode As Long, slBus As String, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AAE_As_Aired Where aaeSheCode = " & llSheCode & " AND aaeBusName = '" & slBus & "'"
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_AAE_As_AiredByBus = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_AAE_As_AiredByBus = False
    Exit Function
    
End Function

Public Function gPutDelete_AEE_AutoEquip(ilCode As Integer, slForm_Module As String) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'AEE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'ACE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'ADE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'AFE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'APE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "SELECT aeeCode FROM AEE_Auto_Equip WHERE aeeOrigAeeCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    While Not rst.EOF
        sgSQLQuery = "DELETE FROM ACE_Auto_Contact WHERE aceAeeCode = " & rst!aeeCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        sgSQLQuery = "DELETE FROM ADE_Auto_Data_Flags WHERE adeAeeCode = " & rst!aeeCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        sgSQLQuery = "DELETE FROM AFE_Auto_Format WHERE afeAeeCode = " & rst!aeeCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        sgSQLQuery = "DELETE FROM APE_Auto_Path WHERE apeAeeCode = " & rst!aeeCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        rst.MoveNext
    Wend
    rst.Close
    sgSQLQuery = "DELETE FROM AEE_Auto_Equip WHERE aeeOrigAeeCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_AEE_AutoEquip = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_AEE_AutoEquip = False
    Exit Function
End Function
Public Function gPutDelete_ANE_AudioName(ilCode As Integer, slForm_Module As String) As Integer
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'ANE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM ANE_Audio_Name WHERE aneOrigAneCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_ANE_AudioName = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_ANE_AudioName = False
    Exit Function
    
End Function

Public Function gPutDelete_ATE_AudioType(ilCode As Integer, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'ATE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM ATE_Audio_Type WHERE ateOrigAteCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_ATE_AudioType = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_ATE_AudioType = False
    Exit Function
    
End Function

Public Function gPutDelete_BDE_BusDefinition(ilCode As Integer, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'BDE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM BSE_Bus_Sel_Group WHERE bseBdeCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM BDE_Bus_Definition WHERE bdeOrigBdeCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_BDE_BusDefinition = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_BDE_BusDefinition = False
    Exit Function
    
End Function

Public Function gPutDelete_BGE_BusGroup(ilCode As Integer, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'BGE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM BGE_Bus_Group WHERE bgeOrigBgeCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_BGE_BusGroup = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_BGE_BusGroup = False
    Exit Function
    
End Function
Public Function gPutDelete_CCE_ControlChar(ilCode As Integer, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'CCE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM CCE_Control_Char WHERE cceOrigCceCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_CCE_ControlChar = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_CCE_ControlChar = False
    Exit Function
    
End Function

Public Function gPutDelete_CEE_Conflict_Events(llGenDate As Long, llGenTime As Long, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM CEE_Conflict_Events WHERE ceeGenDate = " & llGenDate & " AND ceeGenTime = " & llGenTime
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_CEE_Conflict_Events = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_CEE_Conflict_Events = False
    Exit Function
    
End Function
Public Function gPutDelete_CME_Conflict_Master(slSource As String, llSHEDHECode As Long, llSEECode As Long, llDeleteDate As Long, slForm_Module As String, hlCME As Integer) As Integer
    Dim ilUpper As Integer
    Dim ilCME As Integer
    Dim ilRet As Integer
    Dim tlSvCME As CME
    ReDim tlCME(0 To 0) As CME
    Dim rst As ADODB.Recordset
    On Error GoTo ErrHand
    
    gPutDelete_CME_Conflict_Master = True
    Exit Function
    
    sgSQLQuery = ""
    If (slSource = "") And (llSHEDHECode <= 0) And (llSEECode <= 0) And (llDeleteDate <= 0) Then
        sgSQLQuery = "DELETE FROM CME_Conflict_Master"
    Else
        If slSource <> "" Then
            sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeSource = '" & slSource & "'"
            If llSHEDHECode > 0 Then
                sgSQLQuery = sgSQLQuery & " AND cmeSHEDHECode = " & llSHEDHECode
                If llSEECode > 0 Then
                    sgSQLQuery = sgSQLQuery & " AND cmeSEECode = " & llSEECode
                End If
            Else
                If llSEECode > 0 Then
                    sgSQLQuery = sgSQLQuery & " AND cmeSEECode = " & llSEECode
                End If
            End If
        Else
            If llSHEDHECode > 0 Then
                sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeSHEDHECode = " & llSHEDHECode
                If llSEECode > 0 Then
                    sgSQLQuery = sgSQLQuery & " AND cmeSEECode = " & llSEECode
                End If
            Else
                If llSEECode > 0 Then
                    sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeSEECode = " & llSEECode
                End If
            End If
        End If
        If (llDeleteDate > 0) And ((slSource = "S") Or (slSource = "T")) Then
            sgSQLQuery = sgSQLQuery & " AND cmeStartDate = " & llDeleteDate
        ElseIf (llDeleteDate > 0) And ((slSource = "L") Or (slSource = "")) Then
            'Remove date from source = "L"
            sgSQLQuery = "SELECT * FROM CME_Conflict_Master WHERE cmeSource = 'L' AND cmeStartDate <= " & llDeleteDate & " AND cmeEndDate >= " & llDeleteDate
            Set rst = cnn.Execute(sgSQLQuery)
            While Not rst.EOF
                ilUpper = UBound(tlCME)
                tlCME(ilUpper).lCode = rst!cmeCode
                tlCME(ilUpper).sSource = rst!cmeSource
                tlCME(ilUpper).lSHEDHECode = rst!cmeSHEDHECode
                tlCME(ilUpper).lDseCode = rst!cmeDSECode
                tlCME(ilUpper).lDeeCode = rst!cmeDEECode
                tlCME(ilUpper).lSeeCode = rst!cmeSEECode
                tlCME(ilUpper).sEvtType = rst!cmeEvtType
                tlCME(ilUpper).iBdeCode = rst!cmeBDECode
                tlCME(ilUpper).iANECode = rst!cmeANECode
                tlCME(ilUpper).lStartDate = rst!cmeStartDate
                tlCME(ilUpper).lEndDate = rst!cmeEndDate
                tlCME(ilUpper).sDay = rst!cmeDay
                tlCME(ilUpper).lStartTime = rst!cmeStartTime
                tlCME(ilUpper).lEndTime = rst!cmeEndTime
                tlCME(ilUpper).sItemID = rst!cmeItemID
                tlCME(ilUpper).sXMidNight = rst!cmeXMidNight
                tlCME(ilUpper).sUnused = rst!cmeUnused
                ReDim Preserve tlCME(0 To ilUpper + 1) As CME
                rst.MoveNext
            Wend
            rst.Close
            For ilCME = 0 To UBound(tlCME) - 1 Step 1
                LSet tlSvCME = tlCME(ilCME)
                If tlCME(ilCME).lStartDate < llDeleteDate Then
                    tlCME(ilCME).lEndDate = llDeleteDate - 1
                    If tlCME(ilCME).lStartDate <= tlCME(ilCME).lEndDate Then
                        tlCME(ilCME).lCode = 0
                        ilRet = gPutInsert_CME_Conflict_Master(tlCME(ilCME), slForm_Module, hlCME)
                    End If
                    LSet tlCME(ilCME) = tlSvCME
                End If
'                tlCME(ilCME).lStartDate = tlCME(ilCME).lStartDate + 1
'                If tlCME(ilCME).lStartDate <= tlCME(ilCME).lEndDate Then
'                    tlCME(ilCME).lCode = 0
'                    ilRet = gPutInsert_CME_Conflict_Master(tlCME(ilCME), slForm_Module)
'                End If
                sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeCode = " & tlSvCME.lCode
                cnn.Execute sgSQLQuery    ', rdExecDirect
            Next ilCME
            If slSource = "" Then
                sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeStartDate = " & llDeleteDate
                sgSQLQuery = sgSQLQuery & " AND (cmeSource = 'S' OR cmeSource = 'T')"
            Else
                gPutDelete_CME_Conflict_Master = True
                Exit Function
            End If
        End If
    End If
    If sgSQLQuery <> "" Then
        cnn.Execute sgSQLQuery    ', rdExecDirect
    End If
    gPutDelete_CME_Conflict_Master = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_CME_Conflict_Master = False
    Exit Function
    
End Function

Public Function gPutDelete_CTE_CommtsTitle(llCode As Long, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    'sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'CTE' AND aieOrigFileCode = " & llCode
    'cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM CTE_Commts_And_Title WHERE cteCode = " & llCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_CTE_CommtsTitle = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_CTE_CommtsTitle = False
    Exit Function
    
End Function



Public Function gPutDelete_DHE_DayHeaderInfo(llCode As Long, slForm_Module As String) As Integer
    Dim rst As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'DHE' AND aieOrigFileCode = " & llCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'DEE' AND aieOrigFileCode = " & llCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'DBE' AND aieOrigFileCode = " & llCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    'sgSQLQuery = "SELECT dheCode, dheType FROM DHE_Day_Header_Info WHERE dheOrigDheCode = " & llCode
    sgSQLQuery = "SELECT dheCode, dheType FROM DHE_Day_Header_Info WHERE dheCode = " & llCode
    Set rst = cnn.Execute(sgSQLQuery)
    While Not rst.EOF
        sgSQLQuery = "SELECT deeCode, dee1CteCode FROM DEE_Day_Event_Info WHERE deeDheCode = " & rst!dheCode
        Set rst2 = cnn.Execute(sgSQLQuery)
        While Not rst2.EOF
            sgSQLQuery = "DELETE FROM EBE_Event_Bus_Sel WHERE ebeDeeCode = " & rst2!deeCode
            cnn.Execute sgSQLQuery    ', rdExecDirect
            '1/8/12: Retain Comment and let gCommentDelete handle the removal
            'sgSQLQuery = "DELETE FROM CTE_Commts_And_Title WHERE cteCode = " & rst2!dee1CteCode
            'cnn.Execute sgSQLQuery    ', rdExecDirect
            'sgSQLQuery = "DELETE FROM CTE_Commts_And_Title WHERE cteCode = " & rst2!dee2CteCode
            'cnn.Execute sgSQLQuery    ', rdExecDirect
            rst2.MoveNext
        Wend
        sgSQLQuery = "DELETE FROM DEE_Day_Event_Info WHERE deeDheCode = " & rst!dheCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        sgSQLQuery = "DELETE FROM DBE_Day_Bus_Sel WHERE dbeDheCode = " & rst!dheCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        'sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeSHEDHECode = " & rst!dheCode & " AND cmeSource = '" & rst!dheType & "'"
        'cnn.Execute sgSQLQuery    ', rdExecDirect
        sgSQLQuery = "DELETE FROM TSE_Template_Schd WHERE tseDheCode = " & rst!dheCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        rst.MoveNext
    Wend
    rst.Close
    sgSQLQuery = "DELETE FROM DHE_Day_Header_Info WHERE dheCode = " & llCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_DHE_DayHeaderInfo = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_DHE_DayHeaderInfo = False
    Exit Function
End Function



Public Function gPutDelete_ETE_EventType(ilCode As Integer, slForm_Module As String) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'ETE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'EPE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "SELECT eteCode FROM ETE_Event_Type WHERE eteOrigEteCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    While Not rst.EOF
        sgSQLQuery = "DELETE FROM EPE_Event_Properties WHERE epeEteCode = " & rst!eteCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        rst.MoveNext
    Wend
    rst.Close
    sgSQLQuery = "DELETE FROM ETE_Event_Type WHERE eteOrigEteCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_ETE_EventType = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_ETE_EventType = False
    Exit Function
End Function

Public Function gPutDelete_FNE_FollowName(ilCode As Integer, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'FNE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM FNE_Follow_Name WHERE fneOrigFneCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_FNE_FollowName = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_FNE_FollowName = False
    Exit Function
    
End Function

Public Function gPutDelete_MTE_MaterialType(ilCode As Integer, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'MTE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM MTE_Material_Type WHERE mteOrigMteCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_MTE_MaterialType = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_MTE_MaterialType = False
    Exit Function
    
End Function

Public Function gPutDelete_NNE_NetcueName(ilCode As Integer, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'NNE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM NNE_Netcue_Name WHERE nneOrigNneCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_NNE_NetcueName = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_NNE_NetcueName = False
    Exit Function
    
End Function

Public Function gPutDelete_RLE_Record_Locks(llCode As Long, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM RLE_Record_Locks WHERE rleCode = " & llCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_RLE_Record_Locks = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_RLE_Record_Locks = False
    Exit Function
    
End Function

Public Function gPutDelete_RNE_RelayName(ilCode As Integer, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'RNE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM RNE_Relay_Name WHERE rneOrigRneCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_RNE_RelayName = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_RNE_RelayName = False
    Exit Function
    
End Function

Public Function gPutDelete_SCE_SilenceChar(ilCode As Integer, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'SCE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM SCE_Silence_Char WHERE sceOrigSceCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_SCE_SilenceChar = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_SCE_SilenceChar = False
    Exit Function
End Function

Public Function gPutDelete_SEE_Schedule_Events(llCode As Long, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'SEE' AND aieOrigFileCode = " & llCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM SEE_Schedule_Events WHERE seeCode = " & llCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_SEE_Schedule_Events = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_SEE_Schedule_Events = False
    Exit Function
End Function

Public Function gPutDelete_TTE_TimeType(ilCode As Integer, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'TTE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM TTE_Time_Type WHERE tteOrigTteCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_TTE_TimeType = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_TTE_TimeType = False
    Exit Function
End Function

Public Function gPutDelete_UIE_UserInfo(ilCode As Integer, slForm_Module As String) As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'UIE' AND aieOrigFileCode = " & CLng(ilCode)
    cnn.Execute sgSQLQuery    ', rdExecDirect
    'sgSQLQuery = "DELETE FROM UTE_User_Tasks WHERE uteUieCode = " & ilCode
    'cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "SELECT uieCode FROM UIE_User_Info WHERE uieOrigUieCode = " & ilCode
    Set rst = cnn.Execute(sgSQLQuery)
    While Not rst.EOF
        sgSQLQuery = "DELETE FROM UTE_User_Tasks WHERE uteUieCode = " & rst!uieCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        rst.MoveNext
    Wend
    rst.Close
    sgSQLQuery = "DELETE FROM UIE_User_Info WHERE uieOrigUieCode = " & ilCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_UIE_UserInfo = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_UIE_UserInfo = False
    Exit Function
End Function

Public Function gPutInsert_DBE_DayBusSel(tlDBE As DBE, slForm_Module As String) As Integer
'
'   tlDBE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    llLastCode = 0
    llInitCode = tlDBE.lCode
    Do
        If tlDBE.lCode <= 0 Then
            sgSQLQuery = "Select MAX(dbeCode) from DBE_Day_Bus_Sel"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlDBE.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlDBE.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlDBE.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlDBE.lCode
        sgSQLQuery = "Insert Into DBE_Day_Bus_Sel ( "
        sgSQLQuery = sgSQLQuery & "dbeCode, "
        sgSQLQuery = sgSQLQuery & "dbeType, "
        sgSQLQuery = sgSQLQuery & "dbeDheCode, "
        sgSQLQuery = sgSQLQuery & "dbeBdeCode, "
        sgSQLQuery = sgSQLQuery & "dbeBgeCode, "
        sgSQLQuery = sgSQLQuery & "dbeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlDBE.lCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDBE.sType) & "', "
        sgSQLQuery = sgSQLQuery & tlDBE.lDheCode & ", "
        sgSQLQuery = sgSQLQuery & tlDBE.iBdeCode & ", "
        sgSQLQuery = sgSQLQuery & tlDBE.iBgeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlDBE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_DBE_DayBusSel = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlDBE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_DBE_DayBusSel = False
    Exit Function
End Function

Public Function gPutDelete_DNE_DayName(llCode As Long, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'DNE' AND aieOrigFileCode = " & llCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM DNE_Day_Name WHERE dneOrigDneCode = " & llCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_DNE_DayName = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_DNE_DayName = False
    Exit Function
    
End Function

Public Function gPutDelete_DSE_DaySubName(llCode As Long, slForm_Module As String) As Integer
    
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM AIE_Active_Info WHERE aieRefFileName = 'DSE' AND aieOrigFileCode = " & llCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM DSE_Day_SubName WHERE dseOrigDseCode = " & llCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutDelete_DSE_DaySubName = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutDelete_DSE_DaySubName = False
    Exit Function
    
End Function

Public Function gPutInsert_EBE_EventBusSel(tlEBE As EBE, slForm_Module As String) As Integer
'
'   tlEBE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    llLastCode = 0
    llInitCode = tlEBE.lCode
    Do
        If tlEBE.lCode <= 0 Then
            sgSQLQuery = "Select MAX(ebeCode) from EBE_Event_Bus_Sel"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlEBE.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlEBE.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlEBE.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlEBE.lCode
        sgSQLQuery = "Insert Into EBE_Event_Bus_Sel ( "
        sgSQLQuery = sgSQLQuery & "ebeCode, "
        sgSQLQuery = sgSQLQuery & "ebeDeeCode, "
        sgSQLQuery = sgSQLQuery & "ebeBdeCode, "
        sgSQLQuery = sgSQLQuery & "ebeUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlEBE.lCode & ", "
        sgSQLQuery = sgSQLQuery & tlEBE.lDeeCode & ", "
        sgSQLQuery = sgSQLQuery & tlEBE.iBdeCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlEBE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_EBE_EventBusSel = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlEBE.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_EBE_EventBusSel = False
    Exit Function
End Function

Public Function gPutUpdate_Library(tlDHE As DHE, slForm_Module As String, llNewAgedDHECode As Long) As Integer
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim slDEEStamp As String
    Dim tlDEE() As DEE
    Dim ilDEE As Integer
    Dim slDBEStamp As String
    Dim tlDBE() As DBE
    Dim ilDBE As Integer
    Dim slEBEStamp As String
    Dim tlEBE() As EBE
    Dim ilEBE As Integer
    
    ilRet = gPutUpdate_DHE_DayHeaderInfo(3, tlDHE, slForm_Module, llNewAgedDHECode)
    If ilRet Then
        
        ilRet = gGetRecs_DEE_DayEvent(slDEEStamp, tlDHE.lCode, slForm_Module, tlDEE())
        For ilDEE = 0 To UBound(tlDEE) - 1 Step 1
            ilRet = gGetRecs_EBE_EventBusSel(slEBEStamp, tlDEE(ilDEE).lCode, slForm_Module, tlEBE())
            tlDEE(ilDEE).lCode = 0
            tlDEE(ilDEE).lDheCode = llNewAgedDHECode
            ilRet = gPutInsert_DEE_DayEvent(tlDEE(ilDEE), slForm_Module)
            For ilEBE = 0 To UBound(tlEBE) - 1 Step 1
                tlEBE(ilEBE).lCode = 0
                tlEBE(ilEBE).lDeeCode = tlDEE(ilDEE).lCode
                ilRet = gPutInsert_EBE_EventBusSel(tlEBE(ilEBE), slForm_Module)
            Next ilEBE
        Next ilDEE
        ilRet = gGetRecs_DBE_DayBusSel(slDBEStamp, tlDHE.lCode, slForm_Module, tlDBE())
        For ilDBE = 0 To UBound(tlDBE) - 1 Step 1
            tlDBE(ilDBE).lCode = 0
            tlDBE(ilDBE).lDheCode = llNewAgedDHECode
            ilRet = gPutInsert_DBE_DayBusSel(tlDBE(ilDBE), slForm_Module)
        Next ilDBE
        gPutUpdate_Library = True
        Exit Function
    Else
        gPutUpdate_Library = False
        Exit Function
    End If
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_Library = False
    Exit Function

End Function


Public Function gPutUpdate_SEE_DEECode(llSEECode As Long, llDeeCode As Long, slForm_Module As String) As Integer
    Dim ilANE As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    If llSEECode > 0 Then
        sgSQLQuery = "Update SEE_Schedule_Events Set "
        sgSQLQuery = sgSQLQuery & "seeDeeCode = " & llDeeCode
        sgSQLQuery = sgSQLQuery & " WHERE seeCode = " & llSEECode
        cnn.Execute sgSQLQuery    ', rdExecDirect
    End If
    gPutUpdate_SEE_DEECode = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_SEE_DEECode = False
    Exit Function
End Function
Public Function gPutUpdate_DEE_CTECode(ilCmmtNo As Integer, llDeeCode As Long, llCteCode As Long, slForm_Module As String) As Integer
    Dim ilANE As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    If llDeeCode > 0 Then
        sgSQLQuery = "Update DEE_Day_Event_Info Set "
        If ilCmmtNo = 2 Then
            sgSQLQuery = sgSQLQuery & "dee2CteCode = " & llCteCode
        Else
            sgSQLQuery = sgSQLQuery & "dee1CteCode = " & llCteCode
        End If
        sgSQLQuery = sgSQLQuery & " WHERE deeCode = " & llDeeCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
    End If
    gPutUpdate_DEE_CTECode = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_DEE_CTECode = False
    Exit Function
End Function

Public Function gPutUpdate_SEE_DHEDEECode(llSEECode As Long, llDheCode As Long, llDeeCode As Long, slForm_Module As String) As Integer
    Dim ilANE As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    If llSEECode > 0 Then
        sgSQLQuery = "Update SEE_Schedule_Events Set "
        sgSQLQuery = sgSQLQuery & "seeDheCode = " & llDheCode & ", "
        sgSQLQuery = sgSQLQuery & "seeDeeCode = " & llDeeCode
        sgSQLQuery = sgSQLQuery & " WHERE seeCode = " & llSEECode
        cnn.Execute sgSQLQuery    ', rdExecDirect
    End If
    gPutUpdate_SEE_DHEDEECode = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_SEE_DHEDEECode = False
    Exit Function
End Function

Public Function gPutUpdate_SEE_Schedule_Events(tlSEE As SEE, slForm_Module As String) As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    If tlSEE.lCode > 0 Then
        sgSQLQuery = "Update SEE_Schedule_Events Set "
        'sgSQLQuery = sgSQLQuery & "seeCode = " & tlSEE.lCode & ", "
        sgSQLQuery = sgSQLQuery & "seeSheCode = " & tlSEE.lSheCode & ", "
        sgSQLQuery = sgSQLQuery & "seeAction = '" & gFixQuote(tlSEE.sAction) & "', "
        sgSQLQuery = sgSQLQuery & "seeDeeCode = " & tlSEE.lDeeCode & ", "
        sgSQLQuery = sgSQLQuery & "seeBdeCode = " & tlSEE.iBdeCode & ", "
        sgSQLQuery = sgSQLQuery & "seeBusCceCode = " & tlSEE.iBusCceCode & ", "
        sgSQLQuery = sgSQLQuery & "seeSchdType = '" & gFixQuote(tlSEE.sSchdType) & "', "
        sgSQLQuery = sgSQLQuery & "seeEteCode = " & tlSEE.iEteCode & ", "
        sgSQLQuery = sgSQLQuery & "seeTime = " & tlSEE.lTime & ", "
        sgSQLQuery = sgSQLQuery & "seeStartTteCode = " & tlSEE.iStartTteCode & ", "
        sgSQLQuery = sgSQLQuery & "seeFixedTime = '" & gFixQuote(tlSEE.sFixedTime) & "', "
        sgSQLQuery = sgSQLQuery & "seeEndTteCode = " & tlSEE.iEndTteCode & ", "
        sgSQLQuery = sgSQLQuery & "seeDuration = " & tlSEE.lDuration & ", "
        sgSQLQuery = sgSQLQuery & "seeMteCode = " & tlSEE.iMteCode & ", "
        sgSQLQuery = sgSQLQuery & "seeAudioAseCode = " & tlSEE.iAudioAseCode & ", "
        sgSQLQuery = sgSQLQuery & "seeAudioItemID = '" & gFixQuote(tlSEE.sAudioItemID) & "', "
        sgSQLQuery = sgSQLQuery & "seeAudioItemIDChk = '" & gFixQuote(tlSEE.sAudioItemIDChk) & "', "
        sgSQLQuery = sgSQLQuery & "seeAudioISCI = '" & gFixQuote(tlSEE.sAudioISCI) & "', "
        sgSQLQuery = sgSQLQuery & "seeAudioCceCode = " & tlSEE.iAudioCceCode & ", "
        sgSQLQuery = sgSQLQuery & "seeBkupAneCode = " & tlSEE.iBkupAneCode & ", "
        sgSQLQuery = sgSQLQuery & "seeBkupCceCode = " & tlSEE.iBkupCceCode & ", "
        sgSQLQuery = sgSQLQuery & "seeProtAneCode = " & tlSEE.iProtAneCode & ", "
        sgSQLQuery = sgSQLQuery & "seeProtItemID = '" & gFixQuote(tlSEE.sProtItemID) & "', "
        sgSQLQuery = sgSQLQuery & "seeProtItemIDChk = '" & gFixQuote(tlSEE.sProtItemIDChk) & "', "
        sgSQLQuery = sgSQLQuery & "seeProtISCI = '" & gFixQuote(tlSEE.sProtISCI) & "', "
        sgSQLQuery = sgSQLQuery & "seeProtCceCode = " & tlSEE.iProtCceCode & ", "
        sgSQLQuery = sgSQLQuery & "see1RneCode = " & tlSEE.i1RneCode & ", "
        sgSQLQuery = sgSQLQuery & "see2RneCode = " & tlSEE.i2RneCode & ", "
        sgSQLQuery = sgSQLQuery & "seeFneCode = " & tlSEE.iFneCode & ", "
        sgSQLQuery = sgSQLQuery & "seeSilenceTime = " & tlSEE.lSilenceTime & ", "
        sgSQLQuery = sgSQLQuery & "see1SceCode = " & tlSEE.i1SceCode & ", "
        sgSQLQuery = sgSQLQuery & "see2SceCode = " & tlSEE.i2SceCode & ", "
        sgSQLQuery = sgSQLQuery & "see3SceCode = " & tlSEE.i3SceCode & ", "
        sgSQLQuery = sgSQLQuery & "see4SceCode = " & tlSEE.i4SceCode & ", "
        sgSQLQuery = sgSQLQuery & "seeStartNneCode = " & tlSEE.iStartNneCode & ", "
        sgSQLQuery = sgSQLQuery & "seeEndNneCode = " & tlSEE.iEndNneCode & ", "
        sgSQLQuery = sgSQLQuery & "see1CteCode = " & tlSEE.l1CteCode & ", "
        sgSQLQuery = sgSQLQuery & "see2CteCode = " & tlSEE.l2CteCode & ", "
        sgSQLQuery = sgSQLQuery & "seeAreCode = " & tlSEE.lAreCode & ", "
        sgSQLQuery = sgSQLQuery & "seeSpotTime = " & tlSEE.lSpotTime & ", "
        sgSQLQuery = sgSQLQuery & "seeEventID = " & tlSEE.lEventID & ", "
        sgSQLQuery = sgSQLQuery & "seeAsAirStatus = '" & gFixQuote(tlSEE.sAsAirStatus) & "', "
        sgSQLQuery = sgSQLQuery & "seeSentStatus = '" & gFixQuote(tlSEE.sSentStatus) & "', "
        sgSQLQuery = sgSQLQuery & "seeSentDate = '" & gFixQuote(tlSEE.sSentDate) & "', "
        sgSQLQuery = sgSQLQuery & "seeIgnoreConflicts = '" & gFixQuote(tlSEE.sIgnoreConflicts) & "', "
        sgSQLQuery = sgSQLQuery & "seeDheCode = " & tlSEE.lDheCode & ", "
        sgSQLQuery = sgSQLQuery & "seeOrigDHECode = " & tlSEE.lOrigDHECode & ", "
        sgSQLQuery = sgSQLQuery & "seeInsertFlag = '" & gFixQuote(tlSEE.sInsertFlag) & "', "
        sgSQLQuery = sgSQLQuery & "seeABCFormat = '" & gFixQuote(tlSEE.sABCFormat) & "', "
        sgSQLQuery = sgSQLQuery & "seeABCPgmCode = '" & gFixQuote(tlSEE.sABCPgmCode) & "', "
        sgSQLQuery = sgSQLQuery & "seeABCXDSMode = '" & gFixQuote(tlSEE.sABCXDSMode) & "', "
        sgSQLQuery = sgSQLQuery & "seeABCRecordItem = '" & gFixQuote(tlSEE.sABCRecordItem) & "', "
        sgSQLQuery = sgSQLQuery & "seeUnused = '" & gFixQuote(tlSEE.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & " WHERE seeCode = " & tlSEE.lCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
    End If
    gPutUpdate_SEE_Schedule_Events = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_SEE_Schedule_Events = False
    Exit Function
End Function

Public Function gPutUpdate_SEE_ItemIDCheck(llSEECode As Long, slAudioItemIDChk As String, slProtItemIDChk As String, slForm_Module As String) As Integer
    Dim ilANE As Integer
    Dim ilRet As Integer
    Dim slQuery As String
    
    On Error GoTo ErrHand
    If llSEECode > 0 Then
        sgSQLQuery = "Update SEE_Schedule_Events Set "
        slQuery = ""
        If Trim$(slAudioItemIDChk) <> "" Then
            slQuery = "seeAudioItemIDChk = '" & gFixQuote(slAudioItemIDChk) & "'"
        End If
        If Trim$(slProtItemIDChk) <> "" Then
            If slQuery = "" Then
                slQuery = "seeProtItemIDChk = '" & gFixQuote(slProtItemIDChk) & "'"
            Else
                slQuery = slQuery & ", " & "seeProtItemIDChk = '" & gFixQuote(slProtItemIDChk) & "'"
            End If
        End If
        If slQuery <> "" Then
            sgSQLQuery = sgSQLQuery & " " & slQuery
            sgSQLQuery = sgSQLQuery & " WHERE seeCode = " & llSEECode
            cnn.Execute sgSQLQuery    ', rdExecDirect
        End If
    End If
    gPutUpdate_SEE_ItemIDCheck = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_SEE_ItemIDCheck = False
    Exit Function
End Function

Public Function gPutUpdate_SEE_SentFlag(llSEECode As Long, slForm_Module As String) As Integer
    Dim ilANE As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    If llSEECode > 0 Then
        sgSQLQuery = "Update SEE_Schedule_Events Set "
        sgSQLQuery = sgSQLQuery & "seeSentStatus = 'S'" & ", "
        sgSQLQuery = sgSQLQuery & "seeSentDate = '" & Format$(gNow(), sgSQLDateForm) & "'"
        sgSQLQuery = sgSQLQuery & " WHERE seeCode = " & llSEECode
        cnn.Execute sgSQLQuery    ', rdExecDirect
    End If
    gPutUpdate_SEE_SentFlag = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_SEE_SentFlag = False
    Exit Function
End Function

Public Function gPutUpdate_SEE_UnsentFlag(llSEECode As Long, slAction As String, slForm_Module As String) As Integer
    Dim ilANE As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    If llSEECode > 0 Then
        sgSQLQuery = "Update SEE_Schedule_Events Set "
        sgSQLQuery = sgSQLQuery & "seeSentStatus = 'N'" & ", "
        sgSQLQuery = sgSQLQuery & "seeAction = '" & slAction & "', "
        sgSQLQuery = sgSQLQuery & "seeSentDate = '" & Format$("12/31/2069", sgSQLDateForm) & "'"
        sgSQLQuery = sgSQLQuery & " WHERE seeCode = " & llSEECode
        cnn.Execute sgSQLQuery    ', rdExecDirect
    End If
    gPutUpdate_SEE_UnsentFlag = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_SEE_UnsentFlag = False
    Exit Function
End Function

Public Function gPutUpdate_SHE_SentFlags(llSheCode As Long, slForm_Module As String) As Integer
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    If llSheCode > 0 Then
        sgSQLQuery = "Update SHE_Schedule_Header Set "
        sgSQLQuery = sgSQLQuery & "sheLoadedAutoStatus = 'L'" & ", "
        sgSQLQuery = sgSQLQuery & "sheLoadedAutoDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "sheCreateLoad = 'N'"
        sgSQLQuery = sgSQLQuery & " WHERE sheCode = " & llSheCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
    End If
    gPutUpdate_SHE_SentFlags = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_SHE_SentFlags = False
    Exit Function
End Function


Public Function gSchdAndAsAiredDelete(slPriorToDate As String, slForm_Module As String) As Integer
    Dim ilRet As Integer
    Dim llPriorToDate As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    lgPurgeCount = 0
    If slPriorToDate = "" Then
        gSchdAndAsAiredDelete = True
        Exit Function
    End If
    If Not gIsDate(slPriorToDate) Then
        gSchdAndAsAiredDelete = False
        Exit Function
    End If
    llPriorToDate = gDateValue(DateAdd("d", -1, slPriorToDate))
    'Alternative Delete call that would remove all records from see in one step:
    'DELETE FROM "SEE_Schedule_Events" WHERE seesheCode IN (SELECT seesheCode from "SEE_Schedule_Events", "SHE_Schedule_Header" where seeshecode = shecode and sheAirDate < '2005-01-18')

    'sgSQLQuery = "SELECT cmeCode, cmeEndDate FROM CME_Conflict_Master WHERE cmeSource = " & "'S'" & " AND " & "cmeStartDate < " & llPriorToDate
    'Set rst2 = cnn.Execute(sgSQLQuery)
    'While Not rst2.EOF
    '    If gDateValue(Format$(rst2!cmeEndDate, sgShowDateForm)) < llPriorToDate Then
    '        sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeCode = " & rst2!cmeCode
    '        cnn.Execute sgSQLQuery    ', rdExecDirect
    '    End If
    '    rst2.MoveNext
    'Wend
    
    'sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeSource = " & "'S'" & " AND " & "cmeEndDate < " & llPriorToDate
    'cnn.Execute sgSQLQuery
    
    sgSQLQuery = "SELECT sheCode FROM SHE_Schedule_Header WHERE sheAirDate < '" & Format(slPriorToDate, sgSQLDateForm) & "'"
    Set rst = cnn.Execute(sgSQLQuery)
    While Not rst.EOF
        sgSQLQuery = "DELETE FROM SEE_Schedule_Events Where seeSheCode = " & rst!sheCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        sgSQLQuery = "DELETE FROM AAE_As_Aired Where aaeSheCode = " & rst!sheCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        sgSQLQuery = "DELETE FROM AIE_Active_Info Where aieOrigFileCode = " & rst!sheCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        'sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeSource = " & "'S'"
        'sgSQLQuery = sgSQLQuery & " AND cmeSHEDHECode = " & rst!sheCode
        'sgSQLQuery = sgSQLQuery & " AND cmeXMidNight = 'N'"
        'cnn.Execute sgSQLQuery    ', rdExecDirect
        
        'sgSQLQuery = "SELECT cmeCode, cmeXMidNight FROM CME_Conflict_Master WHERE cmeSource = " & "'S'" & " AND cmeSHEDHECode = " & rst!sheCode
        'Set rst2 = cnn.Execute(sgSQLQuery)
        'While Not rst2.EOF
        '    If rst2!cmeXMidNight = "N" Then
        '        sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeCode = " & rst2!cmeCode
        '        cnn.Execute sgSQLQuery    ', rdExecDirect
        '    End If
        '    rst2.MoveNext
        'Wend
        lgPurgeCount = lgPurgeCount + 1
        rst.MoveNext
    Wend
    On Error Resume Next
    rst.Close
    rst2.Close
    On Error GoTo ErrHand
    'sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeSource = " & "'S'" & " AND " & "cmeEndDate < " & gDateValue(DateAdd("d", -1, slPriorToDate))
    'cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM SHE_Schedule_Header Where sheAirDate < '" & Format(slPriorToDate, sgSQLDateForm) & "'"
    cnn.Execute sgSQLQuery    ', rdExecDirect
    
    gSchdAndAsAiredDelete = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gSchdAndAsAiredDelete = False
    Exit Function

End Function
Public Function gLibraryDelete(slPriorToDate As String, slForm_Module As String) As Integer
    Dim ilRet As Integer
    Dim llPriorToDate As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    lgPurgeCount = 0
    If slPriorToDate = "" Then
        gLibraryDelete = True
        Exit Function
    End If
    If Not gIsDate(slPriorToDate) Then
        gLibraryDelete = False
        Exit Function
    End If
    llPriorToDate = gDateValue(DateAdd("d", -1, slPriorToDate))
    'Alternative Delete call that would remove all records from see in one step:
    'DELETE FROM "SEE_Schedule_Events" WHERE seesheCode IN (SELECT seesheCode from "SEE_Schedule_Events", "SHE_Schedule_Header" where seeshecode = shecode and sheAirDate < '2005-01-18')

    'sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeSource = " & "'L'" & " AND " & "cmeEndDate < " & llPriorToDate
    'cnn.Execute sgSQLQuery
    
    sgSQLQuery = "SELECT dheCode FROM DHE_Day_Header_Info WHERE dheType = 'L' AND dheEndDate < '" & Format(slPriorToDate, sgSQLDateForm) & "'"
    Set rst = cnn.Execute(sgSQLQuery)
    While Not rst.EOF
        sgSQLQuery = "DELETE FROM EBE_Event_Bus_Sel WHERE ebeDeeCode IN (SELECT deeCode FROM DEE_Day_Event_Info WHERE deeDheCode = " & rst!dheCode & ")"
        cnn.Execute sgSQLQuery
        sgSQLQuery = "DELETE FROM DEE_Day_Event_Info Where deeDheCode = " & rst!dheCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        sgSQLQuery = "DELETE FROM AIE_Active_Info Where aieOrigFileCode = " & rst!dheCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        sgSQLQuery = "DELETE FROM DBE_Day_Bus_Sel Where dbeDheCode = " & rst!dheCode
        cnn.Execute sgSQLQuery    ', rdExecDirect
        'sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeSource = " & "'S'"
        'sgSQLQuery = sgSQLQuery & " AND cmeSHEDHECode = " & rst!sheCode
        'sgSQLQuery = sgSQLQuery & " AND cmeXMidNight = 'N'"
        'cnn.Execute sgSQLQuery    ', rdExecDirect
        
        'sgSQLQuery = "SELECT cmeCode, cmeXMidNight FROM CME_Conflict_Master WHERE cmeSource = " & "'L'" & " AND cmeSHEDHECode = " & rst!dheCode
        'Set rst2 = cnn.Execute(sgSQLQuery)
        'While Not rst2.EOF
        '    If rst2!cmeXMidNight = "N" Then
        '        sgSQLQuery = "DELETE FROM CME_Conflict_Master WHERE cmeCode = " & rst2!cmeCode
        '        cnn.Execute sgSQLQuery    ', rdExecDirect
        '    End If
        '    rst2.MoveNext
        'Wend
        lgPurgeCount = lgPurgeCount + 1
        rst.MoveNext
    Wend
    On Error Resume Next
    rst.Close
    rst2.Close
    On Error GoTo ErrHand
    sgSQLQuery = "DELETE FROM CEE_Conflict_Events WHERE ceeEndDate < " & gDateValue(DateAdd("d", -1, slPriorToDate))
    cnn.Execute sgSQLQuery    ', rdExecDirect
    sgSQLQuery = "DELETE FROM DHE_Day_Header_Info Where dheType = 'L' AND dheEndDate < '" & Format(slPriorToDate, sgSQLDateForm) & "'"
    cnn.Execute sgSQLQuery    ', rdExecDirect
    
    gLibraryDelete = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gLibraryDelete = False
    Exit Function

End Function

Public Function gTemplateSchdDelete(slPriorToDate As String, slForm_Module As String) As Integer
    Dim ilRet As Integer
    Dim llPriorToDate As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    lgPurgeCount = 0
    If slPriorToDate = "" Then
        gTemplateSchdDelete = True
        Exit Function
    End If
    If Not gIsDate(slPriorToDate) Then
        gTemplateSchdDelete = False
        Exit Function
    End If
    llPriorToDate = gDateValue(DateAdd("d", -1, slPriorToDate))
    
    sgSQLQuery = "SELECT Count(tseCode) FROM TSE_Template_Schd Where  tseLogDate < '" & Format(slPriorToDate, sgSQLDateForm) & "'"
    Set rst = cnn.Execute(sgSQLQuery)
    lgPurgeCount = rst(0).Value
    
    sgSQLQuery = "DELETE FROM TSE_Template_Schd Where  tseLogDate < '" & Format(slPriorToDate, sgSQLDateForm) & "'"
    cnn.Execute sgSQLQuery    ', rdExecDirect
    
    gTemplateSchdDelete = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gTemplateSchdDelete = False
    Exit Function

End Function
Public Function gCommentDelete(slForm_Module As String) As Integer
    Dim ilRet As Integer
    Dim llPriorToDate As Long
    Dim ilDelete As Integer
    Dim llCode As Long
    Dim llLowLimit As Long
    Dim llSvLowLimit As Long
    Dim llHighLimit As Long
    Dim llMaxCode As Long
    Dim rst As ADODB.Recordset
    
    Dim slStr1 As String
    Dim slStr2 As String
    Dim slStr3 As String
    Dim slStr4 As String
    Dim slStr5 As String
    Dim slStr6 As String
    Dim slStr7 As String
    Dim slStr8 As String
    Dim slStr9 As String
    Dim slStr10 As String
    Dim slStr11 As String
        
    On Error GoTo ErrHand
    lgPurgeCount = 0
    
'    sgSQLQuery = "SELECT cteCode FROM CTE_Commts_And_Title"
'    Set rst = cnn.Execute(sgSQLQuery)
'    While Not rst.EOF
'        ilDelete = True
'        llCode = rst!cteCode
'        sgSQLQuery = "SELECT dheCteCode FROM DHE_Day_Header_Info WHERE dheCteCode = " & llCode
'        Set rst2 = cnn.Execute(sgSQLQuery)
'        If Not rst2.EOF Then
'            ilDelete = False
'        End If
'        If ilDelete Then
'            sgSQLQuery = "SELECT dee1CteCode FROM DEE_Day_Event_Info WHERE dee1CteCode = " & llCode
'            Set rst2 = cnn.Execute(sgSQLQuery)
'            If Not rst2.EOF Then
'                ilDelete = False
'            End If
'        End If
'        If ilDelete Then
'            sgSQLQuery = "SELECT dee2CteCode FROM DEE_Day_Event_Info WHERE dee2CteCode = " & llCode
'            Set rst2 = cnn.Execute(sgSQLQuery)
'            If Not rst2.EOF Then
'                ilDelete = False
'            End If
'        End If
'        If ilDelete Then
'            sgSQLQuery = "SELECT see1CteCode FROM SEE_Schedule_Events WHERE see1CteCode = " & llCode
'            Set rst2 = cnn.Execute(sgSQLQuery)
'            If Not rst2.EOF Then
'                ilDelete = False
'            End If
'        End If
'        If ilDelete Then
'            sgSQLQuery = "SELECT see2CteCode FROM SEE_Schedule_Events WHERE see2CteCode = " & llCode
'            Set rst2 = cnn.Execute(sgSQLQuery)
'            If Not rst2.EOF Then
'                ilDelete = False
'            End If
'        End If
'        If ilDelete Then
'            sgSQLQuery = "SELECT tseCteCode FROM TSE_Template_Schd WHERE tseCteCode = " & llCode
'            Set rst2 = cnn.Execute(sgSQLQuery)
'            If Not rst2.EOF Then
'                ilDelete = False
'            End If
'        End If
'        If ilDelete Then
'            sgSQLQuery = "DELETE FROM CTE_Commts_And_Title WHERE cteCode = " & llCode
'            cnn.Execute sgSQLQuery    ', rdExecDirect
'        End If
'        rst.MoveNext
'    Wend
'    rst.Close
'    rst2.Close
    slStr1 = Now
    sgSQLQuery = "SELECT Min(cteCode) FROM CTE_Commts_And_Title"
    Set rst = cnn.Execute(sgSQLQuery)
    If Not rst.EOF Then
        If IsNull(rst(0).Value) Then
            llLowLimit = 1
        Else
            llLowLimit = rst(0).Value
        End If
        llSvLowLimit = llLowLimit
        llHighLimit = llLowLimit + 99999
        slStr2 = Now
        sgSQLQuery = "SELECT Max(cteCode) FROM CTE_Commts_And_Title"
        Set rst = cnn.Execute(sgSQLQuery)
        If Not rst.EOF Then
            llMaxCode = rst(0).Value
            slStr3 = Now
            Do
                sgSQLQuery = "UPDATE CTE_Commts_And_Title SET cteUsedFlag = 'N' WHERE cteCode >= " & llLowLimit & " AND cteCode <= " & llHighLimit
                cnn.Execute sgSQLQuery
                llLowLimit = llHighLimit
                llHighLimit = llHighLimit + 100000
            Loop While llLowLimit <= llMaxCode
            slStr4 = Now
            sgSQLQuery = "UPDATE CTE_Commts_And_Title SET cteUsedFlag = 'Y' WHERE cteCode IN (SELECT dheCteCode FROM DHE_Day_Header_Info WHERE dheCteCode > 0)"
            cnn.Execute sgSQLQuery
            slStr5 = Now
            sgSQLQuery = "UPDATE CTE_Commts_And_Title SET cteUsedFlag = 'Y' WHERE cteCode IN (SELECT dee1CteCode FROM DEE_Day_Event_Info WHERE dee1CteCode > 0)"
            cnn.Execute sgSQLQuery
            slStr6 = Now
            sgSQLQuery = "UPDATE CTE_Commts_And_Title SET cteUsedFlag = 'Y' WHERE cteCode IN (SELECT dee2CteCode FROM DEE_Day_Event_Info WHERE dee2CteCode > 0)"
            cnn.Execute sgSQLQuery
            slStr7 = Now
            sgSQLQuery = "UPDATE CTE_Commts_And_Title SET cteUsedFlag = 'Y' WHERE cteCode IN (SELECT see1CteCode FROM SEE_Schedule_Events WHERE see1CteCode > 0)"
            cnn.Execute sgSQLQuery
            slStr8 = Now
            sgSQLQuery = "UPDATE CTE_Commts_And_Title SET cteUsedFlag = 'Y' WHERE cteCode IN (SELECT see2CteCode FROM SEE_Schedule_Events WHERE see2CteCode > 0)"
            cnn.Execute sgSQLQuery
            slStr9 = Now
            sgSQLQuery = "UPDATE CTE_Commts_And_Title SET cteUsedFlag = 'Y' WHERE cteCode IN (SELECT tseCteCode FROM TSE_Template_Schd WHERE tseCteCode > 0)"
            cnn.Execute sgSQLQuery
            slStr10 = Now
            llLowLimit = llSvLowLimit
            llHighLimit = llLowLimit + 99999
            Do
                sgSQLQuery = "SELECT Count(cteCode) FROM CTE_Commts_And_Title WHERE cteUsedFlag = 'N' AND cteCode >= " & llLowLimit & " AND cteCode <= " & llHighLimit
                Set rst = cnn.Execute(sgSQLQuery)
                lgPurgeCount = lgPurgeCount + rst(0).Value
                sgSQLQuery = "DELETE FROM CTE_Commts_And_Title WHERE cteUsedFlag = 'N' AND cteCode >= " & llLowLimit & " AND cteCode <= " & llHighLimit
                cnn.Execute sgSQLQuery
                llLowLimit = llHighLimit
                llHighLimit = llHighLimit + 100000
            Loop While llLowLimit <= llMaxCode
            slStr11 = Now
        End If
    End If
    rst.Close
    gCommentDelete = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gCommentDelete = False
    Exit Function

End Function

Public Function gPutReplace_ASE_BkupANECode(ilOldBkupANECode As Integer, ilNewBkupANECode As Integer, slForm_Module As String)
    On Error GoTo ErrHand
    sgSQLQuery = "Update ASE_Audio_Source Set "
    sgSQLQuery = sgSQLQuery & "aseBkupAneCode = " & ilNewBkupANECode
    If ilNewBkupANECode = 0 Then
        sgSQLQuery = sgSQLQuery & ", " & "aseBkupCceCode = " & ilNewBkupANECode
    End If
    sgSQLQuery = sgSQLQuery & " WHERE aseBkupAneCode = " & ilOldBkupANECode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutReplace_ASE_BkupANECode = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutReplace_ASE_BkupANECode = False
    Exit Function

End Function

Public Function gPutReplace_ASE_ProtANECode(ilOldProtANECode As Integer, ilNewProtANECode As Integer, slForm_Module As String) As Integer
    On Error GoTo ErrHand
    sgSQLQuery = "Update ASE_Audio_Source Set "
    sgSQLQuery = sgSQLQuery & "aseProtAneCode = " & ilNewProtANECode
    If ilNewProtANECode = 0 Then
        sgSQLQuery = sgSQLQuery & ", " & "aseProtCceCode = " & ilNewProtANECode
    End If
    sgSQLQuery = sgSQLQuery & " WHERE aseProtAneCode = " & ilOldProtANECode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutReplace_ASE_ProtANECode = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutReplace_ASE_ProtANECode = False
    Exit Function

End Function


Public Function gClearFile(slFileName As String, slForm_Module As String) As Integer
    On Error GoTo ErrHand
    
    sgSQLQuery = "DELETE FROM " & slFileName
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gClearFile = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gClearFile = False
    Exit Function
End Function

Public Function gPutReplace_CME_Schd(llOldSEECode As Long, llNewSEECode As Long, slForm_Module As String) As Integer
    On Error GoTo ErrHand
    '12/11/09:  Remove to make Saving Libraries faster
    gPutReplace_CME_Schd = True
    Exit Function

    sgSQLQuery = "Update CME_Conflict_Master Set "
    sgSQLQuery = sgSQLQuery & "cmeSEECode = " & llNewSEECode
    sgSQLQuery = sgSQLQuery & " WHERE cmeSEECode = " & llOldSEECode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutReplace_CME_Schd = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutReplace_CME_Schd = False
    Exit Function

End Function

Public Function gPutReplace_SEE_SHECode(llOldSEECode As Long, llOldSHECode As Long, slForm_Module As String) As Integer
    On Error GoTo ErrHand
    sgSQLQuery = "Update SEE_Schedule_Events Set "
    sgSQLQuery = sgSQLQuery & "seeSheCode = " & llOldSHECode
    sgSQLQuery = sgSQLQuery & " WHERE seeCode = " & llOldSEECode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutReplace_SEE_SHECode = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutReplace_SEE_SHECode = False
    Exit Function

End Function

Private Sub mConflictCMEBusRec(slSource As String, tlDEE As DEE, llDSECode As Long, llStartDate As Long, llEndDate As Long, llStartTime As Long, llEndTime As Long, hlCME As Integer)
    Dim tlCME As CME
    Dim ilDay As Integer
    Dim ilEBE As Integer
    Dim ilRet As Integer
    
    
    DoEvents
    '12/11/09:  Remove to make Saving Libraries faster
    Exit Sub
    
    If (tlDEE.sIgnoreConflicts = "B") Or (tlDEE.sIgnoreConflicts = "I") Then
        Exit Sub
    End If
    If slSource = "L" Then
        smCurrLibEBEStamp = ""
        ilRet = gGetRecs_EBE_EventBusSel(smCurrLibEBEStamp, tlDEE.lCode, "Bus Definition-mDEEMoveRecToCtrls", tmCurrLibEBE())
    Else
        ReDim tmCurrLibEBE(0 To 1) As EBE
        tmCurrLibEBE(0).iBdeCode = tlDEE.iFneCode
    End If
    DoEvents
    For ilEBE = LBound(tmCurrLibEBE) To UBound(tmCurrLibEBE) - 1 Step 1
        DoEvents
        For ilDay = 1 To 7 Step 1
            DoEvents
            If Mid$(tlDEE.sDays, ilDay, 1) = "Y" Then
                tlCME.lCode = 0
                tlCME.sSource = slSource
                tlCME.lSHEDHECode = tlDEE.lDheCode
                tlCME.lDseCode = llDSECode
                tlCME.lDeeCode = tlDEE.lCode
                If slSource = "S" Then
                    tlCME.lSeeCode = tlDEE.lEventID
                Else
                    tlCME.lSeeCode = 0
                End If
                tlCME.sEvtType = "B"
                tlCME.iBdeCode = tmCurrLibEBE(ilEBE).iBdeCode
                tlCME.iANECode = 0
                tlCME.sItemID = ""
                tlCME.sXMidNight = "N"
                If llEndTime <= 864000 Then
                    tlCME.lStartDate = llStartDate
                    tlCME.lEndDate = llEndDate
                    Select Case ilDay
                        Case 1
                            tlCME.sDay = "Mo"
                        Case 2
                            tlCME.sDay = "Tu"
                        Case 3
                            tlCME.sDay = "We"
                        Case 4
                            tlCME.sDay = "Th"
                        Case 5
                            tlCME.sDay = "Fr"
                        Case 6
                            tlCME.sDay = "Sa"
                        Case 7
                            tlCME.sDay = "Su"
                    End Select
                    tlCME.lStartTime = llStartTime
                    tlCME.lEndTime = llEndTime
                    tlCME.sUnused = ""
                    ilRet = gPutInsert_CME_Conflict_Master(tlCME, "mConflictCMEBusRec", hlCME)
                Else
                    tlCME.lStartDate = llStartDate
                    tlCME.lEndDate = llEndDate
                    Select Case ilDay
                        Case 1
                            tlCME.sDay = "Mo"
                        Case 2
                            tlCME.sDay = "Tu"
                        Case 3
                            tlCME.sDay = "We"
                        Case 4
                            tlCME.sDay = "Th"
                        Case 5
                            tlCME.sDay = "Fr"
                        Case 6
                            tlCME.sDay = "Sa"
                        Case 7
                            tlCME.sDay = "Su"
                    End Select
                    tlCME.lStartTime = llStartTime
                    tlCME.lEndTime = 864000
                    tlCME.sUnused = ""
                    ilRet = gPutInsert_CME_Conflict_Master(tlCME, "mConflictCMEBusRec", hlCME)
                    tlCME.lCode = 0
                    tlCME.sXMidNight = "Y"
                    tlCME.lStartDate = llStartDate + 1
                    If llEndDate <> gDateValue("12/31/2069") Then
                        tlCME.lEndDate = llEndDate + 1
                    End If
                    Select Case ilDay
                        Case 1
                            tlCME.sDay = "Tu"
                        Case 2
                            tlCME.sDay = "We"
                        Case 3
                            tlCME.sDay = "Th"
                        Case 4
                            tlCME.sDay = "Fr"
                        Case 5
                            tlCME.sDay = "Sa"
                        Case 6
                            tlCME.sDay = "Su"
                        Case 7
                            tlCME.sDay = "Mo"
                    End Select
                    tlCME.lStartTime = 0
                    tlCME.lEndTime = llEndTime - 864000
                    tlCME.sUnused = ""
                    ilRet = gPutInsert_CME_Conflict_Master(tlCME, "mConflictCMEBusRec", hlCME)
                End If
            End If
        Next ilDay
    Next ilEBE
End Sub

Private Sub mConflictCMEAudioRec(slSource As String, tlDEE As DEE, llDSECode As Long, ilANECode As Integer, slItemID As String, llStartDate As Long, llEndDate As Long, llStartTime As Long, llEndTime As Long, hlCME As Integer)
    Dim ilANE As Integer
    Dim ilATE As Integer
    Dim llPreTime As Long
    Dim llPostTime As Long
    Dim llSTime As Long
    Dim llETime As Long
    Dim ilDay As Integer
    Dim ilRet As Integer
    Dim ilEBE As Integer
    Dim ilBdeCode As Integer
    Dim ilETE As Integer
    Dim tlCME As CME
    
    DoEvents
    '12/11/09:  Remove to make Saving Libraries faster
    Exit Sub
    
    If ilANECode <= 0 Then
        Exit Sub
    End If
    llPreTime = 0
    llPostTime = 0
    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
    '    If ilANECode = tgCurrANE(ilANE).iCode Then
        ilANE = gBinarySearchANE(ilANECode, tgCurrANE())
        If ilANE <> -1 Then
            For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
                If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                    If tgCurrANE(ilANE).sCheckConflicts = "N" Then
                        Exit Sub
                    End If
                    llPreTime = tgCurrATE(ilATE).lPreBufferTime
                    llPostTime = tgCurrATE(ilATE).lPostBufferTime
                End If
            Next ilATE
    '        Exit For
        End If
    'Next ilANE
    DoEvents
    llSTime = llStartTime - llPreTime
    llETime = llEndTime + llPostTime
    If slSource = "L" Then
        smCurrLibEBEStamp = ""
        ilRet = gGetRecs_EBE_EventBusSel(smCurrLibEBEStamp, tlDEE.lCode, "Bus Definition-mDEEMoveRecToCtrls", tmCurrLibEBE())
    Else
        ReDim tmCurrLibEBE(0 To 1) As EBE
        tmCurrLibEBE(0).iBdeCode = tlDEE.iFneCode
    End If
    DoEvents
    For ilEBE = LBound(tmCurrLibEBE) To UBound(tmCurrLibEBE) - 1 Step 1
        DoEvents
        'Only test bus is avail event
        ilBdeCode = tmCurrLibEBE(ilEBE).iBdeCode
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tgCurrETE(ilETE).iCode = tlDEE.iEteCode Then
                If tgCurrETE(ilETE).sCategory <> "A" Then
                    ilBdeCode = 0
                End If
                Exit For
            End If
        Next ilETE
        For ilDay = 1 To 7 Step 1
            DoEvents
            If Mid$(tlDEE.sDays, ilDay, 1) = "Y" Then
                tlCME.lCode = 0
                tlCME.sSource = slSource
                tlCME.lSHEDHECode = tlDEE.lDheCode
                tlCME.lDseCode = llDSECode
                If slSource = "S" Then
                    tlCME.lSeeCode = tlDEE.lEventID
                Else
                    tlCME.lSeeCode = 0
                End If
                tlCME.lDeeCode = tlDEE.lCode
                tlCME.sEvtType = "A"
                tlCME.iBdeCode = ilBdeCode
                tlCME.iANECode = ilANECode
                tlCME.sXMidNight = "N"
                tlCME.sItemID = slItemID
                If llSTime < 0 Then
                    tlCME.sXMidNight = "Y"
                    tlCME.lStartDate = llStartDate - 1
                    tlCME.lEndDate = llEndDate - 1
                    Select Case ilDay
                        Case 1
                            tlCME.sDay = "Su"
                        Case 2
                            tlCME.sDay = "Mo"
                        Case 3
                            tlCME.sDay = "Tu"
                        Case 4
                            tlCME.sDay = "We"
                        Case 5
                            tlCME.sDay = "Th"
                        Case 6
                            tlCME.sDay = "Fr"
                        Case 7
                            tlCME.sDay = "Sa"
                    End Select
                    tlCME.lStartTime = 864000 + llSTime
                    tlCME.lEndTime = 864000
                    tlCME.sUnused = ""
                    ilRet = gPutInsert_CME_Conflict_Master(tlCME, "mConflictCMEAudioRec", hlCME)
                    tlCME.lCode = 0
                    tlCME.sXMidNight = "N"
                    tlCME.lStartDate = llStartDate
                    tlCME.lEndDate = llEndDate
                    Select Case ilDay
                        Case 1
                            tlCME.sDay = "Mo"
                        Case 2
                            tlCME.sDay = "Tu"
                        Case 3
                            tlCME.sDay = "We"
                        Case 4
                            tlCME.sDay = "Th"
                        Case 5
                            tlCME.sDay = "Fr"
                        Case 6
                            tlCME.sDay = "Sa"
                        Case 7
                            tlCME.sDay = "Su"
                    End Select
                    tlCME.lStartTime = 0
                    tlCME.lEndTime = llETime
                    tlCME.sUnused = ""
                    ilRet = gPutInsert_CME_Conflict_Master(tlCME, "mConflictCMEAudioRec", hlCME)
                Else
                    If llEndTime <= 864000 Then
                        tlCME.lStartDate = llStartDate
                        tlCME.lEndDate = llEndDate
                        Select Case ilDay
                            Case 1
                                tlCME.sDay = "Mo"
                            Case 2
                                tlCME.sDay = "Tu"
                            Case 3
                                tlCME.sDay = "We"
                            Case 4
                                tlCME.sDay = "Th"
                            Case 5
                                tlCME.sDay = "Fr"
                            Case 6
                                tlCME.sDay = "Sa"
                            Case 7
                                tlCME.sDay = "Su"
                        End Select
                        tlCME.lStartTime = llStartTime
                        tlCME.lEndTime = llEndTime
                        tlCME.sUnused = ""
                        ilRet = gPutInsert_CME_Conflict_Master(tlCME, "mConflictCMEAudioRec", hlCME)
                    Else
                        tlCME.lStartDate = llStartDate
                        tlCME.lEndDate = llEndDate
                        Select Case ilDay
                            Case 1
                                tlCME.sDay = "Mo"
                            Case 2
                                tlCME.sDay = "Tu"
                            Case 3
                                tlCME.sDay = "We"
                            Case 4
                                tlCME.sDay = "Th"
                            Case 5
                                tlCME.sDay = "Fr"
                            Case 6
                                tlCME.sDay = "Sa"
                            Case 7
                                tlCME.sDay = "Su"
                        End Select
                        tlCME.lStartTime = llStartTime
                        tlCME.lEndTime = 864000
                        tlCME.sUnused = ""
                        ilRet = gPutInsert_CME_Conflict_Master(tlCME, "mConflictCMEAudioRec", hlCME)
                        tlCME.lCode = 0
                        tlCME.sXMidNight = "Y"
                        tlCME.lStartDate = llStartDate + 1
                        If llEndDate <> gDateValue("12/31/2069") Then
                            tlCME.lEndDate = llEndDate + 1
                        End If
                        Select Case ilDay
                            Case 1
                                tlCME.sDay = "Tu"
                            Case 2
                                tlCME.sDay = "We"
                            Case 3
                                tlCME.sDay = "Th"
                            Case 4
                                tlCME.sDay = "Fr"
                            Case 5
                                tlCME.sDay = "Sa"
                            Case 6
                                tlCME.sDay = "Su"
                            Case 7
                                tlCME.sDay = "Mo"
                        End Select
                        tlCME.lStartTime = 0
                        tlCME.lEndTime = llEndTime - 864000
                        tlCME.sUnused = ""
                        ilRet = gPutInsert_CME_Conflict_Master(tlCME, "mConflictCMEAudioRec", hlCME)
                    End If
                End If
            End If
        Next ilDay
    Next ilEBE
End Sub

Public Function gCreateCMEForLib(tlDHE As DHE, slDateStartRange As String, hlCME As Integer) As Integer
    Dim ilRet As Integer
    Dim ilDEE As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llOffsetStartTime As Long
    Dim llOffsetEndTime As Long
    Dim ilPriAneCode As Integer
    Dim ilProtAneCode As Integer
    Dim ilBkupAneCode As Integer
    Dim ilASE As Integer
    Dim ilHour As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    
    gCreateCMEForLib = True
    '12/11/09:  Remove to make Saving Libraries faster
    Exit Function
    
    If (tlDHE.sState = "D") Or (tlDHE.sState = "L") Then
        Exit Function
    End If
    slStartDate = tlDHE.sStartDate
    If gDateValue(slDateStartRange) > gDateValue(slStartDate) Then
        slStartDate = slDateStartRange
    End If
    slEndDate = tlDHE.sEndDate
    If gDateValue(slEndDate) < gDateValue(slStartDate) Then
        Exit Function
    End If
    ilRet = gGetRecs_DEE_DayEvent(smCurrLibDEEStamp, tlDHE.lCode, "EngrLibDef-mPopulate", tmCurrLibDEE())
    For ilDEE = 0 To UBound(tmCurrLibDEE) - 1 Step 1
        llOffsetStartTime = tmCurrLibDEE(ilDEE).lTime
        llOffsetEndTime = llOffsetStartTime + tmCurrLibDEE(ilDEE).lDuration ' - 1
        If llOffsetEndTime < llOffsetStartTime Then
            llOffsetEndTime = llOffsetStartTime
        End If
        ilPriAneCode = 0
        ilProtAneCode = 0
        ilBkupAneCode = 0
        If (tmCurrLibDEE(ilDEE).sIgnoreConflicts <> "A") And (tmCurrLibDEE(ilDEE).sIgnoreConflicts <> "I") Then
            If tmCurrLibDEE(ilDEE).iAudioAseCode > 0 Then
                'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                '    If tmCurrLibDEE(ilDEE).iAudioAseCode = tgCurrASE(ilASE).iCode Then
                '        ilPriAneCode = tgCurrASE(ilASE).iPriAneCode
                '        Exit For
                '    End If
                'Next ilASE
                ilASE = gBinarySearchASE(tmCurrLibDEE(ilDEE).iAudioAseCode, tgCurrASE())
                If ilASE <> -1 Then
                    ilPriAneCode = tgCurrASE(ilASE).iPriAneCode
                End If
            End If
            ilProtAneCode = tmCurrLibDEE(ilDEE).iProtAneCode
            ilBkupAneCode = tmCurrLibDEE(ilDEE).iBkupAneCode
        End If
        For ilHour = 1 To 24 Step 1
            If Mid$(tmCurrLibDEE(ilDEE).sHours, ilHour, 1) = "Y" Then
                llStartTime = 36000 * (ilHour - 1) + llOffsetStartTime
                llEndTime = 36000 * (ilHour - 1) + llOffsetEndTime
                llStartDate = gDateValue(slStartDate)
                llEndDate = gDateValue(slEndDate)
                mConflictCMEBusRec "L", tmCurrLibDEE(ilDEE), tlDHE.lDseCode, llStartDate, llEndDate, llStartTime, llEndTime, hlCME
                mConflictCMEAudioRec "L", tmCurrLibDEE(ilDEE), tlDHE.lDseCode, ilPriAneCode, tmCurrLibDEE(ilDEE).sAudioItemID, llStartDate, llEndDate, llStartTime, llEndTime, hlCME
                mConflictCMEAudioRec "L", tmCurrLibDEE(ilDEE), tlDHE.lDseCode, ilProtAneCode, tmCurrLibDEE(ilDEE).sProtItemID, llStartDate, llEndDate, llStartTime, llEndTime, hlCME
                mConflictCMEAudioRec "L", tmCurrLibDEE(ilDEE), tlDHE.lDseCode, ilBkupAneCode, tmCurrLibDEE(ilDEE).sAudioItemID, llStartDate, llEndDate, llStartTime, llEndTime, hlCME
            End If
        Next ilHour
    Next ilDEE
End Function

Public Function gCreateCMEForTemp(tlDHE As DHE, tlTSE As TSE, hlCME As Integer) As Integer
    Dim ilRet As Integer
    Dim ilDEE As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llOffsetStartTime As Long
    Dim llOffsetEndTime As Long
    Dim ilPriAneCode As Integer
    Dim ilProtAneCode As Integer
    Dim ilBkupAneCode As Integer
    Dim ilASE As Integer
    Dim ilHour As Integer
    Dim slHours As String
    Dim ilLoop As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llLength As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    
    gCreateCMEForTemp = True
    '12/11/09:  Remove to make Saving Libraries faster
    Exit Function
    
    If (tlDHE.sState <> "D") And (tlDHE.sState <> "L") And (tlTSE.sState <> "D") And (tlTSE.sState <> "L") Then
        If (tlDHE.sIgnoreConflicts <> "I") Then
            smCurrLibDEEStamp = ""
            ilRet = gGetRecs_DEE_DayEvent(smCurrLibDEEStamp, tlDHE.lCode, "EngrLibDef-mPopulate", tmCurrLibDEE())
            For ilDEE = 0 To UBound(tmCurrLibDEE) - 1 Step 1
                tmCurrLibDEE(ilDEE).sIgnoreConflicts = "N"
                ilHour = Hour(tlTSE.sStartTime)
                If ilHour <> 0 Then
                    slHours = tmCurrLibDEE(ilDEE).sHours
                    tmCurrLibDEE(ilDEE).sHours = String(24, "N")
                    For ilLoop = 0 To 23 Step 1
                        Mid$(tmCurrLibDEE(ilDEE).sHours, ilHour + 1, 1) = Mid$(slHours, ilLoop + 1, 1)
                        ilHour = ilHour + 1
                        If ilHour > 23 Then
                            Exit For
                        End If
                    Next ilLoop
                End If
                tmCurrLibDEE(ilDEE).sDays = String(7, "N")
                Select Case Weekday(tlTSE.sLogDate)
                    Case vbMonday
                        Mid(tmCurrLibDEE(ilDEE).sDays, 1, 1) = "Y"
                    Case vbTuesday
                        Mid(tmCurrLibDEE(ilDEE).sDays, 2, 1) = "Y"
                    Case vbWednesday
                        Mid(tmCurrLibDEE(ilDEE).sDays, 3, 1) = "Y"
                    Case vbThursday
                        Mid(tmCurrLibDEE(ilDEE).sDays, 4, 1) = "Y"
                    Case vbFriday
                        Mid(tmCurrLibDEE(ilDEE).sDays, 5, 1) = "Y"
                    Case vbSaturday
                        Mid(tmCurrLibDEE(ilDEE).sDays, 6, 1) = "Y"
                    Case vbSunday
                        Mid(tmCurrLibDEE(ilDEE).sDays, 7, 1) = "Y"
                End Select
                tmCurrLibDEE(ilDEE).iFneCode = tlTSE.iBdeCode
                
                llOffsetStartTime = tmCurrLibDEE(ilDEE).lTime
                llOffsetEndTime = llOffsetStartTime + tmCurrLibDEE(ilDEE).lDuration ' - 1
                If llOffsetEndTime < llOffsetStartTime Then
                    llOffsetEndTime = llOffsetStartTime
                End If
                ilPriAneCode = 0
                ilProtAneCode = 0
                ilBkupAneCode = 0
                If (tmCurrLibDEE(ilDEE).sIgnoreConflicts <> "A") And (tmCurrLibDEE(ilDEE).sIgnoreConflicts <> "I") Then
                    If tmCurrLibDEE(ilDEE).iAudioAseCode > 0 Then
                        'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                        '    If tmCurrLibDEE(ilDEE).iAudioAseCode = tgCurrASE(ilASE).iCode Then
                        '        ilPriAneCode = tgCurrASE(ilASE).iPriAneCode
                        '        Exit For
                        '    End If
                        'Next ilASE
                        ilASE = gBinarySearchASE(tmCurrLibDEE(ilDEE).iAudioAseCode, tgCurrASE())
                        If ilASE <> -1 Then
                            ilPriAneCode = tgCurrASE(ilASE).iPriAneCode
                        End If
                    End If
                    ilProtAneCode = tmCurrLibDEE(ilDEE).iProtAneCode
                    ilBkupAneCode = tmCurrLibDEE(ilDEE).iBkupAneCode
                End If
                For ilHour = 1 To 24 Step 1
                    If Mid$(tmCurrLibDEE(ilDEE).sHours, ilHour, 1) = "Y" Then
                        llStartTime = 36000 * (ilHour - 1) + llOffsetStartTime
                        llEndTime = 36000 * (ilHour - 1) + llOffsetEndTime
                        llLength = 10 * (gTimeToLong(tlTSE.sStartTime, False) Mod 3600)
                        llStartTime = llStartTime + llLength
                        llEndTime = llEndTime + llLength
                        llStartDate = gDateValue(slStartDate)
                        llEndDate = gDateValue(slEndDate)
                        mConflictCMEBusRec "T", tmCurrLibDEE(ilDEE), tlDHE.lDseCode, llStartDate, llEndDate, llStartTime, llEndTime, hlCME
                        mConflictCMEAudioRec "T", tmCurrLibDEE(ilDEE), tlDHE.lDseCode, ilPriAneCode, tmCurrLibDEE(ilDEE).sAudioItemID, llStartDate, llEndDate, llStartTime, llEndTime, hlCME
                        mConflictCMEAudioRec "T", tmCurrLibDEE(ilDEE), tlDHE.lDseCode, ilProtAneCode, tmCurrLibDEE(ilDEE).sProtItemID, llStartDate, llEndDate, llStartTime, llEndTime, hlCME
                        mConflictCMEAudioRec "T", tmCurrLibDEE(ilDEE), tlDHE.lDseCode, ilBkupAneCode, tmCurrLibDEE(ilDEE).sAudioItemID, llStartDate, llEndDate, llStartTime, llEndTime, hlCME
                    End If
                Next ilHour
            Next ilDEE
        End If
    End If
End Function


Public Function gCreateCMEForSchd(tlSHE As SHE, tlSEE As SEE, ilSpotETECode As Integer, hlCME As Integer) As Integer
    Dim ilPriAneCode As Integer
    Dim ilProtAneCode As Integer
    Dim ilBkupAneCode As Integer
    Dim ilASE As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    ReDim tmCurrLibDEE(0 To 0) As DEE
    
    gCreateCMEForSchd = True
    '12/11/09:  Remove to make Saving Libraries faster
    Exit Function
    
    If (tlSEE.sAction <> "D") And (tlSEE.sAction <> "R") And (ilSpotETECode <> tlSEE.iEteCode) Then
        ilPriAneCode = 0
        ilProtAneCode = 0
        ilBkupAneCode = 0
        If (tlSEE.sIgnoreConflicts <> "A") And (tlSEE.sIgnoreConflicts <> "I") Then
            If tlSEE.iAudioAseCode > 0 Then
                'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                '    If tlSEE.iAudioAseCode = tgCurrASE(ilASE).iCode Then
                '        ilPriAneCode = tgCurrASE(ilASE).iPriAneCode
                '        Exit For
                '    End If
                'Next ilASE
                ilASE = gBinarySearchASE(tlSEE.iAudioAseCode, tgCurrASE())
                If ilASE <> -1 Then
                    ilPriAneCode = tgCurrASE(ilASE).iPriAneCode
                End If
            End If
            ilProtAneCode = tlSEE.iProtAneCode
            ilBkupAneCode = tlSEE.iBkupAneCode
        End If
        llStartTime = tlSEE.lTime
        llEndTime = tlSEE.lTime + tlSEE.lDuration
        llStartDate = gDateValue(tlSHE.sAirDate)
        llEndDate = gDateValue(tlSHE.sAirDate)
        tmCurrLibDEE(0).lDheCode = tlSHE.lCode
        tmCurrLibDEE(0).lCode = tlSEE.lDeeCode
        tmCurrLibDEE(0).lEventID = tlSEE.lCode
        tmCurrLibDEE(0).sIgnoreConflicts = tlSEE.sIgnoreConflicts
        tmCurrLibDEE(0).iFneCode = tlSEE.iBdeCode
        tmCurrLibDEE(0).iEteCode = tlSEE.iEteCode
        tmCurrLibDEE(0).sDays = String(7, "N")
        Select Case Weekday(tlSHE.sAirDate)
            Case vbMonday
                Mid(tmCurrLibDEE(0).sDays, 1, 1) = "Y"
            Case vbTuesday
                Mid(tmCurrLibDEE(0).sDays, 2, 1) = "Y"
            Case vbWednesday
                Mid(tmCurrLibDEE(0).sDays, 3, 1) = "Y"
            Case vbThursday
                Mid(tmCurrLibDEE(0).sDays, 4, 1) = "Y"
            Case vbFriday
                Mid(tmCurrLibDEE(0).sDays, 5, 1) = "Y"
            Case vbSaturday
                Mid(tmCurrLibDEE(0).sDays, 6, 1) = "Y"
            Case vbSunday
                Mid(tmCurrLibDEE(0).sDays, 7, 1) = "Y"
        End Select
        mConflictCMEBusRec "S", tmCurrLibDEE(0), 0, llStartDate, llEndDate, llStartTime, llEndTime, hlCME
        mConflictCMEAudioRec "S", tmCurrLibDEE(0), 0, ilPriAneCode, tlSEE.sAudioItemID, llStartDate, llEndDate, llStartTime, llEndTime, hlCME
        mConflictCMEAudioRec "S", tmCurrLibDEE(0), 0, ilProtAneCode, tlSEE.sProtItemID, llStartDate, llEndDate, llStartTime, llEndTime, hlCME
        mConflictCMEAudioRec "S", tmCurrLibDEE(0), 0, ilBkupAneCode, tlSEE.sAudioItemID, llStartDate, llEndDate, llStartTime, llEndTime, hlCME
    End If

End Function

Public Function gPutInsert_MIE_MessageInfo(tlMie As MIE, slForm_Module As String) As Integer
'
'   tlUIE(I)- Record to be added to Database
'
    Dim llRet As Long
    Dim llLastCode As Long
    Dim llInitCode As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    llLastCode = 0
    llInitCode = tlMie.lCode
    Do
        If tlMie.lCode <= 0 Then
            sgSQLQuery = "Select MAX(mieCode) from MIE_Message_Info"
            Set rst = cnn.Execute(sgSQLQuery)
            If IsNull(rst(0).Value) Then
                tlMie.lCode = 1
            Else
                If rst(0).Value > 0 Then
                    tlMie.lCode = rst(0).Value + 1
                End If
            End If
            rst.Close
            If llLastCode = tlMie.lCode Then
                GoTo ErrHand1:
            End If
        End If
        llLastCode = tlMie.lCode
        sgSQLQuery = "Insert Into MIE_Message_Info ( "
        sgSQLQuery = sgSQLQuery & "mieCode, "
        sgSQLQuery = sgSQLQuery & "mieType, "
        sgSQLQuery = sgSQLQuery & "mieID, "
        sgSQLQuery = sgSQLQuery & "mieMessage, "
        sgSQLQuery = sgSQLQuery & "mieEnteredDate, "
        sgSQLQuery = sgSQLQuery & "mieEnteredTime, "
        sgSQLQuery = sgSQLQuery & "mieUieCode, "
        sgSQLQuery = sgSQLQuery & "mieUnused "
        sgSQLQuery = sgSQLQuery & ") "
        sgSQLQuery = sgSQLQuery & "Values ( "
        sgSQLQuery = sgSQLQuery & tlMie.lCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlMie.sType) & "', "
        sgSQLQuery = sgSQLQuery & tlMie.lID & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlMie.sMessage) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlMie.sEnteredDate, sgSQLDateForm) & "', "
        sgSQLQuery = sgSQLQuery & "'" & Format$(tlMie.sEnteredTime, sgSQLTimeForm) & "', "
        sgSQLQuery = sgSQLQuery & tlMie.iUieCode & ", "
        sgSQLQuery = sgSQLQuery & "'" & gFixQuote(tlMie.sUnused) & "' "
        sgSQLQuery = sgSQLQuery & ") "
        llRet = 0
        cnn.Execute sgSQLQuery    ', rdExecDirect
    Loop While llRet = BTRV_ERR_DUPLICATE_KEY
    gPutInsert_MIE_MessageInfo = True
    Exit Function
ErrHand:
    If (llInitCode = 0) Then
        For Each gErrSQL In cnn.Errors
            llRet = gErrSQL.NativeError
            If llRet < 0 Then
                llRet = llRet + 4999
            End If
            If (llRet = BTRV_ERR_DUPLICATE_KEY) And (llInitCode = 0) Then
                tlMie.lCode = 0
                Resume Next
            End If
        Next gErrSQL
    End If
ErrHand1:
    gShowErrorMsg slForm_Module
    gPutInsert_MIE_MessageInfo = False
    Exit Function
End Function

Public Function gPutUpdate_MIE_MessageInfo(tlMie As MIE, slForm_Module As String) As Integer
    Dim ilRet As Integer
    On Error GoTo ErrHand
    

    sgSQLQuery = "Update MIE_Message_Info Set "
    sgSQLQuery = sgSQLQuery & "mieCode = " & tlMie.lCode & ", "
    sgSQLQuery = sgSQLQuery & "mieType = '" & gFixQuote(tlMie.sType) & "', "
    sgSQLQuery = sgSQLQuery & "mieID = " & tlMie.lID & ", "
    sgSQLQuery = sgSQLQuery & "mieMessage = '" & gFixQuote(tlMie.sMessage) & "', "
    sgSQLQuery = sgSQLQuery & "mieEnteredDate = '" & Format$(tlMie.sEnteredDate, sgSQLDateForm) & "', "
    sgSQLQuery = sgSQLQuery & "mieEnteredTime = '" & Format$(tlMie.sEnteredTime, sgSQLTimeForm) & "', "
    sgSQLQuery = sgSQLQuery & "mieUieCode = " & tlMie.iUieCode & ", "
    sgSQLQuery = sgSQLQuery & "mieUnused = '" & gFixQuote(tlMie.sUnused) & "' "
    sgSQLQuery = sgSQLQuery & " WHERE mieCode = " & tlMie.lCode
    cnn.Execute sgSQLQuery    ', rdExecDirect
    gPutUpdate_MIE_MessageInfo = True
    Exit Function
ErrHand:
    gShowErrorMsg slForm_Module
    gPutUpdate_MIE_MessageInfo = False
    Exit Function
End Function


Public Function gExecGenSQLCall(slSQLCall As String) As Integer
    On Error GoTo ErrHand
    cnn.Execute slSQLCall    ', rdExecDirect
    gExecGenSQLCall = True
    Exit Function
ErrHand:
    gExecGenSQLCall = False
    Exit Function
End Function
