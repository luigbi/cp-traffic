Attribute VB_Name = "EngrSchdSubs"

'
' Release: 1.0
'
' Description:
'   This file contains the Constants

Option Explicit
Private smCurrLibDEEStamp As String
Private tmCurrLibDEE() As DEE
Private smCurrTempTSEStamp As String
Private tmCurrTempTSE() As TSE
Private smCurrEBEStamp As String
Private tmCurrEBE() As EBE
Private smCurrLibDateDHEStamp As String
Private tmCurrLibDateDHE() As DHE
Private smCurrTempDateDHETSEStamp As String
Private tmCurrTempDateDHETSE() As DHETSE

Private tmSHE As SHE

Private tmAAE As AAE
Private hmImport As Integer

Private tmARE As ARE

Private tmAdjSHE() As SHE
Private smAdjSEEStamp As String
Private tmAdjSEE() As SEE
Private smNewChgStamp As String
Private tmNewChgDHE As DHE
Private smNewChgDEEStamp As String
Private tmNewChgDEE() As DEE
Private tmNewChgSEE() As SEE
Private smCheckStamp As String
Private tmCheckDHE As DHE
Private smCheckDEEStamp As String
Private tmCheckDEE() As DEE
Private tmCheckSEE() As SEE
Private smSplitStamp As String
Private tmSplitDHE As DHE
Private smSplitDEEStamp As String
Private tmSplitDEE() As DEE
Private tmSplitSEE() As SEE
Private smExpandStamp As String
Private tmExpandDHE As DHE
Private smExpandDEEStamp As String
Private tmExpandDEE() As DEE
Private tmExpandSEE() As SEE
Private tmMergeInfo() As MERGEINFO

Private smBusDeleted() As String

'Constant must match ones defined in EngrServiceMain and EngrLoad
Const EVENTTYPEINDEX = 0
Const EVENTIDINDEX = 1
Const BUSNAMEINDEX = 2
Const BUSCTRLINDEX = 3
Const TIMEINDEX = 4
Const STARTTYPEINDEX = 5
Const FIXEDINDEX = 6
Const ENDTYPEINDEX = 7
Const DURATIONINDEX = 8
Const MATERIALINDEX = 9
Const AUDIONAMEINDEX = 10
Const AUDIOITEMIDINDEX = 11
Const AUDIOISCIINDEX = 12
Const AUDIOCTRLINDEX = 13
Const BACKUPNAMEINDEX = 18
Const BACKUPCTRLINDEX = 19
Const PROTNAMEINDEX = 14
Const PROTITEMIDINDEX = 15
Const PROTISCIINDEX = 16
Const PROTCTRLINDEX = 17
Const RELAY1INDEX = 20
Const RELAY2INDEX = 21
Const FOLLOWINDEX = 22
Const SILENCETIMEINDEX = 23
Const SILENCE1INDEX = 24
Const SILENCE2INDEX = 25
Const SILENCE3INDEX = 26
Const SILENCE4INDEX = 27
Const NETCUE1INDEX = 28
Const NETCUE2INDEX = 29
Const TITLE1INDEX = 30
Const TITLE2INDEX = 31
Const ABCFORMATINDEX = 32
Const ABCPGMCODEINDEX = 33
Const ABCXDSMODEINDEX = 34
Const ABCRECORDITEMINDEX = 35








Public Sub gCreateHeader(slAirDate As String, tlSHE As SHE)
    Dim slNowDate As String
    Dim slNowTime As String
    
    slNowDate = Format(Now, sgShowDateForm)
    slNowTime = Format(Now, sgShowTimeWSecForm)
    tlSHE.lCode = 0
    tlSHE.iAeeCode = tgAEE.iCode
    tlSHE.sAirDate = Format$(slAirDate, sgShowDateForm)
    tlSHE.sLoadedAutoStatus = "N"
    tlSHE.sLoadedAutoDate = Format$("12/31/2069", sgShowDateForm)
    tlSHE.iChgSeqNo = 0
    tlSHE.sAsAirStatus = "N"
    tlSHE.sLoadedAsAirDate = Format$("12/31/2069", sgShowDateForm)
    tlSHE.sLastDateItemChk = Format$("12/31/2069", sgShowDateForm)
    tlSHE.sCreateLoad = "N"
    tlSHE.iVersion = 0
    tlSHE.lOrigSheCode = 0
    tlSHE.sCurrent = "Y"
    tlSHE.sEnteredDate = slNowDate
    tlSHE.sEnteredTime = slNowTime
    tlSHE.iUieCode = tgUIE.iCode
    tlSHE.sConflictExist = "N"
    tlSHE.sSpotMergeStatus = "N"
    tlSHE.sLoadStatus = "N"
    tlSHE.sUnused = ""

End Sub



Public Function gGetEventsFromLibraries(slAirDate As String) As Integer
    Dim ilRet As Integer
    Dim ilDHE As Integer
    Dim llUpper As Long
    Dim ilDEE As Integer
    Dim ilHours As Integer
    Dim ilDay As Integer
    Dim ilEBE As Integer
    Dim slAirDateMinus1 As String
    
    gGetEventsFromLibraries = True
    ReDim tgCurrSEE(0 To 0) As SEE
    ReDim lgLibDheUsed(0 To 0) As Long
    ilDay = Weekday(slAirDate, vbMonday)
    ilRet = gGetTypeOfRecs_DHE_DayHeaderInfoForLibByDate("C", "A", slAirDate, smCurrLibDateDHEStamp, "gGetEventsFromLibraries- Library DHE", tmCurrLibDateDHE())
    For ilDHE = 0 To UBound(tmCurrLibDateDHE) - 1 Step 1
        ilRet = gGetRecs_DEE_DayEvent(smCurrLibDEEStamp, tmCurrLibDateDHE(ilDHE).lCode, "gGetEventsFromLibraries-Library DEE", tmCurrLibDEE())
        If (UBound(tmCurrLibDEE) > LBound(tmCurrLibDEE)) And (tmCurrLibDateDHE(ilDHE).sUsedFlag <> "Y") Then
            lgLibDheUsed(UBound(lgLibDheUsed)) = tmCurrLibDateDHE(ilDHE).lCode
            ReDim Preserve lgLibDheUsed(0 To UBound(lgLibDheUsed) + 1) As Long
        End If
        mTransferDEEToSEE slAirDate, tmCurrLibDEE(), tmCurrLibDateDHE(ilDHE).lOrigDHECode, tmCurrLibDateDHE(ilDHE).sIgnoreConflicts, tgCurrSEE()
    Next ilDHE
    smCurrTempDateDHETSEStamp = ""
    slAirDateMinus1 = DateAdd("d", -1, slAirDate)
    ilRet = gGetTypeOfRecs_DHETSE_DayHeaderInfoForTempByRange("C", slAirDateMinus1, slAirDate, smCurrTempDateDHETSEStamp, "gGetEventsFromLibraries-Template DHE", tmCurrTempDateDHETSE())
    For ilDHE = 0 To UBound(tmCurrTempDateDHETSE) - 1 Step 1
        If (tmCurrTempDateDHETSE(ilDHE).tDHE.sState <> "D") And (tmCurrTempDateDHETSE(ilDHE).tDHE.sState <> "L") And (tmCurrTempDateDHETSE(ilDHE).tTSE.sState <> "D") And (tmCurrTempDateDHETSE(ilDHE).tTSE.sState <> "L") Then
            ilRet = gGetRecs_DEE_DayEvent(smCurrLibDEEStamp, tmCurrTempDateDHETSE(ilDHE).tDHE.lCode, "gGetEventsFromLibraries-Template DEE", tmCurrLibDEE())
            If (UBound(tmCurrLibDEE) > LBound(tmCurrLibDEE)) And (tmCurrTempDateDHETSE(ilDHE).tDHE.sUsedFlag <> "Y") Then
                lgLibDheUsed(UBound(lgLibDheUsed)) = tmCurrTempDateDHETSE(ilDHE).tDHE.lCode
                ReDim Preserve lgLibDheUsed(0 To UBound(lgLibDheUsed) + 1) As Long
            End If
            mTransferTSEDEEToSEE slAirDate, tmCurrTempDateDHETSE(ilDHE).tTSE, tmCurrLibDEE(), tmCurrTempDateDHETSE(ilDHE).tDHE.lOrigDHECode, tgCurrSEE()
        End If
    Next ilDHE
    Erase tmCurrTempDateDHETSE
    Erase tmCurrLibDateDHE
    Erase tmCurrEBE
    Erase tmCurrLibDEE
End Function
Public Function gGetEventsFromLibrariesHourRange(slAirDate As String, ilStartHour As Integer, ilEndHour As Integer) As Integer
    Dim ilRet As Integer
    Dim ilDHE As Integer
    Dim llUpper As Long
    Dim ilDEE As Integer
    Dim ilHours As Integer
    Dim ilDay As Integer
    Dim ilEBE As Integer
    Dim slAirDateMinus1 As String
    Dim ilHour As Integer
    
    gGetEventsFromLibrariesHourRange = True
    ReDim tgCurrSEE(0 To 0) As SEE
    ReDim lgLibDheUsed(0 To 0) As Long
    ilDay = Weekday(slAirDate, vbMonday)
    ilRet = gGetTypeOfRecs_DHE_DayHeaderInfoForLibByDate("C", "A", slAirDate, smCurrLibDateDHEStamp, "gGetEventsFromLibraries- Library DHE", tmCurrLibDateDHE())
    For ilDHE = 0 To UBound(tmCurrLibDateDHE) - 1 Step 1
        For ilHour = ilStartHour To ilEndHour Step 1
            If Mid$(tmCurrLibDateDHE(ilDHE).sHours, ilHour + 1, 1) = "Y" Then
                ilRet = gGetRecs_DEE_DayEvent(smCurrLibDEEStamp, tmCurrLibDateDHE(ilDHE).lCode, "gGetEventsFromLibraries-Library DEE", tmCurrLibDEE())
                If (UBound(tmCurrLibDEE) > LBound(tmCurrLibDEE)) And (tmCurrLibDateDHE(ilDHE).sUsedFlag <> "Y") Then
                    lgLibDheUsed(UBound(lgLibDheUsed)) = tmCurrLibDateDHE(ilDHE).lCode
                    ReDim Preserve lgLibDheUsed(0 To UBound(lgLibDheUsed) + 1) As Long
                End If
                mTransferDEEToSEE slAirDate, tmCurrLibDEE(), tmCurrLibDateDHE(ilDHE).lOrigDHECode, tmCurrLibDateDHE(ilDHE).sIgnoreConflicts, tgCurrSEE()
                Exit For
            End If
        Next ilHour
    Next ilDHE
    smCurrTempDateDHETSEStamp = ""
    slAirDateMinus1 = DateAdd("d", -1, slAirDate)
    ilRet = gGetTypeOfRecs_DHETSE_DayHeaderInfoForTempByRange("C", slAirDateMinus1, slAirDate, smCurrTempDateDHETSEStamp, "gGetEventsFromLibraries-Template DHE", tmCurrTempDateDHETSE())
    For ilDHE = 0 To UBound(tmCurrTempDateDHETSE) - 1 Step 1
        If (tmCurrTempDateDHETSE(ilDHE).tDHE.sState <> "D") And (tmCurrTempDateDHETSE(ilDHE).tDHE.sState <> "L") And (tmCurrTempDateDHETSE(ilDHE).tTSE.sState <> "D") And (tmCurrTempDateDHETSE(ilDHE).tTSE.sState <> "L") Then
            ilRet = gGetRecs_DEE_DayEvent(smCurrLibDEEStamp, tmCurrTempDateDHETSE(ilDHE).tDHE.lCode, "gGetEventsFromLibraries-Template DEE", tmCurrLibDEE())
            If (UBound(tmCurrLibDEE) > LBound(tmCurrLibDEE)) And (tmCurrTempDateDHETSE(ilDHE).tDHE.sUsedFlag <> "Y") Then
                lgLibDheUsed(UBound(lgLibDheUsed)) = tmCurrTempDateDHETSE(ilDHE).tDHE.lCode
                ReDim Preserve lgLibDheUsed(0 To UBound(lgLibDheUsed) + 1) As Long
            End If
            mTransferTSEDEEToSEE slAirDate, tmCurrTempDateDHETSE(ilDHE).tTSE, tmCurrLibDEE(), tmCurrTempDateDHETSE(ilDHE).tDHE.lOrigDHECode, tgCurrSEE()
        End If
    Next ilDHE
    Erase tmCurrTempDateDHETSE
    Erase tmCurrLibDateDHE
    Erase tmCurrEBE
    Erase tmCurrLibDEE
End Function
Public Sub gSetCTE(slComment As String, slType As String, tlCTE As CTE)
    Dim slNowDate As String
    Dim slNowTime As String
    
    slNowDate = Format(Now, sgShowDateForm)
    slNowTime = Format(Now, sgShowTimeWSecForm)
    tlCTE.lCode = 0
    tlCTE.sComment = slComment
    tlCTE.sState = "A"
    tlCTE.sType = slType    '"DH" or "T1"
    tlCTE.sUsedFlag = "Y"
    tlCTE.iVersion = 0
    tlCTE.lOrigCteCode = 0
    tlCTE.sCurrent = "Y"
    tlCTE.sEnteredDate = slNowDate
    tlCTE.sEnteredTime = slNowTime
    tlCTE.iUieCode = tgUIE.iCode
    tlCTE.sUnused = ""

End Sub

Public Sub gSetUsedFlags(tlSEE As SEE, hlCTE As Integer)
    Dim ilRet As Integer
    
    ilRet = gPutUpdate_ANE_UsedFlag(tlSEE.iBkupAneCode, tgCurrANE())
    DoEvents
    ilRet = gPutUpdate_ANE_UsedFlag(tlSEE.iProtAneCode, tgCurrANE())
    DoEvents
    ilRet = gPutUpdate_ASE_UsedFlag(tlSEE.iAudioAseCode, tgCurrASE())
    DoEvents
    ilRet = gPutUpdate_BDE_UsedFlag(tlSEE.iBdeCode, tgCurrBDE())
    DoEvents
    ilRet = gPutUpdate_CCE_UsedFlag(tlSEE.iAudioCceCode, tgCurrAudioCCE())
    DoEvents
    ilRet = gPutUpdate_CCE_UsedFlag(tlSEE.iBkupCceCode, tgCurrAudioCCE())
    DoEvents
    ilRet = gPutUpdate_CCE_UsedFlag(tlSEE.iProtCceCode, tgCurrAudioCCE())
    DoEvents
    ilRet = gPutUpdate_CCE_UsedFlag(tlSEE.iBusCceCode, tgCurrBusCCE())
    DoEvents
    'ilRet = gPutUpdate_CTE_UsedFlag(tlSEE.l2CteCode, tgCurrCTE(), hlCTE)
    'DoEvents
    ilRet = gPutUpdate_ETE_UsedFlag(tlSEE.iEteCode, tgCurrETE())
    DoEvents
    ilRet = gPutUpdate_FNE_UsedFlag(tlSEE.iFneCode, tgCurrFNE())
    DoEvents
    ilRet = gPutUpdate_MTE_UsedFlag(tlSEE.iMteCode, tgCurrMTE())
    DoEvents
    ilRet = gPutUpdate_NNE_UsedFlag(tlSEE.iEndNneCode, tgCurrNNE())
    DoEvents
    ilRet = gPutUpdate_NNE_UsedFlag(tlSEE.iStartNneCode, tgCurrNNE())
    DoEvents
    ilRet = gPutUpdate_RNE_UsedFlag(tlSEE.i1RneCode, tgCurrRNE())
    DoEvents
    ilRet = gPutUpdate_RNE_UsedFlag(tlSEE.i2RneCode, tgCurrRNE())
    DoEvents
    ilRet = gPutUpdate_SCE_UsedFlag(tlSEE.i1SceCode, tgCurrSCE())
    DoEvents
    ilRet = gPutUpdate_SCE_UsedFlag(tlSEE.i2SceCode, tgCurrSCE())
    DoEvents
    ilRet = gPutUpdate_SCE_UsedFlag(tlSEE.i3SceCode, tgCurrSCE())
    DoEvents
    ilRet = gPutUpdate_SCE_UsedFlag(tlSEE.i4SceCode, tgCurrSCE())
    DoEvents
    ilRet = gPutUpdate_TTE_UsedFlag(tlSEE.iEndTteCode, tgCurrEndTTE())
    DoEvents
    ilRet = gPutUpdate_TTE_UsedFlag(tlSEE.iStartTteCode, tgCurrStartTTE())
    DoEvents
End Sub

Public Function gExportStrLength() As Integer
    Dim ilLength As Integer
    ilLength = 0
    If tgStartColAFE.iBus + tgNoCharAFE.iBus > ilLength Then
        ilLength = tgStartColAFE.iBus + tgNoCharAFE.iBus - 1
    End If
    If tgStartColAFE.iBusControl + tgNoCharAFE.iBusControl > ilLength Then
        ilLength = tgStartColAFE.iBusControl + tgNoCharAFE.iBusControl - 1
    End If
    If tgStartColAFE.iEventType + tgNoCharAFE.iEventType > ilLength Then
        ilLength = tgStartColAFE.iEventType + tgNoCharAFE.iEventType - 1
    End If
    If tgStartColAFE.iTime + tgNoCharAFE.iTime > ilLength Then
        ilLength = tgStartColAFE.iTime + tgNoCharAFE.iTime - 1
    End If
    If tgStartColAFE.iStartType + tgNoCharAFE.iStartType > ilLength Then
        ilLength = tgStartColAFE.iStartType + tgNoCharAFE.iStartType - 1
    End If
    If tgStartColAFE.iFixedTime + tgNoCharAFE.iFixedTime > ilLength Then
        ilLength = tgStartColAFE.iFixedTime + tgNoCharAFE.iFixedTime - 1
    End If
    If tgStartColAFE.iEndType + tgNoCharAFE.iEndType > ilLength Then
        ilLength = tgStartColAFE.iEndType + tgNoCharAFE.iEndType - 1
    End If
    If tgStartColAFE.iDuration + tgNoCharAFE.iDuration > ilLength Then
        ilLength = tgStartColAFE.iDuration + tgNoCharAFE.iDuration - 1
    End If
    If tgStartColAFE.iEndTime + tgNoCharAFE.iEndTime > ilLength Then
        ilLength = tgStartColAFE.iEndTime + tgNoCharAFE.iEndTime - 1
    End If
    If tgStartColAFE.iMaterialType + tgNoCharAFE.iMaterialType > ilLength Then
        ilLength = tgStartColAFE.iMaterialType + tgNoCharAFE.iMaterialType - 1
    End If
    If tgStartColAFE.iAudioName + tgNoCharAFE.iAudioName > ilLength Then
        ilLength = tgStartColAFE.iAudioName + tgNoCharAFE.iAudioName - 1
    End If
    If tgStartColAFE.iAudioISCI + tgNoCharAFE.iAudioISCI > ilLength Then
        ilLength = tgStartColAFE.iAudioISCI + tgNoCharAFE.iAudioISCI - 1
    End If
    If tgStartColAFE.iAudioItemID + tgNoCharAFE.iAudioItemID > ilLength Then
        ilLength = tgStartColAFE.iAudioItemID + tgNoCharAFE.iAudioItemID - 1
    End If
    If tgStartColAFE.iAudioControl + tgNoCharAFE.iAudioControl > ilLength Then
        ilLength = tgStartColAFE.iAudioControl + tgNoCharAFE.iAudioControl - 1
    End If
    If tgStartColAFE.iBkupAudioName + tgNoCharAFE.iBkupAudioName > ilLength Then
        ilLength = tgStartColAFE.iBkupAudioName + tgNoCharAFE.iBkupAudioName - 1
    End If
    If tgStartColAFE.iBkupAudioControl + tgNoCharAFE.iBkupAudioControl > ilLength Then
        ilLength = tgStartColAFE.iBkupAudioControl + tgNoCharAFE.iBkupAudioControl - 1
    End If
    If tgStartColAFE.iProtAudioName + tgNoCharAFE.iProtAudioName > ilLength Then
        ilLength = tgStartColAFE.iProtAudioName + tgNoCharAFE.iProtAudioName - 1
    End If
    If tgStartColAFE.iProtISCI + tgNoCharAFE.iProtISCI > ilLength Then
        ilLength = tgStartColAFE.iProtISCI + tgNoCharAFE.iProtISCI - 1
    End If
    If tgStartColAFE.iProtItemID + tgNoCharAFE.iProtItemID > ilLength Then
        ilLength = tgStartColAFE.iProtItemID + tgNoCharAFE.iProtItemID - 1
    End If
    If tgStartColAFE.iProtAudioControl + tgNoCharAFE.iProtAudioControl > ilLength Then
        ilLength = tgStartColAFE.iProtAudioControl + tgNoCharAFE.iProtAudioControl - 1
    End If
    If tgStartColAFE.iRelay1 + tgNoCharAFE.iRelay1 > ilLength Then
        ilLength = tgStartColAFE.iRelay1 + tgNoCharAFE.iRelay1 - 1
    End If
    If tgStartColAFE.iRelay2 + tgNoCharAFE.iRelay2 > ilLength Then
        ilLength = tgStartColAFE.iRelay2 + tgNoCharAFE.iRelay2 - 1
    End If
    If tgStartColAFE.iFollow + tgNoCharAFE.iFollow > ilLength Then
        ilLength = tgStartColAFE.iFollow + tgNoCharAFE.iFollow - 1
    End If
    If tgStartColAFE.iSilenceTime + tgNoCharAFE.iSilenceTime > ilLength Then
        ilLength = tgStartColAFE.iSilenceTime + tgNoCharAFE.iSilenceTime - 1
    End If
    If tgStartColAFE.iSilence1 + tgNoCharAFE.iSilence1 > ilLength Then
        ilLength = tgStartColAFE.iSilence1 + tgNoCharAFE.iSilence1 - 1
    End If
    If tgStartColAFE.iSilence2 + tgNoCharAFE.iSilence2 > ilLength Then
        ilLength = tgStartColAFE.iSilence2 + tgNoCharAFE.iSilence2 - 1
    End If
    If tgStartColAFE.iSilence3 + tgNoCharAFE.iSilence3 > ilLength Then
        ilLength = tgStartColAFE.iSilence3 + tgNoCharAFE.iSilence3 - 1
    End If
    If tgStartColAFE.iSilence4 + tgNoCharAFE.iSilence4 > ilLength Then
        ilLength = tgStartColAFE.iSilence4 + tgNoCharAFE.iSilence4 - 1
    End If
    If tgStartColAFE.iStartNetcue + tgNoCharAFE.iStartNetcue > ilLength Then
        ilLength = tgStartColAFE.iStartNetcue + tgNoCharAFE.iStartNetcue - 1
    End If
    If tgStartColAFE.iStopNetcue + tgNoCharAFE.iStopNetcue > ilLength Then
        ilLength = tgStartColAFE.iStopNetcue + tgNoCharAFE.iStopNetcue - 1
    End If
    If tgStartColAFE.iTitle1 + tgNoCharAFE.iTitle1 > ilLength Then
        ilLength = tgStartColAFE.iTitle1 + tgNoCharAFE.iTitle1 - 1
    End If
    If tgStartColAFE.iTitle2 + tgNoCharAFE.iTitle2 > ilLength Then
        ilLength = tgStartColAFE.iTitle2 + tgNoCharAFE.iTitle2 - 1
    End If
    If tgStartColAFE.iABCFormat + tgNoCharAFE.iABCFormat > ilLength Then
        ilLength = tgStartColAFE.iABCFormat + tgNoCharAFE.iABCFormat - 1
    End If
    If tgStartColAFE.iABCPgmCode + tgNoCharAFE.iABCPgmCode > ilLength Then
        ilLength = tgStartColAFE.iABCPgmCode + tgNoCharAFE.iABCPgmCode - 1
    End If
    If tgStartColAFE.iABCXDSMode + tgNoCharAFE.iABCXDSMode > ilLength Then
        ilLength = tgStartColAFE.iABCXDSMode + tgNoCharAFE.iABCXDSMode - 1
    End If
    If tgStartColAFE.iABCRecordItem + tgNoCharAFE.iABCRecordItem > ilLength Then
        ilLength = tgStartColAFE.iABCRecordItem + tgNoCharAFE.iABCRecordItem - 1
    End If
    If tgStartColAFE.iEventID + tgNoCharAFE.iEventID > ilLength Then
        ilLength = tgStartColAFE.iEventID + tgNoCharAFE.iEventID - 1
    End If
    If tgStartColAFE.iDate + tgNoCharAFE.iDate > ilLength Then
        ilLength = tgStartColAFE.iDate + tgNoCharAFE.iDate - 1
    End If
    gExportStrLength = ilLength
End Function

Public Sub gGetAuto()
    Dim ilRet As Integer
    Dim ilCountActive As Integer
    Dim ilLoop As Integer
    
    ilRet = gGetTypeOfRecs_AEE_AutoEquip("C", sgCurrAEEStamp, "EngrAutomation-mPopulate", tgCurrAEE())
    ilRet = gGetTypeOfRecs_ACE_AutoContact("C", sgCurrACEStamp, "EngrEventType-mPopulate", tgCurrACE())
    ilRet = gGetTypeOfRecs_ADE_AutoDataFlags("C", sgCurrADEStamp, "EngrEventType-mPopulate", tgCurrADE())
    ilRet = gGetTypeOfRecs_AFE_AutoFormat("C", sgCurrAFEStamp, "EngrEventType-mPopulate", tgCurrAFE())
    ilRet = gGetTypeOfRecs_APE_AutoPath("C", sgCurrAPEStamp, "EngrEventType-mPopulate", tgCurrAPE())
    ilCountActive = 0
    For ilLoop = 0 To UBound(tgCurrAEE) - 1 Step 1
        If tgCurrAEE(ilLoop).sState = "A" Then
            ilCountActive = ilCountActive + 1
        End If
    Next ilLoop
    If ilCountActive = 1 Then
        LSet tgAEE = tgCurrAEE(0)
        LSet tgACE = tgCurrACE(0)
        LSet tgADE = tgCurrADE(0)
        If tgCurrAFE(0).sSubType = "S" Then
            LSet tgStartColAFE = tgCurrAFE(0)
        ElseIf tgCurrAFE(1).sSubType = "S" Then
            LSet tgStartColAFE = tgCurrAFE(1)
        End If
        If tgCurrAFE(0).sSubType = "N" Then
            LSet tgNoCharAFE = tgCurrAFE(0)
        ElseIf tgCurrAFE(1).sSubType = "N" Then
            LSet tgNoCharAFE = tgCurrAFE(1)
        End If
        
        'LSet tgAPE = tgCurrAPE(0)
        For ilLoop = 0 To UBound(tgCurrAPE) - 1 Step 1
            If (Not igTestSystem) And ((tgCurrAPE(ilLoop).sSubType = "P") Or (Trim$(tgCurrAPE(ilLoop).sSubType) = "")) Then
                LSet tgAPE = tgCurrAPE(ilLoop)
                Exit For
            ElseIf (igTestSystem) And (tgCurrAPE(ilLoop).sSubType = "T") Then
                LSet tgAPE = tgCurrAPE(ilLoop)
                Exit For
            End If
        Next ilLoop
    Else
        gInitMaxAFE
    End If

End Sub

Public Sub gGetSiteOption()
    Dim ilRet As Integer
    
    On Error GoTo gGetSiteOptionErr
    ilRet = gGetTypeOfRecs_SOE_SiteOption("C", sgCurrSOEStamp, "Start Up-mChkForSiteOption", tgCurrSOE())
    LSet tgSOE = tgCurrSOE(0)
    ilRet = gGetRecs_ITE_ItemTest(sgCurrITEStamp, tgSOE.iCode, "Start Up-mChkForSiteOpion: Get ITE", tgCurrITE())
    ilRet = gGetRecs_SGE_SiteGenSchd(sgCurrSGEStamp, tgSOE.iCode, "Start Up-mChkForSiteOpion: Get SGE", tgCurrSGE())
    ilRet = gGetRecs_SPE_SitePath(sgCurrSPEStamp, tgSOE.iCode, "Start Up-mChkForSiteOpion: Get SPE", tgCurrSPE())
    ilRet = gGetRec_SSE_Site_SMTP_Info(tgSOE.iCode, "Start Up-mChkForSiteOpion: Get SSE", tgCurrSSE)
    Exit Sub
gGetSiteOptionErr:
    ReDim tgCurrSOE(0 To 0) As SOE
    ReDim tgCurrITE(0 To 0) As ITE
    ReDim tgCurrSGE(0 To 0) As SGE
    ReDim tgCurrSPE(0 To 0) As SPE
    Exit Sub

End Sub

Public Function gAdjustSEE(tlSchdChgInfo As SCHDCHGINFO, hlSEE As Integer, hlSOE As Integer, ilSpotRomoved As Integer, tlUPDSee() As SEE) As Integer
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim ilSHE As Integer
    Dim llNewChgStartDate As Long
    Dim llNewChgEndDate As Long
    Dim llCheckStartDate As Long
    Dim llCheckEndDate As Long
    Dim llSplitStartDate As Long
    Dim llSplitEndDate As Long
    Dim llExpandStartDate As Long
    Dim llExpandEndDate As Long
    Dim slAirDate As String
    Dim llAirDate As Long
    Dim llSEE As Long
    Dim llCheckSEE As Long
    Dim llSplitSEE As Long
    Dim llNewSEE As Long
    Dim llExpandSEE As Long
    Dim ilMatchFound As Integer
    Dim ilETE As Integer
    Dim ilSpotETECode As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilAdd As Integer
    Dim ilTSE As Integer
    Dim ilStartTSE As Integer
    Dim tlTSE As TSE
    Dim ilProcessDate As Integer
    Dim slCategory As String
    Dim llOldSHECode As Long
    Dim blBypassRemove As Boolean
    ReDim llBuildNewLoadDate(0 To 0) As Long
    ReDim llNewCodesAdded(0 To 0) As Long
    
    ReDim tlUPDSee(0 To 0) As SEE
    ilSpotRomoved = 0
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = gDateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    ilRet = gGetTypeOfRecs_SHE_ScheduleHeaderByDate(slNowDate, "gAdjust SEE, get SHE", tmAdjSHE())
    If Not ilRet Then
        gAdjustSEE = False
        Exit Function
    End If
    If UBound(tmAdjSHE) <= LBound(tmAdjSHE) Then
        gAdjustSEE = True
        Exit Function
    End If
    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "gAdjustSEE-Get ETE", tgCurrETE())
    ilSpotETECode = 0
    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        If tgCurrETE(ilETE).sCategory = "S" Then
            ilSpotETECode = tgCurrETE(ilETE).iCode
            Exit For
        End If
    Next ilETE
    'Get DHE records so that the Value dates can be checked
    llNewChgStartDate = 0
    llNewChgEndDate = 0
    If tlSchdChgInfo.lNewChgDHE > 0 Then
        ilRet = gGetRec_DHE_DayHeaderInfo(tlSchdChgInfo.lNewChgDHE, smNewChgStamp, tmNewChgDHE)
        If ilRet Then
            llNewChgStartDate = gDateValue(tmNewChgDHE.sStartDate)
            llNewChgEndDate = gDateValue(tmNewChgDHE.sEndDate)
        End If
    End If
    llCheckStartDate = 0
    llCheckEndDate = 0
    If tlSchdChgInfo.lCheckDHE > 0 Then
        ilRet = gGetRec_DHE_DayHeaderInfo(tlSchdChgInfo.lCheckDHE, smCheckStamp, tmCheckDHE)
        If ilRet Then
            If (tmCheckDHE.sState <> "D") And (tmCheckDHE.sCurrent <> "N") Then
                llCheckStartDate = gDateValue(tmCheckDHE.sStartDate)
                llCheckEndDate = gDateValue(tmCheckDHE.sEndDate)
            End If
        End If
    End If
    llSplitStartDate = 0
    llSplitEndDate = 0
    If tlSchdChgInfo.lSplitDHE > 0 Then
        ilRet = gGetRec_DHE_DayHeaderInfo(tlSchdChgInfo.lSplitDHE, smSplitStamp, tmSplitDHE)
        If ilRet Then
            llSplitStartDate = gDateValue(tmSplitDHE.sStartDate)
            llSplitEndDate = gDateValue(tmSplitDHE.sEndDate)
        End If
    End If
    llExpandStartDate = 0
    llExpandEndDate = 0
    If tlSchdChgInfo.lExpandDHE > 0 Then
        ilRet = gGetRec_DHE_DayHeaderInfo(tlSchdChgInfo.lExpandDHE, smExpandStamp, tmExpandDHE)
        If ilRet Then
            llExpandStartDate = gDateValue(tmExpandDHE.sStartDate)
            llExpandEndDate = gDateValue(tmExpandDHE.sEndDate)
        End If
    End If
    If (tmNewChgDHE.sType = "T") And (tlSchdChgInfo.lNewChgDHE > 0) Then
        ilRet = gGetRecs_TSE_TemplateSchd(smCurrTempTSEStamp, tlSchdChgInfo.lNewChgDHE, "EngrTempDef-mPopulate for TSE", tmCurrTempTSE())
    End If
    If tlSchdChgInfo.lNewChgDHE > 0 Then
        ilRet = gGetRecs_DEE_DayEvent(smNewChgDEEStamp, tlSchdChgInfo.lNewChgDHE, "gAdjustSEE, Get New/Chg DEE", tmNewChgDEE())
    End If
    If tlSchdChgInfo.lSplitDHE > 0 Then
        ilRet = gGetRecs_DEE_DayEvent(smSplitDEEStamp, tlSchdChgInfo.lSplitDHE, "gAdjustSEE, Get Split DEE", tmSplitDEE())
    End If
    If tlSchdChgInfo.lCheckDHE > 0 Then
        ilRet = gGetRecs_DEE_DayEvent(smCheckDEEStamp, tlSchdChgInfo.lCheckDHE, "gAdjustSEE, Get Check DEE", tmCheckDEE())
    End If
    If tlSchdChgInfo.lExpandDHE > 0 Then
        ilRet = gGetRecs_DEE_DayEvent(smCheckDEEStamp, tlSchdChgInfo.lExpandDHE, "gAdjustSEE, Get Expand DEE", tmExpandDEE())
    End If
    For ilSHE = 0 To UBound(tmAdjSHE) - 1 Step 1
        'Select * FROM "SEE_Schedule_Events" WHERE seedeeCode IN (SELECT seedeeCode from "SEE_Schedule_Events", "DEE_Day_Event_Info", "DHE_Day_Header_Info" where seedeecode = deecode and deedhecode = dhecode and dheCode = 3)
        'The Inner select is done first building a temp table.  Then the outer select is done with the IN part change to an OR
        'select * from "SEE_Schedule_Events" WHERE seeDEECode = XXX or seeDEECode = YYY
        'The XXX and YYY are the values stored into the temp table
        'Select * FROM "SEE_Schedule_Events" WHERE seedeeCode IN (SELECT seedeeCode from "SEE_Schedule_Events", "DEE_Day_Event_Info", "DHE_Day_Header_Info" where seedeecode = deecode and deedhecode = dhecode and dheCurrent = 'Y' and dheCode = 3) Order by seeCode
        slAirDate = tmAdjSHE(ilSHE).sAirDate
        llAirDate = gDateValue(slAirDate)
        ilStartTSE = 0
        Do
            ilProcessDate = 0
            If (tmNewChgDHE.sType = "T") And (tlSchdChgInfo.lNewChgDHE > 0) Then
                For ilTSE = ilStartTSE To UBound(tmCurrTempTSE) - 1 Step 1
                    If (gDateValue(tmCurrTempTSE(ilTSE).sLogDate) = llAirDate) Or (gDateValue(tmCurrTempTSE(ilTSE).sLogDate) + 1 = llAirDate) Then
                        llNewChgStartDate = llAirDate
                        llNewChgEndDate = llAirDate
                        If tlSchdChgInfo.lCheckDHE > 0 Then
                            llCheckStartDate = 0
                            llCheckEndDate = 0
                            'Get TSE
                            ilRet = gGetRec_TSE_TemplateSchdByDHETSE(tlSchdChgInfo.lCheckDHE, tmCurrTempTSE(ilTSE).lOrigTseCode, "gAdjustSEE, Get TSE for Check DHE", tlTSE)
                        End If
                        ilProcessDate = 2
                        If (tmNewChgDHE.sState = "D") Or (tmCurrTempTSE(ilTSE).sState = "D") Then
                            ilProcessDate = 3
                        End If
                        ilStartTSE = ilTSE + 1
                        Exit For
                    End If
                Next ilTSE
            Else
                ilProcessDate = 1
            End If
            If ilProcessDate = 0 Then
                Exit Do
            End If
            ReDim tmNewChgSEE(0 To 0) As SEE
            If tlSchdChgInfo.lNewChgDHE > 0 Then
                If (tmNewChgDHE.sType = "T") Then
                    mTransferTSEDEEToSEE slAirDate, tmCurrTempTSE(ilStartTSE - 1), tmNewChgDEE(), tmNewChgDHE.lOrigDHECode, tmNewChgSEE()
                Else
                    mTransferDEEToSEE slAirDate, tmNewChgDEE(), tmNewChgDHE.lOrigDHECode, tmNewChgDHE.sIgnoreConflicts, tmNewChgSEE()
                End If
            End If
            ReDim tmSplitSEE(0 To 0) As SEE
            If tlSchdChgInfo.lSplitDHE > 0 Then
                mTransferDEEToSEE slAirDate, tmSplitDEE(), tmSplitDHE.lOrigDHECode, tmSplitDHE.sIgnoreConflicts, tmSplitSEE()
            End If
            ReDim tmCheckSEE(0 To 0) As SEE
            If tlSchdChgInfo.lCheckDHE > 0 Then
                If (tmCheckDHE.sType = "T") And (tlSchdChgInfo.lCheckDHE > 0) Then
                    If tlTSE.lCode > 0 Then
                        mTransferTSEDEEToSEE slAirDate, tlTSE, tmCheckDEE(), tmNewChgDHE.lOrigDHECode, tmCheckSEE()
                    Else
                        ReDim tmCheckSEE(0 To 0) As SEE
                    End If
                Else
                    mTransferDEEToSEE slAirDate, tmCheckDEE(), tmCheckDHE.lOrigDHECode, tmCheckDHE.sIgnoreConflicts, tmCheckSEE()
                End If
            End If
            ReDim tmExpandSEE(0 To 0) As SEE
            If tlSchdChgInfo.lExpandDHE > 0 Then
                mTransferDEEToSEE slAirDate, tmExpandDEE(), tmExpandDHE.lOrigDHECode, tmExpandDHE.sIgnoreConflicts, tmExpandSEE()
            End If
            'Loop by schedule dates
            '  Set flag in SplitDHE, NewChgDHE and ExpandDHE SEE records indicating not processed
            '  Get SEE by date and matching lCheckDHE
            '  Loop thru SEE looking for matching tmCheckSEE
            '    If so leave it
            '    If not see if a matching event exist in tmNewChgSEE or in tmSplitSEE (first check date, then event)
            '    If match found, then change DEE reference to DHE only
            '    If no match found, then delete and set flag to re-gnerate schedule date
            '  End Loop
            '  Get SEE by date and matching SplitDHE
            '  Loop thru SEE looking for matched tmSplitSEE
            '    If Match, leave
            '    If no match found, remove
            '  End Loop
            '  Add any tmSplitSEE not found
            '  Get SEE by date and matching NewChgDHE
            '  Loop thru SEE looking for match tmNewChgSEE
            '    If Match, Update if changed or leave if not changed
            '    If no match found, remove
            '  End Loop
            '  Add any tmNewChgSEE not found
            '  Get SEE by date and matching tmExpandDHE (none should be found)
            '    If Match, Update if changed or leave if not changed
            '    If no match found, remove
            '  End Loop
            '  Add any tmExpandDHE not found
            'If not found, add
            'ilRet = gGetRecs_SEE_ScheduleEventsByDHEandSHE(smAdjSEEStamp, tlSchdChgInfo.lCheckDHE, tmAdjSHE(ilSHE).lCode, "gAdjust SEE, Get SEE", tmAdjSEE())
            ilRet = gGetRecs_SEE_ScheduleEventsAPI(hlSEE, smAdjSEEStamp, tlSchdChgInfo.lCheckDHE, tmAdjSHE(ilSHE).lCode, "gAdjust SEE, Get SEE", tmAdjSEE())
            If UBound(tmAdjSEE) <= LBound(tmAdjSEE) Then
                ilRet = gGetRecs_SEE_ScheduleEventsAPI(hlSEE, smAdjSEEStamp, tmNewChgDHE.lOrigDHECode, tmAdjSHE(ilSHE).lCode, "gAdjust SEE, Get SEE", tmAdjSEE())
            End If
            If UBound(tmCheckSEE) > LBound(tmCheckSEE) Then
                For llSEE = 0 To UBound(tmAdjSEE) - 1 Step 1
                    'Bypass Spot Event Types
                    If tmAdjSEE(llSEE).iEteCode <> ilSpotETECode Then
                        slCategory = ""
                        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                            If tgCurrETE(ilETE).iCode = tmAdjSEE(llSEE).iEteCode Then
                                slCategory = tgCurrETE(ilETE).sCategory
                                Exit For
                            End If
                        Next ilETE
                        ilMatchFound = False
                        If ilProcessDate <> 3 Then
                            If (llAirDate >= llCheckStartDate) And (llAirDate <= llCheckEndDate) Then
                                For llCheckSEE = LBound(tmCheckSEE) To UBound(tmCheckSEE) - 1 Step 1
                                    If mCompareSEE(tmAdjSEE(llSEE), tmCheckSEE(llCheckSEE)) Then
                                        ilMatchFound = True
                                        If tmAdjSEE(llSEE).lDeeCode <> tmCheckSEE(llCheckSEE).lDeeCode Then
                                            tmAdjSEE(llSEE).lDheCode = tmCheckSEE(llCheckSEE).lDheCode
                                            tmAdjSEE(llSEE).lDeeCode = tmCheckSEE(llCheckSEE).lDeeCode
                                            ilRet = gPutUpdate_SEE_DHEDEECode(tmAdjSEE(llSEE).lCode, tmAdjSEE(llSEE).lDheCode, tmAdjSEE(llSEE).lDeeCode, "gAdjustSEE, Update SEE")
                                        End If
                                        Exit For
                                    End If
                                Next llCheckSEE
                            End If
                            If Not ilMatchFound Then
                                If (llAirDate >= llNewChgStartDate) And (llAirDate <= llNewChgEndDate) Then
                                    For llNewSEE = LBound(tmNewChgSEE) To UBound(tmNewChgSEE) - 1 Step 1
                                        If mCompareSEE(tmAdjSEE(llSEE), tmNewChgSEE(llNewSEE)) Then
                                            'Update DEE
                                            tmAdjSEE(llSEE).lDheCode = tmNewChgSEE(llNewSEE).lDheCode
                                            tmAdjSEE(llSEE).lDeeCode = tmNewChgSEE(llNewSEE).lDeeCode
                                            ilMatchFound = True
                                            ilRet = gPutUpdate_SEE_DHEDEECode(tmAdjSEE(llSEE).lCode, tmAdjSEE(llSEE).lDheCode, tmAdjSEE(llSEE).lDeeCode, "gAdjustSEE, Update SEE")
                                            Exit For
                                        End If
                                    Next llNewSEE
                                End If
                            End If
                            If Not ilMatchFound Then
                                If (llAirDate >= llSplitStartDate) And (llAirDate <= llSplitEndDate) Then
                                    For llSplitSEE = LBound(tmSplitSEE) To UBound(tmSplitSEE) - 1 Step 1
                                        If mCompareSEE(tmAdjSEE(llSEE), tmSplitSEE(llSplitSEE)) Then
                                            'Update DEE
                                            tmAdjSEE(llSEE).lDheCode = tmSplitSEE(llSplitSEE).lDheCode
                                            tmAdjSEE(llSEE).lDeeCode = tmSplitSEE(llSplitSEE).lDeeCode
                                            ilMatchFound = True
                                            ilRet = gPutUpdate_SEE_DHEDEECode(tmAdjSEE(llSEE).lCode, tmAdjSEE(llSEE).lDheCode, tmAdjSEE(llSEE).lDeeCode, "gAdjustSEE, Update SEE")
                                            Exit For
                                        End If
                                    Next llSplitSEE
                                End If
                            End If
                        End If
                        If Not ilMatchFound Then
                            'Remove record
                            If tmAdjSEE(llSEE).sAction <> "D" Then
                                'If tmAdjSEE(llSEE).sSentStatus = "S" Then
                                '    If slCategory = "A" Then
                                '        ilSpotRomoved = 1
                                '    End If
                                '    ilRet = gPutUpdate_SEE_UnsentFlag(tmAdjSEE(llSEE).lCode, "D", "Schedule Definition-mSave: SEE")
                                '    ilFound = False
                                '    For ilLoop = 0 To UBound(llBuildNewLoadDate) - 1 Step 1
                                '        If llAirDate = llBuildNewLoadDate(ilLoop) Then
                                '            ilFound = True
                                '            Exit For
                                '        End If
                                '    Next ilLoop
                                '    If Not ilFound Then
                                '        llBuildNewLoadDate(UBound(llBuildNewLoadDate)) = llAirDate
                                '        ReDim Preserve llBuildNewLoadDate(0 To UBound(llBuildNewLoadDate) + 1) As Long
                                '    End If
                                'Else
                                '    If ilSpotRomoved = 0 Then
                                '        If slCategory = "A" Then
                                '            ilSpotRomoved = 2
                                '        End If
                                '    End If
                                '    'Delete record as it has not been sent
                                '    ilRet = gPutUpdate_SEE_UnsentFlag(tmAdjSEE(llSEE).lCode, "R", "Schedule Definition-mSave: SEE")
                                'End If
                                blBypassRemove = False
                                For llNewSEE = LBound(llNewCodesAdded) To UBound(llNewCodesAdded) - 1 Step 1
                                    If tmAdjSEE(llSEE).lCode = llNewCodesAdded(llNewSEE) Then
                                        blBypassRemove = True
                                        Exit For
                                    End If
                                Next llNewSEE
                                If Not blBypassRemove Then
                                    mRemoveEvent llSEE, llAirDate, slCategory, ilSpotRomoved, llBuildNewLoadDate()
                                End If
                            End If
                        Else
                            llNewCodesAdded(UBound(llNewCodesAdded)) = tmAdjSEE(llSEE).lCode
                            ReDim Preserve llNewCodesAdded(0 To UBound(llNewCodesAdded) + 1) As Long
                        End If
                    Else
                        If ilProcessDate = 3 Then
                            blBypassRemove = False
                            For llNewSEE = LBound(llNewCodesAdded) To UBound(llNewCodesAdded) - 1 Step 1
                                If tmAdjSEE(llSEE).lCode = llNewCodesAdded(llNewSEE) Then
                                    blBypassRemove = True
                                    Exit For
                                End If
                            Next llNewSEE
                            If Not blBypassRemove Then
                                mRemoveEvent llSEE, llAirDate, "S", ilSpotRomoved, llBuildNewLoadDate()
                            End If
                        End If
                    End If
                Next llSEE
                ilRet = mUpdateSpots(tmAdjSEE(), tmCheckSEE(), ilSpotETECode)
                ilRet = mUpdateSpots(tmAdjSEE(), tmNewChgSEE(), ilSpotETECode)
            End If
            If (UBound(tmNewChgSEE) > LBound(tmNewChgSEE)) And (llAirDate >= llNewChgStartDate) And (llAirDate <= llNewChgEndDate) And (ilProcessDate <> 3) Then
                'Bypass records added above
                'ilRet = gGetRecs_SEE_ScheduleEventsByDHEandSHE(smAdjSEEStamp, tlSchdChgInfo.lDEEDHE, tmAdjSHE(ilSHE).lCode, "gAdjust SEE, Get SEE", tmAdjSEE())
                ilRet = gGetRecs_SEE_ScheduleEventsAPI(hlSEE, smAdjSEEStamp, tlSchdChgInfo.lDEEDHE, tmAdjSHE(ilSHE).lCode, "gAdjust SEE, Get SEE", tmAdjSEE())
                For llNewSEE = LBound(tmNewChgSEE) To UBound(tmNewChgSEE) - 1 Step 1
                    ilMatchFound = False
                    For llSEE = 0 To UBound(tmAdjSEE) - 1 Step 1
                        If mCompareSEE(tmAdjSEE(llSEE), tmNewChgSEE(llNewSEE)) Then
                            'Update DEE
                            ilMatchFound = True
                            If tmAdjSEE(llSEE).lDeeCode <> tmNewChgSEE(llNewSEE).lDeeCode Then
                                tmAdjSEE(llSEE).lDheCode = tmNewChgSEE(llNewSEE).lDheCode
                                tmAdjSEE(llSEE).lDeeCode = tmNewChgSEE(llNewSEE).lDeeCode
                                ilRet = gPutUpdate_SEE_DHEDEECode(tmAdjSEE(llSEE).lCode, tmAdjSEE(llSEE).lDheCode, tmAdjSEE(llSEE).lDeeCode, "gAdjustSEE, Update SEE")
                            End If
                            Exit For
                        End If
                    Next llSEE
                    If Not ilMatchFound Then
                        'Add Record
                        ilAdd = True
                        If llAirDate = llNowDate Then
                            If llNowTime > tmNewChgSEE(llNewSEE).lTime Then
                                ilAdd = False
                            End If
                        ElseIf llAirDate < llNowDate Then
                            ilAdd = False
                        End If
                        If ilAdd Then
                            tmNewChgSEE(llNewSEE).lCode = 0
                            tmNewChgSEE(llNewSEE).lSheCode = tmAdjSHE(ilSHE).lCode
                            tmNewChgSEE(llNewSEE).sAction = "N"
                            tmNewChgSEE(llNewSEE).sSentStatus = "N"
                            tmNewChgSEE(llNewSEE).sSentDate = Format$("12/31/2069", sgShowDateForm)
                            ilRet = gPutInsert_SEE_ScheduleEvents(tmNewChgSEE(llNewSEE), "Schedule Definition-mSave: SEE", hlSEE, hlSOE)
                            llNewCodesAdded(UBound(llNewCodesAdded)) = tmNewChgSEE(llNewSEE).lCode
                            ReDim Preserve llNewCodesAdded(0 To UBound(llNewCodesAdded) + 1) As Long
                            If tmAdjSHE(ilSHE).sLoadedAutoStatus = "L" Then
                                LSet tlUPDSee(UBound(tlUPDSee)) = tmNewChgSEE(llNewSEE)
                                ReDim Preserve tlUPDSee(0 To UBound(tlUPDSee) + 1) As SEE
                                ilFound = False
                                For ilLoop = 0 To UBound(llBuildNewLoadDate) - 1 Step 1
                                    If llAirDate = llBuildNewLoadDate(ilLoop) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilLoop
                                If Not ilFound Then
                                    llBuildNewLoadDate(UBound(llBuildNewLoadDate)) = llAirDate
                                    ReDim Preserve llBuildNewLoadDate(0 To UBound(llBuildNewLoadDate) + 1) As Long
                                End If
                            End If
                        End If
                    Else
                        llNewCodesAdded(UBound(llNewCodesAdded)) = tmNewChgSEE(llNewSEE).lCode
                        ReDim Preserve llNewCodesAdded(0 To UBound(llNewCodesAdded) + 1) As Long
                    End If
                Next llNewSEE
                ilRet = mUpdateSpots(tmAdjSEE(), tmNewChgSEE(), ilSpotETECode)
            End If
            If (UBound(tmSplitSEE) > LBound(tmSplitSEE)) And (llAirDate >= llSplitStartDate) And (llAirDate <= llSplitEndDate) Then
                'Bypass records added above
                'ilRet = gGetRecs_SEE_ScheduleEventsByDHEandSHE(smAdjSEEStamp, tlSchdChgInfo.lSplitDHE, tmAdjSHE(ilSHE).lCode, "gAdjust SEE, Get SEE", tmAdjSEE())
                ilRet = gGetRecs_SEE_ScheduleEventsAPI(hlSEE, smAdjSEEStamp, tlSchdChgInfo.lSplitDHE, tmAdjSHE(ilSHE).lCode, "gAdjust SEE, Get SEE", tmAdjSEE())
                For llSplitSEE = LBound(tmSplitSEE) To UBound(tmSplitSEE) - 1 Step 1
                    ilMatchFound = False
                    For llSEE = 0 To UBound(tmAdjSEE) - 1 Step 1
                        If mCompareSEE(tmAdjSEE(llSEE), tmSplitSEE(llSplitSEE)) Then
                            tmAdjSEE(llSEE).lDheCode = tmSplitSEE(llSplitSEE).lDheCode
                            tmAdjSEE(llSEE).lDeeCode = tmSplitSEE(llSplitSEE).lDeeCode
                            ilMatchFound = True
                            ilRet = gPutUpdate_SEE_DHEDEECode(tmAdjSEE(llSEE).lCode, tmAdjSEE(llSEE).lDheCode, tmAdjSEE(llSEE).lDeeCode, "gAdjustSEE, Update SEE")
                            Exit For
                        End If
                    Next llSEE
                    If Not ilMatchFound Then
                        'Add Record- none should be found
                        ilAdd = True
                        If llAirDate = llNowDate Then
                            If llNowTime > tmNewChgSEE(llNewSEE).lTime Then
                                ilAdd = False
                            End If
                        ElseIf llAirDate < llNowDate Then
                            ilAdd = False
                        End If
                        If ilAdd Then
                            tmSplitSEE(llSplitSEE).lCode = 0
                            tmSplitSEE(llSplitSEE).lSheCode = tmAdjSHE(ilSHE).lCode
                            tmSplitSEE(llSplitSEE).sAction = "N"
                            tmSplitSEE(llSplitSEE).sSentStatus = "N"
                            tmSplitSEE(llSplitSEE).sSentDate = Format$("12/31/2069", sgShowDateForm)
                            ilRet = gPutInsert_SEE_ScheduleEvents(tmSplitSEE(llSplitSEE), "Schedule Definition-mSave: SEE", hlSEE, hlSOE)
                            If tmAdjSHE(ilSHE).sLoadedAutoStatus = "L" Then
                                ilFound = False
                                For ilLoop = 0 To UBound(llBuildNewLoadDate) - 1 Step 1
                                    If llAirDate = llBuildNewLoadDate(ilLoop) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilLoop
                                If Not ilFound Then
                                    llBuildNewLoadDate(UBound(llBuildNewLoadDate)) = llAirDate
                                    ReDim Preserve llBuildNewLoadDate(0 To UBound(llBuildNewLoadDate) + 1) As Long
                                End If
                            End If
                        End If
                    End If
                Next llSplitSEE
            End If
            If (UBound(tmExpandSEE) > LBound(tmExpandSEE)) And (llAirDate >= llExpandStartDate) And (llAirDate <= llExpandEndDate) Then
                'Bypass records added above
                'ilRet = gGetRecs_SEE_ScheduleEventsByDHEandSHE(smAdjSEEStamp, tlSchdChgInfo.lExpandDHE, tmAdjSHE(ilSHE).lCode, "gAdjust SEE, Get SEE", tmAdjSEE())
                ilRet = gGetRecs_SEE_ScheduleEventsAPI(hlSEE, smAdjSEEStamp, tlSchdChgInfo.lExpandDHE, tmAdjSHE(ilSHE).lCode, "gAdjust SEE, Get SEE", tmAdjSEE())
                For llExpandSEE = LBound(tmExpandSEE) To UBound(tmExpandSEE) - 1 Step 1
                    'Should not find any matches as these are new records being added
                    ilMatchFound = False
                    For llSEE = 0 To UBound(tmAdjSEE) - 1 Step 1
                        If mCompareSEE(tmAdjSEE(llSEE), tmExpandSEE(llExpandSEE)) Then
                            ilMatchFound = True
                            Exit For
                        End If
                    Next llSEE
                    If Not ilMatchFound Then
                        'Add Record
                        ilAdd = True
                        If llAirDate = llNowDate Then
                            If llNowTime > tmNewChgSEE(llNewSEE).lTime Then
                                ilAdd = False
                            End If
                        ElseIf llAirDate < llNowDate Then
                            ilAdd = False
                        End If
                        If ilAdd Then
                            tmExpandSEE(llExpandSEE).lCode = 0
                            tmExpandSEE(llExpandSEE).lSheCode = tmAdjSHE(ilSHE).lCode
                            tmExpandSEE(llExpandSEE).sAction = "N"
                            tmExpandSEE(llExpandSEE).sSentStatus = "N"
                            tmExpandSEE(llExpandSEE).sSentDate = Format$("12/31/2069", sgShowDateForm)
                            ilRet = gPutInsert_SEE_ScheduleEvents(tmExpandSEE(llExpandSEE), "Schedule Definition-mSave: SEE", hlSEE, hlSOE)
                            If tmAdjSHE(ilSHE).sLoadedAutoStatus = "L" Then
                                ilFound = False
                                For ilLoop = 0 To UBound(llBuildNewLoadDate) - 1 Step 1
                                    If llAirDate = llBuildNewLoadDate(ilLoop) Then
                                        ilFound = True
                                        Exit For
                                    End If
                                Next ilLoop
                                If Not ilFound Then
                                    llBuildNewLoadDate(UBound(llBuildNewLoadDate)) = llAirDate
                                    ReDim Preserve llBuildNewLoadDate(0 To UBound(llBuildNewLoadDate) + 1) As Long
                                End If
                            End If
                        End If
                    End If
                Next llExpandSEE
            End If
            If (ilProcessDate <> 2) And (ilProcessDate <> 3) Then
                Exit Do
            End If
        Loop
    Next ilSHE
    For ilLoop = 0 To UBound(llBuildNewLoadDate) - 1 Step 1
        slAirDate = Format$(llBuildNewLoadDate(ilLoop), sgSQLDateForm)
        ilRet = gGetRec_SHE_ScheduleHeaderByDate(slAirDate, "EngrSchedule-Get Schedule by Date", tmSHE)
        If tmSHE.sCreateLoad <> "Y" Then
            tmSHE.sCreateLoad = "Y"
            ilRet = gPutUpdate_SHE_ScheduleHeader(4, tmSHE, "EngrServiceMain: mCreateAuto- Update SHE", llOldSHECode)
        End If
    Next ilLoop
End Function


Private Sub mTransferDEEToSEE(slAirDate As String, tlDEE() As DEE, llOrigDHECode As Long, slIgnoreConflicts As String, tlSEE() As SEE)
    Dim ilRet As Integer
    Dim ilDHE As Integer
    Dim llUpper As Long
    Dim ilDEE As Integer
    Dim ilHours As Integer
    Dim ilDay As Integer
    Dim ilEBE As Integer
    
    ilDay = Weekday(slAirDate, vbMonday)
    For ilDEE = 0 To UBound(tlDEE) - 1 Step 1
        tlDEE(ilDEE).sIgnoreConflicts = slIgnoreConflicts
        smCurrEBEStamp = ""
        Erase tmCurrEBE
        ilRet = gGetRecs_EBE_EventBusSel(smCurrEBEStamp, tlDEE(ilDEE).lCode, "Bus Definition-mDEEMoveRecToCtrls", tmCurrEBE())
        If Mid(tlDEE(ilDEE).sDays, ilDay, 1) = "Y" Then
            For ilHours = 1 To 24 Step 1
                If Mid$(tlDEE(ilDEE).sHours, ilHours, 1) = "Y" Then
                    For ilEBE = 0 To UBound(tmCurrEBE) - 1 Step 1
                        llUpper = UBound(tlSEE)
                        ilRet = mCopyDEEToSEE(tlDEE(ilDEE), tmCurrEBE(ilEBE).iBdeCode, llOrigDHECode, ilHours, tlSEE(llUpper))
                        If ilRet Then
                            llUpper = llUpper + 1
                            ReDim Preserve tlSEE(0 To llUpper) As SEE
                        End If
                    Next ilEBE
                End If
            Next ilHours
        End If
    Next ilDEE
    Exit Sub
End Sub

Private Function mCompareSEE(tlAdjSEE As SEE, tlSEE As SEE) As Integer

    Dim ilSEENew As Integer
    Dim ilSEEOld As Integer
    Dim ilEBE As Integer
    Dim slStr As String
    Dim ilBDE As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    
                        
    'Buses
    If tlAdjSEE.iBdeCode <> tlSEE.iBdeCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iBusCceCode <> tlSEE.iBusCceCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iEteCode <> tlSEE.iEteCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.lTime <> tlSEE.lTime Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iStartTteCode <> tlSEE.iStartTteCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.sFixedTime <> tlSEE.sFixedTime Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iEndTteCode <> tlSEE.iEndTteCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.lDuration <> tlSEE.lDuration Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iMteCode <> tlSEE.iMteCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iAudioAseCode <> tlSEE.iAudioAseCode Then
        mCompareSEE = False
        Exit Function
    End If
    If StrComp(tlAdjSEE.sAudioItemID, tlSEE.sAudioItemID, vbTextCompare) <> 0 Then
        mCompareSEE = False
        Exit Function
    End If
    If StrComp(tlAdjSEE.sAudioISCI, tlSEE.sAudioISCI, vbTextCompare) <> 0 Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iAudioCceCode <> tlSEE.iAudioCceCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iBkupAneCode <> tlSEE.iBkupAneCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iBkupCceCode <> tlSEE.iBkupCceCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iProtAneCode <> tlSEE.iProtAneCode Then
        mCompareSEE = False
        Exit Function
    End If
    If StrComp(tlAdjSEE.sProtItemID, tlSEE.sProtItemID, vbTextCompare) <> 0 Then
        mCompareSEE = False
        Exit Function
    End If
    If StrComp(tlAdjSEE.sProtISCI, tlSEE.sProtISCI, vbTextCompare) <> 0 Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iProtCceCode <> tlSEE.iProtCceCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.i1RneCode <> tlSEE.i1RneCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.i2RneCode <> tlSEE.i2RneCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iFneCode <> tlSEE.iFneCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.lSilenceTime <> tlSEE.lSilenceTime Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.i1SceCode <> tlSEE.i1SceCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.i2SceCode <> tlSEE.i2SceCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.i3SceCode <> tlSEE.i3SceCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.i4SceCode <> tlSEE.i4SceCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iStartNneCode <> tlSEE.iStartNneCode Then
        mCompareSEE = False
        Exit Function
    End If
    If tlAdjSEE.iEndNneCode <> tlSEE.iEndNneCode Then
        mCompareSEE = False
        Exit Function
    End If
    'Comment
'    If tlAdjSEE.l1CteCode <> tlSEE.l1CteCode Then
'        mCompareSEE = False
'        Exit Function
'    End If
    '7/8/11: Make T2 work like T1
    'If tlAdjSEE.l2CteCode <> tlSEE.l2CteCode Then
    '    mCompareSEE = False
    '    Exit Function
    'End If
    If StrComp(tlAdjSEE.sABCFormat, tlSEE.sABCFormat, vbTextCompare) <> 0 Then
        mCompareSEE = False
        Exit Function
    End If
    If StrComp(tlAdjSEE.sABCPgmCode, tlSEE.sABCPgmCode, vbTextCompare) <> 0 Then
        mCompareSEE = False
        Exit Function
    End If
    If StrComp(tlAdjSEE.sABCXDSMode, tlSEE.sABCXDSMode, vbTextCompare) <> 0 Then
        mCompareSEE = False
        Exit Function
    End If
    If StrComp(tlAdjSEE.sABCRecordItem, tlSEE.sABCRecordItem, vbTextCompare) <> 0 Then
        mCompareSEE = False
        Exit Function
    End If
    mCompareSEE = True
    Exit Function
End Function

Public Function gMerge(ilFrom As Integer, slAirDate As String, hlMerge As Integer, hlMsg As Integer, tlCurrSEE() As SEE, slT1Comment() As String, slT2Comment() As String, lbcCommercialSort As ListBox, ilMergeError As Integer) As Integer
    '
    '  ilFrom(I)- 0 = EngrService; 1=EngrSchd
    '
    Dim ilRet As Integer
    Dim ilEof As Integer
    Dim slLine As String
    Dim slDate As String
    Dim llAirDate As Long
    Dim slTime As String
    Dim llTime As Long
    Dim slTitle As String
    Dim slLen As String
    Dim slBus As String
    Dim slCopy As String
    Dim slISCI As String
    Dim llLoop As Long
    Dim ilETE As Integer
    Dim ilBDE As Integer
    Dim ilBus As Integer
    Dim llRow As Long
    Dim llUpper As Long
    Dim ilFound As Integer
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim ilRemove As Integer
    Dim ilFindMatch As Integer
    Dim llAvailLength As Long
    Dim llCheck As Long
    Dim llCounter As Long
    Dim slCounter As String
    Dim slStr As String
    Dim ilSpotETECode As Integer
    Dim llLastPAIndex As Long
    Dim ilPass As Integer
    Dim ilBdeCode As Integer
    Dim ilOverbooked As Integer
    Dim ilUnderbooked As Integer
    Dim ilAvailNotFound As Integer
    Dim llTest As Long
    Dim tlSHE As SHE
    
    gMerge = True
    ilMergeError = False
    llAirDate = gDateValue(slAirDate)
    ReDim tgSpotCurrSEE(0 To 0) As SEE
    slDateTime = gNow()
    slNowDate = Format(slDateTime, "ddddd")
    slNowTime = Format(slDateTime, "ttttt")
    llNowDate = gDateValue(slNowDate)
    llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
    If llAirDate < llNowDate Then
        Print #hlMsg, "Commercial Merge Spots Prior to " & gLongToTime(llNowTime) & " on " & slAirDate & " not checked"
        gMerge = False
        Exit Function
    End If
    ilSpotETECode = 0
    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        If tgCurrETE(ilETE).sCategory = "S" Then
            ilSpotETECode = tgCurrETE(ilETE).iCode
            Exit For
        End If
    Next ilETE
    'Remove Spots
    llLoop = LBound(tlCurrSEE)
    Do While llLoop < UBound(tlCurrSEE)
        ilFound = False
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tlCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
                '3/15/13: Not removing all spots from tlCurrSEE
                '         Moved found after Spot type test
                'ilFound = True
                If tgCurrETE(ilETE).sCategory = "S" Then
                    ilFound = True
                    ilRemove = True
                    If llAirDate = llNowDate Then
                        If llNowTime > tlCurrSEE(llLoop).lTime Then
                            ilRemove = False
                        End If
                    End If
                    If ilRemove Then
                        LSet tgSpotCurrSEE(UBound(tgSpotCurrSEE)) = tlCurrSEE(llLoop)
                        ReDim Preserve tgSpotCurrSEE(0 To UBound(tgSpotCurrSEE) + 1) As SEE
                        For llRow = llLoop + 1 To UBound(tlCurrSEE) - 1 Step 1
                            LSet tlCurrSEE(llRow - 1) = tlCurrSEE(llRow)
                            If ilFrom = 1 Then
                                slT1Comment(llRow - 1) = slT1Comment(llRow)
                                slT2Comment(llRow - 1) = slT2Comment(llRow)
                            End If
                        Next llRow
                        ReDim Preserve tlCurrSEE(0 To UBound(tlCurrSEE) - 1) As SEE
                        If ilFrom = 1 Then
                            ReDim Preserve slT1Comment(0 To UBound(slT1Comment) - 1) As String
                            ReDim Preserve slT2Comment(0 To UBound(slT2Comment) - 1) As String
                        End If
                    '3/15/13: Not removing all spots from tlCurrSEE
                    'Else
                    '    llLoop = llLoop + 1
                    End If
                '3/15/13: Not removing all spots from tlCurrSEE
                'Else
                '    llLoop = llLoop + 1
                End If
                '3/15/13: Not removing all spots from tlCurrSEE
                Exit For
            End If
        Next ilETE
        If Not ilFound Then
            llLoop = llLoop + 1
        End If
    Loop
    lbcCommercialSort.Clear
    llCounter = 0
    ilEof = False
    Do
        'Get Lines
        ilRet = 0
        On Error GoTo gMergeErr:
        Line Input #hlMerge, slLine
        On Error GoTo 0
        If ilRet <> 0 Then
            Exit Do
        End If
        If Trim$(slLine) <> "" Then
            If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                Exit Do
            End If
        End If
        DoEvents
        If Trim$(slLine) <> "" Then
            slTime = Mid$(slLine, 11, 2) & ":" & Mid$(slLine, 13, 2) & ":" & Mid$(slLine, 15, 2)
            llTime = 10 * gLengthToLong(slTime)
            slStr = Trim$(Str$(llTime))
            Do While Len(slStr) < 8
                slStr = "0" & slStr
            Loop
            llCounter = llCounter + 1
            slCounter = Trim$(Str$(llCounter))
            Do While Len(slCounter) < 6
                slCounter = "0" & slCounter
            Loop
            lbcCommercialSort.AddItem slStr & "|" & slCounter & "|" & slLine
        End If
    Loop Until ilEof
    llLastPAIndex = UBound(tlCurrSEE)
    ReDim tmMergeInfo(0 To UBound(tlCurrSEE)) As MERGEINFO
    For llLoop = 0 To UBound(tlCurrSEE) - 1 Step 1
        tmMergeInfo(llLoop).lAvailRunTime = tlCurrSEE(llLoop).lTime
        tmMergeInfo(llLoop).iSpotSoldTime = 0
        tmMergeInfo(llLoop).lLastSpotAddedIndex = -1
    Next llLoop
    ilOverbooked = False
    ilUnderbooked = False
    ilAvailNotFound = False
    For llCounter = 0 To lbcCommercialSort.ListCount - 1 Step 1
        slStr = lbcCommercialSort.List(llCounter)
        slLine = Mid$(slStr, 17)
            
        slDate = Mid$(slLine, 3, 2) & "/" & Mid$(slLine, 5, 2) & "/" & Mid$(slLine, 1, 2)
        If gDateValue(slDate) <> llAirDate Then
            Erase tmMergeInfo
            'gMerge = False
            Print #hlMsg, "Commercial Merge Spot Date " & slDate & " does not Match Schedule Date " & slAirDate
            Exit Function
        End If
        slTime = Mid$(slLine, 11, 2) & ":" & Mid$(slLine, 13, 2) & ":" & Mid$(slLine, 15, 2)
        llTime = 10 * gLengthToLong(slTime)
        slBus = Trim$(Mid$(slLine, 18, 5))
        slCopy = Mid$(slLine, 24, 5)
        slISCI = Mid$(slLine, 53, 20)
        slTitle = Trim$(Mid$(slLine, 30, 15))
        slLen = "00:" & Mid$(slLine, 46, 2) & ":" & Mid$(slLine, 48, 2)
        ilFound = False
        ilFindMatch = True
        If llAirDate = llNowDate Then
            If llNowTime > llTime Then
                ilFindMatch = False
            End If
        End If
        If ilFindMatch Then
            ilBDE = gBinarySearchName(slBus, tgCurrBDE_Name())
            If ilBDE <> -1 Then
                ilBdeCode = tgCurrBDE_Name(ilBDE).iCode
            Else
                ilBdeCode = -1
            End If
            'Pass 0=> look only at avails without any spots, pass 1 => look at avail with spots only
            'This is done to handle avails that are back to back (5:50 and 5:51 with 5:50 having a 60 sec spot)
            For ilPass = 0 To 1 Step 1
                For llLoop = 0 To llLastPAIndex - 1 Step 1
                    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                        If tlCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
                            If tgCurrETE(ilETE).sCategory = "A" Then
                                'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                                '    If tlCurrSEE(llLoop).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                                '        ilBus = StrComp(Trim$(tgCurrBDE(ilBDE).sName), slBus, vbTextCompare)
                                '        If ilBus = 0 Then
                                        If ilBdeCode = tlCurrSEE(llLoop).iBdeCode Then
                                            ilBus = 0
                                        Else
                                            ilBus = -1
                                        End If
                                        If ilBus = 0 Then
                                            'If (tlCurrSEE(llLoop).lTime = llTime) Then  'Or ((tlCurrSEE(llLoop).lTime > llTime) And (llPrevAvailLoop <> -1)) Then
                                            If ((ilPass = 0) And (tmMergeInfo(llLoop).lAvailRunTime = llTime) And (tmMergeInfo(llLoop).lLastSpotAddedIndex = -1)) Or ((ilPass = 1) And (tmMergeInfo(llLoop).lAvailRunTime = llTime) And (tmMergeInfo(llLoop).lLastSpotAddedIndex <> -1)) Then
                                                ilFound = True
                                                'Create event
                                                llUpper = UBound(tlCurrSEE)
                                                gInitSEE tlCurrSEE(llUpper)
                                                If ilFrom = 1 Then
                                                    slT1Comment(llUpper) = ""
                                                    slT2Comment(llUpper) = ""
                                                End If
                                                LSet tlCurrSEE(llUpper) = tlCurrSEE(llLoop)
                                                tlCurrSEE(llUpper).lCode = 0
                                                tlCurrSEE(llUpper).iEteCode = ilSpotETECode
                                                tlCurrSEE(llUpper).lDuration = 10 * gLengthToLong(slLen)
                                                If tlCurrSEE(llUpper).iAudioAseCode > 0 Then
                                                    tlCurrSEE(llUpper).sAudioItemID = slCopy
                                                    tlCurrSEE(llUpper).sAudioISCI = slISCI
                                                End If
                                                If tlCurrSEE(llUpper).iProtAneCode > 0 Then
                                                    tlCurrSEE(llUpper).sProtItemID = slCopy
                                                    tlCurrSEE(llUpper).sProtISCI = slISCI
                                                End If
                                                tlCurrSEE(llUpper).lSpotTime = llTime
                                                tlCurrSEE(llUpper).sInsertFlag = "Y"
                                                tmARE.lCode = 0
                                                tmARE.sName = slTitle
                                                tmARE.sUnusued = ""
                                                If tlCurrSEE(llUpper).lDuration + tmMergeInfo(llLoop).iSpotSoldTime <= tlCurrSEE(llLoop).lDuration Then
                                                    If tmMergeInfo(llLoop).lLastSpotAddedIndex <> -1 Then
                                                        'Remove start netcue and previous, end netcue
                                                        tlCurrSEE(llUpper).iStartNneCode = 0
                                                        tlCurrSEE(tmMergeInfo(llLoop).lLastSpotAddedIndex).iEndNneCode = 0
                                                    End If
                                                    tmMergeInfo(llLoop).lAvailRunTime = tmMergeInfo(llLoop).lAvailRunTime + tlCurrSEE(llUpper).lDuration
                                                    tmMergeInfo(llLoop).iSpotSoldTime = tmMergeInfo(llLoop).iSpotSoldTime + tlCurrSEE(llUpper).lDuration
                                                    ilRet = gPutInsert_ARE_AdvertiserRefer(tmARE, "EngrSchd-Merge Insert Advertiser Name")
                                                    If ilRet Then
                                                        tmMergeInfo(llLoop).lLastSpotAddedIndex = llUpper
                                                        tlCurrSEE(llUpper).lAreCode = tmARE.lCode
                                                        mSpotMatch tlCurrSEE(llUpper)
                                                        ReDim Preserve tlCurrSEE(0 To llUpper + 1) As SEE
                                                        If ilFrom = 1 Then
                                                            ReDim Preserve slT1Comment(0 To llUpper + 1) As String
                                                            ReDim Preserve slT2Comment(0 To llUpper + 1) As String
                                                        End If
                                                    Else
                                                        'gMerge = False
                                                        Print #hlMsg, "Unable to Add Advertiser/Product " & slDate & " " & slTime & " " & slTitle
                                                        gInitSEE tlCurrSEE(llUpper)
                                                    End If
                                                Else
                                                    'gMerge = False
                                                    ilOverbooked = True
                                                    Print #hlMsg, "Commercial Merge Spot Overbooked Avail " & slDate & " " & slTime & " Bus " & slBus & " Advertiser " & slTitle
                                                End If
                                '                Exit For
                                            End If
                                        End If
                                '    End If
                                'Next ilBDE
                                If ilFound Then
                                    Exit For
                                End If
                            End If
                        End If
                    Next ilETE
                    If ilFound Then
                        Exit For
                    End If
                Next llLoop
                If ilFound Then
                    Exit For
                End If
            Next ilPass
            If Not ilFound Then
                'gMerge = False
                ilAvailNotFound = True
                Print #hlMsg, "Commercial Merge Spot Avail Not Found " & slDate & " " & slTime & " Bus " & slBus & " Advertiser " & slTitle
            End If
        End If
    Next llCounter
    For llLoop = 0 To UBound(tgSpotCurrSEE) - 1 Step 1
        tgSpotCurrSEE(llLoop).lCode = Abs(tgSpotCurrSEE(llLoop).lCode)
    Next llLoop
    For llLoop = 0 To UBound(tlCurrSEE) - 1 Step 1
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tlCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
                If tgCurrETE(ilETE).sCategory = "A" Then
                    llAvailLength = tlCurrSEE(llLoop).lDuration
                    For llTest = 0 To UBound(tlCurrSEE) - 1 Step 1
                        If tlCurrSEE(llTest).sAction <> "D" Then
                            If (tlCurrSEE(llLoop).iBdeCode = tlCurrSEE(llTest).iBdeCode) And (tlCurrSEE(llLoop).lTime = tlCurrSEE(llTest).lTime) And (tlCurrSEE(llTest).iEteCode = ilSpotETECode) Then
                                llAvailLength = llAvailLength - tlCurrSEE(llTest).lDuration
                            End If
                        End If
                    Next llTest
                    If llAvailLength > 0 Then
                        ilBDE = gBinarySearchBDE(tlCurrSEE(llLoop).iBdeCode, tgCurrBDE())
                        If ilBDE <> -1 Then
                            slTime = gLongToStrTimeInTenth(tlCurrSEE(llLoop).lTime)
                            Print #hlMsg, "Commercial Merge Spot Underbooked Avail " & slDate & " " & slTime & " Bus " & Trim$(tgCurrBDE(ilBDE).sName)
                        Else
                            slTime = gLongToStrTimeInTenth(tlCurrSEE(llLoop).lTime)
                            Print #hlMsg, "Commercial Merge Spot Underbooked Avail " & slDate & " " & slTime
                        End If
                        ilUnderbooked = True
                    End If
                End If
            End If
        Next ilETE
    Next llLoop
    If ilOverbooked Or ilAvailNotFound Or ilUnderbooked Then
        ilMergeError = True
    End If
    Erase tmMergeInfo
    Exit Function
gMergeErr:
    ilRet = Err.Number
    Resume Next
End Function

Public Sub gInitSEE(tlSEE As SEE)
    tlSEE.lCode = 0
    tlSEE.lSheCode = 0
    tlSEE.sAction = "N"
    tlSEE.lDeeCode = 0
    tlSEE.iBdeCode = 0
    tlSEE.iBusCceCode = 0
    tlSEE.sSchdType = "I"
    tlSEE.iEteCode = 0
    tlSEE.lTime = 0
    tlSEE.iStartTteCode = 0
    tlSEE.sFixedTime = ""
    tlSEE.iEndTteCode = 0
    tlSEE.lDuration = 0
    tlSEE.iMteCode = 0
    tlSEE.iAudioAseCode = 0
    tlSEE.sAudioItemID = ""
    tlSEE.sAudioISCI = ""
    tlSEE.sAudioItemIDChk = "N"
    tlSEE.iAudioCceCode = 0
    tlSEE.iBkupAneCode = 0
    tlSEE.iBkupCceCode = 0
    tlSEE.iProtAneCode = 0
    tlSEE.sProtItemID = ""
    tlSEE.sProtISCI = ""
    tlSEE.sProtItemIDChk = "N"
    tlSEE.iProtCceCode = 0
    tlSEE.i1RneCode = 0
    tlSEE.i2RneCode = 0
    tlSEE.iFneCode = 0
    tlSEE.lSilenceTime = 0
    tlSEE.i1SceCode = 0
    tlSEE.i2SceCode = 0
    tlSEE.i3SceCode = 0
    tlSEE.i4SceCode = 0
    tlSEE.iStartNneCode = 0
    tlSEE.iEndNneCode = 0
    tlSEE.l1CteCode = 0
    tlSEE.l2CteCode = 0
    tlSEE.sABCFormat = ""
    tlSEE.sABCPgmCode = ""
    tlSEE.sABCXDSMode = ""
    tlSEE.sABCRecordItem = ""
    tlSEE.lAreCode = 0
    tlSEE.lSpotTime = -1
    tlSEE.lEventID = 0
    tlSEE.sAsAirStatus = "N"
    tlSEE.sIgnoreConflicts = "N"
    tlSEE.lDheCode = 0
    tlSEE.lOrigDHECode = 0
    tlSEE.sInsertFlag = "N"
    tlSEE.sUnused = ""
End Sub

Private Sub mSpotMatch(tlSEE As SEE)
    Dim llLoop As Long
    
    tlSEE.lCode = 0
    For llLoop = 0 To UBound(tgSpotCurrSEE) - 1 Step 1
        If tgSpotCurrSEE(llLoop).lCode > 0 Then
            If tlSEE.iBdeCode = tgSpotCurrSEE(llLoop).iBdeCode Then
                If tlSEE.lTime = tgSpotCurrSEE(llLoop).lTime Then
                    If tlSEE.lDuration = tgSpotCurrSEE(llLoop).lDuration Then
                        If tlSEE.lAreCode = tgSpotCurrSEE(llLoop).lAreCode Then
                            tlSEE.lCode = tgSpotCurrSEE(llLoop).lCode
                            tlSEE.sAudioItemIDChk = tgSpotCurrSEE(llLoop).sAudioItemIDChk
                            tlSEE.sProtItemIDChk = tgSpotCurrSEE(llLoop).sProtItemIDChk
                            tgSpotCurrSEE(llLoop).lCode = -tgSpotCurrSEE(llLoop).lCode
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Next llLoop
End Sub

Public Sub gItemIDCheck(spcComm As MSComm, tlCurrSEE() As SEE)
    Dim llLoop As Long
    Dim ilETE As Integer
    Dim ilITE As Integer
    Dim tlPriITE As ITE
    Dim tlSecITE As ITE
    Dim slCart As String
    Dim slQuery As String
    Dim slPriQuery As String
    Dim slResultTitle As String
    Dim slResultLength As String
    Dim slTitle As String
    Dim ilASE As Integer
    Dim slTestItemID As String
    Dim ilATE As Integer
    Dim ilANE As Integer
    Dim ilRet As Integer
    Dim ilTestPort As Integer
    Dim ilDoConnectTest As Integer

    For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
        If tgCurrITE(ilITE).sType = "P" Then
            LSet tlPriITE = tgCurrITE(ilITE)
            Exit For
        End If
    Next ilITE
    For ilITE = LBound(tgCurrITE) To UBound(tgCurrITE) - 1 Step 1
        If tgCurrITE(ilITE).sType = "S" Then
            LSet tlSecITE = tgCurrITE(ilITE)
            Exit For
        End If
    Next ilITE
    ilTestPort = True
    ilDoConnectTest = True
    For llLoop = 0 To UBound(tlCurrSEE) - 1 Step 1
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tlCurrSEE(llLoop).iEteCode = tgCurrETE(ilETE).iCode Then
                If tgCurrETE(ilETE).sCategory = "S" Then
                    slCart = Trim$(tlCurrSEE(llLoop).sAudioItemID)
                    slTitle = ""
                    If slCart <> "" Then
                        slTestItemID = ""
                        'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                        '    If tlCurrSEE(llLoop).iAudioAseCode = tgCurrASE(ilASE).iCode Then
                            ilASE = gBinarySearchASE(tlCurrSEE(llLoop).iAudioAseCode, tgCurrASE())
                            If ilASE <> -1 Then
                                'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                                    ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                                    If ilANE <> -1 Then
                                        For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
                                            If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                                                slTestItemID = tgCurrATE(ilATE).sTestItemID
                                                Exit For
                                            End If
                                        Next ilATE
                                '        If slTestItemID <> "" Then
                                '            Exit For
                                '        End If
                                    End If
                                'Next ilANE
                        '        If slTestItemID <> "" Then
                        '            Exit For
                        '        End If
                            End If
                        'Next ilASE
                        If (slTestItemID = "Y") And (ilTestPort) Then
                            ilRet = gGetRec_ARE_AdvertiserRefer(tlCurrSEE(llLoop).lAreCode, "EngrItemIDChk-mBuildItemIDbyDate: Advertiser", tmARE)
                            If ilRet Then
                                slTitle = gFileNameFilter(Trim$(tmARE.sName))
                            End If
                        End If
                        If (slTestItemID = "Y") And (slTitle <> "") And (ilTestPort) And (Trim$(tlPriITE.sName) <> "") Then
                            'gBuildItemIDQuery slCart, tlPriITE, slQuery, slPriQuery
                            'ilRet = gTestItemID(spcComm, tlPriITE, slQuery, slPriQuery, slResult)
                            ilRet = gTestItemID(spcComm, tlPriITE, slCart, ilDoConnectTest, slResultTitle, slResultLength)
                            If ilRet Then
                                'slResult = Mid$(slResult, Len(slPriQuery) + 1)
                                If StrComp(Trim$(slTitle), slResultTitle, vbTextCompare) = 0 Then
                                    tlCurrSEE(llLoop).sAudioItemIDChk = "O"
                                Else
                                    tlCurrSEE(llLoop).sAudioItemIDChk = "F"
                                End If
                            Else
                                If StrComp(slResultTitle, "Failed", vbTextCompare) = 0 Then
                                    ilTestPort = False
                                End If
                                tlCurrSEE(llLoop).sAudioItemIDChk = "N"
                            End If
                            ilDoConnectTest = False
                        Else
                            If (slTestItemID = "Y") And (slTitle = "") And (Trim$(tlPriITE.sName) = "") Then
                                tlCurrSEE(llLoop).sAudioItemIDChk = "N"
                            End If
                        End If
                    End If
                    slCart = Trim$(tlCurrSEE(llLoop).sProtItemID)
                    If slCart <> "" Then
                        slTestItemID = ""
                        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                        '    If tlCurrSEE(llLoop).iProtAneCode = tgCurrANE(ilANE).iCode Then
                            ilANE = gBinarySearchANE(tlCurrSEE(llLoop).iProtAneCode, tgCurrANE())
                            If ilANE <> -1 Then
                                For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
                                    If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                                        slTestItemID = tgCurrATE(ilATE).sTestItemID
                                        Exit For
                                    End If
                                Next ilATE
                        '        If slTestItemID <> "" Then
                        '            Exit For
                        '        End If
                            End If
                        'Next ilANE
                        If (slTestItemID = "Y") And (slTitle = "") And (ilTestPort) Then
                            ilRet = gGetRec_ARE_AdvertiserRefer(tlCurrSEE(llLoop).lAreCode, "EngrItemIDChk-mBuildItemIDbyDate: Advertiser", tmARE)
                            If ilRet Then
                                slTitle = gFileNameFilter(Trim$(tmARE.sName))
                            End If
                        End If
                        If (slTestItemID = "Y") And (slTitle <> "") And (ilTestPort) And (Trim$(tlSecITE.sName) <> "") Then
                            'gBuildItemIDQuery slCart, tlSecITE, slQuery, slPriQuery
                            'ilRet = gTestItemID(spcComm, tlSecITE, slQuery, slPriQuery, slResult)
                            ilRet = gTestItemID(spcComm, tlSecITE, slCart, ilDoConnectTest, slResultTitle, slResultLength)
                            If ilRet Then
                                'slResult = Mid$(slResult, Len(slPriQuery) + 1)
                                If StrComp(Trim$(slTitle), slResultTitle, vbTextCompare) = 0 Then
                                    tlCurrSEE(llLoop).sProtItemIDChk = "O"
                                Else
                                    tlCurrSEE(llLoop).sProtItemIDChk = "F"
                                End If
                            Else
                                If StrComp(slResultTitle, "Failed", vbTextCompare) = 0 Then
                                    ilTestPort = False
                                End If
                                tlCurrSEE(llLoop).sProtItemIDChk = "N"
                            End If
                            ilDoConnectTest = False
                        Else
                            If (slTestItemID = "Y") And (slTitle = "") And (Trim$(tlSecITE.sName) = "") Then
                                tlCurrSEE(llLoop).sProtItemIDChk = "N"
                            End If
                        End If
                    End If
                End If
            End If
        Next ilETE
    Next llLoop
End Sub

Private Sub mTransferTSEDEEToSEE(slAirDate As String, tlTSE As TSE, tlInDEE() As DEE, llOrigDHECode As Long, tlSEE() As SEE)
    Dim ilDEE As Integer
    Dim ilLoop As Integer
    Dim ilHour As Integer
    Dim slHours As String
    Dim ilHours As Integer
    Dim ilDay As Integer
    Dim llUpper As Long
    Dim ilRet As Integer
    Dim tlDEE As DEE
    Dim slHours48 As String * 48
    'Displace hours

    For ilDEE = 0 To UBound(tlInDEE) - 1 Step 1
        LSet tlDEE = tlInDEE(ilDEE)
        tlDEE.sIgnoreConflicts = "N"
        ilDay = Weekday(slAirDate, vbMonday)
        ilHour = Hour(tlTSE.sStartTime)
        If ilHour <> 0 Then
            slHours = tlDEE.sHours
            slHours48 = String(48, "N")
            For ilLoop = 0 To 23 Step 1
                Mid$(slHours48, ilHour + 1, 1) = Mid$(slHours, ilLoop + 1, 1)
                ilHour = ilHour + 1
                If gDateValue(tlTSE.sLogDate) < gDateValue(slAirDate) Then
                    If ilHour > 47 Then
                        Exit For
                    End If
                Else
                    If ilHour > 23 Then
                        Exit For
                    End If
                End If
            Next ilLoop
        End If
        tlDEE.sDays = String(7, "N")
        If gDateValue(tlTSE.sLogDate) < gDateValue(slAirDate) Then
            Select Case Weekday(DateAdd("d", 1, tlTSE.sLogDate))
                Case vbMonday
                    Mid(tlDEE.sDays, 1, 1) = "Y"
                Case vbTuesday
                    Mid(tlDEE.sDays, 2, 1) = "Y"
                Case vbWednesday
                    Mid(tlDEE.sDays, 3, 1) = "Y"
                Case vbThursday
                    Mid(tlDEE.sDays, 4, 1) = "Y"
                Case vbFriday
                    Mid(tlDEE.sDays, 5, 1) = "Y"
                Case vbSaturday
                    Mid(tlDEE.sDays, 6, 1) = "Y"
                Case vbSunday
                    Mid(tlDEE.sDays, 7, 1) = "Y"
            End Select
        End If
        Select Case Weekday(tlTSE.sLogDate)
            Case vbMonday
                Mid(tlDEE.sDays, 1, 1) = "Y"
            Case vbTuesday
                Mid(tlDEE.sDays, 2, 1) = "Y"
            Case vbWednesday
                Mid(tlDEE.sDays, 3, 1) = "Y"
            Case vbThursday
                Mid(tlDEE.sDays, 4, 1) = "Y"
            Case vbFriday
                Mid(tlDEE.sDays, 5, 1) = "Y"
            Case vbSaturday
                Mid(tlDEE.sDays, 6, 1) = "Y"
            Case vbSunday
                Mid(tlDEE.sDays, 7, 1) = "Y"
        End Select
        If Mid(tlDEE.sDays, ilDay, 1) = "Y" Then
            For ilHours = 1 To 48 Step 1
                If Mid$(slHours48, ilHours, 1) = "Y" Then
                    llUpper = UBound(tlSEE)
                    ilRet = mCopyDEEToSEE(tlDEE, tlTSE.iBdeCode, llOrigDHECode, ilHours, tlSEE(llUpper))
                    If ilRet Then
                        'tlSEE().lTime is Offset + hour, need to add minute
                        'Hour based on hour from tlTSE.sStartTime
                        If gDateValue(tlTSE.sLogDate) < gDateValue(slAirDate) Then
                            'Ignore events that don't start in the Air Date
                            If tlSEE(llUpper).lTime + 10 * (gTimeToLong(tlTSE.sStartTime, False) Mod 3600) >= 864000 Then
                                tlSEE(llUpper).lTime = (tlSEE(llUpper).lTime + 10 * (gTimeToLong(tlTSE.sStartTime, False) Mod 3600)) - 864000
                                llUpper = llUpper + 1
                                ReDim Preserve tlSEE(0 To llUpper) As SEE
                            End If
                        Else
                            'Ignore events that start in next date
                            If tlSEE(llUpper).lTime + 10 * (gTimeToLong(tlTSE.sStartTime, False) Mod 3600) < 864000 Then
                                tlSEE(llUpper).lTime = tlSEE(llUpper).lTime + 10 * (gTimeToLong(tlTSE.sStartTime, False) Mod 3600)
                                llUpper = llUpper + 1
                                ReDim Preserve tlSEE(0 To llUpper) As SEE
                            End If
                        End If
                    End If
                End If
            Next ilHours
        End If
    Next ilDEE
End Sub

Private Function mCopyDEEToSEE(tlDEE As DEE, ilBdeCode As Integer, llOrigDHECode As Long, ilHours As Integer, tlSEE As SEE) As Integer
    On Error GoTo mCopyDEEToSEEErr:
    tlSEE.lCode = 0
    tlSEE.lSheCode = 0
    tlSEE.sAction = "N"
    tlSEE.lDeeCode = tlDEE.lCode
    tlSEE.iBdeCode = ilBdeCode
    tlSEE.iBusCceCode = tlDEE.iCceCode
    tlSEE.sSchdType = "L"
    tlSEE.iEteCode = tlDEE.iEteCode
    tlSEE.lTime = tlDEE.lTime + (ilHours - 1) * CLng(36000)
    tlSEE.iStartTteCode = tlDEE.iStartTteCode
    tlSEE.sFixedTime = tlDEE.sFixedTime
    tlSEE.iEndTteCode = tlDEE.iEndTteCode
    tlSEE.lDuration = tlDEE.lDuration
    tlSEE.iMteCode = tlDEE.iMteCode
    tlSEE.iAudioAseCode = tlDEE.iAudioAseCode
    tlSEE.sAudioItemID = tlDEE.sAudioItemID
    tlSEE.sAudioISCI = tlDEE.sAudioISCI
    tlSEE.sAudioItemIDChk = "N"
    tlSEE.iAudioCceCode = tlDEE.iAudioCceCode
    tlSEE.iBkupAneCode = tlDEE.iBkupAneCode
    tlSEE.iBkupCceCode = tlDEE.iBkupCceCode
    tlSEE.iProtAneCode = tlDEE.iProtAneCode
    tlSEE.sProtItemID = tlDEE.sProtItemID
    tlSEE.sProtISCI = tlDEE.sProtISCI
    tlSEE.sProtItemIDChk = "N"
    tlSEE.iProtCceCode = tlDEE.iProtCceCode
    tlSEE.i1RneCode = tlDEE.i1RneCode
    tlSEE.i2RneCode = tlDEE.i2RneCode
    tlSEE.iFneCode = tlDEE.iFneCode
    tlSEE.lSilenceTime = tlDEE.lSilenceTime
    tlSEE.i1SceCode = tlDEE.i1SceCode
    tlSEE.i2SceCode = tlDEE.i2SceCode
    tlSEE.i3SceCode = tlDEE.i3SceCode
    tlSEE.i4SceCode = tlDEE.i4SceCode
    tlSEE.iStartNneCode = tlDEE.iStartNneCode
    tlSEE.iEndNneCode = tlDEE.iEndNneCode
    tlSEE.l1CteCode = tlDEE.l1CteCode
    tlSEE.l2CteCode = tlDEE.l2CteCode
    tlSEE.sABCFormat = tlDEE.sABCFormat
    tlSEE.sABCPgmCode = tlDEE.sABCPgmCode
    tlSEE.sABCXDSMode = tlDEE.sABCXDSMode
    tlSEE.sABCRecordItem = tlDEE.sABCRecordItem
    tlSEE.lAreCode = 0
    tlSEE.lEventID = 0
    tlSEE.sAsAirStatus = "N"
    tlSEE.sIgnoreConflicts = tlDEE.sIgnoreConflicts
    tlSEE.lDheCode = tlDEE.lDheCode
    tlSEE.lOrigDHECode = llOrigDHECode
    tlSEE.sUnused = ""
    'Field not part of record
    tlSEE.lAvailLength = tlSEE.lDuration
    mCopyDEEToSEE = True
    Exit Function
mCopyDEEToSEEErr:
    mCopyDEEToSEE = False
    Exit Function
End Function

Public Function gLoadAsAirLog(slPathAndFile As String, slAsAirDate As String, hlSEE As Integer) As Integer
    Dim ilRet As Integer
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slLine As String
    Dim ilEof As Integer
    Dim slToFile As String
    
    On Error GoTo gLoadAsAirLogErr:
    ReDim smBusDeleted(0 To 0) As String
    ilRet = 0
    slToFile = slPathAndFile
    On Error GoTo gLoadAsAirLogErr:
    hmImport = FreeFile
    Open slToFile For Input Access Read As hmImport
    If ilRet = 0 Then
        'Clear out any previous import results
        ilRet = gGetRec_SHE_ScheduleHeaderByDate(slAsAirDate, "gLoadAsAirLog", tmSHE)
        'Delete by Bus
        'If ilRet Then
        '    ilRet = gPutDelete_AAE_As_Aired(tmSHE.lCode, "gLoadAsAirLog")
        'End If
        ilRet = gGetRecs_SEE_ScheduleEventsAPI(hlSEE, sgCurrSEEStamp, -1, tmSHE.lCode, "EngrSchd-Get Events", tgCurrSEE())
        If Not ilRet Then
            If igOperationMode = 1 Then
                gLogMsg "Unable to Access Schedule Event File", "EngrServiceErrors.Txt", False
            End If
            Close #hmImport
            Exit Function
        End If
       'Process file
        ilEof = False
        Do
            'Get Lines
            ilRet = 0
            On Error GoTo gLoadAsAirLogErr:
            Line Input #hmImport, slLine
            On Error GoTo 0
            If ilRet <> 0 Then
                Exit Do
            End If
            If Trim$(slLine) <> "" Then
                If (Asc(slLine) = 26) Or (ilRet <> 0) Then    'Ctrl Z
                    Exit Do
                End If
            End If
            DoEvents
            If Trim$(slLine) <> "" Then
                ilRet = mBuildAAE(slLine)
            End If
        Loop Until ilEof
    End If
    Close #hmImport
    gLoadAsAirLog = True
    Exit Function
gLoadAsAirLogErr:
    ilRet = 1
    Resume Next
End Function

Private Function mBuildAAE(slLine As String) As Integer
    Dim llCol As Long
    Dim slEventID As String
    Dim llEventID As Long
    Dim llSEE As Long
    'Dim ilETE As Integer
    'Dim ilEteCode As Integer
    Dim ilOffset As Integer
    Dim slAirTime As String
    Dim ilRet As Integer
    Dim ilBus As Integer
    Dim ilFound As Integer
    Dim slAirDate As String
    Dim slDate As String
    
    If tgADE.iScheduleData <= 0 Then
        mBuildAAE = False
        Exit Function
    End If
    ilOffset = tgADE.iScheduleData - 1
    mInitAAE
    tmAAE.lSheCode = tmSHE.lCode
    'Event ID
    llEventID = -1
    'ilEteCode = -1
    If tgStartColAFE.iEventID > 0 Then
        slEventID = Mid$(slLine, tgStartColAFE.iEventID + ilOffset, tgNoCharAFE.iEventID)
        llEventID = Val(slEventID)
        'The event auto code is returned, not the event type
        'For llSEE = 0 To UBound(tgCurrSEE) - 1 Step 1
        '    If tgCurrSEE(llSEE).lEventID = llEventID Then
        '        tmAAE.lSeeCode = tgCurrSEE(llSEE).lCode
        '        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        '            If tgCurrSEE(llSEE).iEteCode = tgCurrETE(ilETE).iCode Then
        '                ilEteCode = tgCurrSEE(llSEE).iEteCode
        '                Exit For
        '            End If
        '        Next ilETE
        '    End If
        'Next llSEE
    End If
    'If (llEventID = -1) Or (ilEteCode = -1) Then
    If (llEventID = -1) Or (Trim$(slEventID) = "") Then
        If tgStartColAFE.iEventType > 0 Then
            'Check if Default Event (these are auto added events and should be ignored)
            If Mid$(slLine, tgStartColAFE.iEventType + ilOffset, tgNoCharAFE.iEventType) = "D" Then
                mBuildAAE = True
                Exit Function
            End If
        End If
    End If
    'If (llEventID = -1) Or (ilEteCode = -1) Then
    '    mBuildAAE = False
    '    Exit Function
    'End If
    If llEventID = -1 Then
        tmAAE.lEventID = 0
    Else
        tmAAE.lEventID = llEventID
    End If
    For llCol = BUSNAMEINDEX To ABCRECORDITEMINDEX Step 1
        'If gExportCol(ilEteCode, llCol) Then
            Select Case llCol
                Case BUSNAMEINDEX
                    If (tgStartColAFE.iBus > 0) And (tgNoCharAFE.iBus > 0) Then
                        tmAAE.sBusName = Mid$(slLine, tgStartColAFE.iBus + ilOffset, tgNoCharAFE.iBus)
                        ilFound = False
                        For ilBus = 0 To UBound(smBusDeleted) - 1 Step 1
                            If StrComp(Trim$(tmAAE.sBusName), Trim$(smBusDeleted(ilBus)), vbTextCompare) = 0 Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilBus
                        If Not ilFound Then
                            ilRet = gPutDelete_AAE_As_AiredByBus(tmSHE.lCode, Trim$(tmAAE.sBusName), "gLoadAsAirLog")
                            smBusDeleted(UBound(smBusDeleted)) = Trim$(tmAAE.sBusName)
                            ReDim Preserve smBusDeleted(0 To UBound(smBusDeleted) + 1) As String
                        End If
                    End If
                Case BUSCTRLINDEX
                    If (tgStartColAFE.iBusControl > 0) And (tgNoCharAFE.iBusControl > 0) Then
                        tmAAE.sBusControl = Mid$(slLine, tgStartColAFE.iBusControl + ilOffset, tgNoCharAFE.iBusControl)
                    End If
                Case EVENTTYPEINDEX
                    If (tgStartColAFE.iEventType > 0) And (tgNoCharAFE.iEventType > 0) Then
                        tmAAE.sEventType = Mid$(slLine, tgStartColAFE.iEventType + ilOffset, tgNoCharAFE.iEventType)
                    End If
                Case EVENTIDINDEX
                    'Event ID exported if any other column is exported and tgStartColAFE.iEventID > 0
                Case TIMEINDEX
                    If (tgStartColAFE.iTime > 0) And (tgNoCharAFE.iTime > 0) Then
                        tmAAE.sStartTime = Mid$(slLine, tgStartColAFE.iTime + ilOffset, tgNoCharAFE.iTime)
                    End If
                Case STARTTYPEINDEX
                    If (tgStartColAFE.iStartType > 0) And (tgNoCharAFE.iStartType > 0) Then
                        tmAAE.sStartType = Mid$(slLine, tgStartColAFE.iStartType + ilOffset, tgNoCharAFE.iStartType)
                    End If
                Case FIXEDINDEX
                    If (tgStartColAFE.iFixedTime > 0) And (tgNoCharAFE.iFixedTime > 0) Then
                        tmAAE.sFixedTime = Mid$(slLine, tgStartColAFE.iFixedTime + ilOffset, tgNoCharAFE.iFixedTime)
                    End If
                Case ENDTYPEINDEX
                    If (tgStartColAFE.iEndType > 0) And (tgNoCharAFE.iEndType > 0) Then
                        tmAAE.sEndType = Mid$(slLine, tgStartColAFE.iEndType + ilOffset, tgNoCharAFE.iEndType)
                    End If
                Case DURATIONINDEX
                    If (tgStartColAFE.iDuration > 0) And (tgNoCharAFE.iDuration > 0) Then
                        tmAAE.sDuration = Mid$(slLine, tgStartColAFE.iDuration + ilOffset, tgNoCharAFE.iDuration)
                    End If
                Case MATERIALINDEX
                    If (tgStartColAFE.iMaterialType > 0) And (tgNoCharAFE.iMaterialType > 0) Then
                        tmAAE.sMaterialType = Mid$(slLine, tgStartColAFE.iMaterialType + ilOffset, tgNoCharAFE.iMaterialType)
                    End If
                Case AUDIONAMEINDEX
                    If (tgStartColAFE.iAudioName > 0) And (tgNoCharAFE.iAudioName > 0) Then
                        tmAAE.sAudioName = Mid$(slLine, tgStartColAFE.iAudioName + ilOffset, tgNoCharAFE.iAudioName)
                    End If
                Case AUDIOITEMIDINDEX
                    If (tgStartColAFE.iAudioItemID > 0) And (tgNoCharAFE.iAudioItemID > 0) Then
                        tmAAE.sAudioItemID = Mid$(slLine, tgStartColAFE.iAudioItemID + ilOffset, tgNoCharAFE.iAudioItemID)
                    End If
                Case AUDIOISCIINDEX
                    If (tgStartColAFE.iAudioISCI > 0) And (tgNoCharAFE.iAudioISCI > 0) Then
                        tmAAE.sAudioISCI = Mid$(slLine, tgStartColAFE.iAudioISCI + ilOffset, tgNoCharAFE.iAudioISCI)
                    End If
                Case AUDIOCTRLINDEX
                    If (tgStartColAFE.iAudioControl > 0) And (tgNoCharAFE.iAudioControl > 0) Then
                        tmAAE.sAudioCrtlChar = Mid$(slLine, tgStartColAFE.iAudioControl + ilOffset, tgNoCharAFE.iAudioControl)
                    End If
                Case BACKUPNAMEINDEX
                    If (tgStartColAFE.iBkupAudioName > 0) And (tgNoCharAFE.iBkupAudioName > 0) Then
                        tmAAE.sBkupAudioName = Mid$(slLine, tgStartColAFE.iBkupAudioName + ilOffset, tgNoCharAFE.iBkupAudioName)
                    End If
                Case BACKUPCTRLINDEX
                    If (tgStartColAFE.iBkupAudioControl > 0) And (tgNoCharAFE.iBkupAudioControl > 0) Then
                        tmAAE.sBkupCtrlChar = Mid$(slLine, tgStartColAFE.iBkupAudioControl + ilOffset, tgNoCharAFE.iBkupAudioControl)
                    End If
                Case PROTNAMEINDEX
                    If (tgStartColAFE.iProtAudioName > 0) And (tgNoCharAFE.iProtAudioName > 0) Then
                        tmAAE.sProtAudioName = Mid$(slLine, tgStartColAFE.iProtAudioName + ilOffset, tgNoCharAFE.iProtAudioName)
                    End If
                Case PROTITEMIDINDEX
                    If (tgStartColAFE.iProtItemID > 0) And (tgNoCharAFE.iProtItemID > 0) Then
                        tmAAE.sProtItemID = Mid$(slLine, tgStartColAFE.iProtItemID + ilOffset, tgNoCharAFE.iProtItemID)
                    End If
                Case PROTISCIINDEX
                    If (tgStartColAFE.iProtISCI > 0) And (tgNoCharAFE.iProtISCI > 0) Then
                        tmAAE.sProtISCI = Mid$(slLine, tgStartColAFE.iProtISCI + ilOffset, tgNoCharAFE.iProtISCI)
                    End If
                Case PROTCTRLINDEX
                    If (tgStartColAFE.iProtAudioControl > 0) And (tgNoCharAFE.iProtAudioControl > 0) Then
                        tmAAE.sProtCtrlChar = Mid$(slLine, tgStartColAFE.iProtAudioControl + ilOffset, tgNoCharAFE.iProtAudioControl)
                    End If
                Case RELAY1INDEX
                    If (tgStartColAFE.iRelay1 > 0) And (tgNoCharAFE.iRelay1 > 0) Then
                        tmAAE.sRelay1 = Mid$(slLine, tgStartColAFE.iRelay1 + ilOffset, tgNoCharAFE.iRelay1)
                    End If
                Case RELAY2INDEX
                    If (tgStartColAFE.iRelay2 > 0) And (tgNoCharAFE.iRelay2 > 0) Then
                        tmAAE.sRelay2 = Mid$(slLine, tgStartColAFE.iRelay2 + ilOffset, tgNoCharAFE.iRelay2)
                    End If
                Case FOLLOWINDEX
                    If (tgStartColAFE.iFollow > 0) And (tgNoCharAFE.iFollow > 0) Then
                        tmAAE.sFollow = Mid$(slLine, tgStartColAFE.iFollow + ilOffset, tgNoCharAFE.iFollow)
                    End If
                Case SILENCETIMEINDEX
                    If (tgStartColAFE.iSilenceTime > 0) And (tgNoCharAFE.iSilenceTime > 0) Then
                        tmAAE.sSilenceTime = Mid$(slLine, tgStartColAFE.iSilenceTime + ilOffset, tgNoCharAFE.iSilenceTime)
                    End If
                Case SILENCE1INDEX
                    If (tgStartColAFE.iSilence1 > 0) And (tgNoCharAFE.iSilence1 > 0) Then
                        tmAAE.sSilence1 = Mid$(slLine, tgStartColAFE.iSilence1 + ilOffset, tgNoCharAFE.iSilence1)
                    End If
                Case SILENCE2INDEX
                    If (tgStartColAFE.iSilence2 > 0) And (tgNoCharAFE.iSilence2 > 0) Then
                        tmAAE.sSilence2 = Mid$(slLine, tgStartColAFE.iSilence2 + ilOffset, tgNoCharAFE.iSilence2)
                    End If
                Case SILENCE3INDEX
                    If (tgStartColAFE.iSilence3 > 0) And (tgNoCharAFE.iSilence3 > 0) Then
                        tmAAE.sSilence3 = Mid$(slLine, tgStartColAFE.iSilence3 + ilOffset, tgNoCharAFE.iSilence3)
                    End If
                Case SILENCE4INDEX
                    If (tgStartColAFE.iSilence4 > 0) And (tgNoCharAFE.iSilence4 > 0) Then
                        tmAAE.sSilence4 = Mid$(slLine, tgStartColAFE.iSilence4 + ilOffset, tgNoCharAFE.iSilence4)
                    End If
                Case NETCUE1INDEX
                    If (tgStartColAFE.iStartNetcue > 0) And (tgNoCharAFE.iStartNetcue > 0) Then
                        tmAAE.sNetcueStart = Mid$(slLine, tgStartColAFE.iStartNetcue + ilOffset, tgNoCharAFE.iStartNetcue)
                    End If
                Case NETCUE2INDEX
                    If (tgStartColAFE.iStopNetcue > 0) And (tgNoCharAFE.iStopNetcue > 0) Then
                        tmAAE.sNetcueEnd = Mid$(slLine, tgStartColAFE.iStopNetcue + ilOffset, tgNoCharAFE.iStopNetcue)
                    End If
                Case TITLE1INDEX
                    If (tgStartColAFE.iTitle1 > 0) And (tgNoCharAFE.iTitle2 > 0) Then
                        tmAAE.sTitle1 = Mid$(slLine, tgStartColAFE.iTitle1 + ilOffset, tgNoCharAFE.iTitle1)
                    End If
                Case TITLE2INDEX
                    If (tgStartColAFE.iTitle2 > 0) And (tgNoCharAFE.iTitle1 > 0) Then
                        tmAAE.sTitle2 = Mid$(slLine, tgStartColAFE.iTitle2 + ilOffset, tgNoCharAFE.iTitle2)
                    End If
                Case ABCFORMATINDEX
                    If (tgStartColAFE.iABCFormat > 0) And (tgNoCharAFE.iABCFormat > 0) Then
                        tmAAE.sABCFormat = Mid$(slLine, tgStartColAFE.iABCFormat + ilOffset, tgNoCharAFE.iABCFormat)
                    End If
                Case ABCPGMCODEINDEX
                    If (tgStartColAFE.iABCPgmCode > 0) And (tgNoCharAFE.iABCPgmCode > 0) Then
                        tmAAE.sABCPgmCode = Mid$(slLine, tgStartColAFE.iABCPgmCode + ilOffset, tgNoCharAFE.iABCPgmCode)
                    End If
                Case ABCXDSMODEINDEX
                    If (tgStartColAFE.iABCXDSMode > 0) And (tgNoCharAFE.iABCXDSMode > 0) Then
                        tmAAE.sABCXDSMode = Mid$(slLine, tgStartColAFE.iABCXDSMode + ilOffset, tgNoCharAFE.iABCXDSMode)
                    End If
                Case ABCRECORDITEMINDEX
                    If (tgStartColAFE.iABCRecordItem > 0) And (tgNoCharAFE.iABCRecordItem > 0) Then
                        tmAAE.sABCRecordItem = Mid$(slLine, tgStartColAFE.iABCRecordItem + ilOffset, tgNoCharAFE.iABCRecordItem)
                    End If
            End Select
        'End If
    Next llCol
    'Out Time (End Time)
    'If gExportCol(ilEteCode, DURATIONINDEX) Then
        If (tgStartColAFE.iEndTime > 0) And (tgNoCharAFE.iEndTime > 0) Then
            tmAAE.sOutTime = Mid$(slLine, tgStartColAFE.iEndTime + ilOffset, tgNoCharAFE.iEndTime)
        End If
    'End If
    'Extract Air Date
    slAirDate = Mid$(slLine, tgADE.iDate, tgADE.iDateNoChar)
    If tgADE.iDateNoChar = 8 Then
        slDate = Mid$(slAirDate, 5, 2) & "/" & Mid$(slAirDate, 7, 2) & "/" & Left$(slAirDate, 4)
    ElseIf tgADE.iDateNoChar = 6 Then
        slDate = Mid$(slAirDate, 3, 2) & "/" & Mid$(slAirDate, 5, 2) & "/" & Left$(slAirDate, 2)
    End If
    tmAAE.sAirDate = slDate
    'Extract Air Time
    slAirTime = Mid$(slLine, tgADE.iTime, tgADE.iTimeNoChar)
    tmAAE.lAirTime = gStrTimeInTenthToLong(slAirTime, False)
    'Flags
    tmAAE.sAutoOff = Mid$(slLine, tgADE.iAutoOff, 1)
    tmAAE.sData = Mid$(slLine, tgADE.iData, 1)
    tmAAE.sSchedule = Mid$(slLine, tgADE.iSchedule, 1)
    tmAAE.sTrueTime = Mid$(slLine, tgADE.iTrueTime, 1)
    tmAAE.sSourceConflict = Mid$(slLine, tgADE.iSourceConflict, 1)
    tmAAE.sSourceUnavail = Mid$(slLine, tgADE.iSourceUnavail, 1)
    tmAAE.sSourceItem = Mid$(slLine, tgADE.iSourceItem, 1)
    tmAAE.sBkupSrceUnavail = Mid$(slLine, tgADE.iBkupSrceUnavail, 1)
    tmAAE.sBkupSrceItem = Mid$(slLine, tgADE.iBkupSrceItem, 1)
    tmAAE.sProtSrceUnavail = Mid$(slLine, tgADE.iProtSrceUnavail, 1)
    tmAAE.sProtSrceItem = Mid$(slLine, tgADE.iProtSrceItem, 1)
    ilRet = gPutInsert_AAE_As_Aired(tmAAE, "Load As Aired-mBuildAAE: AAEE")
    mBuildAAE = ilRet
    
End Function

Private Sub mInitAAE()
    tmAAE.lCode = 0
    tmAAE.lSheCode = 0
    tmAAE.lSeeCode = 0
    tmAAE.sAirDate = ""
    tmAAE.lAirTime = 0
    tmAAE.sAutoOff = ""
    tmAAE.sData = ""
    tmAAE.sSchedule = ""
    tmAAE.sTrueTime = ""
    tmAAE.sSourceConflict = ""
    tmAAE.sSourceUnavail = ""
    tmAAE.sSourceItem = ""
    tmAAE.sBkupSrceUnavail = ""
    tmAAE.sBkupSrceItem = ""
    tmAAE.sProtSrceUnavail = ""
    tmAAE.sProtSrceItem = ""
    tmAAE.sDate = ""
    tmAAE.lEventID = 0
    tmAAE.sBusName = ""
    tmAAE.sBusControl = ""
    tmAAE.sEventType = ""
    tmAAE.sStartTime = ""
    tmAAE.sStartType = ""
    tmAAE.sFixedTime = ""
    tmAAE.sEndType = ""
    tmAAE.sDuration = ""
    tmAAE.sOutTime = ""
    tmAAE.sMaterialType = ""
    tmAAE.sAudioName = ""
    tmAAE.sAudioItemID = ""
    tmAAE.sAudioISCI = ""
    tmAAE.sAudioCrtlChar = ""
    tmAAE.sBkupAudioName = ""
    tmAAE.sBkupCtrlChar = ""
    tmAAE.sProtAudioName = ""
    tmAAE.sProtItemID = ""
    tmAAE.sProtISCI = ""
    tmAAE.sProtCtrlChar = ""
    tmAAE.sRelay1 = ""
    tmAAE.sRelay2 = ""
    tmAAE.sFollow = ""
    tmAAE.sSilenceTime = ""
    tmAAE.sSilence1 = ""
    tmAAE.sSilence2 = ""
    tmAAE.sSilence3 = ""
    tmAAE.sSilence4 = ""
    tmAAE.sNetcueStart = ""
    tmAAE.sNetcueEnd = ""
    tmAAE.sTitle1 = ""
    tmAAE.sTitle2 = ""
    tmAAE.sABCFormat = ""
    tmAAE.sABCPgmCode = ""
    tmAAE.sABCXDSMode = ""
    tmAAE.sABCRecordItem = ""
    tmAAE.sEnteredDate = Format(gNow(), "ddddd")
    tmAAE.sEnteredTime = Format(gNow(), "ttttt")
    tmAAE.sUnused = ""

End Sub

Public Function gExportCol(ilEteCode As Integer, llCol As Long) As Integer
    Dim slStr As String
    Dim ilETE As Integer
    Dim ilEPE As Integer
    
    gExportCol = True
    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        If tgCurrETE(ilETE).iCode = ilEteCode Then
            For ilEPE = 0 To UBound(tgCurrEPE) - 1 Step 1
                If tgCurrEPE(ilEPE).sType = "U" Then
                    If tgCurrETE(ilETE).iCode = tgCurrEPE(ilEPE).iEteCode Then
                        Select Case llCol
                            Case BUSNAMEINDEX
                                If tgCurrEPE(ilEPE).sBus <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case BUSCTRLINDEX
                                If tgCurrEPE(ilEPE).sBusControl <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case EVENTTYPEINDEX
                                'Event Type exported if any other column exported and tgStartColAFE.iEventType >0
                            Case EVENTIDINDEX
                                'Event ID exported if any other column is exported and tgStartColAFE.iEventID > 0
                            Case TIMEINDEX
                                If tgCurrEPE(ilEPE).sTime <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case STARTTYPEINDEX
                                If tgCurrEPE(ilEPE).sStartType <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case FIXEDINDEX
                                If tgCurrEPE(ilEPE).sFixedTime <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case ENDTYPEINDEX
                                If tgCurrEPE(ilEPE).sEndType <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case DURATIONINDEX
                                If tgCurrEPE(ilEPE).sDuration <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case MATERIALINDEX
                                If tgCurrEPE(ilEPE).sMaterialType <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case AUDIONAMEINDEX
                                If tgCurrEPE(ilEPE).sAudioName <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case AUDIOITEMIDINDEX
                                If tgCurrEPE(ilEPE).sAudioItemID <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case AUDIOISCIINDEX
                                If tgCurrEPE(ilEPE).sAudioISCI <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case AUDIOCTRLINDEX
                                If tgCurrEPE(ilEPE).sAudioControl <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case BACKUPNAMEINDEX
                                If tgCurrEPE(ilEPE).sBkupAudioName <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case BACKUPCTRLINDEX
                                If tgCurrEPE(ilEPE).sBkupAudioControl <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case PROTNAMEINDEX
                                If tgCurrEPE(ilEPE).sProtAudioName <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case PROTITEMIDINDEX
                                If tgCurrEPE(ilEPE).sProtAudioItemID <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case PROTISCIINDEX
                                If tgCurrEPE(ilEPE).sProtAudioISCI <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case PROTCTRLINDEX
                                If tgCurrEPE(ilEPE).sProtAudioControl <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case RELAY1INDEX
                                If tgCurrEPE(ilEPE).sRelay1 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case RELAY2INDEX
                                If tgCurrEPE(ilEPE).sRelay2 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case FOLLOWINDEX
                                If tgCurrEPE(ilEPE).sFollow <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case SILENCETIMEINDEX
                                If tgCurrEPE(ilEPE).sSilenceTime <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case SILENCE1INDEX
                                If tgCurrEPE(ilEPE).sSilence1 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case SILENCE2INDEX
                                If tgCurrEPE(ilEPE).sSilence2 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case SILENCE3INDEX
                                If tgCurrEPE(ilEPE).sSilence3 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case SILENCE4INDEX
                                If tgCurrEPE(ilEPE).sSilence4 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case NETCUE1INDEX
                                If tgCurrEPE(ilEPE).sStartNetcue <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case NETCUE2INDEX
                                If tgCurrEPE(ilEPE).sStopNetcue <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case TITLE1INDEX
                                If tgCurrEPE(ilEPE).sTitle1 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case TITLE2INDEX
                                If tgCurrEPE(ilEPE).sTitle2 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case ABCFORMATINDEX
                                If (tgCurrEPE(ilEPE).sABCFormat <> "Y") Or (sgClientFields <> "A") Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case ABCPGMCODEINDEX
                                If (tgCurrEPE(ilEPE).sABCPgmCode <> "Y") Or (sgClientFields <> "A") Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case ABCXDSMODEINDEX
                                If (tgCurrEPE(ilEPE).sABCXDSMode <> "Y") Or (sgClientFields <> "A") Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case ABCRECORDITEMINDEX
                                If (tgCurrEPE(ilEPE).sABCRecordItem <> "Y") Or (sgClientFields <> "A") Then
                                    gExportCol = False
                                    Exit Function
                                End If
                        End Select
                        Exit For
                    End If
                End If
            Next ilEPE
            For ilEPE = 0 To UBound(tgCurrEPE) - 1 Step 1
                If tgCurrEPE(ilEPE).sType = "E" Then
                    If tgCurrETE(ilETE).iCode = tgCurrEPE(ilEPE).iEteCode Then
                        Select Case llCol
                            Case BUSNAMEINDEX
                                If tgCurrEPE(ilEPE).sBus <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case BUSCTRLINDEX
                                If tgCurrEPE(ilEPE).sBusControl <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case EVENTTYPEINDEX
                                'Always exported if any other col is exported
                            Case EVENTIDINDEX
                                'Always exported if any other col is exported
                            Case TIMEINDEX
                                If tgCurrEPE(ilEPE).sTime <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case STARTTYPEINDEX
                                If tgCurrEPE(ilEPE).sStartType <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case FIXEDINDEX
                                If tgCurrEPE(ilEPE).sFixedTime <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case ENDTYPEINDEX
                                If tgCurrEPE(ilEPE).sEndType <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case DURATIONINDEX
                                If tgCurrEPE(ilEPE).sDuration <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case MATERIALINDEX
                                If tgCurrEPE(ilEPE).sMaterialType <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case AUDIONAMEINDEX
                                If tgCurrEPE(ilEPE).sAudioName <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case AUDIOITEMIDINDEX
                                If tgCurrEPE(ilEPE).sAudioItemID <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case AUDIOISCIINDEX
                                If tgCurrEPE(ilEPE).sAudioISCI <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case AUDIOCTRLINDEX
                                If tgCurrEPE(ilEPE).sAudioControl <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case BACKUPNAMEINDEX
                                If tgCurrEPE(ilEPE).sBkupAudioName <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case BACKUPCTRLINDEX
                                If tgCurrEPE(ilEPE).sBkupAudioControl <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case PROTNAMEINDEX
                                If tgCurrEPE(ilEPE).sProtAudioName <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case PROTITEMIDINDEX
                                If tgCurrEPE(ilEPE).sProtAudioItemID <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case PROTISCIINDEX
                                If tgCurrEPE(ilEPE).sProtAudioISCI <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case PROTCTRLINDEX
                                If tgCurrEPE(ilEPE).sProtAudioControl <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case RELAY1INDEX
                                If tgCurrEPE(ilEPE).sRelay1 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case RELAY2INDEX
                                If tgCurrEPE(ilEPE).sRelay2 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case FOLLOWINDEX
                                If tgCurrEPE(ilEPE).sFollow <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case SILENCETIMEINDEX
                                If tgCurrEPE(ilEPE).sSilenceTime <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case SILENCE1INDEX
                                If tgCurrEPE(ilEPE).sSilence1 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case SILENCE2INDEX
                                If tgCurrEPE(ilEPE).sSilence2 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case SILENCE3INDEX
                                If tgCurrEPE(ilEPE).sSilence3 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case SILENCE4INDEX
                                If tgCurrEPE(ilEPE).sSilence4 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case NETCUE1INDEX
                                If tgCurrEPE(ilEPE).sStartNetcue <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case NETCUE2INDEX
                                If tgCurrEPE(ilEPE).sStopNetcue <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case TITLE1INDEX
                                If tgCurrEPE(ilEPE).sTitle1 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case TITLE2INDEX
                                If tgCurrEPE(ilEPE).sTitle2 <> "Y" Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case ABCFORMATINDEX
                                If (tgCurrEPE(ilEPE).sABCFormat <> "Y") Or (sgClientFields <> "A") Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case ABCPGMCODEINDEX
                                If (tgCurrEPE(ilEPE).sABCPgmCode <> "Y") Or (sgClientFields <> "A") Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case ABCXDSMODEINDEX
                                If (tgCurrEPE(ilEPE).sABCXDSMode <> "Y") Or (sgClientFields <> "A") Then
                                    gExportCol = False
                                    Exit Function
                                End If
                            Case ABCRECORDITEMINDEX
                                If (tgCurrEPE(ilEPE).sABCRecordItem <> "Y") Or (sgClientFields <> "A") Then
                                    gExportCol = False
                                    Exit Function
                                End If
                        End Select
                        Exit For
                    End If
                End If
            Next ilEPE
            Exit For
        End If
    Next ilETE
End Function

Private Function mUpdateSpots(tlAdjSEE() As SEE, tlSEE() As SEE, ilSpotETECode As Integer) As Integer
    Dim llSEE As Long
    Dim llAdjSEE As Long
    Dim llSpot As Long
    Dim ilETE As Integer
    Dim ilUpdate As Integer
    Dim ilRet As Integer
    
    For llAdjSEE = 0 To UBound(tlAdjSEE) - 1 Step 1
        For llSEE = 0 To UBound(tlSEE) Step 1
            If (tlAdjSEE(llAdjSEE).lTime = tlSEE(llSEE).lTime) And (tlAdjSEE(llAdjSEE).iEteCode = tlSEE(llSEE).iEteCode) Then
                For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                    If tgCurrETE(ilETE).iCode = tlAdjSEE(llAdjSEE).iEteCode Then
                        If tgCurrETE(ilETE).sCategory = "A" Then
                            For llSpot = 0 To UBound(tlAdjSEE) - 1 Step 1
                                If (llAdjSEE <> llSpot) And (tlAdjSEE(llSpot).iEteCode = ilSpotETECode) Then
                                    If (tlAdjSEE(llAdjSEE).lTime = tlAdjSEE(llSpot).lTime) And (tlAdjSEE(llAdjSEE).iBdeCode = tlAdjSEE(llSpot).iBdeCode) And (tlAdjSEE(llSpot).iBdeCode = tlSEE(llSEE).iBdeCode) Then
                                        ilUpdate = False
                                        'If tlAdjSEE(llSpot).iBdeCode <> tlSEE(llSEE).iBdeCode Then
                                        '    ilUpdate = True
                                        '    tlAdjSEE(llSpot).iBdeCode = tlSEE(llSEE).iBdeCode
                                        'End If
                                        If tlAdjSEE(llSpot).iBusCceCode <> tlSEE(llSEE).iBusCceCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).iBusCceCode = tlSEE(llSEE).iBusCceCode
                                        End If
                                        If tlAdjSEE(llSpot).iStartTteCode <> tlSEE(llSEE).iStartTteCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).iStartTteCode = tlSEE(llSEE).iStartTteCode
                                        End If
                                        If tlAdjSEE(llSpot).sFixedTime <> tlSEE(llSEE).sFixedTime Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).sFixedTime = tlSEE(llSEE).sFixedTime
                                        End If
                                        If tlAdjSEE(llSpot).iEndTteCode <> tlSEE(llSEE).iEndTteCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).iEndTteCode = tlSEE(llSEE).iEndTteCode
                                        End If
                                        If tlAdjSEE(llSpot).iMteCode <> tlSEE(llSEE).iMteCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).iMteCode = tlSEE(llSEE).iMteCode
                                        End If
                                        'If tlAdjSEE(llSpot).lDuration <> tlSEE(llSEE).lDuration Then
                                        '    ilUpdate = True
                                        '    tlAdjSEE(llSpot).lDuration = tlSEE(llSEE).lDuration
                                        'End If
                                        If tlAdjSEE(llSpot).iAudioAseCode <> tlSEE(llSEE).iAudioAseCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).iAudioAseCode = tlSEE(llSEE).iAudioAseCode
                                        End If
                                        'If StrComp(tlAdjSEE(llSpot).sAudioItemID, tlSEE(llSEE).sAudioItemID, vbTextCompare) <> 0 Then
                                        '    ilUpdate = True
                                        '    tlAdjSEE(llSpot).sAudioItemID = tlSEE(llSEE).sAudioItemID
                                        'End If
                                        'If StrComp(tlAdjSEE(llSpot).sAudioISCI, tlSEE(llSEE).sAudioISCI, vbTextCompare) <> 0 Then
                                        '    ilUpdate = True
                                        '    tlAdjSEE(llSpot).sAudioISCI = tlSEE(llSEE).sAudioISCI
                                        'End If
                                        If tlAdjSEE(llSpot).iAudioCceCode <> tlSEE(llSEE).iAudioCceCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).iAudioCceCode = tlSEE(llSEE).iAudioCceCode
                                        End If
                                        If tlAdjSEE(llSpot).iBkupAneCode <> tlSEE(llSEE).iBkupAneCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).iBkupAneCode = tlSEE(llSEE).iBkupAneCode
                                        End If
                                        If tlAdjSEE(llSpot).iBkupCceCode <> tlSEE(llSEE).iBkupCceCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).iBkupCceCode = tlSEE(llSEE).iBkupCceCode
                                        End If
                                        If tlAdjSEE(llSpot).iProtAneCode <> tlSEE(llSEE).iProtAneCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).iProtAneCode = tlSEE(llSEE).iProtAneCode
                                        End If
                                        'If StrComp(tlAdjSEE(llSpot).sProtItemID, tlSEE(llSEE).sProtItemID, vbTextCompare) <> 0 Then
                                        '    ilUpdate = True
                                        '    tlAdjSEE(llSpot).sProtItemID = tlSEE(llSEE).sProtItemID
                                        'End If
                                        'If StrComp(tlAdjSEE(llSpot).sProtISCI, tlSEE(llSEE).sProtISCI, vbTextCompare) <> 0 Then
                                        '    ilUpdate = True
                                        '    tlAdjSEE(llSpot).sProtISCI = tlSEE(llSEE).sProtISCI
                                        'End If
                                        If tlAdjSEE(llSpot).iProtCceCode <> tlSEE(llSEE).iProtCceCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).iProtCceCode = tlSEE(llSEE).iProtCceCode
                                        End If
                                        If tlAdjSEE(llSpot).i1RneCode <> tlSEE(llSEE).i1RneCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).i1RneCode = tlSEE(llSEE).i1RneCode
                                        End If
                                        If tlAdjSEE(llSpot).i2RneCode <> tlSEE(llSEE).i2RneCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).i2RneCode = tlSEE(llSEE).i2RneCode
                                        End If
                                        If tlAdjSEE(llSpot).iFneCode <> tlSEE(llSEE).iFneCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).iFneCode = tlSEE(llSEE).iFneCode
                                        End If
                                        If tlAdjSEE(llSpot).lSilenceTime <> tlSEE(llSEE).lSilenceTime Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).lSilenceTime = tlSEE(llSEE).lSilenceTime
                                        End If
                                        If tlAdjSEE(llSpot).i1SceCode <> tlSEE(llSEE).i1SceCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).i1SceCode = tlSEE(llSEE).i1SceCode
                                        End If
                                        If tlAdjSEE(llSpot).i2SceCode <> tlSEE(llSEE).i2SceCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).i2SceCode = tlSEE(llSEE).i2SceCode
                                        End If
                                        If tlAdjSEE(llSpot).i3SceCode <> tlSEE(llSEE).i3SceCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).i3SceCode = tlSEE(llSEE).i3SceCode
                                        End If
                                        If tlAdjSEE(llSpot).i4SceCode <> tlSEE(llSEE).i4SceCode Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).i4SceCode = tlSEE(llSEE).i4SceCode
                                        End If
                                        'If tlAdjSEE(llSpot).iStartNneCode <> tlSEE(llSEE).iStartNneCode Then
                                        '    ilUpdate = True
                                        '    tlAdjSEE(llSpot).iStartNneCode = tlSEE(llSEE).iStartNneCode
                                        'End If
                                        'If tlAdjSEE(llSpot).iEndNneCode <> tlSEE(llSEE).iEndNneCode Then
                                        '    ilUpdate = True
                                        '    tlAdjSEE(llSpot).iEndNneCode = tlSEE(llSEE).iEndNneCode
                                        'End If
                                        '7/8/11: Make T2 work like T1
                                        'If tlAdjSEE(llSpot).l2CteCode <> tlSEE(llSEE).l2CteCode Then
                                        '    ilUpdate = True
                                        '    tlAdjSEE(llSpot).l2CteCode = tlSEE(llSEE).l2CteCode
                                        'End If
                                        If StrComp(tlAdjSEE(llSpot).sABCFormat, tlSEE(llSEE).sABCFormat, vbTextCompare) <> 0 Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).sABCFormat = tlSEE(llSEE).sABCFormat
                                        End If
                                        If StrComp(tlAdjSEE(llSpot).sABCPgmCode, tlSEE(llSEE).sABCPgmCode, vbTextCompare) <> 0 Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).sABCPgmCode = tlSEE(llSEE).sABCPgmCode
                                        End If
                                        If StrComp(tlAdjSEE(llSpot).sABCXDSMode, tlSEE(llSEE).sABCXDSMode, vbTextCompare) <> 0 Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).sABCXDSMode = tlSEE(llSEE).sABCXDSMode
                                        End If
                                        If StrComp(tlAdjSEE(llSpot).sABCRecordItem, tlSEE(llSEE).sABCRecordItem, vbTextCompare) <> 0 Then
                                            ilUpdate = True
                                            tlAdjSEE(llSpot).sABCRecordItem = tlSEE(llSEE).sABCRecordItem
                                        End If
                                        If ilUpdate Then
                                            tlAdjSEE(llSpot).lDeeCode = tlSEE(llSEE).lDeeCode
                                            ilRet = gPutUpdate_SEE_Schedule_Events(tlAdjSEE(llSpot), "gAdjustSEE: mUpdateSpots")
                                        End If
                                    End If
                                End If
                            Next llSpot
                        End If
                    End If
                Next ilETE
            End If
        Next llSEE
    Next llAdjSEE
    mUpdateSpots = True
    Exit Function
End Function



Private Sub mRemoveEvent(llSEE As Long, llAirDate As Long, slCategory As String, ilSpotRomoved As Integer, llBuildNewLoadDate() As Long)
    Dim ilFound As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    
    If tmAdjSEE(llSEE).sAction <> "D" Then
        If tmAdjSEE(llSEE).sSentStatus = "S" Then
            If slCategory = "A" Then
                ilSpotRomoved = 1
            End If
            ilRet = gPutUpdate_SEE_UnsentFlag(tmAdjSEE(llSEE).lCode, "D", "Schedule Definition-mSave: SEE")
            ilFound = False
            For ilLoop = 0 To UBound(llBuildNewLoadDate) - 1 Step 1
                If llAirDate = llBuildNewLoadDate(ilLoop) Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoop
            If Not ilFound Then
                llBuildNewLoadDate(UBound(llBuildNewLoadDate)) = llAirDate
                ReDim Preserve llBuildNewLoadDate(0 To UBound(llBuildNewLoadDate) + 1) As Long
            End If
        Else
            If ilSpotRomoved = 0 Then
                If slCategory = "A" Then
                    ilSpotRomoved = 2
                End If
            End If
            'Delete record as it has not been sent
            ilRet = gPutUpdate_SEE_UnsentFlag(tmAdjSEE(llSEE).lCode, "R", "Schedule Definition-mSave: SEE")
        End If
    End If

End Sub
