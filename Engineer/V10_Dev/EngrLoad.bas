Attribute VB_Name = "EngrLoad"

'
' Release: 1.0
'
' Description:
'   This file contains the Constants

Option Explicit
Public sgLoadFileName As String
Public sgTmpLoadFileName As String

Private tmSeeTimeSort() As SEETIMESORT
Private smExportStr As String
Private tmCTE As CTE
Private tmARE As ARE

'Constant must match ones defined in EngrSchdSub
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


Public Function gOpenAutoMsgFile(slAirDate As String, slMsgFileName As String, hlMsg As Integer) As Integer
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim slChar As String
    Dim slAirYear As String
    Dim slAirMonth As String
    Dim slAirDay As String
    Dim slNowDate As String
    Dim slNowYear As String
    Dim slNowMonth As String
    Dim slNowDay As String

    On Error GoTo gOpenAutoMsgFileErr:
    ilRet = 0
    slAirYear = Year(slAirDate)
    If Len(slAirYear) = 4 Then
        slAirYear = Mid$(slAirYear, 3)
    End If
    slAirMonth = Month(slAirDate)
    If Len(slAirMonth) = 1 Then
        slAirMonth = "0" & slAirMonth
    End If
    slAirDay = Day(slAirDate)
    If Len(slAirDay) = 1 Then
        slAirDay = "0" & slAirDay
    End If
    slNowDate = Format$(gNow(), sgShowDateForm)
    slNowYear = Year(slNowDate)
    If Len(slNowYear) = 4 Then
        slNowYear = Mid$(slNowYear, 3)
    End If
    slNowMonth = Month(slNowDate)
    If Len(slNowMonth) = 1 Then
        slNowMonth = "0" & slNowMonth
    End If
    slNowDay = Day(slNowDate)
    If Len(slNowDay) = 1 Then
        slNowDay = "0" & slNowDay
    End If
    slChar = "A"
    Do
        ilRet = 0
        slToFile = sgMsgDirectory & "AutoExport_For_" & slAirYear & slAirMonth & slAirDay & "_On_" & slNowYear & slNowMonth & slNowDay & "_" & slChar & ".Txt"
        slDateTime = FileDateTime(slToFile)
        slChar = Chr(Asc(slChar) + 1)
    Loop While ilRet = 0
    On Error GoTo 0
    ilRet = 0
    On Error GoTo gOpenAutoMsgFileErr:
    hlMsg = FreeFile
    Open slToFile For Output As hlMsg
    If ilRet <> 0 Then
        Close hlMsg
        hlMsg = -1
        If igOperationMode = 1 Then
            gLogMsg "Open File " & slToFile & " error#" & Str$(Err.Number), "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Open File " & slToFile & " error#" & Str$(Err.Number), "EngrErrors.Txt", False
            MsgBox "Open File " & slToFile & Err.Description & "; Error #" & Err.Number, vbCritical
        End If
        gOpenAutoMsgFile = False
        Exit Function
    End If
    On Error GoTo 0
    slMsgFileName = slToFile
    gOpenAutoMsgFile = True
    Exit Function
gOpenAutoMsgFileErr:
    ilRet = 1
    Resume Next
End Function

Public Function gOpenAutoExportFile(tlSHE As SHE, slAirDate As String, slMsgFileName As String, hlExport As Integer) As Integer
    Dim slToFile As String
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim slTime As String
    Dim ilLoop As Integer
    Dim ilPos As Integer
    Dim ilPosE As Integer
    Dim slName As String
    Dim slPath As String
    Dim slDateTime As String
    Dim slChar As String
    Dim slSeqNo As String

    On Error GoTo gOpenAutoExportFileErr:
    'slNowDate = Format$(gNow(), sgShowDateForm)
    slName = ""
    slPath = ""
    For ilLoop = 0 To UBound(tgCurrAPE) - 1 Step 1
        If ((tgCurrAPE(ilLoop).sType = "CE") And (igRunningFrom = 1)) Or ((tgCurrAPE(ilLoop).sType = "SE") And (igRunningFrom = 0)) Then
            If ((tgCurrAPE(ilLoop).sSubType = "P") Or (Trim$(tgCurrAPE(ilLoop).sSubType) = "")) Then
                If (tlSHE.sLoadedAutoStatus = "L") Then
                    slName = Trim$(tgCurrAPE(ilLoop).sChgFileName) & "." & Trim$(tgCurrAPE(ilLoop).sChgFileExt)
                Else
                    slName = Trim$(tgCurrAPE(ilLoop).sNewFileName) & "." & Trim$(tgCurrAPE(ilLoop).sNewFileExt)
                End If
                ilPos = InStr(1, slName, "Date", vbTextCompare)
                If ilPos > 0 Then
                    If Trim$(tgCurrAPE(ilLoop).sDateFormat) <> "" Then
                        'slDate = Format$(slNowDate, tgCurrAPE(ilLoop).sDateFormat)
                        slDate = Format$(slAirDate, Trim$(tgCurrAPE(ilLoop).sDateFormat))
                    Else
                        'slDate = Format$(slNowDate, "yymmdd")
                        slDate = Format$(slAirDate, "yymmdd")
                    End If
                    slName = Left$(slName, ilPos - 1) & slDate & Mid(slName, ilPos + 4)
                End If
                ilPos = InStr(1, slName, "Time", vbTextCompare)
                If ilPos > 0 Then
                    If Trim$(tgCurrAPE(ilLoop).sTimeFormat) <> "" Then
                        slTime = Format$(slNowDate, Trim$(tgCurrAPE(ilLoop).sTimeFormat))
                    Else
                        slTime = Format$(slNowDate, "hhmmss")
                    End If
                    slName = Left$(slName, ilPos - 1) & slTime & Mid(slName, ilPos + 4)
                End If
                'Check for Sequence number
                If (tlSHE.sLoadedAutoStatus = "L") Then
                    ilPos = InStr(1, slName, "S", vbTextCompare)
                    If ilPos > 0 Then
                        ilPosE = ilPos + 1
                        Do While ilPosE <= Len(slName)
                            slChar = Mid$(slName, ilPosE, 1)
                            If StrComp(slChar, "S", vbTextCompare) <> 0 Then
                                Exit Do
                            End If
                            ilPosE = ilPosE + 1
                        Loop
                        slSeqNo = Trim$(Str$(tlSHE.iChgSeqNo + 1))
                        Do While Len(slSeqNo) < ilPosE - ilPos
                            slSeqNo = "0" & slSeqNo
                        Loop
                        Mid$(slName, ilPos, ilPosE - ilPos) = slSeqNo
                    End If
                End If
            End If
            'slPath = Trim$(tgCurrAPE(ilLoop).sPath)
            If (Not igTestSystem) And ((tgCurrAPE(ilLoop).sSubType = "P") Or (Trim$(tgCurrAPE(ilLoop).sSubType) = "")) Then
                slPath = Trim$(tgCurrAPE(ilLoop).sPath)
            ElseIf (igTestSystem) And (tgCurrAPE(ilLoop).sSubType = "T") Then
                slPath = Trim$(tgCurrAPE(ilLoop).sPath)
            End If
            If slPath <> "" Then
                If right(slPath, 1) <> "\" Then
                    slPath = slPath & "\"
                End If
            End If
            'Exit For
        End If
    Next ilLoop
    If slName = "" Then
        If igOperationMode = 1 Then
            gLogMsg "Load File Name missing for Client from Automation Equipment Definition", "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Load File Name missing for Client from Automation Equipment Definition", "EngrErrors.Txt", False
            MsgBox "Load File Name missing for Client from Automation Equipment Definition", vbCritical
        End If
        gOpenAutoExportFile = False
        Exit Function
    End If
    If slPath = "" Then
        If igOperationMode = 1 Then
            gLogMsg "Load Path missing for Client from Automation Equipment Definition", "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Load Path missing for Client from Automation Equipment Definition", "EngrErrors.Txt", False
            MsgBox "Load Path missing for Client from Automation Equipment Definition", vbCritical
        End If
        gOpenAutoExportFile = False
        Exit Function
    End If
    
    ilRet = 0
    slToFile = slPath & slName
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        Kill slToFile
    End If
    
    '3/31/13: Create temp file name
    ilRet = 0
    sgLoadFileName = slToFile
    ilPos = InStr(1, slToFile, ".", vbBinaryCompare)
    If ilPos > 0 Then
        Mid(slToFile, ilPos, 1) = "_"
        slToFile = slToFile & ".txt"
    Else
        gOpenAutoExportFile = False
        Exit Function
    End If
    sgTmpLoadFileName = slToFile
    slDateTime = FileDateTime(slToFile)
    If ilRet = 0 Then
        Kill slToFile
    End If
    
    ilRet = 0
    On Error GoTo gOpenAutoExportFileErr:
    hlExport = FreeFile
    Open slToFile For Output As hlExport
    If ilRet <> 0 Then
        Close hlExport
        hlExport = -1
        If igOperationMode = 1 Then
            gLogMsg "Open File " & slToFile & " error# " & Err.Number, "EngrServiceErrors.Txt", False
        Else
            gLogMsg "Open File " & slToFile & " error# " & Err.Number, "EngrErrors.Txt", False
            MsgBox "Open File " & slToFile & " error# " & Err.Number, vbCritical
        End If
        gOpenAutoExportFile = False
        Exit Function
    End If
    On Error GoTo 0
'    Print #hlExport, "** Test : " & Format$(Now, "m/d/yyyy") & " at " & Format$(Now, sgShowTimeWSecForm) & " **"
'    Print #hlExport, ""
    slMsgFileName = slToFile
    gOpenAutoExportFile = True
    Exit Function
gOpenAutoExportFileErr:
    ilRet = 1
    Resume Next
End Function

Public Sub gAutoSortTime(tlSEE() As SEE)
    Dim llSEE As Long
    Dim slEventCategory As String
    Dim slTime As String
    Dim slBusName As String
    Dim slSpotTime As String
    Dim ilETE As Integer
    Dim ilBDE As Integer
    
    ReDim tmSeeTimeSort(0 To UBound(tlSEE)) As SEETIMESORT
    For llSEE = 0 To UBound(tlSEE) - 1 Step 1
        slTime = Trim$(Str$(tlSEE(llSEE).lTime))
        Do While Len(slTime) < 10
            slTime = "0" & slTime
        Loop
        slBusName = ""
        For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
            If tlSEE(llSEE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                slBusName = Trim$(tgCurrBDE(ilBDE).sName)
                Exit For
            End If
        Next ilBDE
        Do While Len(slBusName) < 10
            slBusName = slBusName & " "
        Loop
        slSpotTime = Trim$(Str$(tlSEE(llSEE).lSpotTime))
        Do While Len(slSpotTime) < 10
            slSpotTime = "0" & slSpotTime
        Loop
        slEventCategory = ""
        For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
            If tlSEE(llSEE).iEteCode = tgCurrETE(ilETE).iCode Then
                slEventCategory = tgCurrETE(ilETE).sCategory
                Exit For
            End If
        Next ilETE
        If (slEventCategory = "S") Then
            tmSeeTimeSort(llSEE).sKey = slSpotTime & slBusName
        Else
            tmSeeTimeSort(llSEE).sKey = slTime & slBusName
        End If
        tmSeeTimeSort(llSEE).tSEE = tlSEE(llSEE)
    Next llSEE
    'Sort by Time
    If UBound(tmSeeTimeSort) - 1 > 0 Then
        ArraySortTyp fnAV(tmSeeTimeSort(), 0), UBound(tmSeeTimeSort), 0, LenB(tmSeeTimeSort(0)), 0, LenB(tmSeeTimeSort(0).sKey), 0
    End If
    For llSEE = 0 To UBound(tmSeeTimeSort) - 1 Step 1
        tlSEE(llSEE) = tmSeeTimeSort(llSEE).tSEE
    Next llSEE
    Erase tmSeeTimeSort
End Sub

Public Function gAutoExportRow(ilEteCode As Integer, slEventCategory As String, slEventAutoCode As String) As Integer
    Dim slStr As String
    Dim ilETE As Integer
    Dim ilEPE As Integer
    
    gAutoExportRow = False
    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        If tgCurrETE(ilETE).iCode = ilEteCode Then
            slEventCategory = tgCurrETE(ilETE).sCategory
            slEventAutoCode = tgCurrETE(ilETE).sAutoCodeChar
            If tgCurrETE(ilETE).sCategory = "A" Then
                Exit Function
            End If
            For ilEPE = 0 To UBound(tgCurrEPE) - 1 Step 1
                If tgCurrEPE(ilEPE).sType = "E" Then
                    If tgCurrETE(ilETE).iCode = tgCurrEPE(ilEPE).iEteCode Then
                        If tgCurrEPE(ilEPE).sBus = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sBusControl = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        'Event Type exported if any other column exported and tgStartColAFE.iEventType >0
                        'Event ID exported if any other column is exported and tgStartColAFE.iEventID > 0
                        If tgCurrEPE(ilEPE).sTime = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sStartType = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sFixedTime = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sEndType = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sDuration = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sMaterialType = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sAudioName = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sAudioItemID = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sAudioISCI = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sAudioControl = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sBkupAudioName = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sBkupAudioControl = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sProtAudioName = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sProtAudioItemID = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sProtAudioISCI = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sProtAudioControl = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sRelay1 = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sRelay2 = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sFollow = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sSilenceTime = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sSilence1 = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sSilence2 = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sSilence3 = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sSilence4 = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sStartNetcue = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sStopNetcue = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sTitle1 = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If tgCurrEPE(ilEPE).sTitle2 = "Y" Then
                            gAutoExportRow = True
                            Exit Function
                        End If
                        If (sgClientFields = "A") Then
                            If tgCurrEPE(ilEPE).sABCFormat = "Y" Then
                                gAutoExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sABCPgmCode = "Y" Then
                                gAutoExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sABCXDSMode = "Y" Then
                                gAutoExportRow = True
                                Exit Function
                            End If
                            If tgCurrEPE(ilEPE).sABCRecordItem = "Y" Then
                                gAutoExportRow = True
                                Exit Function
                            End If
                        End If
                        Exit For
                    End If
                End If
            Next ilEPE
            Exit For
        End If
    Next ilETE
End Function

Public Sub gAutoSendSEE(hlExport As Integer, slEventCategory As String, slEventAutoCode As String, slDate As String, ilEteCode As Integer, ilLength As Integer, tlSEE As SEE)
    Dim slStr As String
    Dim ilETE As Integer
    Dim ilBDE As Integer
    Dim ilCCE As Integer
    Dim ilTTE As Integer
    Dim ilMTE As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim ilRNE As Integer
    Dim ilFNE As Integer
    Dim ilSCE As Integer
    Dim ilNNE As Integer
    Dim slComment As String
    Dim ilRet As Integer
    Dim llEndTime As Long
    Dim slEndType As String
    
    '9/12/11: Bypass Spots wil Live copy, hard code test for L with copy cart name (Jim)
    If slEventCategory = "S" Then
        If Left(tlSEE.sAudioItemID, 1) = "L" Then
            Exit Sub
        End If
    End If
    slComment = ""
    If slEventCategory = "P" Then
        If tlSEE.l1CteCode > 0 Then
            ilRet = gGetRec_CTE_CommtsTitle(tlSEE.l1CteCode, "EngrLib- mMoveRecToCtrl for CTE", tmCTE)
            If ilRet Then
                slComment = Trim$(tmCTE.sComment)
            End If
        End If
    ElseIf slEventCategory = "S" Then
        If tlSEE.lAreCode > 0 Then
            ilRet = gGetRec_ARE_AdvertiserRefer(tlSEE.lAreCode, "EngrSchd-mMoveSEERecToCtrls: Advertiser", tmARE)
            If ilRet Then
                slComment = Trim$(tmARE.sName)
            End If
        End If
    End If
    smExportStr = String(ilLength, " ")
    slStr = ""
    For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
        If tlSEE.iBdeCode = tgCurrBDE(ilBDE).iCode Then
            slStr = Trim$(tgCurrBDE(ilBDE).sName)
            Exit For
        End If
    Next ilBDE
    mMakeExportStr tgStartColAFE.iBus, tgNoCharAFE.iBus, BUSNAMEINDEX, True, ilEteCode, slStr
    slStr = ""
    For ilCCE = 0 To UBound(tgCurrBusCCE) - 1 Step 1
        If tlSEE.iBusCceCode = tgCurrBusCCE(ilCCE).iCode Then
            slStr = Trim$(tgCurrBusCCE(ilCCE).sAutoChar)
            Exit For
        End If
    Next ilCCE
    mMakeExportStr tgStartColAFE.iBusControl, tgNoCharAFE.iBusControl, BUSCTRLINDEX, True, ilEteCode, slStr
    If slEventCategory = "P" Then
        slStr = gLongToStrTimeInTenth(tlSEE.lTime)
    Else
        slStr = gLongToStrTimeInTenth(tlSEE.lSpotTime)
    End If
    mMakeExportStr tgStartColAFE.iTime, tgNoCharAFE.iTime, TIMEINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilTTE = 0 To UBound(tgCurrStartTTE) - 1 Step 1
        If tlSEE.iStartTteCode = tgCurrStartTTE(ilTTE).iCode Then
            slStr = Trim$(tgCurrStartTTE(ilTTE).sName)
            Exit For
        End If
    Next ilTTE
    mMakeExportStr tgStartColAFE.iStartType, tgNoCharAFE.iStartType, STARTTYPEINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilTTE = 0 To UBound(tgCurrEndTTE) - 1 Step 1
        If tlSEE.iEndTteCode = tgCurrEndTTE(ilTTE).iCode Then
            slStr = Trim$(tgCurrEndTTE(ilTTE).sName)
            Exit For
        End If
    Next ilTTE
    mMakeExportStr tgStartColAFE.iEndType, tgNoCharAFE.iEndType, ENDTYPEINDEX, False, ilEteCode, slStr
    slEndType = slStr
    ''12/11/12: Show Duration of zero as 00:00:00.0
    ''If (tlSEE.lDuration > 0) Then
    '2/22/13: Don't show duration if zero and End Type = MAN or EXT
    'If (tlSEE.lDuration >= 0) Then
    If (tlSEE.lDuration > 0) Or ((tlSEE.lDuration = 0) And (Trim$(slEndType) <> "MAN") And (Trim$(slEndType) <> "EXT")) Then
        slStr = gLongToStrLengthInTenth(tlSEE.lDuration, True)
    Else
        slStr = ""
    End If
    mMakeExportStr tgStartColAFE.iDuration, tgNoCharAFE.iDuration, DURATIONINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilMTE = 0 To UBound(tgCurrMTE) - 1 Step 1
        If tlSEE.iMteCode = tgCurrMTE(ilMTE).iCode Then
            slStr = Trim$(tgCurrMTE(ilMTE).sName)
            Exit For
        End If
    Next ilMTE
    mMakeExportStr tgStartColAFE.iMaterialType, tgNoCharAFE.iMaterialType, MATERIALINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
        If tlSEE.iAudioAseCode = tgCurrASE(ilASE).iCode Then
            For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                    slStr = Trim$(tgCurrANE(ilANE).sName)
                End If
            Next ilANE
            Exit For
        End If
    Next ilASE
    mMakeExportStr tgStartColAFE.iAudioName, tgNoCharAFE.iAudioName, AUDIONAMEINDEX, True, ilEteCode, slStr
    slStr = Trim$(tlSEE.sAudioItemID)
    mMakeExportStr tgStartColAFE.iAudioItemID, tgNoCharAFE.iAudioItemID, AUDIOITEMIDINDEX, False, ilEteCode, slStr
    slStr = Trim$(tlSEE.sAudioISCI)
    mMakeExportStr tgStartColAFE.iAudioISCI, tgNoCharAFE.iAudioISCI, AUDIOISCIINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        If tlSEE.iAudioCceCode = tgCurrAudioCCE(ilCCE).iCode Then
            slStr = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
            Exit For
        End If
    Next ilCCE
    mMakeExportStr tgStartColAFE.iAudioControl, tgNoCharAFE.iAudioControl, AUDIOCTRLINDEX, True, ilEteCode, slStr
    slStr = ""
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        If tlSEE.iBkupAneCode = tgCurrANE(ilANE).iCode Then
            slStr = Trim$(tgCurrANE(ilANE).sName)
            Exit For
        End If
    Next ilANE
    mMakeExportStr tgStartColAFE.iBkupAudioName, tgNoCharAFE.iBkupAudioName, BACKUPNAMEINDEX, True, ilEteCode, slStr
    slStr = ""
    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        If tlSEE.iBkupCceCode = tgCurrAudioCCE(ilCCE).iCode Then
            slStr = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
            Exit For
        End If
    Next ilCCE
    mMakeExportStr tgStartColAFE.iBkupAudioControl, tgNoCharAFE.iBkupAudioControl, BACKUPCTRLINDEX, True, ilEteCode, slStr
    slStr = ""
    For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        If tlSEE.iProtAneCode = tgCurrANE(ilANE).iCode Then
            slStr = Trim$(tgCurrANE(ilANE).sName)
            Exit For
        End If
    Next ilANE
    mMakeExportStr tgStartColAFE.iProtAudioName, tgNoCharAFE.iProtAudioName, PROTNAMEINDEX, True, ilEteCode, slStr
    slStr = Trim$(tlSEE.sProtItemID)
    mMakeExportStr tgStartColAFE.iProtItemID, tgNoCharAFE.iProtItemID, PROTITEMIDINDEX, False, ilEteCode, slStr
    slStr = Trim$(tlSEE.sProtISCI)
    mMakeExportStr tgStartColAFE.iProtISCI, tgNoCharAFE.iProtISCI, PROTISCIINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilCCE = 0 To UBound(tgCurrAudioCCE) - 1 Step 1
        If tlSEE.iProtCceCode = tgCurrAudioCCE(ilCCE).iCode Then
            slStr = Trim$(tgCurrAudioCCE(ilCCE).sAutoChar)
            Exit For
        End If
    Next ilCCE
    mMakeExportStr tgStartColAFE.iProtAudioControl, tgNoCharAFE.iProtAudioControl, PROTCTRLINDEX, True, ilEteCode, slStr
    slStr = ""
    For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
        If tlSEE.i1RneCode = tgCurrRNE(ilRNE).iCode Then
            slStr = Trim$(tgCurrRNE(ilRNE).sName)
            Exit For
        End If
    Next ilRNE
    mMakeExportStr tgStartColAFE.iRelay1, tgNoCharAFE.iRelay1, RELAY1INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilRNE = 0 To UBound(tgCurrRNE) - 1 Step 1
        If tlSEE.i2RneCode = tgCurrRNE(ilRNE).iCode Then
            slStr = Trim$(tgCurrRNE(ilRNE).sName)
            Exit For
        End If
    Next ilRNE
    mMakeExportStr tgStartColAFE.iRelay2, tgNoCharAFE.iRelay2, RELAY2INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilFNE = 0 To UBound(tgCurrFNE) - 1 Step 1
        If tlSEE.iFneCode = tgCurrFNE(ilFNE).iCode Then
            slStr = Trim$(tgCurrFNE(ilFNE).sName)
            Exit For
        End If
    Next ilFNE
    mMakeExportStr tgStartColAFE.iFollow, tgNoCharAFE.iFollow, FOLLOWINDEX, False, ilEteCode, slStr
    If tlSEE.lSilenceTime > 0 Then
        slStr = gLongToLength(tlSEE.lSilenceTime, False)    'gLongToStrLengthInTenth(tlSEE.lSilenceTime, False)
    Else
        slStr = ""
    End If
    mMakeExportStr tgStartColAFE.iSilenceTime, tgNoCharAFE.iSilenceTime, SILENCETIMEINDEX, False, ilEteCode, slStr
    slStr = ""
    For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
        If tlSEE.i1SceCode = tgCurrSCE(ilSCE).iCode Then
            slStr = Trim$(tgCurrSCE(ilSCE).sAutoChar)
            Exit For
        End If
    Next ilSCE
    mMakeExportStr tgStartColAFE.iSilence1, tgNoCharAFE.iSilence1, SILENCE1INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
        If tlSEE.i2SceCode = tgCurrSCE(ilSCE).iCode Then
            slStr = Trim$(tgCurrSCE(ilSCE).sAutoChar)
            Exit For
        End If
    Next ilSCE
    mMakeExportStr tgStartColAFE.iSilence2, tgNoCharAFE.iSilence2, SILENCE2INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
        If tlSEE.i3SceCode = tgCurrSCE(ilSCE).iCode Then
            slStr = Trim$(tgCurrSCE(ilSCE).sAutoChar)
            Exit For
        End If
    Next ilSCE
    mMakeExportStr tgStartColAFE.iSilence3, tgNoCharAFE.iSilence3, SILENCE3INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilSCE = 0 To UBound(tgCurrSCE) - 1 Step 1
        If tlSEE.i4SceCode = tgCurrSCE(ilSCE).iCode Then
            slStr = Trim$(tgCurrSCE(ilSCE).sAutoChar)
            Exit For
        End If
    Next ilSCE
    mMakeExportStr tgStartColAFE.iSilence4, tgNoCharAFE.iSilence4, SILENCE4INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
        If tlSEE.iStartNneCode = tgCurrNNE(ilNNE).iCode Then
            slStr = Trim$(tgCurrNNE(ilNNE).sName)
            Exit For
        End If
    Next ilNNE
    mMakeExportStr tgStartColAFE.iStartNetcue, tgNoCharAFE.iStartNetcue, NETCUE1INDEX, False, ilEteCode, slStr
    slStr = ""
    For ilNNE = 0 To UBound(tgCurrNNE) - 1 Step 1
        If tlSEE.iEndNneCode = tgCurrNNE(ilNNE).iCode Then
            slStr = Trim$(tgCurrNNE(ilNNE).sName)
            Exit For
        End If
    Next ilNNE
    mMakeExportStr tgStartColAFE.iStopNetcue, tgNoCharAFE.iStopNetcue, NETCUE2INDEX, False, ilEteCode, slStr
    If (slEventCategory = "P") Then
        mMakeExportStr tgStartColAFE.iTitle1, tgNoCharAFE.iTitle1, TITLE1INDEX, False, ilEteCode, slComment
    Else
        mMakeExportStr tgStartColAFE.iTitle1, tgNoCharAFE.iTitle1, TITLE1INDEX, False, ilEteCode, slComment
    End If
    slStr = ""
    If tlSEE.l2CteCode > 0 Then
        ilRet = gGetRec_CTE_CommtsTitle(tlSEE.l2CteCode, "EngrLib- mMoveRecToCtrl for CTE", tmCTE)
        If ilRet Then
            slStr = Trim$(tmCTE.sComment)
        End If
    End If
    mMakeExportStr tgStartColAFE.iTitle2, tgNoCharAFE.iTitle2, TITLE2INDEX, False, ilEteCode, slStr
    If sgClientFields = "A" Then
        slStr = Trim$(tlSEE.sABCFormat)
        mMakeExportStr tgStartColAFE.iABCFormat, tgNoCharAFE.iABCFormat, ABCFORMATINDEX, False, ilEteCode, slStr
        slStr = Trim$(tlSEE.sABCPgmCode)
        mMakeExportStr tgStartColAFE.iABCPgmCode, tgNoCharAFE.iABCPgmCode, ABCPGMCODEINDEX, False, ilEteCode, slStr
        slStr = Trim$(tlSEE.sABCXDSMode)
        mMakeExportStr tgStartColAFE.iABCXDSMode, tgNoCharAFE.iABCXDSMode, ABCXDSMODEINDEX, False, ilEteCode, slStr
        slStr = Trim$(tlSEE.sABCRecordItem)
        mMakeExportStr tgStartColAFE.iABCRecordItem, tgNoCharAFE.iABCRecordItem, ABCRECORDITEMINDEX, False, ilEteCode, slStr
    End If
    'Event Type
    'If mColOk(ilEteCode, EVENTTYPEINDEX) Then
        If tgStartColAFE.iEventType > 0 Then
            slStr = slEventAutoCode
            Do While Len(slStr) < tgNoCharAFE.iEventType
                slStr = slStr & " "
            Loop
            Mid(smExportStr, tgStartColAFE.iEventType, tgNoCharAFE.iEventType) = slStr
        End If
    'End If
    'Fixed
    If gExportCol(ilEteCode, FIXEDINDEX) Then
        slStr = Trim$(tlSEE.sFixedTime)
        If slStr = "Y" Then
            If tgStartColAFE.iFixedTime > 0 Then
                slStr = Trim$(tgAEE.sFixedTimeChar)
                Do While Len(slStr) < tgNoCharAFE.iFixedTime
                    slStr = slStr & " "
                Loop
                Mid(smExportStr, tgStartColAFE.iFixedTime, tgNoCharAFE.iFixedTime) = slStr
            End If
        End If
    End If
    'Date
    If tgStartColAFE.iDate > 0 Then
        slStr = slDate
        Do While Len(slStr) < tgNoCharAFE.iDate
            slStr = slStr & " "
        Loop
        Mid(smExportStr, tgStartColAFE.iDate, tgNoCharAFE.iDate) = slStr
    End If
    'End Time
    If gExportCol(ilEteCode, DURATIONINDEX) Then
        If tgStartColAFE.iEndTime > 0 Then
            '2/22/13: Don't show Out Time if duration is zero and End Type = MAN or EXT
            If (tlSEE.lDuration > 0) Or ((tlSEE.lDuration = 0) And (Trim$(slEndType) <> "MAN") And (Trim$(slEndType) <> "EXT")) Then
                If slEventCategory = "P" Then
                    llEndTime = tlSEE.lTime + tlSEE.lDuration
                Else
                    llEndTime = tlSEE.lSpotTime + tlSEE.lDuration
                End If
                If llEndTime > 864000 Then
                    llEndTime = llEndTime - 864000
                End If
                slStr = gLongToStrLengthInTenth(llEndTime, True)
            Else
                slStr = ""
            End If
            Do While Len(slStr) < tgNoCharAFE.iEndTime
                slStr = slStr & " "
            Loop
            Mid(smExportStr, tgStartColAFE.iEndTime, tgNoCharAFE.iEndTime) = slStr
        End If
    End If
    'Event ID
    If tgStartColAFE.iEventID > 0 Then
        slStr = Trim$(Str$(tlSEE.lEventID))
        Do While Len(slStr) < tgNoCharAFE.iEventID
            slStr = "0" & slStr
        Loop
        Mid(smExportStr, tgStartColAFE.iEventID, tgNoCharAFE.iEventID) = slStr
    End If
    Print #hlExport, smExportStr

End Sub

Private Sub mMakeExportStr(ilStartCol As Integer, ilNoChar As Integer, llCol As Long, ilUCase As Integer, ilEteCode As Integer, slInStr As String)
    Dim slStr As String
    If (ilStartCol > 0) And (gExportCol(ilEteCode, llCol)) Then
        slStr = Trim$(slInStr)
        Do While Len(slStr) < ilNoChar
            slStr = slStr & " "
        Loop
        If ilUCase Then
            slStr = UCase$(slStr)
        End If
        Mid(smExportStr, ilStartCol, ilNoChar) = slStr
    End If
End Sub

Public Sub gRenameExportFile()
    Name sgTmpLoadFileName As sgLoadFileName
End Sub
