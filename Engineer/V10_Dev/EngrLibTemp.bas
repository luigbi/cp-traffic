Attribute VB_Name = "EngrLibTemp"
'
' Release: 1.0
'
' Description:
'   This file contains the Routines used by Library and Template Definition

Option Explicit

Private smDays() As String
Private smHours() As String

Private smCurrLibDBEStamp As String
Private tmCurrLibDBE() As DBE
Private smCurrLibDEEStamp As String
Private tmCurrLibDEE() As DEE
Private smCurrLibEBEStamp As String
Private tmCurrLibEBE() As EBE
Private smCurrLibDHEStamp As String
Private tmCurrLibDHE() As DHE
Private tmCurrTempTSE() As TSE
Private smCurrTempDHETSEStamp As String
Private tmCurrTempDHETSE() As DHETSE

Private tmSHE As SHE
Private tmCurrSHE() As SHE
Private smCurrSEEStamp As String
Private tmCurrSEE() As SEE

Private lmGenDate As Long
Private lmGenTime As Long
Private lmGridEventRow As Long
Private smTypeSource As String
Private tmConflictResults() As CONFLICTRESULTS

Type CONFLICTTEST
    lRow As Long
    sType As String * 1     'B=Bus; 1=Primary audio; 2=Protection audio; 3=Backup audio
    sDays As String * 7
    lEventStartTime As Long
    lEventEndTime As Long
End Type
Dim tmConflictTest() As CONFLICTTEST
Dim tmConflictLib() As CONFLICTTEST





Public Function gAudioConflicts(slEvtType1 As String, slEvtType2 As String, slAudio1 As String, slAudio2 As String, slItemID1 As String, slItemID2 As String, llStartTime1 As Long, llEndTime1 As Long, llStartTime2 As Long, llEndTime2 As Long, ilAdjTime As Integer, slInBus1 As String, slInBus2 As String) As Integer
    Dim ilANE As Integer
    Dim ilATE As Integer
    Dim llPreTime As Long
    Dim llPostTime As Long
    Dim llAdjStartTime1 As Long
    Dim llAdjEndTime1 As Long
    Dim llAdjStartTime2 As Long
    Dim llAdjEndTime2 As Long
    Dim slBus1 As String
    Dim slBus2 As String
    Dim ilETE As Integer
    
    
    gAudioConflicts = False
    If slAudio1 = "" Then
        Exit Function
    End If
    If StrComp(slAudio1, "[None]", vbTextCompare) = 0 Then
        Exit Function
    End If
    'Only test bus is avail event
    slBus1 = slInBus1
    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        If StrComp(Trim$(tgCurrETE(ilETE).sName), Trim$(slEvtType1), vbTextCompare) = 0 Then
            If tgCurrETE(ilETE).sCategory <> "A" Then
                slBus1 = ""
            End If
            Exit For
        End If
    Next ilETE
    slBus2 = slInBus2
    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        If StrComp(Trim$(tgCurrETE(ilETE).sName), Trim$(slEvtType2), vbTextCompare) = 0 Then
            If tgCurrETE(ilETE).sCategory <> "A" Then
                slBus2 = ""
            End If
            Exit For
        End If
    Next ilETE
    If StrComp(slAudio1, slAudio2, vbTextCompare) = 0 Then
        'Same audio allowed if times match
        
        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
        '    If StrComp(slAudio1, Trim$(tgCurrANE(ilANE).sName), vbTextCompare) = 0 Then
        ilANE = gBinarySearchName(slAudio1, tgCurrANE_Name())
        If ilANE <> -1 Then
            ilANE = gBinarySearchANE(tgCurrANE_Name(ilANE).iCode, tgCurrANE())
            If ilANE <> -1 Then

                If tgCurrANE(ilANE).sCheckConflicts <> "N" Then
                    If (llStartTime1 <> llStartTime2) Or (llEndTime1 <> llEndTime2) Then
                        If tgSOE.sMatchANotT <> "N" Then
                            For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
                                If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                                    If ilAdjTime Then
                                        llPreTime = tgCurrATE(ilATE).lPreBufferTime
                                        llPostTime = tgCurrATE(ilATE).lPostBufferTime
                                        llAdjStartTime1 = llStartTime1 - llPreTime
                                        If llAdjStartTime1 < 0 Then
                                            llAdjStartTime1 = 0
                                        End If
                                        llAdjEndTime1 = llEndTime1 + llPostTime
                                        If llAdjEndTime1 > 864000 Then
                                            llAdjEndTime1 = 864000
                                        End If
                                        llAdjStartTime2 = llStartTime2 - llPreTime
                                        If llAdjStartTime2 < 0 Then
                                            llAdjStartTime2 = 0
                                        End If
                                        llAdjEndTime2 = llEndTime2 + llPostTime
                                        If llAdjEndTime2 > 864000 Then
                                            llAdjEndTime2 = 864000
                                        End If
                                    Else
                                        llAdjStartTime1 = llStartTime1
                                        llAdjEndTime1 = llEndTime1
                                        llAdjStartTime2 = llStartTime2
                                        llAdjEndTime2 = llEndTime2
                                    End If
                                    'Don't need to twest if start times match as done with buses because audio can switch ok
                                    If (llAdjEndTime2 > llAdjStartTime1) And (llAdjStartTime2 < llAdjEndTime1) Then
                                        gAudioConflicts = True
                                        Exit Function
                                    End If
                                End If
                            Next ilATE
                        End If
                    Else
                        If StrComp(Trim$(slBus1), Trim$(slBus2), vbTextCompare) <> 0 Then
                            If tgSOE.sMatchATNotB <> "N" Then
                                gAudioConflicts = True
                                Exit Function
                            End If
                        Else
                            If tgSOE.sMatchATBNotI <> "N" Then
                                If ((Trim$(slItemID1) <> "") And (Trim$(slItemID2) <> "")) Or ((Trim$(slItemID1) = "") And (Trim$(slItemID2) <> "")) Or ((Trim$(slItemID1) <> "") And (Trim$(slItemID2) = "")) Then
                                    If StrComp(slItemID1, slItemID2, vbTextCompare) <> 0 Then
                                        gAudioConflicts = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
        '        Exit For
            End If
        End If
        'Next ilANE
    End If
    
End Function

Public Function gCheckConflicts(slType As String, llDheCode As Long, llCurrentDHE As Long, slInStartDate As String, slInEndDate As String, slTempStartTime As String, grdEvents As MSHFlexGrid, ilCols() As Integer, tlConflictList() As CONFLICTLIST, hlSEE As Integer) As Integer
    'slType- L=Library, T=Template, S=Schedule
    'ilCols(0) = ERRORFLAGINDEX
    'ilCols(1) = EVENTTYPEINDEX
    'ilCols(2) = AIRHOURSINDEX
    'ilCols(3) = AIRDAYSINDEX
    'ilCols(4) = TIMEINDEX
    'ilCols(5) = DURATIONINDEX
    'ilCols(6) = BUSNAMEINDEX
    'ilCols(7) = AUDIONAMEINDEX
    'ilCols(8) = PROTNAMEINDEX
    'ilCols(9) = BACKUPNAMEINDEX
    'ilCols(10) = AUDIOITEMIDINDEX
    'ilCols(11) = PROTITEMIDINDEX
    'ilCols(12) = BACKUPITEMIDINDEX
    'ilCols(13) = CHGSTATUSINDEX
    'ilCols(14) = Ignore Conflicts (A=Audio, B=Bus; I=Both)
    'ilCols(15) = DEECode Index
    Dim llRow As Long
    Dim ilRet As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llLibStartDate As Long
    Dim slStartTime As String
    Dim slEndTime As String
    Dim llOffsetStartTime As Long
    Dim llOffsetEndTime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilDHE As Integer
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim slStart As String
    Dim slEnd As String
    Dim slDays As String
    Dim ilDays As Integer
    Dim ilDay As Integer
    Dim ilPos As Integer
    Dim slStr As String
    Dim ilDBE As Integer
    Dim ilSet As Integer
    Dim llLength As Long
    Dim ilTest As Integer
    Dim ilLoop As Integer
    Dim ilLib As Integer
    Dim llEvent As Long
    Dim llOffsetEventStartTime As Long
    Dim llOffsetEventEndTime As Long
    Dim ilEventTempHour As Integer
    Dim llEventStartTime As Long
    Dim llEventEndTime As Long
    Dim slBuses As String
    Dim ilEBE As Integer
    Dim ilBDE As Integer
    Dim ilASE As Integer
    Dim ilANE As Integer
    Dim slEvtType1 As String
    Dim slEvtType2 As String
    Dim slPriAudio As String
    Dim slProtAudio As String
    Dim slBkupAudio As String
    Dim slPriItemID1 As String
    Dim slProtItemID1 As String
    Dim slBkupItemID1 As String
    Dim slPriAudio2 As String
    Dim slProtAudio2 As String
    Dim slBkupAudio2 As String
    Dim slPriItemID2 As String
    Dim slProtItemID2 As String
    Dim slBkupItemID2 As String
    Dim slHours As String
    Dim slTestHours As String
    Dim ilHour As Integer
    Dim ilTempHour As Integer
    Dim ilATE As Integer
    Dim ilFound As Integer
    Dim ilError As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim ilUpper As Integer
    Dim slDateTime As String
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llNowTime As Long
    Dim ilSHE As Integer
    Dim ilBus As Integer
    Dim ilDEE As Integer
    'Dim tlDEE As DEE
    Dim tlTSE As TSE
    'Dim tlDHE As DHE
    ReDim slBusNames(1 To 1) As String
    Dim ilConflictIndex As Integer
    Dim ilStartConflictIndex As Integer
    Dim ilStartHour As Integer
    Dim ilTestHour As Integer
    Dim ilCheck As Integer
    Dim slDateStartRange As String
    Dim slDateEndRange As String
    Dim llDateStartRange As Long
    Dim llDateEndRange As Long
    Dim slLibDays As String
    Dim slEventDays As String
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llUpper As Long
    Dim llLibEvent As Long
    Dim ilPass As Integer       'Test from start date to End Date; 1= Test only events that map to previous date from start date; 2= only test events the map to next date passed end date
    Dim slSpecDay As String
    Dim slIgnoreConflict As String
    Dim ilETE As Integer
    
    gCheckConflicts = False
    gConflictPop
    'For llEvent = grdEvents.FixedRows To grdEvents.Rows - 1 Step 1
    '    If Trim$(grdEvents.TextMatrix(llEvent, ilCols(1))) <> "" Then
    '        grdEvents.TextMatrix(llEvent, ilCols(0)) = "0"
    '    End If
    'Next llEvent
    For ilPass = 0 To 2 Step 1
        slStartDate = slInStartDate
        slEndDate = slInEndDate
        If ilPass = 1 Then
            llStartDate = gDateValue(slStartDate)
            slStartDate = gAdjYear(Format$(llStartDate - 1, "ddddd"))
            slEndDate = slStartDate
        ElseIf ilPass = 2 Then
            If gDateValue(slInEndDate) = gDateValue("12/31/2069") Then
                Exit For
            End If
            llEndDate = gDateValue(slEndDate)
            slStartDate = gAdjYear(Format$(llEndDate + 1, "ddddd"))
            slEndDate = slStartDate
        End If
        ReDim tmConflictTest(1 To 1) As CONFLICTTEST
        llUpper = 1
        For llEvent = grdEvents.FixedRows To grdEvents.Rows - 1 Step 1
            If (Trim$(grdEvents.TextMatrix(llEvent, ilCols(1))) <> "") Then
                ilSet = True
                If (ilCols(13) <> -1) Then
                    If Trim$(grdEvents.TextMatrix(llEvent, ilCols(13))) = "N" Then
                        ilSet = False
                    End If
                End If
                If ilSet Then
                    If slType = "L" Then
                        slStr = Trim$(grdEvents.TextMatrix(llEvent, ilCols(2)))
                        slHours = gCreateHourStr(slStr)
                        slStr = Trim$(grdEvents.TextMatrix(llEvent, ilCols(3)))
                        slDays = gCreateDayStr(slStr)
                        slTestHours = slHours
                    ElseIf slType = "S" Then
                        slDays = String(7, "N")
                        Select Case Weekday(slInStartDate)
                            Case vbMonday
                                Mid(slDays, 1, 1) = "Y"
                            Case vbTuesday
                                Mid(slDays, 2, 1) = "Y"
                            Case vbWednesday
                                Mid(slDays, 3, 1) = "Y"
                            Case vbThursday
                                Mid(slDays, 4, 1) = "Y"
                            Case vbFriday
                                Mid(slDays, 5, 1) = "Y"
                            Case vbSaturday
                                Mid(slDays, 6, 1) = "Y"
                            Case vbSunday
                                Mid(slDays, 7, 1) = "Y"
                        End Select
                        slTestHours = String(24, "Y")
                    Else
                        slStr = Trim$(grdEvents.TextMatrix(llEvent, ilCols(2)))
                        slHours = gCreateHourStr(slStr)
                        slDays = String(7, "N")
                        Select Case Weekday(slInStartDate)
                            Case vbMonday
                                Mid(slDays, 1, 1) = "Y"
                            Case vbTuesday
                                Mid(slDays, 2, 1) = "Y"
                            Case vbWednesday
                                Mid(slDays, 3, 1) = "Y"
                            Case vbThursday
                                Mid(slDays, 4, 1) = "Y"
                            Case vbFriday
                                Mid(slDays, 5, 1) = "Y"
                            Case vbSaturday
                                Mid(slDays, 6, 1) = "Y"
                            Case vbSunday
                                Mid(slDays, 7, 1) = "Y"
                        End Select
                        ilStartHour = Hour(slTempStartTime)
                        If ilStartHour <> 0 Then
                            slTestHours = String(24, "N")
                            ilHour = ilStartHour
                            For ilLoop = 0 To 23 Step 1
                                Mid$(slTestHours, ilHour + 1, 1) = Mid$(slHours, ilLoop + 1, 1)
                                ilHour = ilHour + 1
                                If ilHour > 23 Then
                                    Exit For
                                End If
                            Next ilLoop
                        Else
                            slTestHours = slHours
                        End If
                    End If
                    slStr = grdEvents.TextMatrix(llEvent, ilCols(4))
                    llOffsetEventStartTime = gStrLengthInTenthToLong(slStr)
                    slStr = grdEvents.TextMatrix(llEvent, ilCols(5))
                    llOffsetEventEndTime = llOffsetEventStartTime + gStrLengthInTenthToLong(slStr)  ' - 1
                    If llOffsetEventEndTime < llOffsetEventStartTime Then
                        llOffsetEventEndTime = llOffsetEventStartTime
                    End If
                    slIgnoreConflict = Trim$(grdEvents.TextMatrix(llEvent, ilCols(14)))
                    If (slIgnoreConflict = "A") Or (slIgnoreConflict = "I") Then
                        slPriAudio = ""
                        slProtAudio = ""
                        slBkupAudio = ""
                    Else
                        slPriAudio = Trim$(grdEvents.TextMatrix(llEvent, ilCols(7)))
                        slProtAudio = Trim$(grdEvents.TextMatrix(llEvent, ilCols(8)))
                        slBkupAudio = Trim$(grdEvents.TextMatrix(llEvent, ilCols(9)))
                    End If
                    If slType = "S" Then
                        llEventStartTime = llOffsetEventStartTime
                        llEventEndTime = llOffsetEventEndTime
                        If ilPass = 0 Then
                            mCreateBusRecs 1, llEvent, "B", slIgnoreConflict, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                            mCreateAudioRecs 1, llEvent, "1", slPriAudio, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                            mCreateAudioRecs 1, llEvent, "2", slProtAudio, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                            mCreateAudioRecs 1, llEvent, "3", slBkupAudio, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                            mCreateBusRecs 2, llEvent, "B", slIgnoreConflict, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                            mCreateAudioRecs 2, llEvent, "1", slPriAudio, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                            mCreateAudioRecs 2, llEvent, "2", slProtAudio, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                            mCreateAudioRecs 2, llEvent, "3", slBkupAudio, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                        Else
                            mCreateBusRecs ilPass, llEvent, "B", slIgnoreConflict, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                            mCreateAudioRecs ilPass, llEvent, "1", slPriAudio, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                            mCreateAudioRecs ilPass, llEvent, "2", slProtAudio, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                            mCreateAudioRecs ilPass, llEvent, "3", slBkupAudio, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                        End If
                    Else
                        For ilHour = 1 To 24 Step 1
                            If (Mid$(slTestHours, ilHour, 1) = "Y") Then
                                llEventStartTime = 36000 * (ilHour - 1) + llOffsetEventStartTime
                                llEventEndTime = 36000 * (ilHour - 1) + llOffsetEventEndTime
                                If slType = "T" Then
                                    llEventStartTime = llEventStartTime + 10 * (gTimeToLong(slTempStartTime, False) Mod 3600)
                                    llEventEndTime = llEventEndTime + 10 * (gTimeToLong(slTempStartTime, False) Mod 3600)
                                End If
                                mCreateBusRecs ilPass, llEvent, "B", slIgnoreConflict, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                                mCreateAudioRecs ilPass, llEvent, "1", slPriAudio, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                                mCreateAudioRecs ilPass, llEvent, "2", slProtAudio, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                                mCreateAudioRecs ilPass, llEvent, "3", slBkupAudio, llEventStartTime, llEventEndTime, slDays, tmConflictTest()
                            End If
                        Next ilHour
                    End If
                End If
            End If
        Next llEvent
        If UBound(tmConflictTest) > LBound(tmConflictTest) Then
            If ilPass = 0 Then
                llDateStartRange = gDateValue(slStartDate) - 1
                slDateStartRange = Format$(llDateStartRange, "mm/dd/yyyy")
                If gDateValue(slEndDate) < gDateValue("12/31/2069") Then
                    llDateEndRange = gDateValue(slEndDate) + 1
                    slDateEndRange = Format$(llDateEndRange, "mm/dd/yyyy")
                Else
                    llDateEndRange = gDateValue(slEndDate)
                    slDateEndRange = Format$(llDateEndRange, "mm/dd/yyyy")
                End If
            Else
                llDateStartRange = gDateValue(slStartDate)
                llDateEndRange = gDateValue(slEndDate)
                slDateStartRange = Format$(llDateStartRange, "mm/dd/yyyy")
                slDateEndRange = Format$(llDateEndRange, "mm/dd/yyyy")
            End If
            'Check for conflicts
            smCurrLibDHEStamp = ""
            'ilRet = gGetTypeOfRecs_DHE_DayHeaderInfoForLibByRange("C", "A", slStartDate, slEndDate, smCurrLibDHEStamp, "EngrLib-mPopulate", tmCurrLibDHE())
            ilRet = gGetTypeOfRecs_DHE_DayHeaderInfoForLibByRange("C", "A", slDateStartRange, slDateEndRange, smCurrLibDHEStamp, "EngrLib-mPopulate", tmCurrLibDHE())
            ReDim tmCurrTempTSE(0 To UBound(tmCurrLibDHE)) As TSE
            smCurrTempDHETSEStamp = ""
            ilRet = gGetTypeOfRecs_DHETSE_DayHeaderInfoForTempByRange("C", slDateStartRange, slDateEndRange, smCurrTempDHETSEStamp, "EngrLib-mPopulate", tmCurrTempDHETSE())
            'Note:  Start and End Date set to Template Log Date.  Days is also set based on Log Date in the routine
            For ilDHE = 0 To UBound(tmCurrTempDHETSE) - 1 Step 1
                If (tmCurrTempDHETSE(ilDHE).tDHE.sState <> "D") And (tmCurrTempDHETSE(ilDHE).tDHE.sState <> "L") And (tmCurrTempDHETSE(ilDHE).tTSE.sState <> "D") And (tmCurrTempDHETSE(ilDHE).tTSE.sState <> "L") Then
                    If (tmCurrTempDHETSE(ilDHE).tDHE.sIgnoreConflicts <> "I") Then
                        LSet tmCurrLibDHE(UBound(tmCurrLibDHE)) = tmCurrTempDHETSE(ilDHE).tDHE
                        LSet tmCurrTempTSE(UBound(tmCurrTempTSE)) = tmCurrTempDHETSE(ilDHE).tTSE
                        tmCurrLibDHE(UBound(tmCurrLibDHE)).sStartDate = tmCurrTempDHETSE(ilDHE).tTSE.sLogDate
                        tmCurrLibDHE(UBound(tmCurrLibDHE)).sEndDate = tmCurrTempDHETSE(ilDHE).tTSE.sLogDate
                        tmCurrLibDHE(UBound(tmCurrLibDHE)).sStartTime = tmCurrTempDHETSE(ilDHE).tTSE.sStartTime
                        tmCurrLibDHE(UBound(tmCurrLibDHE)).lLength = 10 * (gTimeToLong(tmCurrTempDHETSE(ilDHE).tTSE.sStartTime, False) Mod 3600)
                        ReDim Preserve tmCurrLibDHE(0 To UBound(tmCurrLibDHE) + 1) As DHE
                        ReDim Preserve tmCurrTempTSE(0 To UBound(tmCurrTempTSE) + 1) As TSE
                    End If
                End If
            Next ilDHE
            'Merge the Library and template
            llStartDate = gDateValue(slDateStartRange)   'gDateValue(slStartDate)
            llEndDate = gDateValue(slDateEndRange)   'gDateValue(slEndDate)
            llLibStartDate = llStartDate
            slDateTime = gNow()
            slNowDate = Format(slDateTime, "ddddd")
            slNowTime = Format(slDateTime, "ttttt")
            llNowDate = gDateValue(slNowDate)
            llNowTime = 10 * (gTimeToLong(slNowTime, False) + tgSOE.lChgInterval)
            llDate = 0
            If llStartDate > llNowDate Then
                ilRet = gGetTypeOfRecs_SHE_ScheduleHeaderByDate(slDateStartRange, "gCheckConflicts", tmCurrSHE())
            Else
                ilRet = gGetTypeOfRecs_SHE_ScheduleHeaderByDate(slNowDate, "gCheckConflicts", tmCurrSHE())
            End If
            'For llDate = llStartDate To llEndDate Step 1
            For ilSHE = 0 To UBound(tmCurrSHE) - 1 Step 1
                LSet tmSHE = tmCurrSHE(ilSHE)
                slDate = tmSHE.sAirDate
                If gDateValue(slDate) > llDate Then
                    llDate = gDateValue(slDate)
                End If
                ilUpper = UBound(tmCurrLibDHE)
                tmCurrLibDHE(ilUpper).lCode = -1
                tmCurrLibDHE(ilUpper).sType = "S"
                tmCurrLibDHE(ilUpper).lDneCode = tmSHE.lCode
                tmCurrLibDHE(ilUpper).lDseCode = 0
                tmCurrLibDHE(ilUpper).sStartTime = "00:00:00.0"
                tmCurrLibDHE(ilUpper).lLength = 86400
                tmCurrLibDHE(ilUpper).sHours = String(24, "Y")
                tmCurrLibDHE(ilUpper).sStartDate = slDate
                tmCurrLibDHE(ilUpper).sEndDate = slDate
                tmCurrLibDHE(ilUpper).sDays = String(7, "N")
                Select Case Weekday(slDate)
                    Case vbMonday
                        Mid(tmCurrLibDHE(ilUpper).sDays, 1, 1) = "Y"
                    Case vbTuesday
                        Mid(tmCurrLibDHE(ilUpper).sDays, 2, 1) = "Y"
                    Case vbWednesday
                        Mid(tmCurrLibDHE(ilUpper).sDays, 3, 1) = "Y"
                    Case vbThursday
                        Mid(tmCurrLibDHE(ilUpper).sDays, 4, 1) = "Y"
                    Case vbFriday
                        Mid(tmCurrLibDHE(ilUpper).sDays, 5, 1) = "Y"
                    Case vbSaturday
                        Mid(tmCurrLibDHE(ilUpper).sDays, 6, 1) = "Y"
                    Case vbSunday
                        Mid(tmCurrLibDHE(ilUpper).sDays, 7, 1) = "Y"
                End Select
                ReDim Preserve tmCurrLibDHE(0 To ilUpper + 1) As DHE
            Next ilSHE
            If (llDate <> 0) And (slType <> "T") Then
                llLibStartDate = llDate + 1
            End If
            If llLibStartDate <= llEndDate Then
                For ilDHE = 0 To UBound(tmCurrLibDHE) - 1 Step 1
                    slStartDate = tmCurrLibDHE(ilDHE).sStartDate
                    slEndDate = tmCurrLibDHE(ilDHE).sEndDate
                    If (llDheCode <> tmCurrLibDHE(ilDHE).lCode) And (llCurrentDHE <> tmCurrLibDHE(ilDHE).lCode) Then
                        ''If (gDateValue(slEndDate) >= llStartDate) And (gDateValue(slStartDate) <= llEndDate) Then
                        'If ((gDateValue(slEndDate) >= llLibStartDate) And (gDateValue(slStartDate) <= llEndDate) And (tmCurrLibDHE(ilDHE).sType <> "S")) Or (tmCurrLibDHE(ilDHE).sType = "S") Then
                        If ((gDateValue(slEndDate) >= llLibStartDate) And (gDateValue(slStartDate) <= llDateEndRange) And (tmCurrLibDHE(ilDHE).sType <> "S")) Or (tmCurrLibDHE(ilDHE).sType = "S") Then
                            ilCheck = True
                            If ilPass = 1 Then
                                If gDateValue(slEndDate) <> llDateStartRange Then
                                    ilCheck = False
                                End If
                            ElseIf ilPass = 2 Then
                                If gDateValue(slStartDate) <> llDateEndRange Then
                                    ilCheck = False
                                End If
                            End If
                            If ilCheck Then
                                If tmCurrLibDHE(ilDHE).sType <> "S" Then
                                    ilRet = gGetRecs_DEE_DayEvent(smCurrLibDEEStamp, tmCurrLibDHE(ilDHE).lCode, "EngrLibDef-mPopulate", tmCurrLibDEE())
                                    If tmCurrLibDHE(ilDHE).sType = "T" Then
                                        'Adjust Hours and Days and set Bus
                                        LSet tlTSE = tmCurrTempTSE(ilDHE)
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
                                        Next ilDEE
                                    Else
                                        For ilDEE = 0 To UBound(tmCurrLibDEE) - 1 Step 1
                                            tmCurrLibDEE(ilDEE).sIgnoreConflicts = tmCurrLibDHE(ilDHE).sIgnoreConflicts
                                        Next ilDEE
                                    End If
                                Else
                                    'Note: tmCurrLibDHE(ilDHE).lDneCode contains SHECode
                                    ilRet = gGetRecs_SEE_ScheduleEventsAPI(hlSEE, smCurrSEEStamp, -1, tmCurrLibDHE(ilDHE).lDneCode, "EngrSchd-Get Events", tmCurrSEE())
                                    ReDim tmCurrLibDEE(0 To 0) As DEE
                                    smCurrLibDEEStamp = ""
                                    If ilRet Then
                                        For ilLib = 0 To UBound(tmCurrSEE) - 1 Step 1
                                            If (tmCurrSEE(ilLib).sAction <> "D") And (tmCurrSEE(ilLib).sAction <> "R") Then
                                                If (gDateValue(tmCurrLibDHE(ilDHE).sStartDate) > llNowDate) Or ((gDateValue(tmCurrLibDHE(ilDHE).sStartDate) = llNowDate) And (tmCurrSEE(ilLib).lTime > llNowTime)) Then
                                                    'ilRet = gGetRec_DEE_DayEvent(tmCurrSEE(ilLib).lDeeCode, "EngrSchd-Get Library Events", tlDEE)
                                                    'ilRet = gGetRec_DHE_DayHeaderInfo(tlDEE.lDheCode, "EngrSchd-Get Library Events", tlDHE)
                                                    'If ((llDheCode <> tlDEE.lDheCode) And (llCurrentDHE <> tlDEE.lDheCode)) And ((llDheCode <> tlDHE.lOrigDheCode) And (llCurrentDHE <> tlDHE.lOrigDheCode)) Then
                                                    If ((llDheCode <> tmCurrSEE(ilLib).lDheCode) And (llCurrentDHE <> tmCurrSEE(ilLib).lDheCode)) And ((llDheCode <> tmCurrSEE(ilLib).lOrigDHECode) And (llCurrentDHE <> tmCurrSEE(ilLib).lOrigDHECode)) Then
                                                        ilUpper = UBound(tmCurrLibDEE)
                                                        tmCurrLibDEE(ilUpper).lDheCode = tmCurrSEE(ilLib).lDheCode  'tlDEE.lDheCode
                                                        tmCurrLibDEE(ilUpper).lCode = tmCurrSEE(ilLib).lDeeCode
                                                        tmCurrLibDEE(ilUpper).l1CteCode = tmCurrSEE(ilLib).lCode
                                                        tmCurrLibDEE(ilUpper).lTime = tmCurrSEE(ilLib).lTime
                                                        tmCurrLibDEE(ilUpper).lDuration = tmCurrSEE(ilLib).lDuration
                                                        tmCurrLibDEE(ilUpper).sHours = String(24, "N")
                                                        ilHour = tmCurrLibDEE(ilUpper).lTime \ 36000 + 1
                                                        Mid(tmCurrLibDEE(ilUpper).sHours, ilHour, 1) = "Y"
                                                        tmCurrLibDEE(ilUpper).lTime = tmCurrLibDEE(ilUpper).lTime Mod 36000
                                                        tmCurrLibDEE(ilUpper).sDays = tmCurrLibDHE(ilDHE).sDays
                                                        tmCurrLibDEE(ilUpper).iFneCode = tmCurrSEE(ilLib).iBdeCode
                                                        tmCurrLibDEE(ilUpper).iEteCode = tmCurrSEE(ilLib).iEteCode
                                                        tmCurrLibDEE(ilUpper).iAudioAseCode = tmCurrSEE(ilLib).iAudioAseCode
                                                        tmCurrLibDEE(ilUpper).iProtAneCode = tmCurrSEE(ilLib).iProtAneCode
                                                        tmCurrLibDEE(ilUpper).iBkupAneCode = tmCurrSEE(ilLib).iBkupAneCode
                                                        tmCurrLibDEE(ilUpper).sIgnoreConflicts = tmCurrSEE(ilLib).sIgnoreConflicts
                                                        ReDim Preserve tmCurrLibDEE(0 To ilUpper + 1) As DEE
                                                    End If
                                                End If
                                            End If
                                        Next ilLib
                                    End If
                                End If
                                ReDim tmConflictLib(1 To 1) As CONFLICTTEST
                                For ilLib = 0 To UBound(tmCurrLibDEE) - 1 Step 1
                                    llOffsetStartTime = tmCurrLibDEE(ilLib).lTime
                                    llOffsetEndTime = llOffsetStartTime + tmCurrLibDEE(ilLib).lDuration ' - 1
                                    If llOffsetEndTime < llOffsetStartTime Then
                                        llOffsetEndTime = llOffsetStartTime
                                    End If
                                    slPriAudio2 = ""
                                    slProtAudio2 = ""
                                    slBkupAudio2 = ""
                                    If (tmCurrLibDEE(ilLib).sIgnoreConflicts <> "A") And (tmCurrLibDEE(ilLib).sIgnoreConflicts <> "I") Then
                                        If tmCurrLibDEE(ilLib).iAudioAseCode > 0 Then
                                            'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                                            '    If tmCurrLibDEE(ilLib).iAudioAseCode = tgCurrASE(ilASE).iCode Then
                                                ilASE = gBinarySearchASE(tmCurrLibDEE(ilLib).iAudioAseCode, tgCurrASE())
                                                If ilASE <> -1 Then
                                                    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                                    '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                                                        ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                                                        If ilANE <> -1 Then
                                                            slPriAudio2 = Trim$(tgCurrANE(ilANE).sName)
                                                        End If
                                                    'Next ilANE
                                            '        If slPriAudio2 <> "" Then
                                            '            Exit For
                                            '        End If
                                                End If
                                            'Next ilASE
                                        End If
                                        If (tmCurrLibDEE(ilLib).iProtAneCode > 0) Then
                                            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                            '    If tmCurrLibDEE(ilLib).iProtAneCode = tgCurrANE(ilANE).iCode Then
                                                ilANE = gBinarySearchANE(tmCurrLibDEE(ilLib).iProtAneCode, tgCurrANE())
                                                If ilANE <> -1 Then
                                                    slProtAudio2 = Trim$(tgCurrANE(ilANE).sName)
                                            '        Exit For
                                                End If
                                            'Next ilANE
                                        End If
                                        If (tmCurrLibDEE(ilLib).iBkupAneCode > 0) Then
                                            'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                            '    If tmCurrLibDEE(ilLib).iBkupAneCode = tgCurrANE(ilANE).iCode Then
                                                ilANE = gBinarySearchANE(tmCurrLibDEE(ilLib).iBkupAneCode, tgCurrANE())
                                                If ilANE <> -1 Then
                                                    slBkupAudio2 = Trim$(tgCurrANE(ilANE).sName)
                                            '        Exit For
                                                End If
                                            'Next ilANE
                                        End If
                                    End If
                                    For ilHour = 1 To 24 Step 1
                                        If Mid$(tmCurrLibDEE(ilLib).sHours, ilHour, 1) = "Y" Then
                                            llStartTime = 36000 * (ilHour - 1) + llOffsetStartTime
                                            llEndTime = 36000 * (ilHour - 1) + llOffsetEndTime
                                            If tmCurrLibDHE(ilDHE).sType = "T" Then
                                                llStartTime = llStartTime + tmCurrLibDHE(ilDHE).lLength
                                                llEndTime = llEndTime + tmCurrLibDHE(ilDHE).lLength
                                            End If
                                            If slType = "S" Then
                                                If ilPass = 0 Then
                                                    If (gDateValue(slEndDate) = llDateStartRange) Then
                                                        slSpecDay = String(7, "N")
                                                        ilDay = Weekday(slEndDate, vbMonday)
                                                        Mid(slSpecDay, ilDay, 1) = "Y"
                                                        If Mid(tmCurrLibDEE(ilDEE).sDays, ilDay, 1) = "Y" Then
                                                            mCreateBusRecs 2, CLng(ilLib), "B", tmCurrLibDEE(ilDEE).sIgnoreConflicts, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                            mCreateAudioRecs 2, CLng(ilLib), "1", slPriAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                            mCreateAudioRecs 2, CLng(ilLib), "2", slProtAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                            mCreateAudioRecs 2, CLng(ilLib), "3", slBkupAudio2, llStartTime, llEndTime, tmCurrLibDEE(ilLib).sDays, tmConflictLib()
                                                        End If
                                                    ElseIf (gDateValue(slStartDate) = llDateEndRange) And (llDateEndRange <> gDateValue("12/31/2069")) Then
                                                        slSpecDay = String(7, "N")
                                                        ilDay = Weekday(slStartDate, vbMonday)
                                                        Mid(slSpecDay, ilDay, 1) = "Y"
                                                        If Mid(tmCurrLibDEE(ilDEE).sDays, ilDay, 1) = "Y" Then
                                                            mCreateBusRecs 1, CLng(ilLib), "B", tmCurrLibDEE(ilDEE).sIgnoreConflicts, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                            mCreateAudioRecs 1, CLng(ilLib), "1", slPriAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                            mCreateAudioRecs 1, CLng(ilLib), "2", slProtAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                            mCreateAudioRecs 1, CLng(ilLib), "3", slBkupAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                        End If
                                                    Else
                                                        slSpecDay = tmCurrLibDEE(ilLib).sDays
                                                        ilDay = Weekday(slInStartDate, vbMonday)
                                                        Mid(slSpecDay, ilDay, 1) = "N"
                                                        mCreateBusRecs 0, CLng(ilLib), "B", tmCurrLibDEE(ilDEE).sIgnoreConflicts, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                        mCreateAudioRecs 0, CLng(ilLib), "1", slPriAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                        mCreateAudioRecs 0, CLng(ilLib), "2", slProtAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                        mCreateAudioRecs 0, CLng(ilLib), "3", slBkupAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                    End If
                                                Else
                                                    mCreateBusRecs 0, CLng(ilLib), "B", tmCurrLibDEE(ilDEE).sIgnoreConflicts, llStartTime, llEndTime, tmCurrLibDEE(ilLib).sDays, tmConflictLib()
                                                    mCreateAudioRecs 0, CLng(ilLib), "1", slPriAudio2, llStartTime, llEndTime, tmCurrLibDEE(ilLib).sDays, tmConflictLib()
                                                    mCreateAudioRecs 0, CLng(ilLib), "2", slProtAudio2, llStartTime, llEndTime, tmCurrLibDEE(ilLib).sDays, tmConflictLib()
                                                    mCreateAudioRecs 0, CLng(ilLib), "3", slBkupAudio2, llStartTime, llEndTime, tmCurrLibDEE(ilLib).sDays, tmConflictLib()
                                                End If
                                            Else
                                                If (ilPass = 0) And (gDateValue(slEndDate) = llDateStartRange) Then
                                                    slSpecDay = String(7, "N")
                                                    ilDay = Weekday(slEndDate, vbMonday)
                                                    Mid(slSpecDay, ilDay, 1) = "Y"
                                                    If Mid(tmCurrLibDEE(ilDEE).sDays, ilDay, 1) = "Y" Then
                                                        mCreateBusRecs 2, CLng(ilLib), "B", tmCurrLibDEE(ilDEE).sIgnoreConflicts, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                        mCreateAudioRecs 2, CLng(ilLib), "1", slPriAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                        mCreateAudioRecs 2, CLng(ilLib), "2", slProtAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                        mCreateAudioRecs 2, CLng(ilLib), "3", slBkupAudio2, llStartTime, llEndTime, tmCurrLibDEE(ilLib).sDays, tmConflictLib()
                                                    End If
                                                ElseIf (ilPass = 0) And (gDateValue(slStartDate) = llDateEndRange) And (llDateEndRange <> gDateValue("12/31/2069")) Then
                                                    slSpecDay = String(7, "N")
                                                    ilDay = Weekday(slStartDate, vbMonday)
                                                    Mid(slSpecDay, ilDay, 1) = "Y"
                                                    If Mid(tmCurrLibDEE(ilDEE).sDays, ilDay, 1) = "Y" Then
                                                        mCreateBusRecs 1, CLng(ilLib), "B", tmCurrLibDEE(ilDEE).sIgnoreConflicts, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                        mCreateAudioRecs 1, CLng(ilLib), "1", slPriAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                        mCreateAudioRecs 1, CLng(ilLib), "2", slProtAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                        mCreateAudioRecs 1, CLng(ilLib), "3", slBkupAudio2, llStartTime, llEndTime, slSpecDay, tmConflictLib()
                                                    End If
                                                Else
                                                    mCreateBusRecs 0, CLng(ilLib), "B", tmCurrLibDEE(ilDEE).sIgnoreConflicts, llStartTime, llEndTime, tmCurrLibDEE(ilLib).sDays, tmConflictLib()
                                                    mCreateAudioRecs 0, CLng(ilLib), "1", slPriAudio2, llStartTime, llEndTime, tmCurrLibDEE(ilLib).sDays, tmConflictLib()
                                                    mCreateAudioRecs 0, CLng(ilLib), "2", slProtAudio2, llStartTime, llEndTime, tmCurrLibDEE(ilLib).sDays, tmConflictLib()
                                                    mCreateAudioRecs 0, CLng(ilLib), "3", slBkupAudio2, llStartTime, llEndTime, tmCurrLibDEE(ilLib).sDays, tmConflictLib()
                                                End If
                                            End If
                                        End If
                                    Next ilHour
                                Next ilLib
                                'Test at Events
                                For llRow = 1 To UBound(tmConflictTest) - 1 Step 1
                                    llEvent = tmConflictTest(llRow).lRow
                                    If (Trim$(grdEvents.TextMatrix(llEvent, ilCols(1))) <> "") Then 'And (grdEvents.TextMatrix(llEvent, ilCols(0)) = "0") Then
                                        slEvtType1 = Trim$(grdEvents.TextMatrix(llEvent, ilCols(1)))
                                        slBuses = Trim$(grdEvents.TextMatrix(llEvent, ilCols(6)))
                                        gParseCDFields slBuses, False, slBusNames()
                                        slPriAudio = Trim$(grdEvents.TextMatrix(llEvent, ilCols(7)))
                                        slProtAudio = Trim$(grdEvents.TextMatrix(llEvent, ilCols(8)))
                                        slBkupAudio = Trim$(grdEvents.TextMatrix(llEvent, ilCols(9)))
                                        slPriItemID1 = Trim$(grdEvents.TextMatrix(llEvent, ilCols(10)))
                                        slProtItemID1 = Trim$(grdEvents.TextMatrix(llEvent, ilCols(11)))
                                        slBkupItemID1 = Trim$(grdEvents.TextMatrix(llEvent, ilCols(12)))
                                        For llLibEvent = 1 To UBound(tmConflictLib) - 1 Step 1
                                            ilLib = tmConflictLib(llLibEvent).lRow
                                            ilConflictIndex = UBound(tlConflictList)
                                            tlConflictList(ilConflictIndex).sType = tmCurrLibDHE(ilDHE).sType
                                            tlConflictList(ilConflictIndex).sStartDate = tmCurrLibDHE(ilDHE).sStartDate
                                            tlConflictList(ilConflictIndex).sEndDate = tmCurrLibDHE(ilDHE).sEndDate
                                            tlConflictList(ilConflictIndex).lIndex = -1
                                            tlConflictList(ilConflictIndex).iNextIndex = -1
                                            If tmCurrLibDHE(ilDHE).sType <> "S" Then
                                                tlConflictList(ilConflictIndex).lSheCode = 0
                                                tlConflictList(ilConflictIndex).lSeeCode = 0
                                                tlConflictList(ilConflictIndex).lDheCode = tmCurrLibDHE(ilDHE).lCode
                                                tlConflictList(ilConflictIndex).lDseCode = tmCurrLibDHE(ilDHE).lDseCode
                                                tlConflictList(ilConflictIndex).lDeeCode = tmCurrLibDEE(ilLib).lCode
                                            Else
                                                tlConflictList(ilConflictIndex).lSheCode = tmCurrLibDHE(ilDHE).lDneCode
                                                tlConflictList(ilConflictIndex).lDheCode = 0
                                                tlConflictList(ilConflictIndex).lDseCode = 0
                                                tlConflictList(ilConflictIndex).lDeeCode = 0
                                                tlConflictList(ilConflictIndex).lSeeCode = tmCurrLibDEE(ilLib).l1CteCode
                                            End If
                                            smCurrLibEBEStamp = ""
                                            Erase tmCurrLibEBE
                                            If (tmCurrLibDHE(ilDHE).sType <> "S") And (tmCurrLibDHE(ilDHE).sType <> "T") Then
                                                ilRet = gGetRecs_EBE_EventBusSel(smCurrLibEBEStamp, tmCurrLibDEE(ilLib).lCode, "Bus Definition-mDEEMoveRecToCtrls", tmCurrLibEBE())
                                            Else
                                                ReDim tmCurrLibEBE(0 To 1) As EBE
                                                tmCurrLibEBE(0).iBdeCode = tmCurrLibDEE(ilLib).iFneCode
                                            End If
                                            ilError = False
                                            slEventDays = tmConflictTest(llRow).sDays
                                            slLibDays = tmConflictLib(llLibEvent).sDays
                                            For ilDay = 1 To 7 Step 1
                                                If (Mid$(slEventDays, ilDay, 1) = "Y") And (Mid$(slLibDays, ilDay, 1) = "Y") Then
                                                    ilStartConflictIndex = Val(grdEvents.TextMatrix(llEvent, ilCols(0)))
                                                    'Include start and end points
                                                    If (tmConflictLib(llLibEvent).sType = "B") And (tmConflictTest(llRow).sType = "B") And ((tmConflictLib(llLibEvent).lEventEndTime > tmConflictTest(llRow).lEventStartTime) And (tmConflictLib(llLibEvent).lEventStartTime < tmConflictTest(llRow).lEventEndTime) Or (tmConflictLib(llLibEvent).lEventStartTime = tmConflictTest(llRow).lEventStartTime)) Then
                                                       'Check Bus
                                                        If (grdEvents.ColWidth(ilCols(6)) > 0) Or (slBuses <> "") Then
                                                            ilFound = False
                                                            For ilEBE = 0 To UBound(tmCurrLibEBE) - 1 Step 1
                                                                'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                                                                '    If tmCurrLibEBE(ilEBE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                                                                    ilBDE = gBinarySearchBDE(tmCurrLibEBE(ilEBE).iBdeCode, tgCurrBDE())
                                                                    If ilBDE <> -1 Then
                                                                        For ilBus = LBound(slBusNames) To UBound(slBusNames) Step 1
                                                                            If StrComp(slBusNames(ilBus), Trim$(tgCurrBDE(ilBDE).sName), vbTextCompare) = 0 Then
                                                                                gCheckConflicts = True
                                                                                grdEvents.Row = llEvent
                                                                                grdEvents.Col = ilCols(6)
                                                                                grdEvents.CellForeColor = vbRed
                                                                                ilFound = True
                                                                                If Not ilError Then
                                                                                    grdEvents.TextMatrix(llEvent, ilCols(0)) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tlConflictList())))
                                                                                End If
                                                                                ilError = True
                                                                                Exit For
                                                                            End If
                                                                        Next ilBus
                                                                '        Exit For
                                                                    End If
                                                                'Next ilBDE
                                                                If ilFound Then
                                                                    Exit For
                                                                End If
                                                            Next ilEBE
                                                        End If
                                                    End If
                                                    slEvtType2 = ""
                                                    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                                                        If tgCurrETE(ilETE).iCode = tmCurrLibDEE(ilLib).iEteCode Then
                                                            slEvtType2 = tgCurrETE(ilETE).sName
                                                            Exit For
                                                        End If
                                                    Next ilETE
                                                    slPriAudio2 = ""
                                                    If tmCurrLibDEE(ilLib).iAudioAseCode > 0 Then
                                                        'For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
                                                        '    If tmCurrLibDEE(ilLib).iAudioAseCode = tgCurrASE(ilASE).iCode Then
                                                            ilASE = gBinarySearchASE(tmCurrLibDEE(ilLib).iAudioAseCode, tgCurrASE())
                                                            If ilASE <> -1 Then
                                                                'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                                                '    If tgCurrASE(ilASE).iPriAneCode = tgCurrANE(ilANE).iCode Then
                                                                    ilANE = gBinarySearchANE(tgCurrASE(ilASE).iPriAneCode, tgCurrANE())
                                                                    If ilANE <> -1 Then
                                                                        slPriAudio2 = Trim$(tgCurrANE(ilANE).sName)
                                                                    End If
                                                                'Next ilANE
                                                        '        If slPriAudio2 <> "" Then
                                                        '            Exit For
                                                        '        End If
                                                            End If
                                                        'Next ilASE
                                                    End If
                                                    slProtAudio2 = ""
                                                    If (tmCurrLibDEE(ilLib).iProtAneCode > 0) Then
                                                        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                                        '    If tmCurrLibDEE(ilLib).iProtAneCode = tgCurrANE(ilANE).iCode Then
                                                            ilANE = gBinarySearchANE(tmCurrLibDEE(ilLib).iProtAneCode, tgCurrANE())
                                                            If ilANE <> -1 Then
                                                                slProtAudio2 = Trim$(tgCurrANE(ilANE).sName)
                                                        '        Exit For
                                                            End If
                                                        'Next ilANE
                                                    End If
                                                    slBkupAudio2 = ""
                                                    If (tmCurrLibDEE(ilLib).iBkupAneCode > 0) Then
                                                        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                                                        '    If tmCurrLibDEE(ilLib).iBkupAneCode = tgCurrANE(ilANE).iCode Then
                                                            ilANE = gBinarySearchANE(tmCurrLibDEE(ilLib).iBkupAneCode, tgCurrANE())
                                                            If ilANE <> -1 Then
                                                                slBkupAudio2 = Trim$(tgCurrANE(ilANE).sName)
                                                        '        Exit For
                                                            End If
                                                        'Next ilANE
                                                    End If
                                                    slPriItemID2 = Trim$(tmCurrLibDEE(ilLib).sAudioItemID)
                                                    slProtItemID2 = Trim$(tmCurrLibDEE(ilLib).sProtItemID)
                                                    slBkupItemID2 = Trim$(tmCurrLibDEE(ilLib).sAudioItemID)
                                                    For ilEBE = 0 To UBound(tmCurrLibEBE) - 1 Step 1
                                                        'For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
                                                        '    If tmCurrLibEBE(ilEBE).iBdeCode = tgCurrBDE(ilBDE).iCode Then
                                                            ilBDE = gBinarySearchBDE(tmCurrLibEBE(ilEBE).iBdeCode, tgCurrBDE())
                                                            If ilBDE <> -1 Then
                                                                For ilBus = LBound(slBusNames) To UBound(slBusNames) Step 1
                                                                    If (tmConflictTest(llRow).sType = "1") And (tmConflictLib(llLibEvent).sType = "1") Then
                                                                        If gAudioConflicts(slEvtType1, slEvtType2, slPriAudio, slPriAudio2, slPriItemID1, slPriItemID2, tmConflictTest(llRow).lEventStartTime, tmConflictTest(llRow).lEventEndTime, tmConflictLib(llLibEvent).lEventStartTime, tmConflictLib(llLibEvent).lEventEndTime, False, slBusNames(ilBus), Trim$(tgCurrBDE(ilBDE).sName)) Then
                                                                            gCheckConflicts = True
                                                                            grdEvents.Row = llEvent
                                                                            grdEvents.Col = ilCols(Val(tmConflictTest(llRow).sType) + 6)
                                                                            grdEvents.CellForeColor = vbRed
                                                                            If Not ilError Then
                                                                                grdEvents.TextMatrix(llEvent, ilCols(0)) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tlConflictList())))
                                                                            End If
                                                                            ilError = True
                                                                        End If
                                                                    End If
                                                                    If (tmConflictTest(llRow).sType = "1") And (tmConflictLib(llLibEvent).sType = "2") Then
                                                                        If gAudioConflicts(slEvtType1, slEvtType2, slPriAudio, slProtAudio2, slPriItemID1, slProtItemID2, tmConflictTest(llRow).lEventStartTime, tmConflictTest(llRow).lEventEndTime, tmConflictLib(llLibEvent).lEventStartTime, tmConflictLib(llLibEvent).lEventEndTime, False, slBusNames(ilBus), Trim$(tgCurrBDE(ilBDE).sName)) Then
                                                                            gCheckConflicts = True
                                                                            grdEvents.Row = llEvent
                                                                            grdEvents.Col = ilCols(Val(tmConflictTest(llRow).sType) + 6)
                                                                            grdEvents.CellForeColor = vbRed
                                                                            If Not ilError Then
                                                                                grdEvents.TextMatrix(llEvent, ilCols(0)) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tlConflictList())))
                                                                            End If
                                                                            ilError = True
                                                                        End If
                                                                    End If
                                                                    If (tmConflictTest(llRow).sType = "1") And (tmConflictLib(llLibEvent).sType = "3") Then
                                                                        If gAudioConflicts(slEvtType1, slEvtType2, slPriAudio, slBkupAudio2, slPriItemID1, slBkupItemID2, tmConflictTest(llRow).lEventStartTime, tmConflictTest(llRow).lEventEndTime, tmConflictLib(llLibEvent).lEventStartTime, tmConflictLib(llLibEvent).lEventEndTime, False, slBusNames(ilBus), Trim$(tgCurrBDE(ilBDE).sName)) Then
                                                                            gCheckConflicts = True
                                                                            grdEvents.Row = llEvent
                                                                            grdEvents.Col = ilCols(Val(tmConflictTest(llRow).sType) + 6)
                                                                            grdEvents.CellForeColor = vbRed
                                                                            If Not ilError Then
                                                                                grdEvents.TextMatrix(llEvent, ilCols(0)) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tlConflictList())))
                                                                            End If
                                                                            ilError = True
                                                                        End If
                                                                    End If
                                                                    If (tmConflictTest(llRow).sType = "2") And (tmConflictLib(llLibEvent).sType = "1") Then
                                                                        If gAudioConflicts(slEvtType1, slEvtType2, slProtAudio, slPriAudio2, slProtItemID1, slPriItemID2, tmConflictTest(llRow).lEventStartTime, tmConflictTest(llRow).lEventEndTime, tmConflictLib(llLibEvent).lEventStartTime, tmConflictLib(llLibEvent).lEventEndTime, False, slBusNames(ilBus), Trim$(tgCurrBDE(ilBDE).sName)) Then
                                                                            gCheckConflicts = True
                                                                            grdEvents.Row = llEvent
                                                                            grdEvents.Col = ilCols(Val(tmConflictTest(llRow).sType) + 6)
                                                                            grdEvents.CellForeColor = vbRed
                                                                            If Not ilError Then
                                                                                grdEvents.TextMatrix(llEvent, ilCols(0)) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tlConflictList())))
                                                                            End If
                                                                            ilError = True
                                                                        End If
                                                                    End If
                                                                    If (tmConflictTest(llRow).sType = "2") And (tmConflictLib(llLibEvent).sType = "2") Then
                                                                        If gAudioConflicts(slEvtType1, slEvtType2, slProtAudio, slProtAudio2, slProtItemID1, slProtItemID2, tmConflictTest(llRow).lEventStartTime, tmConflictTest(llRow).lEventEndTime, tmConflictLib(llLibEvent).lEventStartTime, tmConflictLib(llLibEvent).lEventEndTime, False, slBusNames(ilBus), Trim$(tgCurrBDE(ilBDE).sName)) Then
                                                                            gCheckConflicts = True
                                                                            grdEvents.Row = llEvent
                                                                            grdEvents.Col = ilCols(Val(tmConflictTest(llRow).sType) + 6)
                                                                            grdEvents.CellForeColor = vbRed
                                                                            If Not ilError Then
                                                                                grdEvents.TextMatrix(llEvent, ilCols(0)) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tlConflictList())))
                                                                            End If
                                                                            ilError = True
                                                                        End If
                                                                    End If
                                                                    If (tmConflictTest(llRow).sType = "2") And (tmConflictLib(llLibEvent).sType = "3") Then
                                                                        If gAudioConflicts(slEvtType1, slEvtType2, slProtAudio, slBkupAudio2, slProtItemID1, slBkupItemID2, tmConflictTest(llRow).lEventStartTime, tmConflictTest(llRow).lEventEndTime, tmConflictLib(llLibEvent).lEventStartTime, tmConflictLib(llLibEvent).lEventEndTime, False, slBusNames(ilBus), Trim$(tgCurrBDE(ilBDE).sName)) Then
                                                                            gCheckConflicts = True
                                                                            grdEvents.Row = llEvent
                                                                            grdEvents.Col = ilCols(Val(tmConflictTest(llRow).sType) + 6)
                                                                            grdEvents.CellForeColor = vbRed
                                                                            If Not ilError Then
                                                                                grdEvents.TextMatrix(llEvent, ilCols(0)) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tlConflictList())))
                                                                            End If
                                                                            ilError = True
                                                                        End If
                                                                    End If
                                                                    If (tmConflictTest(llRow).sType = "3") And (tmConflictLib(llLibEvent).sType = "1") Then
                                                                        If gAudioConflicts(slEvtType1, slEvtType2, slBkupAudio, slPriAudio2, slBkupItemID1, slPriItemID2, tmConflictTest(llRow).lEventStartTime, tmConflictTest(llRow).lEventEndTime, tmConflictLib(llLibEvent).lEventStartTime, tmConflictLib(llLibEvent).lEventEndTime, False, slBusNames(ilBus), Trim$(tgCurrBDE(ilBDE).sName)) Then
                                                                            gCheckConflicts = True
                                                                            grdEvents.Row = llEvent
                                                                            grdEvents.Col = ilCols(Val(tmConflictTest(llRow).sType) + 6)
                                                                            grdEvents.CellForeColor = vbRed
                                                                            If Not ilError Then
                                                                                grdEvents.TextMatrix(llEvent, ilCols(0)) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tlConflictList())))
                                                                            End If
                                                                            ilError = True
                                                                        End If
                                                                    End If
                                                                    If (tmConflictTest(llRow).sType = "3") And (tmConflictLib(llLibEvent).sType = "2") Then
                                                                        If gAudioConflicts(slEvtType1, slEvtType2, slBkupAudio, slProtAudio2, slBkupItemID1, slProtItemID2, tmConflictTest(llRow).lEventStartTime, tmConflictTest(llRow).lEventEndTime, tmConflictLib(llLibEvent).lEventStartTime, tmConflictLib(llLibEvent).lEventEndTime, False, slBusNames(ilBus), Trim$(tgCurrBDE(ilBDE).sName)) Then
                                                                            gCheckConflicts = True
                                                                            grdEvents.Row = llEvent
                                                                            grdEvents.Col = ilCols(Val(tmConflictTest(llRow).sType) + 6)
                                                                            grdEvents.CellForeColor = vbRed
                                                                            If Not ilError Then
                                                                                grdEvents.TextMatrix(llEvent, ilCols(0)) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tlConflictList())))
                                                                            End If
                                                                            ilError = True
                                                                        End If
                                                                    End If
                                                                    If (tmConflictTest(llRow).sType = "3") And (tmConflictLib(llLibEvent).sType = "3") Then
                                                                        If gAudioConflicts(slEvtType1, slEvtType2, slBkupAudio, slBkupAudio2, slBkupItemID1, slBkupItemID2, tmConflictTest(llRow).lEventStartTime, tmConflictTest(llRow).lEventEndTime, tmConflictLib(llLibEvent).lEventStartTime, tmConflictLib(llLibEvent).lEventEndTime, False, slBusNames(ilBus), Trim$(tgCurrBDE(ilBDE).sName)) Then
                                                                            gCheckConflicts = True
                                                                            grdEvents.Row = llEvent
                                                                            grdEvents.Col = ilCols(Val(tmConflictTest(llRow).sType) + 6)
                                                                            grdEvents.CellForeColor = vbRed
                                                                            If Not ilError Then
                                                                                grdEvents.TextMatrix(llEvent, ilCols(0)) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tlConflictList())))
                                                                            End If
                                                                            ilError = True
                                                                        End If
                                                                    End If
                                                                Next ilBus
                                                            End If
                                                        'Next ilBDE
                                                    Next ilEBE
                                                End If
                                            Next ilDay
                                        Next llLibEvent
                                    End If
                                Next llRow
                                ilError = True
                                For llEvent = grdEvents.FixedRows To grdEvents.Rows - 1 Step 1
                                    If (Trim$(grdEvents.TextMatrix(llEvent, ilCols(1))) <> "") And (grdEvents.TextMatrix(llEvent, ilCols(0)) = "0") Then
                                        ilError = False
                                        Exit For
                                    End If
                                Next llEvent
                                If ilError Then
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next ilDHE
            End If
        End If
    Next ilPass
    Erase tmConflictLib
    Erase tmConflictTest
End Function
Public Function gCreateHourStr(slHourName As String) As String
    Dim slStr As String
    Dim ilHours As Integer
    Dim ilPos As Integer
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim ilSet As Integer
    
    slStr = Trim$(slHourName)
    gParseCDFields slStr, False, smHours()
    'slStr = "NNNNNNNNNNNNNNNNNNNNNNNN"
    slStr = String(24, "N")
    For ilHours = LBound(smHours) To UBound(smHours) Step 1
        ilPos = InStr(1, smHours(ilHours), "-", vbTextCompare)
        If ilPos <= 0 Then
            ilStart = Val(smHours(ilHours))
            ilEnd = ilStart
        Else
            ilStart = Val(Left$(smHours(ilHours), ilPos - 1))
            ilEnd = Val(Mid$(smHours(ilHours), ilPos + 1))
        End If
        If (ilStart < 0) Or (ilStart > 23) Or (ilEnd < 0) Or (ilEnd > 23) Or (ilEnd < ilStart) Then
            slStr = ""  'String(24, "N")
            Exit For
        Else
            For ilSet = ilStart To ilEnd Step 1
                Mid$(slStr, ilSet + 1, 1) = "Y"
            Next ilSet
        End If
    Next ilHours
    gCreateHourStr = slStr
End Function


Public Function gCreateDayStr(slDayName As String) As String
    Dim slStr As String
    Dim slDays As String
    Dim ilPos As Integer
    Dim Days As Integer
    Dim slStart As String
    Dim slEnd As String
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim ilSet As Integer
    Dim ilDays As Integer
    
    slStr = Trim$(slDayName)
    slDays = String(7, "N")
    gParseCDFields slStr, False, smDays()
    For ilDays = LBound(smDays) To UBound(smDays) Step 1
        ilPos = InStr(1, smDays(ilDays), "-", vbTextCompare)
        If ilPos <= 0 Then
            slStart = smDays(ilDays)
            slEnd = slStart
        Else
            slStart = Left$(smDays(ilDays), ilPos - 1)
            slEnd = Mid$(smDays(ilDays), ilPos + 1)
        End If
        slStr = UCase(slStart)
        'Could use switch to get the index
        'ilStart = Switch(slStr = "M", 1, slStr = "MO", 1, slStr = "TU", 2, slStr = "W", 3, slStr = "WE", 3,...)
        If slStr = "M" Or slStr = "MO" Then
            ilStart = 1
        ElseIf (slStr = "TU") Then
            ilStart = 2
        ElseIf slStr = "W" Or slStr = "WE" Then
            ilStart = 3
        ElseIf (slStr = "TH") Then
            ilStart = 4
        ElseIf slStr = "F" Or slStr = "FR" Then
            ilStart = 5
        ElseIf (slStr = "SA") Then
            ilStart = 6
        ElseIf (slStr = "SU") Then
            ilStart = 7
        End If
        slStr = UCase(slEnd)
        If slStr = "M" Or slStr = "MO" Then
            ilEnd = 1
        ElseIf (slStr = "TU") Then
            ilEnd = 2
        ElseIf slStr = "W" Or slStr = "WE" Then
            ilEnd = 3
        ElseIf (slStr = "TH") Then
            ilEnd = 4
        ElseIf slStr = "F" Or slStr = "FR" Then
            ilEnd = 5
        ElseIf (slStr = "SA") Then
            ilEnd = 6
        ElseIf (slStr = "SU") Then
            ilEnd = 7
        End If
        If (ilStart < 1) Or (ilStart > 7) Or (ilEnd < 1) Or (ilEnd > 7) Or (ilEnd < ilStart) Then
            slDays = "" 'String(7, "N")
            Exit For
        Else
            For ilSet = ilStart To ilEnd Step 1
                Mid$(slDays, ilSet, 1) = "Y"
            Next ilSet
        End If
    Next ilDays
    gCreateDayStr = slDays
End Function

Public Sub gConflictPop()
    Dim ilRet As Integer
    
    ilRet = gGetTypeOfRecs_BDE_BusDefinition("C", sgCurrBDEStamp, "EngrSchd-mPopBDE Bus Definition", tgCurrBDE())
    ilRet = gGetTypeOfRecs_ASE_AudioSource("C", sgCurrASEStamp, "EngrSchd-mPopASE Audio Source", tgCurrASE())
    ilRet = gGetTypeOfRecs_ANE_AudioName("C", sgCurrANEStamp, "EngrSchd-mPopANE Audio Audio Names", tgCurrANE())
    ilRet = gGetTypeOfRecs_ATE_AudioType("C", sgCurrATEStamp, "EngrAudioType-mPopulate", tgCurrATE())
End Sub

Public Function gGetTempDateRange(llStartSelectionDate As Long, llEndSelectionDate As Long, tlTSE() As TSE) As String
    Dim llSdate As Long
    Dim llEDate As Long
    Dim llDate As Long
    Dim ilTSE As Integer
    If UBound(tlTSE) <= LBound(tlTSE) Then
        gGetTempDateRange = "No Dates"
        Exit Function
    End If
    llSdate = 99999999
    For ilTSE = LBound(tlTSE) To UBound(tlTSE) - 1 Step 1
        llDate = gDateValue(tlTSE(ilTSE).sLogDate)
        If (llDate >= llStartSelectionDate) And (llDate <= llEndSelectionDate) Then
            If llDate < llSdate Then
                llSdate = llDate
            End If
            If llDate > llEDate Then
                llEDate = llDate
            End If
        End If
    Next ilTSE
    If llSdate = 99999999 Then
        gGetTempDateRange = ""
        Exit Function
    Else
        If llSdate <> llEDate Then
            gGetTempDateRange = Format$(llSdate, sgShowDateForm) & "-" & Format$(llEDate, sgShowDateForm)
        Else
            gGetTempDateRange = Format$(llSdate, sgShowDateForm)
        End If
    End If
End Function

Private Function mAddConflict(ilStartIndex As Integer, tlConflictList() As CONFLICTLIST) As Integer
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilUpper As Integer
    
    If ilStartIndex = 0 Then
        mAddConflict = UBound(tlConflictList)
        ReDim Preserve tlConflictList(1 To UBound(tlConflictList) + 1) As CONFLICTLIST
        Exit Function
    End If
    mAddConflict = ilStartIndex
    ilIndex = ilStartIndex
    ilUpper = UBound(tlConflictList)
    Do
        If (tlConflictList(ilIndex).sType = "S") And (tlConflictList(ilUpper).sType = "S") Then
            If (tlConflictList(ilIndex).lSheCode = tlConflictList(ilUpper).lSheCode) And (tlConflictList(ilIndex).lSeeCode = tlConflictList(ilUpper).lSeeCode) Then
                Exit Function
            End If
        ElseIf (tlConflictList(ilIndex).sType = "L") And (tlConflictList(ilUpper).sType = "L") Then
            If (tlConflictList(ilIndex).lDheCode = tlConflictList(ilUpper).lDheCode) And (tlConflictList(ilIndex).lDeeCode = tlConflictList(ilUpper).lDeeCode) Then
                Exit Function
            End If
        ElseIf (tlConflictList(ilIndex).sType = "T") And (tlConflictList(ilUpper).sType = "T") Then
            If (tlConflictList(ilIndex).lDheCode = tlConflictList(ilUpper).lDheCode) And (tlConflictList(ilIndex).lDeeCode = tlConflictList(ilUpper).lDeeCode) Then
                Exit Function
            End If
        End If
        ilIndex = tlConflictList(ilIndex).iNextIndex
    Loop While ilIndex > 0
    'Add to chain
    ilIndex = ilStartIndex
    Do
        If tlConflictList(ilIndex).iNextIndex = -1 Then
            tlConflictList(ilIndex).iNextIndex = UBound(tlConflictList)
            ReDim Preserve tlConflictList(1 To UBound(tlConflictList) + 1) As CONFLICTLIST
            Exit Function
        End If
        ilIndex = tlConflictList(ilIndex).iNextIndex
    Loop While ilIndex <> -1
End Function

Public Sub gGetPreAndPostAudioTime(slAudio As String, llPreTime As Long, llPostTime As Long, slCheckConflicts As String)
    Dim ilANE As Integer
    Dim ilATE As Integer
    
    llPreTime = 0
    llPostTime = 0
    slCheckConflicts = "Y"
    If slAudio = "" Then
        Exit Sub
    End If
    If StrComp(slAudio, "[None]", vbTextCompare) = 0 Then
        Exit Sub
    End If
    
    'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
    '    If StrComp(slAudio, Trim$(tgCurrANE(ilANE).sName), vbTextCompare) = 0 Then
    ilANE = gBinarySearchName(slAudio, tgCurrANE_Name())
    If ilANE <> -1 Then
        ilANE = gBinarySearchANE(tgCurrANE_Name(ilANE).iCode, tgCurrANE())
        If ilANE <> -1 Then
            slCheckConflicts = tgCurrANE(ilANE).sCheckConflicts
            If tgCurrANE(ilANE).sCheckConflicts <> "N" Then
                For ilATE = 0 To UBound(tgCurrATE) - 1 Step 1
                    If tgCurrANE(ilANE).iAteCode = tgCurrATE(ilATE).iCode Then
                        llPreTime = tgCurrATE(ilATE).lPreBufferTime
                        llPostTime = tgCurrATE(ilATE).lPostBufferTime
                        Exit Sub
                    End If
                Next ilATE
            End If
            Exit Sub
        End If
    'Next ilANE
    End If
End Sub

Private Sub mCreateAudioRecs(ilPass As Integer, llRow As Long, slType As String, slAudio As String, llEventStartTime As Long, llEventEndTime As Long, slDays As String, tlConflict() As CONFLICTTEST)
    Dim llUpper As Long
    Dim llPreTime As Long
    Dim llPostTime As Long
    Dim slEventDays As String
    Dim ilDay As Integer
    Dim slCheckConflicts As String
    
    If slAudio = "" Then
        Exit Sub
    End If
    llUpper = UBound(tlConflict)
    tlConflict(llUpper).lRow = llRow
    tlConflict(llUpper).sType = slType
    tlConflict(llUpper).sDays = slDays
    gGetPreAndPostAudioTime slAudio, llPreTime, llPostTime, slCheckConflicts
    If slCheckConflicts = "N" Then
        Exit Sub
    End If
    If llEventEndTime <= 864000 Then
        If llEventStartTime - llPreTime >= 0 Then
            If llEventEndTime + llPostTime <= 864000 Then
                If ilPass = 0 Then
                    tlConflict(llUpper).lEventStartTime = llEventStartTime - llPreTime
                    tlConflict(llUpper).lEventEndTime = llEventEndTime + llPostTime
                    llUpper = llUpper + 1
                    ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
                End If
            Else
                If ilPass = 0 Then
                    tlConflict(llUpper).lEventStartTime = llEventStartTime - llPreTime
                    tlConflict(llUpper).lEventEndTime = 864000
                    llUpper = llUpper + 1
                    ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
                End If
                If (ilPass = 0) Or (ilPass = 2) Then
                    tlConflict(llUpper).lRow = llRow
                    tlConflict(llUpper).sType = slType
                    tlConflict(llUpper).lEventStartTime = 0
                    tlConflict(llUpper).lEventEndTime = llEventEndTime - 864000 + llPostTime
                    slEventDays = String(7, "N")
                    For ilDay = 1 To 7 Step 1
                        If ilDay = 7 Then
                            If Mid$(slDays, ilDay, 1) = "Y" Then
                                Mid(slEventDays, 1, 1) = "Y"
                            End If
                        Else
                            If Mid$(slDays, ilDay, 1) = "Y" Then
                                Mid(slEventDays, ilDay + 1, 1) = "Y"
                            End If
                        End If
                    Next ilDay
                    tlConflict(llUpper).sDays = slEventDays
                    llUpper = llUpper + 1
                    ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
                End If
            End If
        Else
            If (ilPass = 0) Or (ilPass = 1) Then
                tlConflict(llUpper).lEventStartTime = 864000 + (llEventStartTime - llPreTime)
                tlConflict(llUpper).lEventEndTime = 864000
                slEventDays = String(7, "N")
                For ilDay = 7 To 1 Step -1
                    If ilDay = 1 Then
                        If Mid$(slDays, ilDay, 1) = "Y" Then
                            Mid(slEventDays, 7, 1) = "Y"
                        End If
                    Else
                        If Mid$(slDays, ilDay, 1) = "Y" Then
                            Mid(slEventDays, ilDay - 1, 1) = "Y"
                        End If
                    End If
                Next ilDay
                tlConflict(llUpper).sDays = slEventDays
                llUpper = llUpper + 1
                ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
            End If
            If ilPass = 0 Then
                tlConflict(llUpper).lRow = llRow
                tlConflict(llUpper).sType = slType
                tlConflict(llUpper).sDays = slDays
                tlConflict(llUpper).lEventStartTime = 0
                tlConflict(llUpper).lEventEndTime = llEventEndTime + llPostTime
                llUpper = llUpper + 1
                ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
            End If
        End If
    Else
        If ilPass = 0 Then
            tlConflict(llUpper).lEventStartTime = llEventStartTime - llPreTime
            tlConflict(llUpper).lEventEndTime = 864000
            llUpper = llUpper + 1
            ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
        End If
        If (ilPass = 0) Or (ilPass = 2) Then
            tlConflict(llUpper).lRow = llRow
            tlConflict(llUpper).sType = slType
            tlConflict(llUpper).sDays = slDays
            tlConflict(llUpper).lEventStartTime = 0
            tlConflict(llUpper).lEventEndTime = llEventEndTime - 864000 + llPostTime
            slEventDays = String(7, "N")
            For ilDay = 1 To 7 Step 1
                If ilDay = 7 Then
                    If Mid$(slDays, ilDay, 1) = "Y" Then
                        Mid(slEventDays, 1, 1) = "Y"
                    End If
                Else
                    If Mid$(slDays, ilDay, 1) = "Y" Then
                        Mid(slEventDays, ilDay + 1, 1) = "Y"
                    End If
                End If
            Next ilDay
            tlConflict(llUpper).sDays = slEventDays
            llUpper = llUpper + 1
            ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
        End If
    End If
End Sub

Private Sub mCreateBusRecs(ilPass As Integer, llRow As Long, slType As String, slIgnoreConflicts As String, llEventStartTime As Long, llEventEndTime As Long, slDays As String, tlConflict() As CONFLICTTEST)
    Dim llUpper As Long
    Dim ilDay As Integer
    Dim slEventDays As String
    
    If (slIgnoreConflicts = "B") Or (slIgnoreConflicts = "I") Then
        Exit Sub
    End If
    llUpper = UBound(tlConflict)
    tlConflict(llUpper).lRow = llRow
    tlConflict(llUpper).sType = slType
    tlConflict(llUpper).sDays = slDays
    If llEventEndTime <= 864000 Then
        If ilPass = 0 Then
            tlConflict(llUpper).lEventStartTime = llEventStartTime
            tlConflict(llUpper).lEventEndTime = llEventEndTime
            llUpper = llUpper + 1
            ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
        End If
    Else
        If ilPass = 0 Then
            tlConflict(llUpper).lEventStartTime = llEventStartTime
            tlConflict(llUpper).lEventEndTime = 864000
            llUpper = llUpper + 1
            ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
        End If
        If (ilPass = 0) Or (ilPass = 2) Then
            tlConflict(llUpper).lRow = llRow
            tlConflict(llUpper).sType = slType
            tlConflict(llUpper).lEventStartTime = 0
            tlConflict(llUpper).lEventEndTime = llEventEndTime - 864000
            slEventDays = String(7, "N")
            For ilDay = 1 To 7 Step 1
                If ilDay = 7 Then
                    If Mid$(slDays, ilDay, 1) = "Y" Then
                        Mid(slEventDays, 1, 1) = "Y"
                    End If
                Else
                    If Mid$(slDays, ilDay, 1) = "Y" Then
                        Mid(slEventDays, ilDay + 1, 1) = "Y"
                    End If
                End If
            Next ilDay
            tlConflict(llUpper).sDays = slEventDays
            llUpper = llUpper + 1
            ReDim Preserve tlConflict(1 To llUpper) As CONFLICTTEST
        End If
    End If
End Sub

Public Function gConflictTableCreations(slInStartDate As String, slInEndDate As String, hlCME As Integer, hlSEE As Integer) As Integer
    Dim ilRet As Integer
    Dim slDateStartRange As String
    Dim slDateEndRange As String
    Dim llDateStartRange As Long
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilDHE As Integer
    Dim ilDEE As Integer
    Dim ilASE As Integer
    Dim ilSHE As Integer
    Dim ilPriAneCode As Integer
    Dim ilProtAneCode As Integer
    Dim ilBkupAneCode As Integer
    Dim ilHour As Integer
    Dim slHours As String
    Dim ilLoop As Integer
    Dim llOffsetStartTime As Long
    Dim llOffsetEndTime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llLength As Long
    Dim slDate As String
    Dim tlTSE As TSE
    Dim ilSEE As Integer
    Dim ilETE As Integer
    Dim ilSpotETECode As Integer

    gConflictTableCreations = True
    gConflictPop
    slDateStartRange = slInStartDate
    slDateEndRange = slInEndDate
    llDateStartRange = gDateValue(slDateStartRange)
    
    'Clear CME of all records
    ilRet = gPutDelete_CME_Conflict_Master("", 0, 0, 0, "gConflictTableCreations", hlCME)
    ilRet = gClearFile("CEE_Conflict_Events", "gConflictTableCreations")
    ilRet = gGetTypeOfRecs_ETE_EventType("C", sgCurrETEStamp, "gConflictTableCreations-Get ETE", tgCurrETE())
    ilSpotETECode = 0
    For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
        If tgCurrETE(ilETE).sCategory = "S" Then
            ilSpotETECode = tgCurrETE(ilETE).iCode
            Exit For
        End If
    Next ilETE
    'Create Schedule
    ilRet = gGetTypeOfRecs_SHE_ScheduleHeaderByDate(slDateStartRange, "gConflictTableCreations-Get SHE", tmCurrSHE())
    For ilSHE = 0 To UBound(tmCurrSHE) - 1 Step 1
        LSet tmSHE = tmCurrSHE(ilSHE)
        slDate = tmSHE.sAirDate
        If gDateValue(slDate) >= llDateStartRange Then
            llDateStartRange = gDateValue(slDate) + 1
        End If
        smCurrSEEStamp = ""
        ilRet = gGetRecs_SEE_ScheduleEventsAPI(hlSEE, smCurrSEEStamp, -1, tmSHE.lCode, "EngrSchd-Get Events", tmCurrSEE())
        ReDim tmCurrLibDEE(0 To 0) As DEE
        smCurrLibDEEStamp = ""
        If ilRet Then
            For ilSEE = 0 To UBound(tmCurrSEE) - 1 Step 1
'                If (tmCurrSEE(ilSEE).sAction <> "D") And (tmCurrSEE(ilSEE).sAction <> "R") Then
'                    ilPriAneCode = 0
'                    ilProtAneCode = 0
'                    ilBkupAneCode = 0
'                    If (tmCurrSEE(ilSEE).sIgnoreConflicts <> "A") And (tmCurrSEE(ilSEE).sIgnoreConflicts <> "I") Then
'                        If tmCurrSEE(ilSEE).iAudioAseCode > 0 Then
'                            For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
'                                If tmCurrSEE(ilSEE).iAudioAseCode = tgCurrASE(ilASE).iCode Then
'                                    ilPriAneCode = tgCurrASE(ilASE).iPriAneCode
'                                    Exit For
'                                End If
'                            Next ilASE
'                        End If
'                        ilProtAneCode = tmCurrSEE(ilSEE).iProtAneCode
'                        ilBkupAneCode = tmCurrSEE(ilSEE).iBkupAneCode
'                    End If
'                    llStartTime = tmCurrSEE(ilSEE).lTime
'                    llEndTime = tmCurrSEE(ilSEE).lTime + tmCurrSEE(ilSEE).lDuration
'                    llStartDate = gDateValue(tmSHE.sAirDate)
'                    llEndDate = gDateValue(tmSHE.sAirDate)
'                    tmCurrLibDEE(0).lDheCode = tmSHE.lCode
'                    tmCurrLibDEE(0).lCode = tmCurrSEE(ilSEE).lCode
'                    tmCurrLibDEE(0).sIgnoreConflicts = tmCurrSEE(ilSEE).sIgnoreConflicts
'                    tmCurrLibDEE(0).iFneCode = tmCurrSEE(ilSEE).iBdeCode
'                    tmCurrLibDEE(0).sDays = String(7, "N")
'                    Select Case Weekday(tmSHE.sAirDate)
'                        Case vbMonday
'                            Mid(tmCurrLibDEE(0).sDays, 1, 1) = "Y"
'                        Case vbTuesday
'                            Mid(tmCurrLibDEE(0).sDays, 2, 1) = "Y"
'                        Case vbWednesday
'                            Mid(tmCurrLibDEE(0).sDays, 3, 1) = "Y"
'                        Case vbThursday
'                            Mid(tmCurrLibDEE(0).sDays, 4, 1) = "Y"
'                        Case vbFriday
'                            Mid(tmCurrLibDEE(0).sDays, 5, 1) = "Y"
'                        Case vbSaturday
'                            Mid(tmCurrLibDEE(0).sDays, 6, 1) = "Y"
'                        Case vbSunday
'                            Mid(tmCurrLibDEE(0).sDays, 7, 1) = "Y"
'                    End Select
'                    mConflictCMEBusRec "S", tmCurrLibDEE(0), 0, llStartDate, llEndDate, llStartTime, llEndTime
'                    mConflictCMEAudioRec "S", tmCurrLibDEE(0), 0, ilPriAneCode, tmCurrSEE(ilSEE).sAudioItemID, llStartDate, llEndDate, llStartTime, llEndTime
'                    mConflictCMEAudioRec "S", tmCurrLibDEE(0), 0, ilProtAneCode, tmCurrSEE(ilSEE).sProtItemID, llStartDate, llEndDate, llStartTime, llEndTime
'                    mConflictCMEAudioRec "S", tmCurrLibDEE(0), 0, ilBkupAneCode, tmCurrSEE(ilSEE).sAudioItemID, llStartDate, llEndDate, llStartTime, llEndTime
'                End If
                ilRet = gCreateCMEForSchd(tmSHE, tmCurrSEE(ilSEE), ilSpotETECode, hlCME)
            Next ilSEE
        End If
    Next ilSHE
    slDateStartRange = Format$(llDateStartRange, "ddddd")

    smCurrLibDHEStamp = ""
    ilRet = gGetTypeOfRecs_DHE_DayHeaderInfoForLibByRange("C", "A", slDateStartRange, slDateEndRange, smCurrLibDHEStamp, "EngrLib-mPopulate", tmCurrLibDHE())
    For ilDHE = 0 To UBound(tmCurrLibDHE) - 1 Step 1
'        slStartDate = tmCurrLibDHE(ilDHE).sStartDate
'        If gDateValue(slDateStartRange) > gDateValue(slStartDate) Then
'            slStartDate = slDateStartRange
'        End If
'        slEndDate = tmCurrLibDHE(ilDHE).sEndDate
'        ilRet = gGetRecs_DEE_DayEvent(smCurrLibDEEStamp, tmCurrLibDHE(ilDHE).lCode, "EngrLibDef-mPopulate", tmCurrLibDEE())
'        For ilDEE = 0 To UBound(tmCurrLibDEE) - 1 Step 1
'            llOffsetStartTime = tmCurrLibDEE(ilDEE).lTime
'            llOffsetEndTime = llOffsetStartTime + tmCurrLibDEE(ilDEE).lDuration ' - 1
'            If llOffsetEndTime < llOffsetStartTime Then
'                llOffsetEndTime = llOffsetStartTime
'            End If
'            ilPriAneCode = 0
'            ilProtAneCode = 0
'            ilBkupAneCode = 0
'            If (tmCurrLibDEE(ilDEE).sIgnoreConflicts <> "A") And (tmCurrLibDEE(ilDEE).sIgnoreConflicts <> "I") Then
'                If tmCurrLibDEE(ilDEE).iAudioAseCode > 0 Then
'                    For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
'                        If tmCurrLibDEE(ilDEE).iAudioAseCode = tgCurrASE(ilASE).iCode Then
'                            ilPriAneCode = tgCurrASE(ilASE).iPriAneCode
'                            Exit For
'                        End If
'                    Next ilASE
'                End If
'                ilProtAneCode = tmCurrLibDEE(ilDEE).iProtAneCode
'                ilBkupAneCode = tmCurrLibDEE(ilDEE).iBkupAneCode
'            End If
'            For ilHour = 1 To 24 Step 1
'                If Mid$(tmCurrLibDEE(ilDEE).sHours, ilHour, 1) = "Y" Then
'                    llStartTime = 36000 * (ilHour - 1) + llOffsetStartTime
'                    llEndTime = 36000 * (ilHour - 1) + llOffsetEndTime
'                    llStartDate = gDateValue(slStartDate)
'                    llEndDate = gDateValue(slEndDate)
'                    mConflictCMEBusRec "L", tmCurrLibDEE(ilDEE), tmCurrLibDHE(ilDHE).lDSECode, llStartDate, llEndDate, llStartTime, llEndTime
'                    mConflictCMEAudioRec "L", tmCurrLibDEE(ilDEE), tmCurrLibDHE(ilDHE).lDSECode, ilPriAneCode, tmCurrLibDEE(ilDEE).sAudioItemID, llStartDate, llEndDate, llStartTime, llEndTime
'                    mConflictCMEAudioRec "L", tmCurrLibDEE(ilDEE), tmCurrLibDHE(ilDHE).lDSECode, ilProtAneCode, tmCurrLibDEE(ilDEE).sProtItemID, llStartDate, llEndDate, llStartTime, llEndTime
'                    mConflictCMEAudioRec "L", tmCurrLibDEE(ilDEE), tmCurrLibDHE(ilDHE).lDSECode, ilBkupAneCode, tmCurrLibDEE(ilDEE).sAudioItemID, llStartDate, llEndDate, llStartTime, llEndTime
'                End If
'            Next ilHour
'        Next ilDEE
        ilRet = gCreateCMEForLib(tmCurrLibDHE(ilDHE), slDateStartRange, hlCME)
    Next ilDHE
    
    'Create Templates
    smCurrTempDHETSEStamp = ""
    ilRet = gGetTypeOfRecs_DHETSE_DayHeaderInfoForTempByRange("C", slDateStartRange, slDateEndRange, smCurrTempDHETSEStamp, "EngrLib-mPopulate", tmCurrTempDHETSE())
    For ilDHE = 0 To UBound(tmCurrTempDHETSE) - 1 Step 1
'        If (tmCurrTempDHETSE(ilDHE).tDHE.sState <> "D") And (tmCurrTempDHETSE(ilDHE).tDHE.sState <> "L") And (tmCurrTempDHETSE(ilDHE).tTSE.sState <> "D") And (tmCurrTempDHETSE(ilDHE).tTSE.sState <> "L") Then
'            If (tmCurrTempDHETSE(ilDHE).tDHE.sIgnoreConflicts <> "I") Then
'                smCurrLibDEEStamp = ""
'                ilRet = gGetRecs_DEE_DayEvent(smCurrLibDEEStamp, tmCurrTempDHETSE(ilDHE).tDHE.lCode, "EngrLibDef-mPopulate", tmCurrLibDEE())
'                LSet tlTSE = tmCurrTempDHETSE(ilDHE).tTSE
'                For ilDEE = 0 To UBound(tmCurrLibDEE) - 1 Step 1
'                    tmCurrLibDEE(ilDEE).sIgnoreConflicts = "N"
'                    ilHour = Hour(tlTSE.sStartTime)
'                    If ilHour <> 0 Then
'                        slHours = tmCurrLibDEE(ilDEE).sHours
'                        tmCurrLibDEE(ilDEE).sHours = String(24, "N")
'                        For ilLoop = 0 To 23 Step 1
'                            Mid$(tmCurrLibDEE(ilDEE).sHours, ilHour + 1, 1) = Mid$(slHours, ilLoop + 1, 1)
'                            ilHour = ilHour + 1
'                            If ilHour > 23 Then
'                                Exit For
'                            End If
'                        Next ilLoop
'                    End If
'                    tmCurrLibDEE(ilDEE).sDays = String(7, "N")
'                    Select Case Weekday(tlTSE.sLogDate)
'                        Case vbMonday
'                            Mid(tmCurrLibDEE(ilDEE).sDays, 1, 1) = "Y"
'                        Case vbTuesday
'                            Mid(tmCurrLibDEE(ilDEE).sDays, 2, 1) = "Y"
'                        Case vbWednesday
'                            Mid(tmCurrLibDEE(ilDEE).sDays, 3, 1) = "Y"
'                        Case vbThursday
'                            Mid(tmCurrLibDEE(ilDEE).sDays, 4, 1) = "Y"
'                        Case vbFriday
'                            Mid(tmCurrLibDEE(ilDEE).sDays, 5, 1) = "Y"
'                        Case vbSaturday
'                            Mid(tmCurrLibDEE(ilDEE).sDays, 6, 1) = "Y"
'                        Case vbSunday
'                            Mid(tmCurrLibDEE(ilDEE).sDays, 7, 1) = "Y"
'                    End Select
'                    tmCurrLibDEE(ilDEE).iFneCode = tlTSE.iBdeCode
'
'                    llOffsetStartTime = tmCurrLibDEE(ilDEE).lTime
'                    llOffsetEndTime = llOffsetStartTime + tmCurrLibDEE(ilDEE).lDuration ' - 1
'                    If llOffsetEndTime < llOffsetStartTime Then
'                        llOffsetEndTime = llOffsetStartTime
'                    End If
'                    ilPriAneCode = 0
'                    ilProtAneCode = 0
'                    ilBkupAneCode = 0
'                    If (tmCurrLibDEE(ilDEE).sIgnoreConflicts <> "A") And (tmCurrLibDEE(ilDEE).sIgnoreConflicts <> "I") Then
'                        If tmCurrLibDEE(ilDEE).iAudioAseCode > 0 Then
'                            For ilASE = 0 To UBound(tgCurrASE) - 1 Step 1
'                                If tmCurrLibDEE(ilDEE).iAudioAseCode = tgCurrASE(ilASE).iCode Then
'                                    ilPriAneCode = tgCurrASE(ilASE).iPriAneCode
'                                    Exit For
'                                End If
'                            Next ilASE
'                        End If
'                        ilProtAneCode = tmCurrLibDEE(ilDEE).iProtAneCode
'                        ilBkupAneCode = tmCurrLibDEE(ilDEE).iBkupAneCode
'                    End If
'                    For ilHour = 1 To 24 Step 1
'                        If Mid$(tmCurrLibDEE(ilDEE).sHours, ilHour, 1) = "Y" Then
'                            llStartTime = 36000 * (ilHour - 1) + llOffsetStartTime
'                            llEndTime = 36000 * (ilHour - 1) + llOffsetEndTime
'                            llLength = 10 * (gTimeToLong(tlTSE.sStartTime, False) Mod 3600)
'                            llStartTime = llStartTime + llLength
'                            llEndTime = llEndTime + llLength
'                            llStartDate = gDateValue(slStartDate)
'                            llEndDate = gDateValue(slEndDate)
'                            mConflictCMEBusRec "T", tmCurrLibDEE(ilDEE), tmCurrLibDHE(ilDHE).lDSECode, llStartDate, llEndDate, llStartTime, llEndTime
'                            mConflictCMEAudioRec "T", tmCurrLibDEE(ilDEE), tmCurrLibDHE(ilDHE).lDSECode, ilPriAneCode, tmCurrLibDEE(ilDEE).sAudioItemID, llStartDate, llEndDate, llStartTime, llEndTime
'                            mConflictCMEAudioRec "T", tmCurrLibDEE(ilDEE), tmCurrLibDHE(ilDHE).lDSECode, ilProtAneCode, tmCurrLibDEE(ilDEE).sProtItemID, llStartDate, llEndDate, llStartTime, llEndTime
'                            mConflictCMEAudioRec "T", tmCurrLibDEE(ilDEE), tmCurrLibDHE(ilDHE).lDSECode, ilBkupAneCode, tmCurrLibDEE(ilDEE).sAudioItemID, llStartDate, llEndDate, llStartTime, llEndTime
'                        End If
'                    Next ilHour
'                Next ilDEE
'            End If
'        End If
        ilRet = gCreateCMEForTemp(tmCurrTempDHETSE(ilDHE).tDHE, tmCurrTempDHETSE(ilDHE).tTSE, hlCME)
    Next ilDHE


End Function

Private Sub mConflictCEEBusRec(ilGridEventCol As Integer, slSource As String, slIgnoreConflicts As String, slDays As String, slBusNames() As String, llStartDate As Long, llEndDate As Long, llStartTime As Long, llEndTime As Long)
    Dim tlCEE As CEE
    Dim ilDay As Integer
    Dim ilBus As Integer
    Dim ilRet As Integer
    Dim ilBDE As Integer
    
    
    DoEvents
    If (slIgnoreConflicts = "B") Or (slIgnoreConflicts = "I") Then
        Exit Sub
    End If
    If tgSOE.sMatchBNotT = "N" Then
        Exit Sub
    End If
    For ilBus = LBound(slBusNames) To UBound(slBusNames) Step 1
        DoEvents
        For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
            DoEvents
            If StrComp(Trim$(tgCurrBDE(ilBDE).sName), Trim$(slBusNames(ilBus)), vbTextCompare) = 0 Then
                DoEvents
                For ilDay = 1 To 7 Step 1
                    DoEvents
                    If Mid$(slDays, ilDay, 1) = "Y" Then
                        tlCEE.lCode = 0
                        tlCEE.sEvtType = "B"
                        tlCEE.iBdeCode = tgCurrBDE(ilBDE).iCode
                        tlCEE.iANECode = 0
                        tlCEE.lGenDate = lmGenDate
                        tlCEE.lGenTime = lmGenTime
                        tlCEE.lGridEventRow = lmGridEventRow
                        tlCEE.iGridEventCol = ilGridEventCol
                        If llEndTime <= 864000 Then
                            tlCEE.lStartDate = llStartDate
                            tlCEE.lEndDate = llEndDate
                            Select Case ilDay
                                Case 1
                                    tlCEE.sDay = "Mo"
                                Case 2
                                    tlCEE.sDay = "Tu"
                                Case 3
                                    tlCEE.sDay = "We"
                                Case 4
                                    tlCEE.sDay = "Th"
                                Case 5
                                    tlCEE.sDay = "Fr"
                                Case 6
                                    tlCEE.sDay = "Sa"
                                Case 7
                                    tlCEE.sDay = "Su"
                            End Select
                            tlCEE.lStartTime = llStartTime
                            tlCEE.lEndTime = llEndTime
                            tlCEE.sUnused = ""
                            'ilRet = gPutInsert_CEE_Conflict_Events(tlCEE, "mConflictCEEBusRec")
                            ilRet = gGetConflicts_CME(tlCEE, smTypeSource, "gConflictTableCheck", tmConflictResults())
                        Else
                            tlCEE.lStartDate = llStartDate
                            tlCEE.lEndDate = llEndDate
                            Select Case ilDay
                                Case 1
                                    tlCEE.sDay = "Mo"
                                Case 2
                                    tlCEE.sDay = "Tu"
                                Case 3
                                    tlCEE.sDay = "We"
                                Case 4
                                    tlCEE.sDay = "Th"
                                Case 5
                                    tlCEE.sDay = "Fr"
                                Case 6
                                    tlCEE.sDay = "Sa"
                                Case 7
                                    tlCEE.sDay = "Su"
                            End Select
                            tlCEE.lStartTime = llStartTime
                            tlCEE.lEndTime = 864000
                            tlCEE.sUnused = ""
                            'ilRet = gPutInsert_CEE_Conflict_Events(tlCEE, "mConflictCEEBusRec")
                            ilRet = gGetConflicts_CME(tlCEE, smTypeSource, "gConflictTableCheck", tmConflictResults())
                            tlCEE.lCode = 0
                            tlCEE.lStartDate = llStartDate + 1
                            If llEndDate <> gDateValue("12/31/2069") Then
                                tlCEE.lEndDate = llEndDate + 1
                            End If
                            Select Case ilDay
                                Case 1
                                    tlCEE.sDay = "Tu"
                                Case 2
                                    tlCEE.sDay = "We"
                                Case 3
                                    tlCEE.sDay = "Th"
                                Case 4
                                    tlCEE.sDay = "Fr"
                                Case 5
                                    tlCEE.sDay = "Sa"
                                Case 6
                                    tlCEE.sDay = "Su"
                                Case 7
                                    tlCEE.sDay = "Mo"
                            End Select
                            tlCEE.lStartTime = 0
                            tlCEE.lEndTime = llEndTime - 864000
                            tlCEE.sUnused = ""
                            'ilRet = gPutInsert_CEE_Conflict_Events(tlCEE, "mConflictCEEBusRec")
                            ilRet = gGetConflicts_CME(tlCEE, smTypeSource, "gConflictTableCheck", tmConflictResults())
                        End If
                    End If
                Next ilDay
                Exit For
            End If
        Next ilBDE
    Next ilBus
End Sub


Private Sub mConflictCEEAudioRec(slEventType As String, ilGridEventCol As Integer, slSource As String, slDays As String, slBusNames() As String, ilANECode As Integer, llStartDate As Long, llEndDate As Long, llStartTime As Long, llEndTime As Long)
    Dim ilANE As Integer
    Dim ilATE As Integer
    Dim llPreTime As Long
    Dim llPostTime As Long
    Dim llSTime As Long
    Dim llETime As Long
    Dim ilDay As Integer
    Dim ilRet As Integer
    Dim ilBus As Integer
    Dim ilBDE As Integer
    Dim ilETE As Integer
    Dim ilBdeCode As Integer
    Dim slEventCategory As String
    Dim tlCEE As CEE
    
    DoEvents
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
    For ilBus = LBound(slBusNames) To UBound(slBusNames) Step 1
        DoEvents
        For ilBDE = 0 To UBound(tgCurrBDE) - 1 Step 1
            DoEvents
            If StrComp(Trim$(tgCurrBDE(ilBDE).sName), Trim$(slBusNames(ilBus)), vbTextCompare) = 0 Then
                DoEvents
                'Only test bus is avail event
                ilBdeCode = tgCurrBDE(ilBDE).iCode
                For ilETE = 0 To UBound(tgCurrETE) - 1 Step 1
                    If StrComp(Trim$(tgCurrETE(ilETE).sName), Trim$(slEventType), vbTextCompare) = 0 Then
                        If tgCurrETE(ilETE).sCategory <> "A" Then
                            ilBdeCode = 0
                        End If
                        Exit For
                    End If
                Next ilETE
                For ilDay = 1 To 7 Step 1
                    DoEvents
                    If Mid$(slDays, ilDay, 1) = "Y" Then
                        tlCEE.lCode = 0
                        tlCEE.sEvtType = "A"
                        tlCEE.iBdeCode = ilBdeCode
                        tlCEE.iANECode = ilANECode
                        tlCEE.lGenDate = lmGenDate
                        tlCEE.lGenTime = lmGenTime
                        tlCEE.lGridEventRow = lmGridEventRow
                        tlCEE.iGridEventCol = ilGridEventCol
                        If llSTime < 0 Then
                            tlCEE.lStartDate = llStartDate - 1
                            tlCEE.lEndDate = llEndDate - 1
                            Select Case ilDay
                                Case 1
                                    tlCEE.sDay = "Su"
                                Case 2
                                    tlCEE.sDay = "Mo"
                                Case 3
                                    tlCEE.sDay = "Tu"
                                Case 4
                                    tlCEE.sDay = "We"
                                Case 5
                                    tlCEE.sDay = "Th"
                                Case 6
                                    tlCEE.sDay = "Fr"
                                Case 7
                                    tlCEE.sDay = "Sa"
                            End Select
                            tlCEE.lStartTime = 864000 + llSTime
                            tlCEE.lEndTime = 864000
                            tlCEE.sUnused = ""
                            'ilRet = gPutInsert_CEE_Conflict_Events(tlCEE, "mConflictCEEAudioRec")
                            ilRet = gGetConflicts_CME(tlCEE, smTypeSource, "gConflictTableCheck", tmConflictResults())
                            tlCEE.lCode = 0
                            tlCEE.lStartDate = llStartDate + 1
                            If llEndDate <> gDateValue("12/31/2069") Then
                                tlCEE.lEndDate = llEndDate + 1
                            End If
                            Select Case ilDay
                                Case 1
                                    tlCEE.sDay = "Mo"
                                Case 2
                                    tlCEE.sDay = "Tu"
                                Case 3
                                    tlCEE.sDay = "We"
                                Case 4
                                    tlCEE.sDay = "Th"
                                Case 5
                                    tlCEE.sDay = "Fr"
                                Case 6
                                    tlCEE.sDay = "Sa"
                                Case 7
                                    tlCEE.sDay = "Su"
                            End Select
                            tlCEE.lStartTime = 0
                            tlCEE.lEndTime = llETime
                            tlCEE.sUnused = ""
                            'ilRet = gPutInsert_CEE_Conflict_Events(tlCEE, "mConflictCEEAudioRec")
                            ilRet = gGetConflicts_CME(tlCEE, smTypeSource, "gConflictTableCheck", tmConflictResults())
                        Else
                            If llEndTime <= 864000 Then
                                tlCEE.lStartDate = llStartDate
                                tlCEE.lEndDate = llEndDate
                                Select Case ilDay
                                    Case 1
                                        tlCEE.sDay = "Mo"
                                    Case 2
                                        tlCEE.sDay = "Tu"
                                    Case 3
                                        tlCEE.sDay = "We"
                                    Case 4
                                        tlCEE.sDay = "Th"
                                    Case 5
                                        tlCEE.sDay = "Fr"
                                    Case 6
                                        tlCEE.sDay = "Sa"
                                    Case 7
                                        tlCEE.sDay = "Su"
                                End Select
                                tlCEE.lStartTime = llStartTime
                                tlCEE.lEndTime = llEndTime
                                tlCEE.sUnused = ""
                                'ilRet = gPutInsert_CEE_Conflict_Events(tlCEE, "mConflictCEEAudioRec")
                                ilRet = gGetConflicts_CME(tlCEE, smTypeSource, "gConflictTableCheck", tmConflictResults())
                            Else
                                tlCEE.lStartDate = llStartDate
                                tlCEE.lEndDate = llEndDate
                                Select Case ilDay
                                    Case 1
                                        tlCEE.sDay = "Mo"
                                    Case 2
                                        tlCEE.sDay = "Tu"
                                    Case 3
                                        tlCEE.sDay = "We"
                                    Case 4
                                        tlCEE.sDay = "Th"
                                    Case 5
                                        tlCEE.sDay = "Fr"
                                    Case 6
                                        tlCEE.sDay = "Sa"
                                    Case 7
                                        tlCEE.sDay = "Su"
                                End Select
                                tlCEE.lStartTime = llStartTime
                                tlCEE.lEndTime = 864000
                                tlCEE.sUnused = ""
                                'ilRet = gPutInsert_CEE_Conflict_Events(tlCEE, "mConflictCEEAudioRec")
                                ilRet = gGetConflicts_CME(tlCEE, smTypeSource, "gConflictTableCheck", tmConflictResults())
                                tlCEE.lCode = 0
                                tlCEE.lStartDate = llStartDate + 1
                                If llEndDate <> gDateValue("12/31/2069") Then
                                    tlCEE.lEndDate = llEndDate + 1
                                End If
                                Select Case ilDay
                                    Case 1
                                        tlCEE.sDay = "Tu"
                                    Case 2
                                        tlCEE.sDay = "We"
                                    Case 3
                                        tlCEE.sDay = "Th"
                                    Case 4
                                        tlCEE.sDay = "Fr"
                                    Case 5
                                        tlCEE.sDay = "Sa"
                                    Case 6
                                        tlCEE.sDay = "Su"
                                    Case 7
                                        tlCEE.sDay = "Mo"
                                End Select
                                tlCEE.lStartTime = 0
                                tlCEE.lEndTime = llEndTime - 864000
                                tlCEE.sUnused = ""
                                'ilRet = gPutInsert_CEE_Conflict_Events(tlCEE, "mConflictCEEAudioRec")
                                ilRet = gGetConflicts_CME(tlCEE, smTypeSource, "gConflictTableCheck", tmConflictResults())
                            End If
                        End If
                    End If
                Next ilDay
                Exit For
            End If
        Next ilBDE
    Next ilBus
End Sub



Public Function gConflictTableCheck(slType As String, llDheCode As Long, llOverlappedDHE As Long, slInStartDate As String, slInEndDate As String, slTempStartTime As String, grdEvents As MSHFlexGrid, ilCols() As Integer, tlConflictList() As CONFLICTLIST) As Integer
    'slType- L=Library, T=Template, S=Schedule
    'ilCols(0) = ERRORFLAGINDEX
    'ilCols(1) = EVENTTYPEINDEX
    'ilCols(2) = AIRHOURSINDEX
    'ilCols(3) = AIRDAYSINDEX
    'ilCols(4) = TIMEINDEX
    'ilCols(5) = DURATIONINDEX
    'ilCols(6) = BUSNAMEINDEX
    'ilCols(7) = AUDIONAMEINDEX
    'ilCols(8) = PROTNAMEINDEX
    'ilCols(9) = BACKUPNAMEINDEX
    'ilCols(10) = AUDIOITEMIDINDEX
    'ilCols(11) = PROTITEMIDINDEX
    'ilCols(12) = BACKUPITEMIDINDEX
    'ilCols(13) = CHGSTATUSINDEX
    'ilCols(14) = Ignore Conflicts (A=Audio, B=Bus; I=Both)
    'ilCols(15) = DEECode Index
    Dim llRow As Long
    Dim ilRet As Integer
    Dim llUpper As Long
    Dim llEvent As Long
    Dim ilANE As Integer
    Dim slAudio As String
    Dim ilPriAneCode As Integer
    Dim ilProtAneCode As Integer
    Dim ilBkupAneCode As Integer
    Dim ilSet As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llOffsetStartTime As Long
    Dim llOffsetEndTime As Long
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim slBuses As String
    ReDim slBusNames(1 To 1) As String
    Dim slTestHours As String
    Dim slHours As String
    Dim slStr As String
    Dim ilHour As Integer
    Dim slDays As String
    Dim slDateTime As String
    Dim slIgnoreConflict As String
    Dim ilStartHour As Integer
    Dim ilLoop As Integer
    Dim ilConflictIndex As Integer
    Dim ilStartConflictIndex As Integer
    Dim ilShowConflict As Integer
    Dim slItemID1 As String
    Dim slItemID2 As String
    Dim slEventType As String
    Dim tlSEE As SEE

    gConflictTableCheck = False
    gConflictPop
    slDateTime = gNow()
    lmGenDate = gDateValue(Format$(slDateTime, "ddddd"))
    lmGenTime = gTimeToLong(Format$(slDateTime, "ttttt"), False)
    
    slStartDate = slInStartDate
    slEndDate = slInEndDate
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    
    smTypeSource = slType
    ReDim Preserve tmConflictResults(0 To 0) As CONFLICTRESULTS

    llUpper = 1
    For llEvent = grdEvents.FixedRows To grdEvents.Rows - 1 Step 1
        If (Trim$(grdEvents.TextMatrix(llEvent, ilCols(1))) <> "") Then
            slEventType = Trim$(grdEvents.TextMatrix(llEvent, ilCols(1)))
            lmGridEventRow = llEvent
            ilSet = True
            If (ilCols(13) <> -1) Then
                If Trim$(grdEvents.TextMatrix(llEvent, ilCols(13))) = "N" Then
                    ilSet = False
                End If
            End If
            If ilSet Then
                ilPriAneCode = 0
                ilProtAneCode = 0
                ilBkupAneCode = 0
                slIgnoreConflict = Trim$(grdEvents.TextMatrix(llEvent, ilCols(14)))
                slBuses = Trim$(grdEvents.TextMatrix(llEvent, ilCols(6)))
                gParseCDFields slBuses, False, slBusNames()
                If (slIgnoreConflict <> "A") And (slIgnoreConflict <> "I") Then
                    slAudio = Trim$(grdEvents.TextMatrix(llEvent, ilCols(7)))
                    If slAudio <> "" Then
                        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                        '    If StrComp(slAudio, Trim$(tgCurrANE(ilANE).sName), vbTextCompare) = 0 Then
                            ilANE = gBinarySearchName(slAudio, tgCurrANE_Name())
                            If ilANE <> -1 Then
                                ilPriAneCode = tgCurrANE_Name(ilANE).iCode    'tgCurrANE(ilANE).iCode
                        '        Exit For
                            End If
                        'Next ilANE
                    End If
                    slAudio = Trim$(grdEvents.TextMatrix(llEvent, ilCols(8)))
                    If slAudio <> "" Then
                        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                        '    If StrComp(slAudio, Trim$(tgCurrANE(ilANE).sName), vbTextCompare) = 0 Then
                            ilANE = gBinarySearchName(slAudio, tgCurrANE_Name())
                            If ilANE <> -1 Then
                                ilProtAneCode = tgCurrANE_Name(ilANE).iCode 'tgCurrANE(ilANE).iCode
                        '        Exit For
                            End If
                        'Next ilANE
                    End If
                    slAudio = Trim$(grdEvents.TextMatrix(llEvent, ilCols(9)))
                    If slAudio <> "" Then
                        'For ilANE = 0 To UBound(tgCurrANE) - 1 Step 1
                        '    If StrComp(slAudio, Trim$(tgCurrANE(ilANE).sName), vbTextCompare) = 0 Then
                            ilANE = gBinarySearchName(slAudio, tgCurrANE_Name())
                            If ilANE <> -1 Then
                                ilBkupAneCode = tgCurrANE_Name(ilANE).iCode 'tgCurrANE(ilANE).iCode
                        '        Exit For
                            End If
                        'Next ilANE
                    End If
                End If
                slStr = grdEvents.TextMatrix(llEvent, ilCols(4))
                llOffsetStartTime = gStrLengthInTenthToLong(slStr)
                slStr = grdEvents.TextMatrix(llEvent, ilCols(5))
                llOffsetEndTime = llOffsetStartTime + gStrLengthInTenthToLong(slStr)  ' - 1
                If llOffsetEndTime < llOffsetStartTime Then
                    llOffsetEndTime = llOffsetStartTime
                End If
                
                If slType = "L" Then
                    slStr = Trim$(grdEvents.TextMatrix(llEvent, ilCols(2)))
                    slHours = gCreateHourStr(slStr)
                    slStr = Trim$(grdEvents.TextMatrix(llEvent, ilCols(3)))
                    slDays = gCreateDayStr(slStr)
                    slTestHours = slHours
                    For ilHour = 1 To 24 Step 1
                        If (Mid$(slTestHours, ilHour, 1) = "Y") Then
                            llStartTime = 36000 * (ilHour - 1) + llOffsetStartTime
                            llEndTime = 36000 * (ilHour - 1) + llOffsetEndTime
                            If slBuses <> "" Then
                                mConflictCEEBusRec ilCols(6), "S", slIgnoreConflict, slDays, slBusNames(), llStartDate, llEndDate, llStartTime, llEndTime
                            End If
                            mConflictCEEAudioRec slEventType, ilCols(7), "S", slDays, slBusNames(), ilPriAneCode, llStartDate, llEndDate, llStartTime, llEndTime
                            mConflictCEEAudioRec slEventType, ilCols(8), "S", slDays, slBusNames(), ilProtAneCode, llStartDate, llEndDate, llStartTime, llEndTime
                            mConflictCEEAudioRec slEventType, ilCols(9), "S", slDays, slBusNames(), ilBkupAneCode, llStartDate, llEndDate, llStartTime, llEndTime
                        End If
                    Next ilHour
                ElseIf slType = "S" Then
                    slDays = String(7, "N")
                    Select Case Weekday(slInStartDate)
                        Case vbMonday
                            Mid(slDays, 1, 1) = "Y"
                        Case vbTuesday
                            Mid(slDays, 2, 1) = "Y"
                        Case vbWednesday
                            Mid(slDays, 3, 1) = "Y"
                        Case vbThursday
                            Mid(slDays, 4, 1) = "Y"
                        Case vbFriday
                            Mid(slDays, 5, 1) = "Y"
                        Case vbSaturday
                            Mid(slDays, 6, 1) = "Y"
                        Case vbSunday
                            Mid(slDays, 7, 1) = "Y"
                    End Select
                    llStartTime = llOffsetStartTime
                    llEndTime = llOffsetEndTime
                    If slBuses <> "" Then
                        mConflictCEEBusRec ilCols(6), "S", slIgnoreConflict, slDays, slBusNames(), llStartDate, llEndDate, llStartTime, llEndTime
                    End If
                    mConflictCEEAudioRec slEventType, ilCols(7), "S", slDays, slBusNames(), ilPriAneCode, llStartDate, llEndDate, llStartTime, llEndTime
                    mConflictCEEAudioRec slEventType, ilCols(8), "S", slDays, slBusNames(), ilProtAneCode, llStartDate, llEndDate, llStartTime, llEndTime
                    mConflictCEEAudioRec slEventType, ilCols(9), "S", slDays, slBusNames(), ilBkupAneCode, llStartDate, llEndDate, llStartTime, llEndTime
                Else
                    slStr = Trim$(grdEvents.TextMatrix(llEvent, ilCols(2)))
                    slHours = gCreateHourStr(slStr)
                    slDays = String(7, "N")
                    Select Case Weekday(slInStartDate)
                        Case vbMonday
                            Mid(slDays, 1, 1) = "Y"
                        Case vbTuesday
                            Mid(slDays, 2, 1) = "Y"
                        Case vbWednesday
                            Mid(slDays, 3, 1) = "Y"
                        Case vbThursday
                            Mid(slDays, 4, 1) = "Y"
                        Case vbFriday
                            Mid(slDays, 5, 1) = "Y"
                        Case vbSaturday
                            Mid(slDays, 6, 1) = "Y"
                        Case vbSunday
                            Mid(slDays, 7, 1) = "Y"
                    End Select
                    ilStartHour = Hour(slTempStartTime)
                    If ilStartHour <> 0 Then
                        slTestHours = String(24, "N")
                        ilHour = ilStartHour
                        For ilLoop = 0 To 23 Step 1
                            Mid$(slTestHours, ilHour + 1, 1) = Mid$(slHours, ilLoop + 1, 1)
                            ilHour = ilHour + 1
                            If ilHour > 23 Then
                                Exit For
                            End If
                        Next ilLoop
                    Else
                        slTestHours = slHours
                    End If
                    For ilHour = 1 To 24 Step 1
                        If (Mid$(slTestHours, ilHour, 1) = "Y") Then
                            llStartTime = 36000 * (ilHour - 1) + llOffsetStartTime
                            llEndTime = 36000 * (ilHour - 1) + llOffsetEndTime
                            llStartTime = llStartTime + 10 * (gTimeToLong(slTempStartTime, False) Mod 3600)
                            llEndTime = llEndTime + 10 * (gTimeToLong(slTempStartTime, False) Mod 3600)
                            If slBuses <> "" Then
                                mConflictCEEBusRec ilCols(6), "S", slIgnoreConflict, slDays, slBusNames(), llStartDate, llEndDate, llStartTime, llEndTime
                            End If
                            mConflictCEEAudioRec slEventType, ilCols(7), "S", slDays, slBusNames(), ilPriAneCode, llStartDate, llEndDate, llStartTime, llEndTime
                            mConflictCEEAudioRec slEventType, ilCols(8), "S", slDays, slBusNames(), ilProtAneCode, llStartDate, llEndDate, llStartTime, llEndTime
                            mConflictCEEAudioRec slEventType, ilCols(9), "S", slDays, slBusNames(), ilBkupAneCode, llStartDate, llEndDate, llStartTime, llEndTime
                        End If
                    Next ilHour
                    
                End If
            End If
        End If
    Next llEvent
    'ilRet = gGetConflicts_CEE_CME(lmGenDate, lmGenTime, slType, "gConflictTableCheck", tmConflictResults())
    For ilLoop = 0 To UBound(tmConflictResults) - 1 Step 1
        llEvent = tmConflictResults(ilLoop).tCEE.lGridEventRow
        ilStartConflictIndex = Val(grdEvents.TextMatrix(llEvent, ilCols(0)))
        ilConflictIndex = UBound(tlConflictList)
        tlConflictList(ilConflictIndex).sType = tmConflictResults(ilLoop).tCME.sSource
        tlConflictList(ilConflictIndex).sStartDate = Format$(tmConflictResults(ilLoop).tCME.lStartDate, "ddddd")
        tlConflictList(ilConflictIndex).sEndDate = Format$(tmConflictResults(ilLoop).tCME.lEndDate, "ddddd")
        tlConflictList(ilConflictIndex).lIndex = -1
        tlConflictList(ilConflictIndex).iNextIndex = -1
        ilShowConflict = True
        If tmConflictResults(ilLoop).tCEE.iGridEventCol <> ilCols(6) Then
            If (tmConflictResults(ilLoop).tCEE.lStartTime = tmConflictResults(ilLoop).tCME.lStartTime) And (tmConflictResults(ilLoop).tCEE.lEndTime = tmConflictResults(ilLoop).tCME.lEndTime) Then
                If tmConflictResults(ilLoop).tCEE.iBdeCode = tmConflictResults(ilLoop).tCME.iBdeCode Then
                    If tmConflictResults(ilLoop).tCEE.iGridEventCol = ilCols(10) Then
                        slItemID1 = Trim$(grdEvents.TextMatrix(llEvent, ilCols(10)))
                        slItemID2 = Trim$(tmConflictResults(ilLoop).tCME.sItemID)
                        If StrComp(slItemID1, slItemID2, vbTextCompare) = 0 Then
                            ilShowConflict = False
                        End If
                    End If
                    If tmConflictResults(ilLoop).tCEE.iGridEventCol = ilCols(11) Then
                        slItemID1 = Trim$(grdEvents.TextMatrix(llEvent, ilCols(11)))
                        slItemID2 = Trim$(tmConflictResults(ilLoop).tCME.sItemID)
                        If StrComp(slItemID1, slItemID2, vbTextCompare) = 0 Then
                            ilShowConflict = False
                        End If
                    End If
                    If tmConflictResults(ilLoop).tCEE.iGridEventCol = ilCols(12) Then
                        slItemID1 = Trim$(grdEvents.TextMatrix(llEvent, ilCols(12)))
                        slItemID2 = Trim$(tmConflictResults(ilLoop).tCME.sItemID)
                        If StrComp(slItemID1, slItemID2, vbTextCompare) = 0 Then
                            ilShowConflict = False
                        End If
                    End If
                End If
            End If
        End If
        If ilShowConflict Then
            If Val(grdEvents.TextMatrix(llEvent, ilCols(15))) = tmConflictResults(ilLoop).tCME.lDeeCode Then
                ilShowConflict = False
            End If
        End If
        If (ilShowConflict) And (llOverlappedDHE > 0) Then
            If (tmConflictResults(ilLoop).tCME.sSource = "L") And (tmConflictResults(ilLoop).tCME.lSHEDHECode = llOverlappedDHE) Then
                ilShowConflict = False
            End If
        End If
        If (ilShowConflict) And (llDheCode > 0) Then
            If (tmConflictResults(ilLoop).tCME.sSource <> "S") And (tmConflictResults(ilLoop).tCME.lSHEDHECode = llDheCode) Then
                ilShowConflict = False
            End If
        End If
        If tmConflictResults(ilLoop).tCME.sSource = "S" Then
            ilRet = gGetRec_SEE_ScheduleEvent(tmConflictResults(ilLoop).tCME.lSeeCode, "gConflictTableCheck", tlSEE)
            If (tlSEE.sAction = "D") Or (tlSEE.sAction = "R") Then
                ilShowConflict = False
            Else
                If (llDheCode = tlSEE.lOrigDHECode) Or (llOverlappedDHE = tlSEE.lOrigDHECode) Then
                    ilShowConflict = False
                End If
            End If
        End If
        If ilShowConflict Then
            If tmConflictResults(ilLoop).tCME.sSource <> "S" Then
                tlConflictList(ilConflictIndex).lSheCode = 0
                tlConflictList(ilConflictIndex).lSeeCode = 0
                tlConflictList(ilConflictIndex).lDheCode = tmConflictResults(ilLoop).tCME.lSHEDHECode
                tlConflictList(ilConflictIndex).lDseCode = tmConflictResults(ilLoop).tCME.lDseCode
                tlConflictList(ilConflictIndex).lDeeCode = tmConflictResults(ilLoop).tCME.lDeeCode
            Else
                tlConflictList(ilConflictIndex).lSheCode = tmConflictResults(ilLoop).tCME.lSHEDHECode
                tlConflictList(ilConflictIndex).lDheCode = 0
                tlConflictList(ilConflictIndex).lDseCode = 0
                tlConflictList(ilConflictIndex).lDeeCode = 0
                tlConflictList(ilConflictIndex).lSeeCode = tmConflictResults(ilLoop).tCME.lSeeCode
            End If
            gConflictTableCheck = True
            grdEvents.Row = llEvent
            grdEvents.Col = tmConflictResults(ilLoop).tCEE.iGridEventCol
            grdEvents.CellForeColor = vbRed
            grdEvents.TextMatrix(llEvent, ilCols(0)) = Trim$(Str$(mAddConflict(ilStartConflictIndex, tlConflictList())))
        End If
    Next ilLoop
    'ilRet = gPutDelete_CEE_Conflict_Events(lmGenDate, lmGenTime, "gConflictTableCheck")
End Function


