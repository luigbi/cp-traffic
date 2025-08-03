Attribute VB_Name = "RPTLinkSubs"

Option Explicit
Option Compare Text

Dim tmSsf As SSF
Dim tmSsfSrchKey As SSFKEY0
Dim tmCTSSF As SSF
Dim imSsfRecLen As Integer

Dim tmVlf As VLF
Dim imVlfRecLen As Integer
Dim tmVlfSrchKey1 As VLFKEY1

Dim tmSdf As SDF
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey3 As LONGKEY0

Dim tmAvailTest As AVAILSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
'
'                       gGetAiringSpots -Obtain spots from selling vehicle given the airing vehicle code
'                       Find the Airing vehicle links; then get its closes SSF.
'                       Cycle thru SSF for breaks only.  For each break, find the associated
'                       selling SSF and matching link.  Then obtain the SDF to update into
'                       array of spots with the time of the airing link.
'                       <input> ilVefCode - airing vehicle internal code
'                                    slSDate - earliest date to search
'                                slEDate - latest date to search
'                    llStartTime - earliest time filter
'                                        llLatestTime - latest time filter
'                       <output> tlAiringSDF() array of airing spots generated from selling SDF
Public Sub gGetAiringSpots(hlVlf As Integer, hlSsf As Integer, hlCTSsf As Integer, hlSdf As Integer, ilVefCode As Integer, slSDate As String, slEDate As String, llStartTime As Long, llEndTime As Long, tlAiringSDF() As SPOTTYPESORT)
    Dim ilType As Integer
    Dim llSsfReclec As Integer
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim slDate As String
    Dim ilDay As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim slDay As String
    Dim ilVlfDate0 As Integer
    Dim ilVlfDate1 As Integer
    Dim ilRet As Integer
    Dim ilTerminated As Integer
    Dim ilWithinTime As Integer
    Dim ilEvt As Integer
    Dim ilSIndex As Integer
    Dim ilSpot As Integer
    Dim llTime As Long
    Dim ilSSFDate(0 To 1) As Integer
    Dim slSSFDate As String
    
    ReDim tlAiringSDF(0 To 0) As SPOTTYPESORT
    ilType = 0              'airing doesnt have game #s
    imSsfRecLen = Len(tmSsf)  'Get and save SSF record length
    On Error GoTo 0
    llSDate = gDateValue(slSDate)
    llEDate = gDateValue(slEDate)
    For llDate = llSDate To llEDate Step 1  'Process next date for avails that map back one day
        slDate = Format$(llDate, "m/d/yy")
        ilDay = gWeekDayStr(slDate)
        gPackDate slDate, ilLogDate0, ilLogDate1

        If (ilDay >= 0) And (ilDay <= 4) Then
            slDay = "0"
        ElseIf ilDay = 5 Then
            slDay = "6"
        Else
            slDay = "7"
        End If
        imVlfRecLen = Len(tmVlf)
        ilVlfDate0 = 0
        ilVlfDate1 = 0
        tmVlfSrchKey1.iAirCode = ilVefCode
        tmVlfSrchKey1.iAirDay = Val(slDay)
        tmVlfSrchKey1.iEffDate(0) = ilLogDate0
        tmVlfSrchKey1.iEffDate(1) = ilLogDate1
        tmVlfSrchKey1.iAirTime(0) = 0
        tmVlfSrchKey1.iAirTime(1) = 6144    '24*256
        tmVlfSrchKey1.iAirPosNo = 32000
        tmVlfSrchKey1.iAirSeq = 32000
        'AIRING VLF
        ilRet = btrGetLessOrEqual(hlVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVefCode)
            ilTerminated = False
            'Check for CBS
            If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                    ilTerminated = True
                End If
            End If
            If (tmVlf.sStatus <> "P") And (tmVlf.iAirDay = Val(slDay)) And (Not ilTerminated) Then
                ilVlfDate0 = tmVlf.iEffDate(0)
                ilVlfDate1 = tmVlf.iEffDate(1)
                Exit Do
            End If
            ilRet = btrGetPrevious(hlVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop

        ilDay = gWeekDayStr(slDate)
        gPackDate slDate, ilLogDate0, ilLogDate1
        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
        ilSSFDate(0) = ilLogDate0
        ilSSFDate(1) = ilLogDate1
        tmSsfSrchKey.iType = ilType
        tmSsfSrchKey.iVefCode = ilVefCode
        tmSsfSrchKey.iDate(0) = ilSSFDate(0)
        tmSsfSrchKey.iDate(1) = ilSSFDate(1)
        tmSsfSrchKey.iStartTime(0) = 0
        tmSsfSrchKey.iStartTime(1) = 0
        'AIRING SSF
        ilRet = gSSFGetEqual(hlSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iType <> ilType) Or (tmSsf.iVefCode <> ilVefCode) Or (tmSsf.iDate(0) <> ilSSFDate(0)) Or (tmSsf.iDate(1) <> ilSSFDate(1)) Then
            ilSSFDate(0) = 0
            ilSSFDate(1) = 0
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            tmSsfSrchKey.iType = ilType
            tmSsfSrchKey.iVefCode = ilVefCode
            tmSsfSrchKey.iDate(0) = ilLogDate0
            tmSsfSrchKey.iDate(1) = ilLogDate1
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 6144   '24*256
            ilRet = gSSFGetLessOrEqual(hlSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode)
                gUnpackDate tmSsf.iDate(0), tmSsf.iDate(1), slSSFDate
                If (ilDay = gWeekDayStr(slSSFDate)) And (tmSsf.iStartTime(0) = 0) And (tmSsf.iStartTime(1) = 0) Then
                    ilSSFDate(0) = tmSsf.iDate(0)
                    ilSSFDate(1) = tmSsf.iDate(1)
                    Exit Do
                End If
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetPrevious(hlSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
        DoEvents
        If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode) Then
            gUnpackDate ilSSFDate(0), ilSSFDate(1), slSSFDate
            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode) And (tmSsf.iDate(0) = ilSSFDate(0)) And (tmSsf.iDate(1) = ilSSFDate(1))
                ilEvt = 1
                Do While ilEvt <= tmSsf.iCount
                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Avail
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                        If llTime >= llStartTime And llTime < llEndTime Then      'if not within the time requests, no need to look for the associated selling link
                            tmVlfSrchKey1.iAirCode = ilVefCode
                            tmVlfSrchKey1.iAirDay = Val(slDay)
                            tmVlfSrchKey1.iEffDate(0) = ilVlfDate0
                            tmVlfSrchKey1.iEffDate(1) = ilVlfDate1
                            tmVlfSrchKey1.iAirTime(0) = tmAvail.iTime(0)
                            tmVlfSrchKey1.iAirTime(1) = tmAvail.iTime(1)
                            tmVlfSrchKey1.iAirPosNo = 0
                            tmVlfSrchKey1.iAirSeq = 1
                            'find associated Selling link
                            ilRet = btrGetGreaterOrEqual(hlVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                            Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVefCode) And (tmVlf.iAirDay = Val(slDay)) And (tmVlf.iEffDate(0) = ilVlfDate0) And (tmVlf.iEffDate(1) = ilVlfDate1) And (tmVlf.iAirTime(0) = tmAvail.iTime(0)) And (tmVlf.iAirTime(1) = tmAvail.iTime(1))
                                ilTerminated = False
                                If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                    If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                        ilTerminated = True
                                    End If
                                End If
                                If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                    If (tmCTSSF.iType <> ilType) Or (tmCTSSF.iVefCode <> tmVlf.iSellCode) Or (tmCTSSF.iDate(0) <> ilLogDate0) Or (tmCTSSF.iDate(1) <> ilLogDate1) Then
                                        imSsfRecLen = Len(tmCTSSF) 'Max size of variable length record
                                        'tmSSFSrchKey.sType = slType
                                        tmSsfSrchKey.iType = ilType
                                        tmSsfSrchKey.iVefCode = tmVlf.iSellCode
                                        tmSsfSrchKey.iDate(0) = ilLogDate0
                                        tmSsfSrchKey.iDate(1) = ilLogDate1
                                        tmSsfSrchKey.iStartTime(0) = 0
                                        tmSsfSrchKey.iStartTime(1) = 0
                                        ilRet = gSSFGetEqual(hlCTSsf, tmCTSSF, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
    
                                    End If
                                    Do While (ilRet = BTRV_ERR_NONE) And (tmCTSSF.iType = ilType) And (tmCTSSF.iVefCode = tmVlf.iSellCode) And (tmCTSSF.iDate(0) = ilLogDate0) And (tmCTSSF.iDate(1) = ilLogDate1)
                                        For ilSIndex = 1 To tmCTSSF.iCount Step 1
                                            tmAvailTest = tmCTSSF.tPas(ADJSSFPASBZ + ilSIndex)
                                            If ((tmAvailTest.iRecType >= 2) And (tmAvailTest.iRecType <= 9)) Then
                                                If (tmVlf.iSellTime(0) = tmAvailTest.iTime(0)) And (tmVlf.iSellTime(1) = tmAvailTest.iTime(1)) Then     'compare the airing time that is being processing to the correct selling ssf entry
                                                    'filter time of spots based on user selectivity
                                                    For ilSpot = 1 To tmAvailTest.iNoSpotsThis Step 1
                                                       LSet tmSpot = tmCTSSF.tPas(ADJSSFPASBZ + ilSpot + ilSIndex)
                                                        tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                        ilRet = btrGetEqual(hlSdf, tmSdf, Len(tmSdf), tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            tlAiringSDF(UBound(tlAiringSDF)).tSdf = tmSdf
                                                            tlAiringSDF(UBound(tlAiringSDF)).tSdf.iTime(0) = tmAvail.iTime(0)      'airing time
                                                            tlAiringSDF(UBound(tlAiringSDF)).tSdf.iTime(1) = tmAvail.iTime(1)
                                                            tlAiringSDF(UBound(tlAiringSDF)).iVefCode = tmVlf.iAirCode                 'airing vehicle
                                                            ReDim Preserve tlAiringSDF(0 To UBound(tlAiringSDF) + 1) As SPOTTYPESORT
                                                        End If
                                                    Next ilSpot
                                                    Exit Do
                                                End If
                                            End If
                                        Next ilSIndex
                                        imSsfRecLen = Len(tmCTSSF) 'Max size of variable length record
                                        ilRet = gSSFGetNext(hlCTSsf, tmCTSSF, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                    Loop
                                End If
                                'Get next selling
                                ilRet = btrGetNext(hlVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                        End If              'filter out time
                    End If
                    ilEvt = ilEvt + 1
                Loop
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hlSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
    Next llDate
    Exit Sub
End Sub
