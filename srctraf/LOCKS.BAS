Attribute VB_Name = "LOCKSSUBS"

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Schedule.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the schedule subs and functions

'   Type     Subtype       Definition
'   C        C             Altering a Contract
'   C        S             Contract Scheduling
'
'   S        C             Spot Scheduling
'   S        P             Posting spots
'   S        S             Booking Spot
'
'   I        F             Generating Final Invoices
'   I        U             Running Undo
'
'   M                      Defining Copy Inventory
'
'   R        R             Running Reconcile
'   R        Z             Running Zero-Purge
'   R        P             In Collection
'

'Posting spots and generating Final invoices disallowed
'Only one person can change the same contract
'Can't schedule a contract while some one is altering the contract and visa verse
'Only one person can book a spot within same vehicle and day/game
'Only one person can work within same vehicle and  day/game
'Only one person can run final invoices
'Can't post while final invoices are running and visa verse
'Only one person can be in Collections
'Only one person can run reconcile
'Only one person can run Zero-Purge
'Only one person can change same copy inventory




Option Explicit
Option Compare Text
Dim tmSsf As SSF                'SSF record image
'Dim tmSsfOld As SSF
Dim imSsfRecLen As Integer
Dim tmAvail As AVAILSS

Dim tmRlf As RLF
Dim imRlfRecLen As Integer
Dim tmRlfSrchKey0 As LONGKEY0
Dim tmRlfSrchKey1 As RLFKEY1     'RLF key record image

Dim hmAlf As Integer
Dim tmAlf As ALF
Dim tmAlfSrchkey1 As ALFKEY1
Dim tmAlfSrchKey2 As ALFKEY2
Dim imAlfRecLen As Integer



'*******************************************************
'*                                                     *
'*      Procedure Name:gInschSpots                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Unschedule spots               *
'*                                                     *
'*******************************************************
Sub gSetLockStatus(ilVehCode As Integer, ilLockAvail As Integer, ilLockSpot As Integer, slStartDate As String, slEndDate As String, ilGameNo As Integer, slStartTime As String, slEndTime As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilType                                                                                *
'******************************************************************************************

'
'   gUnschSpots ilVehCode, ilLockAvail, ilLockSpot, slStartDate, slEndDate, slStartTime, slEndTime
'
'   Where:
'       ilVehCode (I)-Vehicle code number
'       ilLockAvail(I)- Avail: 0=Unlock; 1=Lock; -1=Leave
'       ilLockSpot(I)- Spot: 0=Unlock; 1=Lock; -1=Leave
'       slStartDate (I)- Start Date that that spots are to be removed
'       slEndDate (I)- End Date that that spots are to be removed
'       ilGameNo(I) - Game Number
'       slStartTime (I)- Start Time (included)
'       slEndTime (I)- End time (not included)
'
    Dim ilRet As Integer
    Dim ilRPRet As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim llSsfRecPos As Long
    Dim llDate As Long
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slDate As String
    Dim ilDay As Integer
    Dim slTime As String
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llAvailTime As Long
    Dim ilLoop As Integer
    Dim llRecPos As Long
    Dim llAlfStartTime As Long
    Dim llAlfEndTime As Long
    Dim ilAlfChgd As Integer
    Dim ilFound As Integer
    Dim tlAlf As ALF

    'Spot summary
    Dim hlSsf As Integer        'Spot summary file handle
    llStartTime = CLng(gTimeToCurrency(slStartTime, False))
    llEndTime = CLng(gTimeToCurrency(slEndTime, True)) - 1
    hlSsf = CBtrvTable(TWOHANDLES)        'Create SSF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlSsf)
        btrDestroy hlSsf
        Exit Sub
    End If
    hmAlf = CBtrvTable(TWOHANDLES)        'Create SSF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmAlf, "", sgDBPath & "Alf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlSsf)
        btrDestroy hlSsf
        ilRet = btrClose(hmAlf)
        btrDestroy hmAlf
        Exit Sub
    End If
    imAlfRecLen = Len(tmAlf)
    If ilGameNo <= 0 Then
        llStartDate = gDateValue(slStartDate)
        llEndDate = gDateValue(slEndDate)
    Else
        If slStartDate = "" Then
            llStartDate = 0
            llEndDate = 0
        Else
            llStartDate = gDateValue(slStartDate)
            llEndDate = llStartDate
        End If
    End If
    For llDate = llStartDate To llEndDate Step 1
        tmSsf.iType = ilGameNo
        tmSsf.iVefCode = ilVehCode
        If ilGameNo <= 0 Then
            slDate = Format$(llDate, "m/d/yy")
            ilDay = gWeekDayStr(slDate)
            gPackDate slDate, ilLogDate0, ilLogDate1
        Else
            If slStartDate = "" Then
                ilLogDate0 = 0
                ilLogDate1 = 0
            Else
                slDate = Format$(llDate, "m/d/yy")
                ilDay = gWeekDayStr(slDate)
                gPackDate slDate, ilLogDate0, ilLogDate1
            End If
        End If
        tmSsf.iDate(0) = ilLogDate0
        tmSsf.iDate(1) = ilLogDate1
        tmSsf.iStartTime(0) = 0
        tmSsf.iStartTime(1) = 0
        imSsfRecLen = Len(tmSsf)
        ilRet = gGetByKeyForUpdateSSF(hlSsf, tmSsf)
        Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = ilVehCode)   'And (((tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1) And (tmSsf.iType = 0) And (ilGameNo = 0)) Or ((tmSsf.iType = ilGameNo) And (ilGameNo <> 0)))
            If (((tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1) And (tmSsf.iType = 0) And (ilGameNo = 0)) Or ((tmSsf.iType = ilGameNo) And (ilGameNo <> 0))) Then
                ilRPRet = gSSFGetPosition(hlSsf, llSsfRecPos)
                Do
                    For ilLoop = 1 To tmSsf.iCount Step 1
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilLoop)
                        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Contract Avail subrecord
                            'Test time-
                            gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                            llAvailTime = CLng(gTimeToCurrency(slTime, False))
                            If (llAvailTime >= llStartTime) And (llAvailTime <= llEndTime) Then
                                If ilLockAvail = 0 Then 'Unlock
                                    tmAvail.iAvInfo = tmAvail.iAvInfo And (Not SSLOCK)
                                ElseIf ilLockAvail = 1 Then
                                    tmAvail.iAvInfo = tmAvail.iAvInfo Or SSLOCK
                                End If
                                If ilLockSpot = 0 Then
                                    tmAvail.iAvInfo = tmAvail.iAvInfo And (Not SSLOCKSPOT)
                                ElseIf ilLockSpot = 1 Then
                                    tmAvail.iAvInfo = tmAvail.iAvInfo Or SSLOCKSPOT
                                End If
                                tmSsf.tPas(ADJSSFPASBZ + ilLoop) = tmAvail
                            End If
                        End If
                    Next ilLoop
                    imSsfRecLen = igSSFBaseLen + tmSsf.iCount * Len(tmAvail)
                    ilRet = gSSFUpdate(hlSsf, tmSsf, imSsfRecLen)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        imSsfRecLen = Len(tmSsf)
                        ilRPRet = gSSFGetDirect(hlSsf, tmSsf, imSsfRecLen, llSsfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        ilRPRet = gGetByKeyForUpdateSSF(hlSsf, tmSsf)
                        If ilRPRet <> BTRV_ERR_NONE Then
                            ilRet = btrClose(hlSsf)
                            btrDestroy hlSsf
                            Exit Sub
                        End If
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
            Else
                If (ilGameNo = 0) Then
                    If (tmSsf.iDate(0) <> ilLogDate0) Or (tmSsf.iDate(1) <> ilLogDate1) Or (tmSsf.iType <> 0) Then
                        Exit Do
                    End If
                Else
                    If slStartDate <> "" Then
                        If (tmSsf.iDate(0) <> ilLogDate0) Or (tmSsf.iDate(1) <> ilLogDate1) Then
                            Exit Do
                        End If
                    End If
                End If
            End If
            Do
                imSsfRecLen = Len(tmSsf)
                ilRet = gSSFGetNext(hlSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet = BTRV_ERR_END_OF_FILE Then
                    Exit Do
                End If
                ilRPRet = gSSFGetPosition(hlSsf, llRecPos)
            Loop While llRecPos = llSsfRecPos
        Loop
    Next llDate
    If ilGameNo <= 0 Then
        llStartDate = gDateValue(slStartDate)
        llEndDate = gDateValue(slEndDate)
    Else
        llStartDate = 0
        llEndDate = 0
    End If
    For llDate = llStartDate To llEndDate Step 1
        If (ilLockAvail = 0) Or (ilLockSpot = 0) Then
            'Remove records
            If ilGameNo <= 0 Then
                tmAlfSrchkey1.iVefCode = ilVehCode
                slDate = Format$(llDate, "m/d/yy")
                gPackDate slDate, ilLogDate0, ilLogDate1
                tmAlfSrchkey1.iDate(0) = ilLogDate0
                tmAlfSrchkey1.iDate(1) = ilLogDate1
                ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchkey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
            Else
                tmAlfSrchKey2.iVefCode = ilVehCode
                tmAlfSrchKey2.iGameNo = ilGameNo
                ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
            End If
            Do While (ilRet = BTRV_ERR_NONE) And (tmAlf.iVefCode = ilVehCode) And (((tmAlf.iDate(0) = ilLogDate0) And (tmAlf.iDate(1) = ilLogDate1) And (ilGameNo = 0)) Or ((tmAlf.iGameNo = ilGameNo) And (ilGameNo <> 0)))
                ilAlfChgd = False
                If ((tmAlf.sLockType = "A") And (ilLockAvail = 0)) Or ((tmAlf.sLockType = "S") And (ilLockSpot = 0)) Then
                    'Times intersect
                    gUnpackTimeLong tmAlf.iStartTime(0), tmAlf.iStartTime(1), False, llAlfStartTime
                    gUnpackTimeLong tmAlf.iEndTime(0), tmAlf.iEndTime(1), True, llAlfEndTime
                    llAlfEndTime = llAlfEndTime - 1
                    If (llStartTime <= llAlfStartTime) And (llEndTime >= llAlfEndTime) Then
                        ilAlfChgd = True
                        ilRet = btrDelete(hmAlf)
                    ElseIf (llEndTime >= llAlfStartTime) And (llStartTime <= llAlfEndTime) Then
                        tmAlf.iUrfCode = tgUrf(0).iCode
                        If (llStartTime <= llAlfStartTime) And (llEndTime < llAlfEndTime) Then
                            ilAlfChgd = True
                            gPackTimeLong llEndTime + 1, tmAlf.iStartTime(0), tmAlf.iStartTime(1)
                            ilRet = btrUpdate(hmAlf, tmAlf, imAlfRecLen)
                        ElseIf (llEndTime >= llAlfEndTime) And (llStartTime > llAlfStartTime) Then
                            ilAlfChgd = True
                            gPackTimeLong llStartTime, tmAlf.iEndTime(0), tmAlf.iEndTime(1)
                            ilRet = btrUpdate(hmAlf, tmAlf, imAlfRecLen)
                        ElseIf (llStartTime > llAlfStartTime) And (llEndTime < llAlfEndTime) Then
                            ilAlfChgd = True
                            tlAlf = tmAlf
                            gPackTimeLong llStartTime, tmAlf.iEndTime(0), tmAlf.iEndTime(1)
                            ilRet = btrUpdate(hmAlf, tmAlf, imAlfRecLen)
                            tmAlf = tlAlf
                            tmAlf.lCode = 0
                            gPackTimeLong llEndTime + 1, tmAlf.iStartTime(0), tmAlf.iStartTime(1)
                            ilRet = btrInsert(hmAlf, tmAlf, imAlfRecLen, INDEXKEY2)
                            If ilRet <> BTRV_ERR_NONE Then
                                If ilRet = 30002 Then
                                    ilRet = csiHandleValue(0, 7)
                                End If
                                gMsgBox "gSetLockStatus: Insert ALF Error = " & ilRet, -1, ""
                            End If
                        End If
                    End If
                End If
                If ilAlfChgd Then
                    If ilGameNo <= 0 Then
                        tmAlfSrchkey1.iVefCode = ilVehCode
                        slDate = Format$(llDate, "m/d/yy")
                        gPackDate slDate, ilLogDate0, ilLogDate1
                        tmAlfSrchkey1.iDate(0) = ilLogDate0
                        tmAlfSrchkey1.iDate(1) = ilLogDate1
                        ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchkey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
                    Else
                        tmAlfSrchKey2.iVefCode = ilVehCode
                        tmAlfSrchKey2.iGameNo = ilGameNo
                        ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
                    End If
                Else
                    If ilGameNo <= 0 Then
                        ilRet = btrGetNext(hmAlf, tmAlf, imAlfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    Else
                        ilRet = btrGetNext(hmAlf, tmAlf, imAlfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    End If
                End If
            Loop
        End If
        If (ilLockAvail = 1) Or (ilLockSpot = 1) Then
            'Add Records
            If ilGameNo <= 0 Then
                tmAlfSrchkey1.iVefCode = ilVehCode
                slDate = Format$(llDate, "m/d/yy")
                gPackDate slDate, ilLogDate0, ilLogDate1
                tmAlfSrchkey1.iDate(0) = ilLogDate0
                tmAlfSrchkey1.iDate(1) = ilLogDate1
                ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchkey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
            Else
                tmAlfSrchKey2.iVefCode = ilVehCode
                tmAlfSrchKey2.iGameNo = ilGameNo
                ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
            End If
            ilFound = False
            Do While (ilRet = BTRV_ERR_NONE) And (tmAlf.iVefCode = ilVehCode) And ((tmAlf.iDate(0) = ilLogDate0) And (tmAlf.iDate(1) = ilLogDate1) And (ilGameNo = 0)) Or ((tmAlf.iGameNo = ilGameNo) And (ilGameNo <> 0))
                If ((tmAlf.sLockType = "A") And (ilLockAvail = 1)) Or ((tmAlf.sLockType = "S") And (ilLockSpot = 1)) Then
                    'Times intersect
                    gUnpackTimeLong tmAlf.iStartTime(0), tmAlf.iStartTime(1), False, llAlfStartTime
                    gUnpackTimeLong tmAlf.iEndTime(0), tmAlf.iEndTime(1), True, llAlfEndTime
                    llAlfEndTime = llAlfEndTime - 1
                    'If (llAlfStartTime >= llStartTime) And (llAlfEndTime <= llEndTime) Then
                    If (llStartTime >= llAlfStartTime) And (llEndTime <= llAlfEndTime) Then
                        ilFound = True
                        Exit Do
                    End If
                End If
                If ilGameNo <= 0 Then
                    ilRet = btrGetNext(hmAlf, tmAlf, imAlfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                Else
                    ilRet = btrGetNext(hmAlf, tmAlf, imAlfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                End If
            Loop
            If Not ilFound Then
                If ilGameNo <= 0 Then
                    tmAlfSrchkey1.iVefCode = ilVehCode
                    slDate = Format$(llDate, "m/d/yy")
                    gPackDate slDate, ilLogDate0, ilLogDate1
                    tmAlfSrchkey1.iDate(0) = ilLogDate0
                    tmAlfSrchkey1.iDate(1) = ilLogDate1
                    ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchkey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
                Else
                    tmAlfSrchKey2.iVefCode = ilVehCode
                    tmAlfSrchKey2.iGameNo = ilGameNo
                    ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
                End If
                ilFound = False
                Do While (ilRet = BTRV_ERR_NONE) And (tmAlf.iVefCode = ilVehCode) And (((tmAlf.iDate(0) = ilLogDate0) And (tmAlf.iDate(1) = ilLogDate1) And (ilGameNo = 0)) Or ((tmAlf.iGameNo = ilGameNo) And (ilGameNo <> 0)))
                    ilAlfChgd = False
                    If ((tmAlf.sLockType = "A") And (ilLockAvail = 1)) Or ((tmAlf.sLockType = "S") And (ilLockSpot = 1)) Then
                        'Times intersect
                        gUnpackTimeLong tmAlf.iStartTime(0), tmAlf.iStartTime(1), False, llAlfStartTime
                        gUnpackTimeLong tmAlf.iEndTime(0), tmAlf.iEndTime(1), True, llAlfEndTime
                        llAlfEndTime = llAlfEndTime - 1
                        If (llStartTime <= llAlfStartTime) And (llEndTime >= llAlfEndTime) Then
                            ilAlfChgd = True
                            ilRet = btrDelete(hmAlf)
                        ElseIf (llEndTime >= llAlfStartTime) And (llStartTime <= llAlfEndTime) Then
                            If (llStartTime <= llAlfStartTime) And (llEndTime < llAlfEndTime) Then
                                ilAlfChgd = True
                                gPackTimeLong llEndTime + 1, tmAlf.iStartTime(0), tmAlf.iStartTime(1)
                                ilRet = btrUpdate(hmAlf, tmAlf, imAlfRecLen)
                            ElseIf (llEndTime >= llAlfEndTime) And (llStartTime > llAlfStartTime) Then
                                ilAlfChgd = True
                                gPackTimeLong llStartTime, tmAlf.iEndTime(0), tmAlf.iEndTime(1)
                                ilRet = btrUpdate(hmAlf, tmAlf, imAlfRecLen)
                            ElseIf (llStartTime > llAlfStartTime) And (llEndTime < llAlfEndTime) Then
                                ilAlfChgd = True
                                tlAlf = tmAlf
                                gPackTimeLong llStartTime, tmAlf.iEndTime(0), tmAlf.iEndTime(1)
                                ilRet = btrUpdate(hmAlf, tmAlf, imAlfRecLen)
                                tmAlf = tlAlf
                                tmAlf.lCode = 0
                                gPackTimeLong llEndTime + 1, tmAlf.iStartTime(0), tmAlf.iStartTime(1)
                                ilRet = btrInsert(hmAlf, tmAlf, imAlfRecLen, INDEXKEY2)
                                If ilRet <> BTRV_ERR_NONE Then
                                    If ilRet = 30002 Then
                                        ilRet = csiHandleValue(0, 7)
                                    End If
                                    gMsgBox "gSetLockStatus: Insert ALF Error = " & ilRet, -1, ""
                                End If
                            End If
                        End If
                    End If
                    If ilAlfChgd Then
                        If ilGameNo <= 0 Then
                            tmAlfSrchkey1.iVefCode = ilVehCode
                            slDate = Format$(llDate, "m/d/yy")
                            gPackDate slDate, ilLogDate0, ilLogDate1
                            tmAlfSrchkey1.iDate(0) = ilLogDate0
                            tmAlfSrchkey1.iDate(1) = ilLogDate1
                            ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchkey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
                        Else
                            tmAlfSrchKey2.iVefCode = ilVehCode
                            tmAlfSrchKey2.iGameNo = ilGameNo
                            ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
                        End If
                    Else
                        If ilGameNo <= 0 Then
                            ilRet = btrGetNext(hmAlf, tmAlf, imAlfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                        Else
                            ilRet = btrGetNext(hmAlf, tmAlf, imAlfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                        End If
                    End If
                Loop
                tmAlf.iVefCode = ilVehCode
                If ilGameNo = 0 Then
                    slDate = Format$(llDate, "m/d/yy")
                    gPackDate slDate, tmAlf.iDate(0), tmAlf.iDate(1)
                    tmAlf.iGameNo = 0
                Else
                    gPackDate "", tmAlf.iDate(0), tmAlf.iDate(1)
                    tmAlf.iGameNo = ilGameNo
                End If
                gPackTime slStartTime, tmAlf.iStartTime(0), tmAlf.iStartTime(1)
                gPackTime slEndTime, tmAlf.iEndTime(0), tmAlf.iEndTime(1)
                tmAlf.iUrfCode = tgUrf(0).iCode
                tmAlf.sUnused = ""
                ilRet = BTRV_ERR_NONE
                If ilLockAvail = 1 Then
                    tmAlf.lCode = 0
                    tmAlf.sLockType = "A"
                    ilRet = btrInsert(hmAlf, tmAlf, imAlfRecLen, INDEXKEY0)
                End If
                If ilLockSpot = 1 Then
                    tmAlf.lCode = 0
                    tmAlf.sLockType = "S"
                    ilRet = btrInsert(hmAlf, tmAlf, imAlfRecLen, INDEXKEY0)
                End If
                If ilRet <> BTRV_ERR_NONE Then
                    If ilRet = 30002 Then
                        ilRet = csiHandleValue(0, 7)
                    End If
                    gMsgBox "gSetLockStatus: Insert ALF Error = " & ilRet, -1, ""
                End If
            End If
        End If
    Next llDate
    ilRet = btrClose(hmAlf)
    btrDestroy hmAlf
    ilRet = btrClose(hlSsf)
    btrDestroy hlSsf
    Exit Sub
End Sub

'
'       12-14-04    Change return value definition from integer to long
Public Function gCreateLockRec(hlRlf As Integer, slType As String, slSubType As String, llRecCode As Long, ilRetryFlag As Integer, slUserNameWithLock As String) As Long
    Dim slNowDate As String
    Dim slNowTime As String
    Dim ilRet As Integer
    Dim llDate As Long
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim ilForceRetry As Integer
    Dim tlRlf As RLF

    If hlRlf <= 0 Then
        gCreateLockRec = 0
        Exit Function
    End If
    'Create Lock
    imRlfRecLen = Len(tmRlf)
    ilCount = 0
    Do
        ilForceRetry = False
        slUserNameWithLock = ""
        tmRlf.lCode = 0
        tmRlf.iUrfCode = tgUrf(0).iCode
        tmRlf.sType = slType
        tmRlf.sSubType = slSubType
        tmRlf.lRecCode = llRecCode
        slNowDate = Format$(gNow(), "m/d/yy")
        slNowTime = Format$(gNow(), "h:mm:ssAM/PM")
        gPackDate slNowDate, tmRlf.iEnteredDate(0), tmRlf.iEnteredDate(1)
        gPackTime slNowTime, tmRlf.iEnteredTime(0), tmRlf.iEnteredTime(1)
        tmRlf.sUnused = ""
        ilRet = btrInsert(hlRlf, tmRlf, imRlfRecLen, INDEXKEY0)
        If (ilRet <> BTRV_ERR_NONE) Then
            If ilRet = 30002 Then
                ilRet = csiHandleValue(0, 7)
            End If
            If (ilRet <> BTRV_ERR_DUPLICATE_KEY) Then
                gMsgBox "gCreateLockRec: Insert RLF Error = " & ilRet, -1, ""
            End If
        End If
        If ilRet = BTRV_ERR_DUPLICATE_KEY Then
            'Check if longer then Day for lock, if so remove it
            tmRlfSrchKey1.sType = slType
            tmRlfSrchKey1.lRecCode = llRecCode
            ilRet = btrGetEqual(hlRlf, tlRlf, imRlfRecLen, tmRlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                gUnpackDateLong tlRlf.iEnteredDate(0), tlRlf.iEnteredDate(1), llDate
                '6/30/06:  Remove checking for same user.  This is being removed to aviod the problem
                '          where the same user is running schedule on two terminals
                'If (llDate < gDateValue(slNowDate)) Or (tlRlf.iurfCode = tgUrf(0).iCode) Then
                If (llDate < gDateValue(slNowDate)) Then
                    ilRet = btrDelete(hlRlf)
                    If ilRet = BTRV_ERR_NONE Then
                        ilForceRetry = True
                    End If
                Else
                    For ilLoop = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
                        If tgPopUrf(ilLoop).iCode = tlRlf.iUrfCode Then
                            If Trim$(tgPopUrf(ilLoop).sRept) <> "" Then
                                slUserNameWithLock = Trim$(tgPopUrf(ilLoop).sRept)
                            Else
                                slUserNameWithLock = Trim$(tgPopUrf(ilLoop).sName)
                            End If
                            Exit For
                        End If
                    Next ilLoop
                    If Not ilRetryFlag Then
                        gCreateLockRec = 0
                        Exit Function
                    End If
                End If
            End If
        ElseIf ilRet = BTRV_ERR_DEADLOCK_DETECT Then
            ilForceRetry = True
        ElseIf ilRet = BTRV_ERR_NONE Then
            gCreateLockRec = tmRlf.lCode
            Exit Function
        End If
        If Not ilForceRetry Then
            If Not ilRetryFlag Then
                gCreateLockRec = 0
                Exit Function
            End If
        End If
        Sleep 100
        ilCount = ilCount + 1
    Loop While ilCount < 20
    gCreateLockRec = 0
    Exit Function
End Function

Public Function gDeleteLockRec_ByRlfCode(hlRlf As Integer, llRlfCode As Long) As Integer
    Dim ilRet As Integer

    If (llRlfCode > 0) And (hlRlf > 0) Then
        imRlfRecLen = Len(tmRlf)
        tmRlfSrchKey0.lCode = llRlfCode
        ilRet = btrGetEqual(hlRlf, tmRlf, imRlfRecLen, tmRlfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrDelete(hlRlf)
            If ilRet = BTRV_ERR_NONE Then
                gDeleteLockRec_ByRlfCode = True
                llRlfCode = 0
            Else
                gDeleteLockRec_ByRlfCode = False
            End If
        Else
            gDeleteLockRec_ByRlfCode = False
        End If
    Else
        gDeleteLockRec_ByRlfCode = True
        llRlfCode = 0
    End If
End Function

Public Function gDeleteLockRec_ByType(hlRlf As Integer, slType As String, llRecCode As Long) As Integer 'VBC NR
    Dim ilRet As Integer 'VBC NR

    If hlRlf <= 0 Then 'VBC NR
        gDeleteLockRec_ByType = False 'VBC NR
        Exit Function 'VBC NR
    End If 'VBC NR
    imRlfRecLen = Len(tmRlf) 'VBC NR
    tmRlfSrchKey1.sType = slType 'VBC NR
    tmRlfSrchKey1.lRecCode = llRecCode 'VBC NR
    ilRet = btrGetEqual(hlRlf, tmRlf, imRlfRecLen, tmRlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE) 'VBC NR
    If ilRet = BTRV_ERR_NONE Then 'VBC NR
        ilRet = btrDelete(hlRlf) 'VBC NR
        If ilRet = BTRV_ERR_NONE Then 'VBC NR
            gDeleteLockRec_ByType = True 'VBC NR
        Else 'VBC NR
            gDeleteLockRec_ByType = False 'VBC NR
        End If 'VBC NR
    Else 'VBC NR
        gDeleteLockRec_ByType = False 'VBC NR
    End If 'VBC NR

End Function 'VBC NR

Public Function gDeleteLockRec_ByUser(ilUrfCode As Integer) As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim hlRlf As Integer
    ReDim llRlfCode(0 To 0) As Long

    hlRlf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlRlf, "", sgDBPath & "Rlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        gDeleteLockRec_ByUser = True
        Exit Function
    End If
    imRlfRecLen = Len(tmRlf)
    ilRet = btrGetFirst(hlRlf, tmRlf, imRlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        If tmRlf.iUrfCode = ilUrfCode Then
            llRlfCode(UBound(llRlfCode)) = tmRlf.lCode
            ReDim Preserve llRlfCode(0 To UBound(llRlfCode) + 1) As Long
        End If
        ilRet = btrGetNext(hlRlf, tmRlf, imRlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    For ilLoop = 0 To UBound(llRlfCode) - 1 Step 1
        ilRet = gDeleteLockRec_ByRlfCode(hlRlf, llRlfCode(ilLoop))
        If ilRet = False Then
            ilRet = btrClose(hlRlf)
            btrDestroy hlRlf
            gDeleteLockRec_ByUser = True
            Exit Function
        End If
    Next ilLoop
    ilRet = btrClose(hlRlf)
    btrDestroy hlRlf
    Erase llRlfCode
    gDeleteLockRec_ByUser = True

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gIsLocked                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Unschedule spots               *
'*                                                     *
'*******************************************************
Function gIsLocked(tlSdf As SDF) As Integer
'
'   ilRet = gIsLocked(tlSdf)
'
'   Where:
'       tlSdf (I)-Spot record
'       ilRet(O)- True = avail or spot is locked; False = spot not in locked avail or locked
'
    Dim ilRet As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim llSpotDate As Long
    Dim slDate As String
    Dim ilVefCode As Integer
    Dim llSpotTime As Long
    Dim ilGameNo As Integer
    Dim llAvailTime As Long
    Dim ilLoop As Integer
    Dim llAlfStartTime As Long
    Dim llAlfEndTime As Long
    Dim tlAlf As ALF
    Dim llSsfMemDate As Long
    Dim llSsfRecPos As Long
    Dim slTime As String
    
    'Spot summary
    Dim hlSsf As Integer        'Spot summary file handle
    
    gIsLocked = False
    llSsfMemDate = 0
    gUnpackTimeLong tlSdf.iTime(0), tlSdf.iTime(1), False, llSpotTime
    gUnpackTime tlSdf.iTime(0), tlSdf.iTime(1), "A", "1", slTime
    ilGameNo = tlSdf.iGameNo
    ilVefCode = tlSdf.iVefCode
    gUnpackDateLong tlSdf.iDate(0), tlSdf.iDate(1), llSpotDate
    hlSsf = CBtrvTable(TWOHANDLES)        'Create SSF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlSsf)
        btrDestroy hlSsf
        Exit Function
    End If
    hmAlf = CBtrvTable(TWOHANDLES)        'Create SSF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmAlf, "", sgDBPath & "Alf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hlSsf)
        btrDestroy hlSsf
        ilRet = btrClose(hmAlf)
        btrDestroy hmAlf
        Exit Function
    End If
    imAlfRecLen = Len(tmAlf)
    If ilGameNo <= 0 Then
        slDate = Format$(llSpotDate, "m/d/yy")
        gPackDate slDate, ilLogDate0, ilLogDate1
    Else
        'ilLogDate0 = 0
        'ilLogDate1 = 0
        slDate = Format$(llSpotDate, "m/d/yy")
        gPackDate slDate, ilLogDate0, ilLogDate1
    End If
    ilRet = gObtainSsfForDateOrGame(ilVefCode, llSpotDate, slTime, ilGameNo, hlSsf, tmSsf, llSsfMemDate, llSsfRecPos)
    If ilRet Then
        ilRet = BTRV_ERR_NONE
    Else
        ilRet = -1
    End If
    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = ilVefCode) And (((tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1) And (tmSsf.iType = 0) And (ilGameNo = 0)) Or ((tmSsf.iType = ilGameNo) And (ilGameNo <> 0)))
        For ilLoop = 1 To tmSsf.iCount Step 1
           LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilLoop)
            If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Contract Avail subrecord
                'Test time-
                gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llAvailTime
                If (llAvailTime = llSpotTime) Then
                    If ((tmAvail.iAvInfo And SSLOCK) = SSLOCK) Or ((tmAvail.iAvInfo And SSLOCKSPOT) = SSLOCKSPOT) Then
                        ilRet = btrClose(hmAlf)
                        btrDestroy hmAlf
                        ilRet = btrClose(hlSsf)
                        btrDestroy hlSsf
                        gIsLocked = True
                        Exit Function
                    End If
                End If
                If llAvailTime > llSpotTime Then
                    Exit Do
                End If
            End If
        Next ilLoop
        imSsfRecLen = Len(tmSsf)
        ilRet = gSSFGetNext(hlSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_END_OF_FILE Then
            Exit Do
        End If
    Loop
    
    
    If ilGameNo <= 0 Then
        tmAlfSrchkey1.iVefCode = ilVefCode
        tmAlfSrchkey1.iDate(0) = ilLogDate0
        tmAlfSrchkey1.iDate(1) = ilLogDate1
        ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchkey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Else
        tmAlfSrchKey2.iVefCode = ilVefCode
        tmAlfSrchKey2.iGameNo = ilGameNo
        ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
    End If
    Do While ilRet = BTRV_ERR_NONE
        If tmAlf.iVefCode <> ilVefCode Then
            Exit Do
        End If
        If ilGameNo <= 0 Then
            If (tmAlf.iDate(0) <> ilLogDate0) Or (tmAlf.iDate(1) <> ilLogDate1) Then
                Exit Do
            End If
        Else
            If tmAlf.iGameNo <> ilGameNo Then
                Exit Do
            End If
        End If
        gUnpackTimeLong tmAlf.iStartTime(0), tmAlf.iStartTime(1), False, llAlfStartTime
        gUnpackTimeLong tmAlf.iEndTime(0), tmAlf.iEndTime(1), True, llAlfEndTime
        If (llSpotTime >= llAlfStartTime) And (llSpotTime <= llAlfEndTime) Then
            If (tmAlf.sLockType = "S") Or (tmAlf.sLockType = "A") Then
                ilRet = btrClose(hmAlf)
                btrDestroy hmAlf
                ilRet = btrClose(hlSsf)
                btrDestroy hlSsf
                Exit Function
            End If
        End If
        ilRet = btrGetNext(hmAlf, tmAlf, imAlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    ilRet = btrClose(hmAlf)
    btrDestroy hmAlf
    ilRet = btrClose(hlSsf)
    btrDestroy hlSsf
    gIsLocked = False
    Exit Function
End Function

Public Function gFindLockRec(hlRlf As Integer, slType As String, slSubType As String, llRecCode As Long, blIgnoreCurrentUser As Boolean, slUserNameWithLock As String) As Integer
    Dim ilRet As Integer
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llNowDate As Long
    Dim llDate As Long
    Dim ilLoop As Integer
    Dim blFound As Boolean
    
    slNowDate = Format$(gNow(), "m/d/yy")
    slNowTime = Format$(gNow(), "h:mm:ssAM/PM")
    llNowDate = gDateValue(slNowDate)
    slUserNameWithLock = ""
    imRlfRecLen = Len(tmRlf)
    tmRlfSrchKey1.sType = slType
    tmRlfSrchKey1.lRecCode = llRecCode
    ilRet = btrGetGreaterOrEqual(hlRlf, tmRlf, imRlfRecLen, tmRlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmRlf.sType = slType)
        If (tgUrf(0).iCode <> tmRlf.iUrfCode) Or ((Not blIgnoreCurrentUser) And (tgUrf(0).iCode = tmRlf.iUrfCode)) Then
            gUnpackDateLong tmRlf.iEnteredDate(0), tmRlf.iEnteredDate(1), llDate
            If (llDate >= llNowDate) Then
                blFound = False
                If llRecCode <> -1 Then
                    If tmRlf.lRecCode = llRecCode Then
                        blFound = True
                    End If
                End If
                If slSubType <> "" Then
                    If tmRlf.sSubType = slSubType Then
                        blFound = True
                    End If
                End If
                If blFound Then
                    For ilLoop = LBound(tgPopUrf) To UBound(tgPopUrf) - 1 Step 1
                        If tgPopUrf(ilLoop).iCode = tmRlf.iUrfCode Then
                            If Trim$(tgPopUrf(ilLoop).sRept) <> "" Then
                                slUserNameWithLock = Trim$(tgPopUrf(ilLoop).sRept)
                            Else
                                slUserNameWithLock = Trim$(tgPopUrf(ilLoop).sName)
                            End If
                            gFindLockRec = True
                            Exit Function
                        End If
                    Next ilLoop
                End If
            End If
        End If
        ilRet = btrGetNext(hlRlf, tmRlf, imRlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    gFindLockRec = False
End Function

