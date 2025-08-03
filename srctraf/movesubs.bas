Attribute VB_Name = "MoveSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of movesubs.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
'
'
'
Option Explicit
Option Compare Text
' Vehicle File
Dim tmVef As VEF            'VEF record image
Dim imVefRecLen As Integer     'VEF record length
'Contract record information
Dim hmCHF As Integer        'Contract header file handle
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image
'Contract record information
Dim hmClf As Integer        'Contract line file handle
Dim tmClfSrchKey As CLFKEY0 'CLF key record image
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF            'CLF record image
'Contract record information
Dim hmCff As Integer        'Contract line Flight file handle
Dim tmCffSrchKey As CFFKEY0 'CFF key record image
Dim imCffRecLen As Integer  'CFF record length
Dim tmCff As CFF            'CFF record image
'Contract Games
Dim hmCgf As Integer
Dim tmCgf As CGF
Dim imCgfRecLen As Integer
Dim tmCgfSrchKey1 As CGFKEY1    'CntrNo; CntRevNo; PropVer
Dim tmCgfCff() As CFF
'Spot record
Dim hmSdf As Integer
Dim tmSdf As SDF
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey1 As SDFKEY1    '4-26-05 SDF record image (key 1)
Dim tmSdfSrchKey3 As LONGKEY0   '8/10/06 SDF record image (key 3)

'Spot MG record
Dim hmSmf As Integer
Dim tmSmf As SMF
Dim imSmfRecLen As Integer
'Copy Rotation
Dim hmCrf As Integer
' Rate Card Programs/Times File
Dim hmRdf As Integer        'Rate Card Programs/Times file handle
Dim tmLnRdf As RDF            'RDF record image
Dim tmRdfSrchKey As INTKEY0 'RDF key record image
Dim imRdfRecLen As Integer     'RDF record length
'Required by gMakeSsf
Dim hmSsf As Integer
Dim lmSsfDate(0 To 6) As Long    'Dates of the days stored into tmSsf
Dim lmSsfRecPos(0 To 6) As Long  'Record positions
Dim tmSsf(0 To 6) As SSF         'Spot summary for one week (0 index for monday;
'Dim tmSsfOld As SSF
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim imSsfRecLen As Integer
Dim imSelectedDay As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim imVefCode As Integer
Dim imVpfIndex As Integer
Dim imBkQH As Integer   'Rank
Dim imPriceLevel As Integer
'Dim lmTBStartTime(1 To 49) As Long  'Allowed times if time buy (-1 indicates end of times) and library times if library buy
'Dim lmTBEndTime(1 To 49) As Long
Dim lmTBStartTime(0 To 48) As Long  'Allowed times if time buy (-1 indicates end of times) and library times if library buy
Dim lmTBEndTime(0 To 48) As Long
Dim lmSepLength As Long 'Separation length for advertiser
Dim lmStartDateLen As Long  'Start date that separartion is valid for
Dim lmEndDateLen As Long    'End date that separation is valid for
'Dim imHour(1 To 24) As Integer     'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
'Dim imDay(1 To 7) As Integer       'Day count (Index 1 for monday; Index 2 for tuesday;..)
'Dim imQH(1 To 4) As Integer        'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
''Actual for the day or week be processed- this will be a subset from
''imC---- or imP----
'Dim imAHour(1 To 24) As Integer    'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
'Dim imADay(1 To 7) As Integer      'Day count (Index 1 for monday; Index 2 for tuesday;..)
'Dim imAQH(1 To 4) As Integer       'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
'Dim imSkip(1 To 24, 1 To 4, 0 To 6) As Integer  '-1=Skip all test;0=All test;
'                                    'Bit 0=Skip insert;
'                                    'Bit 1=Skip move;
'                                    'Bit 2=Skip competitive pack;
'                                    'Bit 3=Skip Preempt
Dim imHour(0 To 23) As Integer     'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
Dim imDay(0 To 6) As Integer       'Day count (Index 1 for monday; Index 2 for tuesday;..)
Dim imQH(0 To 3) As Integer        'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
'Actual for the day or week be processed- this will be a subset from
'imC---- or imP----
Dim imAHour(0 To 23) As Integer    'Hour count (Index 1 for 12m-1a; Index 2 for 1a-2a;...)
Dim imADay(0 To 6) As Integer      'Day count (Index 1 for monday; Index 2 for tuesday;..)
Dim imAQH(0 To 3) As Integer       'Quarter hour count (Index 1 for 0min to 15min; Index 2 for 15min to 30min;..)
Dim imSkip(0 To 23, 0 To 3, 0 To 6) As Integer  '-1=Skip all test;0=All test;
                                    'Bit 0=Skip insert;
                                    'Bit 1=Skip move;
                                    'Bit 2=Skip competitive pack;
                                    'Bit 3=Skip Preempt

Type PREEMPTVEH                 '5-22-06 added to skip game spots from importing (or skipping the preempted vehicles spots)
    iVefCode As Integer         'vehicle that game preempts
    iAirDate(0 To 1) As Integer
    iPrevDay As Integer         'true if previous day to check form start of day across midnight
    iGameNo As Integer
End Type

Type IGNORETIMES                'start and end times to ignore for imported spots
    lStartTime As Long
    lEndTime As Long
End Type

Type RECONVEHICLES
    iVefCode As Integer
    iGameNo As Integer          '0 if not game
End Type
'Dim tmRec As LPOPREC
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainSDF                      *
'*      Extended read to get all matching SDF          *
'*      records with passed sched date&time& vehicle        *
'*      <input>  hlSdf - SDF  handle (file must be open*
'*               ilDate(0 to 1) - Sch date to match    *
'*               ilTime(0 to 1) - sch time to match
'                ilVehCode - vehicle to match          *
'*      <I/O>    tlSdf() - array of matching SDF recds *
'*                                                     *
'*             Created:5/11/01       By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read all SDF records by
'                       date, time &vehicle            *
'*      4-26-05 speed up retrieval by setting the initial*
'*              keyfields
'*******************************************************
Function gObtainSDFbyVehDate(RptForm As Form, hlSdf As Integer, ilSchDate() As Integer, ilVefCode As Integer, tlSdf() As SDF) As Integer
'
'    gObtainSDF (hlSDf, ilSchDate(), tlSdf())
'
    Dim ilRet As Integer    'Return status
    Dim slStr As String
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilSdfUpper As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlIntTypeBuff As POPICODE

    ReDim tlSdf(0 To 0) As SDF
    btrExtClear hlSdf   'Clear any previous extend operation
    ilExtLen = Len(tlSdf(0))  'Extract operation record size
    imSdfRecLen = Len(tmSdf)
    ilSdfUpper = UBound(tlSdf)
    'ilRet = btrGetFirst(hlSdf, tmSdf, imSdfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    '4-26-05 speed up retrieval by setting the initial keyfields
    tmSdfSrchKey1.iVefCode = ilVefCode
    tmSdfSrchKey1.iDate(0) = ilSchDate(0)
    tmSdfSrchKey1.iDate(1) = ilSchDate(1)
    tmSdfSrchKey1.iTime(0) = 0
    tmSdfSrchKey1.iTime(1) = 0
    tmSdfSrchKey1.sSchStatus = ""
    ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation    If ilRet <> BTRV_ERR_END_OF_FILE Then

    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        'Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC") '"EG") 'Set extract limits (all records)
        Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDF", "") '"EG") 'Set extract limits (all records)
        'Match on vehicle & date, get all times & Status'
        tlIntTypeBuff.iCode = ilVefCode
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)

        tlDateTypeBuff.iDate0 = ilSchDate(0)
        tlDateTypeBuff.iDate1 = ilSchDate(1)
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        On Error GoTo mObtainSdfErr
        gBtrvErrorMsg ilRet, "gObtainSDFbyVehDate (btrExtAddLogicConst):" & "Sdf.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtAddField(hlSdf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainSdfErr
        gBtrvErrorMsg ilRet, "gObtainSDFbyVehDate (btrExtAddField):" & "Sdf.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainSdfErr
            gBtrvErrorMsg ilRet, "gObtainSDFbyVehDate (btrExtGetNextExt):" & "Sdf.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmSdf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                slStr = ""
                tlSdf(UBound(tlSdf)) = tmSdf           'save entire record
                ReDim Preserve tlSdf(0 To UBound(tlSdf) + 1) As SDF
                ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSdf, tmSdf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainSDFbyVehDate = True
    Exit Function
mObtainSdfErr:
    On Error GoTo 0
    gObtainSDFbyVehDate = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gProcessSpot                    *
'*                                                     *
'*      Created:5-11-01       By:D. Hosaka
'       (copied from Post Log)
'*                                                     *
'*
'*                                                     *
'*  4-6-05 Remove exceeds reads to SSF to speed-up process
'*         Remove error logging, its done for each vehicle not
'*         each spot
'*  6-10-05 exit without aborting when smf or contract
'           not found/read
'*******************************************************
Function gProcessSpot(hlSsf As Integer, hlSdf As Integer, hlSmf As Integer, hlChf As Integer, hlClf As Integer, hlCff As Integer, hlRdf As Integer, hlCrf As Integer, hlCgf As Integer, ttSdf As SDF, ilInVefCode As Integer, slSdfDate As String, slSdfTime As String, slImportDate As String, slImportTime As String, slMOrO As String, hlSxf As Integer, Optional hlGsf As Integer = 0, Optional hlGhf As Integer = 0) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilRetC                                                                                *
'******************************************************************************************

'
'   Where:
'       hlSsf (I)- SSF handle
'       hlSdf (I)- SDF handle
'       hlSmf (I)- SMF handle
'       hlChf (I) - CHF handle
'       hlClf (I) - CLF handle
'       hlCff (I) - CFF handle
'       hlRdf (I) - RDF handle
'       hlCRF (I) - CRF handle
'       ttSdf (I) - SDF buffer
'       ilVefCode (I) - vehicle code
'       slSDFDate (I) - orig date of spot
'       slSdfTime (I) - orig time of spot
'       slImportDate (I) - new date of spot
'       slImportTime (I) - new time of spot
'       slMorO(I) - from site preference (what to do with discrepant spots: mg or outside)
'
'    ilRet = True(ok) or False(error)
'
    Dim llSdfRecPos As Long
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilError As Integer
    Dim ilFindAdjAvail As Integer
    Dim slOrigTime As String
    Dim ilAvailIndex As Integer
    Dim slRet As String
    Dim ilBkQH As Integer
    Dim tlSdf As SDF
    Dim ilMissedToSch As Integer
    Dim ilGameNo As Integer
    Dim ilFirst As Integer

    imVefRecLen = Len(tmVef)
    hmSsf = hlSsf
    hmSdf = hlSdf
    imSdfRecLen = Len(tmSdf)
    hmSmf = hlSmf
    imSmfRecLen = Len(tmSmf)
    hmCHF = hlChf
    imCHFRecLen = Len(tmChf)
    hmClf = hlClf
    imClfRecLen = Len(tmClf)
    hmCff = hlCff
    imCffRecLen = Len(tmCff)
    hmCgf = hlCgf
    imCgfRecLen = Len(tmCgf)
    hmRdf = hlRdf
    imRdfRecLen = Len(tmLnRdf)
    hmCrf = hlCrf
    imVpfIndex = gVpfFindIndex(ilInVefCode)
    imVefCode = ilInVefCode

    tmSdf = ttSdf
    '8/10/06:  Added to make sure that the getdirect is Ok
    tmSdfSrchKey3.lCode = tmSdf.lCode
    ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
    If ilRet <> BTRV_ERR_NONE Then
        gProcessSpot = 1
        Exit Function
    End If

    If tmSdf.sSchStatus = "M" Or tmSdf.sSchStatus = "C" Or tmSdf.sSchStatus = "H" Then
        ilMissedToSch = True
        gProcessSpot = 2
        Exit Function
    Else
        '5/20/11
        If gDateValue(slSdfDate) = gDateValue(slImportDate) Then
            ilFirst = True
            Do
                '1/10/11: Check if time outside contract times.  If so, make spot MG or Outside
                If (tmSdf.sSchStatus <> "G") And (tmSdf.sSchStatus <> "O") Then
                    slRet = mMoveTest(slSdfDate, slImportDate, slImportTime, slMOrO)
                    If (slRet = "G") Or (slRet = "O") Then
                        If ilFirst Then
                            tmSmf.lChfCode = 0
                            ilRet = gMakeSmf(hmSmf, tmSmf, slRet, tmSdf, tmSdf.iVefCode, slSdfDate, slSdfTime, tmSdf.iGameNo, slImportDate, slImportTime)
                            If ilRet Then
                                tmSdf.lSmfCode = tmSmf.lCode
                                tmSdf.sSchStatus = slRet
                            Else
                                gProcessSpot = 1
                                Exit Function
                            End If
                        Else
                            tmSdf.lSmfCode = tmSmf.lCode
                            tmSdf.sSchStatus = slRet
                        End If
                        ilFirst = False
                    ElseIf slRet = "" Or slRet = "1" Or slRet = "2" Then
                        gProcessSpot = 1
                        Exit Function
                    End If
                Else
                    '1/10/11: Later, add test if MG/Outside should be removed
                End If
                gPackTime slImportTime, tmSdf.iTime(0), tmSdf.iTime(1)
                tmSdf.sAffChg = "Y"
                tmSdf.sXCrossMidnight = ttSdf.sXCrossMidnight       '3-28-07
                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen) '  Update File
                If ilRet = BTRV_ERR_CONFLICT Then
                    tmSdfSrchKey3.lCode = tmSdf.lCode
                    ilCRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                gProcessSpot = 1
                Exit Function
            End If
            gProcessSpot = 0
            Exit Function
        Else
            gProcessSpot = 3
            Exit Function
        End If
        ilMissedToSch = False
    End If
    ilGameNo = tmSdf.iGameNo
    If ilGameNo > 0 Then
        imVefCode = tmSdf.iVefCode
        imVpfIndex = gVpfFindIndex(imVefCode)
    End If
    Do
        '4-6-05 take out error logging and do on vehicle rather thanby spot
        'ilRet = btrBeginTrans(hmSdf, 1000)
        'If ilRet <> BTRV_ERR_NONE Then
        '    Screen.MousePointer = vbDefault
        '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "gProcessSpot")
        '    gProcessSpot = 1
        '    Exit Function
        'End If

        'If StrComp(slSdfDate, slImportDate) <> 0 Or StrComp(slSdfTime, slImportTime) <> 0 Then    'any changes in sch date/time vs aired date/time?
         If gDateValue(slSdfDate) <> gDateValue(slImportDate) Or gTimeToLong(slSdfTime, False) <> gTimeToLong(slImportTime, False) Then      '4-5-05 any changes in sch date/time vs aired date/time?
           ilRet = btrGetPosition(hlSdf, llSdfRecPos) 'retrieve the SDF position for direct reads
            ilFindAdjAvail = True
            If ilGameNo = 0 Then
                imSelectedDay = gWeekDayStr(slSdfDate)
            Else
                imSelectedDay = 0
            End If
            slOrigTime = slSdfTime
            If Not mFindAvail(slImportDate, slImportTime, ilGameNo, ilFindAdjAvail, ilAvailIndex) Then      'obtain the ssf for the avail
                'ilRet = btrAbortTrans(hmSdf)
                Screen.MousePointer = vbDefault
                gProcessSpot = 1
                Exit Function
            End If
            If Not mAvailRoom(ilAvailIndex) Then
                'ilRet = btrAbortTrans(hmSdf)
                Screen.MousePointer = vbDefault
                gProcessSpot = 1
                Exit Function
            End If
            slRet = mMoveTest(slSdfDate, slImportDate, slImportTime, slMOrO)
            If slRet = "" Or slRet = "1" Or slRet = "2" Then    'cant read contract or cant find smf
                'ilRet = btrAbortTrans(hmSdf)
                Screen.MousePointer = vbDefault
                gProcessSpot = 1
                Exit Function
            End If

            If Not ilMissedToSch Then       'if missed to Sch, need to book new time;  bypass this status change
                'Unschedule, then schedule (gChgSchSpot removes Smf if exist)
                gPackDate slSdfDate, tmSdf.iDate(0), tmSdf.iDate(1)
                gPackTime slOrigTime, tmSdf.iTime(0), tmSdf.iTime(1)
                ilRet = gChgSchSpot("TM", hmSdf, tmSdf, hmSmf, ilGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hlSxf, hlGsf, hlGhf)
                If Not ilRet Then
                    '6-10-05 let process go thru remainder of file
                    'ilRet = btrAbortTrans(hmSdf)
                    'Screen.MousePointer = vbDefault
                    'ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "ProcessSpot")
                    gProcessSpot = 1
                    Exit Function
                End If
                '5/20/11
                'If Not mRemoveAvail(slSdfDate, slOrigTime, ilGameNo) Then
                '    '6-10-05 let process go thru remainder of file
                '    'ilRet = btrAbortTrans(hmSdf)
                '    'Screen.MousePointer = vbDefault
                '    'ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "ProcessSpot")
                '    gProcessSpot = 1
                '    Exit Function
                'End If

                If Not mFindAvail(slImportDate, slImportTime, ilGameNo, ilFindAdjAvail, ilAvailIndex) Then
                    '6-10-05 let process go thru remainder of file
                    'ilRet = btrAbortTrans(hmSdf)
                    'Screen.MousePointer = vbDefault
                    'ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "ProcessSpot")
                    gProcessSpot = 1
                    Exit Function
                End If
            End If

            '4-6-05 moved to inside If Not ilMissedToSch test above
            'If Not mFindAvail(slImportDate, slImportTime, ilFindAdjAvail, ilAvailIndex) Then
            '    ilRet = btrAbortTrans(hmSdf)
            '    Screen.MousePointer = vbDefault
            '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
            '    gProcessSpot = 1
            '    Exit Function
            'End If

            If imBkQH <= 1000 Then  'Above 1000 is DR; Remnant; PI; Trade; PSA; Promo
                If (slRet = "G") Or (slRet = "O") Then
                    ilBkQH = 0
                Else
                    ilBkQH = imBkQH
                End If
            Else    'D.R.; Remnant; PI, Trade; Promo; PSA
                ilBkQH = imBkQH
            End If
            'Schedule spot, Smf created if required
            tlSdf = tmSdf
            ilRet = gBookSpot(slRet, hmSdf, tmSdf, llSdfRecPos, ilBkQH, hmSsf, tmSsf(imSelectedDay), lmSsfRecPos(imSelectedDay), ilAvailIndex, -1, tmChf, tmClf, tmLnRdf, imVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, imPriceLevel, False, hlSxf, hlGsf)
            If Not ilRet Then
                '6-10-05 let process go thru remainder of file
                'ilRet = btrAbortTrans(hmSdf)
                'Screen.MousePointer = vbDefault
                'ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "ProcessSpot")
                gProcessSpot = 1
                Exit Function
            End If
            'Reset copy
            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                '6-10-05 let process go thru remainder of file
                'ilRet = btrAbortTrans(hmSdf)
               ' Screen.MousePointer = vbDefault
                'ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "gProcessSpot")
                gProcessSpot = 1
                Exit Function
            End If
            'tmRec = tmSdf
            'ilRet = gGetByKeyForUpdate("SDF", hmSdf, tmRec)
            'tmSdf = tmRec
            'If ilRet <> BTRV_ERR_NONE Then
            '    ilRet = btrAbortTrans(hmSdf)
            '    Screen.MousePointer = vbDefault
            '    ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "Erase")
            '    gProcessSpot = 1
            '    Exit Function
            'End If

            If ilMissedToSch Then           'missed to schedule vs just time update
                tmSdf.sTracer = "N"     'N/A: Import change
                tmSdf.iMnfMissed = 0    'Missed reason
                tmSdf.sAffChg = "A"     'added
                gPackTime slImportTime, tmSdf.iTime(0), tmSdf.iTime(1)
            Else
                tmSdf.iRotNo = tlSdf.iRotNo
                tmSdf.sPtType = tlSdf.sPtType
                tmSdf.lCopyCode = tlSdf.lCopyCode
                If ilFindAdjAvail Then
                    gPackDate slImportDate, tmSdf.iDate(0), tmSdf.iDate(1)
                    gPackTime slImportTime, tmSdf.iTime(0), tmSdf.iTime(1)
                End If
                tmSdf.sAffChg = "Y"
            End If

            tmSdf.sXCrossMidnight = ttSdf.sXCrossMidnight       '3-28-07
            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen) '  Update File
            'If ilRet = BTRV_ERR_CONFLICT Then
            '    ilRetC = btrAbortTrans(hmSdf)
            'End If
        End If

    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        '6-10-05 let process go thru remainder of file
        'ilRet = btrAbortTrans(hmSdf)
        'Screen.MousePointer = vbDefault
        'ilRet = MsgBox("Save Not Completed, Try Later", vbOkOnly + vbExclamation, "gProcessSpot")
        gProcessSpot = 1
        Exit Function
    End If
    ilRet = btrEndTrans(hmSdf)
    gProcessSpot = 0
    Exit Function

    On Error GoTo 0
    ilError = True
    Resume Next
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mAvailRoom                      *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if room exist for    *
'*                      spot within avail              *
'*                                                     *
'*******************************************************
Function mAvailRoom(ilAvailIndex) As Integer
'
'   ilRet = mAvailRoom(ilAvailIndex)
'   where:
'       ilAvailIndex(I)- location of avail within Ssf (use mFindAvail)
'       ilRet(O)- True=Avail has room; False=insufficient room within avail
'
'       tmSdf(I)- spot records
'
'       Code later: ask if avail should be overbooked
'                   If so, create a version zero (0) of the library with the new
'                   units/seconds
'
    Dim ilAvailUnits As Integer
    Dim ilAvailSec As Integer
    Dim ilUnitsSold As Integer
    Dim ilSecSold As Integer
    Dim ilSpotLen As Integer
    Dim ilSpotUnits As Integer
    Dim ilSpotIndex As Integer
    Dim ilNewUnit As Integer
    Dim ilNewSec As Integer
    Dim ilRet As Integer
    tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex)
    ilAvailUnits = tmAvail.iAvInfo And &H1F
    ilAvailSec = tmAvail.iLen
    For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tmAvail.iNoSpotsThis Step 1
        LSet tmSpot = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilSpotIndex)
        If tmSpot.lSdfCode = tmSdf.lCode Then
            mAvailRoom = True
            Exit Function
        End If
        If (tmSpot.iRecType And &HF) >= 10 Then
            ilSpotLen = tmSpot.iPosLen And &HFFF
            If (tgVpf(imVpfIndex).sSSellOut = "T") Then
                ilSpotUnits = ilSpotLen \ 30
                If ilSpotUnits <= 0 Then
                    ilSpotUnits = 1
                End If
                ilSpotLen = 0
            Else
                ilSpotUnits = 1
                'If (tgVpf(imVpfIndex).sSSellOut = "U") Then
                '    ilSpotLen = 0
                'End If
            End If
            ilUnitsSold = ilUnitsSold + ilSpotUnits
            ilSecSold = ilSecSold + ilSpotLen
        End If
        Next ilSpotIndex
        ilSpotLen = tmSdf.iLen
        If (tgVpf(imVpfIndex).sSSellOut = "T") Then
                ilSpotUnits = ilSpotLen \ 30
            If ilSpotUnits <= 0 Then
                ilSpotUnits = 1
            End If
            ilSpotLen = 0
        Else
            ilSpotUnits = 1
            'If (tgVpf(imVpfIndex).sSSellOut = "U") Then
            '    ilSpotLen = 0
            'End If
        End If
        ilNewUnit = 0
        ilNewSec = 0
        If (tgVpf(imVpfIndex).sSSellOut = "M") Then
            If (ilSpotLen + ilSecSold <> ilAvailSec) Or (ilSpotUnits + ilUnitsSold <> ilAvailUnits) Then
                ilNewSec = ilSpotLen + ilSecSold
                ilNewUnit = ilSpotUnits + ilUnitsSold
            Else
                mAvailRoom = True
                Exit Function
            End If
            Else
            If (ilSpotLen + ilSecSold > ilAvailSec) Or (ilSpotUnits + ilUnitsSold > ilAvailUnits) Then
                ilNewSec = ilSpotLen + ilSecSold
                ilNewUnit = ilSpotUnits + ilUnitsSold
            Else
                mAvailRoom = True
                Exit Function
            End If
        End If
        Do
        imSsfRecLen = Len(tmSsf(imSelectedDay))
        ilRet = gSSFGetDirect(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, lmSsfRecPos(imSelectedDay), INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mAvailRoom = False
            Exit Function
        End If
        ilRet = gGetByKeyForUpdateSSF(hmSsf, tmSsf(imSelectedDay))
        If ilRet <> BTRV_ERR_NONE Then
            mAvailRoom = False
            Exit Function
        End If
        '5/20/11
        If (tmAvail.iOrigUnit = 0) And (tmAvail.iOrigLen = 0) Then
            tmAvail.iOrigUnit = tmAvail.iAvInfo And &H1F
            tmAvail.iOrigLen = tmAvail.iLen
        End If
        tmAvail.iAvInfo = (tmAvail.iAvInfo And (Not &H1F)) + ilNewUnit
        tmAvail.iLen = ilNewSec
        tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex) = tmAvail
        imSsfRecLen = 17 + tmSsf(imSelectedDay).iCount * Len(tmProg)
        ilRet = gSSFUpdate(hmSsf, tmSsf(imSelectedDay), imSsfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        mAvailRoom = False
    Exit Function
    End If
    mAvailRoom = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mFindAvail                      *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get avail within Ssf           *
'*                                                     *
'*******************************************************
Function mFindAvail(slSchDate As String, slFindTime As String, ilGameNo As Integer, ilFindAdjAvail As Integer, ilAvailIndex As Integer) As Integer
'
'   ilRet = mFindAvail(slSchDate, slSchTime, ilAvailIndex)
'   Where:
'       slSchDate(I)- Scheduled Date
'       slSchTime(I)- Time that avail is to be found at
'       ilFindAdjAvail(I)- Find closest avail to specified time
'       llSsfRecPos(O)- Ssf record position
'       ilAvailIndex(O)- Index into Ssf where avail is located
'       ilRet(O)- True=Avail found; False=Avail not found
'       lmSsfRecPos(O)- Ssf record position
'
    Dim ilRet As Integer
    Dim llSchDate As Long
    Dim llTime As Long
    Dim llTstTime As Long
    Dim llFndAdjTime As Long
    Dim ilLoop As Integer
    llTime = CLng(gTimeToCurrency(slFindTime, False))
    llSchDate = gDateValue(slSchDate)
    If ilGameNo = 0 Then
        imSelectedDay = gWeekDayStr(slSchDate)
    Else
        imSelectedDay = 0
    End If
    'lmSsfDate(imSelectedDay) = 0     gObtainSsfForDateOrGame reads record back inwith direct call if dates are the same
    ilRet = gObtainSsfForDateOrGame(imVefCode, llSchDate, slFindTime, ilGameNo, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay))
    If Not ilRet Then
        mFindAvail = False
        Exit Function
    End If
    llFndAdjTime = -1
    For ilLoop = 1 To tmSsf(imSelectedDay).iCount Step 1
        tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilLoop)
        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTstTime
            If llTime = llTstTime Then 'Replace
                ilAvailIndex = ilLoop
                mFindAvail = True
                Exit Function
            ElseIf (llTstTime < llTime) And (ilFindAdjAvail) Then
                ilAvailIndex = ilLoop
                llFndAdjTime = llTstTime
            ElseIf (llTime < llTstTime) And (ilFindAdjAvail) Then
                If llFndAdjTime = -1 Then
                    ilAvailIndex = ilLoop
                    mFindAvail = True
                    Exit Function
                Else
                    If (llTime - llFndAdjTime) < (llTstTime - llTime) Then
                        mFindAvail = True
                        Exit Function
                    Else
                        ilAvailIndex = ilLoop
                        mFindAvail = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next ilLoop
    If (llFndAdjTime <> -1) And (ilFindAdjAvail) Then
        mFindAvail = True
        Exit Function
    End If
    mFindAvail = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mMoveTest                       *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Determine if spot can be moved *
'*                      into specified location        *
'*                                                     *
'*******************************************************
Function mMoveTest(slSchDate As String, slMoveDate As String, slMoveTime As String, slMOrO As String) As String
'
'   slRet = mMoveTest(slSchDate, slMoveDate, slMoveTime)
'       Where:
'           tmSdf (I)- Contains spot
'
'           slRet(O)-   "1" = Abort move, unable to read contract
'                       "2" = abort, unable to find smf
'                       "S"=Move
'                       "G"=Move as MG
'                       "O"=Move and set as moved outside contract limits
'
    Dim ilMoveDay As Integer
    Dim llDate As Long
    Dim ilRet As Integer
    Dim ilMGMove As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim slMsg As String
    Dim ilWeekMoveOk As Integer
    Dim ilDayMoveOk As Integer
    Dim ilTimeMoveOk As Integer
    Dim llChfCode As Long
    Dim ilLineNo As Integer
    Dim ilAdfCode As Integer
    Dim ilVehComp As Integer
    Dim slDate As String
    Dim slType As String
    Dim slWkDate As String
    Dim llTime As Long
    Dim llMoDate As Long
    Dim llSuDate As Long
    Dim slTime As String
    Dim tlSmf As SMF
    ReDim tlCff(0 To 1) As CFF
    Dim llEarliestAllowedDate As Long
    Dim ilGameNo As Integer
    Dim slMissedDate As String

    ilGameNo = tmSdf.iGameNo
    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
        If gFindSmf(tmSdf, hmSmf, tlSmf) Then
            ilGameNo = tlSmf.iGameNo
        End If
    End If
    ilRet = mReadChfClfRdfCffRec(tmSdf.lChfCode, tmSdf.iLineNo, ilGameNo, slMoveDate)
    If Not ilRet Then
        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slWkDate
        gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
        'sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Unable to Read Contract " & slWkDate & " " & slTime & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
        'ReDim Preserve sgSSFErrorMsg(1 To UBound(sgSSFErrorMsg) + 1) As String
        'MsgBox = "Unable to Read Contract " & slWkDate & " " & slTime & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
        mMoveTest = "1"  'Abort
        Exit Function
    End If
    ilWeekMoveOk = True
    ilDayMoveOk = True
    ilTimeMoveOk = True
    llTime = CLng(gTimeToCurrency(slMoveTime, False))
    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
        llMoDate = 0
        If gFindSmf(tmSdf, hmSmf, tlSmf) Then
            gUnpackDate tlSmf.iMissedDate(0), tlSmf.iMissedDate(1), slWkDate
            slMissedDate = slWkDate
            llMoDate = gDateValue(gObtainPrevMonday(slWkDate))
            llSuDate = gDateValue(gObtainNextSunday(slWkDate))
        End If
        If llMoDate = 0 Then
            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slWkDate
            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
            'sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Unable to find SMF " & slWkDate & " " & slTime & " Cntr #=" & Str$(tmChf.lCntrNo) & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
            'ReDim Preserve sgSSFErrorMsg(1 To UBound(sgSSFErrorMsg) + 1) As String
            'MsgBox = "Unable to find SMF " & slWkDate & " " & slTime & " Cntr #=" & Str$(tmChf.lCntrNo) & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
            mMoveTest = "2"  'Abort
            Exit Function
        End If
    Else
        'gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slWkDate
        slWkDate = slSchDate
        llMoDate = gDateValue(gObtainPrevMonday(slWkDate))
        llSuDate = gDateValue(gObtainNextSunday(slWkDate))
    End If
    slType = tmSdf.sSpotType
    ilMoveDay = gWeekDayStr(slMoveDate)
    llDate = gDateValue(slMoveDate)
    llDate = gDateValue(slMoveDate) 'smSelectedDate)
    tlCff(0) = tmCff
    llEarliestAllowedDate = 0
    gGetLineSchParameters hmSsf, tmSsf(), lmSsfDate(), lmSsfRecPos(), llDate, imVefCode, tmChf.iAdfCode, ilGameNo, tlCff(), tmClf, tmLnRdf, lmSepLength, lmStartDateLen, lmEndDateLen, llChfCode, ilLineNo, ilAdfCode, ilVehComp, imHour(), imDay(), imQH(), imAHour(), imADay(), imAQH(), lmTBStartTime(), lmTBEndTime(), llEarliestAllowedDate, imSkip(), tmChf.sType, tmChf.iPctTrade, imBkQH, True, imPriceLevel, False
    'If cff not found, then spot is outside date definition
    ilMGMove = 0
    If tmCff.sDelete = "Y" Then
        ilWeekMoveOk = False
    Else
        'Test if within same week
        If (llDate < llMoDate) Or (llDate > llSuDate) Then
            ilWeekMoveOk = False
        End If
        'Test days
        If (tmCff.iSpotsWk > 0) Or (tmCff.iXSpotsWk > 0) Then 'Weekly
            If (tmCff.iDay(ilMoveDay) <= 0) And (tmCff.sXDay(ilMoveDay) <> "Y") Then
                ilDayMoveOk = False
            End If
        Else
            If tmCff.iDay(ilMoveDay) <= 0 Then
                ilDayMoveOk = False
            End If
            If ((tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O")) And (slMissedDate <> "") Then
                slDate = slMissedDate
            Else
                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
            End If
            If gDateValue(slDate) <> gDateValue(slMoveDate) Then
                ilDayMoveOk = False
            End If
        End If
    End If
    'Check Times
    ilFound = False
    For ilLoop = LBound(lmTBStartTime) To UBound(lmTBEndTime) Step 1
        If lmTBStartTime(ilLoop) <> -1 Then
            If (llTime >= lmTBStartTime(ilLoop)) And (llTime < lmTBEndTime(ilLoop)) Then
                ilFound = True
                Exit For
            End If
        End If
    Next ilLoop
    If Not ilFound Then
        ilTimeMoveOk = False
    End If
    slMsg = ""
    If (tmSdf.sSchStatus <> "G") And (tmSdf.sSchStatus <> "O") Then
        If tmClf.iVefCode <> imVefCode Then
            slMsg = "Move violates Vehicle"
        End If
        If Not ilWeekMoveOk Then
            If slMsg = "" Then
                slMsg = "Move violates Weeks"
            Else
                slMsg = slMsg & ", weeks"
            End If
        End If
        If Not ilTimeMoveOk Then
            If slMsg = "" Then
                slMsg = "Move violates Times"
            Else
                slMsg = slMsg & ", times"
            End If
        End If
        If Not ilDayMoveOk Then
            If slMsg = "" Then
                slMsg = "Move violates Days"
            Else
                slMsg = slMsg & ", days"
            End If
        End If
    Else
        If (tmSdf.sSchStatus = "G") Then
            If slType = "X" Then
                ilMGMove = vbYes
            End If
            If tmClf.iVefCode <> imVefCode Then
                ilMGMove = vbYes
            End If
            If Not ilWeekMoveOk Then
                ilMGMove = vbYes
            End If
            If Not ilTimeMoveOk Then
                ilMGMove = vbYes
            End If
            If Not ilDayMoveOk Then
                ilMGMove = vbYes
            End If
        Else
            If slType = "X" Then
                ilMGMove = vbNo
            End If
            If tmClf.iVefCode <> imVefCode Then
                ilMGMove = vbNo
            End If
            If Not ilWeekMoveOk Then
                ilMGMove = vbNo
            End If
            If Not ilTimeMoveOk Then
                ilMGMove = vbNo
            End If
            If Not ilDayMoveOk Then
                ilMGMove = vbNo
            End If
        End If
    End If
    If slMsg <> "" Then
        If (slType <> "S") And (slType <> "M") And (slType <> "T") And (slType <> "Q") Then
            If slType <> "X" Then
                If slMOrO = "G" Then
                    ilMGMove = vbYes       'MG
                Else
                    ilMGMove = vbNo        'Outside
                End If
            Else
                ilMGMove = vbNo
            End If
        Else
            ilMGMove = vbNo
        End If
    End If
    If (ilMGMove = vbYes) Or (ilMGMove = vbOK) Then
        mMoveTest = "G"
    ElseIf ilMGMove = vbNo Then
        mMoveTest = "O"
    Else
        mMoveTest = "S"
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadChfClfRdfCffRec            *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read records                   *
'*                                                     *
'*******************************************************
Function mReadChfClfRdfCffRec(llChfCode As Long, ilLineNo As Integer, ilGameNo As Integer, slSpotDate As String) As Integer
'
'   iRet = mReadChfClfRdpfCffRec(llChfCode, ilLineNo, slMissedDate, SlStartDate, slEndDate, slNoSpots)
'   Where:
'       llChfCode (I) - Contract code
'       ilLineNo(I)- Line number
'       slMissedDate(I)- Missed date or date to find bracketing week
'       tmCff(O)- contains valid flight week (if sDelete = "Y", then week is invalid)
'       iRet (O)- True if all records read,
'                 False if error in read
'
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStartDate As String
    Dim llStartDate As Long
    Dim slEndDate As String
    Dim llEndDate As Long
    Dim llSpotDate As Long
    Dim ilDay As Integer
    Dim ilIndex As Integer
    Dim ilLastDay As Integer
    Dim ilFirstDay As Integer
    Dim tlCff As CFF
    ilDay = gWeekDayStr(slSpotDate)
    ilLastDay = -1
    ilFirstDay = -1
    tmCff.sDelete = "Y"  'Set as flag that illegal week
    If mReadChfClfRdfRec(llChfCode, ilLineNo) Then
        llStartDate = 0
        llEndDate = 0
        llSpotDate = gDateValue(slSpotDate)
        If ilGameNo = 0 Then
            tmCffSrchKey.lChfCode = llChfCode
            tmCffSrchKey.iClfLine = ilLineNo
            tmCffSrchKey.iCntRevNo = tmClf.iCntRevNo
            tmCffSrchKey.iPropVer = tmClf.iPropVer
            tmCffSrchKey.iStartDate(0) = 0
            tmCffSrchKey.iStartDate(1) = 0
            ilRet = btrGetGreaterOrEqual(hmCff, tlCff, imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tlCff.lChfCode = llChfCode) And (tlCff.iClfLine = ilLineNo)
                If (tlCff.iCntRevNo = tmClf.iCntRevNo) And (tlCff.iPropVer = tmClf.iPropVer) And (tlCff.sDelete <> "Y") Then
                    gUnpackDate tlCff.iStartDate(0), tlCff.iStartDate(1), slStartDate    'Week Start date
                    gUnpackDate tlCff.iEndDate(0), tlCff.iEndDate(1), slEndDate    'Week Start date
                    If llStartDate = 0 Then
                        llStartDate = gDateValue(slStartDate)
                        llEndDate = gDateValue(slEndDate)
                    Else
                        If gDateValue(slStartDate) < llStartDate Then
                            llStartDate = gDateValue(slStartDate)
                        End If
                        If gDateValue(slEndDate) > llEndDate Then
                            llEndDate = gDateValue(slEndDate)
                        End If
                    End If
                    If (llSpotDate >= gDateValue(slStartDate)) And (llSpotDate <= gDateValue(slEndDate)) Then
                        tmCff = tlCff
                        If (tmCff.iSpotsWk <> 0) Or (tmCff.iXSpotsWk <> 0) Then 'Weekly
                            For ilIndex = 0 To 6 Step 1
                                If tmCff.iDay(ilIndex) > 0 Then
                                    If ilFirstDay = -1 Then
                                        ilFirstDay = ilIndex
                                    End If
                                    ilLastDay = ilIndex
                                End If
                            Next ilIndex
                            If (ilDay < ilFirstDay) Or (ilDay > ilLastDay) Then
                                ilLastDay = -1
                                ilFirstDay = -1
                                For ilIndex = 0 To 6 Step 1
                                    If tmCff.sXDay(ilIndex) = "Y" Then
                                        If ilFirstDay = -1 Then
                                            ilFirstDay = ilIndex
                                        End If
                                        ilLastDay = ilIndex
                                    End If
                                Next ilIndex
                            End If
                        Else    'Daily
                            ilFirstDay = ilDay
                            ilLastDay = ilDay
                        End If
                        Exit Do
                    End If
                End If
                ilRet = btrGetNext(hmCff, tlCff, imCffRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Else
            tmCgfSrchKey1.lClfCode = tmClf.lCode
            ilRet = btrGetEqual(hmCgf, tmCgf, imCgfRecLen, tmCgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            Do While tmClf.lCode = tmCgf.lClfCode
                If tmCgf.iGameNo = ilGameNo Then
                    gCgfToCff tmClf, tmCgf, tmCgfCff()
                    tmCff = tmCgfCff(0) 'tmCgfCff(1)
                    Exit Do
                End If
                ilRet = btrGetNext(hmCgf, tmCgf, imCgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            Erase tmCgfCff
        End If
        'Determine times
        For ilLoop = LBound(lmTBStartTime) To UBound(lmTBStartTime) Step 1
            lmTBStartTime(ilLoop) = -1
            lmTBEndTime(ilLoop) = -1
        Next ilLoop
        mReadChfClfRdfCffRec = True
        'set of lmTBStartTime and lmTBEndTime are now set in mMoveTest which calls the function

        'If (tmCff.sDelete <> "Y") And (ilFirstDay <> -1) Then
        '    ilTBIndex = 1
        '    If (tmLnRdf.iLtfCode(0) <> 0) Or (tmLnRdf.iLtfCode(1) <> 0) Or (tmLnRdf.iLtfCode(2) <> 0) Then
        '        'Read Ssf for date- test for library- code removed- as Ssf not read into memory
        '        'this can be added if required
        '        'See gGetLineSchParameters for code
        '        'For now set time as 12m-12m
        '        lmTBStartTime(ilTBIndex) = 0
        '        lmTBEndTime(ilTBIndex) = 86400  '24*3600
        '    Else    'Time buy- check if override times defined (if so, use them as bump times)
        '        If (tmClf.iStartTime(0) = 1) And (tmClf.iStartTime(1) = 0) Then
        '            For ilLoop = LBound(tmLnRdf.iStartTime, 2) To UBound(tmLnRdf.iStartTime, 2) Step 1
        '                If (tmLnRdf.iStartTime(0, ilLoop) <> 1) Or (tmLnRdf.iStartTime(1, ilLoop) <> 0) Then
        '                    If (tmCff.iSpotsWk = 0) And (tmCff.iXSpotsWk = 0) Then 'Daily- Test if valid day
        '                        If tmLnRdf.sWkDays(ilLoop, ilDay + 1) = "Y" Then
        '                            gUnpackTime tmLnRdf.iStartTime(0, ilLoop), tmLnRdf.iStartTime(1, ilLoop), "A", "1", slLnStart
        '                            gUnpackTime tmLnRdf.iEndTime(0, ilLoop), tmLnRdf.iEndTime(1, ilLoop), "A", "1", slLnEnd
        '                            lmTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLnStart, False))
        '                            lmTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLnEnd, True))
        '                            ilTBIndex = ilTBIndex + 1
        '                        End If
        '                    Else    'Add time for each valid day
        '                        For ilIndex = ilFirstDay To ilLastDay Step 1
        '                            If (tmCff.iDay(ilIndex) = 1) Or (tmCff.sXDay(ilIndex) = "Y") Then
        '                                If tmLnRdf.sWkDays(ilLoop, ilIndex + 1) = "Y" Then
        '                                    gUnpackTime tmLnRdf.iStartTime(0, ilLoop), tmLnRdf.iStartTime(1, ilLoop), "A", "1", slLnStart
        '                                    gUnpackTime tmLnRdf.iEndTime(0, ilLoop), tmLnRdf.iEndTime(1, ilLoop), "A", "1", slLnEnd
        '                                    lmTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLnStart, False))
        '                                    lmTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLnEnd, True))
        '                                    ilTBIndex = ilTBIndex + 1
        '                                End If
        '                            End If
        '                        Next ilIndex
        '                    End If
         '               End If
        '            Next ilLoop
        '        Else
        '            gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slLnStart
        '            gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slLnEnd
        '            lmTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLnStart, False))
        '            lmTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLnEnd, True))
        '        End If
        '    End If
        '    mReadChfClfRdfCffRec = True
        'Else
        '    mReadChfClfRdfCffRec = True
        'End If
    Else
        mReadChfClfRdfCffRec = False
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadChfClfRdfRec               *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read records                   *
'*                                                     *
'*******************************************************
Function mReadChfClfRdfRec(llChfCode As Long, ilLineNo As Integer) As Integer
'
'   iRet = mReadChfClfRdfRec(llChfCode, ilLineNo)
'   Where:
'       llChfCode (I) - Contract code
'       ilLineNo(I)- Line number
'       iRet (O)- True if all records read,
'                 False if error in read
'
    Dim ilRet As Integer
    'If llChfCode <> tmChf.lCode Then
        tmChfSrchKey.lCode = llChfCode
        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then
            mReadChfClfRdfRec = False
            Exit Function
        End If
    'End If
    'If (tmClf.lChfCode <> llChfCode) Or (tmClf.iLine <> ilLineNo) Then
        tmClfSrchKey.lChfCode = llChfCode
        tmClfSrchKey.iLine = ilLineNo
        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iLine = ilLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))    'And (tmClf.sSchStatus = "A")
            ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    'Else
    '    ilRet = BTRV_ERR_NONE
    'End If
    If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = llChfCode) And (tmClf.iLine = ilLineNo) Then
        If tmLnRdf.iCode <> tmClf.iRdfCode Then
            tmRdfSrchKey.iCode = tmClf.iRdfCode  ' Rate card program/time File Code
            ilRet = btrGetEqual(hmRdf, tmLnRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet <> BTRV_ERR_NONE Then
                mReadChfClfRdfRec = False
                Exit Function
            End If
        End If
        mReadChfClfRdfRec = True
    Else
        mReadChfClfRdfRec = False
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mRemoveAvail                    *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get avail within Ssf           *
'*                                                     *
'*******************************************************
Function mRemoveAvail(slSchDate As String, slTime As String, ilGameNo As Integer) As Integer
    Dim ilRet As Integer
    Dim llTime As Long
    Dim llATime As Long
    Dim ilLoop As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilEvt As Integer
    Dim ilType As Integer

    ilType = ilGameNo
    llTime = gTimeToCurrency(slTime, False)
    If ilGameNo = 0 Then
        imSelectedDay = gWeekDayStr(slSchDate)
    Else
        imSelectedDay = 0
    End If
    imSsfRecLen = Len(tmSsf(imSelectedDay)) 'Max size of variable length record
    'tmSsfSrchKey.sType = "O" 'slType-On Air
    tmSsfSrchKey.iType = ilType 'slType-On Air
    tmSsfSrchKey.iVefCode = imVefCode
    gPackDate slSchDate, ilDate0, ilDate1
    tmSsfSrchKey.iDate(0) = ilDate0
    tmSsfSrchKey.iDate(1) = ilDate1
    tmSsfSrchKey.iStartTime(0) = 0
    tmSsfSrchKey.iStartTime(1) = 0
    ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    'Do While (ilRet = BTRV_ERR_NONE) And (tmSsf(imSelectedDay).sType = "O") And (tmSsf(imSelectedDay).iVefCode = imvefCode) And (tmSsf(imSelectedDay).iDate(0) = ilDate0) And (tmSsf(imSelectedDay).iDate(1) = ilDate1)
    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf(imSelectedDay).iType = ilType) And (tmSsf(imSelectedDay).iVefCode = imVefCode) And (tmSsf(imSelectedDay).iDate(0) = ilDate0) And (tmSsf(imSelectedDay).iDate(1) = ilDate1)
        For ilLoop = 1 To tmSsf(imSelectedDay).iCount Step 1
            tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilLoop)
            If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llATime
                If llTime = llATime Then
                    If (tmAvail.ianfCode = igPLAnfCode) And (tmAvail.iNoSpotsThis = 0) Then
                        'Remove avail
                        ilRet = gSSFGetPosition(hmSsf, lmSsfRecPos(imSelectedDay))
                        Do
                            imSsfRecLen = Len(tmSsf(imSelectedDay))
                            ilRet = gSSFGetDirect(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, lmSsfRecPos(imSelectedDay), INDEXKEY0, BTRV_LOCK_NONE)
                            If ilRet <> BTRV_ERR_NONE Then
                                mRemoveAvail = False
                                Exit Function
                            End If
                            ilRet = gGetByKeyForUpdateSSF(hmSsf, tmSsf(imSelectedDay))
                            If ilRet <> BTRV_ERR_NONE Then
                                mRemoveAvail = False
                                Exit Function
                            End If
                            'Move events donw and added avail
                            For ilEvt = ilLoop To tmSsf(imSelectedDay).iCount - 1 Step 1
                                tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilEvt) = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilEvt + 1)
                            Next ilEvt
                            tmSsf(imSelectedDay).iCount = tmSsf(imSelectedDay).iCount - 1
                            imSsfRecLen = 17 + tmSsf(imSelectedDay).iCount * Len(tmProg)
                            ilRet = gSSFUpdate(hmSsf, tmSsf(imSelectedDay), imSsfRecLen)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilRet <> BTRV_ERR_NONE Then
                            mRemoveAvail = False
                            Exit Function
                        End If
                    End If
                    mRemoveAvail = True
                    Exit Function
                End If
            End If
        Next ilLoop
        imSsfRecLen = Len(tmSsf(imSelectedDay)) 'Max size of variable length record
        ilRet = gSSFGetNext(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    mRemoveAvail = True
    Exit Function
End Function
