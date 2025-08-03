Attribute VB_Name = "SSFCHECKSubs"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of SSFCHECK.BAS on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: SSFCheck.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the ssfcheck function
'
'
Option Explicit
Option Compare Text
' Vehicle File
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer     'VEF record length
'Contract record information
Dim hmCHF As Integer        'Contract header file handle
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim tmChfSrchKey1 As CHFKEY1
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image
Dim tmTChf As CHF
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
'Spot record
Dim hmSdf As Integer
Dim tmSdf As SDF
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey As SDFKEY1
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

Dim hmCgf As Integer
Dim tmCgf As CGF
Dim imCgfRecLen As Integer
Dim tmCgfSrchKey1 As CGFKEY1    'CntrNo; CntRevNo; PropVer
Dim tmCgfCff() As CFF
'Required by gMakeSsf
Dim hmSsf As Integer
Dim tmTstSsf As SSF                'SSF record image
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
Dim lmSdfRecPos As Long
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


'Dim tmRec As LPOPREC
'*******************************************************
'*                                                     *
'*      Procedure Name:gSynchSdfSsf                    *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Synch Sdf and Ssf               *
'*                                                     *
'*******************************************************
Function gSynchSdfSsf(ilTestOnly As Integer, hlSsf As Integer, hlSdf As Integer, hlSmf As Integer, hlVef As Integer, hlChf As Integer, hlClf As Integer, hlCff As Integer, hlRdf As Integer, hlCrf As Integer, hlCgf As Integer, ilType As Integer, ilVefCode As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, slMOrO As String, hlSxf As Integer, Optional hlGsf As Integer = 0, Optional hlGhf As Integer = 0) As Integer
'
'   ilRet = gSynchSdfSsf(ilTestOnly, hlSsf, hlSdf, hlSmf, slType, ilVefCode, ilLogDate0, ilLogDate1)
'   Where:
'       ilTestOnly(I)- True=Test SSF only; False=Test and Update
'       hlSsf (I)- SSF handle (obtained from CBtrvTable)
'       hlSdf (I)- SDF handle (obtained from CBtrvTable)
'       hlSmf (I)- SMF handle
'       slType (I)- "O" = On Air; "A" = Alternate
'       ilVefCode (I)- Vehicle code
'       ilLogDate0 (I)- Log date to be checked and created
'       ilLogDate1
'
'    Dim tlSSf As SSF                'SSF record image
'
'    ilRet = True(ok) or False(error)
'
    Dim tlSsfSrchKey As SSFKEY0     'SSF key record image
    Dim ilSsfRecLen As Integer      'SSF record length
    Dim llSdfRecLoc As Long
    Dim llSdfRecPos As Long
    Dim llClfRecPos As Long
    Dim llCffRecPos As Long
    Dim ilSpotFound As Integer
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilIndex As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilTime0 As Integer
    Dim ilTime1 As Integer
    Dim ilError As Integer
    Dim ilMadeMissed As Integer
    Dim ilSpotTime0 As Integer
    Dim ilSpotTime1 As Integer
    Dim ilOrigSchVef As Integer
    Dim ilOrigGameNo As Integer
    Dim slTime As String
    Dim slDate As String
    Dim slMonDate As String
    Dim slMsg As String
    Dim llCntrNo As Long
    Dim llChfCode As Long
    Dim ilAdfCode As Integer
    Dim llChfRecPos As Long
    Dim llDate As Long
    Dim llMonDate As Long
    Dim llStartMDate As Long
    Dim llSdfCode As Long
    Dim slOrigSchStatus As String
    Dim slXSpotType As String
    Dim ilBBLen As Integer
    ReDim sgSSFErrorMsg(0 To 0) As String
    imVefRecLen = Len(tmVef)  'Get and save VEF record length
    hmSsf = hlSsf
    hmSdf = hlSdf
    imSdfRecLen = Len(tmSdf)
    hmSmf = hlSmf
    imSmfRecLen = Len(tmSmf)
    hmVef = hlVef
    imVefRecLen = Len(tmVef)
    hmCHF = hlChf
    imCHFRecLen = Len(tmChf)
    hmClf = hlClf
    imClfRecLen = Len(tmClf)
    hmCff = hlCff
    imCffRecLen = Len(tmCff)
    hmRdf = hlRdf
    imRdfRecLen = Len(tmLnRdf)
    hmCrf = hlCrf
    imCgfRecLen = Len(tmCgf)
    hmCgf = hlCgf
    imVefCode = ilVefCode
    imVpfIndex = gVpfFindIndex(imVefCode)
    tmVefSrchKey.iCode = ilVefCode
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        tmVef.sName = "Vehicle Missing"
    End If
    gUnpackDate ilLogDate0, ilLogDate1, slDate
    'Loop on spots
    tmTstSsf.iVefCode = 0  'Force read of Ssf
    imSdfRecLen = Len(tmSdf)
    tmSdfSrchKey.iVefCode = ilVefCode
    tmSdfSrchKey.iDate(0) = ilLogDate0
    tmSdfSrchKey.iDate(1) = ilLogDate1
    tmSdfSrchKey.iTime(0) = 0
    tmSdfSrchKey.iTime(1) = 0
    tmSdfSrchKey.sSchStatus = " "
    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.iDate(0) = ilLogDate0) And (tmSdf.iDate(1) = ilLogDate1)
        ilMadeMissed = False
        ilRet = btrGetPosition(hmSdf, llSdfRecLoc)
        lmSdfRecPos = llSdfRecLoc
        llSdfCode = tmSdf.lCode
        slOrigSchStatus = tmSdf.sSchStatus
        If ((tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C")) And ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O")) Then
            'Find match record within Ssf
            ilSpotFound = False
            ilSsfRecLen = Len(tmTstSsf)
            ilType = tmSdf.iGameNo
            tlSsfSrchKey.iType = ilType
            tlSsfSrchKey.iVefCode = ilVefCode
            tlSsfSrchKey.iDate(0) = ilLogDate0
            tlSsfSrchKey.iDate(1) = ilLogDate1
            tlSsfSrchKey.iStartTime(0) = 0
            tlSsfSrchKey.iStartTime(1) = 0
            'If (tmTstSsf.iVefCode = ilVefCode) And (tmTstSsf.iDate(0) = ilLogDate0) And (tmTstSsf.iDate(1) = ilLogDate1) And (tmTstSsf.iStartTime(0) = 0) And (tmTstSsf.iStartTime(1) = 0) Then
            '    ilRet = BTRV_ERR_NONE
            'Else
                ilRet = gSSFGetGreaterOrEqual(hmSsf, tmTstSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            'End If
            Do While (ilRet = BTRV_ERR_NONE) And (tmTstSsf.iType = ilType) And (tmTstSsf.iVefCode = ilVefCode) And (tmTstSsf.iDate(0) = ilLogDate0) And (tmTstSsf.iDate(1) = ilLogDate1)
                ilSpotTime0 = tmSdf.iTime(0)
                ilSpotTime1 = tmSdf.iTime(1)
                For ilIndex = 1 To tmTstSsf.iCount Step 1
                    LSet tgSpot = tmTstSsf.tPas(ADJSSFPASBZ + ilIndex)
                    If ((tgSpot.iRecType And &HF) >= 10) Then
                        If tgSpot.lSdfCode = tmSdf.lCode Then
                            ilSpotFound = True
                            '7/30/97: Bypass Checking of dates/times if altered in post log
                            If (Trim$(tmSdf.sAffChg) = "") Or (tmSdf.sAffChg = "N") Then
                                ilRet = mTestDateTime(ilTestOnly, ilLogDate0, ilLogDate1, tgAvail, slMOrO, ilMadeMissed, hlSxf, hlGsf, hlGhf)
                            Else
                                ilRet = True
                            End If
                            Exit For
                        End If
                    ElseIf (tgSpot.iRecType >= 2) And (tgSpot.iRecType <= 9) Then
                        tgAvail = tmTstSsf.tPas(ADJSSFPASBZ + ilIndex)
                    End If
                Next ilIndex
                If ilSpotFound Then
                    Exit Do
                End If
                ilSsfRecLen = Len(tmTstSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmTstSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If (Not ilSpotFound) And (Trim$(tmSdf.sAffChg) <> "") And (tmSdf.sAffChg <> "N") Then
                'Test other dates within week
                slMonDate = gObtainPrevMonday(slDate)
                llMonDate = gDateValue(slMonDate)
                For llDate = llMonDate To llMonDate + 6 Step 1
                    ilSpotFound = False
                    ilSsfRecLen = Len(tmTstSsf)
                    gPackDateLong llDate, ilDate0, ilDate1
                    tlSsfSrchKey.iType = ilType
                    tlSsfSrchKey.iVefCode = ilVefCode
                    tlSsfSrchKey.iDate(0) = ilDate0
                    tlSsfSrchKey.iDate(1) = ilDate1
                    tlSsfSrchKey.iStartTime(0) = 0
                    tlSsfSrchKey.iStartTime(1) = 0
                    If (tmTstSsf.iVefCode = ilVefCode) And (tmTstSsf.iDate(0) = ilDate0) And (tmTstSsf.iDate(1) = ilDate1) And (tmTstSsf.iStartTime(0) = 0) And (tmTstSsf.iStartTime(1) = 0) Then
                        ilRet = BTRV_ERR_NONE
                    Else
                        ilRet = gSSFGetGreaterOrEqual(hmSsf, tmTstSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                    End If
                    Do While (ilRet = BTRV_ERR_NONE) And (tmTstSsf.iType = ilType) And (tmTstSsf.iVefCode = ilVefCode) And (tmTstSsf.iDate(0) = ilDate0) And (tmTstSsf.iDate(1) = ilDate1)
                        For ilIndex = 1 To tmTstSsf.iCount Step 1
                            LSet tgSpot = tmTstSsf.tPas(ADJSSFPASBZ + ilIndex)
                            If ((tgSpot.iRecType And &HF) >= 10) Then
                                If tgSpot.lSdfCode = tmSdf.lCode Then
                                    ilSpotFound = True
                                    'If date and time don't match and status and smf must not exist, create it
                                    'and change status
                                    ilRet = mTestDateTime(ilTestOnly, ilDate0, ilDate1, tgAvail, slMOrO, ilMadeMissed, hlSxf, hlGsf, hlGhf)
                                    Exit For
                                End If
                            ElseIf (tgSpot.iRecType >= 2) And (tgSpot.iRecType <= 9) Then
                                tgAvail = tmTstSsf.tPas(ADJSSFPASBZ + ilIndex)
                            End If
                        Next ilIndex
                        If ilSpotFound Then
                            Exit Do
                        End If
                        ilSsfRecLen = Len(tmTstSsf) 'Max size of variable length record
                        ilRet = gSSFGetNext(hmSsf, tmTstSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If ilSpotFound Then
                        Exit For
                    End If
                Next llDate
                If Not ilSpotFound Then
                    llStartMDate = gDateValue(gObtainStartStd(slMonDate))
                    For llDate = llStartMDate To llMonDate - 1 Step 1
                        ilSpotFound = False
                        ilSsfRecLen = Len(tmTstSsf)
                        gPackDateLong llDate, ilDate0, ilDate1
                        tlSsfSrchKey.iType = ilType
                        tlSsfSrchKey.iVefCode = ilVefCode
                        tlSsfSrchKey.iDate(0) = ilDate0
                        tlSsfSrchKey.iDate(1) = ilDate1
                        tlSsfSrchKey.iStartTime(0) = 0
                        tlSsfSrchKey.iStartTime(1) = 0
                        If (tmTstSsf.iVefCode = ilVefCode) And (tmTstSsf.iDate(0) = ilDate0) And (tmTstSsf.iDate(1) = ilDate1) And (tmTstSsf.iStartTime(0) = 0) And (tmTstSsf.iStartTime(1) = 0) Then
                            ilRet = BTRV_ERR_NONE
                        Else
                            ilRet = gSSFGetGreaterOrEqual(hmSsf, tmTstSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                        End If
                        Do While (ilRet = BTRV_ERR_NONE) And (tmTstSsf.iType = ilType) And (tmTstSsf.iVefCode = ilVefCode) And (tmTstSsf.iDate(0) = ilDate0) And (tmTstSsf.iDate(1) = ilDate1)
                            For ilIndex = 1 To tmTstSsf.iCount Step 1
                                LSet tgSpot = tmTstSsf.tPas(ADJSSFPASBZ + ilIndex)
                                If ((tgSpot.iRecType And &HF) >= 10) Then
                                    If tgSpot.lSdfCode = tmSdf.lCode Then
                                        ilSpotFound = True
                                        ilRet = mTestDateTime(ilTestOnly, ilDate0, ilDate1, tgAvail, slMOrO, ilMadeMissed, hlSxf, hlGsf, hlGhf)
                                       Exit For
                                    End If
                                ElseIf (tgSpot.iRecType >= 2) And (tgSpot.iRecType <= 9) Then
                                    tgAvail = tmTstSsf.tPas(ADJSSFPASBZ + ilIndex)
                                End If
                            Next ilIndex
                            If ilSpotFound Then
                                Exit Do
                            End If
                            ilSsfRecLen = Len(tmTstSsf) 'Max size of variable length record
                            ilRet = gSSFGetNext(hmSsf, tmTstSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If ilSpotFound Then
                            Exit For
                        End If
                    Next llDate
                End If
                If Not ilSpotFound Then
                    llStartMDate = gDateValue(gObtainEndStd(slMonDate))
                    For llDate = llMonDate + 7 To llStartMDate Step 1
                        ilSpotFound = False
                        ilSsfRecLen = Len(tmTstSsf)
                        gPackDateLong llDate, ilDate0, ilDate1
                        tlSsfSrchKey.iType = ilType
                        tlSsfSrchKey.iVefCode = ilVefCode
                        tlSsfSrchKey.iDate(0) = ilDate0
                        tlSsfSrchKey.iDate(1) = ilDate1
                        tlSsfSrchKey.iStartTime(0) = 0
                        tlSsfSrchKey.iStartTime(1) = 0
                        If (tmTstSsf.iVefCode = ilVefCode) And (tmTstSsf.iDate(0) = ilDate0) And (tmTstSsf.iDate(1) = ilDate1) And (tmTstSsf.iStartTime(0) = 0) And (tmTstSsf.iStartTime(1) = 0) Then
                            ilRet = BTRV_ERR_NONE
                        Else
                            ilRet = gSSFGetGreaterOrEqual(hmSsf, tmTstSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                        End If
                        Do While (ilRet = BTRV_ERR_NONE) And (tmTstSsf.iType = ilType) And (tmTstSsf.iVefCode = ilVefCode) And (tmTstSsf.iDate(0) = ilDate0) And (tmTstSsf.iDate(1) = ilDate1)
                            For ilIndex = 1 To tmTstSsf.iCount Step 1
                                LSet tgSpot = tmTstSsf.tPas(ADJSSFPASBZ + ilIndex)
                                If ((tgSpot.iRecType And &HF) >= 10) Then
                                    If tgSpot.lSdfCode = tmSdf.lCode Then
                                        ilSpotFound = True
                                        ilRet = mTestDateTime(ilTestOnly, ilDate0, ilDate1, tgAvail, slMOrO, ilMadeMissed, hlSxf, hlGsf, hlGhf)
                                        Exit For
                                    End If
                                ElseIf (tgSpot.iRecType >= 2) And (tgSpot.iRecType <= 9) Then
                                    tgAvail = tmTstSsf.tPas(ADJSSFPASBZ + ilIndex)
                                End If
                            Next ilIndex
                            If ilSpotFound Then
                                Exit Do
                            End If
                            ilSsfRecLen = Len(tmTstSsf) 'Max size of variable length record
                            ilRet = gSSFGetNext(hmSsf, tmTstSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                        If ilSpotFound Then
                            Exit For
                        End If
                    Next llDate
                End If
            End If
            If Not ilSpotFound Then
                slXSpotType = ""
                ilMadeMissed = True
                ilSpotTime0 = tmSdf.iTime(0)
                ilSpotTime1 = tmSdf.iTime(1)
                gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                If Not ilTestOnly Then
                    ilRet = gRemoveSmf(hmSmf, tmSmf, tmSdf, hlSxf)   'resets missed date
                    If Not ilRet Then
                        '7/6/06: Force spot to be deleted since can't determine where it came from
                        slXSpotType = "X"
                        'gSynchSdfSsf = False
                        'Exit Function
                    End If
                Else
                    ilRet = gFindSmf(tmSdf, hlSmf, tmSmf)
                End If
                'tmChfSrchKey.lCode = tmSdf.lChfCode
                'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                'If (tmChf.sType = "T") Or (tmChf.sType = "Q") Or (tmChf.sType = "S") Or (tmChf.sType = "M") Then
                If tmSdf.sSpotType = "X" Then
                    slXSpotType = "X"
                    If ((tmSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                        slXSpotType = ""
                    End If
                Else
                    '7/6/05: Set above
                    'slXSpotType = ""
                End If
                'If (tmSdf.sSpotType = "T") Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "S") Or (tmSdf.sSpotType = "M") Or (slXSpotType = "X") Then
                If ((tmSdf.sSpotType = "T") And (tgSpf.sSchdRemnant = "Y")) Or (tmSdf.sSpotType = "Q") Or ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA = "Y")) Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo = "Y")) Or (slXSpotType = "X") Then
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilTestOnly Then
                        slMsg = "Can't Find SSF Spot: "
                    Else
                        slMsg = "Can't Find SSF Spot, Removed: "
                    End If
                    If ilRet = BTRV_ERR_NONE Then
                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                    Else
                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " ChfCode=" & str$(tmSdf.lChfCode) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                    End If
                    ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                    If Not ilTestOnly Then
                        ilRet = btrGetPosition(hmSdf, llSdfRecPos)
                        Do
                            'tmSRec = tmSdf
                            'ilCRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                            'tmSdf = tmSRec
                            'If ilCRet <> BTRV_ERR_NONE Then
                            '    igBtrError = ilCRet
                            '    sgErrLoc = "gSynchSdfSsf- Get by Key Sdf(4)"
                            '    gSynchSdfSsf = False
                            '    Exit Function
                            'End If
                            ilRet = btrDelete(hmSdf)
                            If ilRet = BTRV_ERR_CONFLICT Then
                                ilCRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                    End If
                Else
                    ilDate0 = tmSdf.iDate(0)
                    ilDate1 = tmSdf.iDate(1)
                    ilTime0 = tmSdf.iTime(0)
                    ilTime1 = tmSdf.iTime(1)
                    ilOrigSchVef = tmSdf.iVefCode
                    ilOrigGameNo = tmSdf.iGameNo
                    tmChfSrchKey.lCode = tmSdf.lChfCode
                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilTestOnly Then
                        slMsg = "Can't Find SSF Spot: "
                    Else
                        slMsg = "Can't Find SSF Spot, Missed: "
                    End If
                    If ilRet = BTRV_ERR_NONE Then
                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                    Else
                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " ChfCode=" & str$(tmSdf.lChfCode) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                    End If
                    ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                    If Not ilTestOnly Then
                        'Update Sdf record
                        Do
                            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecLoc, INDEXKEY0, BTRV_LOCK_NONE)
                            'tmSRec = tmSdf
                            'ilRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                            'tmSdf = tmSRec
                            'If ilRet <> BTRV_ERR_NONE Then
                            '    igBtrError = ilRet
                            '    sgErrLoc = "gSynchSdfSsf- Get by Key Sdf(5)"
                            '    gSynchSdfSsf = False
                            '    Exit Function
                            'End If
                            tmSdf.sSchStatus = "M"
                            tmSdf.iMnfMissed = igMnfMissed
                            tmSdf.iDate(0) = ilDate0
                            tmSdf.iDate(1) = ilDate1
                            tmSdf.iTime(0) = ilTime0
                            tmSdf.iTime(1) = ilTime1
                            tmSdf.iVefCode = ilOrigSchVef
                            tmSdf.iGameNo = ilOrigGameNo
                            tmSdf.lSmfCode = 0
                            tmSdf.sXCrossMidnight = "N"
                            tmSdf.sWasMG = "N"
                            tmSdf.sFromWorkArea = "N"
                            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                        'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                        ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                    End If
                End If
            End If
        Else
            'Test if ChfCode or advertiser or length in error
            tmChfSrchKey.lCode = tmSdf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                llCntrNo = tmChf.lCntrNo
                ilAdfCode = tmChf.iAdfCode
                llChfCode = tmChf.lCode
                If tmChf.sDelete = "Y" Then
                    tmChfSrchKey1.lCntrNo = tmChf.lCntrNo
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llCntrNo) And (tmChf.sSchStatus <> "F") And (tmChf.sSchStatus <> "M") And (tmChf.sDelete = "Y")
                        If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.sDelete <> "Y") Then
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (tmChf.lCntrNo = llCntrNo) And (tmChf.sDelete <> "Y") Then
                        llChfCode = tmChf.lCode
                        ilAdfCode = tmChf.iAdfCode
                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Contract Code Error:" & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
                        ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                    Else
                        tmChfSrchKey.lCode = tmSdf.lChfCode
                        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            'ilRemoveSpot = True
                            tmChfSrchKey1.lCntrNo = tmChf.lCntrNo
                            tmChfSrchKey1.iCntRevNo = 32000
                            tmChfSrchKey1.iPropVer = 32000
                            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llCntrNo)
                                If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) Then
                                    Exit Do
                                End If
                                ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                            If (ilRet <> BTRV_ERR_NONE) Or (tmChf.lCntrNo <> llCntrNo) Then
                                tmChfSrchKey.lCode = tmSdf.lChfCode
                                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            End If
                            llChfCode = tmChf.lCode
                            ilAdfCode = tmChf.iAdfCode
                            llCntrNo = tmChf.lCntrNo
                            sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Contract Delete Flag Error:" & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
                            ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                            If Not ilTestOnly Then
                                ilRet = btrGetPosition(hmCHF, llChfRecPos)
                                Do
                                    ilRet = btrGetDirect(hmCHF, tmChf, imCHFRecLen, llChfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                    'tmSRec = tmChf
                                    'ilRet = gGetByKeyForUpdate("CHF", hmChf, tmSRec)
                                    'tmChf = tmSRec
                                    tmChf.sDelete = "N"
                                    ilRet = btrUpdate(hmCHF, tmChf, imCHFRecLen)
                                Loop While ilRet = BTRV_ERR_CONFLICT
                            End If
                        End If
                    End If
                End If
                tmClfSrchKey.lChfCode = llChfCode
                tmClfSrchKey.iLine = tmSdf.iLineNo
                tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                If (ilRet <> BTRV_ERR_NONE) Or (tmClf.lChfCode <> llChfCode) Or (tmClf.iLine <> tmSdf.iLineNo) Then
                    'Determine if Line Not moved to correct header
                    tmChfSrchKey1.lCntrNo = llCntrNo
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmTChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmTChf.lCntrNo = llCntrNo)
                        tmClfSrchKey.lChfCode = tmTChf.lCode
                        tmClfSrchKey.iLine = tmSdf.iLineNo
                        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmTChf.lCode) And (tmClf.iLine = tmSdf.iLineNo) And (tmClf.sDelete <> "Y") Then
                            'Set Line and Flight to Current header
                            If tmTChf.lCode <> llChfCode Then
                                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Line Contract Code Error:" & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
                                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                If ilTestOnly Then
                                    Exit Do
                                End If
                                ilRet = btrGetPosition(hmClf, llClfRecPos)
                                Do
                                    ilRet = btrGetDirect(hmClf, tmClf, imClfRecLen, llClfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                    'tmSRec = tmClf
                                    'ilRet = gGetByKeyForUpdate("CLF", hmClf, tmSRec)
                                    'tmClf = tmSRec
                                    'If ilCRet <> BTRV_ERR_NONE Then
                                    '    Exit Do
                                    'End If
                                    ilRet = btrDelete(hmClf)
                                Loop While ilRet = BTRV_ERR_CONFLICT
                                tmClf.lChfCode = llChfCode
                                ilRet = btrInsert(hmClf, tmClf, imClfRecLen, INDEXKEY0)
                                tmCffSrchKey.lChfCode = tmTChf.lCode
                                tmCffSrchKey.iClfLine = tmSdf.iLineNo
                                tmCffSrchKey.iCntRevNo = tmClf.iCntRevNo
                                tmCffSrchKey.iPropVer = tmClf.iPropVer
                                tmCffSrchKey.iStartDate(0) = 0
                                tmCffSrchKey.iStartDate(1) = 0
                                ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                Do While (ilRet = BTRV_ERR_NONE) And (tmCff.lChfCode = tmTChf.lCode) And (tmCff.iClfLine = tmSdf.iLineNo)
                                    ilRet = btrGetPosition(hmCff, llCffRecPos)
                                    Do
                                        ilRet = btrGetDirect(hmCff, tmCff, imCffRecLen, llCffRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                        'tmSRec = tmCff
                                        'ilRet = gGetByKeyForUpdate("CFF", hmCff, tmSRec)
                                        'tmCff = tmSRec
                                        'If ilCRet <> BTRV_ERR_NONE Then
                                        '    Exit Do
                                        'End If
                                        ilRet = btrDelete(hmCff)
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    tmCff.lChfCode = llChfCode
                                    tmCff.lCode = 0
                                    ilRet = btrInsert(hmCff, tmCff, imCffRecLen, INDEXKEY1)
                                    tmCffSrchKey.lChfCode = tmTChf.lCode
                                    tmCffSrchKey.iClfLine = tmSdf.iLineNo
                                    tmCffSrchKey.iCntRevNo = tmClf.iCntRevNo
                                    tmCffSrchKey.iPropVer = tmClf.iPropVer
                                    tmCffSrchKey.iStartDate(0) = 0
                                    tmCffSrchKey.iStartDate(1) = 0
                                    ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                Loop
                            End If
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmCHF, tmTChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
                If (tmClf.lChfCode = llChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                    If (tmSdf.lChfCode <> llChfCode) Then
                        If ilTestOnly Then
                            slMsg = "Contract Code Error: "
                        Else
                            slMsg = "Contract Error Fixed: "
                        End If
                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                        ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                    End If
                    If (tmSdf.iAdfCode <> ilAdfCode) Then
                        If ilTestOnly Then
                            slMsg = "Advertiser Name Error: "
                        Else
                            slMsg = "Advertiser Name Error Fixed: "
                        End If
                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                        ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                    End If
                    If tmSdf.sSpotType = "O" Then
                        ilBBLen = tmClf.iBBOpenLen
                        If (tmSdf.iLen <> ilBBLen) Then
                            If ilTestOnly Then
                                slMsg = "OBB Spot Length Error: "
                            Else
                                slMsg = "OBB Spot Length Error Fixed: "
                            End If
                            sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                            ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                        End If
                    ElseIf tmSdf.sSpotType = "C" Then
                        If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                            ilBBLen = tmClf.iBBOpenLen
                        Else
                            ilBBLen = tmClf.iBBCloseLen
                        End If
                        If (tmSdf.iLen <> ilBBLen) Then
                            If ilTestOnly Then
                                slMsg = "CBB Spot Length Error: "
                            Else
                                slMsg = "CBB Spot Length Error Fixed: "
                            End If
                            sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                            ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                        End If
                    Else
                        If (tmSdf.iLen <> tmClf.iLen) Then
                            If ilTestOnly Then
                                slMsg = "Spot Length Error: "
                            Else
                                slMsg = "Spot Length Error Fixed: "
                            End If
                            sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                            ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                        End If
                    End If
                    If Not ilTestOnly Then
                        If (tmSdf.lChfCode <> llChfCode) Or (tmSdf.iAdfCode <> ilAdfCode) Or (tmSdf.iLen <> tmClf.iLen) Then
                            'Update Sdf record
                            Do
                                ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecLoc, INDEXKEY0, BTRV_LOCK_NONE)
                                'tmSRec = tmSdf
                                'ilRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                                'tmSdf = tmSRec
                                'If ilRet <> BTRV_ERR_NONE Then
                                '    igBtrError = ilRet
                                '    sgErrLoc = "gSynchSdfSsf- Get by Key Sdf(6)"
                                '    gSynchSdfSsf = False
                                '    Exit Function
                                'End If
                                tmSdf.lChfCode = llChfCode
                                tmSdf.iAdfCode = ilAdfCode
                                If tmSdf.sSpotType = "O" Then
                                    tmSdf.iLen = tmClf.iBBOpenLen
                                ElseIf tmSdf.sSpotType = "C" Then
                                    If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                                        tmSdf.iLen = tmClf.iBBOpenLen
                                    Else
                                        tmSdf.iLen = tmClf.iBBCloseLen
                                    End If
                                Else
                                    tmSdf.iLen = tmClf.iLen
                                End If
                                ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen)
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            'Reposition for GetNext
                            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecLoc, INDEXKEY1, BTRV_LOCK_NONE)
                        End If
                    End If
                Else
                    If ilTestOnly Then
                        slMsg = "Spot Line Missing Error: "
                    Else
                        slMsg = "Spot Line Missing, Removed: "
                    End If
                    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                    ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                    If Not ilTestOnly Then
                        'Delete Sdf record
                        Do
                            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, llSdfRecLoc, INDEXKEY0, BTRV_LOCK_NONE)
                            'tmSRec = tmSdf
                            'ilRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                            'tmSdf = tmSRec
                            'If ilRet <> BTRV_ERR_NONE Then
                            '    igBtrError = ilRet
                            '    sgErrLoc = "gSynchSdfSsf- Get by Key Sdf(7)"
                            '    gSynchSdfSsf = False
                            '    Exit Function
                            'End If
                            ilRet = btrDelete(hmSdf)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        ilMadeMissed = True
                    End If
                End If
            End If
        End If
        If Not ilTestOnly Then
            If ilMadeMissed Then
                tmSdfSrchKey.iVefCode = ilVefCode
                tmSdfSrchKey.iDate(0) = ilLogDate0
                tmSdfSrchKey.iDate(1) = ilLogDate1
                tmSdfSrchKey.iTime(0) = ilSpotTime0
                tmSdfSrchKey.iTime(1) = ilSpotTime1
                tmSdfSrchKey.sSchStatus = " "
                ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                If tmSdf.lCode = llSdfCode Then
                    ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                End If
            Else
                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            End If
        Else
            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        End If
    Loop
    gSynchSdfSsf = True
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
Function mAvailRoom(ilAvailIndex As Integer, ilSameAvail As Integer) As Integer
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
    ilSameAvail = False
   LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilAvailIndex)
    ilAvailUnits = tmAvail.iAvInfo And &H1F
    ilAvailSec = tmAvail.iLen
    For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tmAvail.iNoSpotsThis Step 1
       LSet tmSpot = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilSpotIndex)
        If tmSpot.lSdfCode = tmSdf.lCode Then
            ilSameAvail = True
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
            If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                ilUnitsSold = ilUnitsSold + ilSpotUnits
                ilSecSold = ilSecSold + ilSpotLen
            End If
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
Function mFindAvail(slSchDate As String, slFindTime As String, ilFindAdjAvail As Integer, ilAvailIndex As Integer) As Integer
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
    imSelectedDay = gWeekDayStr(slSchDate)
    lmSsfDate(imSelectedDay) = 0
    ilRet = gObtainSsfForDateOrGame(imVefCode, llSchDate, slFindTime, 0, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay))
    If Not ilRet Then
        mFindAvail = False
        Exit Function
    End If
    llFndAdjTime = -1
    For ilLoop = 1 To tmSsf(imSelectedDay).iCount Step 1
       LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilLoop)
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
'           slRet(O)-   "" = Abort move
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
    Dim ilGameNo As Integer
    Dim tlSmf As SMF
    ReDim tlCff(0 To 1) As CFF
    Dim llEarliestAllowedDate As Long
    Dim slMissedDate As String

    ilGameNo = tmSdf.iGameNo
    If (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then
        If gFindSmf(tmSdf, hmSmf, tlSmf) Then
            ilGameNo = tlSmf.iGameNo
        End If
    End If
    ilRet = mReadChfClfRdfCffRec(tmSdf.lChfCode, tmSdf.iLineNo, slMoveDate, ilGameNo)
    If Not ilRet Then
        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slWkDate
        gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Unable to Read Contract " & slWkDate & " " & slTime & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
        ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
        mMoveTest = ""  'Abort
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
            sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Unable to find SMF " & slWkDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
            ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
            mMoveTest = ""  'Abort
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
    'gGetLineSchParameters hmSsf, tmSsf(), lmSsfDate(), lmSsfRecPos(), llDate, imVefCode, tmChf.iAdfCode, tmSdf.iGameNo, tlCff(), tmClf, tmLnRdf, lmSepLength, lmStartDateLen, lmEndDateLen, llChfCode, ilLineNo, ilAdfCode, ilVehComp, imHour(), imDay(), imQH(), imAHour(), imADay(), imAQH(), lmTBStartTime(), lmTBEndTime(), llEarliestAllowedDate, imSkip(), tmChf.sType, tmChf.iPctTrade, imBkQH, False, imPriceLevel
    gGetLineSchParameters hmSsf, tmSsf(), lmSsfDate(), lmSsfRecPos(), llDate, imVefCode, tmChf.iAdfCode, ilGameNo, tlCff(), tmClf, tmLnRdf, lmSepLength, lmStartDateLen, lmEndDateLen, llChfCode, ilLineNo, ilAdfCode, ilVehComp, imHour(), imDay(), imQH(), imAHour(), imADay(), imAQH(), lmTBStartTime(), lmTBEndTime(), llEarliestAllowedDate, imSkip(), tmChf.sType, tmChf.iPctTrade, imBkQH, False, imPriceLevel, False
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
        'If (slType <> "S") And (slType <> "M") And (slType <> "T") And (slType <> "Q") Then
        If (slType <> "S") And (slType <> "M") And (slType <> "T") And (slType <> "Q") Or ((slType = "M") And (tgSpf.sSchdPromo = "Y")) Or ((slType = "S") And (tgSpf.sSchdPSA = "Y")) Or ((slType = "T") And (tgSpf.sSchdRemnant = "Y")) Then
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
Function mReadChfClfRdfCffRec(llChfCode As Long, ilLineNo As Integer, slSpotDate As String, ilGameNo As Integer) As Integer
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
    llTime = gTimeToCurrency(slTime, False)
    imSelectedDay = gWeekDayStr(slSchDate)
    imSsfRecLen = Len(tmSsf(imSelectedDay)) 'Max size of variable length record
    '11/24/12
    tmSsfSrchKey.iType = ilGameNo   '0 'slType-On Air
    tmSsfSrchKey.iVefCode = imVefCode
    gPackDate slSchDate, ilDate0, ilDate1
    tmSsfSrchKey.iDate(0) = ilDate0
    tmSsfSrchKey.iDate(1) = ilDate1
    tmSsfSrchKey.iStartTime(0) = 0
    tmSsfSrchKey.iStartTime(1) = 0
    ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf(imSelectedDay), imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    'Do While (ilRet = BTRV_ERR_NONE) And (tmSsf(imSelectedDay).iType = 0) And (tmSsf(imSelectedDay).iVefCode = imVefCode) And (tmSsf(imSelectedDay).iDate(0) = ilDate0) And (tmSsf(imSelectedDay).iDate(1) = ilDate1)
    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf(imSelectedDay).iType = ilGameNo) And (tmSsf(imSelectedDay).iVefCode = imVefCode) And (tmSsf(imSelectedDay).iDate(0) = ilDate0) And (tmSsf(imSelectedDay).iDate(1) = ilDate1)
        For ilLoop = 1 To tmSsf(imSelectedDay).iCount Step 1
           LSet tmAvail = tmSsf(imSelectedDay).tPas(ADJSSFPASBZ + ilLoop)
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
'*******************************************************
'*                                                     *
'*      Procedure Name:mTestDateTime                   *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Synch Sdf and Ssf               *
'*                                                     *
'*******************************************************
Function mTestDateTime(ilTestOnly, ilLogDate0 As Integer, ilLogDate1 As Integer, tlAvail As AVAILSS, slMOrO As String, ilMovedSpot As Integer, hlSxf As Integer, hlGsf As Integer, hlGhf As Integer)
    Dim ilDateTimeDisparity As Integer
    Dim slSchDate As String
    Dim slSchTime As String
    Dim slAirDate As String
    Dim slAirTime As String
    Dim slOrigTime As String
    Dim ilFindAdjAvail As Integer
    Dim ilAvailIndex As Integer
    Dim ilRet As Integer
    Dim ilBkQH As Integer
    Dim slRet As String
    Dim slDate As String
    Dim slTime As String
    Dim ilSameAvail As Integer
    Dim ilGameNo As Integer
    Dim tlSdf1 As SDF
    ilMovedSpot = False
    ilDateTimeDisparity = False
    imSdfRecLen = Len(tmSdf)
    If (tlAvail.iTime(0) <> tmSdf.iTime(0)) Or (tlAvail.iTime(1) <> tmSdf.iTime(1)) Then
        ilDateTimeDisparity = True
    Else
        If (ilLogDate0 <> tmSdf.iDate(0)) Or (ilLogDate1 <> tmSdf.iDate(1)) Then
            ilDateTimeDisparity = True
        End If
    End If
    If ilDateTimeDisparity And (tmSdf.sSchStatus <> "G") And (tmSdf.sSchStatus <> "O") Then
        If ilTestOnly Then
            If (ilLogDate0 = tmSdf.iDate(0)) And (ilLogDate1 = tmSdf.iDate(1)) Then
                gUnpackDate ilLogDate0, ilLogDate1, slSchDate
                gUnpackTime tlAvail.iTime(0), tlAvail.iTime(1), "A", "1", slSchTime
                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slAirDate
                gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slAirTime
                ilFindAdjAvail = True
                slOrigTime = slSchTime
                ilRet = mFindAvail(slSchDate, slSchTime, ilFindAdjAvail, ilAvailIndex)
                ilRet = mFindAvail(slAirDate, slAirTime, ilFindAdjAvail, ilSameAvail)
                If ilAvailIndex <> ilSameAvail Then
                    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Spot Not Defined as MG or Outside or Moved as Required " & slAirDate & " " & slAirTime & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                    ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                End If
            Else
                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Spot Not Defined as MG or Outside or Moved as Required " & slDate & " " & slTime & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
            End If
        Else
            gUnpackDate ilLogDate0, ilLogDate1, slSchDate
            gUnpackTime tlAvail.iTime(0), tlAvail.iTime(1), "A", "1", slSchTime
            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slAirDate
            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slAirTime
            ilFindAdjAvail = True
            slOrigTime = slSchTime
            If Not mFindAvail(slAirDate, slAirTime, ilFindAdjAvail, ilAvailIndex) Then
                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Unable to Find Avail for MG or Outside " & slAirDate & " " & slAirTime & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                mTestDateTime = False
                Exit Function
            End If
            If Not mAvailRoom(ilAvailIndex, ilSameAvail) Then
                mTestDateTime = False
                Exit Function
            End If
            slRet = mMoveTest(slSchDate, slAirDate, slAirTime, slMOrO)
            If slRet = "" Then
                mTestDateTime = False
                Exit Function
            End If
            If (slRet = "S") And (ilSameAvail) Then
                mTestDateTime = True
                Exit Function
            End If
            ilMovedSpot = True
            'Unschedule, then schedule (gChgSchSpot removes Smf if exist)
            gPackDate slSchDate, tmSdf.iDate(0), tmSdf.iDate(1)
            gPackTime slOrigTime, tmSdf.iTime(0), tmSdf.iTime(1)
            ilRet = gChgSchSpot("TM", hmSdf, tmSdf, hmSmf, tmSdf.iGameNo, tmSmf, hmSsf, tmSsf(imSelectedDay), lmSsfDate(imSelectedDay), lmSsfRecPos(imSelectedDay), hlSxf, hlGsf, hlGhf)
            If Not ilRet Then
                mTestDateTime = False
                Exit Function
            End If
            ilGameNo = tmSdf.iGameNo
            If Not mRemoveAvail(slSchDate, slOrigTime, ilGameNo) Then
                mTestDateTime = False
                Exit Function
            End If
            If Not mFindAvail(slAirDate, slAirTime, ilFindAdjAvail, ilAvailIndex) Then
                mTestDateTime = False
                Exit Function
            End If
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
            tlSdf1 = tmSdf
            ilRet = gBookSpot(slRet, hmSdf, tmSdf, lmSdfRecPos, ilBkQH, hmSsf, tmSsf(imSelectedDay), lmSsfRecPos(imSelectedDay), ilAvailIndex, -1, tmChf, tmClf, tmLnRdf, imVpfIndex, hmSmf, tmSmf, hmClf, hmCrf, imPriceLevel, False, hlSxf, hlGsf)
            If Not ilRet Then
                mTestDateTime = False
                Exit Function
            End If
            'Reset copy
            ilRet = btrGetDirect(hmSdf, tmSdf, imSdfRecLen, lmSdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                mTestDateTime = False
                Exit Function
            End If
            'tmRec = tmSdf
            'ilRet = gGetByKeyForUpdate("SDF", hmSdf, tmRec)
            'tmSdf = tmRec
            'If ilRet <> BTRV_ERR_NONE Then
            '    mTestDateTime = False
            '    Exit Function
            'End If
            tmSdf.iRotNo = tlSdf1.iRotNo
            tmSdf.sPtType = tlSdf1.sPtType
            tmSdf.lCopyCode = tlSdf1.lCopyCode
            gPackDate slAirDate, tmSdf.iDate(0), tmSdf.iDate(1)
            gPackTime slAirTime, tmSdf.iTime(0), tmSdf.iTime(1)
            ilRet = btrUpdate(hmSdf, tmSdf, imSdfRecLen) '  Update File
            gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
            If tmSdf.sSchStatus = "G" Then
                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Spot Defined as MG at " & slDate & " " & slTime & " Cntr # " & Trim(str$(tmChf.lCntrNo)) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
            ElseIf tmSdf.sSchStatus = "O" Then
                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Spot Defined as Outside at " & slDate & " " & slTime & " Cntr # " & Trim(str$(tmChf.lCntrNo)) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
            Else
                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Spot Moved to " & slDate & " " & slTime & " Cntr # " & Trim(str$(tmChf.lCntrNo)) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
            End If
            ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
        End If
    End If
    mTestDateTime = True
    Exit Function
End Function
