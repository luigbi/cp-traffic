Attribute VB_Name = "SCHEDULE"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Schedule.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmChkChf                                                                              *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Schedule.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the schedule function
'
'
Option Explicit
Option Compare Text

'Required by gMakeSsf
Dim tmSsf As SSF                'SSF record image
'Dim tmSsfOld As SSF
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey1 As SSFKEY1      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmBBSpot As BBSPOTSS
Dim tmProgTest As PROGRAMSS
Dim tmAvailTest As AVAILSS
Dim tmSpotTest As CSPOTSS
Dim tmBBSpotTest As BBSPOTSS
'Spot record
Dim tmSdf As SDF
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey0 As SDFKEY0
Dim tmSdfSrchKey As SDFKEY1
Dim tmSdfSrchKey3 As LONGKEY0
'Library Calendar
Dim hmLcf As Integer
Dim tmLcf As LCF
Dim imLcfRecLen As Integer
Dim tmLcfSrchKey2 As LCFKEY2
'Advertiser
Dim tmAdf As ADF
'Avail name
Dim tmAnf As ANF
Dim tmAnfSrchKey As INTKEY0 'ANF key record image
Dim imAnfRecLen As Integer  'ANF record length
Dim tmSAnf() As ANF
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
Dim tmClfSrchKey1 As CLFKEY1 'CLF key record image
Dim tmClfSrchKey2 As LONGKEY0 'CLF key record image
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF            'CLF record image
Dim tmChkClf() As CLF
Dim tmXMidClf As CLF
'Contract record information
Dim hmCff As Integer        'Contract line Flight file handle
Dim tmCffSrchKey As CFFKEY0 'CFF key record image
Dim tmCffSrchKey1 As LONGKEY0 'CFF key record image
Dim imCffRecLen As Integer  'CFF record length
Dim tmCff As CFF            'CFF record image
Dim tmFCff() As CFF
Dim tmChkCFF() As CFF
'Comments
Dim hmCxf As Integer            'Comments file handle
Dim tmCxf As CXF               'CXF record image
Dim tmCxfSrchKey As LONGKEY0     'CXF key record image
Dim imCxfRecLen As Integer         'CXF record length
' Daypart
Dim hmRdf As Integer        'Daypart file handle
Dim tmRdf As RDF            'RDF record image
Dim tmRdfSrchKey As INTKEY0 'RDF key record image
Dim imRdfRecLen As Integer     'RDF record length
Dim tmXMidRdf As RDF
' Vehicle File
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer     'VEF record length
'Short Title
Dim hmSif As Integer            'Short Title file handle
Dim tmSif As SIF               'SIF record image
Dim imSifRecLen As Integer         'SIF record length
'Rotation
Dim hmCrf As Integer            'Short Title file handle
Dim tmCrf As CRF               'SIF record image
Dim imCrfRecLen As Integer         'SIF record length
'Spot MG record
Dim tmSmf As SMF
Dim imSmfRecLen As Integer
Dim tmSmfSrchKey2 As LONGKEY0
'Spot Track
Dim hmStf As Integer
Dim tmStf As STF
Dim imStfRecLen As Integer
' Virtual Vehicle File
Dim hmVsf As Integer        'Vehicle file handle
Dim tmVsf As VSF            'VSF record image
Dim tmVsfSrchKey As LONGKEY0 'VSF key record image
Dim imVsfRecLen As Integer     'VSF record length
Dim tmCTSsf As SSF               'Ssf for conflict test
'Feed
Dim hmFsf As Integer
Dim imFsfRecLen As Integer
Dim tmFsfSrchKey0 As LONGKEY0
Dim tmFsfSrchKey4 As FSFKEY4
Dim tmFsf As FSF
Dim tmTFsf As FSF
'Feed
Dim hmFnf As Integer
Dim imFnfRecLen As Integer
Dim tmFnf As FNF
'Feed
Dim hmPrf As Integer
Dim imPrfRecLen As Integer
Dim tmPrf As PRF
'Avail Lock
Dim hmAlf As Integer
Dim tmAlf As ALF
Dim tmAlfSrchkey1 As ALFKEY1
Dim tmAlfSrchKey2 As ALFKEY2
Dim imAlfRecLen As Integer

Dim tmATTSrchKey1 As INTKEY0     'ATT key 1 image
Dim imAttRecLen As Integer      'ATT record length
Dim hmAtt As Integer            'Agreement file handle

Dim tmSxf As SXF
Dim imSxfRecLen As Integer
Dim tmSxfSrchKey1 As SXFKEY1
Public tgSxfSdf As SDF

'Public tmSRec As LPOPREC
Dim tmTVlf() As VLF
'TTP 10496 - Affiliate alerts created when log is generated even if there's no spots
'Public sgLogStartDate As String
Public sgLogEndDate As String


'*******************************************************
'*                                                     *
'*      Procedure Name:gAdvtTest                       *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test for advertiser conflicts   *
'*                                                     *
'*******************************************************
Function gAdvtTest(hlSsf As Integer, tlSsf As SSF, llSsfRecPos As Long, tlSpotMove() As SPOTMOVE, ilVpfIndex As Integer, llSepLength As Long, ilAvailIndex As Integer, ilChfAdfCode As Integer, ilMnfComp0 As Integer, ilMnfComp1 As Integer, ilSchMode As Integer, ilBkQH As Integer, slInOut As String, slPreempt As String, ilPriceLevel As Integer, ilCheckAvail As Integer) As Integer
'
'   ilRet = gAdvtTest (tlSsf, llSsfRecPos, tlSpotMove(), llSepLength, ilAvailIndex, ilChfAdfCode)
'   Where:
'       tlSsf(I)- Ssf image
'       llSsfRecPos(I)- Ssf record position
'       tlSpotMove()(I)- Array of spots to bypass
'       llSepLength(I)- advertiser separation length (in seconds)
'       ilAvailIndex(I)- Index into tgSsf for avail to be processed
'       ilChfAdfCode(I)- advertiser code
'       ilSchMode(I)- 0=Insert; 1=Move; 2=Compact; 3=Preempt; 4=Preempt Fill only (call only from SpotMG)
'       ilBkQH(I)- Rank for ilSchMode = 3
'       ilCheckAvail(I)- True; False (bypass checking avail because this is split network)
'       ilRet(O)- True = No conflicts; False= Conflicts
'
'       tmSdf contains the spot record to be checked
'       imDayIndex contains the day to be checked
'       tgSsf contains the days events
'       tmAvail contain the avail to be check
'
    Dim ilIndex As Integer
    Dim ilSpotIndex As Integer
    Dim slTime As String
    Dim llStartAvailTime As Long
    Dim llEndAvailTime As Long
    Dim llAvailTime As Long
    Dim ilBypass As Integer
    Dim ilBypassIndex As Integer
    Dim ilInitPreempt As Integer
    Dim ilMatchComp As Integer  'True=Do Advt Test; False=Ignore Advt Test (treat as diff advt)
    Dim ilPass As Integer
    Dim ilBNoPasses As Integer
    Dim llPass2StartAvailTime As Long
    Dim ilANoPasses As Integer
    Dim llPass2EndAvailTime As Long
    Dim llDate As Long
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilRet As Integer
    Dim ilLBSpotMove As Integer
    
    If llSepLength <= 0 Then
        gAdvtTest = True
        Exit Function
    End If
    ilLBSpotMove = LBound(tlSpotMove)
    ilInitPreempt = UBound(tlSpotMove)
    'Test within current avail
    If ilAvailIndex + tlSsf.tPas(ADJSSFPASBZ + ilAvailIndex).iNoSpotsThis > UBound(tlSsf.tPas) Then
        gAdvtTest = False
        Exit Function
    End If
   LSet tmAvail = tlSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
    '12/7/09:  Remove the CheckAvail test as it was not coded correctly.  It assumed that all split networks were not overlapping
    '          i.e. Split 1:  West Coast; Split 2: East coast.  The two could be scheduled within the same avail
    '          It fails if Split 1 was West + East or Split 1 was exclude West.
    'If ilCheckAvail Then
        For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tmAvail.iNoSpotsThis Step 1
            If (ilSpotIndex < 1) Or (ilSpotIndex > tlSsf.iCount) Then
                gAdvtTest = False
                Exit Function
            End If
            LSet tmSpotTest = tlSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
            ilBypass = False
            ilMatchComp = False
            If (ilMnfComp0 = 0) And (ilMnfComp1 = 0) And (tmSpotTest.iMnfComp(0) = 0) And (tmSpotTest.iMnfComp(1) = 0) Then
                ilMatchComp = True
            Else
                If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp0 = tmSpotTest.iMnfComp(1))) Then
                    ilMatchComp = True
                End If
                If (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp1 = tmSpotTest.iMnfComp(1))) Then
                    ilMatchComp = True
                End If
            End If
            For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                    ilBypass = True
                    Exit For
                End If
            Next ilBypassIndex
            If Not ilBypass Then
                If (tmSpotTest.iAdfCode = ilChfAdfCode) And (ilMatchComp) Then
                    'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvail.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                    If gPreemptible(ilSchMode, tlSpotMove(), tmAvail, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                        If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                            ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                            tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                            tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                            tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                        Else
                            ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                            tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                            tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                            tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                        End If
                    Else
                        If ilInitPreempt < UBound(tlSpotMove) Then
                            ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                        End If
                        gAdvtTest = False
                        Exit Function
                    End If
                End If
            End If
        Next ilSpotIndex
    'End If
    'Determine if time or break advertiser test
    If tgVpf(ilVpfIndex).sAdvtSep = "B" Then
        gAdvtTest = True
        Exit Function
    End If
    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
    llAvailTime = CLng(gTimeToCurrency(slTime, False))
    llStartAvailTime = llAvailTime - llSepLength
    ilBNoPasses = 1
    If llStartAvailTime < 0 Then
        If tlSsf.iType = 0 Then
            ilBNoPasses = 2
            llPass2StartAvailTime = 86400 + llStartAvailTime
        End If
        llStartAvailTime = 0
    End If
    llEndAvailTime = llAvailTime + llSepLength
    ilANoPasses = 1
    If llEndAvailTime > 86400 Then
        If tlSsf.iType = 0 Then
            ilANoPasses = 2
            llPass2EndAvailTime = llEndAvailTime - 86400
        End If
        llEndAvailTime = 86400
    End If
    'Test avails prior to avail being considered
    ilIndex = ilAvailIndex - 1
    tmCTSsf = tlSsf
    For ilPass = 1 To ilBNoPasses Step 1
        Do
            If ilIndex < 1 Then
                Exit Do
            End If
            tmAvailTest = tmCTSsf.tPas(ADJSSFPASBZ + ilIndex)
            If (tmAvailTest.iRecType >= 2) And (tmAvailTest.iRecType <= 9) Then
                gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                llAvailTime = CLng(gTimeToCurrency(slTime, False))
                If llAvailTime < llStartAvailTime Then
                    Exit Do
                End If
                For ilSpotIndex = ilIndex + 1 To ilIndex + tmAvailTest.iNoSpotsThis Step 1
                    LSet tmSpotTest = tmCTSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                    ilBypass = False
                    ilMatchComp = False
                    If (ilMnfComp0 = 0) And (ilMnfComp1 = 0) And (tmSpotTest.iMnfComp(0) = 0) And (tmSpotTest.iMnfComp(1) = 0) Then
                        ilMatchComp = True
                    Else
                        If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp0 = tmSpotTest.iMnfComp(1))) Then
                            ilMatchComp = True
                        End If
                        If (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp1 = tmSpotTest.iMnfComp(1))) Then
                            ilMatchComp = True
                        End If
                    End If
                    For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                        If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                            ilBypass = True
                            Exit For
                        End If
                    Next ilBypassIndex
                    If Not ilBypass Then
                        If (tmSpotTest.iAdfCode = ilChfAdfCode) And (ilMatchComp) Then
                            'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvailTest.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                            If ilPass = 2 Then
                                If ilInitPreempt < UBound(tlSpotMove) Then
                                    ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                End If
                                gAdvtTest = False
                                Exit Function
                            End If
                            If gPreemptible(ilSchMode, tlSpotMove(), tmAvailTest, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                                If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                                    ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                                    tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                                    tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                                    tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                                Else
                                    ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                                    tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                                    tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                                    tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                                End If
                            Else
                                If ilInitPreempt < UBound(tlSpotMove) Then
                                    ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                End If
                                gAdvtTest = False
                                Exit Function
                            End If
                        End If
                    End If
                Next ilSpotIndex
            ElseIf tmAvailTest.iRecType = 1 Then
               LSet tmProg = tmCTSsf.tPas(ADJSSFPASBZ + ilIndex)
                gUnpackTime tmProg.iStartTime(0), tmProg.iStartTime(1), "A", "1", slTime
                llAvailTime = CLng(gTimeToCurrency(slTime, False))
            Else
                llAvailTime = llStartAvailTime + 1
            End If
            ilIndex = ilIndex - 1
        Loop While llAvailTime > llStartAvailTime
        If ilPass = ilBNoPasses Then
            Exit For
        End If
        gUnpackDateLong tlSsf.iDate(0), tlSsf.iDate(1), llDate
        gPackDateLong llDate - 1, ilDate0, ilDate1
        imSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
        tmSsfSrchKey.iType = 0 'slType
        tmSsfSrchKey.iVefCode = tlSsf.iVefCode
        tmSsfSrchKey.iDate(0) = ilDate0
        tmSsfSrchKey.iDate(1) = ilDate1
        tmSsfSrchKey.iStartTime(0) = 0
        tmSsfSrchKey.iStartTime(1) = 0
        ilRet = gSSFGetGreaterOrEqual(hlSsf, tmCTSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If (ilRet <> BTRV_ERR_NONE) Or (tmCTSsf.iType <> 0) Or (tmCTSsf.iVefCode <> tlSsf.iVefCode) Or (tmCTSsf.iDate(0) <> ilDate0) Or (tmCTSsf.iDate(1) <> ilDate1) Then
            Exit For
        End If
        ilIndex = tmCTSsf.iCount
        llStartAvailTime = llPass2StartAvailTime - 1
    Next ilPass
    'Test avails after avail being considered
    ilIndex = ilAvailIndex + tmAvail.iNoSpotsThis + 1
    tmCTSsf = tlSsf
    For ilPass = 1 To ilANoPasses Step 1
        Do
            If ilIndex > tmCTSsf.iCount Then
                Exit Do
            End If
            tmAvailTest = tmCTSsf.tPas(ADJSSFPASBZ + ilIndex)
            If (tmAvailTest.iRecType >= 2) And (tmAvailTest.iRecType <= 9) Then
                gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                llAvailTime = CLng(gTimeToCurrency(slTime, False))
                If llAvailTime > llEndAvailTime Then
                    Exit Do
                End If
                For ilSpotIndex = ilIndex + 1 To ilIndex + tmAvailTest.iNoSpotsThis Step 1
                    LSet tmSpotTest = tmCTSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                    ilBypass = False
                    ilMatchComp = False
                    If (ilMnfComp0 = 0) And (ilMnfComp1 = 0) And (tmSpotTest.iMnfComp(0) = 0) And (tmSpotTest.iMnfComp(1) = 0) Then
                        ilMatchComp = True
                    Else
                        If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp0 = tmSpotTest.iMnfComp(1))) Then
                            ilMatchComp = True
                        End If
                        If (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp1 = tmSpotTest.iMnfComp(1))) Then
                            ilMatchComp = True
                        End If
                    End If
                    For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                        If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                            ilBypass = True
                            Exit For
                        End If
                    Next ilBypassIndex
                    If Not ilBypass Then
                        If (tmSpotTest.iAdfCode = ilChfAdfCode) And (ilMatchComp) Then
                            'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvailTest.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                            If ilPass = 2 Then
                                If ilInitPreempt < UBound(tlSpotMove) Then
                                    ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                End If
                                gAdvtTest = False
                                Exit Function
                            End If
                            If gPreemptible(ilSchMode, tlSpotMove(), tmAvailTest, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                                If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                                    ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                                    tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                                    tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                                    tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                                Else
                                    ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                                    tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                                    tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                                    tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                                End If
                            Else
                                If ilInitPreempt < UBound(tlSpotMove) Then
                                    ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                End If
                                gAdvtTest = False
                                Exit Function
                            End If
                        End If
                    End If
                Next ilSpotIndex
                ilIndex = ilIndex + tmAvailTest.iNoSpotsThis + 1
            ElseIf tmAvailTest.iRecType = 1 Then
               LSet tmProg = tmCTSsf.tPas(ADJSSFPASBZ + ilIndex)
                gUnpackTime tmProg.iStartTime(0), tmProg.iStartTime(1), "A", "1", slTime
                llAvailTime = CLng(gTimeToCurrency(slTime, False))
                ilIndex = ilIndex + 1
            Else
                llAvailTime = llEndAvailTime - 1
                ilIndex = ilIndex + 1
            End If
        Loop While llAvailTime < llEndAvailTime
        If ilPass = ilANoPasses Then
            Exit For
        End If
        gUnpackDateLong tlSsf.iDate(0), tlSsf.iDate(1), llDate
        gPackDateLong llDate + 1, ilDate0, ilDate1
        imSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
        tmSsfSrchKey.iType = 0 'slType
        tmSsfSrchKey.iVefCode = tlSsf.iVefCode
        tmSsfSrchKey.iDate(0) = ilDate0
        tmSsfSrchKey.iDate(1) = ilDate1
        tmSsfSrchKey.iStartTime(0) = 0
        tmSsfSrchKey.iStartTime(1) = 0
        ilRet = gSSFGetGreaterOrEqual(hlSsf, tmCTSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        If (ilRet <> BTRV_ERR_NONE) Or (tmCTSsf.iType <> 0) Or (tmCTSsf.iVefCode <> tlSsf.iVefCode) Or (tmCTSsf.iDate(0) <> ilDate0) Or (tmCTSsf.iDate(1) <> ilDate1) Then
            Exit For
        End If
        ilIndex = 1
        llEndAvailTime = llPass2EndAvailTime
    Next ilPass
    gAdvtTest = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gBookSpot                       *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add spot to Ssf and change its  *
'*                     status                          *
'*                                                     *
'*******************************************************
Function gBookSpot(slInStatus As String, hlSdf As Integer, tlSdf As SDF, llSdfRecPos As Long, ilBkQH As Integer, hlSsf As Integer, tlSsf As SSF, llSsfRecPos As Long, ilAvailIndex As Integer, ilPos As Integer, tlChf As CHF, tlClf As CLF, tlRdf As RDF, ilVpfIndex As Integer, hlSmf As Integer, tlSmf As SMF, hlClf As Integer, hlCrf As Integer, ilPriceLevel As Integer, ilReplacePrimary As Integer, hlSxf As Integer, Optional hlGsf As Integer = 0, Optional slOverbookMode As String = "N") As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilValue                                                                               *
'******************************************************************************************

'
'   ilRet = gBookSpot(slStatus, hlSdf, tlSdf, llSdfRecPos, ilRank, hlSsf, tlSsf, llSsfRecPos, ilAvailIndex, ilPos, tlChf, tlClf, tlRdf, ilVpfIndex, hlSmf, tlSmf, hlClf, hlCrf, ilPriceLevel, ilReplacePri)
'   Where:
'       slStatus(I)- "S" scheduled as spot; "G" schedule as Makegood; "O" schedule as Outside contract limits
'       hlSdf(I)- Sdf open handle
'       tlSdf(I)- Sdf for spot to be added to Ssf
'       llSdfRecPos(I)- Sdf record position
'       ilBkQH(I)- Booking number of quarter hours
'       hlSsf(I)- Ssf open handle
'       tlSsf(I)- Ssf record image to add event into
'       llSsfRecPos(I)- Ssf record position
'       ilAvailIndex(I)- Avail index inwhich to add spot event
'       ilPos(I)- position within avail to insert spot event
'                 0=Use vehicle option rules (not coded)
'                 -1=At end
'                 1-N= position number
'       tlChf(I)- Contract header record
'       tlClf(I)- Contract line record
'       tlRdf(I)- Daypart record
'       imVpfIndex(I)- Vehicle option index
'       hlSmf(I)- Spot MG specification open handle
'       tlSmf(I/O)- old SMF record if existed (tlSmf.lChfCode = 0 => didn't exist)
'                   new SMF if created
'       hlCrf(I)- Copy rotation handle (undex to determine if copy should be retianed)
'       ilPriceLevel(I)-  zero or price level from eithet site ot vehicle option (used to preempt spots if rank match)
'       ilReplacePri(I)- Change primary that is pushed down to secondary
'
'       ilRet = True if spot booked; False if spot not booked
'
'       The calling program should have done a btrBeginTrans(hlFile, 1000) prior to
'       this call if booking more then one spot (Spot Screen)
'

    Dim ilLoop As Integer
    Dim ilFreeIndex As Integer
    Dim ilSHour As Integer
    Dim ilSMin As Integer
    Dim ilHour As Integer
    Dim ilMin As Integer
    Dim ilAvailHour As Integer
    Dim ilAvailMin As Integer
    Dim slDate As String
    Dim ilDay As Integer
    Dim slEntryDate As String
    Dim ilRet As Integer
    Dim slMissedDate As String
    Dim slMissedTime As String
    Dim slSchDate As String
    Dim slSchTime As String
    Dim ilOrigVefCode As Integer
    Dim ilOrigGameNo As Integer
    Dim ilSchGameNo As Integer
    Dim ilLen As Integer
    Dim slStatus As String
    'Copy rotation record information
    Dim llSAsgnDate As Long
    Dim llEAsgnDate As Long
    Dim llSAsgnTime As Long
    Dim llEAsgnTime As Long
    Dim llTime As Long
    Dim slType As String
    Dim ilCrfVefCode As Integer
    Dim tlCrfSrchKey1 As CRFKEY1 'CRF key record image
    Dim ilCrfRecLen As Integer  'CRF record length
    Dim tlCrf As CRF            'CRF record image
    Dim ilPkgVefCode As Integer
    Dim ilLnVefCode As Integer
    Dim ilSchPkgVefCode As Integer
    Dim slLive As String
    Dim ilRdfCode As Integer
    Dim ilOrigCrfVefCode As Integer
    Dim ilFound As Integer
    Dim ilBypassCrf As Integer
    Dim slStr As String
    Dim tlSpot As CSPOTSS

    'Dim tlCClf As CLF
    imSdfRecLen = Len(tlSdf)
    imSxfRecLen = Len(tmSxf)
    slStatus = slInStatus
    '4/19/12: Check the spot schedule status.
    '         If not missed, then reject the request to book the spot as it is already booked
    If llSdfRecPos > 0 Then
        ilRet = btrGetDirect(hlSdf, tlSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
    Else
        tmSdfSrchKey3.lCode = tlSdf.lCode
        ilRet = btrGetEqual(hlSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
    End If
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBookSpot-GetDirect Sdf(4)"
        gBookSpot = False
        Exit Function
    End If
    If (tlSdf.sSchStatus = "S") Or (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O") Then
        igBtrError = -5
        sgErrLoc = "gBookSpot-Spot Previously Booked"
        gBookSpot = False
        Exit Function
    End If
    If slStatus = "S" Then
        '4/19/12: Call moved above
        'If llSdfRecPos > 0 Then
        '    ilRet = btrGetDirect(hlSdf, tlSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        'Else
        '    tmSdfSrchKey3.lCode = tlSdf.lCode
        '    ilRet = btrGetEqual(hlSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
        'End If
        'If ilRet <> BTRV_ERR_NONE Then
        '    igBtrError = gConvertErrorCode(ilRet)
        '    sgErrLoc = "gBookSpot-GetDirect Sdf(4)"
        '    gBookSpot = False
        '    Exit Function
        'End If
        If tlSdf.sTracer = "*" Then
            If tlSdf.sSpotType = "X" Then
                slStatus = "O"
            Else
                slStatus = "G"
            End If
        End If
    End If
    Do
        imSsfRecLen = Len(tlSsf)
        ilRet = gSSFGetDirect(hlSsf, tlSsf, imSsfRecLen, llSsfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = gConvertErrorCode(ilRet)
            sgErrLoc = "gBookSpot-GetDirect SSf(1)"
            gBookSpot = False
            Exit Function
        End If
        ilSchGameNo = tlSsf.iType
        ilRet = gGetByKeyForUpdateSSF(hlSsf, tlSsf)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = gConvertErrorCode(ilRet)
            sgErrLoc = "gBookSpot-GetByKey SSf(2)"
            gBookSpot = False
            Exit Function
        End If
        If tlSsf.iCount >= UBound(tlSsf.tPas) Then
            igBtrError = -4
            sgErrLoc = "gBookSpot-Out of room in SSf"
            gBookSpot = False
            Exit Function
        End If
       LSet tmAvail = tlSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
        'Test that this is an avail
        If (tmAvail.iRecType < 2) Or (tmAvail.iRecType > 9) Then 'Contract Avail subrecord
            igBtrError = -1
            sgErrLoc = "gBookSpot-Not Avail Type in SSf"
            gBookSpot = False
            Exit Function
        End If
        'Test Room
        If tlClf.lRafCode > 0 Then
            If ilPos > 0 Then
                If tmAvail.iNoSpotsThis = 0 Then
                    ilLen = tlClf.iLen
                ElseIf ilReplacePrimary Then
                    ilLen = 0
                Else
                    ilLen = 0
                End If
            Else
                ilLen = tlClf.iLen
            End If
        Else
            ilLen = tlClf.iLen
        End If
        For ilLoop = ilAvailIndex + 1 To ilAvailIndex + tmAvail.iNoSpotsThis Step 1
           LSet tmSpot = tlSsf.tPas(ADJSSFPASBZ + ilLoop)
            If (tmSpot.iRecType And &HF) < 10 Then
                igBtrError = -2
                sgErrLoc = "gBookSpot-Avail Spot Count Error in SSf"
                gBookSpot = False
                Exit Function
            End If
            If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                ilLen = ilLen + (tmSpot.iPosLen And &HFFF)
            End If
        Next ilLoop
        If tmAvail.iLen < ilLen Then
            If slOverbookMode <> "Y" Then
                igBtrError = -3
                sgErrLoc = "gBookSpot-Avail Length error in SSf"
                gBookSpot = False
                Exit Function
            Else
                tmAvail.iLen = ilLen
                If tmAvail.iNoSpotsThis + 1 > (tmAvail.iAvInfo And &H1F) Then
                    tmAvail.iAvInfo = ((tmAvail.iAvInfo And (Not &H1F)) + tmAvail.iNoSpotsThis + 1) Or sSOverBook
                End If
            End If
        End If
        If ilPos = 0 Then   'Use vehicle option spot sorting specifications
            'Code later- insert at end for now
            If tlSdf.sSpotType <> "X" Then
                If (tlClf.iPosition = 1) And ((Asc(tlClf.sOV2DefinedBits) And LN1STPOSITION) = LN1STPOSITION) Then
                    ilFreeIndex = ilAvailIndex + 1
                Else
                    ilFreeIndex = ilAvailIndex + tmAvail.iNoSpotsThis + 1
                End If
            Else
                ilFreeIndex = ilAvailIndex + tmAvail.iNoSpotsThis + 1
            End If
        ElseIf (ilPos = -1) Or (ilPos > tmAvail.iNoSpotsThis) Then
            ilFreeIndex = ilAvailIndex + tmAvail.iNoSpotsThis + 1
        Else
            ilFreeIndex = ilAvailIndex + ilPos
        End If
        'Move all subrecords down one to make room for spot
        For ilLoop = tlSsf.iCount To ilFreeIndex Step -1
            If (ilLoop = ilFreeIndex) And (tlClf.lRafCode > 0) And (ilReplacePrimary) Then
               LSet tlSpot = tlSsf.tPas(ADJSSFPASBZ + ilLoop)
                tlSpot.iRecType = (tlSpot.iRecType And (Not SSSPLITPRI)) Or SSSPLITSEC
                LSet tlSsf.tPas(ADJSSFPASBZ + ilLoop + 1) = tlSpot
            Else
                tlSsf.tPas(ADJSSFPASBZ + ilLoop + 1) = tlSsf.tPas(ADJSSFPASBZ + ilLoop)
            End If
        Next ilLoop
        tlSsf.iCount = tlSsf.iCount + 1
        tmAvail.iNoSpotsThis = tmAvail.iNoSpotsThis + 1
        tlSsf.tPas(ADJSSFPASBZ + ilAvailIndex) = tmAvail
        tmSpot.iRecType = 10
        If (tlClf.sBB = "O") Or (tlClf.sBB = "B") Then
            tmSpot.iRecType = tmSpot.iRecType Or SSOPENBB
        End If
        If (tlClf.sBB = "C") Or (tlClf.sBB = "B") Then
            tmSpot.iRecType = tmSpot.iRecType Or SSCLOSEBB
        End If
        If tlClf.sExtra = "B" Then
            tmSpot.iRecType = tmSpot.iRecType Or SSBOOKEND
        End If
        If tlClf.sExtra = "D" Then
            tmSpot.iRecType = tmSpot.iRecType Or SSDONUT
        End If
        If tlRdf.sInOut = "I" Then
            tmSpot.iRecType = tmSpot.iRecType Or SSAVAILBUY
        End If
        If tlClf.sPreempt = "P" Then
            tmSpot.iRecType = tmSpot.iRecType Or SSPREEMPTIBLE
        End If
        If tlRdf.sInOut = "O" Then
            tmSpot.iRecType = tmSpot.iRecType Or SSEXAVAILBUY
        End If
        If (tlChf.iMnfExcl(0) <> 0) Or (tlChf.iMnfExcl(1) <> 0) Then
            tmSpot.iRecType = tmSpot.iRecType Or SSEXCLUSIONS
        End If
        If tlClf.lRafCode > 0 Then
            If ilPos > 0 Then
                If (ilReplacePrimary) Or (tmAvail.iNoSpotsThis = 1) Then
                    tmSpot.iRecType = tmSpot.iRecType Or SSSPLITPRI
                Else
                    tmSpot.iRecType = tmSpot.iRecType Or SSSPLITSEC
                End If
            Else
                tmSpot.iRecType = tmSpot.iRecType Or SSSPLITPRI
            End If
        End If
        tmSpot.lSdfCode = tlSdf.lCode 'lSdfRecNo = llSdfRecPos
        gUnpackDate tlSsf.iDate(0), tlSsf.iDate(1), slDate
        ilDay = gWeekDayStr(slDate)
        ilAvailHour = tmAvail.iTime(1) \ 256
        ilAvailMin = tmAvail.iTime(1) And &HFF
        'gUnpackDate tlClf.iEntryDate(0), tlClf.iEntryDate(1), slEntryDate
        slEntryDate = Format(gNow(), "m/d/yy")
        tmSpot.lBkInfo = gDateValue(slEntryDate)
        If tlSdf.sSpotType <> "X" Then
            If tlClf.sSoloAvail = "Y" Then
                tmSpot.lBkInfo = tmSpot.lBkInfo Or SSSOLOAVAIL
            End If
            If (tlClf.iPosition = 1) And ((Asc(tlClf.sOV2DefinedBits) And LN1STPOSITION) = LN1STPOSITION) Then
                tmSpot.lBkInfo = tmSpot.lBkInfo Or SS1STPOSITION
            End If
        End If
        If (tlRdf.iLtfCode(0) <> 0) Or (tlRdf.iLtfCode(1) <> 0) Or (tlRdf.iLtfCode(2) <> 0) Then
            tmSpot.iRecType = tmSpot.iRecType Or SSLIBBUY
        Else    'Time buy- check if override times defined (if so, use them as bump times)
            gXMidClfRdfToRdf "", tlClf, tlRdf, tmXMidClf, tmXMidRdf
            If ((tmXMidClf.iStartTime(0) = 1) And (tmXMidClf.iStartTime(1) = 0)) Or (tgVpf(ilVpfIndex).sGMedium = "S") Then
                For ilLoop = LBound(tmXMidRdf.iStartTime, 2) To UBound(tmXMidRdf.iStartTime, 2) Step 1
                    If (tmXMidRdf.iStartTime(0, ilLoop) <> 1) Or (tmXMidRdf.iStartTime(1, ilLoop) <> 0) Then
                        'If tmXMidRdf.sWkDays(ilLoop, ilDay + 1) = "Y" Then
                        If tmXMidRdf.sWkDays(ilLoop, ilDay) = "Y" Then
                            ilSHour = tmXMidRdf.iStartTime(1, ilLoop) \ 256
                            ilSMin = tmXMidRdf.iStartTime(1, ilLoop) And &HFF
                            If ilSHour * 60 + ilSMin <= ilAvailHour * 60 + ilAvailMin Then
                                ilHour = tmXMidRdf.iEndTime(1, ilLoop) \ 256
                                ilMin = tmXMidRdf.iEndTime(0, ilLoop) And &HFF
                                If (ilHour = 0) And (ilMin = 0) Then
                                    ilHour = 23
                                    ilMin = 59
                                End If
                                If ilHour * 60 + ilMin <= ilAvailHour * 60 + ilAvailMin Then
                                    If ilSHour = ilAvailHour Then
                                        tmSpot.lBkInfo = tmSpot.lBkInfo Or (CLng(ilSMin) * SHIFT17)
                                    End If
                                    If ilHour = ilAvailHour Then
                                        tmSpot.lBkInfo = tmSpot.lBkInfo Or (CLng(ilMin) * SHIFT23)
                                    Else
                                        tmSpot.lBkInfo = tmSpot.lBkInfo Or (60& * SHIFT23)
                                    End If
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next ilLoop
            Else
                ilHour = tmXMidClf.iStartTime(1) \ 256
                ilMin = tmXMidClf.iStartTime(1) And &HFF
                If ilHour = ilAvailHour Then
                    tmSpot.lBkInfo = tmSpot.lBkInfo Or (CLng(ilMin) * SHIFT17)
                End If
                ilHour = tmXMidClf.iEndTime(1) \ 256
                ilMin = tmXMidClf.iEndTime(1) And &HFF
                If (ilHour = 0) And (ilMin = 0) Then
                    ilHour = 23
                    ilMin = 59
                End If
                If ilHour = ilAvailHour Then
                    tmSpot.lBkInfo = tmSpot.lBkInfo Or (CLng(ilMin) * SHIFT23)
                Else
                    tmSpot.lBkInfo = tmSpot.lBkInfo Or (60& * SHIFT23)
                End If
            End If
        End If
        tmSpot.iMnfComp(0) = tlChf.iMnfComp(0)
        tmSpot.iMnfComp(1) = tlChf.iMnfComp(1)
        tmSpot.iPosLen = tlClf.iLen + (tlClf.iPosition * SHIFT12)
        tmSpot.iAdfCode = tlChf.iAdfCode
        tmSpot.iRank = (ilPriceLevel * SHIFT11) + ilBkQH '(tmSpot.iRank And PRICELEVELMASK) + ilBkQH
        LSet tlSsf.tPas(ADJSSFPASBZ + ilFreeIndex) = tmSpot
        imSsfRecLen = igSSFBaseLen + tlSsf.iCount * Len(tmAvail)
        ilRet = gSSFUpdate(hlSsf, tlSsf, imSsfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBookSpot-Update SSf(3)"
        gBookSpot = False
        Exit Function
    End If
    'Update Sdf record
    If slStatus <> "S" Then
        gUnpackDate tlSdf.iDate(0), tlSdf.iDate(1), slMissedDate
        gUnpackDate tlSsf.iDate(0), tlSsf.iDate(1), slSchDate
        gUnpackTime tlSdf.iTime(0), tlSdf.iTime(1), "A", "1", slMissedTime
        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slSchTime
        ilOrigGameNo = tlSdf.iGameNo
    End If
    '4/19/12: Set to avoid reading in sdf a second time if not required
    ilRet = BTRV_ERR_NONE
    Do
        
        ''ilRet = btrGetDirect(hlSdf, tlSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        '4/19/12: Only required if conflict as code added at top to access the spot
        If ilRet = BTRV_ERR_CONFLICT Then
            If llSdfRecPos > 0 Then
                ilRet = btrGetDirect(hlSdf, tlSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            Else
                tmSdfSrchKey3.lCode = tlSdf.lCode
                ilRet = btrGetEqual(hlSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
            End If
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = gConvertErrorCode(ilRet)
                sgErrLoc = "gBookSpot-GetDirect Sdf(4)"
                gBookSpot = False
                Exit Function
            End If
        End If
        'tmSRec = tlSdf
        'ilRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
        'tlSdf = tmSRec
        'If ilRet <> BTRV_ERR_NONE Then
        '    igBtrError = gConvertErrorCode(ilRet)
        '    sgErrLoc = "gBookSpot-GetByKey Sdf(5)"
        '    gBookSpot = False
        '    Exit Function
        'End If
        If (tlSdf.sPtType <> "0") And (tlSdf.iRotNo > 0) Then
            'imClfRecLen = Len(tlCClf)
            ilSchPkgVefCode = 0
            ilCrfRecLen = Len(tlCrf)
            ilRet = gGetCrfVefCode(hlClf, tlSdf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
            If (slStatus = "G") Or (slStatus = "O") Then
                slStr = gGetMGCopyAssign(tlSdf, ilPkgVefCode, ilLnVefCode, slLive, hlSmf, hlCrf)
                 If (slStr = "S") Or (slStr = "B") Then
                    ilSchPkgVefCode = gGetMGPkgVefCode(hlClf, tlSdf)
                End If
                ilCrfVefCode = tlSsf.iVefCode
                ilOrigVefCode = tlSdf.iVefCode
                tlSdf.iVefCode = tlSsf.iVefCode 'Booked vehicle
                tlSdf.iDate(0) = tlSsf.iDate(0)
                tlSdf.iDate(1) = tlSsf.iDate(1)
                tlSdf.iTime(0) = tmAvail.iTime(0)
                tlSdf.iTime(1) = tmAvail.iTime(1)
                tlSdf.iMnfMissed = 0    'Field used to retain missed reason and rotation #
                If slStr = "O" Then
                    ilCrfVefCode = ilLnVefCode
                    ilLnVefCode = 0
                ElseIf slStr = "S" Then
                    ilPkgVefCode = ilSchPkgVefCode
                    ilSchPkgVefCode = 0
                    ilLnVefCode = 0
                Else
                    If ilPkgVefCode = ilSchPkgVefCode Then
                        ilSchPkgVefCode = 0
                    End If
                End If
            Else
                ilLnVefCode = 0
                ilCrfVefCode = tlSsf.iVefCode
                ilOrigVefCode = tlSdf.iVefCode
                tlSdf.iVefCode = tlSsf.iVefCode 'Booked vehicle
                tlSdf.iDate(0) = tlSsf.iDate(0)
                tlSdf.iDate(1) = tlSsf.iDate(1)
                tlSdf.iTime(0) = tmAvail.iTime(0)
                tlSdf.iTime(1) = tmAvail.iTime(1)
                tlSdf.iMnfMissed = 0    'Field used to retain missed reason and rotation #
            End If
            ilOrigCrfVefCode = ilCrfVefCode
            ilFound = False
            Do
                slType = "A"
                tlCrfSrchKey1.sRotType = slType
                tlCrfSrchKey1.iEtfCode = 0
                tlCrfSrchKey1.iEnfCode = 0
                tlCrfSrchKey1.iAdfCode = tlSdf.iAdfCode
                tlCrfSrchKey1.lChfCode = tlSdf.lChfCode
                tlCrfSrchKey1.iVefCode = ilCrfVefCode
                tlCrfSrchKey1.iRotNo = tlSdf.iRotNo
                ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                If (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (tlCrf.iVefCode = ilCrfVefCode) And (tlCrf.iRotNo = tlSdf.iRotNo) Then    'ilVefCode)
                    'Test if times and/or days ok
                    ilBypassCrf = False
                    'Test if looking for Live or Recorded rotations
                    If tlCrf.sState <> "D" Then
                        If slLive = "L" Then
                            If tlCrf.sLiveCopy <> "L" Then
                                ilBypassCrf = True
                            End If
                        ElseIf slLive = "M" Then
                            If tlCrf.sLiveCopy <> "M" Then
                                ilBypassCrf = True
                            End If
                        ElseIf slLive = "S" Then
                            If tlCrf.sLiveCopy <> "S" Then
                                ilBypassCrf = True
                            End If
                        ElseIf slLive = "P" Then
                            If tlCrf.sLiveCopy <> "P" Then
                                ilBypassCrf = True
                            End If
                        ElseIf slLive = "Q" Then
                            If tlCrf.sLiveCopy <> "Q" Then
                                ilBypassCrf = True
                            End If
                        Else
                            If (tlCrf.sLiveCopy = "L") Or (tlCrf.sLiveCopy = "M") Or (tlCrf.sLiveCopy = "S") Or (tlCrf.sLiveCopy = "P") Or (tlCrf.sLiveCopy = "Q") Then
                                ilBypassCrf = True
                            End If
                        End If
                    Else
                        ilBypassCrf = True
                    End If
                    If (tlCrf.sDay(ilDay) = "Y") And (Not ilBypassCrf) Then
                        gUnpackDateLong tlCrf.iStartDate(0), tlCrf.iStartDate(1), llSAsgnDate
                        gUnpackDateLong tlCrf.iEndDate(0), tlCrf.iEndDate(1), llEAsgnDate
                        If (gDateValue(slDate) >= llSAsgnDate) And (gDateValue(slDate) <= llEAsgnDate) Then
                            gUnpackTimeLong tlCrf.iStartTime(0), tlCrf.iStartTime(1), False, llSAsgnTime
                            gUnpackTimeLong tlCrf.iEndTime(0), tlCrf.iEndTime(1), True, llEAsgnTime
                            gUnpackTimeLong tlSdf.iTime(0), tlSdf.iTime(1), False, llTime
                            If (llTime >= llSAsgnTime) And (llTime <= llEAsgnTime) Then
                                'Check avail type
                                If (tlCrf.sInOut = "I") Or (tlCrf.sInOut = "O") Then
                                    If tlCrf.sInOut = "I" Then
                                        If tlCrf.ianfCode = tmAvail.ianfCode Then
                                            ilFound = True
                                        End If
                                    Else
                                        If tlCrf.ianfCode <> tmAvail.ianfCode Then
                                            ilFound = True
                                        End If
                                    End If
                                Else
                                    ilFound = True
                                End If
                            End If
                        End If
                    End If
                End If
                If ilFound Then
                    Exit Do
                End If
                If ilPkgVefCode > 0 Then
                    ilCrfVefCode = ilPkgVefCode
                    ilPkgVefCode = 0
                ElseIf ilSchPkgVefCode > 0 Then
                    ilCrfVefCode = ilSchPkgVefCode
                    ilSchPkgVefCode = 0
                Else
                    If (ilOrigCrfVefCode = ilLnVefCode) Or (ilLnVefCode = 0) Then
                        Exit Do
                    End If
                    ilCrfVefCode = ilLnVefCode
                    ilLnVefCode = 0
                End If
            Loop While ilCrfVefCode > 0
            If Not ilFound Then
                tlSdf.sPtType = "0"
                tlSdf.lCopyCode = 0
                tlSdf.iRotNo = 0
            End If
        Else
            ilOrigVefCode = tlSdf.iVefCode
            tlSdf.iVefCode = tlSsf.iVefCode 'Booked vehicle
            tlSdf.iDate(0) = tlSsf.iDate(0)
            tlSdf.iDate(1) = tlSsf.iDate(1)
            tlSdf.iTime(0) = tmAvail.iTime(0)
            tlSdf.iTime(1) = tmAvail.iTime(1)
            tlSdf.iMnfMissed = 0    'Field used to retain missed reason and rotation #
            If tlSdf.lFsfCode <= 0 Then
                tlSdf.sPtType = "0"
                tlSdf.lCopyCode = 0
                tlSdf.iRotNo = 0
            End If
        End If
        tlSdf.sSchStatus = slStatus
        '10319 test regional copy
        mUnassignRegionalAsNeeded hlCrf, tlSdf.lCode, slDate, llTime, tmAvail.ianfCode, slLive, tlSdf.sSchStatus, tlSdf.iVefCode
        tlSdf.iGameNo = ilSchGameNo
        ilHour = tlSdf.iTime(1) \ 256
        tlSdf.sXCrossMidnight = "N"
        If (tmAvail.iAvInfo And SSXMID) = SSXMID Then
            If ilHour < 6 Then
                tlSdf.sXCrossMidnight = "Y"
            End If
        Else
            If ilHour < 12 Then
                ilHour = tmAvail.iTime(1) \ 256
                If ilHour >= 12 Then
                    tlSdf.sXCrossMidnight = "Y"
                End If
            End If
        End If
        ilRet = gSxfDelete(hlSxf, tlSdf)
        tlSdf.iUrfCode = tgUrf(0).iCode
        If slStatus = "S" Then
            tlSdf.lSmfCode = 0
        End If
        ilRet = btrUpdate(hlSdf, tlSdf, imSdfRecLen)
    Loop While ilRet = BTRV_ERR_CONFLICT
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBookSpot-Update Sdf(6)"
        gBookSpot = False
        Exit Function
    End If
    If slStatus <> "S" Then
        ilRet = gMakeSmf(hlSmf, tlSmf, slStatus, tlSdf, ilOrigVefCode, slMissedDate, slMissedTime, ilOrigGameNo, slSchDate, slSchTime)
        If Not ilRet Then
            gBookSpot = False
            Exit Function
        End If
        tlSdf.lSmfCode = tlSmf.lCode
        tlSdf.sTracer = ""
        tlSdf.iUrfCode = tgUrf(0).iCode
        ilRet = btrUpdate(hlSdf, tlSdf, imSdfRecLen)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = gConvertErrorCode(ilRet)
            sgErrLoc = "gBookSpot-Update Sdf(7)"
            gBookSpot = False
            Exit Function
        End If
    End If
    gMakeLogAlert tlSdf, "S", hlGsf
    gBookSpot = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gBuildEventDay                  *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build a day into LLC            *
'*                                                     *
'*******************************************************
Function gBuildEventDay(ilType As Integer, sLCP As String, ilVehCode As Integer, slDate As String, slSTime As String, slETime As String, ilEvtType() As Integer, tlLLC() As LLC) As Integer
'
'   ilRet = gBuildEventDay (slType, slCP, ilVehCode, slDate, slStartTime, slEndTime, ilEvtType, tlLLC)
'
'   Where:
'       slType (I)- "O"=On air; "A"=Alternate
'       slCP (I)- "C"=Current only; "P"=Pending only; "B"=Both
'       ilVehCode (I)-Vehicle code number
'       slDate (I)- Date that events are to be obtained (For TFN use: "TFNMO", "TFNTU", "TFNWE", "TFNTH", "TFNFR", "TFNSA", "TFNSU")
'       slStartTime (I)- Start Time (included)
'       slEndTime (I)- End time (not included)
'       ilEvtType (I)- Array of which events are to be included (True or False)
'                       Index description
'                         0   Library
'                         1   Program event
'                         2   Contract avail
'                         3   Open BB
'                         4   Floating BB
'                         5   Close BB
'                         6   Cmml promo
'                         7   Feed avail
'                         8   PSA avail
'                         9   Promo avail
'                         10  Page eject
'                         11  Line space 1
'                         12  Line space 2
'                         13  Line space 3
'                         14  Other event types
'
'       tlLLC (O)- Event records
'
'       ilRet = True or False (error)
'
    'LCF Variables
    Dim hlLcf As Integer            'Log calendar file handle
    Dim tlLcf As LCF                'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF Key 0 image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim tlDLcf As LCF               'LCF record image
    'Library title
    Dim hlLtf As Integer            'Log library name file handle
    Dim tlLtf As LTF                'LTF record image
    Dim tlLtfSrchKey As INTKEY0     'LTF key record image
    Dim ilLtfRecLen As Integer      'LTF record length
    'Library version
    Dim hlLvf As Integer            'Log library name file handle
    Dim tlLvf As LVF                'LVF record image
    Dim tlLvfSrchKey As LONGKEY0     'LVF key record image
    Dim ilLvfRecLen As Integer      'LVF record length
    'Events
    Dim hlLef As Integer            'Event file handle
    Dim tlLef As LEF                'LEF record image
    Dim tlLefSrchKey As LEFKEY0     'LEF key record image
    Dim ilLefRecLen As Integer      'LEF record length
    'Event names
    Dim hlEnf As Integer            'Event name file handle
    Dim tlEnf As ENF                'Enf record images
    Dim tlEnfSrchKey As INTKEY0     'Enf key record image
    Dim ilEnfRecLen As Integer         'Enf record length
    'Avail names
    Dim hlAnf As Integer            'Avail name file handle
    Dim tlAnf As ANF                'Anf record images
    Dim tlAnfSrchKey As INTKEY0     'Anf key record image
    Dim ilAnfRecLen As Integer      'Anf record length

    Dim ilDay As Integer
    Dim ilCUpper As Integer
    Dim ilPUpper As Integer
    Dim ilSeqNo As Integer
    Dim ilRet As Integer
    Dim ilIndex As Integer
    Dim slStr As String
    Dim ilDate0 As Integer          'Byte 0 start date
    Dim ilDate1 As Integer          'Byte 1 start date
    Dim clGStartTime As Currency        'general start time reference
    Dim clGEndTime As Currency          'general end time reference
    Dim clCEvtTime As Currency          'Current event start time
    Dim clCEvtEndTime As Currency       'Current event end time
    Dim clPEvtTime As Currency          'Pending event start time
    Dim clPEvtEndTime As Currency       'Pending event start time
    Dim ilCFindEvt As Integer           'Current event found flag
    Dim ilPFindEvt As Integer           'Pending event found flag
    Dim slTime As String                'Time string
    Dim ilShowCurrent As Integer        '0=Show current event; 1=show pending only;
                                        '-1=Increment pending as current and pending
                                        'intersect (pending will be shown once current is beyond pending)
    Dim ilUpper As Integer
    Dim ilFound As Integer
    Dim ilEtfCode As Integer
    Dim llLvfCode As Long
    Dim ilLtfCode As Integer
    Dim ilOnlyLib As Integer
    Dim ilLoop As Integer
    Dim ilDeleted As Integer    'True=Library deleted
    Dim ilDel As Integer
    Dim ilTestDel As Integer
    Dim slStartTime As String
    Dim slXMid As String
    ReDim tlCLLC(0 To 0) As LLC
    ReDim tlPLLC(0 To 0) As LLC

    ilOnlyLib = True
    For ilLoop = LBound(ilEvtType) + 1 To UBound(ilEvtType) Step 1
        If ilEvtType(ilLoop) Then
            ilOnlyLib = False
        End If
    Next ilLoop
    Select Case UCase(slDate)
        Case "TFNMO"
            ilDate0 = 1
            ilDate1 = 0
            ilDay = 0
        Case "TFNTU"
            ilDate0 = 2
            ilDate1 = 0
            ilDay = 1
        Case "TFNWE"
            ilDate0 = 3
            ilDate1 = 0
            ilDay = 2
        Case "TFNTH"
            ilDate0 = 4
            ilDate1 = 0
            ilDay = 3
        Case "TFNFR"
            ilDate0 = 5
            ilDate1 = 0
            ilDay = 4
        Case "TFNSA"
            ilDate0 = 6
            ilDate1 = 0
            ilDay = 5
        Case "TFNSU"
            ilDate0 = 7
            ilDate1 = 0
            ilDay = 6
        Case Else
            gPackDate slDate, ilDate0, ilDate1
            ilDay = gWeekDayStr(slDate)
    End Select
    ilLcfRecLen = Len(tlLcf)  'Get and save LCF record length
    hlLcf = CBtrvTable(ONEHANDLE)        'Create LCF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventDay-Open Lcf(1)"
        ilRet = btrClose(hlLcf)
        btrDestroy hlLcf
        gBuildEventDay = False
        Exit Function
    End If
    ilLtfRecLen = Len(tlLtf)  'Get and save LTF record length
    hlLtf = CBtrvTable(ONEHANDLE)        'Create LTF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlLtf, "", sgDBPath & "Ltf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventDay-Open Ltf(2)"
        ilRet = btrClose(hlLcf)
        ilRet = btrClose(hlLtf)
        btrDestroy hlLcf
        btrDestroy hlLtf
        gBuildEventDay = False
        Exit Function
    End If
    ilLvfRecLen = Len(tlLvf)  'Get and save LVF record length
    hlLvf = CBtrvTable(ONEHANDLE)        'Create LVF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventDay-Open Lvf(3)"
        ilRet = btrClose(hlLcf)
        ilRet = btrClose(hlLtf)
        ilRet = btrClose(hlLvf)
        btrDestroy hlLcf
        btrDestroy hlLtf
        btrDestroy hlLvf
        gBuildEventDay = False
        Exit Function
    End If
    ilLefRecLen = Len(tlLef)  'Get and save LEF record length
    hlLef = CBtrvTable(ONEHANDLE)        'Create LEF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlLef, "", sgDBPath & "Lef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventDay-Open Lef(4)"
        ilRet = btrClose(hlLef)
        ilRet = btrClose(hlLcf)
        ilRet = btrClose(hlLtf)
        ilRet = btrClose(hlLvf)
        btrDestroy hlLef
        btrDestroy hlLcf
        btrDestroy hlLtf
        btrDestroy hlLvf
        gBuildEventDay = False
        Exit Function
    End If
    ilEnfRecLen = Len(tlEnf)  'Get and save EnF record length
    hlEnf = CBtrvTable(ONEHANDLE)        'Create ENF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlEnf, "", sgDBPath & "Enf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventDay-Open Enf(5)"
        ilRet = btrClose(hlEnf)
        ilRet = btrClose(hlLef)
        ilRet = btrClose(hlLcf)
        ilRet = btrClose(hlLtf)
        ilRet = btrClose(hlLvf)
        btrDestroy hlEnf
        btrDestroy hlLef
        btrDestroy hlLcf
        btrDestroy hlLtf
        btrDestroy hlLvf
        gBuildEventDay = False
        Exit Function
    End If
    ilAnfRecLen = Len(tlAnf)  'Get and save AnF record length
    hlAnf = CBtrvTable(ONEHANDLE)        'Create ANF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventDay-Open Anf(6)"
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hlEnf)
        ilRet = btrClose(hlLef)
        ilRet = btrClose(hlLcf)
        ilRet = btrClose(hlLtf)
        ilRet = btrClose(hlLvf)
        btrDestroy hlAnf
        btrDestroy hlEnf
        btrDestroy hlLef
        btrDestroy hlLcf
        btrDestroy hlLtf
        btrDestroy hlLvf
        gBuildEventDay = False
        Exit Function
    End If
    ilTestDel = True
    If (sLCP = "P") Or (sLCP = "B") Then
        tlLcfSrchKey.iType = ilType
        tlLcfSrchKey.sStatus = "D"
        tlLcfSrchKey.iVefCode = ilVehCode
        tlLcfSrchKey.iLogDate(0) = ilDate0
        tlLcfSrchKey.iLogDate(1) = ilDate1
        tlLcfSrchKey.iSeqNo = 1
        ilRet = btrGetEqual(hlLcf, tlDLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
        If ilRet <> BTRV_ERR_NONE Then
            ilTestDel = False
        End If
    Else    'Current only- ignore any pending or deletions
        ilTestDel = False
    End If
    ilCUpper = UBound(tlCLLC)
    tlCLLC(ilCUpper).iDay = -1
    If (sLCP = "C") Or (sLCP = "B") Then
        ilSeqNo = 1
        Do
            tlLcfSrchKey.iType = ilType
            tlLcfSrchKey.sStatus = "C"
            tlLcfSrchKey.iVefCode = ilVehCode
            tlLcfSrchKey.iLogDate(0) = ilDate0
            tlLcfSrchKey.iLogDate(1) = ilDate1
            tlLcfSrchKey.iSeqNo = ilSeqNo
            ilRet = btrGetEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
            If ilRet = BTRV_ERR_NONE Then
                ilSeqNo = ilSeqNo + 1
                For ilIndex = LBound(tlLcf.lLvfCode) To UBound(tlLcf.lLvfCode) Step 1
                    If tlLcf.lLvfCode(ilIndex) <> 0 Then
                        'Test if deleted- if so ignore
                        ilDeleted = False
                        If ilTestDel Then
                            For ilDel = LBound(tlDLcf.lLvfCode) To UBound(tlDLcf.lLvfCode) Step 1
                                If tlDLcf.lLvfCode(ilDel) <> 0 Then
                                    If tlLcf.lLvfCode(ilIndex) = tlDLcf.lLvfCode(ilDel) Then
                                        If (tlLcf.iTime(0, ilIndex) = tlDLcf.iTime(0, ilDel)) And (tlLcf.iTime(1, ilIndex) = tlDLcf.iTime(1, ilDel)) Then
                                            ilDeleted = True
                                        End If
                                    End If
                                End If
                            Next ilDel
                        End If
                        If Not ilDeleted Then
                            tlCLLC(ilCUpper).iDay = ilDay
                            gUnpackTime tlLcf.iTime(0, ilIndex), tlLcf.iTime(1, ilIndex), "A", "1", tlCLLC(ilCUpper).sStartTime
                            'Read in Lvf to obtain name and length
                            tlLvfSrchKey.lCode = tlLcf.lLvfCode(ilIndex)
                            ilRet = btrGetEqual(hlLvf, tlLvf, ilLvfRecLen, tlLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                            If ilRet = BTRV_ERR_NONE Then
                                tlCLLC(ilCUpper).lLvfCode = tlLvf.lCode
                                'Read in Ltf
                                tlLtfSrchKey.iCode = tlLvf.iLtfCode
                                ilRet = btrGetEqual(hlLtf, tlLtf, ilLtfRecLen, tlLtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                                If ilRet = BTRV_ERR_NONE Then
                                    tlCLLC(ilCUpper).iLtfCode = tlLtf.iCode
                                    tlCLLC(ilCUpper).sType = tlLtf.sType
                                    gUnpackLength tlLvf.iLen(0), tlLvf.iLen(1), "3", False, tlCLLC(ilCUpper).sLength
                                    If tlLtf.iVar <> 0 Then
                                        tlCLLC(ilCUpper).sName = Trim$(str$(tlLvf.iVersion)) & "/" & Trim$(tlLtf.sName) & "-" & Trim$(str$(tlLtf.iVar))
                                    Else
                                        tlCLLC(ilCUpper).sName = Trim$(str$(tlLvf.iVersion)) & "/" & Trim$(tlLtf.sName)
                                    End If
    '                                tlCLLC(ilCUpper).sName = tlCLLC(ilCUpper).sName & "\" & Trim$(Str$(tlLcf.lLvfCode(ilIndex)))
                                    ilCUpper = ilCUpper + 1
                                    ReDim Preserve tlCLLC(0 To ilCUpper) As LLC
                                    tlCLLC(ilCUpper).iDay = -1
                                Else
                                    tlCLLC(ilCUpper).iDay = -1
                                End If
                            Else
                                tlCLLC(ilCUpper).iDay = -1
                            End If
                        End If
                    Else
                        ilSeqNo = -1
                        Exit For
                    End If
                Next ilIndex
            Else
                ilSeqNo = -1
            End If
        Loop While ilSeqNo > 0
    End If
    ilPUpper = UBound(tlPLLC)
    tlPLLC(ilPUpper).iDay = -1
    If (sLCP = "P") Or (sLCP = "B") Then
        ilSeqNo = 1
        Do
            tlLcfSrchKey.iType = ilType
            tlLcfSrchKey.sStatus = "P"
            tlLcfSrchKey.iVefCode = ilVehCode
            tlLcfSrchKey.iLogDate(0) = ilDate0
            tlLcfSrchKey.iLogDate(1) = ilDate1
            tlLcfSrchKey.iSeqNo = ilSeqNo
            ilRet = btrGetEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
            If ilRet = BTRV_ERR_NONE Then
                ilSeqNo = ilSeqNo + 1
                For ilIndex = LBound(tlLcf.lLvfCode) To UBound(tlLcf.lLvfCode) Step 1
                    If tlLcf.lLvfCode(ilIndex) <> 0 Then
                        ilDeleted = False
                        If ilTestDel Then
                            For ilDel = LBound(tlDLcf.lLvfCode) To UBound(tlDLcf.lLvfCode) Step 1
                                If tlDLcf.lLvfCode(ilDel) <> 0 Then
                                    If tlLcf.lLvfCode(ilIndex) = tlDLcf.lLvfCode(ilDel) Then
                                        If (tlLcf.iTime(0, ilIndex) = tlDLcf.iTime(0, ilDel)) And (tlLcf.iTime(1, ilIndex) = tlDLcf.iTime(1, ilDel)) Then
                                            ilDeleted = True
                                        End If
                                    End If
                                End If
                            Next ilDel
                        End If
                        If Not ilDeleted Then
                            tlPLLC(ilPUpper).iDay = ilDay
                            gUnpackTime tlLcf.iTime(0, ilIndex), tlLcf.iTime(1, ilIndex), "A", "1", tlPLLC(ilPUpper).sStartTime
                            'Read in Lvf to obtain name and length
                            tlLvfSrchKey.lCode = tlLcf.lLvfCode(ilIndex)
                            ilRet = btrGetEqual(hlLvf, tlLvf, ilLvfRecLen, tlLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                            If ilRet = BTRV_ERR_NONE Then
                                tlPLLC(ilPUpper).lLvfCode = tlLvf.lCode
                                tlLtfSrchKey.iCode = tlLvf.iLtfCode
                                ilRet = btrGetEqual(hlLtf, tlLtf, ilLtfRecLen, tlLtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                                If ilRet = BTRV_ERR_NONE Then
                                    tlPLLC(ilPUpper).iLtfCode = tlLtf.iCode
                                    tlPLLC(ilPUpper).sType = tlLtf.sType
                                    gUnpackLength tlLvf.iLen(0), tlLvf.iLen(1), "3", False, tlPLLC(ilPUpper).sLength
                                    If tlLtf.iVar <> 0 Then
                                        tlPLLC(ilPUpper).sName = Trim$(str$(tlLvf.iVersion)) & "/" & Trim$(tlLtf.sName) & "-" & Trim$(str$(tlLtf.iVar))
                                    Else
                                        tlPLLC(ilPUpper).sName = Trim$(str$(tlLvf.iVersion)) & "/" & Trim$(tlLtf.sName)
                                    End If
    '                                tlPLLC(ilPUpper).sName = tlPLLC(ilPUpper).sName & "\" & Trim$(Str$(tlLcf.lLvfCode(ilIndex)))
                                    ilPUpper = ilPUpper + 1
                                    ReDim Preserve tlPLLC(0 To ilPUpper) As LLC
                                    tlPLLC(ilPUpper).iDay = -1
                                Else
                                    tlPLLC(ilPUpper).iDay = -1
                                End If
                            Else
                                tlPLLC(ilPUpper).iDay = -1
                            End If
                        End If
                    Else
                        ilSeqNo = -1
                        Exit For
                    End If
                Next ilIndex
            Else
                ilSeqNo = -1
            End If
        Loop While ilSeqNo > 0
    End If

    ilUpper = UBound(tlLLC)
    tlLLC(ilUpper).iDay = -1
    ilCUpper = LBound(tlCLLC)  'Save lower bound for current
    ilPUpper = LBound(tlPLLC)  'Save lower bound for pending
    clGStartTime = gTimeToCurrency(slSTime, False)
    clGEndTime = gTimeToCurrency(slETime, True) - 1
    Do While (tlCLLC(ilCUpper).iDay = ilDay) Or ((tlPLLC(ilPUpper).iDay = ilDay))
        ilCFindEvt = False
        Do While tlCLLC(ilCUpper).iDay = ilDay
            If (tlCLLC(ilCUpper).sType = "R") Or (tlCLLC(ilCUpper).sType = "S") Or (tlCLLC(ilCUpper).sType = "P") Then
                clCEvtTime = gTimeToCurrency(tlCLLC(ilCUpper).sStartTime, False)
                gAddTimeLength tlCLLC(ilCUpper).sStartTime, tlCLLC(ilCUpper).sLength, "A", "1", slTime, slXMid
                clCEvtEndTime = gTimeToCurrency(slTime, True) - 1
                If (clCEvtEndTime < clGStartTime) Or (clCEvtTime > clGEndTime) Then
                    If (tlCLLC(ilCUpper).iDay = -1) Or (ilCUpper >= UBound(tlCLLC)) Then
                        Exit Do
                    End If
                    ilCUpper = ilCUpper + 1
                Else
                    ilCFindEvt = True
                    Exit Do
                End If
            Else
                If (tlCLLC(ilCUpper).iDay = -1) Or (ilCUpper >= UBound(tlCLLC)) Then
                    Exit Do
                End If
                ilCUpper = ilCUpper + 1
            End If
        Loop
        ilPFindEvt = False
        Do While tlPLLC(ilPUpper).iDay = ilDay
            If (tlPLLC(ilPUpper).sType = "R") Or (tlPLLC(ilPUpper).sType = "S") Or (tlPLLC(ilPUpper).sType = "P") Then
                clPEvtTime = gTimeToCurrency(tlPLLC(ilPUpper).sStartTime, False)
                gAddTimeLength tlPLLC(ilPUpper).sStartTime, tlPLLC(ilPUpper).sLength, "A", "1", slTime, slXMid
                clPEvtEndTime = gTimeToCurrency(slTime, True) - 1
                If (clPEvtEndTime < clGStartTime) Or (clPEvtTime > clGEndTime) Then
                    If (tlPLLC(ilPUpper).iDay = -1) Or (ilPUpper >= UBound(tlPLLC)) Then
                        Exit Do
                    End If
                    ilPUpper = ilPUpper + 1
                Else
                    ilPFindEvt = True
                    Exit Do
                End If
            Else
                If (tlPLLC(ilPUpper).iDay = -1) Or (ilPUpper >= UBound(tlPLLC)) Then
                    Exit Do
                End If
                ilPUpper = ilPUpper + 1
            End If
        Loop
        If ilCFindEvt And ilPFindEvt Then
            If clCEvtEndTime < clPEvtTime Then
                ilShowCurrent = 0
            ElseIf clPEvtEndTime < clCEvtTime Then
                ilShowCurrent = 1
            Else
                ilShowCurrent = -1
            End If
        ElseIf ilCFindEvt And Not ilPFindEvt Then
            ilShowCurrent = 0
        ElseIf Not ilCFindEvt And ilPFindEvt Then
            ilShowCurrent = 1
        Else
            Exit Do
        End If
        If ilShowCurrent = 0 Then
            llLvfCode = tlCLLC(ilCUpper).lLvfCode 'Val(slCode)
            ilLtfCode = tlCLLC(ilCUpper).iLtfCode
            If ilEvtType(0) Then
                tlLLC(ilUpper).iDay = tlCLLC(ilCUpper).iDay
                tlLLC(ilUpper).sType = tlCLLC(ilCUpper).sType
                tlLLC(ilUpper).sStartTime = tlCLLC(ilCUpper).sStartTime
                tlLLC(ilUpper).sLength = tlCLLC(ilCUpper).sLength
                tlLLC(ilUpper).iUnits = 0 'tlCLLC(ilCUpper).iUnits
                tlLLC(ilUpper).sName = tlCLLC(ilCUpper).sName
'                ilRet = gParseItem(tlCLLC(ilCUpper).sName, 1, "\", tlLLC(ilUpper).sName)
'                ilRet = gParseItem(tlCLLC(ilCUpper).sName, 2, "\", slCode)
                tlLLC(ilUpper).lLvfCode = tlCLLC(ilCUpper).lLvfCode 'Val(slCode)
                tlLLC(ilUpper).iLtfCode = tlCLLC(ilCUpper).iLtfCode
                tlLLC(ilUpper).iAvailInfo = 0
                tlLLC(ilUpper).iEtfCode = 0
                tlLLC(ilUpper).iEnfCode = 0
                tlLLC(ilUpper).lCefCode = 0
                tlLLC(ilUpper).lEvtIDCefCode = 0
                ilUpper = ilUpper + 1
                ReDim Preserve tlLLC(0 To ilUpper) As LLC
                tlLLC(ilUpper).iDay = -1
            End If
            'Read Lef and add to tlLLC
            slStartTime = tlCLLC(ilCUpper).sStartTime
            ilCUpper = ilCUpper + 1
        ElseIf ilShowCurrent = 1 Then
            llLvfCode = tlPLLC(ilPUpper).lLvfCode 'Val(slCode)
            ilLtfCode = tlPLLC(ilPUpper).iLtfCode
            If ilEvtType(0) Then
                tlLLC(ilUpper).iDay = tlPLLC(ilPUpper).iDay
                tlLLC(ilUpper).sType = tlPLLC(ilPUpper).sType
                tlLLC(ilUpper).sStartTime = tlPLLC(ilPUpper).sStartTime
                tlLLC(ilUpper).sLength = tlPLLC(ilPUpper).sLength
                tlLLC(ilUpper).iUnits = 0 'tlPLLC(ilPUpper).iUnits
                tlLLC(ilUpper).sName = tlPLLC(ilPUpper).sName
'                ilRet = gParseItem(tlPLLC(ilPUpper).sName, 1, "\", tlLLC(ilUpper).sName)
'                ilRet = gParseItem(tlPLLC(ilPUpper).sName, 2, "\", slCode)
                tlLLC(ilUpper).lLvfCode = tlPLLC(ilPUpper).lLvfCode 'Val(slCode)
                tlLLC(ilUpper).iLtfCode = tlPLLC(ilPUpper).iLtfCode
                tlLLC(ilUpper).iAvailInfo = 0
                tlLLC(ilUpper).iEtfCode = 0
                tlLLC(ilUpper).iEnfCode = 0
                tlLLC(ilUpper).lCefCode = 0
                tlLLC(ilUpper).lEvtIDCefCode = 0
                ilUpper = ilUpper + 1
                ReDim Preserve tlLLC(0 To ilUpper) As LLC
                tlLLC(ilUpper).iDay = -1
            End If
            slStartTime = tlPLLC(ilPUpper).sStartTime
            ilPUpper = ilPUpper + 1
        ElseIf ilShowCurrent = -1 Then  'Current and pending intersect- ignore current
            ilCUpper = ilCUpper + 1
        End If
        If (Not ilOnlyLib) And (ilShowCurrent <> -1) Then
            'Get all events for the library
            tlLefSrchKey.lLvfCode = llLvfCode
            tlLefSrchKey.iStartTime(0) = 0
            tlLefSrchKey.iStartTime(1) = 0
            tlLefSrchKey.iSeqNo = 0
            ilRet = btrGetGreaterOrEqual(hlLef, tlLef, ilLefRecLen, tlLefSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tlLef.lLvfCode = llLvfCode)
                tlLLC(ilUpper).iDay = ilDay
                tlLLC(ilUpper).lLvfCode = llLvfCode
                tlLLC(ilUpper).iLtfCode = ilLtfCode
                gUnpackLength tlLef.iStartTime(0), tlLef.iStartTime(1), "3", False, slStr
                gAddTimeLength slStartTime, slStr, "A", "1", tlLLC(ilUpper).sStartTime, slXMid
                If gTimeToCurrency(tlLLC(ilUpper).sStartTime, False) > clGEndTime Then
                    tlLLC(ilUpper).iDay = -1
                    ilRet = btrClose(hlAnf)
                    ilRet = btrClose(hlEnf)
                    ilRet = btrClose(hlLef)
                    ilRet = btrClose(hlLvf)
                    ilRet = btrClose(hlLtf)
                    ilRet = btrClose(hlLcf)
                    btrDestroy hlAnf
                    btrDestroy hlEnf
                    btrDestroy hlLef
                    btrDestroy hlLvf
                    btrDestroy hlLtf
                    btrDestroy hlLcf
                    gBuildEventDay = True
                    Exit Function
                End If
                If tlLef.iEtfCode <= 13 Then
                    ilEtfCode = tlLef.iEtfCode
                Else
                    ilEtfCode = 14
                End If
                If ilEvtType(ilEtfCode) Then
                    ilFound = True
                    Select Case tlLef.iEtfCode
                        Case 1  'Program
                            tlLLC(ilUpper).sType = "1"
                            gUnpackLength tlLef.iLen(0), tlLef.iLen(1), "3", False, tlLLC(ilUpper).sLength
                            If tlLef.iEnfCode > 0 Then
                                tlEnfSrchKey.iCode = tlLef.iEnfCode
                                ilRet = btrGetEqual(hlEnf, tlEnf, ilEnfRecLen, tlEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    tlLLC(ilUpper).sName = Trim$(tlEnf.sName)
                                Else
                                    tlLLC(ilUpper).sName = ""
                                End If
                            Else
                                tlLLC(ilUpper).sName = ""
                            End If
                            tlLLC(ilUpper).iUnits = tlLef.iMnfExcl(0)
                            tlLLC(ilUpper).iAvailInfo = tlLef.iMnfExcl(1)
                            tlLLC(ilUpper).iEtfCode = tlLef.iEtfCode
                            tlLLC(ilUpper).iEnfCode = tlLef.iEnfCode
                            tlLLC(ilUpper).lCefCode = tlLef.lCefCode
                            tlLLC(ilUpper).lEvtIDCefCode = tlLef.lEvtIDCefCode
                            tlLLC(ilUpper).sXMid = slXMid
                        Case 2, 3, 4, 5, 6, 7, 8, 9  'Avail
                            tlLLC(ilUpper).sType = Trim$(str$(tlLef.iEtfCode)) '"2"
                            tlLLC(ilUpper).iUnits = tlLef.iMaxUnits
                            gUnpackLength tlLef.iLen(0), tlLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                            tlLLC(ilUpper).iAvailInfo = 0
                            tlLLC(ilUpper).sName = Trim$(str$(tlLef.ianfCode))
                            tlAnfSrchKey.iCode = tlLef.ianfCode
                            ilRet = btrGetEqual(hlAnf, tlAnf, ilAnfRecLen, tlAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                If tlAnf.sSustain = "Y" Then
                                    tlLLC(ilUpper).iAvailInfo = tlLLC(ilUpper).iAvailInfo Or SSSUSTAINING
                                End If
                                If tlAnf.sSponsorship = "Y" Then
                                    tlLLC(ilUpper).iAvailInfo = tlLLC(ilUpper).iAvailInfo Or SSSPONSORSHIP
                                End If
                                If tlAnf.sBookLocalFeed = "L" Then
                                    tlLLC(ilUpper).iAvailInfo = tlLLC(ilUpper).iAvailInfo Or SSLOCALONLY
                                End If
                                If tlAnf.sBookLocalFeed = "F" Then
                                    tlLLC(ilUpper).iAvailInfo = tlLLC(ilUpper).iAvailInfo Or SSFEEDONLY
                                End If
                            End If
                            If slXMid = "Y" Then
                                tlLLC(ilUpper).iAvailInfo = tlLLC(ilUpper).iAvailInfo Or SSXMID
                            End If
                            tlLLC(ilUpper).iEtfCode = tlLef.iEtfCode
                            tlLLC(ilUpper).iEnfCode = tlLef.iEnfCode
                            tlLLC(ilUpper).lCefCode = tlLef.lCefCode
                            tlLLC(ilUpper).lEvtIDCefCode = tlLef.lEvtIDCefCode
                        Case 10, 11, 12, 13  'Page eject, Line space 1, 2 or 3
                            tlLLC(ilUpper).sType = Chr$(Asc("A") + tlLef.iEtfCode - 10)
                            tlLLC(ilUpper).iUnits = 0
                            tlLLC(ilUpper).sLength = ""
                            If tlLef.iEtfCode = 10 Then
                                tlLLC(ilUpper).sName = Trim$(str$(tlLef.ianfCode))
                            Else
                                tlLLC(ilUpper).sName = "0"
                            End If
                            tlLLC(ilUpper).iAvailInfo = 0
                            'tlLLC(ilUpper).sName = ""
                            tlLLC(ilUpper).iEtfCode = tlLef.iEtfCode
                            tlLLC(ilUpper).iEnfCode = tlLef.iEnfCode
                            tlLLC(ilUpper).lCefCode = tlLef.lCefCode
                            tlLLC(ilUpper).lEvtIDCefCode = tlLef.lEvtIDCefCode
                            tlLLC(ilUpper).sXMid = slXMid
                        Case Else   'Other
                            tlLLC(ilUpper).sType = "Y"
                            gUnpackLength tlLef.iLen(0), tlLef.iLen(1), "3", True, tlLLC(ilUpper).sLength
                            tlLLC(ilUpper).iUnits = 0
                            tlLLC(ilUpper).iAvailInfo = 0
                            If tlLef.iEnfCode > 0 Then
                                tlEnfSrchKey.iCode = tlLef.iEnfCode
                                ilRet = btrGetEqual(hlEnf, tlEnf, ilEnfRecLen, tlEnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    tlLLC(ilUpper).sName = Trim$(tlEnf.sName)
                                Else
                                    tlLLC(ilUpper).sName = ""
                                End If
                            Else
                                tlLLC(ilUpper).sName = ""
                            End If
                            tlLLC(ilUpper).iEtfCode = tlLef.iEtfCode
                            tlLLC(ilUpper).iEnfCode = tlLef.iEnfCode
                            tlLLC(ilUpper).lCefCode = tlLef.lCefCode
                            tlLLC(ilUpper).lEvtIDCefCode = tlLef.lEvtIDCefCode
                            tlLLC(ilUpper).sXMid = slXMid
                    End Select
                Else
                    ilFound = False
                End If
                If ilFound Then
                    ilUpper = ilUpper + 1
                    ReDim Preserve tlLLC(0 To ilUpper) As LLC
                    tlLLC(ilUpper).iDay = -1
                End If
                ilRet = btrGetNext(hlLef, tlLef, ilLefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
    Loop
    ilRet = btrClose(hlAnf)
    ilRet = btrClose(hlEnf)
    ilRet = btrClose(hlLef)
    ilRet = btrClose(hlLvf)
    ilRet = btrClose(hlLtf)
    ilRet = btrClose(hlLcf)
    btrDestroy hlAnf
    btrDestroy hlEnf
    btrDestroy hlLef
    btrDestroy hlLvf
    btrDestroy hlLtf
    btrDestroy hlLcf
    gBuildEventDay = True
    Exit Function

    ilRet = btrClose(hlAnf)
    ilRet = btrClose(hlEnf)
    ilRet = btrClose(hlLef)
    ilRet = btrClose(hlLvf)
    ilRet = btrClose(hlLtf)
    ilRet = btrClose(hlLcf)
    btrDestroy hlAnf
    btrDestroy hlEnf
    btrDestroy hlLef
    btrDestroy hlLvf
    btrDestroy hlLtf
    btrDestroy hlLcf
    gBuildEventDay = False
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gBuildEventSpotDay              *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build a day into EVTINFO        *
'*                                                     *
'*******************************************************
Function gBuildEventSpotDay(ilVehCode As Integer, ilVpfIndex As Integer, slDate As String, slStartTime As String, slEndTime As String, ilGameNo As Integer, tlVcf0() As VCF, tlVcf6() As VCF, tlVcf7() As VCF, ilEvtType() As Integer, tlEvtSpot() As EVTINFO) As Integer
'
'   ilRet = gBuildEventSpotDay (ilVehCode, ilVpfIndex, slDate, slStartTime, slEndTime, ilEvtType, tlLLC)
'
'   Where:
'       ilVehCode (I)-Vehicle code number
'       ilVpfIndex(I)- Vehicle option index
'       slDate (I)- Date that events are to be obtained
'       slStartTime (I)- Start Time (included)
'       slEndTime (I)- End time (not included)
'       ilEvtType (I)- Array of which events are to be included (True or False)
'                       Index description
'                         0   Library
'                         1   Program event
'                         2   Contract avail
'                         3   Open BB
'                         4   Floating BB
'                         5   Close BB
'                         6   Cmml promo
'                         7   Feed avail
'                         8   PSA avail
'                         9   Promo avail
'                         10  Page eject
'                         11  Line space 1
'                         12  Line space 2
'                         13  Line space 3
'                         14  Other event types
'
'       tlEvt (O)- Event and spot records
'
'       tgCommAdf(I)-Advertiser code, name, abbr
'       tgCompMnf(I)-Multi-name code, name, abbr
'
    Dim slTime As String
    Dim ilType As Integer
    Dim slType As String
    Dim ilRet As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim llSsfRecPos As Long
    Dim ilEvt As Integer
    Dim ilDay As Integer
    Dim ilSpot As Integer
    Dim slStr As String
    Dim sLCP As String
    Dim ilUpperBound As Integer
    Dim ilLoop As Integer
    Dim ilUnits As Integer
    Dim slUnits As String
    Dim ilSec As Integer
    Dim slSpotLen As String
    Dim llDate As Long
    Dim llTime As Long
    Dim ilVehIndex As Integer
    Dim ilConflictSpot As Integer
    Dim slVTime As String
    Dim slChar As String    'One character
    Dim slPrgName As String
    'Spot summary
    Dim hlSsf As Integer        'Spot summary file handle
    Dim tlSsfSrchKey As SSFKEY0 'SSF key record image
    Dim ilSsfRecLen As Integer  'SSF record length
    Dim slSplitNetworkType As String
    Dim ilChkLen As Integer
    Dim ilLen As Integer
    Dim tlSpot As CSPOTSS

    'Spot detail record information
    Dim hlSdf As Integer        'Spot detail file handle
    Dim hlAnf As Integer
    Dim hlCff As Integer
    Dim hlVsf As Integer
    Dim hlRaf As Integer

    'Vehicle conflict
    Dim hlVcf As Integer        'Contract header file handle

    Dim hlSmf As Integer

    ReDim tlEvtSpot(0 To 1) As EVTINFO
    ReDim tlLLC(0 To 0) As LLC  'Image
    ilUpperBound = 1
    ilSsfRecLen = Len(tmSsf)  'Get and save SSF record length
    hlSsf = CBtrvTable(ONEHANDLE)        'Create SSF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Ssf(1)"
        ilRet = btrClose(hlSsf)
        btrDestroy hlSsf
        gBuildEventSpotDay = False
        Exit Function
    End If
    imSdfRecLen = Len(tmSdf)  'Get and save SDF record length
    hlSdf = CBtrvTable(ONEHANDLE)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Sdf(2)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        gBuildEventSpotDay = False
        Exit Function
    End If
    imAnfRecLen = Len(tmAnf)  'Get and save ANF record length
    hlAnf = CBtrvTable(ONEHANDLE)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Anf(3)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        gBuildEventSpotDay = False
        Exit Function
    End If
    imCHFRecLen = Len(tmChf)  'Get and save ADF record length
    hmCHF = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Chf(4)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        gBuildEventSpotDay = False
        Exit Function
    End If
    imClfRecLen = Len(tmClf)  'Get and save CLF record length
    hmClf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Clf(5)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        gBuildEventSpotDay = False
        Exit Function
    End If
    imRdfRecLen = Len(tmRdf)  'Get and save RDF record length
    hmRdf = CBtrvTable(ONEHANDLE)        'Create RDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Rdf(6)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        gBuildEventSpotDay = False
        Exit Function
    End If
    imVefRecLen = Len(tmVef)  'Get and save VEF record length
    hmVef = CBtrvTable(ONEHANDLE)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Vef(7)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVef)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        btrDestroy hmVef
        gBuildEventSpotDay = False
        Exit Function
    End If
    hlVcf = CBtrvTable(ONEHANDLE)        'Create VCF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlVcf, "", sgDBPath & "Vcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Vcf(8)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hlVcf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        btrDestroy hmVef
        btrDestroy hlVcf
        gBuildEventSpotDay = False
        Exit Function
    End If
    hlSmf = CBtrvTable(ONEHANDLE)        'Create SMF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Smf(9)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hlVcf)
        ilRet = btrClose(hlSmf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        btrDestroy hmVef
        btrDestroy hlVcf
        btrDestroy hlSmf
        gBuildEventSpotDay = False
        Exit Function
    End If
    hlVsf = CBtrvTable(ONEHANDLE)        'Create SMF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Vsf(10)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hlVcf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlVsf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        btrDestroy hmVef
        btrDestroy hlVcf
        btrDestroy hlSmf
        btrDestroy hlVsf
        gBuildEventSpotDay = False
        Exit Function
    End If
    hlCff = CBtrvTable(ONEHANDLE)        'Create SMF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Cff(11)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hlVcf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlVsf)
        ilRet = btrClose(hlCff)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        btrDestroy hmVef
        btrDestroy hlVcf
        btrDestroy hlSmf
        btrDestroy hlVsf
        btrDestroy hlCff
        gBuildEventSpotDay = False
        Exit Function
    End If
    imCrfRecLen = Len(tmCrf)  'Get and save ANF record length
    hmCrf = CBtrvTable(ONEHANDLE)        'Create SMF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Crf(12)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hlVcf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlVsf)
        ilRet = btrClose(hlCff)
        ilRet = btrClose(hmCrf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        btrDestroy hmVef
        btrDestroy hlVcf
        btrDestroy hlSmf
        btrDestroy hlVsf
        btrDestroy hlCff
        btrDestroy hmCrf
        gBuildEventSpotDay = False
        Exit Function
    End If
    imSifRecLen = Len(tmSif)  'Get and save ANF record length
    hmSif = CBtrvTable(ONEHANDLE)        'Create SMF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Sif(13)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hlVcf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlVsf)
        ilRet = btrClose(hlCff)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmSif)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        btrDestroy hmVef
        btrDestroy hlVcf
        btrDestroy hlSmf
        btrDestroy hlVsf
        btrDestroy hlCff
        btrDestroy hmCrf
        btrDestroy hmSif
        gBuildEventSpotDay = False
        Exit Function
    End If
    imCxfRecLen = Len(tmCxf)  'Get and save ANF record length
    hmCxf = CBtrvTable(ONEHANDLE)        'Create SMF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Cxf(14)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hlVcf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlVsf)
        ilRet = btrClose(hlCff)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmCxf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        btrDestroy hmVef
        btrDestroy hlVcf
        btrDestroy hlSmf
        btrDestroy hlVsf
        btrDestroy hlCff
        btrDestroy hmCrf
        btrDestroy hmSif
        btrDestroy hmCxf
        gBuildEventSpotDay = False
        Exit Function
    End If
    imFsfRecLen = Len(tmFsf)  'Get and save ANF record length
    hmFsf = CBtrvTable(ONEHANDLE)        'Create SMF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Fsf(14)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hlVcf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlVsf)
        ilRet = btrClose(hlCff)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmCxf)
        ilRet = btrClose(hmFsf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        btrDestroy hmVef
        btrDestroy hlVcf
        btrDestroy hlSmf
        btrDestroy hlVsf
        btrDestroy hlCff
        btrDestroy hmCrf
        btrDestroy hmSif
        btrDestroy hmCxf
        btrDestroy hmFsf
        gBuildEventSpotDay = False
        Exit Function
    End If
    imFnfRecLen = Len(tmFnf)  'Get and save ANF record length
    hmFnf = CBtrvTable(ONEHANDLE)        'Create SMF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmFnf, "", sgDBPath & "Fnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Fsf(14)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hlVcf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlVsf)
        ilRet = btrClose(hlCff)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmCxf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmFnf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        btrDestroy hmVef
        btrDestroy hlVcf
        btrDestroy hlSmf
        btrDestroy hlVsf
        btrDestroy hlCff
        btrDestroy hmCrf
        btrDestroy hmSif
        btrDestroy hmCxf
        btrDestroy hmFsf
        btrDestroy hmFnf
        gBuildEventSpotDay = False
        Exit Function
    End If
    imPrfRecLen = Len(tmPrf)  'Get and save ANF record length
    hmPrf = CBtrvTable(ONEHANDLE)        'Create SMF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Fsf(14)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hlVcf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlVsf)
        ilRet = btrClose(hlCff)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmCxf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmFnf)
        ilRet = btrClose(hmPrf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        btrDestroy hmVef
        btrDestroy hlVcf
        btrDestroy hlSmf
        btrDestroy hlVsf
        btrDestroy hlCff
        btrDestroy hmCrf
        btrDestroy hmSif
        btrDestroy hmCxf
        btrDestroy hmFsf
        btrDestroy hmFnf
        btrDestroy hmPrf
        gBuildEventSpotDay = False
        Exit Function
    End If
    hlRaf = CBtrvTable(ONEHANDLE)        'Create SMF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gBuildEventSpotDay-Open Fsf(14)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hlVcf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlVsf)
        ilRet = btrClose(hlCff)
        ilRet = btrClose(hmCrf)
        ilRet = btrClose(hmSif)
        ilRet = btrClose(hmCxf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmFnf)
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hlRaf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlAnf
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmRdf
        btrDestroy hmVef
        btrDestroy hlVcf
        btrDestroy hlSmf
        btrDestroy hlVsf
        btrDestroy hlCff
        btrDestroy hmCrf
        btrDestroy hmSif
        btrDestroy hmCxf
        btrDestroy hmFsf
        btrDestroy hmFnf
        btrDestroy hmPrf
        btrDestroy hlRaf
        gBuildEventSpotDay = False
        Exit Function
    End If
    gObtainVirtVehList
    tmAnf.iCode = -1
    ilType = ilGameNo
    slType = "O"
    sLCP = "C"
    If (ilEvtType(0) = True) Or (ilEvtType(10) = True) Or (ilEvtType(11) = True) Or (ilEvtType(12) = True) Or (ilEvtType(13) = True) Or (ilEvtType(14) = True) Then
        ilRet = gBuildEventDay(ilType, sLCP, ilVehCode, slDate, slStartTime, slEndTime, ilEvtType(), tlLLC())
    End If
    llDate = gDateValue(slDate)
    gObtainVcf hlVcf, ilVehCode, llDate, tlVcf0(), tlVcf6(), tlVcf7()
    ilDay = gWeekDayStr(slDate)
    gPackDate slDate, ilLogDate0, ilLogDate1
    ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
    tlSsfSrchKey.iType = ilType
    tlSsfSrchKey.iVefCode = ilVehCode
    tlSsfSrchKey.iDate(0) = ilLogDate0
    tlSsfSrchKey.iDate(1) = ilLogDate1
    tlSsfSrchKey.iStartTime(0) = 0
    tlSsfSrchKey.iStartTime(1) = 0
    ilRet = gSSFGetEqual(hlSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVehCode) And (tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1)
        ilRet = gSSFGetPosition(hlSsf, llSsfRecPos)
        'Loop thru Ssf and move records to tm--Evt
        ilEvt = 1
        Do While ilEvt <= tmSsf.iCount
           LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
            If tmProg.iRecType = 1 Then    'Program
                gUnpackTime tmProg.iStartTime(0), tmProg.iStartTime(1), "A", "1", slTime
                If (InStr(slTime, "A") > 0) Or (InStr(slTime, "P") > 0) Then
                    slStr = Left$(slTime, Len(slTime) - 1)
                Else
                    slStr = slTime
                End If
                tlEvtSpot(ilUpperBound).sShow = slStr
                tlEvtSpot(ilUpperBound).iType = 1
                tlEvtSpot(ilUpperBound).lTime = CLng(gTimeToCurrency(slTime, False))
                gUnpackTime tmProg.iEndTime(0), tmProg.iEndTime(1), "A", "1", slTime
                If (InStr(slTime, "A") > 0) Or (InStr(slTime, "P") > 0) Then
                    slStr = Left$(slTime, Len(slTime) - 1)
                Else
                    slStr = slTime
                End If
                tlEvtSpot(ilUpperBound).sShow = Trim$(tlEvtSpot(ilUpperBound).sShow) & "-" & slStr
                tlEvtSpot(ilUpperBound).lChfCode = 0
                tlEvtSpot(ilUpperBound).lSdfCode = 0
                tlEvtSpot(ilUpperBound).lLen = CLng(gTimeToCurrency(slTime, True))
                tlEvtSpot(ilUpperBound).iUnits = tmProg.iLtfCode
                tlEvtSpot(ilUpperBound).lInfo = tmProg.lLvfCode
                tlEvtSpot(ilUpperBound).iMnfComp1 = tmProg.iMnfExcl(0)
                tlEvtSpot(ilUpperBound).iMnfComp2 = tmProg.iMnfExcl(1)
                tlEvtSpot(ilUpperBound).iSsfIndex = ilEvt
                tlEvtSpot(ilUpperBound).lSsfRecPos = llSsfRecPos
                tlEvtSpot(ilUpperBound).lSdfCode = 0
                tlEvtSpot(ilUpperBound).sShow = Trim$(tlEvtSpot(ilUpperBound).sShow) & "  " & "Prog"
                tlEvtSpot(ilUpperBound).sPrice = ""
                If ilEvtType(tmProg.iRecType) Then
                    ilUpperBound = ilUpperBound + 1
                    'ReDim Preserve tlEvtSpot(1 To ilUpperBound) As EVTINFO
                    ReDim Preserve tlEvtSpot(0 To ilUpperBound) As EVTINFO
                End If
            ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then 'Avail
               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                tlEvtSpot(ilUpperBound).iType = tmAvail.iRecType
                tlEvtSpot(ilUpperBound).lTime = CLng(gTimeToCurrency(slTime, False))
                tlEvtSpot(ilUpperBound).lChfCode = 0
                tlEvtSpot(ilUpperBound).lSdfCode = 0
                tlEvtSpot(ilUpperBound).lLen = tmAvail.iLen
                tlEvtSpot(ilUpperBound).iUnits = tmAvail.iAvInfo And &H1F
                tlEvtSpot(ilUpperBound).iLineInfo = tmAvail.iAvInfo And (Not &H1F)
                tlEvtSpot(ilUpperBound).lInfo = tmAvail.ianfCode
                tlEvtSpot(ilUpperBound).iSsfIndex = ilEvt
                tlEvtSpot(ilUpperBound).lSsfRecPos = llSsfRecPos
                tlEvtSpot(ilUpperBound).lSdfCode = 0
                If (InStr(slTime, "A") > 0) Or (InStr(slTime, "P") > 0) Then
                    slTime = Left$(slTime, Len(slTime) - 1)
                Else
                    slTime = slTime
                End If
                If (tgVpf(ilVpfIndex).sSSellOut = "B") Or (tgVpf(ilVpfIndex).sSSellOut = "U") Or (tgVpf(ilVpfIndex).sSSellOut = "M") Then
                    slStr = slTime & " " & Trim$(str$(tlEvtSpot(ilUpperBound).iUnits)) & "/" & Trim$(str$(tmAvail.iLen))     'Time/Units/Seconds
                    ilUnits = tlEvtSpot(ilUpperBound).iUnits
                    slUnits = Trim$(str$(ilUnits)) & ".0"   'For units as thirty
                    ilSec = tmAvail.iLen
                Else
                    slStr = slTime & " " & Trim$(str$(tlEvtSpot(ilUpperBound).iUnits))     'Time/Units
                    ilUnits = tlEvtSpot(ilUpperBound).iUnits
                    ilSec = 0
                End If
                If tmAnf.iCode <> tmAvail.ianfCode Then
                    tmAnfSrchKey.iCode = tmAvail.ianfCode
                    ilRet = btrGetEqual(hlAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
                    If ilRet = BTRV_ERR_NONE Then
                        slStr = slStr & "  " & Trim$(tmAnf.sName)
                    End If
                Else
                    slStr = slStr & "  " & Trim$(tmAnf.sName)
                End If
                slPrgName = gGetPrgName(llDate, llTime, tgPaf())
                slStr = slStr & " " & slPrgName
                Select Case tmAvail.iRecType
                    Case 2  'Contract Avail
                    Case 3  'Open BB
                        slStr = slStr & " OBB"
                    Case 4  'Floater
                    Case 5  'Close BB
                        slStr = slStr & " CBB"
                    Case 6  'Cmml Promo
                        slStr = slStr & " Cmml Promo"
                    Case 7  'Feed
                        slStr = slStr & " Feed"
                    Case 8  'PSA
                        slStr = slStr & " PSA"
                    Case 9  'Promo
                        slStr = slStr & " Promo"
                End Select
                tlEvtSpot(ilUpperBound).sShow = slStr     'Time/Units/Seconds
                tlEvtSpot(ilUpperBound).sPrice = ""
                tlEvtSpot(ilUpperBound).sPrgName = slPrgName
                If ilEvtType(tmAvail.iRecType) Then
                    ilUpperBound = ilUpperBound + 1
                    ReDim Preserve tlEvtSpot(0 To ilUpperBound) As EVTINFO
                End If
                'Loop on spots, then add conflicting spots
                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                    ilEvt = ilEvt + 1
                   LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    tlEvtSpot(ilUpperBound).iType = 100
                    tlEvtSpot(ilUpperBound).sPrgName = slPrgName
                    'Get Sdf
                    tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                    ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        ilRet = btrGetPosition(hlSdf, tlEvtSpot(ilUpperBound).lTime)
                        If ilRet = BTRV_ERR_NONE Then
                            slSplitNetworkType = ""
                            ilLen = (tmSpot.iPosLen And &HFFF)
                            If (tmSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
                                slSplitNetworkType = "P"
                                '10/6/08:  Determine max length
                                For ilChkLen = ilSpot + 1 To tmAvail.iNoSpotsThis Step 1
                                   LSet tlSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilChkLen - ilSpot)
                                    If (tlSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                        If ilLen < (tlSpot.iPosLen And &HFFF) Then
                                            ilLen = (tlSpot.iPosLen And &HFFF)
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next ilChkLen
                            ElseIf (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                slSplitNetworkType = "S"
                            End If
                            gBuildSpotInfo tmSdf, hmCHF, hmClf, hmRdf, hlSmf, hmSif, hmCxf, hlRaf, slDate, slTime, ilEvt, llSsfRecPos, tlEvtSpot(ilUpperBound), True, hlCff, hmVef, hlVsf, True, hmFsf, hmFnf, hmPrf, slSplitNetworkType
                            If (tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                If (tgVpf(ilVpfIndex).sSSellOut = "B") Or (tgVpf(ilVpfIndex).sSSellOut = "U") Then
                                    ilUnits = ilUnits - 1
                                    ilSec = ilSec - ilLen   '(tmSpot.iPosLen And &HFFF)
                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                    ilUnits = ilUnits - 1
                                    ilSec = ilSec - ilLen   '(tmSpot.iPosLen And &HFFF)
                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                    slSpotLen = Trim$(str$(ilLen))  'tmSpot.iPosLen And &HFFF))
                                    slStr = gDivStr(slSpotLen, "30.0")
                                    slUnits = gSubStr(slUnits, slSpotLen)
                                End If
                            End If
                            ilUpperBound = ilUpperBound + 1
                            'ReDim Preserve tlEvtSpot(1 To ilUpperBound) As EVTINFO
                            ReDim Preserve tlEvtSpot(0 To ilUpperBound) As EVTINFO
                        End If
                    End If
                Next ilSpot
                'Create remaining record
                If (tgVpf(ilVpfIndex).sSSellOut = "B") Or (tgVpf(ilVpfIndex).sSSellOut = "U") Or (tgVpf(ilVpfIndex).sSSellOut = "M") Then
                    If (ilUnits > 0) And (ilSec > 0) Then
                        tlEvtSpot(ilUpperBound).iType = 99
                        tlEvtSpot(ilUpperBound).lChfCode = 0
                        tlEvtSpot(ilUpperBound).lSdfCode = 0
                        tlEvtSpot(ilUpperBound).lLen = ilSec
                        tlEvtSpot(ilUpperBound).iUnits = ilUnits
                        slStr = "  " & Trim$(str$(tlEvtSpot(ilUpperBound).iUnits)) & "/" & Trim$(str$(ilSec))     'Time/Units/Seconds
                        tlEvtSpot(ilUpperBound).sShow = slStr     'Time/Units/Seconds
                        tlEvtSpot(ilUpperBound).sPrice = ""
                        tlEvtSpot(ilUpperBound).sPrgName = slPrgName
                        ilUpperBound = ilUpperBound + 1
                        'ReDim Preserve tlEvtSpot(1 To ilUpperBound) As EVTINFO
                        ReDim Preserve tlEvtSpot(0 To ilUpperBound) As EVTINFO
                    End If
                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                    If gCompNumberStr(slUnits, "0.0") > 0 Then
                        tlEvtSpot(ilUpperBound).iType = 99
                        tlEvtSpot(ilUpperBound).lChfCode = 0
                        tlEvtSpot(ilUpperBound).lSdfCode = 0
                        tlEvtSpot(ilUpperBound).lLen = 0
                        tlEvtSpot(ilUpperBound).iUnits = 0
                        tlEvtSpot(ilUpperBound).sShow = "  " & slUnits     'Time/Units/Seconds
                        tlEvtSpot(ilUpperBound).sPrice = ""
                        tlEvtSpot(ilUpperBound).sPrgName = slPrgName
                        ilUpperBound = ilUpperBound + 1
                        'ReDim Preserve tlEvtSpot(1 To ilUpperBound) As EVTINFO
                        ReDim Preserve tlEvtSpot(0 To ilUpperBound) As EVTINFO
                    End If
                End If
                Select Case ilDay
                    Case 0 To 4
                        For ilLoop = LBound(tlVcf0) To UBound(tlVcf0) - 1 Step 1
                            If (tmAvail.iTime(0) = tlVcf0(ilLoop).iSellTime(0)) And (tmAvail.iTime(1) = tlVcf0(ilLoop).iSellTime(1)) Then
                                'Obtain spots for other vehicles
                                For ilVehIndex = 1 To 5 Step 1
                                    If tlVcf0(ilLoop).iCSV(ilVehIndex - 1) > 0 Then
                                        ilConflictSpot = False
                                        tmSdfSrchKey.iVefCode = tlVcf0(ilLoop).iCSV(ilVehIndex - 1)
                                        tmSdfSrchKey.iDate(0) = ilLogDate0
                                        tmSdfSrchKey.iDate(1) = ilLogDate1
                                        tmSdfSrchKey.iTime(0) = tlVcf0(ilLoop).iCST(0, ilVehIndex - 1)
                                        tmSdfSrchKey.iTime(1) = tlVcf0(ilLoop).iCST(1, ilVehIndex - 1)
                                        tmSdfSrchKey.sSchStatus = ""
                                        ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tlVcf0(ilLoop).iCSV(ilVehIndex - 1)) And (tmSdf.iDate(0) = ilLogDate0) And (tmSdf.iDate(1) = ilLogDate1) And (tmSdf.iTime(0) = tlVcf0(ilLoop).iCST(0, ilVehIndex - 1)) And (tmSdf.iTime(1) = tlVcf0(ilLoop).iCST(1, ilVehIndex - 1))
                                            If (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                                                If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then    'Add spot
                                                    tlEvtSpot(ilUpperBound).iType = 101
                                                    ilRet = btrGetPosition(hlSdf, tlEvtSpot(ilUpperBound).lTime)
                                                    If (ilRet = BTRV_ERR_NONE) Then
                                                        ilConflictSpot = True
                                                        gUnpackTime tlVcf0(ilLoop).iCST(0, ilVehIndex - 1), tlVcf0(ilLoop).iCST(1, ilVehIndex - 1), "A", "1", slTime
                                                        If (InStr(slTime, "A") > 0) Or (InStr(slTime, "P") > 0) Then
                                                            slTime = Left$(slTime, Len(slTime) - 1)
                                                        Else
                                                            slTime = slTime
                                                        End If
                                                        slSplitNetworkType = ""
                                                        gBuildSpotInfo tmSdf, hmCHF, hmClf, hmRdf, hlSmf, hmSif, hmCxf, hlRaf, slDate, slTime, 0, 0, tlEvtSpot(ilUpperBound), False, hlCff, hmVef, hlVsf, False, hmFsf, hmFnf, hmPrf, slSplitNetworkType
                                                        'Made vehicle name and spot time to sShow
                                                        tlEvtSpot(ilUpperBound).sSpot = tlEvtSpot(ilUpperBound).sShow
                                                        If tmVef.iCode <> tmSdf.iVefCode Then
                                                            tmVefSrchKey.iCode = tmSdf.iVefCode
                                                            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                            If ilRet <> BTRV_ERR_NONE Then
                                                                tmVef.sName = "Error"
                                                            End If
                                                        End If
                                                        '11/17/11: Replace vehicle name with Vehicle Station Code if it exist
                                                        If Trim$(tmVef.sCodeStn) <> "" Then
                                                            slStr = Trim$(tmVef.sCodeStn)
                                                        Else
                                                            slStr = Trim$(Left$(tmVef.sName, 20))
                                                        End If
                                                        'tlEvtSpot(ilUpperBound).sShow = "  " & Trim$(Left$(tmVef.sName, 20)) & "," & slTime & "," & LTrim$(tlEvtSpot(ilUpperBound).sShow)
                                                        tlEvtSpot(ilUpperBound).sShow = "  " & slStr & "," & slTime & "," & LTrim$(tlEvtSpot(ilUpperBound).sShow)
                                                        slStr = LTrim$(tlEvtSpot(ilUpperBound).sSpot)
                                                        'Remove spot length
                                                        Do
                                                            slChar = Left$(slStr, 1)
                                                            slStr = right$(slStr, Len(slStr) - 1)
                                                        Loop While (slChar <> " ") And (slChar <> "+") And (slChar <> "-") And (slChar <> "!") And (slChar <> "@") And (slChar <> "#")
                                                        tlEvtSpot(ilUpperBound).sSpot = "  " & Left$(tmVef.sName, 2) & ", " & slStr
                                                        ilUpperBound = ilUpperBound + 1
                                                        'ReDim Preserve tlEvtSpot(1 To ilUpperBound) As EVTINFO
                                                        ReDim Preserve tlEvtSpot(0 To ilUpperBound) As EVTINFO
                                                    End If
                                                End If
                                            End If
                                            ilRet = btrGetNext(hlSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                        Loop
                                        If Not ilConflictSpot Then
                                            tlEvtSpot(ilUpperBound).iType = 101
                                            tlEvtSpot(ilUpperBound).iLineInfo = 0
                                            tlEvtSpot(ilUpperBound).lChfCode = 0
                                            tlEvtSpot(ilUpperBound).lSdfCode = 0
                                            tlEvtSpot(ilUpperBound).lLen = 0
                                            tlEvtSpot(ilUpperBound).iUnits = 0
                                            tlEvtSpot(ilUpperBound).lInfo = 0
                                            tlEvtSpot(ilUpperBound).iLineNo = 0
                                            tlEvtSpot(ilUpperBound).iMnfComp1 = 0
                                            tlEvtSpot(ilUpperBound).iMnfComp2 = 0
                                            tlEvtSpot(ilUpperBound).iSsfIndex = 0
                                            tlEvtSpot(ilUpperBound).lSsfRecPos = 0
                                            tlEvtSpot(ilUpperBound).lSdfCode = 0
                                            tlEvtSpot(ilUpperBound).sCntrType = ""
                                            tlEvtSpot(ilUpperBound).sSpot = ""
                                            tlEvtSpot(ilUpperBound).sShow = ""
                                            tlEvtSpot(ilUpperBound).sPrice = ""
                                            If tmVef.iCode <> tlVcf0(ilLoop).iCSV(ilVehIndex - 1) Then
                                                tmVefSrchKey.iCode = tlVcf0(ilLoop).iCSV(ilVehIndex - 1)
                                                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    tmVef.sName = "Error"
                                                End If
                                            End If
                                            gUnpackTime tlVcf0(ilLoop).iCST(0, ilVehIndex - 1), tlVcf0(ilLoop).iCST(1, ilVehIndex - 1), "A", "1", slVTime
                                            tlEvtSpot(ilUpperBound).sSpot = "  " & Left$(tmVef.sName, 2) & "," & slVTime
                                            '11/17/11: Replace vehicle name with Vehicle Station Code if it exist
                                            If Trim$(tmVef.sCodeStn) <> "" Then
                                                slStr = Trim$(tmVef.sCodeStn)
                                            Else
                                                slStr = Trim$(Left$(tmVef.sName, 20))
                                            End If
                                            'tlEvtSpot(ilUpperBound).sShow = "  " & Trim$(Left$(tmVef.sName, 20)) & "," & slVTime
                                            tlEvtSpot(ilUpperBound).sShow = "  " & slStr & "," & slVTime
                                            ilUpperBound = ilUpperBound + 1
                                            'ReDim Preserve tlEvtSpot(1 To ilUpperBound) As EVTINFO
                                            ReDim Preserve tlEvtSpot(0 To ilUpperBound) As EVTINFO
                                        End If
                                    End If
                                Next ilVehIndex
                                Exit For
                            End If
                        Next ilLoop
                    Case 5
                        For ilLoop = LBound(tlVcf6) To UBound(tlVcf6) - 1 Step 1
                            If (tmAvail.iTime(0) = tlVcf6(ilLoop).iSellTime(0)) And (tmAvail.iTime(1) = tlVcf6(ilLoop).iSellTime(1)) Then
                                'Obtain spots for other vehicles
                                For ilVehIndex = 1 To 5 Step 1
                                    If tlVcf6(ilLoop).iCSV(ilVehIndex - 1) > 0 Then
                                        ilConflictSpot = False
                                        tmSdfSrchKey.iVefCode = tlVcf6(ilLoop).iCSV(ilVehIndex - 1)
                                        tmSdfSrchKey.iDate(0) = ilLogDate0
                                        tmSdfSrchKey.iDate(1) = ilLogDate1
                                        tmSdfSrchKey.iTime(0) = tlVcf6(ilLoop).iCST(0, ilVehIndex - 1)
                                        tmSdfSrchKey.iTime(1) = tlVcf6(ilLoop).iCST(1, ilVehIndex - 1)
                                        tmSdfSrchKey.sSchStatus = ""
                                        ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tlVcf6(ilLoop).iCSV(ilVehIndex - 1)) And (tmSdf.iDate(0) = ilLogDate0) And (tmSdf.iDate(1) = ilLogDate1) And (tmSdf.iTime(0) = tlVcf6(ilLoop).iCST(0, ilVehIndex - 1)) And (tmSdf.iTime(1) = tlVcf6(ilLoop).iCST(1, ilVehIndex - 1))
                                            If (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                                                If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then     'Add spot
                                                    tlEvtSpot(ilUpperBound).iType = 101
                                                    ilRet = btrGetPosition(hlSdf, tlEvtSpot(ilUpperBound).lTime)
                                                    If (ilRet = BTRV_ERR_NONE) Then
                                                        ilConflictSpot = True
                                                        gUnpackTime tlVcf6(ilLoop).iCST(0, ilVehIndex - 1), tlVcf6(ilLoop).iCST(1, ilVehIndex - 1), "A", "1", slTime
                                                        If (InStr(slTime, "A") > 0) Or (InStr(slTime, "P") > 0) Then
                                                            slTime = Left$(slTime, Len(slTime) - 1)
                                                        Else
                                                            slTime = slTime
                                                        End If
                                                        slSplitNetworkType = ""
                                                        gBuildSpotInfo tmSdf, hmCHF, hmClf, hmRdf, hlSmf, hmSif, hmCxf, hlRaf, slDate, slTime, 0, 0, tlEvtSpot(ilUpperBound), False, hlCff, hmVef, hlVsf, False, hmFsf, hmFnf, hmPrf, slSplitNetworkType
                                                        'Made vehicle name and spot time to sShow
                                                        tlEvtSpot(ilUpperBound).sSpot = tlEvtSpot(ilUpperBound).sShow
                                                        If tmVef.iCode <> tmSdf.iVefCode Then
                                                            tmVefSrchKey.iCode = tmSdf.iVefCode
                                                            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                            If ilRet <> BTRV_ERR_NONE Then
                                                                tmVef.sName = "Error"
                                                            End If
                                                        End If
                                                        '1/20/12: Replace vehicle name with Vehicle Station Code if it exist
                                                        If Trim$(tmVef.sCodeStn) <> "" Then
                                                            slStr = Trim$(tmVef.sCodeStn)
                                                        Else
                                                            slStr = Trim$(Left$(tmVef.sName, 20))
                                                        End If
                                                        'tlEvtSpot(ilUpperBound).sShow = "  " & Trim$(Left$(tmVef.sName, 20)) & "," & slTime & "," & LTrim$(tlEvtSpot(ilUpperBound).sShow)
                                                        tlEvtSpot(ilUpperBound).sShow = "  " & slStr & "," & slTime & "," & LTrim$(tlEvtSpot(ilUpperBound).sShow)
                                                        slStr = LTrim$(tlEvtSpot(ilUpperBound).sSpot)
                                                        'Remove spot length
                                                        Do While (Asc(slStr) <> Asc(" ")) And (Asc(slStr) <> Asc("+")) And (Asc(slStr) <> Asc("-")) And (Asc(slStr) <> Asc("!")) And (Asc(slStr) <> Asc("@")) And (Asc(slStr) <> Asc("#"))
                                                            slStr = right$(slStr, Len(slStr) - 1)
                                                        Loop
                                                        tlEvtSpot(ilUpperBound).sSpot = "  " & Left$(tmVef.sName, 2) & "," & slStr
                                                        ilUpperBound = ilUpperBound + 1
                                                        'ReDim Preserve tlEvtSpot(1 To ilUpperBound) As EVTINFO
                                                        ReDim Preserve tlEvtSpot(0 To ilUpperBound) As EVTINFO
                                                    End If
                                                End If
                                            End If
                                            ilRet = btrGetNext(hlSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                        Loop
                                        If Not ilConflictSpot Then
                                            tlEvtSpot(ilUpperBound).iType = 101
                                            tlEvtSpot(ilUpperBound).iLineInfo = 0
                                            tlEvtSpot(ilUpperBound).lChfCode = 0
                                            tlEvtSpot(ilUpperBound).lSdfCode = 0
                                            tlEvtSpot(ilUpperBound).lLen = 0
                                            tlEvtSpot(ilUpperBound).iUnits = 0
                                            tlEvtSpot(ilUpperBound).lInfo = 0
                                            tlEvtSpot(ilUpperBound).iLineNo = 0
                                            tlEvtSpot(ilUpperBound).iMnfComp1 = 0
                                            tlEvtSpot(ilUpperBound).iMnfComp2 = 0
                                            tlEvtSpot(ilUpperBound).iSsfIndex = 0
                                            tlEvtSpot(ilUpperBound).lSsfRecPos = 0
                                            tlEvtSpot(ilUpperBound).lSdfCode = 0
                                            tlEvtSpot(ilUpperBound).sCntrType = ""
                                            tlEvtSpot(ilUpperBound).sSpot = ""
                                            tlEvtSpot(ilUpperBound).sShow = ""
                                            tlEvtSpot(ilUpperBound).sPrice = ""
                                            If tmVef.iCode <> tlVcf6(ilLoop).iCSV(ilVehIndex - 1) Then
                                                tmVefSrchKey.iCode = tlVcf6(ilLoop).iCSV(ilVehIndex - 1)
                                                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    tmVef.sName = "Error"
                                                End If
                                            End If
                                            gUnpackTime tlVcf6(ilLoop).iCST(0, ilVehIndex - 1), tlVcf6(ilLoop).iCST(1, ilVehIndex - 1), "A", "1", slVTime
                                            tlEvtSpot(ilUpperBound).sSpot = "  " & Left$(tmVef.sName, 2) & "," & slVTime
                                            '1/20/12: Replace vehicle name with Vehicle Station Code if it exist
                                            If Trim$(tmVef.sCodeStn) <> "" Then
                                                slStr = Trim$(tmVef.sCodeStn)
                                            Else
                                                slStr = Trim$(Left$(tmVef.sName, 20))
                                            End If
                                            'tlEvtSpot(ilUpperBound).sShow = "  " & Trim$(Left$(tmVef.sName, 20)) & "," & slVTime
                                            tlEvtSpot(ilUpperBound).sShow = "  " & slStr & "," & slVTime
                                            ilUpperBound = ilUpperBound + 1
                                            'ReDim Preserve tlEvtSpot(1 To ilUpperBound) As EVTINFO
                                            ReDim Preserve tlEvtSpot(0 To ilUpperBound) As EVTINFO
                                        End If
                                    End If
                                Next ilVehIndex
                                Exit For
                            End If
                        Next ilLoop
                    Case 6
                        For ilLoop = LBound(tlVcf7) To UBound(tlVcf7) - 1 Step 1
                            If (tmAvail.iTime(0) = tlVcf7(ilLoop).iSellTime(0)) And (tmAvail.iTime(1) = tlVcf7(ilLoop).iSellTime(1)) Then
                                'Obtain spots for other vehicles
                                For ilVehIndex = 1 To 5 Step 1
                                    If tlVcf7(ilLoop).iCSV(ilVehIndex - 1) > 0 Then
                                        ilConflictSpot = False
                                        tmSdfSrchKey.iVefCode = tlVcf7(ilLoop).iCSV(ilVehIndex - 1)
                                        tmSdfSrchKey.iDate(0) = ilLogDate0
                                        tmSdfSrchKey.iDate(1) = ilLogDate1
                                        tmSdfSrchKey.iTime(0) = tlVcf7(ilLoop).iCST(0, ilVehIndex - 1)
                                        tmSdfSrchKey.iTime(1) = tlVcf7(ilLoop).iCST(1, ilVehIndex - 1)
                                        tmSdfSrchKey.sSchStatus = ""
                                        ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tlVcf7(ilLoop).iCSV(ilVehIndex - 1)) And (tmSdf.iDate(0) = ilLogDate0) And (tmSdf.iDate(1) = ilLogDate1) And (tmSdf.iTime(0) = tlVcf7(ilLoop).iCST(0, ilVehIndex - 1)) And (tmSdf.iTime(1) = tlVcf7(ilLoop).iCST(1, ilVehIndex - 1))
                                            If (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                                                If (tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O") Then     'Add spot
                                                    tlEvtSpot(ilUpperBound).iType = 101
                                                    ilRet = btrGetPosition(hlSdf, tlEvtSpot(ilUpperBound).lTime)
                                                    If (ilRet = BTRV_ERR_NONE) Then
                                                        ilConflictSpot = True
                                                        gUnpackTime tlVcf7(ilLoop).iCST(0, ilVehIndex - 1), tlVcf7(ilLoop).iCST(1, ilVehIndex - 1), "A", "1", slTime
                                                        If (InStr(slTime, "A") > 0) Or (InStr(slTime, "P") > 0) Then
                                                            slTime = Left$(slTime, Len(slTime) - 1)
                                                        Else
                                                            slTime = slTime
                                                        End If
                                                        slSplitNetworkType = ""
                                                        gBuildSpotInfo tmSdf, hmCHF, hmClf, hmRdf, hlSmf, hmSif, hmCxf, hlRaf, slDate, slTime, 0, 0, tlEvtSpot(ilUpperBound), False, hlCff, hmVef, hlVsf, False, hmFsf, hmFnf, hmPrf, slSplitNetworkType
                                                        'Made vehicle name and spot time to sShow
                                                        tlEvtSpot(ilUpperBound).sSpot = tlEvtSpot(ilUpperBound).sShow
                                                        If tmVef.iCode <> tmSdf.iVefCode Then
                                                            tmVefSrchKey.iCode = tmSdf.iVefCode
                                                            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                            If ilRet <> BTRV_ERR_NONE Then
                                                                tmVef.sName = "Error"
                                                            End If
                                                        End If
                                                        '1/20/12: Replace vehicle name with Vehicle Station Code if it exist
                                                        If Trim$(tmVef.sCodeStn) <> "" Then
                                                            slStr = Trim$(tmVef.sCodeStn)
                                                        Else
                                                            slStr = Trim$(Left$(tmVef.sName, 20))
                                                        End If
                                                        'tlEvtSpot(ilUpperBound).sShow = "  " & Trim$(Left$(tmVef.sName, 20)) & "," & slTime & "," & LTrim$(tlEvtSpot(ilUpperBound).sShow)
                                                        tlEvtSpot(ilUpperBound).sShow = "  " & slStr & "," & slTime & "," & LTrim$(tlEvtSpot(ilUpperBound).sShow)
                                                        slStr = LTrim$(tlEvtSpot(ilUpperBound).sSpot)
                                                        'Remove spot length
                                                        Do While (Asc(slStr) <> Asc(" ")) And (Asc(slStr) <> Asc("+")) And (Asc(slStr) <> Asc("-")) And (Asc(slStr) <> Asc("!")) And (Asc(slStr) <> Asc("@")) And (Asc(slStr) <> Asc("#"))
                                                            slStr = right$(slStr, Len(slStr) - 1)
                                                        Loop
                                                        tlEvtSpot(ilUpperBound).sSpot = "  " & Left$(tmVef.sName, 2) & "," & slStr
                                                        ilUpperBound = ilUpperBound + 1
                                                        'ReDim Preserve tlEvtSpot(1 To ilUpperBound) As EVTINFO
                                                        ReDim Preserve tlEvtSpot(0 To ilUpperBound) As EVTINFO
                                                    End If
                                                End If
                                            End If
                                            ilRet = btrGetNext(hlSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                        Loop
                                        If Not ilConflictSpot Then
                                            tlEvtSpot(ilUpperBound).iType = 101
                                            tlEvtSpot(ilUpperBound).iLineInfo = 0
                                            tlEvtSpot(ilUpperBound).lChfCode = 0
                                            tlEvtSpot(ilUpperBound).lSdfCode = 0
                                            tlEvtSpot(ilUpperBound).lLen = 0
                                            tlEvtSpot(ilUpperBound).iUnits = 0
                                            tlEvtSpot(ilUpperBound).lInfo = 0
                                            tlEvtSpot(ilUpperBound).iLineNo = 0
                                            tlEvtSpot(ilUpperBound).iMnfComp1 = 0
                                            tlEvtSpot(ilUpperBound).iMnfComp2 = 0
                                            tlEvtSpot(ilUpperBound).iSsfIndex = 0
                                            tlEvtSpot(ilUpperBound).lSsfRecPos = 0
                                            tlEvtSpot(ilUpperBound).lSdfCode = 0
                                            tlEvtSpot(ilUpperBound).sCntrType = ""
                                            tlEvtSpot(ilUpperBound).sSpot = ""
                                            tlEvtSpot(ilUpperBound).sShow = ""
                                            tlEvtSpot(ilUpperBound).sPrice = ""
                                            If tmVef.iCode <> tlVcf7(ilLoop).iCSV(ilVehIndex - 1) Then
                                                tmVefSrchKey.iCode = tlVcf7(ilLoop).iCSV(ilVehIndex - 1)
                                                ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    tmVef.sName = "Error"
                                                End If
                                            End If
                                            gUnpackTime tlVcf7(ilLoop).iCST(0, ilVehIndex - 1), tlVcf7(ilLoop).iCST(1, ilVehIndex - 1), "A", "1", slVTime
                                            tlEvtSpot(ilUpperBound).sSpot = "  " & Left$(tmVef.sName, 2) & "," & slVTime
                                            '1/20/12: Replace vehicle name with Vehicle Station Code if it exist
                                            If Trim$(tmVef.sCodeStn) <> "" Then
                                                slStr = Trim$(tmVef.sCodeStn)
                                            Else
                                                slStr = Trim$(Left$(tmVef.sName, 20))
                                            End If
                                            'tlEvtSpot(ilUpperBound).sShow = "  " & Trim$(Left$(tmVef.sName, 20)) & "," & slVTime
                                            tlEvtSpot(ilUpperBound).sShow = "  " & slStr & "," & slVTime
                                            ilUpperBound = ilUpperBound + 1
                                            'ReDim Preserve tlEvtSpot(1 To ilUpperBound) As EVTINFO
                                            ReDim Preserve tlEvtSpot(0 To ilUpperBound) As EVTINFO
                                        End If
                                    End If
                                Next ilVehIndex
                                Exit For
                            End If
                        Next ilLoop
                End Select
            End If
            ilEvt = ilEvt + 1
        Loop
        ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
        ilRet = gSSFGetNext(hlSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilRet = btrClose(hlSsf)
    ilRet = btrClose(hlSdf)
    ilRet = btrClose(hlAnf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hlVcf)
    ilRet = btrClose(hlSmf)
    ilRet = btrClose(hlVsf)
    ilRet = btrClose(hlCff)
    ilRet = btrClose(hmCrf)
    ilRet = btrClose(hmSif)
    ilRet = btrClose(hmCxf)
    ilRet = btrClose(hmFsf)
    ilRet = btrClose(hmFnf)
    ilRet = btrClose(hmPrf)
    ilRet = btrClose(hlRaf)
    btrDestroy hlSsf
    btrDestroy hlSdf
    btrDestroy hlAnf
    btrDestroy hmCHF
    btrDestroy hmClf
    btrDestroy hmRdf
    btrDestroy hmVef
    btrDestroy hlVcf
    btrDestroy hlSmf
    btrDestroy hlVsf
    btrDestroy hlCff
    btrDestroy hmCrf
    btrDestroy hmSif
    btrDestroy hmCxf
    btrDestroy hmFsf
    btrDestroy hmFnf
    btrDestroy hmPrf
    btrDestroy hlRaf
    gBuildEventSpotDay = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gBuildLinkArray                 *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build array of airing or selling*
'*                     vehicles associated with selling*
'*                     airing vehicle                  *
'*                                                     *
'*******************************************************
Sub gBuildLinkArray(hlVlf As Integer, tlVef As VEF, slInDate As String, ilVefCode() As Integer)
    Dim slDate As String
    Dim llDate As Long
    Dim ilCode As Integer
    Dim ilPrevCode As Integer
    Dim ilFound As Integer
    Dim ilTest As Integer
    Dim ilLoop As Integer
    ReDim ilVefCode(0 To 0) As Integer
    If (tlVef.sType = "S") Or (tlVef.sType = "A") Then
        slDate = gObtainPrevMonday(slInDate)
        llDate = gDateValue(slDate)
        ReDim tmTVlf(0 To 0) As VLF
        gObtainVlf tlVef.sType, hlVlf, tlVef.iCode, llDate, tmTVlf()
        ilPrevCode = -1
        For ilLoop = LBound(tmTVlf) To UBound(tmTVlf) - 1 Step 1
            If (tlVef.sType = "S") Then
                ilCode = tmTVlf(ilLoop).iAirCode
            Else
                ilCode = tmTVlf(ilLoop).iSellCode
            End If
            If ilPrevCode <> ilCode Then
                ilFound = False
                For ilTest = 0 To UBound(ilVefCode) - 1 Step 1
                    If ilCode = ilVefCode(ilTest) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    ilVefCode(UBound(ilVefCode)) = ilCode
                    ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                End If
                ilPrevCode = ilCode
            End If
        Next ilLoop
        llDate = llDate + 5
        ReDim tmTVlf(0 To 0) As VLF
        gObtainVlf tlVef.sType, hlVlf, tlVef.iCode, llDate, tmTVlf()
        ilPrevCode = -1
        For ilLoop = LBound(tmTVlf) To UBound(tmTVlf) - 1 Step 1
            If (tlVef.sType = "S") Then
                ilCode = tmTVlf(ilLoop).iAirCode
            Else
                ilCode = tmTVlf(ilLoop).iSellCode
            End If
            If ilPrevCode <> ilCode Then
                ilFound = False
                For ilTest = 0 To UBound(ilVefCode) - 1 Step 1
                    If ilCode = ilVefCode(ilTest) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    ilVefCode(UBound(ilVefCode)) = ilCode
                    ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                End If
                ilPrevCode = ilCode
            End If
        Next ilLoop
        llDate = llDate + 1
        ReDim tmTVlf(0 To 0) As VLF
        gObtainVlf tlVef.sType, hlVlf, tlVef.iCode, llDate, tmTVlf()
        ilPrevCode = -1
        For ilLoop = LBound(tmTVlf) To UBound(tmTVlf) - 1 Step 1
            If (tlVef.sType = "S") Then
                ilCode = tmTVlf(ilLoop).iAirCode
            Else
                ilCode = tmTVlf(ilLoop).iSellCode
            End If
            If ilPrevCode <> ilCode Then
                ilFound = False
                For ilTest = 0 To UBound(ilVefCode) - 1 Step 1
                    If ilCode = ilVefCode(ilTest) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilTest
                If Not ilFound Then
                    ilVefCode(UBound(ilVefCode)) = ilCode
                    ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                End If
                ilPrevCode = ilCode
            End If
        Next ilLoop
        Erase tmTVlf
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gBuildSpotInfo                  *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build spot info for             *
'*                     gBuildEventSpotDay              *
'*
'*  11-23-04    Change reading of SMF to use KEY2 to
'*              speed up Spot screen (and other places
'*              that use this routine
'*******************************************************
Sub gBuildSpotInfo(tlSdf As SDF, hlChf As Integer, hlClf As Integer, hlRdf As Integer, hlSmf As Integer, hlSif As Integer, hlCxf As Integer, hlRaf, slDate As String, slTime As String, ilEvt As Integer, llSsfRecPos As Long, tlEvtSpot As EVTINFO, ilIncludePrice As Integer, hlCff As Integer, hlVef As Integer, hlVsf As Integer, ilTimeTest As Integer, hlFsf As Integer, hlFnf As Integer, hlPrf As Integer, slSplitNetworkType As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  tlSmfSrchKey                                                                          *
'******************************************************************************************

'
'   gBuildSpotInfo
'   Where:
'
'       tlSdf (I)- Contains spot to create EVTINFO record for
    Dim slAdvt As String
    Dim slComp1 As String
    Dim slComp2 As String
    Dim slProduct As String
    Dim slProgTime As String
    Dim ilPendingLine As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slStr As String
    Dim slSymbol As String
    ReDim ilDate(0 To 1) As Integer
    ReDim ilTime(0 To 1) As Integer
    Dim tlSmf As SMF            'SMF record image
    Dim tlSmfSrchKey2 As LONGKEY0   'sdf code
    Dim ilSmfRecLen As Integer  'SMF record length
    Dim slBonusOnInv As String
    Dim ilVpfIndex As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim tlCff As CFF
    Dim ilRafRecLen As Integer
    Dim tlRaf As RAF
    Dim tlRafSrchKey0 As LONGKEY0

    tlEvtSpot.iLineInfo = 0
    tlEvtSpot.lRafCode = 0
    slProduct = ""
    imCHFRecLen = Len(tmChf)
    imClfRecLen = Len(tmClf)
    imRdfRecLen = Len(tmRdf)
    ilSmfRecLen = Len(tlSmf)
    imCrfRecLen = Len(tmCrf)
    imSifRecLen = Len(tmSif)
    imFsfRecLen = Len(tmFsf)
    ilRafRecLen = Len(tlRaf)
    If tlSdf.lChfCode > 0 Then
        tmChfSrchKey.lCode = tlSdf.lChfCode
        ilRet = btrGetEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet = BTRV_ERR_NONE Then
            slProduct = Trim$(Left$(tmChf.sProduct, 15))
        End If
        slProgTime = ""
        ilPendingLine = False
        tmClfSrchKey.lChfCode = tlSdf.lChfCode
        tmClfSrchKey.iLine = tlSdf.iLineNo
        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
        ilRet = btrGetGreaterOrEqual(hlClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tlSdf.lChfCode) And (tmClf.iLine = tlSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
            ilPendingLine = True
            ilRet = btrGetNext(hlClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tlSdf.lChfCode) And (tmClf.iLine = tlSdf.iLineNo) Then
            tlEvtSpot.lRafCode = tmClf.lRafCode
            ilVpfIndex = gBinarySearchVpf(tmClf.iVefCode)
            If (ilVpfIndex <> -1) And (tlSdf.iGameNo <= 0) Then
                If ((tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And (tgVpf(ilVpfIndex).sGMedium <> "S") Then
                    gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slStartTime
                    gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slEndTime
                    slProgTime = slStartTime & "-" & slEndTime
                    ilRet = gGetSpotFlight(tlSdf, tmClf, hlCff, hlSmf, tlCff)
                    If ilRet Then
                        slProgTime = gDayNames(tlCff.iDay(), tlCff.sXDay(), 1, slStr) & " " & slProgTime
                    End If
                Else
                    tmRdfSrchKey.iCode = tmClf.iRdfCode  ' Daypart File Code
                    ilRet = btrGetEqual(hlRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        slProgTime = Trim$(tmRdf.sName)
                    End If
                End If
            Else
                tmRdfSrchKey.iCode = tmClf.iRdfCode  ' Daypart File Code
                ilRet = btrGetEqual(hlRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE Then
                    slProgTime = Trim$(tmRdf.sName)
                End If
            End If
        End If
    Else
        tmFsfSrchKey0.lCode = tlSdf.lFsfCode
        ilRet = btrGetEqual(hlFsf, tmFsf, imFsfRecLen, tmFsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        gMoveFeedToCntr tmFsf, tmRdf, tmChf, tmClf, tmFCff(), hlFnf, hlPrf
    End If
    slAdvt = "Missing"
    'For ilLoop = LBound(tgCommAdf) To UBound(tgCommAdf) Step 1
    '    If tlSdf.iAdfCode = tgCommAdf(ilLoop).iCode Then
        ilLoop = gBinarySearchAdf(tlSdf.iAdfCode)
        If ilLoop <> -1 Then
            slAdvt = Trim$(tgCommAdf(ilLoop).sAbbr)
            slBonusOnInv = tgCommAdf(ilLoop).sBonusOnInv
    '        Exit For
        End If
    'Next ilLoop
    slComp1 = ""
    If tmChf.iMnfComp(0) > 0 Then
        For ilLoop = LBound(tgCompMnf) To UBound(tgCompMnf) Step 1
            If tmChf.iMnfComp(0) = tgCompMnf(ilLoop).iCode Then
                'Jim requested full name on 3/2/03
                'slComp1 = Trim$(tgCompMnf(ilLoop).sUnitType)
                slComp1 = Trim$(tgCompMnf(ilLoop).sName)
                Exit For
            End If
        Next ilLoop
    End If
    slComp2 = ""
    If tmChf.iMnfComp(1) > 0 Then
        For ilLoop = LBound(tgCompMnf) To UBound(tgCompMnf) Step 1
            If tmChf.iMnfComp(1) = tgCompMnf(ilLoop).iCode Then
                'Jim requested full name on 3/2/03
                'slComp2 = Trim$(tgCompMnf(ilLoop).sUnitType)
                slComp2 = Trim$(tgCompMnf(ilLoop).sName)
                Exit For
            End If
        Next ilLoop
    End If
    tlEvtSpot.lLen = tlSdf.lChfCode
    tlEvtSpot.iUnits = tmClf.iLen
    tlEvtSpot.lChfCode = tlSdf.lChfCode
    tlEvtSpot.lFsfCode = tlSdf.lFsfCode
    tlEvtSpot.lInfo = tlSdf.iAdfCode
    If ilPendingLine Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or 1
    End If
    If tlSdf.sSchStatus = "G" Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or 2
    End If
    If tlSdf.sSchStatus = "O" Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or 4
    End If
    For ilLoop = LBound(igVirtVefCode) To UBound(igVirtVefCode) - 1 Step 1
        If tmClf.iVefCode = igVirtVefCode(ilLoop) Then
            tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or 8
            Exit For
        End If
    Next ilLoop
    If (tmClf.sType = "H") Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or 8
    ElseIf (tmClf.sType = "O") Or (tmClf.sType = "A") Or (tmClf.sType = "E") Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or 8
    End If
    If tlSdf.sSpotType = "X" Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H10
    End If
    'If (tlSdf.sSpotType = "T") Or (tlSdf.sSpotType = "Q") Or (tlSdf.sSpotType = "S") Or (tlSdf.sSpotType = "M") Then
    'If (tlSdf.sSpotType = "S") Or (tlSdf.sSpotType = "M") Then
    If (tmChf.sType = "M") Or (tmChf.sType = "S") Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H20
    End If
    If (tlSdf.sSchStatus = "O") Or (tlSdf.sSchStatus = "G") Then
        'Determine if spot original from another vehicle- check smf
        '11-23-04 change which key to use, from key 0 to key2 which is spot code

        'tlSmfSrchKey.lChfCode = tlSdf.lChfCode
        'tlSmfSrchKey.iLineNo = tlSdf.iLineNo
        'tlSmfSrchKey.iMissedDate(0) = 0 'sch date =tlSdf.iDate(0)
        'tlSmfSrchKey.iMissedDate(1) = 0 'sch date =tlSdf.iDate(1)

        tlSmfSrchKey2.lCode = tlSdf.lCode
        ilSmfRecLen = Len(tlSmf)
        'ilRet = btrGetGreaterOrEqual(hlSmf, tlSmf, ilSmfRecLen, tlSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
        ilRet = btrGetGreaterOrEqual(hlSmf, tlSmf, ilSmfRecLen, tlSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get current record

        Do While (ilRet = BTRV_ERR_NONE) And (tlSmf.lChfCode = tlSdf.lChfCode) And (tlSmf.iLineNo = tlSdf.iLineNo)
            'If (tlSmf.sSchStatus = tlSdf.sSchStatus) And (tlSmf.iActualDate(0) = tlSdf.iDate(0)) And (tlSmf.iActualDate(1) = tlSdf.iDate(1)) And (tlSmf.iActualTime(0) = tlSdf.iTime(0)) And (tlSmf.iActualTime(1) = tlSdf.iTime(1)) Then
            If tlSdf.lCode = tlSmf.lSdfCode Then
                If tlSdf.iVefCode <> tlSmf.iOrigSchVef Then
                    tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H40
                End If
                If (tlSmf.lMtfCode > 0) And (lgMtfNoRecs > 0) Then
                    tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H200
                End If
                Exit Do
            End If
            ilRet = btrGetNext(hlSmf, tlSmf, ilSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    ElseIf (tlSdf.sSchStatus = "M") And (tlSdf.sTracer = "*") Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H200
    End If
    If (tmChf.sType = "V") Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H400
    End If
    If (tmChf.sStatus = "H") Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H80
    End If
    If tgSaf(0).sHideDemoOnBR = "Y" And tmChf.sHideDemo = "Y" Then  'Impressions
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H4000
    End If
    tlEvtSpot.iLineNo = tlSdf.iLineNo
    tlEvtSpot.iMnfComp1 = tmChf.iMnfComp(0)
    tlEvtSpot.iMnfComp2 = tmChf.iMnfComp(1)
    tlEvtSpot.iSsfIndex = ilEvt
    tlEvtSpot.lSsfRecPos = llSsfRecPos
    tlEvtSpot.lSdfCode = tlSdf.lCode
    tlEvtSpot.sCntrType = tmChf.sType
    slSymbol = " "
    If tmClf.iPosition = 1 Then
        slSymbol = "@"
    End If
    If tmClf.sSoloAvail = "Y" Then
        slSymbol = "#"
    End If
    If tlSdf.sSpotType = "X" Then
        'If tlSdf.sPriceType <> "N" Then
        If tlSdf.sPriceType = "+" Then
            slSymbol = "+"  '">"
        ElseIf tlSdf.sPriceType = "-" Then
            slSymbol = "-"  '"<"
        Else
            If slBonusOnInv <> "N" Then
                slSymbol = "+"
            Else
                slSymbol = "-"
            End If
        End If
'    Else
'        slSymbol = " "
    End If
    If ilTimeTest Then
        gPackDate slDate, ilDate(0), ilDate(1)
        gPackTime slTime, ilTime(0), ilTime(1)
        If (ilDate(0) <> tlSdf.iDate(0)) Or (ilDate(1) <> tlSdf.iDate(1)) Or (ilTime(0) <> tlSdf.iTime(0)) Or (ilTime(1) <> tlSdf.iTime(1)) Then
            slSymbol = "!"
            gUnpackDate tlSdf.iDate(0), tlSdf.iDate(1), tlEvtSpot.sAirDate
            gUnpackTime tlSdf.iTime(0), tlSdf.iTime(1), "A", "1", tlEvtSpot.sAirTime
        Else
            tlEvtSpot.sAirDate = ""
            tlEvtSpot.sAirTime = ""
        End If
    Else
        tlEvtSpot.sAirDate = ""
        tlEvtSpot.sAirTime = ""
    End If
    slStr = slTime & "   L" & Trim$(str$(tlSdf.iLineNo)) & str$(tmClf.iLen) & slSymbol
    tlEvtSpot.sSpot = slStr     'Time/Line number Length
    'Test tgSpf to see if advt/prod or just prod-- code later
    'tlEvtSpot.sShow = "  " & Trim$(Str$(tmClf.iLen)) & "/" & slAdvt & "/" & slProduct
    If (tgSpf.sUseProdSptScr = "A") Or (tmChf.lSifCode <= 0) Then
        If slProduct <> "" Then
            tlEvtSpot.sShow = "  " & Trim$(str$(tmClf.iLen)) & slSymbol & slAdvt & "," & slProduct
        Else
            tlEvtSpot.sShow = "  " & Trim$(str$(tmClf.iLen)) & slSymbol & slAdvt
        End If
    Else
        'tmAdf is not used so it does not have to be set
        tmAdf.sAbbr = slAdvt
        tlEvtSpot.sShow = "  " & Trim$(str$(tmClf.iLen)) & slSymbol & gGetShortTitle(hlVsf, hlClf, hlSif, tmChf, tmAdf, tlSdf)
        'tmCrfSrchKey.sRotType = "A"
        'tmCrfSrchKey.iEtfCode = 0
        'tmCrfSrchKey.iEnfCode = 0
        'tmCrfSrchKey.iAdfCode = tlSdf.iAdfCode
        'tmCrfSrchKey.lChfCode = tlSdf.lChfCode
        'ilCrfVefCode = gGetCrfVefCode(hlClf, tlSdf)
        'tmCrfSrchKey.iVefCode = ilCrfVefCode
        'tmCrfSrchKey.iRotNo = tlSdf.iRotNo
        'ilRet = btrGetEqual(hlCrf, tmCrf, imCrfRecLen, tmCrfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)    'Get first record as starting point of extend operation
        'If ilRet = BTRV_ERR_NONE Then
        '    tmSifSrchKey.lCode = tmCrf.lSifCode
        '    ilRet = btrGetEqual(hlSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        'Else
        '    tmSifSrchKey.lCode = tmChf.lSifCode
        '    ilRet = btrGetEqual(hlSif, tmSif, imSifRecLen, tmSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        'End If
        'If ilRet = BTRV_ERR_NONE Then
        '    tlEvtSpot.sShow = "  " & Trim$(Str$(tmClf.iLen)) & slSymbol & Trim$(tmSif.sName)
        'Else
        '    If slProduct <> "" Then
        '        tlEvtSpot.sShow = "  " & Trim$(Str$(tmClf.iLen)) & slSymbol & slAdvt & "," & slProduct
        '    Else
        '        tlEvtSpot.sShow = "  " & Trim$(Str$(tmClf.iLen)) & slSymbol & slAdvt
        '    End If
        'End If
    End If
    If slComp1 <> "" Then
        tlEvtSpot.sShow = RTrim$(tlEvtSpot.sShow) & " " & slComp1
    End If
    If slComp2 <> "" Then
        tlEvtSpot.sShow = RTrim$(tlEvtSpot.sShow) & "," & slComp2
    End If
    If slProgTime <> "" Then
        tlEvtSpot.sShow = RTrim$(tlEvtSpot.sShow) & " " & slProgTime
    End If
    If ilIncludePrice Then
        If tlSdf.sSpotType = "X" Then
            'If tlSdf.sPriceType <> "N" Then
            If tlSdf.sPriceType = "+" Then
                tlEvtSpot.sPrice = "+ Fill"  '"> Fill"
            ElseIf tlSdf.sPriceType = "-" Then
                tlEvtSpot.sPrice = "- Fill" '"< Fill"
            Else
                If slBonusOnInv <> "N" Then
                    tlEvtSpot.sPrice = "+ Fill"
                Else
                    tlEvtSpot.sPrice = "- Fill"
                End If
            End If
        Else
            If tlSdf.lChfCode > 0 Then
                ilRet = gGetSpotPrice(tlSdf, tmClf, hlCff, hlSmf, hlVef, hlVsf, tlEvtSpot.sPrice)
            Else
                tlEvtSpot.sPrice = "Feed"
            End If
        End If
    Else
        tlEvtSpot.sPrice = ""
    End If
    'Comments
    tlEvtSpot.lchfcxfCode = 0
    tlEvtSpot.lchfcxfInt = 0
    tlEvtSpot.lClfCxfCode = 0
    If tmChf.lCxfCode > 0 Then
        imCxfRecLen = Len(tmCxf)
        tmCxfSrchKey.lCode = tmChf.lCxfCode
        ilRet = btrGetEqual(hlCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        If ilRet = BTRV_ERR_NONE Then
            If tmCxf.sShSpot = "Y" Then
                tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H100
                tlEvtSpot.lchfcxfCode = tmChf.lCxfCode
            End If
        End If
    End If
     If tmChf.lCxfInt > 0 Then
        imCxfRecLen = Len(tmCxf)
        tmCxfSrchKey.lCode = tmChf.lCxfInt
        ilRet = btrGetEqual(hlCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        If ilRet = BTRV_ERR_NONE Then
            If tmCxf.sShSpot = "Y" Then
                tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H100
                tlEvtSpot.lchfcxfInt = tmChf.lCxfInt
            End If
        End If
    End If
   If (tmClf.lCxfCode > 0) Then
        imCxfRecLen = Len(tmCxf)
        tmCxfSrchKey.lCode = tmClf.lCxfCode
        ilRet = btrGetEqual(hlCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        If ilRet = BTRV_ERR_NONE Then
            If tmCxf.sShSpot = "Y" Then
                tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H100
                tlEvtSpot.lClfCxfCode = tmClf.lCxfCode
            End If
        End If
    End If
    If slSplitNetworkType = "P" Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H800
    ElseIf slSplitNetworkType = "S" Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H1000
    End If
    If (tmClf.iBBOpenLen > 0) Or (tmClf.iBBCloseLen > 0) Then
        tlEvtSpot.iLineInfo = tlEvtSpot.iLineInfo Or &H2000
    End If
    tlEvtSpot.sLiveCopy = tmClf.sLiveCopy
    tlEvtSpot.sPtType = tlSdf.sPtType
    tlEvtSpot.lCopyCode = tlSdf.lCopyCode
    tlEvtSpot.iRotNo = tlSdf.iRotNo
    tlEvtSpot.sNetRegionAbbr = ""
    If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
        If tmClf.lRafCode > 0 Then
            tlRafSrchKey0.lCode = tmClf.lRafCode
            ilRet = btrGetEqual(hlRaf, tlRaf, ilRafRecLen, tlRafSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If ilRet = BTRV_ERR_NONE Then
                If Trim(tlRaf.sAbbr) = "" Then
                    tlEvtSpot.sNetRegionAbbr = Left$(tlRaf.sName, 5)
                Else
                    tlEvtSpot.sNetRegionAbbr = tlRaf.sAbbr
                End If
            End If
        End If
    End If
    tlEvtSpot.sCITFlag = ""
    tlEvtSpot.sCopyCIT = ""
    Erase tmFCff
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gChgSchSpot                     *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Preempt spot from Ssf           *
'*                                                     *
'*******************************************************
'12/29/14: ahndle fills as Outsides
'Function gChgSchSpot(slSchStatus As String, hlSdf As Integer, tlSdf As SDF, hlSmf As Integer, ilGameNo As Integer, tlSmf As SMF, hlSsf As Integer, tlSsf As SSF, llSsfMemDate As Long, llSsfRecPos As Long, Optional hlGsf As Integer = 0, Optional hlGhf As Integer = 0) As Integer
Function gChgSchSpot(slSchStatus As String, hlSdf As Integer, tlSdf As SDF, hlSmf As Integer, ilGameNo As Integer, tlSmf As SMF, hlSsf As Integer, tlSsf As SSF, llSsfMemDate As Long, llSsfRecPos As Long, hlSxf As Integer, Optional hlGsf As Integer = 0, Optional hlGhf As Integer = 0, Optional ilFillSourceVefCode As Integer = 0) As Integer
'
'   ilRet = gChgSchSpot(slSchStatus, hlSdf, tlSdf, hlSmf, ilGameNo, tlSmf, hlSsf, tlSsf, llSsfMemDate, llSsfRecPos)
'   Where:
'       slSchStatus(I)- "M" missed; "TM" temporary missed; "C" for cancel;
'                       "H" Hidden; "D" for delete spot (remove from system)
'                       3/14/13
'                       "F" convert to fill
'                       TM will ignore contract type
'                       This field is ignored for the following:
'                       Manually sch spots will be delete: PSA; Promo; Remnant; per Inquiry; Deferred;
'                       Extra Bonus will be Deleted
'
'       hlSdf(I)- Handle from Sdf open and tlSdf as current record
'       tlSdf(I)- Spot detail record to be preempted
'       hlSmf(I)- Handle from Smf open
'       ilGameNo(I)- Game number or zero
'       tlSmf(O) - Spot MG record if found and deleted (this is returned so specification can be retained if
'                  MG moved within speicifed boundary
'       hlSsf(I)- Handle from Ssf open
'       tlSsf(I/O)- Ssf record image
'       llSsfMemDate(I/O)- Date of Ssf within tlSsf (this is used instead of converting date within tlSsf for speed)
'       llSsfRecPos(O)- Ssf record position (btrGetDirect)
'
'       ilRet(O)- True if Sdf updated
'                 False if Sdf not updated
'
'       Note: igMnfMissed(I)- Missed reason (set to "Missed" code or zero)
'
    Dim ilRet As Integer
    Dim ilMoveLoop As Integer
    Dim ilLoop As Integer
    Dim ilSpotIndex As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilFound As Integer
    Dim slTime As String
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilTime0 As Integer
    Dim ilTime1 As Integer
    Dim ilAvTime0 As Integer
    Dim ilAvTime1 As Integer
    Dim ilOrigSchVef As Integer
    Dim slOrigSchStatus As String
    Dim ilOrigGameNo As Integer
    Dim slXSpotType As String
    Dim llSvSsfRecPos As Long
    Dim ilAvInfo As Integer
    Dim tlSpot As CSPOTSS
    Dim ilGsfRecLen As Integer
    Dim tlGsf As GSF
    Dim tlGsfSrchKey3 As GSFKEY3
    Dim ilGhfRecLen As Integer
    Dim tlGhf As GHF
    Dim tlGhfSrchKey0 As LONGKEY0
    Dim llSeasonStart As Long
    Dim llSeasonEnd As Long
    Dim llSDFDate As Long
    Dim slPriceType As String

    If (Trim$(tlSdf.sAffChg) = "") Or (tlSdf.sAffChg = "N") Then
        gUnpackTime tlSdf.iTime(0), tlSdf.iTime(1), "A", "1", slTime
    Else
        slTime = "12AM"
    End If
    gUnpackDate tlSdf.iDate(0), tlSdf.iDate(1), slDate
    llDate = gDateValue(slDate)
    slOrigSchStatus = tlSdf.sSchStatus
    ilFound = False
    imSdfRecLen = Len(tlSdf)
    imSxfRecLen = Len(tmSxf)
    '3/14/13
    ilLoop = gBinarySearchAdf(tlSdf.iAdfCode)
    If ilLoop <> -1 Then
        If tgCommAdf(ilLoop).sBonusOnInv <> "N" Then
            slPriceType = "B"
        Else
            slPriceType = "N"
        End If
    Else
        slPriceType = "B"
    End If
    'Obtain SSf for Date or if in memory- re-read so handle is pointing to record
    If gObtainSsfForDateOrGame(tlSdf.iVefCode, llDate, slTime, ilGameNo, hlSsf, tlSsf, llSsfMemDate, llSsfRecPos) Then
        Do
            imSsfRecLen = Len(tlSsf)
            ilRet = gSSFGetDirect(hlSsf, tlSsf, imSsfRecLen, llSsfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = gConvertErrorCode(ilRet)
                sgErrLoc = "gChgSchSpot-GetDirect SSf(1)"
                gChgSchSpot = False
                Exit Function
            End If
            ilRet = gGetByKeyForUpdateSSF(hlSsf, tlSsf)
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = gConvertErrorCode(ilRet)
                sgErrLoc = "gChgSchSpot-GetByKey SSf(2)"
                gChgSchSpot = False
                Exit Function
            End If
            ilFound = False
            'Find matching spot and remove
            ilLoop = 1
            Do While ilLoop <= tlSsf.iCount
               LSet tmAvail = tlSsf.tPas(ADJSSFPASBZ + ilLoop)
                If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Contract Avail subrecord
                    If (tmAvail.iTime(0) = tlSdf.iTime(0)) And (tmAvail.iTime(1) = tlSdf.iTime(1)) Then
                        'Matching avail found- look for spot
                        For ilSpotIndex = ilLoop + 1 To ilLoop + tmAvail.iNoSpotsThis Step 1
                            LSet tmSpotTest = tlSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                            If tlSdf.lCode = tmSpotTest.lSdfCode Then
                                '3/14/13
                                If slSchStatus <> "F" Then
                                    ilAvTime0 = tmAvail.iTime(0)
                                    ilAvTime1 = tmAvail.iTime(1)
                                    ilAvInfo = tmAvail.iAvInfo
                                    'Move all event up, change count and update
                                    tmAvail.iNoSpotsThis = tmAvail.iNoSpotsThis - 1
                                    tlSsf.tPas(ADJSSFPASBZ + ilLoop) = tmAvail
                                    For ilMoveLoop = ilSpotIndex To tlSsf.iCount - 1 Step 1
                                        If ilMoveLoop = ilSpotIndex Then
                                            If (tmSpotTest.iRecType And SSSPLITPRI) = SSSPLITPRI Then
                                               LSet tlSpot = tlSsf.tPas(ADJSSFPASBZ + ilMoveLoop + 1)
                                                If (tlSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                                    tlSpot.iRecType = tlSpot.iRecType And (Not SSSPLITSEC)
                                                    tlSpot.iRecType = tlSpot.iRecType Or SSSPLITPRI
                                                    LSet tlSsf.tPas(ADJSSFPASBZ + ilMoveLoop + 1) = tlSpot
                                                End If
                                            End If
                                        End If
                                        tlSsf.tPas(ADJSSFPASBZ + ilMoveLoop) = tlSsf.tPas(ADJSSFPASBZ + ilMoveLoop + 1)
                                    Next ilMoveLoop
                                    'BB, bookends, donuts,... not hamdled
                                    tlSsf.iCount = tlSsf.iCount - 1
                                Else
                                    tmSpotTest.iRank = 1045
                                    LSet tlSsf.tPas(ADJSSFPASBZ + ilSpotIndex) = tmSpotTest
                                End If
                                ilFound = True
                            End If
                        Next ilSpotIndex
                        If ilFound Then
                            Exit Do
                        Else
                            ilLoop = ilLoop + 1
                        End If
                    Else
                        If (Trim$(tlSdf.sAffChg) <> "") And (tlSdf.sAffChg <> "N") Then
                            For ilSpotIndex = ilLoop + 1 To ilLoop + tmAvail.iNoSpotsThis Step 1
                                LSet tmSpotTest = tlSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                                If tlSdf.lCode = tmSpotTest.lSdfCode Then
                                    '3/14/13
                                    If slSchStatus <> "F" Then
                                        ilAvTime0 = tmAvail.iTime(0)
                                        ilAvTime1 = tmAvail.iTime(1)
                                        ilAvInfo = tmAvail.iAvInfo
                                        'Move all event up, change count and update
                                        tmAvail.iNoSpotsThis = tmAvail.iNoSpotsThis - 1
                                        tlSsf.tPas(ADJSSFPASBZ + ilLoop) = tmAvail
                                        For ilMoveLoop = ilSpotIndex To tlSsf.iCount - 1 Step 1
                                            If ilMoveLoop = ilSpotIndex Then
                                                If (tmSpotTest.iRecType And SSSPLITPRI) = SSSPLITPRI Then
                                                   LSet tlSpot = tlSsf.tPas(ADJSSFPASBZ + ilMoveLoop + 1)
                                                    If (tlSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                                        tlSpot.iRecType = tlSpot.iRecType And (Not SSSPLITSEC)
                                                        tlSpot.iRecType = tlSpot.iRecType Or SSSPLITPRI
                                                        LSet tlSsf.tPas(ADJSSFPASBZ + ilMoveLoop + 1) = tlSpot
                                                    End If
                                                End If
                                            End If
                                            tlSsf.tPas(ADJSSFPASBZ + ilMoveLoop) = tlSsf.tPas(ADJSSFPASBZ + ilMoveLoop + 1)
                                        Next ilMoveLoop
                                        'BB, bookends, donuts,... not hamdled
                                        tlSsf.iCount = tlSsf.iCount - 1
                                    Else
                                        tmSpotTest.iRank = 1045
                                        LSet tlSsf.tPas(ADJSSFPASBZ + ilSpotIndex) = tmSpotTest
                                    End If
                                    ilFound = True
                                End If
                            Next ilSpotIndex
                        End If
                        If ilFound Then
                            Exit Do
                        End If
                        ilLoop = ilLoop + 1
                    End If
                Else
                    ilLoop = ilLoop + 1
                End If
            Loop
            If Not ilFound Then
                ilLoop = 1
                Do While ilLoop <= tlSsf.iCount
                   LSet tmAvail = tlSsf.tPas(ADJSSFPASBZ + ilLoop)
                    If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Contract Avail subrecord
                        'Matching avail found- look for spot
                        For ilSpotIndex = ilLoop + 1 To ilLoop + tmAvail.iNoSpotsThis Step 1
                            LSet tmSpotTest = tlSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                            If tlSdf.lCode = tmSpotTest.lSdfCode Then
                                '3/14/13
                                If slSchStatus <> "F" Then
                                    ilAvTime0 = tmAvail.iTime(0)
                                    ilAvTime1 = tmAvail.iTime(1)
                                    ilAvInfo = tmAvail.iAvInfo
                                    'Move all event up, change count and update
                                    tmAvail.iNoSpotsThis = tmAvail.iNoSpotsThis - 1
                                    tlSsf.tPas(ADJSSFPASBZ + ilLoop) = tmAvail
                                    For ilMoveLoop = ilSpotIndex To tlSsf.iCount - 1 Step 1
                                        If ilMoveLoop = ilSpotIndex Then
                                            If (tmSpotTest.iRecType And SSSPLITPRI) = SSSPLITPRI Then
                                               LSet tlSpot = tlSsf.tPas(ADJSSFPASBZ + ilMoveLoop + 1)
                                                If (tlSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                                    tlSpot.iRecType = tlSpot.iRecType And (Not SSSPLITSEC)
                                                    tlSpot.iRecType = tlSpot.iRecType Or SSSPLITPRI
                                                    LSet tlSsf.tPas(ADJSSFPASBZ + ilMoveLoop + 1) = tlSpot
                                                End If
                                            End If
                                        End If
                                        tlSsf.tPas(ADJSSFPASBZ + ilMoveLoop) = tlSsf.tPas(ADJSSFPASBZ + ilMoveLoop + 1)
                                    Next ilMoveLoop
                                    'BB, bookends, donuts,... not hamdled
                                    tlSsf.iCount = tlSsf.iCount - 1
                                Else
                                    tmSpotTest.iRank = 1045
                                    LSet tlSsf.tPas(ADJSSFPASBZ + ilSpotIndex) = tmSpotTest
                                End If
                                ilFound = True
                            End If
                        Next ilSpotIndex
                        If ilFound Then
                            Exit Do
                        Else
                            ilLoop = ilLoop + 1
                        End If
                    Else
                        ilLoop = ilLoop + 1
                    End If
                Loop
            End If
            If ilFound Then
                imSsfRecLen = igSSFBaseLen + tlSsf.iCount * Len(tmAvail)
                ilRet = gSSFUpdate(hlSsf, tlSsf, imSsfRecLen)
            Else
                '11/24/12
                'If ilGameNo = 0 Then
                    imSsfRecLen = Len(tlSsf) 'Max size of variable length record
                    ilRet = gSSFGetNext(hlSsf, tlSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet = BTRV_ERR_NONE Then
                        If (tlSdf.iDate(0) <> tlSsf.iDate(0)) Or (tlSdf.iDate(1) <> tlSsf.iDate(1)) Or (tlSdf.iVefCode <> tlSsf.iVefCode) Or (tlSsf.iType <> ilGameNo) Then
                            ilRet = BTRV_ERR_NONE
                        Else
                            llSvSsfRecPos = llSsfRecPos
                            ilRet = gSSFGetPosition(hlSsf, llSsfRecPos)
                            If llSvSsfRecPos <> llSsfRecPos Then
                                ilRet = BTRV_ERR_CONFLICT
                            End If
                        End If
                    End If
                'End If
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_END_OF_FILE) Then
            igBtrError = gConvertErrorCode(ilRet)
            sgErrLoc = "gChgSchSpot-Update SSf(1)"
            gChgSchSpot = False
            Exit Function
        End If
        
        If (tlSdf.sWasMG = "Y") Or (tlSdf.sFromWorkArea = "Y") Or (slSchStatus = "M") Then
            tmSdfSrchKey3.lCode = tlSdf.lCode
            ilRet = btrGetEqual(hlSdf, tgSxfSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
        End If
        If (tgSxfSdf.sWasMG = "Y") Or (tgSxfSdf.sFromWorkArea = "Y") Then
            ilRet = gSxfDelete(hlSxf, tgSxfSdf)
            ilRet = btrUpdate(hlSdf, tgSxfSdf, imSdfRecLen)
        End If
        If (slSchStatus = "M") Then
            ilRet = gSxfAdd(hlSxf, "G", tgSxfSdf)
            ilRet = btrUpdate(hlSdf, tgSxfSdf, imSdfRecLen)
        End If
        
        tlSmf.lChfCode = 0  'Set as not found
        If ilFound Then
            '3/14/13
            If slSchStatus = "F" Then
                tmSdfSrchKey3.lCode = tlSdf.lCode
                ilRet = btrGetEqual(hlSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    igBtrError = gConvertErrorCode(ilRet)
                    sgErrLoc = "gChgSchSpot-GetEqual Sdf(14)"
                    gChgSchSpot = False
                    Exit Function
                End If
                tlSdf.sPriceType = slPriceType
                tlSdf.sSpotType = "X"
                '12/29/14: Change to fill to outside
                If (tlSdf.sSchStatus <> "G") And (tlSdf.sSchStatus <> "O") Then
                    tlSdf.sSchStatus = "O"
                    If tlSdf.lSmfCode <= 0 Then
                        tlSmf.lCode = 0
                        gUnpackDate tlSdf.iDate(0), tlSdf.iDate(1), slDate
                        gUnpackTime tlSdf.iTime(0), tlSdf.iTime(1), "A", "1", slTime
                        If tlSdf.iGameNo > 0 Then
                            ilRet = gMakeSmf(hlSmf, tlSmf, "O", tlSdf, ilFillSourceVefCode, slDate, slTime, 0, slDate, slTime)
                        Else
                            ilRet = gMakeSmf(hlSmf, tlSmf, "O", tlSdf, ilFillSourceVefCode, slDate, slTime, 1, slDate, slTime)
                        End If
                        tlSdf.lSmfCode = tlSmf.lCode
                    End If
                End If
                tlSdf.iUrfCode = tgUrf(0).iCode
                ilRet = btrUpdate(hlSdf, tlSdf, imSdfRecLen)
                If ilRet <> BTRV_ERR_NONE Then
                    igBtrError = gConvertErrorCode(ilRet)
                    sgErrLoc = "gChgSchSpot-GetEqual Sdf(15)"
                    gChgSchSpot = False
                    Exit Function
                End If
                'gMakeLogAlert tlSdf, "S", hlGsf
                gChgSchSpot = True
                Exit Function
            End If
            ilRet = gRemoveSmf(hlSmf, tlSmf, tlSdf, hlSxf)  'resets missed date and vehicle
            If Not ilRet Then
                gChgSchSpot = False
                Exit Function
            End If
            ilDate0 = tlSdf.iDate(0)
            ilDate1 = tlSdf.iDate(1)
            If (tlSdf.sSchStatus = "O") Or (tlSdf.sSchStatus = "G") Then
                If (tlSdf.iTime(0) = 1) And (tlSdf.iTime(1) = 0) Then
                    ilTime0 = ilAvTime0 ' tlSdf.iTime(0)
                    ilTime1 = ilAvTime1 'tlSdf.iTime(1)
                Else
                    ilTime0 = tlSdf.iTime(0)
                    ilTime1 = tlSdf.iTime(1)
                End If
            Else
                ilTime0 = ilAvTime0 'tlSdf.iTime(0)
                ilTime1 = ilAvTime1 'tlSdf.iTime(1)
            End If
            ilOrigSchVef = tlSdf.iVefCode
            ilOrigGameNo = tlSdf.iGameNo
            '8/2/11: Missed date within smf might be in error.
            '        Possible cause:  Spot sold to game x scheduled as MG in game y.  Game X date changed.  The missed date was NOT updated in GameSchd
            '        Obtain date from SSF instead of gsf because no handle for gsf
            If ilOrigGameNo > 0 Then
                'If tlSsf.iType <> ilOrigGameNo Then
                    '3/9/13: Code scanned and all places now open gsf and ghf so this case should not happen
                    If (hlGsf = 0) Or (hlGhf = 0) Then
                        '3/11/13: Use the spot date.
                        'imSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                        'tmSsfSrchKey1.iVefCode = ilOrigSchVef
                        'tmSsfSrchKey1.iType = ilOrigGameNo
                        'ilRet = gSSFGetEqualKey1(hlSsf, tmCTSsf, imSsfRecLen, tmSsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
                        'If (ilRet = BTRV_ERR_NONE) Then
                        '    ilDate0 = tmCTSsf.iDate(0)
                        '    ilDate1 = tmCTSsf.iDate(1)
                        'End If
                    Else
                        gUnpackDateLong ilDate0, ilDate1, llSDFDate
                        ilGhfRecLen = Len(tlGhf)
                        ilGsfRecLen = Len(tlGsf)
                        tlGsfSrchKey3.iVefCode = ilOrigSchVef
                        tlGsfSrchKey3.iGameNo = ilOrigGameNo
                        ilRet = btrGetEqual(hlGsf, tlGsf, ilGsfRecLen, tlGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                        Do While (ilRet = BTRV_ERR_NONE) And (tlGsf.iGameNo = ilOrigGameNo) And (tlGsf.iVefCode = ilOrigSchVef)
                            tlGhfSrchKey0.lCode = tlGsf.lghfcode
                            ilRet = btrGetEqual(hlGhf, tlGhf, ilGhfRecLen, tlGhfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                            If ilRet = BTRV_ERR_NONE Then
                                gUnpackDateLong tlGhf.iSeasonStartDate(0), tlGhf.iSeasonStartDate(1), llSeasonStart
                                gUnpackDateLong tlGhf.iSeasonEndDate(0), tlGhf.iSeasonEndDate(1), llSeasonEnd
                                If (llSDFDate >= llSeasonStart) And (llSDFDate <= llSeasonEnd) Then
                                    ilDate0 = tlGsf.iAirDate(0)
                                    ilDate1 = tlGsf.iAirDate(1)
                                    Exit Do
                                End If
                            End If
                            ilRet = btrGetNext(hlGsf, tlGsf, ilGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                'Else
                '    ilDate0 = tlSsf.iDate(0)
                '    ilDate1 = tlSsf.iDate(1)
                'End If
            End If
            If tlSdf.sSpotType = "X" Then
                slXSpotType = "X"
                If ((tlSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                    slXSpotType = ""
                End If
            Else
                slXSpotType = ""
            End If
            'If (slSchStatus = "D") Or (((tlSdf.sSpotType = "T") Or (tlSdf.sSpotType = "Q") Or (tlSdf.sSpotType = "S") Or (tlSdf.sSpotType = "M") Or (slXSpotType = "X")) And (slSchStatus <> "TM")) Then
            If (slSchStatus = "D") Or ((((tlSdf.sSpotType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tlSdf.sSpotType = "Q") Or ((tlSdf.sSpotType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tlSdf.sSpotType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (slXSpotType = "X")) And (slSchStatus <> "TM")) Then
                'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                Do
                    'ilCRet = btrGetDirect(hlSdf, tlSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    'If ilCRet <> BTRV_ERR_NONE Then
                    '    igBtrError = ilCRet
                    '    sgErrLoc = "gChgSchSpot-GetDirect Sdf(11)"
                    '    gChgSchSpot = False
                    '    Exit Function
                    'End If
                    'tmSRec = tlSdf
                    'ilCRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                    'tlSdf = tmSRec
                    'If ilCRet <> BTRV_ERR_NONE Then
                    '    igBtrError = ilCRet
                    '    sgErrLoc = "gChgSchSpot-GetByKey Sdf(3)"
                    '    gChgSchSpot = False
                    '    Exit Function
                    'End If
                    tmSdfSrchKey3.lCode = tlSdf.lCode
                    ilRet = btrGetEqual(hlSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        igBtrError = gConvertErrorCode(ilRet)
                        sgErrLoc = "gChgSchSpot-GetEqual Sdf(3)"
                        gChgSchSpot = False
                        Exit Function
                    End If
                    ilRet = btrDelete(hlSdf)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    igBtrError = gConvertErrorCode(ilRet)
                    sgErrLoc = "gChgSchSpot-Delete Sdf(4)"
                    gChgSchSpot = False
                    Exit Function
                End If
           Else
                'tmSdfSrchKey3.lCode = tlSdf.lCode
                'ilRet = btrGetEqual(hlSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                'If ilRet <> BTRV_ERR_NONE Then
                '    igBtrError = gConvertErrorCode(ilRet)
                '    sgErrLoc = "gChgSchSpot-GetEqual Sdf(5)"
                '    gChgSchSpot = False
                '    Exit Function
                'End If
                'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                'If ilRet <> BTRV_ERR_NONE Then
                '    igBtrError = gConvertErrorCode(ilRet)
                '    sgErrLoc = "gChgSchSpot-GetPosition Sdf(6)"
                '    gChgSchSpot = False
                '    Exit Function
                'End If
                Do
                    'ilRet = btrGetDirect(hlSdf, tlSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    igBtrError = gConvertErrorCode(ilRet)
                    '    sgErrLoc = "gChgSchSpot-GetDirect Sdf(12)"
                    '    gChgSchSpot = False
                    '    Exit Function
                    'End If
                    'tmSRec = tlSdf
                    'ilRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                    'tlSdf = tmSRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    igBtrError = gConvertErrorCode(ilRet)
                    '    sgErrLoc = "gChgSchSpot-GetByKey Sdf(7)"
                    '    gChgSchSpot = False
                    '    Exit Function
                    'End If
                    tmSdfSrchKey3.lCode = tlSdf.lCode
                    ilRet = btrGetEqual(hlSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        igBtrError = gConvertErrorCode(ilRet)
                        sgErrLoc = "gChgSchSpot-GetEqual Sdf(7)"
                        gChgSchSpot = False
                        Exit Function
                    End If
                    If slSchStatus = "TM" Then
                        tlSdf.sSchStatus = "M"
                    Else
                        tlSdf.sSchStatus = slSchStatus
                    End If
                    tlSdf.iVefCode = ilOrigSchVef
                    tlSdf.iMnfMissed = igMnfMissed
                    tlSdf.iDate(0) = ilDate0
                    tlSdf.iDate(1) = ilDate1
                    tlSdf.iTime(0) = ilTime0
                    tlSdf.iTime(1) = ilTime1
                    tlSdf.iGameNo = ilOrigGameNo
                    If ((tlSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                        tlSdf.sTracer = "*"
                        tlSdf.lSmfCode = tlSmf.lMtfCode
                    ElseIf tlSdf.sSchStatus = "M" Then
                        tlSdf.lSmfCode = 0
                    End If
                    tlSdf.sXCrossMidnight = "N"
                    If ((slSchStatus = "S") Or (slSchStatus = "G") Or (slSchStatus = "O")) And ((ilAvInfo And SSXMID) = SSXMID) Then
                        tlSdf.sXCrossMidnight = "Y"
                    End If
                    tlSdf.iUrfCode = tgUrf(0).iCode
                    ilRet = btrUpdate(hlSdf, tlSdf, imSdfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    igBtrError = gConvertErrorCode(ilRet)
                    sgErrLoc = "gChgSchSpot-Update Sdf(8)"
                    gChgSchSpot = False
                    Exit Function
                End If
            End If
        End If
    Else
        If (slSchStatus <> "D") Or (tlSdf.sSchStatus <> "M") Then
            igBtrError = -1
            sgErrLoc = "gChgSchSpot-gObtainSsfForDateOrGame SSf(3)"
            gChgSchSpot = False
            Exit Function
        End If
    End If
    If ilFound Then
        gMakeLogAlert tlSdf, "S", hlGsf
        gChgSchSpot = True
    Else
        'Spot only in Sdf- missing from Ssf
        '3/14/13
        If slSchStatus = "F" Then
            tmSdfSrchKey3.lCode = tlSdf.lCode
            ilRet = btrGetEqual(hlSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = gConvertErrorCode(ilRet)
                sgErrLoc = "gChgSchSpot-GetEqual Sdf(16)"
                gChgSchSpot = False
                Exit Function
            End If
            tlSdf.sPriceType = slPriceType
            tlSdf.sSpotType = "X"
            '12/29/14: Change to fill to outside
            If (tlSdf.sSchStatus <> "G") And (tlSdf.sSchStatus <> "O") Then
                tlSdf.sSchStatus = "O"
                tlSmf.lCode = 0
                gUnpackDate tlSdf.iDate(0), tlSdf.iDate(1), slDate
                gUnpackTime tlSdf.iTime(0), tlSdf.iTime(1), "A", "1", slTime
                If tlSdf.iGameNo > 0 Then
                    ilRet = gMakeSmf(hlSmf, tlSmf, "O", tlSdf, ilFillSourceVefCode, slDate, slTime, 0, slDate, slTime)
                Else
                    'Arbitarily pick game 1 as source
                    ilRet = gMakeSmf(hlSmf, tlSmf, "O", tlSdf, ilFillSourceVefCode, slDate, slTime, 1, slDate, slTime)
                End If
                tlSdf.lSmfCode = tlSmf.lCode
            End If
            tlSdf.iUrfCode = tgUrf(0).iCode
            ilRet = btrUpdate(hlSdf, tlSdf, imSdfRecLen)
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = gConvertErrorCode(ilRet)
                sgErrLoc = "gChgSchSpot-GetEqual Sdf(17)"
                gChgSchSpot = False
                Exit Function
            End If
            'gMakeLogAlert tlSdf, "S", hlGsf
            gChgSchSpot = True
            Exit Function
        End If
        If tlSdf.sSpotType = "X" Then
            slXSpotType = "X"
            If ((tlSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                slXSpotType = ""
            End If
        Else
            slXSpotType = ""
        End If
        'If (slSchStatus = "D") Or (((tlSdf.sSpotType = "T") Or (tlSdf.sSpotType = "Q") Or (tlSdf.sSpotType = "S") Or (tlSdf.sSpotType = "M") Or (slXSpotType = "X")) And (slSchStatus <> "TM")) Then
        If (slSchStatus = "D") Or ((((tlSdf.sSpotType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tlSdf.sSpotType = "Q") Or ((tlSdf.sSpotType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tlSdf.sSpotType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (slXSpotType = "X")) And (slSchStatus <> "TM")) Then
            ilRet = gRemoveSmf(hlSmf, tlSmf, tlSdf, hlSxf)  'resets missed date
            If Not ilRet Then
                gChgSchSpot = False
                Exit Function
            End If
            'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
            Do
                'ilCRet = btrGetDirect(hlSdf, tlSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                'If ilCRet <> BTRV_ERR_NONE Then
                '    igBtrError = ilCRet
                '    sgErrLoc = "gChgSchSpot-GetDirect Sdf(13)"
                '    gChgSchSpot = False
                '    Exit Function
                'End If
                'tmSRec = tlSdf
                'ilCRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                'tlSdf = tmSRec
                'If ilCRet <> BTRV_ERR_NONE Then
                '    igBtrError = ilCRet
                '    sgErrLoc = "gChgSchSpot-GetByKey Sdf(9)"
                '    gChgSchSpot = False
                '    Exit Function
                'End If
                tmSdfSrchKey3.lCode = tlSdf.lCode
                ilRet = btrGetEqual(hlSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                If ilRet <> BTRV_ERR_NONE Then
                    igBtrError = gConvertErrorCode(ilRet)
                    sgErrLoc = "gChgSchSpot-GetEqual Sdf(9)"
                    gChgSchSpot = False
                    Exit Function
                End If
                ilRet = btrDelete(hlSdf)
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = gConvertErrorCode(ilRet)
                sgErrLoc = "gChgSchSpot-Delete Sdf(10)"
                gChgSchSpot = False
                Exit Function
            End If
            gMakeLogAlert tlSdf, "S", hlGsf
            gChgSchSpot = True
        Else
            gChgSchSpot = False
        End If
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gCompetitiveTest                *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test for competitive conflicts  *
'*                                                     *
'*******************************************************
Function gCompetitiveTest(llCompTime As Long, hlSsf As Integer, tlSsf As SSF, llSsfRecPos As Long, tlSpotMove() As SPOTMOVE, ilVpfIndex As Integer, ilLen As Integer, ilMnfComp0 As Integer, ilMnfComp1 As Integer, ilAvailIndex As Integer, tlVcf0() As VCF, tlVcf6() As VCF, tlVcf7() As VCF, ilSchMode As Integer, ilBkQH As Integer, slInOut As String, slPreempt As String, ilPriceLevel As Integer, ilCheckAvail) As Integer
'
'   ilRet = gCompetitiveTest (ilCompTime, hlSsf, tlSsf, llSsfRecPos, tlSpotMove(), ilVpfIndex, ilLen, ilMnfComp0, ilMnfComp1, ilAvailIndex, tlVcf0(), tlVcf6(), tlVcf7())
'   Where:
'       llCompTime(I)- competitive separation time (obtained from the vehicle option file (vpf))
'       tlSsf(I)- Ssf image
'       llSsfRecPos(I)- Ssf record position
'       tlSpotMove()(I)- Array of spots to bypass
'       ilLen(I) Spot length (required if scheduling by thirty units and back to back competitive test
'       ilMnfComp0(I)- Contract header Competitive code # 1
'       ilMnfComp1(I)- Contract header Comptitive code # 2
'       ilAvailIndex(I)- Index into tgSsf for avail to be processed
'       ilSchMode(I)- 0=Insert; 1=Move; 2=Compact; 3=Preempt; 4=Preempt Fill only (call only from SpotMG)
'       ilBkQH(I)- Rank for ilSchMode = 3
'       ilCheckAvail(I)- True; False (bypass checking avail because this is split network)
'       ilRet(O)- True = No conflicts; False= Conflicts
'
'       tmSdf contains the spot record to be checked
'       tgSsf contains the days events
'       tmAvail contain the avail to be check
'
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilSpotIndex As Integer
    Dim slDate As String
    Dim slTime As String
    Dim llStartAvailTime As Long
    Dim llEndAvailTime As Long
    Dim llAvailTime As Long
    Dim ilNoCompSpots As Integer
    Dim ilLenSold As Integer
    Dim ilUnits As Integer
    Dim ilBypass As Integer
    Dim ilBypassIndex As Integer
    Dim ilDayIndex As Integer
    Dim ilVcfDefined As Integer
    Dim ilRet As Integer
    Dim ilInitPreempt As Integer
    Dim ilCompIndex As Integer
    Dim ilPass As Integer
    Dim ilBNoPasses As Integer
    Dim llPass2StartAvailTime As Long
    Dim ilANoPasses As Integer
    Dim llPass2EndAvailTime As Long
    Dim llDate As Long
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilLBSpotMove As Integer
    
    If (ilMnfComp0 = 0) And (ilMnfComp1 = 0) Then
        gCompetitiveTest = True
        Exit Function
    End If
    ilLBSpotMove = LBound(tlSpotMove)
    ilInitPreempt = UBound(tlSpotMove)
    gUnpackDate tlSsf.iDate(0), tlSsf.iDate(1), slDate
    ilDayIndex = gWeekDayStr(slDate)
    If ilDayIndex = 5 Then
        'If UBound(tlVcf6) > 1 Then
        If UBound(tlVcf6) > 0 Then
            ilVcfDefined = True
        End If
    ElseIf ilDayIndex = 6 Then
        'If UBound(tlVcf7) > 1 Then
        If UBound(tlVcf7) > 0 Then
            ilVcfDefined = True
        End If
    Else
        'If UBound(tlVcf0) > 1 Then
        If UBound(tlVcf0) > 0 Then
            ilVcfDefined = True
        End If
    End If
    If ilAvailIndex + tlSsf.tPas(ADJSSFPASBZ + ilAvailIndex).iNoSpotsThis > UBound(tlSsf.tPas) Then
        gCompetitiveTest = False
        Exit Function
    End If
   LSet tmAvail = tlSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
    If ilVcfDefined Then
        'Test avail being considered
        For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tmAvail.iNoSpotsThis Step 1
            If (ilSpotIndex < 1) Or (ilSpotIndex > tlSsf.iCount) Then
                gCompetitiveTest = False
                Exit Function
            End If
            LSet tmSpotTest = tlSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
            ilBypass = False
            For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                    ilBypass = True
                    Exit For
                End If
            Next ilBypassIndex
            If Not ilBypass Then
                If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp0 = tmSpotTest.iMnfComp(1))) Then
                    'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvail.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                    If gPreemptible(ilSchMode, tlSpotMove(), tmAvail, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                        If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                            ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                            tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                            tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                            tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                        Else
                            ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                            tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                            tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                            tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                        End If
                    Else
                        If ilInitPreempt < UBound(tlSpotMove) Then
                            ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                        End If
                        gCompetitiveTest = False
                        Exit Function
                    End If
                End If
                'Recheck bypass to avoid check smae spot twice (added above)
                ilBypass = False
                For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                    If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                        ilBypass = True
                        Exit For
                    End If
                Next ilBypassIndex
                If Not ilBypass Then
                    If (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp1 = tmSpotTest.iMnfComp(1))) Then
                        'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvail.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                        If gPreemptible(ilSchMode, tlSpotMove(), tmAvail, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                            If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                                ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                                tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                                tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                                tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                            Else
                                ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                                tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                                tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                                tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                            End If
                        Else
                            If ilInitPreempt < UBound(tlSpotMove) Then
                                ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                            End If
                            gCompetitiveTest = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next ilSpotIndex
        'Check conflicts on other vehicle
        'slDate = Format$(lmMondayDate + imDayIndex, "m/d/yy")
        'gPackDate slDate, ilDate0, ilDate1
        Select Case ilDayIndex
            Case 0 To 4
                For ilLoop = LBound(tlVcf0) To UBound(tlVcf0) - 1 Step 1
                    If (tmAvail.iTime(0) = tlVcf0(ilLoop).iSellTime(0)) And (tmAvail.iTime(1) = tlVcf0(ilLoop).iSellTime(1)) Then
                        'Check for conflict on ther vehicles
                        ilRet = mVehCompConflictTest(hlSsf, tlSsf, tlSpotMove(), ilMnfComp0, ilMnfComp1, tlVcf0(ilLoop))
                        If ilRet = False Then
                            If ilInitPreempt < UBound(tlSpotMove) Then
                                ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                            End If
                        End If
                        gCompetitiveTest = ilRet
                        Exit Function
                    End If
                Next ilLoop
                gCompetitiveTest = True
            Case 5
                For ilLoop = LBound(tlVcf6) To UBound(tlVcf6) - 1 Step 1
                    If (tmAvail.iTime(0) = tlVcf6(ilLoop).iSellTime(0)) And (tmAvail.iTime(1) = tlVcf6(ilLoop).iSellTime(1)) Then
                        'Check for conflict on ther vehicles
                        ilRet = mVehCompConflictTest(hlSsf, tlSsf, tlSpotMove(), ilMnfComp0, ilMnfComp1, tlVcf6(ilLoop))
                        If ilRet = False Then
                            If ilInitPreempt < UBound(tlSpotMove) Then
                                ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                            End If
                        End If
                        gCompetitiveTest = ilRet
                        Exit Function
                    End If
                Next ilLoop
                gCompetitiveTest = True
            Case 6
                For ilLoop = LBound(tlVcf7) To UBound(tlVcf7) - 1 Step 1
                    If (tmAvail.iTime(0) = tlVcf7(ilLoop).iSellTime(0)) And (tmAvail.iTime(1) = tlVcf7(ilLoop).iSellTime(1)) Then
                        'Check for conflict on ther vehicles
                        ilRet = mVehCompConflictTest(hlSsf, tlSsf, tlSpotMove(), ilMnfComp0, ilMnfComp1, tlVcf7(ilLoop))
                        If ilRet = False Then
                            If ilInitPreempt < UBound(tlSpotMove) Then
                                ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                            End If
                        End If
                        gCompetitiveTest = ilRet
                        Exit Function
                    End If
                Next ilLoop
                gCompetitiveTest = True
        End Select
    Else
        If (tgVpf(ilVpfIndex).sSCompType = "T") And (llCompTime <= 0) Then
            gCompetitiveTest = True
            Exit Function
        End If
        'Test within current avail (if Back to Back- then two spots of
        'same competitives allowed if three spots or room for a third spot)
        If (tgVpf(ilVpfIndex).sSCompType = "N") Then    'N="Not Back to Back"
            ilNoCompSpots = 0
            '12/7/09:  Remove the CheckAvail test as it was not coded correctly.  It assumed that all split networks were not overlapping
            '          i.e. Split 1:  West Coast; Split 2: East coast.  The two could be scheduled within the same avail
            '          It fails if Split 1 was West + East or Split 1 was exclude West.
           'If ilCheckAvail Then
                For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tmAvail.iNoSpotsThis Step 1
                    LSet tmSpotTest = tlSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                    ilBypass = False
                    For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                        If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                            ilBypass = True
                            Exit For
                        End If
                    Next ilBypassIndex
                    If Not ilBypass Then
                        If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp0 = tmSpotTest.iMnfComp(1))) Then
                            ilNoCompSpots = ilNoCompSpots + 1
                            ilCompIndex = ilSpotIndex
                        ElseIf (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp1 = tmSpotTest.iMnfComp(1))) Then
                            ilNoCompSpots = ilNoCompSpots + 1
                            ilCompIndex = ilSpotIndex
                        End If
                    End If
                Next ilSpotIndex
            'End If
            If ilNoCompSpots = 0 Then
                gCompetitiveTest = True
                Exit Function
            ElseIf ilNoCompSpots > 1 Then
                gCompetitiveTest = False
                Exit Function
            Else    'only one competitive spot
                ilSpotIndex = ilCompIndex
                LSet tmSpotTest = tlSsf.tPas(ADJSSFPASBZ + ilSpotIndex)   'Reset for preempt test
                If tmAvail.iNoSpotsThis > 1 Then    'One competitve and two spot booked already- spots can be book not back to back
                    gCompetitiveTest = True
                    Exit Function
                Else
                    'If there room for another spot
                    If (ilVpfIndex >= 0) And (tgVpf(ilVpfIndex).sSSellOut = "T") Then
                        ilUnits = ilLen \ 30
                        If ilUnits <= 0 Then
                            ilUnits = 1
                        End If
                    Else
                        ilUnits = 1
                    End If
                    If ilUnits + 1 > (tmAvail.iAvInfo And &H1F) - tmAvail.iNoSpotsThis Then
                        'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvail.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                        If gPreemptible(ilSchMode, tlSpotMove(), tmAvail, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                            If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                                ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                                tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                                tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                                tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                            Else
                                ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                                tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                                tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                                tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                            End If
                            gCompetitiveTest = True
                            Exit Function
                        Else
                            If ilInitPreempt < UBound(tlSpotMove) Then
                                ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                            End If
                            gCompetitiveTest = False
                            Exit Function
                        End If
                    Else    'Test time
                        If (tgVpf(ilVpfIndex).sSSellOut = "B") Or (tgVpf(ilVpfIndex).sSSellOut = "U") Then
                            ilLenSold = 0
                            For ilIndex = ilLoop + 1 To ilLoop + tmAvail.iNoSpotsThis Step 1
                                LSet tmSpotTest = tlSsf.tPas(ADJSSFPASBZ + ilIndex)
                                ilBypass = False
                                For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                                    If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                                        ilBypass = True
                                        Exit For
                                    End If
                                Next ilBypassIndex
                                If (tmSpotTest.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                    ilBypass = True
                                End If
                                If Not ilBypass Then
                                    ilLenSold = ilLenSold + (tmSpotTest.iPosLen And &HFFF)
                                End If
                            Next ilIndex
                            If tmAvail.iLen - ilLenSold - ilLen <= 0 Then
                                'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvail.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                                If gPreemptible(ilSchMode, tlSpotMove(), tmAvail, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                                    If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                                        ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                                        tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                                        tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                                        tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                                    Else
                                        ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                                        tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                                        tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                                        tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                                    End If
                                    gCompetitiveTest = True
                                    Exit Function
                                Else
                                    If ilInitPreempt < UBound(tlSpotMove) Then
                                        ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                    End If
                                    gCompetitiveTest = False
                                    Exit Function
                                End If
                            Else
                                gCompetitiveTest = True
                                Exit Function
                            End If
                        ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                            ilLenSold = 0
                            For ilIndex = ilLoop + 1 To ilLoop + tmAvail.iNoSpotsThis Step 1
                                LSet tmSpotTest = tlSsf.tPas(ADJSSFPASBZ + ilIndex)
                                ilBypass = False
                                For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                                    If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                                        ilBypass = True
                                        Exit For
                                    End If
                                Next ilBypassIndex
                                If (tmSpotTest.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                    ilBypass = True
                                End If
                                If Not ilBypass Then
                                    ilLenSold = ilLenSold + (tmSpotTest.iPosLen And &HFFF)
                                End If
                            Next ilIndex
                            If (tmAvail.iLen - ilLenSold) <> ilLen Then
                                gCompetitiveTest = False
                                Exit Function
                            Else
                                gCompetitiveTest = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        Else    'Not in same break or by time
            'Test avail being considered
            '12/7/09:  Remove the CheckAvail test as it was not coded correctly.  It assumed that all split networks were not overlapping
            '          i.e. Split 1:  West Coast; Split 2: East coast.  The two could be scheduled within the same avail
            '          It fails if Split 1 was West + East or Split 1 was exclude West.
            'If ilCheckAvail Then
                For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tmAvail.iNoSpotsThis Step 1
                    LSet tmSpotTest = tlSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                    ilBypass = False
                    For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                        If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                            ilBypass = True
                            Exit For
                        End If
                    Next ilBypassIndex
                    If Not ilBypass Then
                        If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp0 = tmSpotTest.iMnfComp(1))) Then
                            'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvail.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                            If ilPass = 2 Then
                                If ilInitPreempt < UBound(tlSpotMove) Then
                                    ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                End If
                                gCompetitiveTest = False
                                Exit Function
                            End If
                            If gPreemptible(ilSchMode, tlSpotMove(), tmAvail, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                                If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                                    ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                                    tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                                    tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                                    tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                                Else
                                    ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                                    tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                                    tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                                    tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                                End If
                            Else
                                If ilInitPreempt < UBound(tlSpotMove) Then
                                    ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                End If
                                gCompetitiveTest = False
                                Exit Function
                            End If
                        End If
                        'Recheck bypass to avoid check smae spot twice (added above)
                        ilBypass = False
                        For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                            If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                                ilBypass = True
                                Exit For
                            End If
                        Next ilBypassIndex
                        If Not ilBypass Then
                            If (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp1 = tmSpotTest.iMnfComp(1))) Then
                                'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvail.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                                If ilPass = 2 Then
                                    If ilInitPreempt < UBound(tlSpotMove) Then
                                        ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                    End If
                                    gCompetitiveTest = False
                                    Exit Function
                                End If
                                If gPreemptible(ilSchMode, tlSpotMove(), tmAvail, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                                    If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                                        ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                                        tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                                        tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                                        tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                                    Else
                                        ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                                        tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                                        tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                                        tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                                    End If
                                Else
                                    If ilInitPreempt < UBound(tlSpotMove) Then
                                        ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                    End If
                                    gCompetitiveTest = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next ilSpotIndex
            'End If
        End If
        If llCompTime <= 0 Then
            gCompetitiveTest = True
            Exit Function
        End If
        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
        llAvailTime = CLng(gTimeToCurrency(slTime, False))
        llStartAvailTime = llAvailTime - llCompTime
        ilBNoPasses = 1
        If llStartAvailTime < 0 Then
            If tlSsf.iType = 0 Then
                ilBNoPasses = 2
                llPass2StartAvailTime = 86400 + llStartAvailTime
            End If
            llStartAvailTime = 0
        End If
        llEndAvailTime = llAvailTime + llCompTime
        ilANoPasses = 1
        If llEndAvailTime > 86400 Then
            If tlSsf.iType = 0 Then
                ilANoPasses = 2
                llPass2EndAvailTime = llEndAvailTime - 86400
            End If
            llEndAvailTime = 86400
        End If
        'Test avails prior to avail being considered
        ilIndex = ilAvailIndex - 1
        tmCTSsf = tlSsf
        For ilPass = 1 To ilBNoPasses Step 1
            Do
                If ilIndex < 1 Then
                    Exit Do
                End If
                tmAvailTest = tmCTSsf.tPas(ADJSSFPASBZ + ilIndex)
                If (tmAvailTest.iRecType >= 2) And (tmAvailTest.iRecType <= 9) Then
                    gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                    llAvailTime = CLng(gTimeToCurrency(slTime, False))
                    If llAvailTime < llStartAvailTime Then
                        Exit Do
                    End If
                    For ilSpotIndex = ilIndex + 1 To ilIndex + tmAvailTest.iNoSpotsThis Step 1
                        LSet tmSpotTest = tmCTSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                        ilBypass = False
                        For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                            If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                                ilBypass = True
                                Exit For
                            End If
                        Next ilBypassIndex
                        If Not ilBypass Then
                            If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp0 = tmSpotTest.iMnfComp(1))) Then
                                'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvailTest.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                                If ilPass = 2 Then
                                    If ilInitPreempt < UBound(tlSpotMove) Then
                                        ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                    End If
                                    gCompetitiveTest = False
                                    Exit Function
                                End If
                                If gPreemptible(ilSchMode, tlSpotMove(), tmAvailTest, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                                    If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                                        ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                                        tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                                        tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                                        tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                                    Else
                                        ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                                        tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                                        tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                                        tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                                    End If
                                Else
                                    If ilInitPreempt < UBound(tlSpotMove) Then
                                        ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                    End If
                                    gCompetitiveTest = False
                                    Exit Function
                                End If
                            End If
                            'Recheck bypass to avoid check smae spot twice (added above)
                            ilBypass = False
                            For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                                If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                                    ilBypass = True
                                    Exit For
                                End If
                            Next ilBypassIndex
                            If Not ilBypass Then
                                If (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp1 = tmSpotTest.iMnfComp(1))) Then
                                    'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvailTest.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                                    If ilPass = 2 Then
                                        If ilInitPreempt < UBound(tlSpotMove) Then
                                            ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                        End If
                                        gCompetitiveTest = False
                                        Exit Function
                                    End If
                                    If gPreemptible(ilSchMode, tlSpotMove(), tmAvailTest, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                                        If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                                            ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                                            tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                                            tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                                            tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                                        Else
                                            ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                                            tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                                            tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                                            tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                                        End If
                                    Else
                                        If ilInitPreempt < UBound(tlSpotMove) Then
                                            ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                        End If
                                        gCompetitiveTest = False
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Next ilSpotIndex
                ElseIf tmAvailTest.iRecType = 1 Then
                   LSet tmProg = tmCTSsf.tPas(ADJSSFPASBZ + ilIndex)
                    gUnpackTime tmProg.iStartTime(0), tmProg.iStartTime(1), "A", "1", slTime
                    llAvailTime = CLng(gTimeToCurrency(slTime, False))
                Else
                    llAvailTime = llStartAvailTime + 1
                End If
                ilIndex = ilIndex - 1
            Loop While llAvailTime > llStartAvailTime
            If ilPass = ilBNoPasses Then
                Exit For
            End If
            gUnpackDateLong tlSsf.iDate(0), tlSsf.iDate(1), llDate
            gPackDateLong llDate - 1, ilDate0, ilDate1
            imSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
            tmSsfSrchKey.iType = 0 'slType
            tmSsfSrchKey.iVefCode = tlSsf.iVefCode
            tmSsfSrchKey.iDate(0) = ilDate0
            tmSsfSrchKey.iDate(1) = ilDate1
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetGreaterOrEqual(hlSsf, tmCTSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet <> BTRV_ERR_NONE) Or (tmCTSsf.iType <> 0) Or (tmCTSsf.iVefCode <> tlSsf.iVefCode) Or (tmCTSsf.iDate(0) <> ilDate0) Or (tmCTSsf.iDate(1) <> ilDate1) Then
                Exit For
            End If
            ilIndex = tmCTSsf.iCount
            llStartAvailTime = llPass2StartAvailTime - 1
        Next ilPass
        'Test avails after avail being considered
        ilIndex = ilAvailIndex + tmAvail.iNoSpotsThis + 1
        tmCTSsf = tlSsf
        For ilPass = 1 To ilANoPasses Step 1
            Do
                If ilIndex > tmCTSsf.iCount Then
                    Exit Do
                End If
                tmAvailTest = tmCTSsf.tPas(ADJSSFPASBZ + ilIndex)
                If (tmAvailTest.iRecType >= 2) And (tmAvailTest.iRecType <= 9) Then
                    gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slTime
                    llAvailTime = CLng(gTimeToCurrency(slTime, False))
                    If llAvailTime > llEndAvailTime Then
                        Exit Do
                    End If
                    For ilSpotIndex = ilIndex + 1 To ilIndex + tmAvailTest.iNoSpotsThis Step 1
                        LSet tmSpotTest = tmCTSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                        ilBypass = False
                        For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                            If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                                ilBypass = True
                                Exit For
                            End If
                        Next ilBypassIndex
                        If Not ilBypass Then
                            If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp0 = tmSpotTest.iMnfComp(1))) Then
                                'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvailTest.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                                If gPreemptible(ilSchMode, tlSpotMove(), tmAvailTest, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                                    If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                                        ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                                        tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                                        tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                                        tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                                    Else
                                        ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                                        tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                                        tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                                        tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                                    End If
                                Else
                                    If ilInitPreempt < UBound(tlSpotMove) Then
                                        ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                    End If
                                    gCompetitiveTest = False
                                    Exit Function
                                End If
                            End If
                            'Recheck bypass to avoid check smae spot twice (added above)
                            ilBypass = False
                            For ilBypassIndex = ilLBSpotMove To UBound(tlSpotMove) - 1 Step 1
                                If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                                    ilBypass = True
                                    Exit For
                                End If
                            Next ilBypassIndex
                            If Not ilBypass Then
                                If (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp1 = tmSpotTest.iMnfComp(1))) Then
                                    'If (ilSchMode = 3) And (UBound(tlSpotMove) = 1) And (tmSpotTest.iRank > ilBkQH) And (tmSpotTest.iRank <> RESERVATION) And ((tmSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) And ((tmAvailTest.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
                                    If gPreemptible(ilSchMode, tlSpotMove(), tmAvailTest, tmSpotTest, ilBkQH, slInOut, slPreempt, ilPriceLevel) Then
                                        If (ilSchMode <> 4) Or (UBound(tlSpotMove) = ilLBSpotMove) Then
                                            ReDim tlSpotMove(ilLBSpotMove To ilLBSpotMove + 1) As SPOTMOVE
                                            tlSpotMove(ilLBSpotMove).iSpotIndex = ilSpotIndex
                                            tlSpotMove(ilLBSpotMove).lSpotSsfRecPos = llSsfRecPos
                                            tlSpotMove(ilLBSpotMove).lSdfCode = tmSpotTest.lSdfCode
                                        Else
                                            ReDim Preserve tlSpotMove(ilLBSpotMove To ilLBSpotMove + 2) As SPOTMOVE
                                            tlSpotMove(ilLBSpotMove + 1).iSpotIndex = ilSpotIndex
                                            tlSpotMove(ilLBSpotMove + 1).lSpotSsfRecPos = llSsfRecPos
                                            tlSpotMove(ilLBSpotMove + 1).lSdfCode = tmSpotTest.lSdfCode
                                        End If
                                    Else
                                        If ilInitPreempt < UBound(tlSpotMove) Then
                                            ReDim Preserve tlSpotMove(ilLBSpotMove To ilInitPreempt) As SPOTMOVE
                                        End If
                                        gCompetitiveTest = False
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Next ilSpotIndex
                    ilIndex = ilIndex + tmAvailTest.iNoSpotsThis + 1
                ElseIf tmAvailTest.iRecType = 1 Then
                   LSet tmProg = tmCTSsf.tPas(ADJSSFPASBZ + ilIndex)
                    gUnpackTime tmProg.iStartTime(0), tmProg.iStartTime(1), "A", "1", slTime
                    llAvailTime = CLng(gTimeToCurrency(slTime, False))
                    ilIndex = ilIndex + 1
                Else
                    llAvailTime = llEndAvailTime - 1
                    ilIndex = ilIndex + 1
                End If
            Loop While llAvailTime < llEndAvailTime
            If ilPass = ilANoPasses Then
                Exit For
            End If
            gUnpackDateLong tlSsf.iDate(0), tlSsf.iDate(1), llDate
            gPackDateLong llDate + 1, ilDate0, ilDate1
            imSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
            tmSsfSrchKey.iType = 0 'slType
            tmSsfSrchKey.iVefCode = tlSsf.iVefCode
            tmSsfSrchKey.iDate(0) = ilDate0
            tmSsfSrchKey.iDate(1) = ilDate1
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetGreaterOrEqual(hlSsf, tmCTSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet <> BTRV_ERR_NONE) Or (tmCTSsf.iType <> 0) Or (tmCTSsf.iVefCode <> tlSsf.iVefCode) Or (tmCTSsf.iDate(0) <> ilDate0) Or (tmCTSsf.iDate(1) <> ilDate1) Then
                Exit For
            End If
            ilIndex = 1
            llEndAvailTime = llPass2EndAvailTime
        Next ilPass
    End If
    gCompetitiveTest = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gExtendTFN                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Move TFN library to current or  *
'*                     pending log dates.  If Current, *
'*                     also create Spot summary (Ssf)  *
'*                                                     *
'*******************************************************
Function gExtendTFN(hlLcf As Integer, hlSsf As Integer, hlSdf As Integer, hlSmf As Integer, sLCP As String, ilVefCode As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, ilClearResch As Integer) As Integer
'
'   ilRet = gExtendTFN(hlLcf, hlSsf, hlSdf, hlSmf, slCP, ilVefCode, ilLogDate(0), ilLogDate(1), ilClearResch)
'   Where:
'       hlLcf (I) - Library calendar handle
'       hlSsf (I) - Spot summary handle (required if slCP = "C")
'       hlSdf (I) - Spot detail handle (required if slCP = "C")
'       hlSmf (I) - Spot MG handle (required if slCP = "C")
'       slCP (I) - C=Extend current; P= Extend pending
'       ilVefCode (I) - Vehicle code
'       ilLogDate (I) - Date to move TFN into
'       ilClearResch(I)- Reset the lgReschSdfCode Array
'
'       ilRet = True if OK; False if Error
'
    Dim tlLcf As LCF                'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim slDate As String
    Dim ilSeqNo As Integer
    Dim ilRet As Integer
    Dim ilVef As Integer
    If ilClearResch Then
        'ReDim lgReschSdfCode(1 To 1) As Long
        ReDim lgReschSdfCode(0 To 0) As Long
    End If
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).sType = "G" Then
            'Game creation of dates done in Programming
            gExtendTFN = True
            Exit Function
        End If
    End If
    ilLcfRecLen = Len(tlLcf)
    tlLcfSrchKey.iType = 0
    tlLcfSrchKey.sStatus = sLCP 'P or C
    tlLcfSrchKey.iVefCode = ilVefCode
    tlLcfSrchKey.iLogDate(0) = ilLogDate0
    tlLcfSrchKey.iLogDate(1) = ilLogDate1
    tlLcfSrchKey.iSeqNo = 1
    ilRet = btrGetEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet = BTRV_ERR_NONE Then
        'Day image exist- don't extend
        gExtendTFN = True
        Exit Function
    End If
    ilSeqNo = 1
    Do
        Do  'Loop until record updated or added
            gUnpackDate ilLogDate0, ilLogDate1, slDate
            tlLcfSrchKey.iType = 0
            tlLcfSrchKey.sStatus = sLCP 'P or C
            tlLcfSrchKey.iVefCode = ilVefCode
            tlLcfSrchKey.iLogDate(0) = gWeekDayStr(slDate) + 1 '1=Monday TFN; 2=Tues,...7=Sunday TFN
            tlLcfSrchKey.iLogDate(1) = 0
            tlLcfSrchKey.iSeqNo = ilSeqNo
            ilRet = btrGetEqual(hlLcf, tlLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                'Not TFN to extend- exit
                gExtendTFN = True
                Exit Function
            End If
            tlLcf.iLogDate(0) = ilLogDate0
            tlLcf.iLogDate(1) = ilLogDate1
            tlLcf.iUrfCode = tgUrf(0).iCode
            tlLcf.lCode = 0
            ilRet = btrInsert(hlLcf, tlLcf, ilLcfRecLen, INDEXKEY3)
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = gConvertErrorCode(ilRet)
                sgErrLoc = "gExtendTFN-Insert Lcf(1)"
                gExtendTFN = False
                Exit Function
            End If
            'If current make Ssf file
            If sLCP = "C" Then
                ilRet = gMakeSSF(False, hlSsf, hlSdf, hlSmf, 0, ilVefCode, ilLogDate0, ilLogDate1, 0)
                Erase sgSSFErrorMsg
                If Not ilRet Then
                    gExtendTFN = False
                    Exit Function
                End If
            End If
        Loop While ilRet = BTRV_ERR_CONFLICT
        ilSeqNo = ilSeqNo + 1
    Loop While ilSeqNo > 0
    gExtendTFN = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gFindBBAvail                    *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Find the closest associated BB  *
'*                     avail to the specified avail    *
'*                                                     *
'*******************************************************
Sub gFindBBAvail(tlSsf As SSF, ilAvail As Integer, slType As String, ilBBAvail As Integer)
'
'   gFindBBAvail tlSsf, ilAvailIndex, slType, ilBBAvailIndex
'   Where:
'       tlSsf(I)- Ssf record image
'       ilAvailIndex(I)- index into Ssf for avail that program is to be found for
'       slType (I)- "O" = open avail; "C" = Close avail
'       ilBBAvailIndex(I)- index into Ssf of BB avail (-1 = not found)
'
'       tgSsf (I)- Ssf to be scan for program
'
    Dim ilAvailIndex As Integer
    Dim slAvailTime As String
    Dim llAvailTime As Long
    Dim slProgTime As String
    Dim llSProgTime As Long
    Dim llEProgTime As Long
    Dim llSPrgTime As Long
    Dim llEPrgTime As Long
    Dim ilRecType As Integer
    Dim ilStep As Integer
    Dim ilBBFound As Integer
    Dim ilBBIndex As Integer
    gFindPrgTimes tlSsf, ilAvail, llSPrgTime, llEPrgTime

    If slType = "O" Then
        ilStep = -1
        ilRecType = 3
    Else
        ilStep = 1
        ilRecType = 5
    End If
    ilAvailIndex = ilAvail + ilStep
    ilBBFound = False
    Do While (ilAvailIndex > 0) And (ilAvailIndex <= tlSsf.iCount)
        LSet tmProgTest = tlSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
        If tmProgTest.iRecType = 1 Then    'Program
            gUnpackTime tmProgTest.iStartTime(0), tmProgTest.iStartTime(1), "A", "1", slProgTime
            llSProgTime = CLng(gTimeToCurrency(slProgTime, False))
            gUnpackTime tmProgTest.iEndTime(0), tmProgTest.iEndTime(1), "A", "1", slProgTime
            llEProgTime = CLng(gTimeToCurrency(slProgTime, True))
            If (llSProgTime = llSPrgTime) And (llEProgTime = llEPrgTime) Then
                Exit Do
            End If
            If (llSProgTime > llEPrgTime) Then
                Exit Do
            End If
            If ilBBFound Then
                If (llSProgTime <= llAvailTime) And (llAvailTime <= llEProgTime) Then
                    ilBBFound = False
                End If
            End If
        ElseIf tmProgTest.iRecType = ilRecType Then    'BB Avail
            tmAvailTest = tlSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
            gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slAvailTime
            llAvailTime = CLng(gTimeToCurrency(slAvailTime, False))
            If Not ilBBFound Then
                If (llSPrgTime <= llAvailTime) And (llAvailTime <= llEPrgTime) Then
                    ilBBFound = True
                    ilBBIndex = ilAvailIndex
                Else
                    Exit Do
                End If
            End If
        End If
        ilAvailIndex = ilAvailIndex + ilStep
    Loop
    If ilBBFound Then
        ilBBAvail = ilBBIndex
    Else
        ilBBAvail = -1
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mFindPrgTimes                   *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Find the program that           *
'*                     encompasses an avail            *
'*                                                     *
'*******************************************************
Sub gFindPrgTimes(tlSsf As SSF, ilAvail As Integer, llPrgStartTime As Long, llPrgEndTime As Long)
'
'   mFindPrgTimes ilAvailIndex, llPrgStartTime, llPrgEndTime
'   Where:
'       ilAvailIndex(I)- index into Ssf for avail that program is to be found for
'       llPrgStartTime (O)- start time of the program
'       llPrgEndTime (O)- end time of the program
'
'       tgSsf (I)- Ssf to be scan for program
'
    Dim ilFound As Integer
    Dim ilAvailIndex As Integer
    Dim slAvailTime As String
    Dim llAvailTime As Long
    Dim slProgTime As String
    Dim llSProgTime As Long
    Dim llEProgTime As Long
    ilAvailIndex = ilAvail - 1
    tmAvailTest = tlSsf.tPas(ADJSSFPASBZ + ilAvail)
    gUnpackTime tmAvailTest.iTime(0), tmAvailTest.iTime(1), "A", "1", slAvailTime
    llAvailTime = CLng(gTimeToCurrency(slAvailTime, False))
    ilFound = False
    Do While ilAvailIndex > 0
        LSet tmProgTest = tlSsf.tPas(ADJSSFPASBZ + ilAvailIndex)
        If tmProgTest.iRecType = 1 Then    'Program
            gUnpackTime tmProgTest.iStartTime(0), tmProgTest.iStartTime(1), "A", "1", slProgTime
            llSProgTime = CLng(gTimeToCurrency(slProgTime, False))
            gUnpackTime tmProgTest.iEndTime(0), tmProgTest.iEndTime(1), "A", "1", slProgTime
            llEProgTime = CLng(gTimeToCurrency(slProgTime, True))
            If (llSProgTime <= llAvailTime) And (llAvailTime <= llEProgTime) Then
                llPrgStartTime = llSProgTime
                llPrgEndTime = llEProgTime
                Exit Sub
            End If
        End If
        ilAvailIndex = ilAvailIndex - 1
    Loop
    llPrgStartTime = CLng(gTimeToCurrency("12:00AM", False))
    llPrgEndTime = CLng(gTimeToCurrency("12:00AM", True))
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gGetLineSchParameters           *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Compute advertiser spearation   *
'*                     contracts                       *
'*                                                     *
'*******************************************************
Sub gGetLineSchParameters(hlSsf As Integer, tlSsf() As SSF, llSsfDate() As Long, llSsfRecPos() As Long, llInDate As Long, ilVefCode As Integer, ilChfAdfCode As Integer, ilGameNo As Integer, tlCff() As CFF, tlClf As CLF, tlRdf As RDF, llSepLength As Long, llStartDateLen As Long, llEndDateLen As Long, llChfCode As Long, ilLineNo As Integer, ilAdfCode As Integer, ilVehComp As Integer, ilTHour() As Integer, ilTDay() As Integer, ilTQH() As Integer, ilAHour() As Integer, ilADay() As Integer, ilAQH() As Integer, llTBStartTime() As Long, llTBEndTime() As Long, llEarliestAllowedDate As Long, ilSkip() As Integer, slType As String, ilPctTrade As Integer, ilBkQH As Integer, ilRetRdfTimes As Integer, ilPriceLevel As Integer, Optional blSetSkip As Boolean = True)
'
'   gGetLineSchParameters hmSsf, tlSsf(), llSsfDate(), llSsfRecPos(), llInDate, ilVefCode, ilChfAdfCode, llChfCode, ilLineNo, tlCff(), tlClf, tlRdf, llSepLength, llStartDateLen, llEndDateLen, ilAdfCode, ilVehComp, imCHour(), imCDay(), imCQH(), imCAHour(), imCADay(), imCAQH(), imTBStart(), imTBEnd(), lmEarliestAllowedDate, imSkip(), slType, ilPctTrade, ilBkQH
'   Where:
'       hlSsf(I)- Handle from Ssf open
'       tlSsf()(I/O)- Ssf record image
'       llSsfDate()(I/O)- Date of Ssf within tlSsf (this is used instead of converting date within tlSsf for speed)
'       llSsfRecPos()(I/O)- Ssf record position (btrGetDirect)
'       llInDate(I)- Date that separation is to be computed for
'       ilChfAdfCode(I)- Advertiser code to compute the separation time for
'       tlCff(I)- Flight records of the line
'       tlClf(I)- line record
'       tlRdf(I)- Rate card program record
'       llSepLength(I/O)- Required separation length between same advertisers (in seconds)
'                       llSepLength = Days*Hours (in Seconds) / 2*NoSpots
'                       If llSepLength >= 3600, then set llSepLength = 3600
'                       If llSepLength < 300, then set llSepLength = 0
'       llStartDateLen(I/O)- Last valid date that separation dates was computed for
'       llEndDateLen(I/O)- Last valid end date that separation for computed for
'       llChfCode(I/O)- Contract code that values where computed for
'       ilLineNo(I/O)- Line number that values where computed for
'       iAdfCode(I/O)- Last valid advertiser separation computed for
'       ilVehComp(I/O) - Last veficle values where computed for
'       imCHour(I)- Total hour counts for line
'       imCDay(I)- Total day counts for line
'       imCQH(I)- Total quarter hour counts
'       imCAHour(O)- Hour counts for legal hours
'       imCADay(O)- Day counts for legal days
'       imCAQH(O)- Quarter hour counts for legal hours
'       imTBStart(I/O)- Allowed start times for all buy types
'       imTBEnd(I/O)- Allowed end times for all buy types
'       lmEarliestAllowedDate(I)- earliest allowed date
'       imSkip()(I/O)- Array of days, hours; quarter hours to bypass
'       slType(I)-Contract type (Chf.sType)
'       ilPctTrade(I)- % trade (Chf.iPctTrade)
'       ilBkQH(O)- Booking number of quarter hours = Number of Days*Number of qurater hours
'       ilRetRdfTimes(I)- True=Return RdfTime; False=Return Avail Times
'       ilPriceLevel(O)- Return Price Level (0=Fill; 1=N/C; 2-15 Price level based on Site or Vehicle setting
'
    Dim ilUpperBound As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim ilDay As Integer
    Dim ilLoopDay As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llFStartDate As Long    'Flight Start Date
    Dim llFEndDate As Long      'Flight End Date
    Dim ilSpots As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilLtfStart0 As Integer
    Dim ilLtfStart1 As Integer
    Dim ilLtfEnd0 As Integer
    Dim ilLtfEnd1 As Integer
    Dim slLtfStart As String
    Dim slLtfEnd As String
    Dim slLnStart As String
    Dim slLnEnd As String
    Dim llLength As Long    'Length of time for buy
    Dim llFreedom As Long
    Dim ilCffIndex As Integer
    Dim ilSsfInMem As Integer
    Dim ilRet As Integer
    Dim ilLastDay As Integer
    Dim ilFirstDay As Integer
    Dim llLoopDate As Long
    Dim ilTBIndex As Integer
    Dim llRecPos As Long
    Dim ilRPRet As Integer
    Dim slTime As String
    Dim llTime As Long
    Dim ilHour As Integer
    Dim ilQH As Integer
    Dim ilSsfRecLen As Integer
    Dim tlSsfSrchKey As SSFKEY0      'SSF key record image
    Dim ilVirtVehSpotAdj As Integer
    Dim llAvailStartTime As Long
    Dim llAvailEndTime As Long
    Dim ilAvail As Integer
    Dim ilDayOk As Integer
    Dim ilVpfIndex As Integer
    Dim ilOtherSpots As Integer
    Dim ilSaf As Integer
    Dim ilPrice As Integer
    Dim llActPrice As Long
    Dim llClfStartDate As Long
    Dim llClfEndDate As Long
    Dim llClfMoFirstWkDate As Long
    Dim llClfMoLastWkDate As Long
    Dim ilType As Integer
    Dim slWkType As String
    Dim ilAvailOk As Integer
    Dim llRdfStartTime As Long
    Dim llRdfEndTime As Long
    Dim ilTBDays(0 To 49) As Integer

    llDate = llInDate
    gUnpackDateLong tlClf.iStartDate(0), tlClf.iStartDate(1), llClfStartDate
    If llClfStartDate > 0 Then
        llClfMoFirstWkDate = llClfStartDate
        Do While gWeekDayLong(llClfMoFirstWkDate) <> 0
            llClfMoFirstWkDate = llClfMoFirstWkDate - 1
        Loop
    Else
        llClfMoFirstWkDate = 0
    End If
    gUnpackDateLong tlClf.iEndDate(0), tlClf.iEndDate(1), llClfEndDate
    If llClfEndDate > 0 Then
        llClfMoLastWkDate = llClfEndDate
        Do While gWeekDayLong(llClfMoLastWkDate) <> 0
            llClfMoLastWkDate = llClfMoLastWkDate - 1
        Loop
    Else
        llClfMoLastWkDate = 9999999
    End If
    If llDate < llClfMoLastWkDate Then
        If (llChfCode = tlClf.lChfCode) And (ilLineNo = tlClf.iLine) And (ilChfAdfCode = ilAdfCode) And (ilVefCode = ilVehComp) And (llDate >= llStartDateLen) And (llDate <= llEndDateLen) Then
            Exit Sub
        End If
    End If
    imVefRecLen = Len(tmVef)  'Get and save VEF record length
'    hmVef = CBtrvTable(ONEHANDLE)        'Create SDF object handle
'    On Error GoTo 0
'    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'    If ilRet = BTRV_ERR_NONE Then
'        tmVefSrchKey.iCode = tlClf.iVefCode 'ilVefCode
'        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'        If ilRet <> BTRV_ERR_NONE Then
'            tmVef.sType = ""
'        End If
'    Else
'        tmVef.sType = ""
'    End If
'    ilRet = btrClose(hmVef)
'    btrDestroy hmVef
    ilRet = gBinarySearchVef(tlClf.iVefCode)
    If ilRet <> -1 Then
        tmVef = tgMVef(ilRet)
    Else
        tmVef.sType = ""
    End If
    ilVpfIndex = gVpfFindIndex(ilVefCode)
    If tmVef.sType = "V" Then
        If tmVef.iCode <> ilVefCode Then    'Check if virtual vehicle getting parameters for
            imVsfRecLen = Len(tmVsf)  'Get and save VSEF record length
            hmVsf = CBtrvTable(ONEHANDLE)        'Create SDF object handle
            On Error GoTo 0
            ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE Then
                tmVsfSrchKey.lCode = tmVef.lVsfCode
                ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet = BTRV_ERR_NONE Then
                    ilVirtVehSpotAdj = 1
                    For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                        If tmVsf.iFSCode(ilLoop) = ilVefCode Then
                            ilVirtVehSpotAdj = tmVsf.iNoSpots(ilLoop)
                            Exit For
                        End If
                    Next ilLoop
                Else
                    ilVirtVehSpotAdj = 1
                End If
            Else
                ilVirtVehSpotAdj = 1
            End If
            ilRet = btrClose(hmVsf)
            btrDestroy hmVsf
        Else
            ilVirtVehSpotAdj = 1
        End If
    Else
        ilVirtVehSpotAdj = 1
    End If
    For ilLoop = LBound(llTBStartTime) To UBound(llTBStartTime) Step 1
        llTBStartTime(ilLoop) = -1
        llTBEndTime(ilLoop) = -1
        ilTBDays(ilLoop) = -1
    Next ilLoop
    'Turn skip hours/qh/days off if no avail exist
    For ilHour = LBound(ilSkip, 1) To UBound(ilSkip, 1) Step 1
        For ilQH = LBound(ilSkip, 2) To UBound(ilSkip, 2) Step 1
            For ilDay = LBound(ilSkip, 3) To UBound(ilSkip, 3) Step 1
                ilSkip(ilHour, ilQH, ilDay) = -1
            Next ilDay
        Next ilQH
    Next ilHour
    For ilLoop = LBound(ilAHour) To UBound(ilAHour) Step 1
        ilAHour(ilLoop) = -1
    Next ilLoop
    For ilLoop = LBound(ilADay) To UBound(ilADay) Step 1
        ilADay(ilLoop) = -1
    Next ilLoop
    For ilLoop = LBound(ilAQH) To UBound(ilAQH) Step 1
        ilAQH(ilLoop) = -1
    Next ilLoop
    ilType = ilGameNo
    ilLastDay = -1
    ilFirstDay = -1
    ilUpperBound = UBound(tlCff) - 1
    ilDay = gWeekDayLong(llDate)
    slDate = Format$(llDate, "m/d/yy")
    gPackDate slDate, ilDate0, ilDate1
    ilPriceLevel = 0
    ilSpots = 0
    ilCffIndex = -1
    llActPrice = -1
    For ilLoop = LBound(tlCff) To ilUpperBound Step 1
        gUnpackDate tlCff(ilLoop).iStartDate(0), tlCff(ilLoop).iStartDate(1), slDate
        llStartDate = gDateValue(slDate)
        gUnpackDate tlCff(ilLoop).iEndDate(0), tlCff(ilLoop).iEndDate(1), slDate
        llEndDate = gDateValue(slDate)
        If (llDate >= llStartDate) And (llDate <= llEndDate) Then
            ilCffIndex = ilLoop
            '7/8/18: Retain price so that the zero or N/C Rank can be set
            If tlCff(ilLoop).sPriceType = "T" Then
                llActPrice = tlCff(ilLoop).lActPrice
            Else
                llActPrice = 0
            End If
            llFStartDate = llStartDate
            llFEndDate = llEndDate
            If (tlCff(ilLoop).sDyWk <> "D") Then  'Weekly
                For ilIndex = 0 To 6 Step 1
                    If tlCff(ilLoop).iDay(ilIndex) > 0 Then
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
                        If tlCff(ilLoop).sXDay(ilIndex) = "1" Then
                            If ilFirstDay = -1 Then
                                ilFirstDay = ilIndex
                                ilSpots = tlCff(ilLoop).iXSpotsWk
                            End If
                            ilLastDay = ilIndex
                        End If
                    Next ilIndex
                    '2/20/11: Test if day range found
                    If ilFirstDay = -1 Then
                        'Get here if llDate was not on a valid week day
                        For ilIndex = 0 To 6 Step 1
                            If tlCff(ilLoop).iDay(ilIndex) > 0 Then
                                If ilFirstDay = -1 Then
                                    ilFirstDay = ilIndex
                                End If
                                ilLastDay = ilIndex
                            End If
                        Next ilIndex
                    End If
                Else
                    ilSpots = tlCff(ilLoop).iSpotsWk
                End If
            Else
                ilSpots = tlCff(ilLoop).iDay(ilDay)
                ilFirstDay = ilDay
                ilLastDay = ilDay
            End If
            '7/11/18: Move checking if price levels defined prior to testing of price.
            ilSaf = gBinarySearchSaf(tlClf.iVefCode)
            If ilSaf = -1 Then
                ilSaf = gBinarySearchSaf(0) 'Obtain from Site
            Else
                If (tgSaf(ilSaf).lLowPrice <= 0) And (tgSaf(ilSaf).lHighPrice <= 0) Then
                    ilSaf = gBinarySearchSaf(0) 'Obtain from Site
                End If
            End If
            If ilSaf <> -1 Then
                If tlCff(ilLoop).sPriceType = "T" Then
                    'ilSaf = gBinarySearchSaf(tlClf.iVefCode)
                    'If ilSaf = -1 Then
                    '    ilSaf = gBinarySearchSaf(0)
                    'Else
                    '    If (tgSaf(ilSaf).lLowPrice <= 0) And (tgSaf(ilSaf).lHighPrice <= 0) Then
                    '        ilSaf = gBinarySearchSaf(0)
                    '    End If
                    'End If
                    'If ilSaf <> -1 Then
                        '6/20/06:  Treat zero dollars same as N/C
                        If (tgSaf(ilSaf).lLowPrice <= 0) And (tgSaf(ilSaf).lHighPrice <= 0) Then
                            ilPriceLevel = 0
                        ElseIf tlCff(ilLoop).lActPrice = 0 Then 'treat as N/C
                            ilPriceLevel = 1
                        ElseIf tlCff(ilLoop).lActPrice <= 100 * tgSaf(ilSaf).lLowPrice Then
                            ilPriceLevel = 2
                        Else
                            If tlCff(ilLoop).lActPrice > 100 * tgSaf(ilSaf).lHighPrice Then
                                ilPriceLevel = 15
                            Else
                                For ilPrice = LBound(tgSaf(ilSaf).lLevelToPrice) To UBound(tgSaf(ilSaf).lLevelToPrice) Step 1
                                    If tlCff(ilLoop).lActPrice <= 100 * tgSaf(ilSaf).lLevelToPrice(ilPrice) Then
                                        ilPriceLevel = ilPrice - LBound(tgSaf(ilSaf).lLevelToPrice) + 3
                                        Exit For
                                    End If
                                Next ilPrice
                                If ilPriceLevel = 0 Then
                                    If tlCff(ilLoop).lActPrice > 100 * tgSaf(ilSaf).lLevelToPrice(UBound(tgSaf(ilSaf).lLevelToPrice)) And (tlCff(ilLoop).lActPrice <= 100 * tgSaf(ilSaf).lHighPrice) Then
                                        ilPriceLevel = 14
                                    End If
                                End If
                            End If
                       End If
                    'End If
                Else
                    If (tgSaf(ilSaf).lLowPrice <= 0) And (tgSaf(ilSaf).lHighPrice <= 0) Then
                        ilPriceLevel = 0
                    Else
                        ilPriceLevel = 1
                    End If
                End If
            Else
                ilPriceLevel = 0
            End If
            Exit For
        End If
    Next ilLoop
    If (ilCffIndex >= 0) And (ilFirstDay >= 0) Then
        If (tlCff(ilCffIndex).sDyWk = "D") Then  'Daily- Test if valid day
            llStartDate = llDate
            llEndDate = llDate
        Else
            '2/20/11:  Replace code to handle case where llDate is outside of valid days.  i.e. On Sunday but allowed days are Mo-Fr
            'llStartDate = llDate
            'Do Until gWeekDayLong(llStartDate) = ilFirstDay   'Backup to monday
            '    If llStartDate <= llFStartDate Then
            '        Exit Do
            '    End If
            '    llStartDate = llStartDate - 1
            'Loop
            'llEndDate = llDate
            'Do Until gWeekDayLong(llEndDate) = ilLastDay
            '    If llEndDate >= llFEndDate Then
            '        Exit Do
            '    End If
            '    llEndDate = llEndDate + 1
            'Loop
            '2/20/11: Check that day is within week days
            'Get first valid date of week
            If ilFirstDay <= ilDay Then
                llStartDate = llDate
                Do Until gWeekDayLong(llStartDate) = ilFirstDay   'Backup to monday
                    If llStartDate <= llFStartDate Then
                        Exit Do
                    End If
                    llStartDate = llStartDate - 1
                Loop
            Else
                llStartDate = llDate
                Do Until gWeekDayLong(llStartDate) = ilFirstDay   'Backup to monday
                    If llStartDate >= llFEndDate Then
                        Exit Do
                    End If
                    llStartDate = llStartDate + 1
                Loop
            End If
            'Get last valid date of week
            If ilLastDay >= ilDay Then
                llEndDate = llDate
                Do Until gWeekDayLong(llEndDate) = ilLastDay
                    If llEndDate >= llFEndDate Then
                        Exit Do
                    End If
                    llEndDate = llEndDate + 1
                Loop
            Else
                llEndDate = llDate
                Do Until gWeekDayLong(llEndDate) = ilLastDay
                    If llEndDate <= llFStartDate Then
                        Exit Do
                    End If
                    llEndDate = llEndDate - 1
                Loop
            End If
            If llDate < llStartDate Then
                llDate = llStartDate
            End If
            If llDate > llEndDate Then
                llDate = llEndDate
            End If
            ilDay = gWeekDayLong(llDate)
        End If
    Else
        llStartDate = llDate
        llEndDate = llDate
    End If
    llStartDateLen = llStartDate
    llEndDateLen = llEndDate
    ilAdfCode = ilChfAdfCode
    llChfCode = tlClf.lChfCode
    ilLineNo = tlClf.iLine
    ilVehComp = ilVefCode
    If tlClf.iPriority <= 0 Then
        '1045 used for extra
        If ilPctTrade <> 100 Then
            If slType = "R" Then    'Direct Response
                ilBkQH = DIRECTRESPONSERANK '1010
            ElseIf slType = "T" Then    'Remnant
                ilBkQH = REMNANTRANK    '1020
            ElseIf slType = "Q" Then    'per Inquiry
                ilBkQH = PERINQUIRYRANK '1030
            ElseIf slType = "M" Then    'Promo
                ilBkQH = PROMORANK  '1050
            ElseIf slType = "S" Then    'PSA
                ilBkQH = PSARANK    '1060
            ElseIf slType = "V" Then    'Reservation
                ilBkQH = RESERVATIONRANK
            Else
                ilBkQH = 0   'Outside limits
            End If
        Else    '100 trade
            ilBkQH = TRADERANK  '1040
        End If
        '7/8/18: Zero dollar spots should not pre-empt a charge spot
        If (llActPrice = 0) And (ilBkQH = 0) Then
            ilBkQH = NOCHARGERANK
        End If
    Else
        ilBkQH = tlClf.iPriority
    End If
    'If tmVef.sType = "V" Then   'If virtual vehicle- set to highest priority
    '    ilBkQH = 1
    'End If
    If llEndDateLen < llEarliestAllowedDate Then
        Exit Sub
    End If
    ilSpots = ilVirtVehSpotAdj * ilSpots
    If ilSpots = 0 Then
        llSepLength = 0
        Exit Sub
    End If
    '2/20/11:  Moved up
    'For ilLoop = LBound(llTBStartTime) To UBound(llTBStartTime) Step 1
    '    llTBStartTime(ilLoop) = -1
    '    llTBEndTime(ilLoop) = -1
    '    ilTBDays(ilLoop) = -1
    'Next ilLoop
    If ilFirstDay = -1 Then
        llSepLength = 0
        Exit Sub
    End If
    For ilLoop = ilFirstDay To ilLastDay Step 1
        'ilADay(ilLoop + 1) = ilTDay(ilLoop + 1)
        ilADay(ilLoop) = ilTDay(ilLoop)
    Next ilLoop
    'ilTBIndex = 1
    ilTBIndex = LBound(llTBStartTime)
    'Determine amount of time to schedule the spots within
    llLength = 0&
    ilLtfStart0 = -1
    slWkType = ""
    If (llDate >= llClfMoFirstWkDate) And (llDate <= llClfMoFirstWkDate + 6) Then
        If llDate >= llClfMoLastWkDate Then
            slWkType = "B"
        Else
            slWkType = "F"
        End If
    Else
        If llDate >= llClfMoLastWkDate Then
            slWkType = "L"
        End If
    End If
    gXMidClfRdfToRdf slWkType, tlClf, tlRdf, tmXMidClf, tmXMidRdf
    If (tmXMidRdf.iLtfCode(0) <> 0) Or (tmXMidRdf.iLtfCode(1) <> 0) Or (tmXMidRdf.iLtfCode(2) <> 0) Then
        'Read Ssf for date- test for library
        For llLoopDate = llStartDate To llEndDate Step 1
            ilLoopDay = gWeekDayLong(llLoopDate)
            If (tlCff(ilCffIndex).iDay(ilIndex) > 0) Or (tlCff(ilCffIndex).sXDay(ilIndex) = "1") Then
                ilSsfInMem = False
                slDate = Format$(llLoopDate, "m/d/yy")
                gPackDate slDate, ilDate0, ilDate1
                If ((llSsfDate(ilLoopDay) = llLoopDate) And (ilType = 0)) Or ((llSsfDate(ilLoopDay) = ilType) And (ilType <> 0)) Then
                    If (tlSsf(ilLoopDay).iType = ilType) And (tlSsf(ilLoopDay).iVefCode = ilVefCode) And (tlSsf(ilLoopDay).iStartTime(0) = 0) And (tlSsf(ilLoopDay).iStartTime(1) = 0) Then
                        ilSsfInMem = True
                        llRecPos = llSsfRecPos(ilLoopDay)
                        ilRet = BTRV_ERR_NONE
                    End If
                End If
                If Not ilSsfInMem Then
                    ilSsfRecLen = Len(tlSsf(ilLoopDay)) 'Max size of variable length record
                    tlSsfSrchKey.iType = ilType
                    tlSsfSrchKey.iVefCode = ilVefCode
                    tlSsfSrchKey.iDate(0) = ilDate0
                    tlSsfSrchKey.iDate(1) = ilDate1
                    tlSsfSrchKey.iStartTime(0) = 0
                    tlSsfSrchKey.iStartTime(1) = 0
                    ilRet = gSSFGetGreaterOrEqual(hlSsf, tlSsf(ilLoopDay), ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                    ilRPRet = gSSFGetPosition(hlSsf, llRecPos)
                End If
                Do While (ilRet = BTRV_ERR_NONE) And (tlSsf(ilLoopDay).iType = ilType) And (tlSsf(ilLoopDay).iVefCode = ilVefCode) And (tlSsf(ilLoopDay).iDate(0) = ilDate0) And (tlSsf(ilLoopDay).iDate(1) = ilDate1)
                    If ilType <= 0 Then
                        llSsfDate(ilLoopDay) = llLoopDate
                    Else
                        llSsfDate(ilLoopDay) = ilType
                    End If
                    llSsfRecPos(ilLoopDay) = llRecPos
                    For ilLoop = 1 To tlSsf(ilLoopDay).iCount Step 1
                       LSet tmProg = tlSsf(ilLoopDay).tPas(ADJSSFPASBZ + ilLoop)
                        ilAvailOk = False
                        If tmProg.iRecType = 2 Then
                            If (slType <> "M") And (slType <> "S") Then
                                ilAvailOk = True
                            Else
                                If slType = "M" Then
                                    If tgSpf.sSchdPromo = "Y" Then
                                        If ((Asc(tgSpf.sUsingFeatures3) And PROMOINTOCONTRACTAVAILS) = PROMOINTOCONTRACTAVAILS) Then
                                            ilAvailOk = True
                                        End If
                                    End If
                                End If
                                If slType = "S" Then
                                    If tgSpf.sSchdPSA = "Y" Then
                                        If ((Asc(tgSpf.sUsingFeatures3) And PSAINTOCONTRACTAVAILS) = PSAINTOCONTRACTAVAILS) Then
                                            ilAvailOk = True
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        If tmProg.iRecType = 1 Then 'Program subrecord
                            If (tmProg.iLtfCode = tmXMidRdf.iLtfCode(0)) Or (tmProg.iLtfCode = tmXMidRdf.iLtfCode(1)) Or (tmProg.iLtfCode = tmXMidRdf.iLtfCode(2)) Then
                                If ilLtfStart0 = -1 Then
                                    ilLtfStart0 = tmProg.iStartTime(0)
                                    ilLtfStart1 = tmProg.iStartTime(1)
                                    ilLtfEnd0 = tmProg.iEndTime(0)  'Set end time incase no other events within prog
                                    ilLtfEnd1 = tmProg.iEndTime(1)
                                Else    'set end time as a running time to handle nested progs
                                    ilLtfEnd0 = tmProg.iStartTime(0)
                                    ilLtfEnd1 = tmProg.iStartTime(1)
                                End If
                            Else
                                If ilLtfStart0 <> -1 Then
                                    gUnpackTime ilLtfStart0, ilLtfStart1, "A", "1", slLtfStart
                                    gUnpackTime ilLtfEnd0, ilLtfEnd1, "A", "1", slLtfEnd
                                    mSetSepValues slLtfStart, slLtfEnd, llLength, ilTHour(), ilTQH(), ilAHour(), ilAQH()
                                    'Save library times for spot move time test
                                    If (ilTBIndex <= UBound(llTBStartTime)) And (ilDay = ilLoopDay) Then
                                        llTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLtfStart, False))
                                        llTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLtfEnd, True))
                                        ilTBDays(ilTBIndex) = ilLoopDay
                                        ilTBIndex = ilTBIndex + 1
                                    End If
                                End If
                                ilLtfStart0 = -1
                            End If
                        'ElseIf ((tmProg.iRecType = 2) And (slType <> "M") And (slType <> "S")) Or ((tmProg.iRecType = 8) And (slType = "S")) Or ((tmProg.iRecType = 9) And (slType = "M")) Then 'Contract Avail
                        ElseIf (ilAvailOk) Or ((tmProg.iRecType = 8) And (slType = "S")) Or ((tmProg.iRecType = 9) And (slType = "M")) Or ((tmProg.iRecType = 2) And (slType = "M") And (tgSpf.sSchdPromo <> "Y")) Or ((tmProg.iRecType = 2) And (slType = "S") And (tgSpf.sSchdPSA <> "Y")) Then
                           LSet tmAvail = tlSsf(ilLoopDay).tPas(ADJSSFPASBZ + ilLoop)
                            If (tmAvail.iLtfCode = tmXMidRdf.iLtfCode(0)) Or (tmAvail.iLtfCode = tmXMidRdf.iLtfCode(1)) Or (tmAvail.iLtfCode = tmXMidRdf.iLtfCode(2)) Then
                                If ilLtfStart0 = -1 Then
                                    ilLtfStart0 = tmAvail.iTime(0)
                                    ilLtfStart1 = tmAvail.iTime(1)
                                    'Compute end time of avail incase only one event for the library
                                    ilLtfEnd0 = tmAvail.iTime(0)
                                    ilLtfEnd1 = tmAvail.iTime(1)
                                    gUnpackTime ilLtfEnd0, ilLtfEnd1, "A", "1", slLtfEnd
                                    slLtfEnd = gCurrencyToTime(gTimeToCurrency(slLtfEnd, False) + tmAvail.iLen)
                                    gPackTime slLtfEnd, ilLtfEnd0, ilLtfEnd1
                                Else    'set end time as a running time to handle nested progs
                                    ilLtfEnd0 = tmAvail.iTime(0)
                                    ilLtfEnd1 = tmAvail.iTime(1)
                                End If
                            Else
                                If ilLtfStart0 <> -1 Then
                                    gUnpackTime ilLtfStart0, ilLtfStart1, "A", "1", slLtfStart
                                    gUnpackTime ilLtfEnd0, ilLtfEnd1, "A", "1", slLtfEnd
                                    mSetSepValues slLtfStart, slLtfEnd, llLength, ilTHour(), ilTQH(), ilAHour(), ilAQH()
                                    'Save library times for spot move time test
                                    If (ilTBIndex <= UBound(llTBStartTime)) And (ilDay = ilLoopDay) Then
                                        llTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLtfStart, False))
                                        llTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLtfEnd, True))
                                        ilTBDays(ilTBIndex) = ilLoopDay
                                        ilTBIndex = ilTBIndex + 1
                                    End If
                                End If
                                ilLtfStart0 = -1
                            End If
                        End If
                    Next ilLoop
                    If ilLtfStart0 <> -1 Then
                        gUnpackTime ilLtfStart0, ilLtfStart1, "A", "1", slLtfStart
                        gUnpackTime ilLtfEnd0, ilLtfEnd1, "A", "1", slLtfEnd
                        mSetSepValues slLtfStart, slLtfEnd, llLength, ilTHour(), ilTQH(), ilAHour(), ilAQH()
                        'Save library times for spot move time test
                        If (ilTBIndex <= UBound(llTBStartTime)) And (ilDay = ilLoopDay) Then
                            llTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLtfStart, False))
                            llTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLtfEnd, True))
                            ilTBDays(ilTBIndex) = ilLoopDay
                            ilTBIndex = ilTBIndex + 1
                        End If
                    End If
                    ilLtfStart0 = -1
                    'If (tlSsf(ilLoopDay).iNextTime(0) = 1) And (tlSsf(ilLoopDay).iNextTime(1) = 0) Then
                        Exit Do
                    'Else
                    '    ilSsfRecLen = Len(tlSsf(ilLoopDay)) 'Max size of variable length record
                    '    ilRet = gSSFGetNext(hlSsf, tlSsf(ilLoopDay), ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    '    ilRPRet = gSSFGetPosition(hlSsf, llRecPos)
                    'End If
                Loop
            End If
        Next llLoopDate
    Else    'Time buy- check if override times defined (if so, use them as bump times)
        If ((tmXMidClf.iStartTime(0) = 1) And (tmXMidClf.iStartTime(1) = 0)) Or (tgVpf(ilVpfIndex).sGMedium = "S") Then
            For ilLoop = LBound(tmXMidRdf.iStartTime, 2) To UBound(tmXMidRdf.iStartTime, 2) Step 1
                If (tmXMidRdf.iStartTime(0, ilLoop) <> 1) Or (tmXMidRdf.iStartTime(1, ilLoop) <> 0) Then
                    If (ilTBIndex <= UBound(llTBStartTime)) And (ilRetRdfTimes) Then
                        'gUnpackTimeLong tmXMidRdf.iStartTime(0, ilLoop), tmXMidRdf.iStartTime(1, ilLoop), False, llTBStartTime(ilTBIndex)
                        'gUnpackTimeLong tmXMidRdf.iEndTime(0, ilLoop), tmXMidRdf.iEndTime(1, ilLoop), True, llTBEndTime(ilTBIndex)
                        'ilTBIndex = ilTBIndex + 1
                        If (tlCff(ilCffIndex).sDyWk = "D") Then  'Daily- Test if valid day
                            'If tmXMidRdf.sWkDays(ilLoop, ilDay + 1) = "Y" Then
                            If tmXMidRdf.sWkDays(ilLoop, ilDay) = "Y" Then
                                gUnpackTimeLong tmXMidRdf.iStartTime(0, ilLoop), tmXMidRdf.iStartTime(1, ilLoop), False, llTBStartTime(ilTBIndex)
                                gUnpackTimeLong tmXMidRdf.iEndTime(0, ilLoop), tmXMidRdf.iEndTime(1, ilLoop), True, llTBEndTime(ilTBIndex)
                                ilTBDays(ilTBIndex) = ilDay
                                ilTBIndex = ilTBIndex + 1
                            End If
                        Else
                            For ilIndex = ilFirstDay To ilLastDay Step 1
                                If (tlCff(ilCffIndex).iDay(ilIndex) = 1) Or (tlCff(ilCffIndex).sXDay(ilIndex) = "1") Then
                                    'If tmXMidRdf.sWkDays(ilLoop, ilIndex + 1) = "Y" Then
                                    If tmXMidRdf.sWkDays(ilLoop, ilIndex) = "Y" Then
                                        gUnpackTimeLong tmXMidRdf.iStartTime(0, ilLoop), tmXMidRdf.iStartTime(1, ilLoop), False, llTBStartTime(ilTBIndex)
                                        gUnpackTimeLong tmXMidRdf.iEndTime(0, ilLoop), tmXMidRdf.iEndTime(1, ilLoop), True, llTBEndTime(ilTBIndex)
                                        ilTBDays(ilTBIndex) = ilIndex
                                        ilTBIndex = ilTBIndex + 1
                                    End If
                                End If
                            Next ilIndex
                        End If
                    End If
                    If (tlCff(ilCffIndex).sDyWk = "D") Then  'Daily- Test if valid day
                        'If tmXMidRdf.sWkDays(ilLoop, ilDay + 1) = "Y" Then
                        If tmXMidRdf.sWkDays(ilLoop, ilDay) = "Y" Then
                            'Compute Avail times if smaller then Rate Card time- use them
                            llAvailStartTime = -1
                            llAvailEndTime = 0
                            For ilIndex = ilFirstDay To ilLastDay Step 1
                                llLoopDate = llStartDate + ilIndex - ilFirstDay
                                ilLoopDay = ilIndex
                                slTime = "12:00AM"
                                If gObtainSsfForDateOrGame(ilVefCode, llLoopDate, slTime, ilGameNo, hlSsf, tlSsf(ilLoopDay), llSsfDate(ilLoopDay), llSsfRecPos(ilLoopDay)) Then
                                    For ilAvail = 1 To tlSsf(ilLoopDay).iCount Step 1
                                       LSet tmAvail = tlSsf(ilLoopDay).tPas(ADJSSFPASBZ + ilAvail)
                                        'If ((tmAvail.iRecType = 2) And (slType <> "M") And (slType <> "S")) Or ((tmAvail.iRecType = 8) And (slType = "S")) Or ((tmAvail.iRecType = 9) And (slType = "M")) Then
                                        ilAvailOk = False
                                        If tmAvail.iRecType = 2 Then
                                            If (slType <> "M") And (slType <> "S") Then
                                                ilAvailOk = True
                                            Else
                                                If slType = "M" Then
                                                    If tgSpf.sSchdPromo = "Y" Then
                                                        If ((Asc(tgSpf.sUsingFeatures3) And PROMOINTOCONTRACTAVAILS) = PROMOINTOCONTRACTAVAILS) Then
                                                            ilAvailOk = True
                                                        End If
                                                    End If
                                                End If
                                                If slType = "S" Then
                                                    If tgSpf.sSchdPSA = "Y" Then
                                                        If ((Asc(tgSpf.sUsingFeatures3) And PSAINTOCONTRACTAVAILS) = PSAINTOCONTRACTAVAILS) Then
                                                            ilAvailOk = True
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                        If (ilAvailOk) Or ((tmAvail.iRecType = 8) And (slType = "S")) Or ((tmAvail.iRecType = 9) And (slType = "M")) Or ((tmAvail.iRecType = 2) And (slType = "M") And (tgSpf.sSchdPromo <> "Y")) Or ((tmAvail.iRecType = 2) And (slType = "S") And (tgSpf.sSchdPSA <> "Y")) Then
                                            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                            If llAvailStartTime = -1 Then
                                                llAvailStartTime = llTime
                                            End If
                                            llAvailEndTime = llTime + tmAvail.iLen 'Include last avail (llTBEndTime)
                                        End If
                                    Next ilAvail
                                End If
                            Next ilIndex
                            If llAvailStartTime >= 0 Then
                                'gUnpackTimeLong tmXMidRdf.iStartTime(0, ilLoop), tmXMidRdf.iStartTime(1, ilLoop), False, llTime
                                'If (llAvailStartTime = 0) Or (llTime >= llAvailStartTime) Then
                                '    gUnpackTime tmXMidRdf.iStartTime(0, ilLoop), tmXMidRdf.iStartTime(1, ilLoop), "A", "1", slLnStart
                                'Else
                                '    slLnStart = gFormatTimeLong(llAvailStartTime, "A", "1")
                                'End If
                                'gUnpackTimeLong tmXMidRdf.iEndTime(0, ilLoop), tmXMidRdf.iEndTime(1, ilLoop), True, llTime
                                'If (llAvailEndTime = 0) Or (llTime <= llAvailEndTime) Then
                                gUnpackTimeLong tmXMidRdf.iStartTime(0, ilLoop), tmXMidRdf.iStartTime(1, ilLoop), False, llRdfStartTime
                                llTime = llRdfStartTime
                                gUnpackTimeLong tmXMidRdf.iEndTime(0, ilLoop), tmXMidRdf.iEndTime(1, ilLoop), True, llRdfEndTime
                                If (llAvailStartTime = 0) Or (llAvailStartTime < llTime) Or (llAvailStartTime >= llRdfEndTime) Then
                                    gUnpackTime tmXMidRdf.iStartTime(0, ilLoop), tmXMidRdf.iStartTime(1, ilLoop), "A", "1", slLnStart
                                Else
                                    slLnStart = gFormatTimeLong(llAvailStartTime, "A", "1")
                                End If
                                llTime = llRdfEndTime
                                If (llAvailEndTime = 0) Or (llAvailEndTime >= llTime) Or (llAvailEndTime < llRdfStartTime) Then
                                    gUnpackTime tmXMidRdf.iEndTime(0, ilLoop), tmXMidRdf.iEndTime(1, ilLoop), "A", "1", slLnEnd
                                Else
                                    slLnEnd = gFormatTimeLong(llAvailEndTime, "A", "1")
                                End If
                                If gTimeToLong(slLnEnd, True) >= gTimeToLong(slLnStart, False) Then
                                    mSetSepValues slLnStart, slLnEnd, llLength, ilTHour(), ilTQH(), ilAHour(), ilAQH()
                                    If (ilTBIndex <= UBound(llTBStartTime)) And (Not ilRetRdfTimes) Then
                                        llTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLnStart, False))
                                        llTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLnEnd, True))
                                        ilTBDays(ilTBIndex) = ilDay
                                        ilTBIndex = ilTBIndex + 1
                                    End If
                                Else
                                    mSetSepValues slLnStart, "12M", llLength, ilTHour(), ilTQH(), ilAHour(), ilAQH()
                                    If (ilTBIndex <= UBound(llTBStartTime)) And (Not ilRetRdfTimes) Then
                                        llTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLnStart, False))
                                        llTBEndTime(ilTBIndex) = CLng(gTimeToCurrency("12M", True))
                                        ilTBDays(ilTBIndex) = ilDay
                                        ilTBIndex = ilTBIndex + 1
                                    End If
                                    mSetSepValues "12M", slLnEnd, llLength, ilTHour(), ilTQH(), ilAHour(), ilAQH()
                                    If (ilTBIndex <= UBound(llTBStartTime)) And (Not ilRetRdfTimes) Then
                                        llTBStartTime(ilTBIndex) = CLng(gTimeToCurrency("12M", False))
                                        llTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLnEnd, True))
                                        ilTBDays(ilTBIndex) = ilDay
                                        ilTBIndex = ilTBIndex + 1
                                    End If
                                End If
                            End If
                        End If
                    Else    'Add time for each valid day
                        For ilIndex = ilFirstDay To ilLastDay Step 1
                            If (tlCff(ilCffIndex).iDay(ilIndex) = 1) Or (tlCff(ilCffIndex).sXDay(ilIndex) = "1") Then
                                'If tmXMidRdf.sWkDays(ilLoop, ilIndex + 1) = "Y" Then
                                If tmXMidRdf.sWkDays(ilLoop, ilIndex) = "Y" Then

                                    llAvailStartTime = -1
                                    llAvailEndTime = 0
                                    llLoopDate = llStartDate + ilIndex - ilFirstDay
                                    ilLoopDay = ilIndex
                                    slTime = "12:00AM"
                                    If gObtainSsfForDateOrGame(ilVefCode, llLoopDate, slTime, ilGameNo, hlSsf, tlSsf(ilLoopDay), llSsfDate(ilLoopDay), llSsfRecPos(ilLoopDay)) Then
                                        For ilAvail = 1 To tlSsf(ilLoopDay).iCount Step 1
                                           LSet tmAvail = tlSsf(ilLoopDay).tPas(ADJSSFPASBZ + ilAvail)
                                            ilAvailOk = False
                                            If tmAvail.iRecType = 2 Then
                                                If (slType <> "M") And (slType <> "S") Then
                                                    ilAvailOk = True
                                                Else
                                                    If slType = "M" Then
                                                        If tgSpf.sSchdPromo = "Y" Then
                                                            If ((Asc(tgSpf.sUsingFeatures3) And PROMOINTOCONTRACTAVAILS) = PROMOINTOCONTRACTAVAILS) Then
                                                                ilAvailOk = True
                                                            End If
                                                        End If
                                                    End If
                                                    If slType = "S" Then
                                                        If tgSpf.sSchdPSA = "Y" Then
                                                            If ((Asc(tgSpf.sUsingFeatures3) And PSAINTOCONTRACTAVAILS) = PSAINTOCONTRACTAVAILS) Then
                                                                ilAvailOk = True
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            'If ((tmAvail.iRecType = 2) And (slType <> "M") And (slType <> "S")) Or ((tmAvail.iRecType = 8) And (slType = "S")) Or ((tmAvail.iRecType = 9) And (slType = "M")) Then
                                            If (ilAvailOk) Or ((tmAvail.iRecType = 8) And (slType = "S")) Or ((tmAvail.iRecType = 9) And (slType = "M")) Or ((tmAvail.iRecType = 2) And (slType = "M") And (tgSpf.sSchdPromo <> "Y")) Or ((tmAvail.iRecType = 2) And (slType = "S") And (tgSpf.sSchdPSA <> "Y")) Then
                                                gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                                                If llAvailStartTime = -1 Then
                                                    llAvailStartTime = llTime
                                                End If
                                                llAvailEndTime = llTime + tmAvail.iLen  'Include last avail (llTBEndTime)
                                            End If
                                        Next ilAvail
                                    End If
                                    If llAvailStartTime >= 0 Then
                                        'gUnpackTimeLong tmXMidRdf.iStartTime(0, ilLoop), tmXMidRdf.iStartTime(1, ilLoop), False, llTime
                                        'If (llAvailStartTime = 0) Or (llTime >= llAvailStartTime) Then
                                        '    gUnpackTime tmXMidRdf.iStartTime(0, ilLoop), tmXMidRdf.iStartTime(1, ilLoop), "A", "1", slLnStart
                                        'Else
                                        '    slLnStart = gFormatTimeLong(llAvailStartTime, "A", "1")
                                        'End If
                                        'gUnpackTimeLong tmXMidRdf.iEndTime(0, ilLoop), tmXMidRdf.iEndTime(1, ilLoop), True, llTime
                                        'If (llAvailEndTime = 0) Or (llTime <= llAvailEndTime) Then
                                        gUnpackTimeLong tmXMidRdf.iStartTime(0, ilLoop), tmXMidRdf.iStartTime(1, ilLoop), False, llRdfStartTime
                                        llTime = llRdfStartTime
                                        gUnpackTimeLong tmXMidRdf.iEndTime(0, ilLoop), tmXMidRdf.iEndTime(1, ilLoop), True, llRdfEndTime
                                        If (llAvailStartTime = 0) Or (llAvailStartTime < llTime) Or (llAvailStartTime >= llRdfEndTime) Then
                                            gUnpackTime tmXMidRdf.iStartTime(0, ilLoop), tmXMidRdf.iStartTime(1, ilLoop), "A", "1", slLnStart
                                        Else
                                            slLnStart = gFormatTimeLong(llAvailStartTime, "A", "1")
                                        End If
                                        llTime = llRdfEndTime
                                        If (llAvailEndTime = 0) Or (llAvailEndTime >= llTime) Or (llAvailEndTime < llRdfStartTime) Then
                                            gUnpackTime tmXMidRdf.iEndTime(0, ilLoop), tmXMidRdf.iEndTime(1, ilLoop), "A", "1", slLnEnd
                                        Else
                                            slLnEnd = gFormatTimeLong(llAvailEndTime, "A", "1")
                                        End If
                                        mSetSepValues slLnStart, slLnEnd, llLength, ilTHour(), ilTQH(), ilAHour(), ilAQH()
                                        If (ilTBIndex <= UBound(llTBStartTime)) And (Not ilRetRdfTimes) Then
                                            llTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLnStart, False))
                                            llTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLnEnd, True))
                                            ilTBDays(ilTBIndex) = ilIndex
                                            ilTBIndex = ilTBIndex + 1
                                        End If
                                    End If
                                End If
                            End If
                        Next ilIndex
                    End If
                End If
            Next ilLoop
        Else
            gUnpackTime tmXMidClf.iStartTime(0), tmXMidClf.iStartTime(1), "A", "1", slLnStart
            gUnpackTime tmXMidClf.iEndTime(0), tmXMidClf.iEndTime(1), "A", "1", slLnEnd
            llTBStartTime(ilTBIndex) = CLng(gTimeToCurrency(slLnStart, False))
            llTBEndTime(ilTBIndex) = CLng(gTimeToCurrency(slLnEnd, True))
            If llTBStartTime(ilTBIndex) = llTBEndTime(ilTBIndex) Then
                llTBEndTime(ilTBIndex) = llTBStartTime(ilTBIndex) + 1
            End If
            If (tlCff(ilCffIndex).sDyWk = "D") Then  'Daily- Test if valid day
                mSetSepValues slLnStart, slLnEnd, llLength, ilTHour(), ilTQH(), ilAHour(), ilAQH()
            Else
                For ilIndex = ilFirstDay To ilLastDay Step 1
                    If (tlCff(ilCffIndex).iDay(ilIndex) = 1) Or (tlCff(ilCffIndex).sXDay(ilIndex) = "1") Then
                        mSetSepValues slLnStart, slLnEnd, llLength, ilTHour(), ilTQH(), ilAHour(), ilAQH()
                    End If
                Next ilIndex
            End If
        End If
    End If
    If (tlCff(ilCffIndex).sDyWk <> "D") Then  'Weekly
        llFreedom = llLength
    Else
        'llFreedom = 0
        'For ilIndex = 0 To 6 Step 1
        '    If tlCff(ilCffIndex).iDay(ilIndex) > 0 Then
        '        llFreedom = llFreedom + llLength
        '    End If
        'Next ilIndex
        llFreedom = llLength
    End If
    If ilBkQH = 0 Then
        If ((llFreedom Mod 900) <> 0) Or (llFreedom = 0) Then
            ilBkQH = CInt(llFreedom \ 900) + 1
        Else
            ilBkQH = CInt(llFreedom \ 900)
        End If
        If ilBkQH <= 0 Then
            ilBkQH = 1
        End If
        If ilBkQH > 1000 Then
            ilBkQH = 1000
        End If
    End If
    'If last week, then apply a 75% rule (ilBkQH * .75)
    If (tgSaf(0).iWk1stSoloIndex > 0) And (tgSaf(0).iWk1stSoloIndex < 100) Then
        If (llDate >= llClfMoLastWkDate) Then
            If ilBkQH <= 1000 Then
                ilBkQH = (CLng(ilBkQH) * CLng(tgSaf(0).iWk1stSoloIndex)) / 100
            Else
                ilBkQH = ilBkQH - 1
            End If
        End If
        If (tlClf.iPosition = 1) Or (tlClf.sSoloAvail = "Y") Then
            If ilBkQH <= 1000 Then
                ilBkQH = (CLng(ilBkQH) * CLng(tgSaf(0).iWk1stSoloIndex)) / 100
            Else
                ilBkQH = ilBkQH - 1
            End If
        End If
        If ilBkQH <= 0 Then
            ilBkQH = 1
        End If
    End If

    'llLength = (llLength / (ilSpots * 2))
    ilOtherSpots = mGetOtherLineSpots(llDate, tmXMidClf, tlCff(ilCffIndex), ilVpfIndex, tmXMidRdf)
    llLength = (llLength / ((ilSpots + ilOtherSpots) * 2))
    If llLength >= 3600& Then
        llSepLength = 3599
    ElseIf llLength < 300& Then
        llSepLength = 0
    Else
        llSepLength = llLength - 1
    End If
    '4/12/14: Fill does not use the Hour, Day and Skip information
    If Not blSetSkip Then
        Exit Sub
    End If
    '2/20/11:  Moved up
    ''Turn skip hours/qh/days off if no avail exist
    'For ilHour = LBound(ilSkip, 1) To UBound(ilSkip, 1) Step 1
    '    For ilQH = LBound(ilSkip, 2) To UBound(ilSkip, 2) Step 1
    '        For ilDay = LBound(ilSkip, 3) To UBound(ilSkip, 3) Step 1
    '            ilSkip(ilHour, ilQH, ilDay) = -1
    '        Next ilDay
    '    Next ilQH
    'Next ilHour
    For llLoopDate = llStartDate To llEndDate Step 1
        ilLoopDay = gWeekDayLong(llLoopDate)
        slTime = "12:00AM"
        ilDayOk = False
        If (tlCff(ilCffIndex).sDyWk <> "D") Then  'Weekly
            If (tlCff(ilCffIndex).iDay(ilLoopDay) = 1) Or (tlCff(ilCffIndex).sXDay(ilLoopDay) = "1") Then
                ilDayOk = True
            End If
        Else
            If (tlCff(ilCffIndex).iDay(ilLoopDay) > 0) Then
                ilDayOk = True
            End If
        End If
        If ilDayOk Then
            If gObtainSsfForDateOrGame(ilVefCode, llLoopDate, slTime, ilGameNo, hlSsf, tlSsf(ilLoopDay), llSsfDate(ilLoopDay), llSsfRecPos(ilLoopDay)) Then
                For ilLoop = 1 To tlSsf(ilLoopDay).iCount Step 1
                   LSet tmAvail = tlSsf(ilLoopDay).tPas(ADJSSFPASBZ + ilLoop)
                    'If ((tmAvail.iRecType = 2) And (slType <> "M") And (slType <> "S")) Or ((tmAvail.iRecType = 8) And (slType = "S")) Or ((tmAvail.iRecType = 9) And (slType = "M")) Then
                    ilAvailOk = False
                    If tmAvail.iRecType = 2 Then
                        If (slType <> "M") And (slType <> "S") Then
                            ilAvailOk = True
                        Else
                            If slType = "M" Then
                                If tgSpf.sSchdPromo = "Y" Then
                                    If ((Asc(tgSpf.sUsingFeatures3) And PROMOINTOCONTRACTAVAILS) = PROMOINTOCONTRACTAVAILS) Then
                                        ilAvailOk = True
                                    End If
                                End If
                            End If
                            If slType = "S" Then
                                If tgSpf.sSchdPSA = "Y" Then
                                    If ((Asc(tgSpf.sUsingFeatures3) And PSAINTOCONTRACTAVAILS) = PSAINTOCONTRACTAVAILS) Then
                                        ilAvailOk = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If (ilAvailOk) Or ((tmAvail.iRecType = 8) And (slType = "S")) Or ((tmAvail.iRecType = 9) And (slType = "M")) Or ((tmAvail.iRecType = 2) And (slType = "M") And (tgSpf.sSchdPromo <> "Y")) Or ((tmAvail.iRecType = 2) And (slType = "S") And (tgSpf.sSchdPSA <> "Y")) Then
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                        For ilIndex = LBound(llTBStartTime) To UBound(llTBStartTime) Step 1
                            If llTBStartTime(ilIndex) = -1 Then
                                Exit For
                            End If
                            If (llTime >= llTBStartTime(ilIndex)) And (llTime < llTBEndTime(ilIndex)) And (ilTBDays(ilIndex) = ilLoopDay) Then
                                'Turn hour, qh, day on
                                'ilQH = (tmAvail.iTime(1) And &HFF) \ 15 + 1 'Obtain quarter hour (1 to 4)
                                'ilHour = tmAvail.iTime(1) \ 256 + 1 'Obtain hour index 1-24
                                ilQH = (tmAvail.iTime(1) And &HFF) \ 15 'Obtain quarter hour (0 to 3)
                                ilHour = tmAvail.iTime(1) \ 256  'Obtain hour index 0-23
                                ilSkip(ilHour, ilQH, ilLoopDay) = 0
                                'Exit For
                            End If
                        Next ilIndex
                    End If
                Next ilLoop
            End If
        End If
    Next llLoopDate
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gMakeSmf                        *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create a Spot MG Spec from     *
'*                      missed spot                    *
'*                                                     *
'*******************************************************
Function gMakeSmf(hlSmf As Integer, tlSmf As SMF, slSchStatus As String, tlSdf As SDF, ilOrigVefCode As Integer, slMissedDate As String, slMissedTime As String, ilGameNo As Integer, slSchDate As String, slSchTime As String) As Integer
'
'   ilRet = gMakeSmf(hlSmf, tlSmf, slSchStatus, tlSdf, slMissedDate, slMissedTime, slSchDate, slSchTime)
'   Where:
'       hlSmf(I)- SMF open handle
'       tlSmf(I)- old SMF image if existed (tlSmf.lChfCode = 0 => didn't exist)
'                 old image must have been deleted previously
'       slSchStatus(I)- "G" for makgood, "O' for outside contract limits
'       llChfCode(I)- Contract code number
'       ilLineNo(I)- Contract line number
'       slMissedDate(I)- original spot date (missed or scheduled)
'       slMissedTime(I)- original spot time (missed or scheduled)
'       slSchDate(I)- schedule date
'       slSchTime(I)- schedule time
'
'
'       ilRet = True or False (i/o error)
'
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slSDate As String
    Dim slEDate As String
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim slSTime As String
    Dim slETime As String
    Dim llTime As Long
    Dim llSTime As Long
    Dim llETime As Long
    Dim ilDay As Integer
    If (slSchStatus <> "G") And (slSchStatus <> "O") Then
        gMakeSmf = True
        Exit Function
    End If
    imSmfRecLen = Len(tlSmf)
    'Determine if one required to be created or updated
    'If (tlSmf.lChfCode = tlSdf.lChfCode) And (tlSmf.iLineNo = tlSdf.iLineNo) Then
    'Contract test left because old code set lChfCode = 0 prior to adding lSdfCode
    If (tlSmf.lSdfCode = tlSdf.lCode) And (tlSmf.lChfCode = tlSdf.lChfCode) And (tlSmf.iLineNo = tlSdf.iLineNo) And (tlSmf.lFsfCode = tlSdf.lFsfCode) Then
        'Test if spot is within old specification- if so, retain specifications
        ilDay = gWeekDayStr(slSchDate)
        'If tlSmf.sWkDays(ilDay + 1) = "Y" Then
        If tlSmf.sWkDays(ilDay) = "Y" Then
            gUnpackDate tlSmf.iStartDate(0), tlSmf.iStartDate(1), slSDate
            gUnpackDate tlSmf.iEndDate(0), tlSmf.iEndDate(1), slEDate
            llDate = gDateValue(slSchDate)
            llSDate = gDateValue(slSDate)
            llEDate = gDateValue(slEDate)
            If (llDate >= llSDate) And (llDate <= llEDate) Then
                gUnpackTime tlSmf.iStartTime(0), tlSmf.iStartTime(1), "A", "1", slSTime
                gUnpackTime tlSmf.iEndTime(0), tlSmf.iEndTime(1), "A", "1", slETime
                llTime = CLng(gTimeToCurrency(slSchTime, False))
                llSTime = CLng(gTimeToCurrency(slSTime, False))
                llETime = CLng(gTimeToCurrency(slETime, True))
                If (llTime >= llSTime) And (llTime <= llETime) Then
                    gPackDate slSchDate, tlSmf.iActualDate(0), tlSmf.iActualDate(1)
                    gPackTime slSchTime, tlSmf.iActualTime(0), tlSmf.iActualTime(1)
                    tlSmf.sSchStatus = slSchStatus    'G=Makegoog; O=Outside contract limits
                    If tlSdf.iVefCode <> ilOrigVefCode Then
                        tlSmf.sMGSource = "A"     'A=Create by mouse move to a different vehicle
                    Else
                        tlSmf.sMGSource = "M"     'S=Created by MG Spec, M=Created by mouse move outside contract
                    End If
                    tlSmf.lCode = 0
                    tlSmf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
                    If tlSdf.sBill <> "Y" Then
                        tlSmf.sPtType = tlSdf.sPtType
                        tlSmf.lCopyCode = tlSdf.lCopyCode
                        tlSmf.iRotNo = tlSdf.iRotNo
                    End If
                    'tlSmf.lMtfCode = 0
                    If tlSdf.sTracer = "*" Then
                        tlSmf.lMtfCode = tlSdf.lSmfCode
                    End If
                    ilRet = btrInsert(hlSmf, tlSmf, imSmfRecLen, INDEXKEY0)
                    If ilRet <> BTRV_ERR_NONE Then
                        igBtrError = gConvertErrorCode(ilRet)
                        sgErrLoc = "gMakeSmf-Insert Smf(1)"
                        gMakeSmf = False
                    Else
                        gMakeSmf = True
                    End If
                    Exit Function
                End If
            End If
        End If
    End If
    'Create SMF record
    imSmfRecLen = Len(tlSmf)  'Get and save SMF record length
    tlSmf.lCode = 0
    tlSmf.lSdfCode = tlSdf.lCode
    tlSmf.lChfCode = tlSdf.lChfCode        'Contract code
    tlSmf.iLineNo = tlSdf.iLineNo      'Line number
    tlSmf.lFsfCode = tlSdf.lFsfCode
    gPackDate slMissedDate, tlSmf.iMissedDate(0), tlSmf.iMissedDate(1)
    gPackTime slMissedTime, tlSmf.iMissedTime(0), tlSmf.iMissedTime(1)
    gPackDate slSchDate, tlSmf.iStartDate(0), tlSmf.iStartDate(1)
    gPackDate slSchDate, tlSmf.iEndDate(0), tlSmf.iEndDate(1)
    gPackDate slSchDate, tlSmf.iActualDate(0), tlSmf.iActualDate(1)
    gPackTime slSchTime, tlSmf.iStartTime(0), tlSmf.iStartTime(1)
    gPackTime slSchTime, tlSmf.iEndTime(0), tlSmf.iEndTime(1)
    gPackTime slSchTime, tlSmf.iActualTime(0), tlSmf.iActualTime(1)
    For ilLoop = 1 To 7
        tlSmf.sWkDays(ilLoop - 1) = "N" 'Weekday flag: Y=day allowed; N=Day disallowed
    Next ilLoop
    'tlSmf.sWkDays(gWeekDayStr(slSchDate) + 1) = "Y" 'Weekday flag: Y=day allowed; N=Day disallowed
    tlSmf.sWkDays(gWeekDayStr(slSchDate)) = "Y" 'Weekday flag: Y=day allowed; N=Day disallowed
    If tlSdf.iVefCode <> ilOrigVefCode Then
        tlSmf.sMGSource = "A"     'A=Create by mouse move to a different vehicle
    Else
        tlSmf.sMGSource = "M"     'S=Created by MG Spec, M=Created by mouse move outside contract
    End If
    tlSmf.sSchStatus = slSchStatus    'G=Makegood; O=Outside contract limits
    tlSmf.iUrfCode = tgUrf(0).iCode 'Use first record retained for user
    tlSmf.iOrigSchVef = ilOrigVefCode
    'Retain copy assignment for as Order billing
    tlSmf.sPtType = tlSdf.sPtType
    tlSmf.lCopyCode = tlSdf.lCopyCode
    tlSmf.iRotNo = tlSdf.iRotNo
    If tlSdf.sTracer = "*" Then
        tlSmf.lMtfCode = tlSdf.lSmfCode
    Else
        tlSmf.lMtfCode = 0
    End If
    tlSmf.iGameNo = ilGameNo
    ilRet = btrInsert(hlSmf, tlSmf, imSmfRecLen, INDEXKEY0)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gMakeSmf-Insert Smf(2)"
        gMakeSmf = False
    Else
        gMakeSmf = True
    End If
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gMakeSSF                        *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Make SSF for specified date     *
'*                                                     *
'*******************************************************
Function gMakeSSF(ilTestOnly As Integer, hlSsf As Integer, hlSdf As Integer, hlSmf As Integer, ilType As Integer, ilVefCode As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, llGameAdjTime As Long, Optional llOrigDate As Long = 0) As Integer
'
'   ilRet = gMakeSSF(ilTestOnly, hlSsf, hlSdf, hlSmf, slType, ilVefCode, ilLogDate0, ilLogDate1, llGameAdjTime)
'   Where:
'       ilTestOnly(I)- True=Test SSF only; False=Test and Update
'       hlSsf (I)- SSF handle (obtained from CBtrvTable)
'       hlSdf (I)- SDF handle (obtained from CBtrvTable)
'       hlSmf (I)- SMF handle
'       slType (I)- "O" = On Air; "A" = Alternate
'       ilVefCode (I)- Vehicle code
'       ilLogDate0 (I)- Log date to be checked and created
'       ilLogDate1
'       llGameAdjTime(I)- Amount of time that avails shifted (Used for games only)
'
'       ilRet = True or False
'
'    Dim tlSSf As SSF                'SSF record image
    Dim tlSsfSrchKey As SSFKEY0     'SSF key record image
    Dim tlSsfSrchKey1 As SSFKEY1     'SSF key record image
    Dim tlSsfSrchKey2 As SSFKEY2     'SSF key record image
    Dim ilSsfRecLen As Integer      'SSF record length
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilBuildHdSSF As Integer
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim slTime As String
    Dim llTime As Long
    Dim ilMove As Integer
    Dim slDate As String
    Dim ilUpperBound As Integer
    Dim ilBoundIndex As Integer
    Dim ilRoomReq As Integer
    Dim ilAvailIndex As Integer
    Dim ilSpotIndex As Integer
    Dim ilSdfRecLen As Integer
    Dim ilSpotUnits As Integer
    Dim ilSpotLen As Integer
    Dim ilUnitsRem As Integer
    Dim ilLenRem As Integer
    Dim ilVpfIndex As Integer
    Dim ilVeh As Integer
    Dim ilEvt As Integer
    Dim ilSvCount As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilTime0 As Integer
    Dim ilTime1 As Integer
    Dim ilError As Integer
    Dim ilFound As Integer
    Dim ilOrigSchVef As Integer
    Dim ilOrigGameNo As Integer
    Dim slMsg As String
    Dim ilTestOk As Integer
    'Dim llSdfRecPos As Long
    Dim llChfRecPos As Long
    Dim llClfRecPos As Long
    Dim llCffRecPos As Long
    Dim ilMakeSpotMissed As Integer
    Dim ilRemoveSpot As Integer
    Dim ilUpdateSdf As Integer
    Dim ilMatchTime As Integer  'True= matching avail time found- advance ilAvailIndex
    Dim llCntrNo As Long
    Dim llChfCode As Long
    Dim ilAdfCode As Integer
    Dim llFsfCode As Long
    Dim ilSec As Integer
    Dim ilMin As Integer
    Dim llNowDate As Long
    Dim ilAffChg As Integer
    Dim slOrigSchStatus As String
    Dim slXSpotType As String
    ReDim ilEvtType(0 To 14) As Integer
    ReDim ilLen(0 To 1) As Integer
    Dim slAdvt As String
    Dim ilAdvtIndex As Integer
    Dim ilPrice As Integer
    Dim ilPriceLevel As Integer
    Dim ilSaf As Integer
    Dim llLLCAvailTime As Long
    Dim llSSFAvailTime As Long
    Dim llAlfStartTime As Long
    Dim llAlfEndTime As Long
    Dim slXMid As String
    Dim ilOrigSSFGameDate(0 To 1) As Integer
    Dim tlCff As CFF
    Dim ilNoSpotsThis As Integer
    Dim ilSplitNetworkPriRemoved As Integer
    Dim ilTestSplitNetworkLen As Integer
    Dim tlSpot As CSPOTSS
    Dim ilTest As Integer
    Dim tlAtt As ATT                'ATT record image
    Dim hlGsf As Integer
    Dim hlSxf As Integer
    Dim llSsfDate As Long


    ReDim sgSSFErrorMsg(0 To 0) As String
    imVefRecLen = Len(tmVef)  'Get and save VEF record length
    hmVef = CBtrvTable(ONEHANDLE)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet = BTRV_ERR_NONE Then
        tmVefSrchKey.iCode = ilVefCode
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            tmVef.sName = "Vehicle Missing"
        End If
    Else
        tmVef.sName = "Vehicle Missing"
    End If
    ilRet = btrClose(hmVef)
    btrDestroy hmVef
    imCHFRecLen = Len(tmChf)  'Get and save ADF record length
    hmCHF = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gMakeSSF-Open Chf(1)"
        btrDestroy hmCHF
        gMakeSSF = False
        Exit Function
    End If
    imClfRecLen = Len(tmClf)  'Get and save ADF record length
    hmClf = CBtrvTable(TWOHANDLES)        'Create ADF object handle
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gMakeSSF-Open Clf(2)"
        btrDestroy hmCHF
        btrDestroy hmClf
        gMakeSSF = False
        Exit Function
    End If
    imCffRecLen = Len(tmCff)  'Get and save ADF record length
    hmCff = CBtrvTable(TWOHANDLES)        'Create ADF object handle
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gMakeSSF-Open Cff(3)"
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmCff
        gMakeSSF = False
        Exit Function
    End If
    imFsfRecLen = Len(tmFsf)  'Get and save ADF record length
    hmFsf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gMakeSSF-Open Fsf(3)"
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmCff
        btrDestroy hmFsf
        gMakeSSF = False
        Exit Function
    End If
    imAlfRecLen = Len(tmAlf)  'Get and save ADF record length
    hmAlf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    ilRet = btrOpen(hmAlf, "", sgDBPath & "Alf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gMakeSSF-Open Alf(3)"
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmCff
        btrDestroy hmFsf
        btrDestroy hmAlf
        gMakeSSF = False
        Exit Function
    End If
    imAlfRecLen = Len(tmAlf)  'Get and save ADF record length
    hmAtt = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    ilRet = btrOpen(hmAtt, "", sgDBPath & "Att.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gMakeSSF-Open Att(3)"
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmCff
        btrDestroy hmFsf
        btrDestroy hmAlf
        btrDestroy hmAtt
        gMakeSSF = False
        Exit Function
    End If
    imAttRecLen = Len(tlAtt)  'Get and save ADF record length
    hmStf = CBtrvTable(TWOHANDLES)        'Create ADF object handle
    ilRet = btrOpen(hmStf, "", sgDBPath & "Stf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gMakeSSF-Open Att(3)"
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmCff
        btrDestroy hmFsf
        btrDestroy hmAlf
        btrDestroy hmAtt
        btrDestroy hmStf
        gMakeSSF = False
        Exit Function
    End If
    imStfRecLen = Len(tmStf)  'Get and save ADF record length
    hlGsf = CBtrvTable(TWOHANDLES)        'Create ADF object handle
    ilRet = btrOpen(hlGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gMakeSSF-Open Att(3)"
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmCff
        btrDestroy hmFsf
        btrDestroy hmAlf
        btrDestroy hmAtt
        btrDestroy hmStf
        btrDestroy hlGsf
        gMakeSSF = False
        Exit Function
    End If
    hlSxf = CBtrvTable(TWOHANDLES)        'Create ADF object handle
    ilRet = btrOpen(hlSxf, "", sgDBPath & "Sxf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gMakeSSF-Open Att(3)"
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmCff
        btrDestroy hmFsf
        btrDestroy hmAlf
        btrDestroy hmAtt
        btrDestroy hmStf
        btrDestroy hlGsf
        btrDestroy hlSxf
        gMakeSSF = False
        Exit Function
    End If
    gObtainPostLogAvailCode 'Set igPLAnfCode
    llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    ReDim tlLLC(0 To 0) As LLC  'Merged library names
    For ilIndex = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilIndex) = False
    Next ilIndex
    ilEvtType(1) = True 'Program
    'For ilIndex = 2 To 9 Step 1 'All avails
    ilEvtType(2) = True 'avail
    For ilIndex = 6 To 9 Step 1 'All avails
        ilEvtType(ilIndex) = True
    Next ilIndex
    gUnpackDate ilLogDate0, ilLogDate1, slDate
    ilRet = gBuildEventDay(ilType, "C", ilVefCode, slDate, "12M", "12M", ilEvtType(), tlLLC())
    If Not ilRet Then
        btrDestroy hmCHF
        btrDestroy hmClf
        btrDestroy hmCff
        btrDestroy hmFsf
        btrDestroy hmAlf
        btrDestroy hmAtt
        btrDestroy hmStf
        btrDestroy hlGsf
        btrDestroy hlSxf
        gMakeSSF = False
        Exit Function
    End If
    ilVpfIndex = -1
    'For ilVeh = 0 To UBound(tgVpf) Step 1
    '    If ilVefCode = tgVpf(ilVeh).iVefKCode Then
        ilVeh = gBinarySearchVpf(ilVefCode)
        If ilVeh <> -1 Then
            ilVpfIndex = ilVeh
    '        Exit For
        End If
    'Next ilVeh
    ilOrigSSFGameDate(0) = -1
    ilOrigSSFGameDate(1) = -1
    ilAvailIndex = 1
    ilBoundIndex = LBound(tgSsf)
    ilSdfRecLen = Len(tmSdf)
    ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
    ilUpperBound = LBound(tgSsf)
    If ilType <= 0 Then
        tlSsfSrchKey.iType = ilType
        tlSsfSrchKey.iVefCode = ilVefCode
        tlSsfSrchKey.iDate(0) = ilLogDate0
        tlSsfSrchKey.iDate(1) = ilLogDate1
        tlSsfSrchKey.iStartTime(0) = 0
        tlSsfSrchKey.iStartTime(1) = 0
        ilRet = gSSFGetGreaterOrEqual(hlSsf, tgSsf(ilUpperBound), ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
        'If (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slType) And (tmSsf.iVefCode = ilVefCode) And (tmSsf.iDate(0) = ilLogDate0) And (tmSsf.iDate(1) = ilLogDate1) Then
            'Compare and add avail records as required- move or remove spots as required
        'End If
        Do While (ilRet = BTRV_ERR_NONE) And (tgSsf(ilUpperBound).iType = ilType) And (tgSsf(ilUpperBound).iVefCode = ilVefCode) And (tgSsf(ilUpperBound).iDate(0) = ilLogDate0) And (tgSsf(ilUpperBound).iDate(1) = ilLogDate1)
            ilRet = gSSFGetPosition(hlSsf, lgSsfRecPos(ilUpperBound))
            ilUpperBound = ilUpperBound + 1
            ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
            ilRet = gSSFGetNext(hlSsf, tgSsf(ilUpperBound), ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Else
        'tlSsfSrchKey1.iType = ilType
        'tlSsfSrchKey1.iVefCode = ilVefCode
        'ilRet = btrGetEqual(hlSsf, tgSsf(ilUpperBound), ilSsfRecLen, tlSsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        'If ilRet = BTRV_ERR_NONE Then
        '    ilRet = gSSFGetPosition(hlSsf, lgSsfRecPos(ilUpperBound))
        '    ilOrigSSFGameDate(0) = tgSsf(ilUpperBound).iDate(0)
        '    ilOrigSSFGameDate(1) = tgSsf(ilUpperBound).iDate(1)
        '    ilUpperBound = ilUpperBound + 1
        'End If
        If llOrigDate > 0 Then
            tlSsfSrchKey2.iVefCode = ilVefCode
            gPackDateLong llOrigDate, tlSsfSrchKey2.iDate(0), tlSsfSrchKey2.iDate(1)
        Else
            tlSsfSrchKey2.iVefCode = ilVefCode
            tlSsfSrchKey2.iDate(0) = ilLogDate0
            tlSsfSrchKey2.iDate(1) = ilLogDate1
        End If
        ilRet = gSSFGetEqualKey2(hlSsf, tgSsf(ilUpperBound), ilSsfRecLen, tlSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
        Do While (ilRet = BTRV_ERR_NONE) And (tgSsf(ilUpperBound).iVefCode = ilVefCode)
            gUnpackDateLong tgSsf(ilUpperBound).iDate(0), tgSsf(ilUpperBound).iDate(1), llSsfDate
            If ((llOrigDate > 0) And (llOrigDate = llSsfDate)) Or ((llOrigDate = 0) And (gDateValue(slDate) = llSsfDate)) Then
                If ilType = tgSsf(ilUpperBound).iType Then
                    ilRet = gSSFGetPosition(hlSsf, lgSsfRecPos(ilUpperBound))
                    ilOrigSSFGameDate(0) = tgSsf(ilUpperBound).iDate(0)
                    ilOrigSSFGameDate(1) = tgSsf(ilUpperBound).iDate(1)
                    ilUpperBound = ilUpperBound + 1
                    Exit Do
                End If
            Else
                Exit Do
            End If
            ilRet = gSSFGetNext(hlSsf, tgSsf(ilUpperBound), ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    ilUpperBound = ilUpperBound - 1
    'Build into LLC any Post Log Avails
    For ilLoop = LBound(tgSsf) To ilUpperBound Step 1
        ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
        ilRet = gSSFGetDirect(hlSsf, tgSsf(ilLoop), ilSsfRecLen, lgSsfRecPos(ilLoop), INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            igBtrError = gConvertErrorCode(ilRet)
            sgErrLoc = "gMakeSSF-Get Direct Ssf(4)"
            btrDestroy hmCHF
            btrDestroy hmClf
            btrDestroy hmCff
            btrDestroy hmFsf
            btrDestroy hmAlf
            btrDestroy hmAtt
            btrDestroy hmStf
            btrDestroy hlGsf
            btrDestroy hlSxf
            gMakeSSF = False
            Exit Function
        End If
        ilEvt = 1
        Do While ilEvt <= tgSsf(ilLoop).iCount
           LSet tmAvail = tgSsf(ilLoop).tPas(ADJSSFPASBZ + ilEvt)
            If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Avail
                If tmAvail.ianfCode = igPLAnfCode Then
                    ilFound = False
                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                    For ilIndex = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                        If gTimeToCurrency(tlLLC(ilIndex).sStartTime, False) > llTime Then
                            ilFound = True
                            For ilMove = UBound(tlLLC) - 1 To ilIndex Step -1
                                tlLLC(ilMove + 1).iDay = tlLLC(ilMove).iDay
                                tlLLC(ilMove + 1).sType = tlLLC(ilMove).sType
                                tlLLC(ilMove + 1).sStartTime = tlLLC(ilMove).sStartTime
                                tlLLC(ilMove + 1).sLength = tlLLC(ilMove).sLength
                                tlLLC(ilMove + 1).iUnits = tlLLC(ilMove).iUnits
                                tlLLC(ilMove + 1).sName = tlLLC(ilMove).sName
                                tlLLC(ilMove + 1).lLvfCode = tlLLC(ilMove).lLvfCode
                                tlLLC(ilMove + 1).iLtfCode = tlLLC(ilMove).iLtfCode
                                tlLLC(ilMove + 1).iAvailInfo = tlLLC(ilMove).iAvailInfo
                                tlLLC(ilMove + 1).iEtfCode = tlLLC(ilMove).iEtfCode
                                tlLLC(ilMove + 1).iEnfCode = tlLLC(ilMove).iEnfCode
                                tlLLC(ilMove + 1).lCefCode = tlLLC(ilMove).lCefCode
                                tlLLC(ilMove + 1).lEvtIDCefCode = tlLLC(ilMove).lEvtIDCefCode
                            Next ilMove
                            ilMove = ilIndex
                            Exit For
                        End If
                    Next ilIndex
                    If Not ilFound Then
                        ilIndex = UBound(tlLLC)
                    Else
                        ilIndex = ilMove
                    End If
                    tlLLC(ilIndex).sType = Trim$(str$(tmAvail.iRecType))
                    gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", tlLLC(ilIndex).sStartTime
                    tlLLC(ilIndex).iLtfCode = tmAvail.iLtfCode
                    tlLLC(ilIndex).iUnits = tmAvail.iAvInfo And &H1F
                    tlLLC(ilIndex).iAvailInfo = tmAvail.iAvInfo And (Not &H1F)
                    ilSec = tmAvail.iLen
                    ilMin = 0
                    Do While ilSec > 59
                        ilMin = ilMin + 1
                        ilSec = ilSec - 60
                    Loop
                    ilLen(0) = 256 * ilSec
                    ilLen(1) = ilMin
                    gUnpackLength ilLen(0), ilLen(1), "3", True, tlLLC(ilIndex).sLength
                    tlLLC(ilIndex).sName = Trim$(str$(tmAvail.ianfCode))
                    ReDim Preserve tlLLC(LBound(tlLLC) To UBound(tlLLC) + 1) As LLC
                End If
            End If
            ilEvt = ilEvt + 1
        Loop
    Next ilLoop
    If Not ilTestOnly Then
        For ilLoop = LBound(tgSsf) To ilUpperBound Step 1
            Do
                ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetDirect(hlSsf, tgSsf(ilLoop), ilSsfRecLen, lgSsfRecPos(ilLoop), INDEXKEY0, BTRV_LOCK_NONE)
                ilRet = gGetByKeyForUpdateSSF(hlSsf, tgSsf(ilLoop))
                If ilRet <> BTRV_ERR_NONE Then
                    igBtrError = gConvertErrorCode(ilRet)
                    sgErrLoc = "gMakeSSF-Get By Key Ssf(4)"
                    btrDestroy hmCHF
                    btrDestroy hmClf
                    btrDestroy hmCff
                    btrDestroy hmFsf
                    btrDestroy hmAlf
                    btrDestroy hmAtt
                    btrDestroy hmStf
                    btrDestroy hlGsf
                    btrDestroy hlSxf
                    gMakeSSF = False
                    Exit Function
                End If
                ilRet = btrDelete(hlSsf)
            Loop While ilRet = BTRV_ERR_CONFLICT
        Next ilLoop
    End If
    If (ilUpperBound < LBound(tgSsf)) And (LBound(tlLLC) <= UBound(tlLLC) - 1) Then
        If ilTestOnly Then
            slMsg = "Date Missing: "
        Else
            slMsg = "Date Missing, Added: "
        End If
        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & Trim$(tmVef.sName)
        ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
    End If
    If (ilUpperBound >= LBound(tgSsf)) And (LBound(tlLLC) > UBound(tlLLC) - 1) Then
        If ilTestOnly Then
            slMsg = "Date Extra: "
        Else
            slMsg = "Date Extra, Removed: "
        End If
        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & Trim$(tmVef.sName)
        ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
    End If
    ilBuildHdSSF = False
    tmSsf.iType = ilType
    tmSsf.iVefCode = ilVefCode
    tmSsf.iDate(0) = ilLogDate0
    tmSsf.iDate(1) = ilLogDate1
    tmSsf.iStartTime(0) = 0
    tmSsf.iStartTime(1) = 0
    tmSsf.iCount = 0
    'tmSsf.iNextTime(0) = 1  'Time not defined
    'tmSsf.iNextTime(1) = 0
    For ilIndex = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
        Do
            If ilBuildHdSSF Then
                'gPackTime tlLLC(ilIndex).sStartTime, tmSsf.iNextTime(0), tmSsf.iNextTime(1)
                ilSsfRecLen = igSSFBaseLen + tmSsf.iCount * Len(tmAvail)
                If Not ilTestOnly Then
                    tmSsf.lCode = 0
                    ilRet = gSSFInsert(hlSsf, tmSsf, ilSsfRecLen, INDEXKEY3)
                    If ilRet <> BTRV_ERR_NONE Then
                        igBtrError = gConvertErrorCode(ilRet)
                        sgErrLoc = "gMakeSSF-Insert Ssf(5)"
                        btrDestroy hmCHF
                        btrDestroy hmClf
                        btrDestroy hmCff
                        btrDestroy hmFsf
                        btrDestroy hmAlf
                        btrDestroy hmAtt
                        btrDestroy hmStf
                        btrDestroy hlGsf
                        btrDestroy hlSxf
                        gMakeSSF = False
                        Exit Function
                    End If
                End If
                ilBuildHdSSF = False
                'tmSsf.iType = ilType
                'tmSsf.iVefCode = ilVefCode
                'tmSsf.iDate(0) = ilLogDate0
                'tmSsf.iDate(1) = ilLogDate1
                'gPackTime tlLLC(ilIndex).sStartTime, tmSsf.iStartTime(0), tmSsf.iStartTime(1)
                tmSsf.iCount = 0
                'tmSsf.iNextTime(0) = 1  'Time not defined
                'tmSsf.iNextTime(1) = 0
                Exit Do
            End If
            If tlLLC(ilIndex).sType = "1" Then
                tmProg.iRecType = 1
                gPackTime tlLLC(ilIndex).sStartTime, tmProg.iStartTime(0), tmProg.iStartTime(1)
                gAddTimeLength tlLLC(ilIndex).sStartTime, tlLLC(ilIndex).sLength, "A", "1", slTime, slXMid
                gPackTime slTime, tmProg.iEndTime(0), tmProg.iEndTime(1)
                'Program exclusions
                tmProg.iMnfExcl(0) = tlLLC(ilIndex).iUnits
                tmProg.iMnfExcl(1) = tlLLC(ilIndex).iAvailInfo
                tmProg.lLvfCode = tlLLC(ilIndex).lLvfCode
                tmProg.iLtfCode = tlLLC(ilIndex).iLtfCode
                If tmSsf.iCount + 1 <= UBound(tmSsf.tPas) Then
                    tmSsf.iCount = tmSsf.iCount + 1
                    LSet tmSsf.tPas(ADJSSFPASBZ + tmSsf.iCount) = tmProg
                Else
                    ilBuildHdSSF = True
                End If
            Else    'Avail
                tmAvail.iRecType = Val(tlLLC(ilIndex).sType)
                gPackTime tlLLC(ilIndex).sStartTime, tmAvail.iTime(0), tmAvail.iTime(1)
                llLLCAvailTime = gTimeToLong(tlLLC(ilIndex).sStartTime, False)
                tmAvail.iLtfCode = tlLLC(ilIndex).iLtfCode
                tmAvail.iAvInfo = tlLLC(ilIndex).iAvailInfo Or tlLLC(ilIndex).iUnits
                '4/3/06- Set locks from ALF, later they will be set from tgAvail (this is required as not all avail locks stored into alf)
                If ilType <= 0 Then
                    tmAlfSrchkey1.iVefCode = ilVefCode
                    tmAlfSrchkey1.iDate(0) = ilLogDate0
                    tmAlfSrchkey1.iDate(1) = ilLogDate1
                    ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchkey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                Else
                    tmAlfSrchKey2.iVefCode = ilVefCode
                    tmAlfSrchKey2.iGameNo = ilType
                    ilRet = btrGetEqual(hmAlf, tmAlf, imAlfRecLen, tmAlfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                End If
                Do While ilRet = BTRV_ERR_NONE
                    If tmAlf.iVefCode <> ilVefCode Then
                        Exit Do
                    End If
                    If ilType <= 0 Then
                        If (tmAlf.iDate(0) <> ilLogDate0) Or (tmAlf.iDate(1) <> ilLogDate1) Then
                            Exit Do
                        End If
                    Else
                        If tmAlf.iGameNo <> ilType Then
                            Exit Do
                        End If
                    End If
                    gUnpackTimeLong tmAlf.iStartTime(0), tmAlf.iStartTime(1), False, llAlfStartTime
                    gUnpackTimeLong tmAlf.iEndTime(0), tmAlf.iEndTime(1), True, llAlfEndTime
                    If (llLLCAvailTime >= llAlfStartTime) And (llLLCAvailTime <= llAlfEndTime) Then
                        If tmAlf.sLockType = "S" Then
                            tmAvail.iAvInfo = (tmAvail.iAvInfo) Or (tgAvail.iAvInfo And SSLOCKSPOT)
                        ElseIf tmAlf.sLockType = "A" Then
                            tmAvail.iAvInfo = (tmAvail.iAvInfo) Or (tgAvail.iAvInfo And SSLOCK)
                        End If
                    End If
                    ilRet = btrGetNext(hmAlf, tmAlf, imAlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If (tmAvail.iAvInfo And SSXMID) = SSXMID Then
                    llLLCAvailTime = llLLCAvailTime + 86400
                End If
                tmAvail.iLen = CInt(gLengthToCurrency(tlLLC(ilIndex).sLength))
                tmAvail.ianfCode = Val(tlLLC(ilIndex).sName)
                tmAvail.iNoSpotsThis = 0
                tmAvail.iOrigUnit = 0
                tmAvail.iOrigLen = 0
                'Check for matching avail and if so, check for room
                ilRoomReq = 1
                ilMatchTime = False
                If (ilUpperBound >= LBound(tgSsf)) And (ilBoundIndex <= ilUpperBound) Then
                    Do
                        If ilAvailIndex > tgSsf(ilBoundIndex).iCount Then
                            ilBoundIndex = ilBoundIndex + 1
                            If (ilBoundIndex > ilUpperBound) Then
                                Exit Do
                            End If
                            ilAvailIndex = 1
                        End If
                        tgAvail = tgSsf(ilBoundIndex).tPas(ADJSSFPASBZ + ilAvailIndex)
                        ilAffChg = False
                        If (tgAvail.iRecType >= 2) And (tgAvail.iRecType <= 9) Then
                            'If (tmAvail.iTime(0) = tgAvail.iTime(0)) And (tmAvail.iTime(1) = tgAvail.iTime(1)) Then
                            gUnpackTimeLong tgAvail.iTime(0), tgAvail.iTime(1), False, llSSFAvailTime
                            If (tgAvail.iAvInfo And SSXMID) = SSXMID Then
                                llSSFAvailTime = llSSFAvailTime + 86400
                            End If
                            If llSSFAvailTime + llGameAdjTime = llLLCAvailTime Then
                                'Test if units or seconds max out and in past- then assume spots added via Post Log
                                'and adjust units/seconds
                                ilNoSpotsThis = 0
                                For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tgAvail.iNoSpotsThis Step 1
                                    LSet tgSpot = tgSsf(ilBoundIndex).tPas(ADJSSFPASBZ + ilSpotIndex)
                                    If (tgSpot.iRecType And &HF) >= 10 Then
                                        If (tgSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                            ilNoSpotsThis = ilNoSpotsThis + 1
                                        End If
                                    End If
                                Next ilSpotIndex
                                If (tmAvail.iLen < tgAvail.iLen) Or ((tmAvail.iAvInfo And &H1F) < (tgAvail.iAvInfo And &H1F)) Then
                                    If (gDateValue(slDate) < llNowDate) Or ((tgAvail.iOrigUnit > 0) Or (tgAvail.iOrigLen > 0)) Then
                                        ilLenRem = tgAvail.iLen
                                        For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tgAvail.iNoSpotsThis Step 1
                                            LSet tgSpot = tgSsf(ilBoundIndex).tPas(ADJSSFPASBZ + ilSpotIndex)
                                            If (tgSpot.iRecType And &HF) >= 10 Then
                                                If (tgSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                                    ilSpotLen = (tgSpot.iPosLen And &HFFF)
                                                    ilLenRem = ilLenRem - ilSpotLen
                                                End If
                                            End If
                                            tmSdfSrchKey3.lCode = tgSpot.lSdfCode
                                            ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                            If (ilRet = BTRV_ERR_NONE) Then
                                                If (Trim$(tmSdf.sAffChg) <> "") And (tmSdf.sAffChg <> "N") Then
                                                    ilAffChg = True
                                                End If
                                            End If
                                        Next ilSpotIndex
                                        If ((tgAvail.iAvInfo And &H1F) = ilNoSpotsThis) Or (ilLenRem = 0) Or (ilAffChg) Then
                                            'Adjust units and length
                                            tmAvail.iLen = tgAvail.iLen
                                            tmAvail.iAvInfo = (tmAvail.iAvInfo And (Not &H1F)) Or (tgAvail.iAvInfo And &H1F)
                                        End If
                                    End If
                                ElseIf (tgVpf(ilVpfIndex).sSSellOut = "M") Then
                                    If (gDateValue(slDate) < llNowDate) Or ((tgAvail.iOrigUnit > 0) Or (tgAvail.iOrigLen > 0)) Then
                                        For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tgAvail.iNoSpotsThis Step 1
                                            LSet tgSpot = tgSsf(ilBoundIndex).tPas(ADJSSFPASBZ + ilSpotIndex)
                                            tmSdfSrchKey3.lCode = tgSpot.lSdfCode
                                            ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                            If (ilRet = BTRV_ERR_NONE) Then
                                                If (Trim$(tmSdf.sAffChg) <> "") And (tmSdf.sAffChg <> "N") Then
                                                    ilAffChg = True
                                                End If
                                            End If
                                        Next ilSpotIndex
                                    End If
                                End If
                                'Match found
                                ilRoomReq = 1 + tgAvail.iNoSpotsThis
                                ilMatchTime = True
                                tmAvail.iAvInfo = (tmAvail.iAvInfo) Or (tgAvail.iAvInfo And SSLOCK) Or (tgAvail.iAvInfo And SSLOCKSPOT)
                                Exit Do
                            End If
                            'Add alert here if using the affiliate system
                            If tgSpf.sGUseAffSys = "Y" Then
                                tmATTSrchKey1.iCode = ilVefCode
                                ilRet = btrGetEqual(hmAtt, tlAtt, imAttRecLen, tmATTSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                                If ilRet = BTRV_ERR_NONE Then
                                    ilRet = gAlertAdd("P", "A", 0, ilVefCode, slDate)
                                End If
                            End If
                            If (tgAvail.iTime(1) > tmAvail.iTime(1)) Or ((tgAvail.iTime(1) = tmAvail.iTime(1)) And (tgAvail.iTime(0) > tmAvail.iTime(0))) Then
                                'Add avail- not matching times (next avail is later in time)
                                Exit Do
                            End If
                            'If tgAvail.iAnfCode = igPLAnfCode Then
                            '    ilRoomReq = 1 + tgAvail.iNoSpotsThis
                            '    ilMatchTime = True
                            '   LSet tmAvail = tgAvail
                            '    ilIndex = ilIndex - 1
                            '    Exit Do
                            'End If
                            'No matching avail- set spots to missed
                            gUnpackTime tgAvail.iTime(0), tgAvail.iTime(1), "A", "1", slTime
                            For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tgAvail.iNoSpotsThis Step 1
                                LSet tgSpot = tgSsf(ilBoundIndex).tPas(ADJSSFPASBZ + ilSpotIndex)
                                If (tgSpot.iRecType And &HF) >= 10 Then
                                    ilError = False
                                    On Error GoTo gMakeSSFErr
                                    tmSdfSrchKey3.lCode = tgSpot.lSdfCode
                                    ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                    If (Not ilError) And (ilRet = BTRV_ERR_NONE) Then
                                        slOrigSchStatus = tmSdf.sSchStatus
                                        'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                                        ilTestSplitNetworkLen = False
                                        '6/4/16: Replaced GoSub
                                        'GoSub TestSDF
                                        mTestSDF ilTestOk, ilUpdateSdf, ilRemoveSpot, ilAdfCode, llChfCode, llFsfCode, llCntrNo, ilSplitNetworkPriRemoved, ilTestSplitNetworkLen, ilBoundIndex, ilAvailIndex, ilSpotIndex, ilTestOnly, hlGsf, hlSxf, hlSdf, hlSmf, ilSdfRecLen, slDate, slTime, slMsg
                                        ilSplitNetworkPriRemoved = False
                                        If Not ilTestOk Then
                                            btrDestroy hmCHF
                                            btrDestroy hmClf
                                            btrDestroy hmCff
                                            btrDestroy hmFsf
                                            btrDestroy hmAlf
                                            btrDestroy hmAtt
                                            btrDestroy hmStf
                                            btrDestroy hlGsf
                                            btrDestroy hlSxf
                                            gMakeSSF = False
                                            Exit Function
                                        End If
                                        If Not ilRemoveSpot Then
                                            If Not ilTestOnly Then
                                                ilRet = gMakeTracer(hlSdf, tmSdf, 0, hmStf, -1, "M", "P", tmSdf.iRotNo, hlGsf)
                                                ilRet = gRemoveSmf(hlSmf, tmSmf, tmSdf, hlSxf)   'resets missed date
                                                If Not ilRet Then
                                                    btrDestroy hmCHF
                                                    btrDestroy hmClf
                                                    btrDestroy hmCff
                                                    btrDestroy hmFsf
                                                    btrDestroy hmAlf
                                                    btrDestroy hmAtt
                                                    btrDestroy hmStf
                                                    btrDestroy hlGsf
                                                    btrDestroy hlSxf
                                                    gMakeSSF = False
                                                    Exit Function
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
                                                slXSpotType = ""
                                            End If
                                            'If (tmSdf.sSpotType = "T") Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "S") Or (tmSdf.sSpotType = "M") Or (slXSpotType = "X") Then
                                            If ((tmSdf.sSpotType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmSdf.sSpotType = "Q") Or ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (slXSpotType = "X") Then
                                                'tmChfSrchKey.lCode = tmSdf.lChfCode
                                                'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If ilTestOnly Then
                                                    slMsg = "No Avail: "
                                                Else
                                                    slMsg = "No Avail, Removed: "
                                                End If
                                                'If ilRet = BTRV_ERR_NONE Then
                                                    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                                'Else
                                                '    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " ChfCode=" & Str$(tmSdf.lChfCode) & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                                'End If
                                                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                                If Not ilTestOnly Then
                                                    'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                                                    Do
                                                        tmSdfSrchKey3.lCode = tmSdf.lCode
                                                        ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                        'tmSRec = tmSdf
                                                        'ilCRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                                                        'tmSdf = tmSRec
                                                        'If ilCRet <> BTRV_ERR_NONE Then
                                                        '    igBtrError = ilCRet
                                                        '    sgErrLoc = "gMakeSSF-Get By Key Sdf(6)"
                                                        '    btrDestroy hmChf
                                                        '    btrDestroy hmClf
                                                        '    btrDestroy hmCff
                                                        '    gMakeSSF = False
                                                        '    Exit Function
                                                        'End If
                                                        ilRet = btrDelete(hlSdf)
                                                        'If ilRet = BTRV_ERR_CONFLICT Then
                                                        '    ilCRet = btrGetDirect(hlSdf, tmSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                        'End If
                                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                                End If
                                            Else
                                                ilDate0 = tmSdf.iDate(0)
                                                ilDate1 = tmSdf.iDate(1)
                                                ilTime0 = tmSdf.iTime(0)
                                                ilTime1 = tmSdf.iTime(1)
                                                ilOrigSchVef = tmSdf.iVefCode
                                                ilOrigGameNo = tmSdf.iGameNo
                                                'tmChfSrchKey.lCode = tmSdf.lChfCode
                                                'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If ilTestOnly Then
                                                    slMsg = "No Avail: "
                                                Else
                                                    slMsg = "No Avail, Missed: "
                                                End If
                                                'If ilRet = BTRV_ERR_NONE Then
                                                    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                                'Else
                                                '    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " ChfCode=" & Str$(tmSdf.lChfCode) & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                                'End If
                                                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                                If Not ilTestOnly Then
                                                    'Update Sdf record
                                                    Do
                                                        'ilRet = btrGetDirect(hlSdf, tmSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                        tmSdfSrchKey3.lCode = tmSdf.lCode
                                                        ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                        'tmSRec = tmSdf
                                                        'ilRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                                                        'tmSdf = tmSRec
                                                        'If ilRet <> BTRV_ERR_NONE Then
                                                        '    igBtrError = gConvertErrorCode(ilRet)
                                                        '    sgErrLoc = "gMakeSSF-Get by Key Sdf(7)"
                                                        '    btrDestroy hmChf
                                                        '    btrDestroy hmClf
                                                        '    btrDestroy hmCff
                                                        '    gMakeSSF = False
                                                        '    Exit Function
                                                        'End If
                                                        tmSdf.lChfCode = llChfCode
                                                        tmSdf.iAdfCode = ilAdfCode
                                                        tmSdf.lFsfCode = llFsfCode
                                                        tmSdf.iLen = tmClf.iLen
                                                        tmSdf.sSchStatus = "M"
                                                        tmSdf.iMnfMissed = igMnfMissed
                                                        tmSdf.iDate(0) = ilDate0
                                                        tmSdf.iDate(1) = ilDate1
                                                        tmSdf.iTime(0) = ilTime0
                                                        tmSdf.iTime(1) = ilTime1
                                                        tmSdf.iVefCode = ilOrigSchVef
                                                        tmSdf.iGameNo = ilOrigGameNo
                                                        tmSdf.lSmfCode = 0
                                                        If ((tmSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                                                            tmSdf.sTracer = "*"
                                                            tmSdf.lSmfCode = tmSmf.lMtfCode
                                                        End If
                                                        If (ilType > 0) And (slOrigSchStatus <> "G") And (slOrigSchStatus <> "O") Then
                                                            tmSdf.iDate(0) = ilLogDate0
                                                            tmSdf.iDate(1) = ilLogDate1
                                                        End If
                                                        tmSdf.sXCrossMidnight = "N"
                                                        tmSdf.sWasMG = "N"
                                                        tmSdf.sFromWorkArea = "N"
                                                        tmSdf.iUrfCode = tgUrf(0).iCode
                                                        ilRet = btrUpdate(hlSdf, tmSdf, ilSdfRecLen)
                                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                                    lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                                                    'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                                                    ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                                                End If
                                            End If
                                        End If
                                    Else
                                        If ilTestOnly Then
                                            slMsg = "Can't Find Sdf: "
                                        Else
                                            slMsg = "Can't Find Sdf, Removed: "
                                        End If
                                        ilAdvtIndex = gBinarySearchAdf(tgSpot.iAdfCode)
                                        If ilAdvtIndex <> -1 Then
                                            If (tgCommAdf(ilAdvtIndex).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdvtIndex).sAddrID) <> "") Then
                                                slAdvt = Trim$(tgCommAdf(ilAdvtIndex).sName) & ", " & Trim$(tgCommAdf(ilAdvtIndex).sAddrID)
                                            Else
                                                slAdvt = Trim$(tgCommAdf(ilAdvtIndex).sName)
                                            End If
                                        Else
                                            slAdvt = "Missing" & str$(tgSpot.iAdfCode)
                                        End If
                                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName) & ", " & slAdvt
                                        ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                    End If
                                Else
                                    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Avail Count not matching # of Spots: " & slDate & " " & slTime & " " & Trim$(tmVef.sName)
                                    ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                End If
                            Next ilSpotIndex
                            ilAvailIndex = ilAvailIndex + tgAvail.iNoSpotsThis + 1
                        Else
                            'Test if spot outside of an avail
                            If ((tgAvail.iRecType And &HF) >= 10) And ((tgAvail.iRecType And &HF) <= 11) Then
                                'Spot outside avail- make spot missed
                                LSet tgSpot = tgSsf(ilBoundIndex).tPas(ADJSSFPASBZ + ilAvailIndex)
                                ilError = False
                                On Error GoTo gMakeSSFErr
                                tmSdfSrchKey3.lCode = tgSpot.lSdfCode
                                ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                If (Not ilError) And (ilRet = BTRV_ERR_NONE) Then
                                    slOrigSchStatus = tmSdf.sSchStatus
                                    'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                                    ilTestSplitNetworkLen = False
                                    '6/4/16: Replaced GoSub
                                    'GoSub TestSDF
                                    mTestSDF ilTestOk, ilUpdateSdf, ilRemoveSpot, ilAdfCode, llChfCode, llFsfCode, llCntrNo, ilSplitNetworkPriRemoved, ilTestSplitNetworkLen, ilBoundIndex, ilAvailIndex, ilSpotIndex, ilTestOnly, hlGsf, hlSxf, hlSdf, hlSmf, ilSdfRecLen, slDate, slTime, slMsg
                                    ilSplitNetworkPriRemoved = False
                                    If Not ilTestOk Then
                                        btrDestroy hmCHF
                                        btrDestroy hmClf
                                        btrDestroy hmCff
                                        btrDestroy hmFsf
                                        btrDestroy hmAlf
                                        btrDestroy hmAtt
                                        btrDestroy hmStf
                                        btrDestroy hlGsf
                                        btrDestroy hlSxf
                                        gMakeSSF = False
                                        Exit Function
                                    End If
                                    If Not ilRemoveSpot Then
                                        If Not ilTestOnly Then
                                            ilRet = gMakeTracer(hlSdf, tmSdf, 0, hmStf, -1, "M", "P", tmSdf.iRotNo, hlGsf)
                                            ilRet = gRemoveSmf(hlSmf, tmSmf, tmSdf, hlSxf)  'resets missed date
                                            If Not ilRet Then
                                                btrDestroy hmCHF
                                                btrDestroy hmClf
                                                btrDestroy hmCff
                                                btrDestroy hmFsf
                                                btrDestroy hmAlf
                                                btrDestroy hmAtt
                                                btrDestroy hmStf
                                                btrDestroy hlGsf
                                                btrDestroy hlSxf
                                                gMakeSSF = False
                                                Exit Function
                                            End If
                                        Else
                                            ilRet = gFindSmf(tmSdf, hlSmf, tmSmf)
                                        End If
                                        gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                                        'tmChfSrchKey.lCode = tmSdf.lChfCode
                                        'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                                        'If (tmChf.sType = "T") Or (tmChf.sType = "Q") Or (tmChf.sType = "S") Or (tmChf.sType = "M") Then
                                        If tmSdf.sSpotType = "X" Then
                                            slXSpotType = "X"
                                            If ((tmSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                                                slXSpotType = ""
                                            End If
                                        Else
                                            slXSpotType = ""
                                        End If
                                        'If (tmSdf.sSpotType = "T") Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "S") Or (tmSdf.sSpotType = "M") Or (slXSpotType = "X") Then
                                        If ((tmSdf.sSpotType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmSdf.sSpotType = "Q") Or ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (slXSpotType = "X") Then
                                            'tmChfSrchKey.lCode = tmSdf.lChfCode
                                            'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilTestOnly Then
                                                slMsg = "No Avail: "
                                            Else
                                                slMsg = "No Avail, Removed: "
                                            End If
                                            'If ilRet = BTRV_ERR_NONE Then
                                                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                            'Else
                                            '    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " ChfCode=" & Str$(tmSdf.lChfCode) & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                            'End If
                                            ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                            If Not ilTestOnly Then
                                                'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                                                Do
                                                    tmSdfSrchKey3.lCode = tmSdf.lCode
                                                    ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                    'tmSRec = tmSdf
                                                    'ilCRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                                                    'tmSdf = tmSRec
                                                    'If ilCRet <> BTRV_ERR_NONE Then
                                                    '    igBtrError = ilCRet
                                                    '    sgErrLoc = "gMakeSSF-Get by Key Sdf(8)"
                                                    '    btrDestroy hmChf
                                                    '    btrDestroy hmClf
                                                    '    btrDestroy hmCff
                                                    '    gMakeSSF = False
                                                    '    Exit Function
                                                    'End If
                                                    ilRet = btrDelete(hlSdf)
                                                    'If ilRet = BTRV_ERR_CONFLICT Then
                                                    '    ilCRet = btrGetDirect(hlSdf, tmSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                    'End If
                                                Loop While ilRet = BTRV_ERR_CONFLICT
                                            End If
                                        Else
                                            ilDate0 = tmSdf.iDate(0)
                                            ilDate1 = tmSdf.iDate(1)
                                            ilTime0 = tmSdf.iTime(0)
                                            ilTime1 = tmSdf.iTime(1)
                                            ilOrigSchVef = tmSdf.iVefCode
                                            ilOrigGameNo = tmSdf.iGameNo
                                            'tmChfSrchKey.lCode = tmSdf.lChfCode
                                            'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            If ilTestOnly Then
                                                slMsg = "No Avail: "
                                            Else
                                                slMsg = "No Avail, Missed: "
                                            End If
                                            'If ilRet = BTRV_ERR_NONE Then
                                                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                            'Else
                                            '    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " ChfCode=" & Str$(tmSdf.lChfCode) & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                            'End If
                                            ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                            If Not ilTestOnly Then
                                                'Update Sdf record
                                                Do
                                                    'ilRet = btrGetDirect(hlSdf, tmSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                    tmSdfSrchKey3.lCode = tmSdf.lCode
                                                    ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                    'tmSRec = tmSdf
                                                    'ilRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                                                    'tmSdf = tmSRec
                                                    'If ilRet <> BTRV_ERR_NONE Then
                                                    '    igBtrError = gConvertErrorCode(ilRet)
                                                    '    sgErrLoc = "gMakeSSF-Get by Key Sdf(9)"
                                                    '    btrDestroy hmChf
                                                    '    btrDestroy hmClf
                                                    '    btrDestroy hmCff
                                                    '    gMakeSSF = False
                                                    '    Exit Function
                                                    'End If
                                                    tmSdf.lChfCode = llChfCode
                                                    tmSdf.iAdfCode = ilAdfCode
                                                    tmSdf.lFsfCode = llFsfCode
                                                    tmSdf.iLen = tmClf.iLen
                                                    tmSdf.sSchStatus = "M"
                                                    tmSdf.iMnfMissed = igMnfMissed
                                                    tmSdf.iDate(0) = ilDate0
                                                    tmSdf.iDate(1) = ilDate1
                                                    tmSdf.iTime(0) = ilTime0
                                                    tmSdf.iTime(1) = ilTime1
                                                    tmSdf.iVefCode = ilOrigSchVef
                                                    tmSdf.iGameNo = ilOrigGameNo
                                                    tmSdf.lSmfCode = 0
                                                    If ((tmSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                                                        tmSdf.sTracer = "*"
                                                        tmSdf.lSmfCode = tmSmf.lMtfCode
                                                    End If
                                                    If (ilType > 0) And (slOrigSchStatus <> "G") And (slOrigSchStatus <> "O") Then
                                                        tmSdf.iDate(0) = ilLogDate0
                                                        tmSdf.iDate(1) = ilLogDate1
                                                    End If
                                                    tmSdf.sXCrossMidnight = "N"
                                                    tmSdf.sWasMG = "N"
                                                    tmSdf.sFromWorkArea = "N"
                                                    tmSdf.iUrfCode = tgUrf(0).iCode
                                                    ilRet = btrUpdate(hlSdf, tmSdf, ilSdfRecLen)
                                                Loop While ilRet = BTRV_ERR_CONFLICT
                                                lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                                                'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                                                ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                                            End If
                                        End If
                                    End If
                                Else
                                    If ilTestOnly Then
                                        slMsg = "Can't Find Sdf: "
                                    Else
                                        slMsg = "Can't Find Sdf, Removed: "
                                    End If
                                    ilAdvtIndex = gBinarySearchAdf(tgSpot.iAdfCode)
                                    If ilAdvtIndex <> -1 Then
                                        If (tgCommAdf(ilAdvtIndex).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdvtIndex).sAddrID) <> "") Then
                                            slAdvt = Trim$(tgCommAdf(ilAdvtIndex).sName) & ", " & Trim$(tgCommAdf(ilAdvtIndex).sAddrID)
                                        Else
                                            slAdvt = Trim$(tgCommAdf(ilAdvtIndex).sName)
                                        End If
                                    Else
                                        slAdvt = "Missing" & str$(tgSpot.iAdfCode)
                                    End If
                                    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName) & ", " & slAdvt
                                    ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                End If
                            End If
                            ilAvailIndex = ilAvailIndex + 1
                        End If
                    Loop
                End If
                If tmSsf.iCount + ilRoomReq <= UBound(tmSsf.tPas) Then
                    tmSsf.iCount = tmSsf.iCount + 1
                    ilSvCount = tmSsf.iCount
                    tmSsf.tPas(ADJSSFPASBZ + ilSvCount) = tmAvail
                    If ilRoomReq > 1 Then
                        'Move spots into avail (set extra to missed)
                        'Add code to miss by priority #
                        'Update Sdf record
                        ilUnitsRem = tmAvail.iAvInfo And &H1F
                        ilLenRem = tmAvail.iLen
                        gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                        ilSplitNetworkPriRemoved = False
                        For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tgAvail.iNoSpotsThis Step 1
                            LSet tgSpot = tgSsf(ilBoundIndex).tPas(ADJSSFPASBZ + ilSpotIndex)
                            If (tgSpot.iRecType And &HF) >= 10 Then
                                'Update Sdf record
                                ilError = False
                                On Error GoTo gMakeSSFErr
                                tmSdfSrchKey3.lCode = tgSpot.lSdfCode
                                ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                If (Not ilError) And (ilRet = BTRV_ERR_NONE) Then
                                    slOrigSchStatus = tmSdf.sSchStatus
                                    'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                                    If ilSplitNetworkPriRemoved Then
                                        If (tgSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                            tgSpot.iRecType = (tgSpot.iRecType And (Not SSSPLITSEC)) Or SSSPLITPRI
                                        End If
                                        ilSplitNetworkPriRemoved = False
                                    End If
                                    ilTestSplitNetworkLen = True
                                    '6/4/16: Replaced GoSub
                                    'GoSub TestSDF
                                    mTestSDF ilTestOk, ilUpdateSdf, ilRemoveSpot, ilAdfCode, llChfCode, llFsfCode, llCntrNo, ilSplitNetworkPriRemoved, ilTestSplitNetworkLen, ilBoundIndex, ilAvailIndex, ilSpotIndex, ilTestOnly, hlGsf, hlSxf, hlSdf, hlSmf, ilSdfRecLen, slDate, slTime, slMsg
                                    If Not ilTestOk Then
                                        btrDestroy hmCHF
                                        btrDestroy hmClf
                                        btrDestroy hmCff
                                        btrDestroy hmFsf
                                        btrDestroy hmAlf
                                        btrDestroy hmAtt
                                        btrDestroy hmStf
                                        btrDestroy hlGsf
                                        btrDestroy hlSxf
                                        gMakeSSF = False
                                        Exit Function
                                    End If
                                    If Not ilRemoveSpot Then
                                        'Remake tgSpot for safety
                                        'Note: Chf read in at TestSdf
                                        If (tgSpot.iAdfCode <> tmSdf.iAdfCode) Or (ilAdfCode <> tmSdf.iAdfCode) Then
                                            If ilTestOnly Then
                                                slMsg = "Advertiser Name Error: "
                                            Else
                                                slMsg = "Advertiser Name Error Fixed: "
                                            End If
                                            'tmChfSrchKey.lCode = tmSdf.lChfCode
                                            'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            'If ilRet = BTRV_ERR_NONE Then
                                                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                            'Else
                                            '    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " ChfCode=" & Str$(tmSdf.lChfCode) & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                            'End If
                                            ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                            If tmSdf.iAdfCode <> ilAdfCode Then
                                                ilUpdateSdf = True
                                                tmSdf.iAdfCode = ilAdfCode
                                                'Note: Sdf.iAdfCode is also set when updating sdf
                                            End If
                                        End If
                                        tgSpot.iAdfCode = ilAdfCode
                                        If ((tgSpot.iPosLen And &HFFF) <> (tmSdf.iLen)) Or (tmClf.iLen <> tmSdf.iLen) Then
                                            If ilTestOnly Then
                                                slMsg = "Spot Length Error: "
                                            Else
                                                slMsg = "Spot Length Error Fixed: "
                                            End If
                                            'tmChfSrchKey.lCode = tmSdf.lChfCode
                                            'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                            'If ilRet = BTRV_ERR_NONE Then
                                                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                            'Else
                                            '    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " ChfCode=" & Str$(tmSdf.lChfCode) & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                            'End If
                                            ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                            ilUpdateSdf = True
                                            tmSdf.iLen = tmClf.iLen
                                        End If
                                        slMsg = ""
                                        tgSpot.iPosLen = (tgSpot.iPosLen And &HF000) Or (tmClf.iLen)    '(tmSdf.iLen)
                                        ilMakeSpotMissed = False
                                        If (tmSdf.sSchStatus = "M") Or (tmSdf.sSchStatus = "R") Or (tmSdf.sSchStatus = "U") Or (tmSdf.sSchStatus = "C") Or (tmSdf.sSchStatus = "H") Then 'This case should not happen
                                            ilMakeSpotMissed = True
                                            If ilTestOnly Then
                                                slMsg = "Sch. Status Disparity: "
                                            Else
                                                slMsg = "Sch. Status Disparity, Removed: "
                                            End If
                                        Else
                                            'Check Date and time
                                            If (Trim$(tmSdf.sAffChg) = "") Or (tmSdf.sAffChg = "N") Then
                                                If ilType <= 0 Then
                                                    If (tmAvail.iTime(0) <> tmSdf.iTime(0)) Or (tmAvail.iTime(1) <> tmSdf.iTime(1)) Then
                                                        ilMakeSpotMissed = True
                                                        If ilTestOnly Then
                                                            slMsg = "Time Disparity: "
                                                        Else
                                                            slMsg = "Time Disparity, Removed: "
                                                        End If
                                                    Else
                                                        If (ilLogDate0 <> tmSdf.iDate(0)) Or (ilLogDate1 <> tmSdf.iDate(1)) Then
                                                            ilMakeSpotMissed = True
                                                            If ilTestOnly Then
                                                                slMsg = "Date Disparity: "
                                                            Else
                                                                slMsg = "Date Disparity, Removed: "
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                        'Test if SMF exist
                                        If (ilMakeSpotMissed = False) And ((tmSdf.sSchStatus = "G") Or (tmSdf.sSchStatus = "O")) Then
                                            If Not gFindSmf(tmSdf, hlSmf, tmSmf) Then
                                                ilMakeSpotMissed = True
                                                If ilTestOnly Then
                                                    slMsg = "MG Spec Missing: "
                                                Else
                                                    slMsg = "MG Spec Missing, Removed: "
                                                End If
                                            End If
                                        End If
                                        ilSpotLen = (tgSpot.iPosLen And &HFFF)
                                        If (tgSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
                                            If (ilVpfIndex >= 0) And (tgVpf(ilVpfIndex).sSSellOut = "T") Then
                                                ilSpotUnits = ilSpotLen \ 30
                                                If ilSpotUnits <= 0 Then
                                                    ilSpotUnits = 1
                                                End If
                                                ilSpotLen = 0
                                            Else
                                                ilSpotUnits = 1
                                                'If (ilVpfIndex >= 0) And (tgVpf(ilVpfIndex).sSSellOut = "U") Then
                                                '    ilSpotLen = 0
                                                'End If
                                            End If
                                        Else
                                            ilSpotLen = 0
                                            ilSpotUnits = 0
                                        End If
                                        'If ((ilSpotUnits <= ilUnitsRem) And (ilSpotLen <= ilLenRem) And (ilMakeSpotMissed = False) And (tgVpf(ilVpfIndex).sSSellOut <> "M")) Or ((ilSpotUnits = ilUnitsRem) And (ilSpotLen = ilLenRem) And (ilMakeSpotMissed = False) And (tgVpf(ilVpfIndex).sSSellOut = "M")) Then
                                        If (ilSpotUnits <= ilUnitsRem) And (ilSpotLen <= ilLenRem) And (ilMakeSpotMissed = False) Then
                                            If ((ilSpotUnits <> ilUnitsRem) Or (ilSpotLen <> ilLenRem)) And (tgVpf(ilVpfIndex).sSSellOut = "M") And (Not ilAffChg) And ((tmSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC) Then
                                                If Len(slMsg) = 0 Then
                                                    If ilTestOnly Then
                                                        slMsg = "Warning- Length Override: "
                                                    Else
                                                        slMsg = "Warning- Length Override: "
                                                    End If
                                                End If
                                                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                            End If
                                            ilUnitsRem = ilUnitsRem - ilSpotUnits
                                            ilLenRem = ilLenRem - ilSpotLen
                                            tmSsf.iCount = tmSsf.iCount + 1
                                            If tmSdf.lChfCode > 0 Then
                                                If tmSdf.sSpotType = "X" Then
                                                    ilPriceLevel = 0
                                                Else
                                                    ilRet = gGetSpotFlight(tmSdf, tmClf, hmCff, hlSmf, tlCff)
                                                    If tlCff.sPriceType = "T" Then
                                                        ilSaf = gBinarySearchSaf(tmClf.iVefCode)
                                                        If ilSaf = -1 Then
                                                            ilSaf = gBinarySearchSaf(0) 'Obtain from Site
                                                        Else
                                                            If (tgSaf(ilSaf).lLowPrice <= 0) And (tgSaf(ilSaf).lHighPrice <= 0) Then
                                                                ilSaf = gBinarySearchSaf(0) 'Obtain from Site
                                                            End If
                                                        End If
                                                        If ilSaf <> -1 Then
                                                            '6/20/06:  Treat zero dollars same as N/C
                                                            If (tgSaf(ilSaf).lLowPrice <= 0) And (tgSaf(ilSaf).lHighPrice <= 0) Then
                                                                ilPriceLevel = 0
                                                            ElseIf tlCff.lActPrice = 0 Then 'treat as N/C
                                                                ilPriceLevel = 1
                                                            ElseIf tlCff.lActPrice <= 100 * tgSaf(ilSaf).lLowPrice Then
                                                                ilPriceLevel = 2
                                                            Else
                                                                If tlCff.lActPrice > 100 * tgSaf(ilSaf).lHighPrice Then
                                                                    ilPriceLevel = 15
                                                                Else
                                                                    ilPriceLevel = 0
                                                                    For ilPrice = LBound(tgSaf(ilSaf).lLevelToPrice) To UBound(tgSaf(ilSaf).lLevelToPrice) Step 1
                                                                        If tlCff.lActPrice <= 100 * tgSaf(ilSaf).lLevelToPrice(ilPrice) Then
                                                                            ilPriceLevel = ilPrice - LBound(tgSaf(ilSaf).lLevelToPrice) + 3
                                                                            Exit For
                                                                        End If
                                                                    Next ilPrice
                                                                    If ilPriceLevel = 0 Then
                                                                        If tlCff.lActPrice > 100 * tgSaf(ilSaf).lLevelToPrice(UBound(tgSaf(ilSaf).lLevelToPrice)) And (tlCff.lActPrice <= 100 * tgSaf(ilSaf).lHighPrice) Then
                                                                            ilPriceLevel = 14
                                                                        End If
                                                                    End If
                                                                End If
                                                           End If
                                                        Else
                                                            ilPriceLevel = 0
                                                        End If
                                                    Else
                                                        ilPriceLevel = 1
                                                    End If
                                                End If
                                            Else
                                                ilPriceLevel = 15
                                            End If
                                            tgSpot.iRank = (ilPriceLevel * SHIFT11) + (tgSpot.iRank And RANKMASK)
                                            LSet tmSsf.tPas(ADJSSFPASBZ + tmSsf.iCount) = tgSpot
                                            tmAvail.iNoSpotsThis = tmAvail.iNoSpotsThis + 1
                                            tmSsf.tPas(ADJSSFPASBZ + ilSvCount) = tmAvail
                                            If (Not ilUpdateSdf) And (ilType > 0) Then
                                                If llGameAdjTime <> 0 Then
                                                    ilUpdateSdf = True
                                                Else
                                                    If (ilOrigSSFGameDate(0) <> ilLogDate0) Or (ilOrigSSFGameDate(1) <> ilLogDate1) Then
                                                        ilUpdateSdf = True
                                                    End If
                                                End If
                                            End If
                                            If (Not ilTestOnly) And (ilUpdateSdf) Then
                                                ilRet = gMakeTracer(hlSdf, tmSdf, 0, hmStf, -1, "S", "P", tmSdf.iRotNo, hlGsf)
                                                Do
                                                    'ilRet = btrGetDirect(hlSdf, tmSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                    tmSdfSrchKey3.lCode = tmSdf.lCode
                                                    ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                    'tmSRec = tmSdf
                                                    'ilRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                                                    'tmSdf = tmSRec
                                                    'If ilRet <> BTRV_ERR_NONE Then
                                                    '    igBtrError = gConvertErrorCode(ilRet)
                                                    '    sgErrLoc = "gMakeSSF-Get by Key Sdf(10)"
                                                    '    btrDestroy hmChf
                                                    '    btrDestroy hmClf
                                                    '    btrDestroy hmCff
                                                    '    gMakeSSF = False
                                                    '    Exit Function
                                                    'End If
                                                    tmSdf.lChfCode = llChfCode
                                                    tmSdf.iAdfCode = ilAdfCode
                                                    tmSdf.lFsfCode = llFsfCode
                                                    tmSdf.iLen = tmClf.iLen
                                                    If (ilType > 0) And (Not ilAffChg) Then
                                                        tmSdf.iDate(0) = ilLogDate0
                                                        tmSdf.iDate(1) = ilLogDate1
                                                        tmSdf.iTime(0) = tmAvail.iTime(0)
                                                        tmSdf.iTime(1) = tmAvail.iTime(1)
                                                    End If
                                                    tmSdf.sXCrossMidnight = "N"
                                                    If (tmAvail.iAvInfo And SSXMID) = SSXMID Then
                                                        tmSdf.sXCrossMidnight = "Y"
                                                    End If
                                                    tmSdf.sWasMG = "N"
                                                    tmSdf.sFromWorkArea = "N"
                                                    tmSdf.iUrfCode = tgUrf(0).iCode
                                                    ilRet = btrUpdate(hlSdf, tmSdf, ilSdfRecLen)
                                                Loop While ilRet = BTRV_ERR_CONFLICT
                                            End If
                                        Else
                                            If Not ilTestOnly Then
                                                If (tgSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
                                                    ilSplitNetworkPriRemoved = True
                                                Else
                                                    ilSplitNetworkPriRemoved = False
                                                End If
                                                ilRet = gMakeTracer(hlSdf, tmSdf, 0, hmStf, -1, "M", "P", tmSdf.iRotNo, hlGsf)
                                                ilRet = gRemoveSmf(hlSmf, tmSmf, tmSdf, hlSxf)  'resets missed date
                                                If Not ilRet Then
                                                    btrDestroy hmCHF
                                                    btrDestroy hmClf
                                                    btrDestroy hmCff
                                                    btrDestroy hmFsf
                                                    btrDestroy hmAlf
                                                    btrDestroy hmAtt
                                                    btrDestroy hmStf
                                                    btrDestroy hlGsf
                                                    btrDestroy hlSxf
                                                    gMakeSSF = False
                                                    Exit Function
                                                End If
                                            Else
                                                ilRet = gFindSmf(tmSdf, hlSmf, tmSmf)
                                            End If
                                            'If (tmSdf.sSpotType = "T") Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "S") Or (tmSdf.sSpotType = "M") Or (tmSdf.sSpotType = "X") Then
                                            If tmSdf.sSpotType = "X" Then
                                                slXSpotType = "X"
                                                If ((tmSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                                                    slXSpotType = ""
                                                End If
                                            Else
                                                slXSpotType = ""
                                            End If
                                            'If (tmSdf.sSpotType = "T") Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "S") Or (tmSdf.sSpotType = "M") Or (slXSpotType = "X") Then
                                            If ((tmSdf.sSpotType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmSdf.sSpotType = "Q") Or ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (slXSpotType = "X") Then
                                                'tmChfSrchKey.lCode = tmSdf.lChfCode
                                                'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If Len(slMsg) = 0 Then
                                                    If ilTestOnly Then
                                                        slMsg = "Overbooked Avail: "
                                                    Else
                                                        slMsg = "Overbooked Avail, Removed: "
                                                    End If
                                                End If
                                                'If ilRet = BTRV_ERR_NONE Then
                                                    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                                'Else
                                                '    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " ChfCode=" & Str$(tmSdf.lChfCode) & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                                'End If
                                                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                                If Not ilTestOnly Then
                                                    'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                                                    Do
                                                        tmSdfSrchKey3.lCode = tmSdf.lCode
                                                        ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                        'tmSRec = tmSdf
                                                        'ilCRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                                                        'tmSdf = tmSRec
                                                        'If ilCRet <> BTRV_ERR_NONE Then
                                                        '    igBtrError = ilCRet
                                                        '    sgErrLoc = "gMakeSSF-Get by Key Sdf(11)"
                                                        '    btrDestroy hmChf
                                                        '    btrDestroy hmClf
                                                        '    btrDestroy hmCff
                                                        '    gMakeSSF = False
                                                        '    Exit Function
                                                        'End If
                                                        ilRet = btrDelete(hlSdf)
                                                        'If ilRet = BTRV_ERR_CONFLICT Then
                                                        '    ilCRet = btrGetDirect(hlSdf, tmSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                        'End If
                                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                                End If
                                            Else
                                                ilDate0 = tmSdf.iDate(0)
                                                ilDate1 = tmSdf.iDate(1)
                                                ilTime0 = tmSdf.iTime(0)
                                                ilTime1 = tmSdf.iTime(1)
                                                ilOrigSchVef = tmSdf.iVefCode
                                                ilOrigGameNo = tmSdf.iGameNo
                                                'tmChfSrchKey.lCode = tmSdf.lChfCode
                                                'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                If Len(slMsg) = 0 Then
                                                    If ilTestOnly Then
                                                        slMsg = "Overbooked Avail: "
                                                    Else
                                                        slMsg = "Overbooked Avail, Removed: "
                                                    End If
                                                End If
                                                'If ilRet = BTRV_ERR_NONE Then
                                                    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                                'Else
                                                '    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " ChfCode=" & Str$(tmSdf.lChfCode) & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                                'End If
                                                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                                If Not ilTestOnly Then
                                                    Do
                                                        'ilRet = btrGetDirect(hlSdf, tmSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                        tmSdfSrchKey3.lCode = tmSdf.lCode
                                                        ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                        'tmSRec = tmSdf
                                                        'ilRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                                                        'tmSdf = tmSRec
                                                        'If ilRet <> BTRV_ERR_NONE Then
                                                        '    igBtrError = gConvertErrorCode(ilRet)
                                                        '    sgErrLoc = "gMakeSSF-Get by Key Sdf(12)"
                                                        '    btrDestroy hmChf
                                                        '    btrDestroy hmClf
                                                        '    btrDestroy hmCff
                                                        '    gMakeSSF = False
                                                        '    Exit Function
                                                        'End If
                                                        tmSdf.lChfCode = llChfCode
                                                        tmSdf.iAdfCode = ilAdfCode
                                                        tmSdf.lFsfCode = llFsfCode
                                                        tmSdf.iLen = tmClf.iLen
                                                        tmSdf.sSchStatus = "M"
                                                        tmSdf.iDate(0) = ilDate0
                                                        tmSdf.iDate(1) = ilDate1
                                                        tmSdf.iTime(0) = ilTime0
                                                        tmSdf.iTime(1) = ilTime1
                                                        tmSdf.iVefCode = ilOrigSchVef
                                                        tmSdf.iGameNo = ilOrigGameNo
                                                        tmSdf.lSmfCode = 0
                                                        If ((tmSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                                                            tmSdf.sTracer = "*"
                                                            tmSdf.lSmfCode = tmSmf.lMtfCode
                                                        End If
                                                        If (ilType > 0) And (slOrigSchStatus <> "G") And (slOrigSchStatus <> "O") Then
                                                            tmSdf.iDate(0) = ilLogDate0
                                                            tmSdf.iDate(1) = ilLogDate1
                                                        End If
                                                        tmSdf.sXCrossMidnight = "N"
                                                        tmSdf.sWasMG = "N"
                                                        tmSdf.sFromWorkArea = "N"
                                                        tmSdf.iUrfCode = tgUrf(0).iCode
                                                        ilRet = btrUpdate(hlSdf, tmSdf, ilSdfRecLen)
                                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                                    lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                                                    ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    If ilTestOnly Then
                                        slMsg = "Can't Find Sdf: "
                                    Else
                                        slMsg = "Can't Find Sdf, Removed: "
                                    End If
                                    If Not ilTestOnly Then
                                        If (tgSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
                                            ilSplitNetworkPriRemoved = True
                                        Else
                                            ilSplitNetworkPriRemoved = False
                                        End If
                                    End If
                                    ilAdvtIndex = gBinarySearchAdf(tgSpot.iAdfCode)
                                    If ilAdvtIndex <> -1 Then
                                        If (tgCommAdf(ilAdvtIndex).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdvtIndex).sAddrID) <> "") Then
                                            slAdvt = Trim$(tgCommAdf(ilAdvtIndex).sName) & ", " & Trim$(tgCommAdf(ilAdvtIndex).sAddrID)
                                        Else
                                            slAdvt = Trim$(tgCommAdf(ilAdvtIndex).sName)
                                        End If
                                    Else
                                        slAdvt = "Missing" & str$(tgSpot.iAdfCode)
                                    End If
                                    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName) & ", " & slAdvt
                                    ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                End If
                            Else
                                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Avail Count not matching # of Spots: " & slDate & " " & slTime & " " & Trim$(tmVef.sName)
                                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                            End If
                        Next ilSpotIndex
                    End If
                    If ilMatchTime Then
                        ilAvailIndex = ilAvailIndex + ilRoomReq
                    End If
                Else
                    ilBuildHdSSF = True
                End If
            End If
        Loop While ilBuildHdSSF
    Next ilIndex
    If Not ilTestOnly Then
        If tmSsf.iCount > 0 Then
            ilSsfRecLen = igSSFBaseLen + tmSsf.iCount * Len(tmAvail)
            tmSsf.lCode = 0
            ilRet = gSSFInsert(hlSsf, tmSsf, ilSsfRecLen, INDEXKEY3)
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = gConvertErrorCode(ilRet)
                sgErrLoc = "gMakeSSF-Insert Ssf(13)"
                btrDestroy hmCHF
                btrDestroy hmClf
                btrDestroy hmCff
                btrDestroy hmFsf
                btrDestroy hmAlf
                btrDestroy hmAtt
                btrDestroy hmStf
                btrDestroy hlGsf
                btrDestroy hlSxf
                gMakeSSF = False
                Exit Function
            End If
        End If
    End If
    'Set any remaining spots as Missed
    If (ilUpperBound >= LBound(tgSsf)) And (ilBoundIndex <= ilUpperBound) Then
        Do
            If ilAvailIndex > tgSsf(ilBoundIndex).iCount Then
                ilBoundIndex = ilBoundIndex + 1
                If (ilBoundIndex > ilUpperBound) Then
                    Exit Do
                End If
                ilAvailIndex = 1
            End If
            'Check each record incase spot outside an avail instead of looping
            'on spots within an avail
            'tgAvail = tgSsf(ilBoundIndex).tPas(ADJSSFPASBZ + ilAvailIndex)
            LSet tgSpot = tgSsf(ilBoundIndex).tPas(ADJSSFPASBZ + ilAvailIndex)
            'If (tgAvail.iRecType >= 2) And (tgAvail.iRecType <= 9) Then
            If ((tgSpot.iRecType And &HF) >= 10) And ((tgSpot.iRecType And &HF) <= 11) Then
                'set spots to missed
                'For ilSpotIndex = ilAvailIndex + 1 To ilAvailIndex + tgAvail.iNoSpotsThis Step 1
                    'LSet tgSpot = tgSsf(ilBoundIndex).tPas(ADJSSFPASBZ + ilSpotIndex)
                    'Update Sdf record
                    tmSdfSrchKey3.lCode = tgSpot.lSdfCode
                    ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                    If (ilRet = BTRV_ERR_NONE) Then
                        slOrigSchStatus = tmSdf.sSchStatus
                        'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                        ilTestSplitNetworkLen = False
                        '6/4/16: Replaced GoSub
                        'GoSub TestSDF
                        mTestSDF ilTestOk, ilUpdateSdf, ilRemoveSpot, ilAdfCode, llChfCode, llFsfCode, llCntrNo, ilSplitNetworkPriRemoved, ilTestSplitNetworkLen, ilBoundIndex, ilAvailIndex, ilSpotIndex, ilTestOnly, hlGsf, hlSxf, hlSdf, hlSmf, ilSdfRecLen, slDate, slTime, slMsg
                        ilSplitNetworkPriRemoved = False
                        If Not ilTestOk Then
                            btrDestroy hmCHF
                            btrDestroy hmClf
                            btrDestroy hmCff
                            btrDestroy hmFsf
                            btrDestroy hmAlf
                            btrDestroy hmAtt
                            btrDestroy hmStf
                            btrDestroy hlGsf
                            btrDestroy hlSxf
                            gMakeSSF = False
                            Exit Function
                        End If
                        If Not ilRemoveSpot Then
                            If Not ilTestOnly Then
                                ilRet = gMakeTracer(hlSdf, tmSdf, 0, hmStf, -1, "M", "P", tmSdf.iRotNo, hlGsf)
                                ilRet = gRemoveSmf(hlSmf, tmSmf, tmSdf, hlSxf)  'resets missed date
                                If Not ilRet Then
                                    btrDestroy hmCHF
                                    btrDestroy hmClf
                                    btrDestroy hmCff
                                    btrDestroy hmFsf
                                    btrDestroy hmAlf
                                    btrDestroy hmAtt
                                    btrDestroy hmStf
                                    btrDestroy hlGsf
                                    btrDestroy hlSxf
                                    gMakeSSF = False
                                    Exit Function
                                End If
                            Else
                                ilRet = gFindSmf(tmSdf, hlSmf, tmSmf)
                            End If
                            gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                            ''tmChfSrchKey.lCode = tmSdf.lChfCode
                            ''ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                            ''If (tmChf.sType = "T") Or (tmChf.sType = "Q") Or (tmChf.sType = "S") Or (tmChf.sType = "M") Then
                            'If (tmSdf.sSpotType = "T") Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "S") Or (tmSdf.sSpotType = "M") Or (tmSdf.sSpotType = "X") Then
                            If tmSdf.sSpotType = "X" Then
                                slXSpotType = "X"
                                If ((tmSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                                    slXSpotType = ""
                                End If
                            Else
                                slXSpotType = ""
                            End If
                            'If (tmSdf.sSpotType = "T") Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "S") Or (tmSdf.sSpotType = "M") Or (slXSpotType = "X") Then
                            If ((tmSdf.sSpotType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmSdf.sSpotType = "Q") Or ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (slXSpotType = "X") Then
                                'tmChfSrchKey.lCode = tmSdf.lChfCode
                                'ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                If ilTestOnly Then
                                    slMsg = "No Avail: "
                                Else
                                    slMsg = "No Avail, Removed: "
                                End If
                                'If ilRet = BTRV_ERR_NONE Then
                                    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                'Else
                                '    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " ChfCode=" & Str$(tmSdf.lChfCode) & " Line=" & Str$(tmSdf.iLineNo) & " Sdf ID=" & Str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                'End If
                                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                If Not ilTestOnly Then
                                    'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                                    Do
                                        tmSdfSrchKey3.lCode = tmSdf.lCode
                                        ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                        'tmSRec = tmSdf
                                        'ilCRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                                        'tmSdf = tmSRec
                                        'If ilCRet <> BTRV_ERR_NONE Then
                                        '    igBtrError = gConvertErrorCode(ilRet)
                                        '    sgErrLoc = "gMakeSSF-Get by Key Sdf(14)"
                                        '    btrDestroy hmChf
                                        '    btrDestroy hmClf
                                        '    btrDestroy hmCff
                                        '    gMakeSSF = False
                                        '    Exit Function
                                        'End If
                                        ilRet = btrDelete(hlSdf)
                                        'If ilRet = BTRV_ERR_CONFLICT Then
                                        '    ilCRet = btrGetDirect(hlSdf, tmSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                        'End If
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                End If
                            Else
                                ilDate0 = tmSdf.iDate(0)
                                ilDate1 = tmSdf.iDate(1)
                                ilTime0 = tmSdf.iTime(0)
                                ilTime1 = tmSdf.iTime(1)
                                ilOrigSchVef = tmSdf.iVefCode
                                ilOrigGameNo = tmSdf.iGameNo
                                If tmSdf.lChfCode > 0 Then
                                    tmChfSrchKey.lCode = tmSdf.lChfCode
                                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                Else
                                    tmFsfSrchKey0.lCode = tmSdf.lFsfCode
                                    ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                End If
                                If ilTestOnly Then
                                    slMsg = "No Avail: "
                                Else
                                    slMsg = "No Avail, Missed: "
                                End If
                                If tmSdf.lChfCode > 0 Then
                                    If ilRet = BTRV_ERR_NONE Then
                                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                    Else
                                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " ChfCode=" & str$(tmSdf.lChfCode) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                    End If
                                Else
                                    If ilRet = BTRV_ERR_NONE Then
                                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Feed Ref #=" & Trim$(tmFsf.sRefId) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                    Else
                                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " FsfCode=" & str$(tmSdf.lFsfCode) & " Sdf ID=" & str$(tmSdf.lCode) & " " & Trim$(tmVef.sName)
                                    End If
                                End If
                                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                If Not ilTestOnly Then
                                    Do
                                        'ilRet = btrGetDirect(hlSdf, tmSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                        tmSdfSrchKey3.lCode = tmSdf.lCode
                                        ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                        'tmSRec = tmSdf
                                        'ilRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                                        'tmSdf = tmSRec
                                        'If ilRet <> BTRV_ERR_NONE Then
                                        '    igBtrError = gConvertErrorCode(ilRet)
                                        '    sgErrLoc = "gMakeSSF-Get by Key Sdf(15)"
                                        '    btrDestroy hmChf
                                        '    btrDestroy hmClf
                                        '    btrDestroy hmCff
                                        '    gMakeSSF = False
                                        '    Exit Function
                                        'End If
                                        tmSdf.lChfCode = llChfCode
                                        tmSdf.iAdfCode = ilAdfCode
                                        tmSdf.lFsfCode = llFsfCode
                                        tmSdf.iLen = tmClf.iLen
                                        tmSdf.sSchStatus = "M"
                                        tmSdf.iMnfMissed = igMnfMissed
                                        tmSdf.iDate(0) = ilDate0
                                        tmSdf.iDate(1) = ilDate1
                                        tmSdf.iTime(0) = ilTime0
                                        tmSdf.iTime(1) = ilTime1
                                        tmSdf.iVefCode = ilOrigSchVef
                                        tmSdf.iGameNo = ilOrigGameNo
                                        tmSdf.lSmfCode = 0
                                        If ((tmSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                                            tmSdf.sTracer = "*"
                                            tmSdf.lSmfCode = tmSmf.lMtfCode
                                        End If
                                        If (ilType > 0) And (slOrigSchStatus <> "G") And (slOrigSchStatus <> "O") Then
                                            tmSdf.iDate(0) = ilLogDate0
                                            tmSdf.iDate(1) = ilLogDate1
                                        End If
                                        tmSdf.sXCrossMidnight = "N"
                                        tmSdf.sWasMG = "N"
                                        tmSdf.sFromWorkArea = "N"
                                        tmSdf.iUrfCode = tgUrf(0).iCode
                                        ilRet = btrUpdate(hlSdf, tmSdf, ilSdfRecLen)
                                    Loop While ilRet = BTRV_ERR_CONFLICT
                                    lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                                    ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                                End If
                            End If
                        End If
                    Else
                        If ilTestOnly Then
                            slMsg = "Can't Find Sdf: "
                        Else
                            slMsg = "Can't Find Sdf, Removed: "
                        End If
                        ilAdvtIndex = gBinarySearchAdf(tgSpot.iAdfCode)
                        If ilAdvtIndex <> -1 Then
                            If (tgCommAdf(ilAdvtIndex).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdvtIndex).sAddrID) <> "") Then
                                slAdvt = Trim$(tgCommAdf(ilAdvtIndex).sName) & ", " & Trim$(tgCommAdf(ilAdvtIndex).sAddrID)
                            Else
                                slAdvt = Trim$(tgCommAdf(ilAdvtIndex).sName)
                            End If
                        Else
                            slAdvt = "Missing" & str$(tgSpot.iAdfCode)
                        End If
                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName) & ", " & slAdvt
                        ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                    End If
                'Next ilSpotIndex
                'ilAvailIndex = ilAvailIndex + tgAvail.iNoSpotsThis + 1
                ilAvailIndex = ilAvailIndex + 1
            Else
                ilAvailIndex = ilAvailIndex + 1
            End If
        Loop
    End If
    ilRet = btrClose(hmCHF)
    btrDestroy hmCHF
    ilRet = btrClose(hmClf)
    btrDestroy hmClf
    ilRet = btrClose(hmCff)
    btrDestroy hmCff
    ilRet = btrClose(hmFsf)
    btrDestroy hmFsf
    ilRet = btrClose(hmAlf)
    btrDestroy hmAlf
    ilRet = btrClose(hmAtt)
    btrDestroy hmAtt
    ilRet = btrClose(hmStf)
    btrDestroy hmStf
    ilRet = btrClose(hlGsf)
    btrDestroy hlGsf
    btrDestroy hlSxf
    gMakeSSF = True
    Exit Function
gMakeSSFErr:
    On Error GoTo 0
    ilError = True
    Resume Next
'TestSDF:
'    ilTestOk = True
'    ilUpdateSdf = False
'    ilRemoveSpot = False
'    ilAdfCode = tmSdf.iAdfCode
'    llChfCode = tmSdf.lChfCode
'    llFsfCode = tmSdf.lFsfCode
'    llCntrNo = 0
'    If (tmSdf.lChfCode = 0) And (tmSdf.lFsfCode = 0) Then
'        ilRemoveSpot = True
'    Else
'        If tmSdf.lChfCode > 0 Then
'        'Test if line exist
'            tmChfSrchKey.lCode = tmSdf.lChfCode
'            ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'            If ilRet = BTRV_ERR_NONE Then
'                If (tmSdf.sSpotType = "X") And ((tgSpot.iRank And RANKMASK) < 1000) Then
'                    If tmChf.sType = "S" Then
'                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + 1060
'                    ElseIf tmChf.sType = "M" Then
'                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + 1050
'                    ElseIf tmChf.sType = "Q" Then
'                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + 1030
'                    Else
'                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + 1045
'                    End If
'                ElseIf (tmSdf.sSpotType <> "X") Then
'                    If tmChf.sType = "S" Then
'                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + 1060
'                    ElseIf tmChf.sType = "M" Then
'                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + 1050
'                    ElseIf tmChf.sType = "Q" Then
'                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + 1030
'                    ElseIf tmChf.sType = "T" Then
'                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + 1020
'                    ElseIf tmChf.sType = "R" Then
'                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + 1010
'                    End If
'                End If
'                ilAdfCode = tmChf.iAdfCode
'                llCntrNo = tmChf.lCntrNo
'                If tmChf.sDelete = "Y" Then
'                    tmChfSrchKey1.lCntrNo = tmChf.lCntrNo
'                    tmChfSrchKey1.iCntRevNo = 32000
'                    tmChfSrchKey1.iPropVer = 32000
'                    ilRet = btrGetGreaterOrEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llCntrNo)
'                        If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.sDelete <> "Y") Then
'                            Exit Do
'                        End If
'                        ilRet = btrGetNext(hmChf, tmChf, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                    Loop
'                    If (tmChf.lCntrNo = llCntrNo) And (tmChf.sDelete <> "Y") Then
'                        llChfCode = tmChf.lCode
'                        ilAdfCode = tmChf.iAdfCode
'                        llCntrNo = tmChf.lCntrNo
'                        ilUpdateSdf = True
'                        sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Contract Code Error:" & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
'                        ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
'                    Else
'                        tmChfSrchKey.lCode = tmSdf.lChfCode
'                        ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'                        If ilRet = BTRV_ERR_NONE Then
'                            'ilRemoveSpot = True
'                            tmChfSrchKey1.lCntrNo = tmChf.lCntrNo
'                            tmChfSrchKey1.iCntRevNo = 32000
'                            tmChfSrchKey1.iPropVer = 32000
'                            ilRet = btrGetGreaterOrEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'                            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llCntrNo)
'                                If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) Then
'                                    Exit Do
'                                End If
'                                ilRet = btrGetNext(hmChf, tmChf, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                            Loop
'                            If (ilRet <> BTRV_ERR_NONE) Or (tmChf.lCntrNo <> llCntrNo) Then
'                                tmChfSrchKey.lCode = tmSdf.lChfCode
'                                ilRet = btrGetEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'                            End If
'                            If llChfCode <> tmChf.lCode Then
'                                ilUpdateSdf = True
'                            End If
'                            llChfCode = tmChf.lCode
'                            ilAdfCode = tmChf.iAdfCode
'                            llCntrNo = tmChf.lCntrNo
'                            sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Contract Delete Flag Error:" & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
'                            ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
'                            If Not ilTestOnly Then
'                                ilRet = btrGetPosition(hmChf, llChfRecPos)
'                                Do
'                                    ilRet = btrGetDirect(hmChf, tmChf, imChfRecLen, llChfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'                                    'tmSRec = tmChf
'                                    'ilRet = gGetByKeyForUpdate("CHF", hmChf, tmSRec)
'                                    'tmChf = tmSRec
'                                    tmChf.sDelete = "N"
'                                    ilRet = btrUpdate(hmChf, tmChf, imChfRecLen)
'                                Loop While ilRet = BTRV_ERR_CONFLICT
'                            End If
'                        Else
'                            ilRemoveSpot = True
'                        End If
'                    End If
'                End If
'                tmClfSrchKey.lChfCode = llChfCode
'                tmClfSrchKey.iLine = tmSdf.iLineNo
'                tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
'                tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
'                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
'                If (ilRet <> BTRV_ERR_NONE) Or (tmClf.lChfCode <> llChfCode) Or (tmClf.iLine <> tmSdf.iLineNo) Then
'                    ilRemoveSpot = True
'                    'Determine if Line Not moved to correct header
'                    tmChfSrchKey1.lCntrNo = llCntrNo
'                    tmChfSrchKey1.iCntRevNo = 32000
'                    tmChfSrchKey1.iPropVer = 32000
'                    ilRet = btrGetGreaterOrEqual(hmChf, tmTChf, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'                    Do While (ilRet = BTRV_ERR_NONE) And (tmTChf.lCntrNo = llCntrNo)
'                        tmClfSrchKey.lChfCode = tmTChf.lCode
'                        tmClfSrchKey.iLine = tmSdf.iLineNo
'                        tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
'                        tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
'                        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
'                        If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmTChf.lCode) And (tmClf.iLine = tmSdf.iLineNo) And (tmClf.sDelete <> "Y") Then
'                            'Set Line and Flight to Current header
'                            If tmTChf.lCode <> llChfCode Then
'                                ilRemoveSpot = False
'                                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Line Contract Code Error:" & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
'                                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
'                                If ilTestOnly Then
'                                    Exit Do
'                                End If
'                                ilRet = btrGetPosition(hmClf, llClfRecPos)
'                                Do
'                                    ilRet = btrGetDirect(hmClf, tmClf, imClfRecLen, llClfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'                                    'tmSRec = tmClf
'                                    'ilRet = gGetByKeyForUpdate("CLF", hmClf, tmSRec)
'                                    'tmClf = tmSRec
'                                    'If ilCRet <> BTRV_ERR_NONE Then
'                                    '    Exit Do
'                                    'End If
'                                    ilRet = btrDelete(hmClf)
'                                Loop While ilRet = BTRV_ERR_CONFLICT
'                                tmClf.lChfCode = llChfCode
'                                ilRet = btrInsert(hmClf, tmClf, imClfRecLen, INDEXKEY0)
'                                tmCffSrchKey.lChfCode = tmTChf.lCode
'                                tmCffSrchKey.iClfLine = tmSdf.iLineNo
'                                tmCffSrchKey.iCntRevNo = tmClf.iCntRevNo
'                                tmCffSrchKey.iPropVer = tmClf.iPropVer
'                                tmCffSrchKey.iStartDate(0) = 0
'                                tmCffSrchKey.iStartDate(1) = 0
'                                ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
'                                Do While (ilRet = BTRV_ERR_NONE) And (tmCff.lChfCode = tmTChf.lCode) And (tmCff.iClfLine = tmSdf.iLineNo)
'                                    ilRet = btrGetPosition(hmCff, llCffRecPos)
'                                    Do
'                                        ilRet = btrGetDirect(hmCff, tmCff, imCffRecLen, llCffRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'                                        'tmSRec = tmCff
'                                        'ilRet = gGetByKeyForUpdate("CFF", hmCff, tmSRec)
'                                        'tmCff = tmSRec
'                                        'If ilCRet <> BTRV_ERR_NONE Then
'                                        '    Exit Do
'                                        'End If
'                                        ilRet = btrDelete(hmCff)
'                                    Loop While ilRet = BTRV_ERR_CONFLICT
'                                    tmCff.lChfCode = llChfCode
'                                    tmCff.lCode = 0
'                                    ilRet = btrInsert(hmCff, tmCff, imCffRecLen, INDEXKEY1)
'                                    tmCffSrchKey.lChfCode = tmTChf.lCode
'                                    tmCffSrchKey.iClfLine = tmSdf.iLineNo
'                                    tmCffSrchKey.iCntRevNo = tmClf.iCntRevNo
'                                    tmCffSrchKey.iPropVer = tmClf.iPropVer
'                                    tmCffSrchKey.iStartDate(0) = 0
'                                    tmCffSrchKey.iStartDate(1) = 0
'                                    ilRet = btrGetGreaterOrEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
'                                Loop
'                            End If
'                            Exit Do
'                        End If
'                        ilRet = btrGetNext(hmChf, tmTChf, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                    Loop
'                Else
'                    If (tmSdf.sSchStatus = "S") And (tmSdf.iVefCode <> tmClf.iVefCode) Then
'                        ilRemoveSpot = True
'                    End If
'                End If
'            Else
'                ilRemoveSpot = True
'            End If
'        Else
'            tmFsfSrchKey0.lCode = tmSdf.lFsfCode
'            ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'            If ilRet = BTRV_ERR_NONE Then
'                ilAdfCode = tmFsf.iAdfCode
'                llCntrNo = 0
'                'Determine if this is the latest fsf
'                tmFsfSrchKey4.lPrevFsfCode = tmFsf.lCode
'                ilRet = btrGetEqual(hmFsf, tmTFsf, imFsfRecLen, tmFsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'                Do While (ilRet = BTRV_ERR_NONE)
'                    If (tmTFsf.sSchStatus <> "F") Then
'                        Exit Do
'                    End If
'                    tmFsf = tmTFsf
'                    tmFsfSrchKey4.lPrevFsfCode = tmTFsf.lCode
'                    ilRet = btrGetEqual(hmFsf, tmTFsf, imFsfRecLen, tmFsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'                Loop
'                If tmSdf.lFsfCode <> tmFsf.lCode Then
'                    ilUpdateSdf = True
'                    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Feed Code Error:" & slDate & " " & slTime & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
'                    ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
'                End If
'                ilAdfCode = tmFsf.iAdfCode
'                llFsfCode = tmFsf.lCode
'                'Check vehicle
'                If (tmSdf.sSchStatus = "S") And (tmSdf.iVefCode <> tmFsf.iVefCode) Then
'                    ilRemoveSpot = True
'                End If
'            Else
'                ilRemoveSpot = True
'            End If
'        End If
'    End If
'    If (Not ilRemoveSpot) And (tmVef.sType = "A") Then
'        ilRemoveSpot = True
'    End If
'    If (Not ilRemoveSpot) And (ilTestSplitNetworkLen) Then
'        If (tgSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
'            For ilTest = ilSpotIndex - 1 To ilAvailIndex + 1 Step -1
'               LSet tlSpot = tgSsf(ilBoundIndex).tPas(ADJSSFPASBZ + ilTest)
'                If (tlSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
'                    If (tgSpot.iPosLen And &HFFF) <> (tlSpot.iPosLen And &HFFF) Then
'                        ilRemoveSpot = True
'                    End If
'                    Exit For
'                End If
'            Next ilTest
'        End If
'    End If
'    If ilRemoveSpot Then
'        If ilTestOnly Then
'            If tmVef.sType = "A" Then
'                slMsg = "Sdf in Air Veh: "
'            Else
'                slMsg = "Sdf Data Missing: "
'            End If
'        Else
'            If (tgSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
'                ilSplitNetworkPriRemoved = True
'            Else
'                ilSplitNetworkPriRemoved = False
'            End If
'            ilRet = gMakeTracer(hlSdf, tmSdf, 0, hmStf, -1, "M", "P", tmSdf.iRotNo, hlGsf)
'            ilRet = gRemoveSmf(hlSmf, tmSmf, tmSdf, hlSxf)  'resets missed date
'            Do
'                tmSdfSrchKey3.lCode = tmSdf.lCode
'                ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
'                'tmSRec = tmSdf
'                'ilCRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
'                'tmSdf = tmSRec
'                'If ilCRet <> BTRV_ERR_NONE Then
'                '    igBtrError = ilCRet
'                '    sgErrLoc = "gMakeSSF-Get by Key Sdf(16)"
'                '    ilTestOk = False
'                'End If
'                ilRet = btrDelete(hlSdf)
'                'If ilRet = BTRV_ERR_CONFLICT Then
'                '    ilCRet = btrGetDirect(hlSdf, tmSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
'                'End If
'            Loop While ilRet = BTRV_ERR_CONFLICT
'            If ilRet <> BTRV_ERR_NONE Then
'                igBtrError = gConvertErrorCode(ilRet)
'                sgErrLoc = "gMakeSSF-Delete Sdf(17)"
'                ilTestOk = False
'            End If
'            If tmVef.sType = "A" Then
'                slMsg = "Sdf in Air Veh, Removed: "
'            Else
'                slMsg = "Sdf Data Missing, Removed: "
'            End If
'        End If
'        If llCntrNo = 0 Then
'            sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Line #=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
'        Else
'            sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(llCntrNo) & " Line #=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
'        End If
'        ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
'    End If
'    Return
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gMakeTracer                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Create a STF if scheduled or   *
'*                      missed.  When set to missed    *
'*                      this routine should only be    *
'*                      called for scheduled spots     *
'*                      being removed                  *
'*                                                     *
'*******************************************************
Function gMakeTracer(hlSdf As Integer, tlSdf As SDF, llSdfRecPos As Long, hlStf As Integer, llInLastLogDate As Long, slStatus As String, slReason As String, ilRotNo As Integer, Optional hlGsf As Integer = 0) As Integer
'
'   ilRet = gMakeTracer(hlSdf, tlSdf, llSdfRecPos, hlStf, llLastLogDate, slStatus, slReason, ilRotNo)
'   Where:
'       hlSdf(I)- Handle for Sdf
'       tlSdf(O)- Stf record work image area
'       llSdfRecPos(I)- Record position used to read Sdf
'       hlStf(I)- Stf handle
'       llLastLogDate(I)- Last Log Date
'       slStatus(I)- Try of operation Add or Remove) "S"=Add, "M" or "C" or "H"=Remove
'       slReason(I)- Default tracer reason (M=Mouse Move, C=Contract Line Change, P=Post Log, U=Unscheduled)
'       ilRotNo(I)- Previous Rotation # (copy changed- sdf will have the new value)
'
'   Note: tlSdf contains spot
'
    Dim slStr As String
    Dim slTracer As String
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim slLogDate As String
    Dim ilSdfRecLen As Integer
    Dim ilStfRecLen As Integer
    Dim tlStf As STF
    Dim llNowDate As Long
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim llLastLogDate As Long

    ilSdfRecLen = Len(tlSdf)
    ilStfRecLen = Len(tlStf)
'    If (slStatus = "M") Or (slStatus = "C") Or (slStatus = "S") Then
        If llSdfRecPos > 0 Then
            ilRet = btrGetDirect(hlSdf, tlSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
        Else
            tmSdfSrchKey3.lCode = tlSdf.lCode
            ilRet = btrGetEqual(hlSdf, tlSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        End If
        'tmSRec = tlSdf
        'ilRet = gGetByKeyForUpdate("Sdf", hlSdf, tmSRec)
        'tlSdf = tmSRec
        'If ilRet <> BTRV_ERR_NONE Then
        '    igBtrError = gConvertErrorCode(ilRet)
        '    sgErrLoc = "gMakeTracer-GetByKey Sdf(1)"
        '    gMakeTracer = False
        '    Exit Function
        'End If
        ilFound = -1
        'For ilLoop = 0 To UBound(tgVpf) Step 1
        '    If tlSdf.iVefCode = tgVpf(ilLoop).iVefKCode Then
            ilLoop = gBinarySearchVpf(tlSdf.iVefCode)
            If ilLoop <> -1 Then
                ilFound = ilLoop
        '        Exit For
            End If
        'Next ilLoop
        If ilFound >= 0 Then
            '6/30/06:  Retain tracer if not allowed to move spots between todays date and the last log date
            '          Act as if the last log date was 10 greater then todays date
            '          This is being done so that we have a record of those adds and moves
            'If tgVpf(ilFound).sMoveLLD = "Y" Then
                If tgVpf(ilFound).sMoveLLD = "Y" Then
                    If llInLastLogDate <= 0 Then
                        gUnpackDateLong tgVpf(ilFound).iLLD(0), tgVpf(ilFound).iLLD(1), llLastLogDate
                    Else
                        llLastLogDate = llInLastLogDate
                    End If
                Else
                    llLastLogDate = gDateValue(Format$(gNow(), "m/d/yy")) + 10
                End If
            '6/30/06: End of change
                llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
                gUnpackDate tlSdf.iDate(0), tlSdf.iDate(1), slLogDate
                '7/14/12: Added routine to handle Log Alerts since different date test required then Tracer date test
                gMakeLogAlert tlSdf, "S", hlGsf
                If (gDateValue(slLogDate) > llNowDate) And (gDateValue(slLogDate) <= llLastLogDate) Then
                    '7/14/12: Created gMakeLogAlert which is added above
                    'ilRet = gAlertAdd("L", "S", 0, tlSdf.iVefCode, slLogDate)
                    'Make track record
                    tlStf.lCode = 0
                    tlStf.iVefCode = tlSdf.iVefCode
                    tlStf.lChfCode = tlSdf.lChfCode
                    tlStf.iLineNo = tlSdf.iLineNo
                    tlStf.lFsfCode = tlSdf.lFsfCode
                    slStr = Format$(gNow(), "m/d/yy")
                    gPackDate slStr, tlStf.iCreateDate(0), tlStf.iCreateDate(1)
                    slStr = Format$(gNow(), "h:m:s AM/PM")
                    gPackTime slStr, tlStf.iCreateTime(0), tlStf.iCreateTime(1)
                    If (slStatus <> "S") Then
                        tlStf.sAction = "R"
                    Else
                        tlStf.sAction = "A"
                    End If
                    tlStf.iLogDate(0) = tlSdf.iDate(0)
                    tlStf.iLogDate(1) = tlSdf.iDate(1)
                    tlStf.iLogTime(0) = tlSdf.iTime(0)
                    tlStf.iLogTime(1) = tlSdf.iTime(1)
                    tlStf.sPrint = "R"  'Ready to print
                    tlStf.iLen = tlSdf.iLen
                    tlStf.lSdfCode = tlSdf.lCode
                    tlStf.iRotNo = ilRotNo
                    tlStf.iUrfCode = tgUrf(0).iCode
                    ilRet = btrInsert(hlStf, tlStf, ilStfRecLen, INDEXKEY0)
                    If ilRet <> BTRV_ERR_NONE Then
                        igBtrError = gConvertErrorCode(ilRet)
                        sgErrLoc = "gMakeTracer-Insert Stf(2)"
                        gMakeTracer = False
                        Exit Function
                    End If
                    If slReason = "U" Then
                        slTracer = "4"
                    ElseIf slReason = "C" Then
                        slTracer = "3"
                    ElseIf slReason = "P" Then
                        slTracer = "2"
                    Else
                        slTracer = "1"
                    End If
                Else
                    slTracer = slReason
                End If
            '6/30/06:  Part of change
            'Else
            '    slTracer = slReason
            'End If
            '6/30/06:  End of this part of the change
        Else
            slTracer = slReason
        End If
    If ((slStatus = "M") Or (slStatus = "C") Or (slStatus = "S") Or (slStatus = "H")) And (tlSdf.sTracer <> "*") Then
        'If tlSdf.sTracer <> "1" Then
            Do
                tlSdf.sTracer = slTracer
                tlSdf.iUrfCode = tgUrf(0).iCode
                ilRet = btrUpdate(hlSdf, tlSdf, ilSdfRecLen)
                If ilRet = BTRV_ERR_CONFLICT Then
                    If llSdfRecPos > 0 Then
                        ilCRet = btrGetDirect(hlSdf, tlSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    Else
                        tmSdfSrchKey3.lCode = tlSdf.lCode
                        ilCRet = btrGetEqual(hlSdf, tlSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    End If
                    'tmSRec = tlSdf
                    'ilCRet = gGetByKeyForUpdate("Sdf", hlSdf, tmSRec)
                    'tlSdf = tmSRec
                    'If ilCRet <> BTRV_ERR_NONE Then
                    '    igBtrError = ilCRet
                    '    sgErrLoc = "gMakeTracer-GetByKey Sdf(3)"
                    '    gMakeTracer = False
                    '    Exit Function
                    'End If
                End If
            Loop While ilRet = BTRV_ERR_CONFLICT
        'End If
    End If
    gMakeTracer = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainSsfForDateOrGame               *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain Ssf for date and hour    *
'*                                                     *
'*******************************************************
Function gObtainSsfForDateOrGame(ilVefCode As Integer, llSsfDate As Long, slFindTime As String, ilGameNo As Integer, hlSsf As Integer, tlSsf As SSF, llSsfMemDate As Long, llSsfRecPos As Long) As Integer
'
'   ilRet = gObtaimSsfForDate(ilVefCode, llSsfDate, slFindTime, ilGameNo, hlSsf, tlSsf, llSsfMemDate, llSsfRecPos)
'   Where:
'       ilVefCode(I)- vehicle code
'       llSsfDate(I)- Date to obtain Ssf for
'       slFindTime(I)- Time to obtain Ssf for
'       ilGameNo(I)- Game number or zero to use date and time
'       hlSsf(I)- Handle from Ssf open
'       tlSsf(I/O)- Ssf record image
'       llSsfMemDate(I/O)- Date of Ssf within tlSsf (this is used instead of converting date within tlSsf for speed)
'       llSsfRecPos(O)- Ssf record position (btrGetDirect)
'       ilRet(O)- True if Ssf obtained
'                 False if Ssf not found
'
    Dim slDate As String        'Date as a string
    Dim ilDate0 As Integer      'Pack date
    Dim ilDate1 As Integer      'Pack date
    Dim slTime As String
    Dim ilTime0 As Integer      'Pack date
    Dim ilTime1 As Integer      'Pack date
    Dim ilSsfInMem As Integer
    Dim ilRet As Integer
    Dim clSsfFromTime As Currency
    Dim clSsfToTime As Currency
    Dim clHourTime As Currency  'Hour to be processed as currency
    Dim ilSsfOk As Integer
    '5/6/11
    Dim blGetDirectFailed As Boolean

    'If ilGameNo = 0 Then
        slDate = Format$(llSsfDate, "m/d/yy")
        gPackDate slDate, ilDate0, ilDate1
    'End If
    clHourTime = gTimeToCurrency(slFindTime, False)
    ilSsfInMem = False
    ''5/6/11: add test of llSsfRecPos
    ''3/15/13:  Add date test as part of game test
    ''If (llSsfRecPos = 0) Or ((llSsfMemDate = llSsfDate) And (ilGameNo = 0)) Or ((tlSsf.iType = ilGameNo) And (ilGameNo <> 0)) Then
    '6/5/14: Add vehicle test
    'If (llSsfRecPos = 0) Or ((llSsfMemDate = llSsfDate) And (ilGameNo = 0)) Or ((llSsfMemDate = llSsfDate) And (tlSsf.iType = ilGameNo) And (ilGameNo <> 0)) Then
    If (llSsfRecPos = 0) Or ((llSsfMemDate = llSsfDate) And (tlSsf.iVefCode = ilVefCode) And (ilGameNo = 0)) Or ((llSsfMemDate = llSsfDate) And (tlSsf.iVefCode = ilVefCode) And (tlSsf.iType = ilGameNo) And (ilGameNo <> 0)) Then
        gUnpackTime tlSsf.iStartTime(0), tlSsf.iStartTime(1), "A", "1", slTime
        clSsfFromTime = gTimeToCurrency(slTime, False)
        'If (tlSsf.iNextTime(0) = 1) And (tlSsf.iNextTime(1) = 0) Then
            clSsfToTime = gTimeToCurrency("12:00AM", True) - 1
        'Else
        '    gUnpackTime tlSsf.iNextTime(0), tlSsf.iNextTime(1), "A", "1", slTime
        '    clSsfToTime = gTimeToCurrency(slTime, True) - 1
        'End If
        If (tlSsf.iType = ilGameNo) And (tlSsf.iVefCode = ilVefCode) And (clHourTime >= clSsfFromTime) And (clHourTime <= clSsfToTime) Then
            ilSsfInMem = True
        End If
    Else
        clSsfToTime = gTimeToCurrency("12:00AM", True) - 1
    End If
    '5/6/11: Retry if GetDirect Fails
    Do
        blGetDirectFailed = False
        If Not ilSsfInMem Then
            gObtainSsfForDateOrGame = False
            llSsfMemDate = 0
            imSsfRecLen = Len(tlSsf) 'Max size of variable length record
            If ilGameNo = 0 Then
                tmSsfSrchKey.iType = 0 'slType
                tmSsfSrchKey.iVefCode = ilVefCode
                tmSsfSrchKey.iDate(0) = ilDate0
                tmSsfSrchKey.iDate(1) = ilDate1
                If clHourTime < clSsfToTime Then
                    ilTime0 = 0
                    ilTime1 = 0
                Else
                    slTime = gCurrencyToTime(clSsfToTime)
                    gPackTime slTime, ilTime0, ilTime1
                End If
                tmSsfSrchKey.iStartTime(0) = ilTime0
                tmSsfSrchKey.iStartTime(1) = ilTime1
                ilRet = gSSFGetGreaterOrEqual(hlSsf, tlSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                If (ilRet = BTRV_ERR_NONE) And (tlSsf.iType = 0) And (tlSsf.iVefCode = ilVefCode) And (tlSsf.iDate(0) = ilDate0) And (tlSsf.iDate(1) = ilDate1) Then
                    ilSsfOk = True
                Else
                    ilSsfOk = False
                End If
            Else
                'tmSsfSrchKey1.iVefCode = ilVefCode
                'tmSsfSrchKey1.iType = ilGameNo
                'ilRet = gSSFGetEqualKey1(hlSsf, tlSsf, imSsfRecLen, tmSsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
                'If (ilRet = BTRV_ERR_NONE) Then
                '    ilSsfOk = True
                'Else
                '    ilSsfOk = False
                'End If
                ilSsfOk = False
                tmSsfSrchKey2.iVefCode = ilVefCode
                tmSsfSrchKey2.iDate(0) = ilDate0
                tmSsfSrchKey2.iDate(1) = ilDate1
                ilRet = gSSFGetEqualKey2(hlSsf, tlSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)
                Do While (ilRet = BTRV_ERR_NONE) And (tlSsf.iVefCode = ilVefCode)
                    If (ilDate0 = tlSsf.iDate(0)) And (ilDate1 = tlSsf.iDate(1)) Then
                        If ilGameNo = tlSsf.iType Then
                            ilSsfOk = True
                            Exit Do
                        End If
                    Else
                        Exit Do
                    End If
                    ilRet = gSSFGetNext(hlSsf, tlSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
            Do While ilSsfOk
                gUnpackTime tlSsf.iStartTime(0), tlSsf.iStartTime(1), "A", "1", slTime
                clSsfFromTime = gTimeToCurrency(slTime, False)
                'If (tlSsf.iNextTime(0) = 1) And (tlSsf.iNextTime(1) = 0) Then
                    clSsfToTime = gTimeToCurrency("12:00AM", True)
                'Else
                '    gUnpackTime tlSsf.iNextTime(0), tlSsf.iNextTime(1), "A", "1", slTime
                '    clSsfToTime = gTimeToCurrency(slTime, True) - 1
                'End If
                If (clHourTime >= clSsfFromTime) And (clHourTime <= clSsfToTime) Then
                    If ilGameNo = 0 Then
                        llSsfMemDate = llSsfDate
                    Else
                        'gUnpackDateLong tlSsf.iDate(0), tlSsf.iDate(1), llSsfMemDate
                        '3/15/13: save date
                        'llSsfMemDate = ilGameNo
                        llSsfMemDate = llSsfDate
                    End If
                    gObtainSsfForDateOrGame = True
                    ilRet = gSSFGetPosition(hlSsf, llSsfRecPos)
                    If ilRet <> BTRV_ERR_NONE Then
                        igBtrError = gConvertErrorCode(ilRet)
                        sgErrLoc = "gObtainSsfForDateOrGame-Get Position Ssf(1)"
                        gObtainSsfForDateOrGame = False
                    End If
                    Exit Do
                End If
                'If (tlSsf.iNextTime(0) = 1) And (tlSsf.iNextTime(1) = 0) Then
                    Exit Do
                'Else
                '    imSsfRecLen = Len(tlSsf) 'Max size of variable length record
                '    ilRet = gSSFGetNext(hlSsf, tlSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                'End If
            Loop
        Else
            'Read back into memory so the current record is set for hlSsf
            imSsfRecLen = Len(tlSsf)
            ilRet = gSSFGetDirect(hlSsf, tlSsf, imSsfRecLen, llSsfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
            gObtainSsfForDateOrGame = True
            If ilRet <> BTRV_ERR_NONE Then
                '5/6/11
                clSsfToTime = gTimeToCurrency("12:00AM", True) - 1
                ilSsfInMem = False
                blGetDirectFailed = True
                'igBtrError = gConvertErrorCode(ilRet)
                'sgErrLoc = "gObtainSsfForDateOrGame-Get Direct Ssf(2)"
                'gObtainSsfForDateOrGame = False
            End If
        End If
    Loop While (blGetDirectFailed)
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainVcf                      *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the vehicle conflict     *
'*                     records for date specified      *
'*                                                     *
'*******************************************************
Sub gObtainVcf(hlVcf As Integer, ilVefCode As Integer, llDate As Long, tlVcf0() As VCF, tlVcf6() As VCF, tlVcf7() As VCF)
'
'   gObtainVcf llDate
'   Where:
'       hlVcf(I)- Vcf handle
'       ilVefCode(I)- Vehicle code
'       llDate(I)- Date within week to obtain Vcf records
'       ilVcfUsed (I/O)- True = Vcf Previously defined- check dates
'       tlVcf0(O)- Array of VCF records for M-F
'       tlVcf6(O)- Array of VCF records for Saturday
'       tlVcf7(O)- Array of VCF records for Sunday
'
    Dim ilRet As Integer
    Dim slDate As String
    Dim llEffDate As Long
    Dim llTermDate As Long
    Dim ilLoop As Integer
    Dim ilUpperBound As Integer
    Dim ilVcfRecLen As Integer
    Dim ilVcfDefined As Integer
    Dim llMoDate As Long
    Dim ilEffDate0 As Integer
    Dim ilEffDate1 As Integer
    Dim tlSrchKey As VCFKEY0
    If tgSpf.sHideGhostSptScr = "Y" Then
        ReDim tlVcf0(0 To 0) As VCF
        ReDim tlVcf6(0 To 0) As VCF
        ReDim tlVcf7(0 To 0) As VCF
        Exit Sub
    End If
    'Convert to Monday date so tlVlf for monday can only have an Monday date
    'Convert to Saturday date so tlVlf for saturday can only have an Monday date
    'Convert to Sunday date so tlVlf for sunday can only have an Monday date
    slDate = Format$(llDate, "m/d/yy")
    slDate = gObtainPrevMonday(slDate)
    llMoDate = gDateValue(slDate)
    For ilLoop = 0 To 2 Step 1
        ilVcfDefined = False
        Select Case ilLoop
            Case 0  'Monday thru friday
                'If UBound(tlVcf0) > 1 Then
                If UBound(tlVcf0) > 0 Then
                    If tlVcf0(0).iSellCode = ilVefCode Then
                        gUnpackDate tlVcf0(0).iEffDate(0), tlVcf0(0).iEffDate(1), slDate
                        llEffDate = gDateValue(slDate)
                        gUnpackDate tlVcf0(0).iTermDate(0), tlVcf0(0).iTermDate(1), slDate
                        If slDate = "" Then
                            slDate = "12/31/2060"
                        End If
                        llTermDate = gDateValue(slDate)
                        If (llMoDate >= llEffDate) And (llMoDate <= llTermDate) Then
                            ilVcfDefined = True
                        End If
                    End If
                End If
            Case 1  'Saturday
                'If UBound(tlVcf6) > 1 Then
                If UBound(tlVcf6) > 0 Then
                    If tlVcf6(0).iSellCode = ilVefCode Then
                        gUnpackDate tlVcf6(0).iEffDate(0), tlVcf6(0).iEffDate(1), slDate
                        llEffDate = gDateValue(slDate)
                        gUnpackDate tlVcf6(0).iTermDate(0), tlVcf6(0).iTermDate(1), slDate
                        If slDate = "" Then
                            slDate = "12/31/2060"
                        End If
                        llTermDate = gDateValue(slDate)
                        If (llMoDate + 5 >= llEffDate) And (llMoDate + 5 <= llTermDate) Then
                            ilVcfDefined = True
                        End If
                    End If
                End If
            Case 2  'Sunday
                'If UBound(tlVcf7) > 1 Then
                If UBound(tlVcf7) > 0 Then
                    If tlVcf7(0).iSellCode = ilVefCode Then
                        gUnpackDate tlVcf7(0).iEffDate(0), tlVcf7(0).iEffDate(1), slDate
                        llEffDate = gDateValue(slDate)
                        gUnpackDate tlVcf7(0).iTermDate(0), tlVcf7(0).iTermDate(1), slDate
                        If slDate = "" Then
                            slDate = "12/31/2060"
                        End If
                        llTermDate = gDateValue(slDate)
                        If (llMoDate + 6 >= llEffDate) And (llMoDate + 6 <= llTermDate) Then
                            ilVcfDefined = True
                        End If
                    End If
                End If
        End Select
        If Not ilVcfDefined Then
            Select Case ilLoop
                Case 0
                    ReDim tlVcf0(0 To 0) As VCF
                    ilUpperBound = UBound(tlVcf0)
                    ilVcfRecLen = Len(tlVcf0(0))
                    slDate = Format$(llMoDate, "m/d/yy")
                    gPackDate slDate, ilEffDate0, ilEffDate1
                    tlSrchKey.iSellCode = ilVefCode
                    tlSrchKey.iSellDay = 0  'Monday thru friday
                    tlSrchKey.iEffDate(0) = ilEffDate0
                    tlSrchKey.iEffDate(1) = ilEffDate1
                    tlSrchKey.iSellTime(0) = 0
                    tlSrchKey.iSellTime(1) = 6144   '24*256
                    tlSrchKey.iSellPosNo = 32000
                    ilRet = btrGetLessOrEqual(hlVcf, tlVcf0(ilUpperBound), ilVcfRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    If (ilRet = BTRV_ERR_NONE) And (tlVcf0(ilUpperBound).iSellCode = ilVefCode) And (tlVcf0(ilUpperBound).iSellDay = 0) Then
                        ilEffDate0 = tlVcf0(ilUpperBound).iEffDate(0)
                        ilEffDate1 = tlVcf0(ilUpperBound).iEffDate(1)
                        ilVcfRecLen = Len(tlVcf0(0))
                        tlSrchKey.iSellCode = ilVefCode
                        tlSrchKey.iSellDay = 0  'Monday thru friday
                        tlSrchKey.iEffDate(0) = ilEffDate0
                        tlSrchKey.iEffDate(1) = ilEffDate1
                        tlSrchKey.iSellTime(0) = 0
                        tlSrchKey.iSellTime(1) = 0
                        tlSrchKey.iSellPosNo = 0
                        ilRet = btrGetGreaterOrEqual(hlVcf, tlVcf0(ilUpperBound), ilVcfRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tlVcf0(ilUpperBound).iSellCode = ilVefCode) And (tlVcf0(ilUpperBound).iSellDay = 0)
                            gUnpackDate tlVcf0(ilUpperBound).iEffDate(0), tlVcf0(ilUpperBound).iEffDate(1), slDate
                            llEffDate = gDateValue(slDate)
                            gUnpackDate tlVcf0(ilUpperBound).iTermDate(0), tlVcf0(ilUpperBound).iTermDate(1), slDate
                            If slDate = "" Then
                                slDate = "12/31/2060"
                            End If
                            llTermDate = gDateValue(slDate)
                            If (llMoDate >= llEffDate) And (llMoDate <= llTermDate) Then
                                ilUpperBound = ilUpperBound + 1
                                ReDim Preserve tlVcf0(0 To ilUpperBound) As VCF
                            End If
                            ilRet = btrGetNext(hlVcf, tlVcf0(ilUpperBound), ilVcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                Case 1
                    ReDim tlVcf6(0 To 0) As VCF
                    ilUpperBound = UBound(tlVcf6)
                    ilVcfRecLen = Len(tlVcf6(0))
                    slDate = Format$(llMoDate + 5, "m/d/yy")
                    gPackDate slDate, ilEffDate0, ilEffDate1
                    tlSrchKey.iSellCode = ilVefCode
                    tlSrchKey.iSellDay = 6  'Saturday
                    tlSrchKey.iEffDate(0) = ilEffDate0
                    tlSrchKey.iEffDate(1) = ilEffDate1
                    tlSrchKey.iSellTime(0) = 0
                    tlSrchKey.iSellTime(1) = 6144   '24*256
                    tlSrchKey.iSellPosNo = 32000
                    ilRet = btrGetLessOrEqual(hlVcf, tlVcf6(ilUpperBound), ilVcfRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    If (ilRet = BTRV_ERR_NONE) And (tlVcf6(ilUpperBound).iSellCode = ilVefCode) And (tlVcf6(ilUpperBound).iSellDay = 6) Then
                        ilEffDate0 = tlVcf6(ilUpperBound).iEffDate(0)
                        ilEffDate1 = tlVcf6(ilUpperBound).iEffDate(1)
                        tlSrchKey.iSellCode = ilVefCode
                        tlSrchKey.iSellDay = 6  'Saturday
                        tlSrchKey.iEffDate(0) = ilEffDate0
                        tlSrchKey.iEffDate(1) = ilEffDate1
                        tlSrchKey.iSellTime(0) = 0
                        tlSrchKey.iSellTime(1) = 0
                        tlSrchKey.iSellPosNo = 0
                        ilRet = btrGetGreaterOrEqual(hlVcf, tlVcf6(ilUpperBound), ilVcfRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tlVcf6(ilUpperBound).iSellCode = ilVefCode) And (tlVcf6(ilUpperBound).iSellDay = 6)
                            gUnpackDate tlVcf6(ilUpperBound).iEffDate(0), tlVcf6(ilUpperBound).iEffDate(1), slDate
                            llEffDate = gDateValue(slDate)
                            gUnpackDate tlVcf6(ilUpperBound).iTermDate(0), tlVcf6(ilUpperBound).iTermDate(1), slDate
                            If slDate = "" Then
                                slDate = "12/31/2060"
                            End If
                            llTermDate = gDateValue(slDate)
                            If (llMoDate + 5 >= llEffDate) And (llMoDate + 5 <= llTermDate) Then
                                ilUpperBound = ilUpperBound + 1
                                ReDim Preserve tlVcf6(0 To ilUpperBound) As VCF
                            End If
                            ilRet = btrGetNext(hlVcf, tlVcf6(ilUpperBound), ilVcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                Case 2
                    ReDim tlVcf7(0 To 0) As VCF
                    ilUpperBound = UBound(tlVcf7)
                    ilVcfRecLen = Len(tlVcf7(0))
                    slDate = Format$(llMoDate + 6, "m/d/yy")
                    gPackDate slDate, ilEffDate0, ilEffDate1
                    tlSrchKey.iSellCode = ilVefCode
                    tlSrchKey.iSellDay = 7  'Sunday
                    tlSrchKey.iEffDate(0) = ilEffDate0
                    tlSrchKey.iEffDate(1) = ilEffDate1
                    tlSrchKey.iSellTime(0) = 0
                    tlSrchKey.iSellTime(1) = 6144   '24*256
                    tlSrchKey.iSellPosNo = 32000
                    ilRet = btrGetLessOrEqual(hlVcf, tlVcf7(ilUpperBound), ilVcfRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    If (ilRet = BTRV_ERR_NONE) And (tlVcf7(ilUpperBound).iSellCode = ilVefCode) And (tlVcf7(ilUpperBound).iSellDay = 7) Then
                        ilEffDate0 = tlVcf7(ilUpperBound).iEffDate(0)
                        ilEffDate1 = tlVcf7(ilUpperBound).iEffDate(1)
                        tlSrchKey.iSellCode = ilVefCode
                        tlSrchKey.iSellDay = 7  'Sunday
                        tlSrchKey.iEffDate(0) = 0
                        tlSrchKey.iEffDate(1) = 0
                        tlSrchKey.iSellTime(0) = 0
                        tlSrchKey.iSellTime(1) = 0
                        tlSrchKey.iSellPosNo = 0
                        ilRet = btrGetGreaterOrEqual(hlVcf, tlVcf7(ilUpperBound), ilVcfRecLen, tlSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tlVcf7(ilUpperBound).iSellCode = ilVefCode) And (tlVcf7(ilUpperBound).iSellDay = 7)
                            gUnpackDate tlVcf7(ilUpperBound).iEffDate(0), tlVcf7(ilUpperBound).iEffDate(1), slDate
                            llEffDate = gDateValue(slDate)
                            gUnpackDate tlVcf7(ilUpperBound).iTermDate(0), tlVcf7(ilUpperBound).iTermDate(1), slDate
                            If slDate = "" Then
                                slDate = "12/31/2060"
                            End If
                            llTermDate = gDateValue(slDate)
                            If (llMoDate + 6 >= llEffDate) And (llMoDate + 6 <= llTermDate) Then
                                ilUpperBound = ilUpperBound + 1
                                ReDim Preserve tlVcf7(0 To ilUpperBound) As VCF
                            End If
                            ilRet = btrGetNext(hlVcf, tlVcf7(ilUpperBound), ilVcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
            End Select
        End If
    Next ilLoop
    Exit Sub
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gObtainVlf                      *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the vehicle link         *
'*                     records for date specified      *
'*                                                     *
'*******************************************************
Sub gObtainVlf(slType As String, hlVlf As Integer, ilVefCode As Integer, llDate As Long, tlVlf() As VLF)
'
'   gObtainVcf llDate
'   Where:
'       slType(I) "S" = Selling; "A" = Airing
'       hlVlf(I)- Vcf handle
'       ilVefCode(I)- Vehicle code
'       llDate(I)- Date within week to obtain Vcf records
'       tlVlf(O)- Array of VLF records
'
    Dim ilRet As Integer
    Dim slDate As String
    Dim llEffDate As Long
    Dim llTermDate As Long
    Dim ilDay As Integer
    Dim ilUpperBound As Integer
    Dim ilLowerBound As Integer
    Dim ilVlfRecLen As Integer
    Dim ilVlfDefined As Integer
    Dim ilEffDate0 As Integer
    Dim ilEffDate1 As Integer
    Dim ilTerminated As Integer
    Dim tlSrchKey0 As VLFKEY0
    Dim tlSrchKey1 As VLFKEY1
    'Convert to Monday date so tlVlf for monday can only have an Monday date
    'Convert to Saturday date so tlVlf for saturday can only have an Monday date
    'Convert to Sunday date so tlVlf for sunday can only have an Monday date
    slDate = Format$(llDate, "m/d/yy")
    gPackDate slDate, ilEffDate0, ilEffDate1
    ilVlfDefined = False
    ilDay = gWeekDayLong(llDate)
    If ilDay = 6 Then
        ilDay = 7
    ElseIf ilDay = 5 Then   'Saturady
        ilDay = 6
    Else
        ilDay = 0
    End If
    ilLowerBound = LBound(tlVlf)
    If UBound(tlVlf) > ilLowerBound Then
        If slType = "S" Then
            If (tlVlf(ilLowerBound).iSellCode = ilVefCode) And (tlVlf(ilLowerBound).iSellDay = ilDay) Then
                gUnpackDate tlVlf(ilLowerBound).iEffDate(0), tlVlf(ilLowerBound).iEffDate(1), slDate
                llEffDate = gDateValue(slDate)
                gUnpackDate tlVlf(ilLowerBound).iTermDate(0), tlVlf(ilLowerBound).iTermDate(1), slDate
                If slDate = "" Then
                    slDate = "12/31/2060"
                End If
                llTermDate = gDateValue(slDate)
                If (llDate >= llEffDate) And (llDate <= llTermDate) Then
                    ilVlfDefined = True
                End If
            End If
        Else
            If (tlVlf(ilLowerBound).iAirCode = ilVefCode) And (tlVlf(ilLowerBound).iAirDay = ilDay) Then
                gUnpackDate tlVlf(ilLowerBound).iEffDate(0), tlVlf(ilLowerBound).iEffDate(1), slDate
                llEffDate = gDateValue(slDate)
                gUnpackDate tlVlf(ilLowerBound).iTermDate(0), tlVlf(ilLowerBound).iTermDate(1), slDate
                If slDate = "" Then
                    slDate = "12/31/2060"
                End If
                llTermDate = gDateValue(slDate)
                If (llDate >= llEffDate) And (llDate <= llTermDate) Then
                    ilVlfDefined = True
                End If
            End If
        End If
    End If
    If Not ilVlfDefined Then
        ReDim tlVlf(ilLowerBound To ilLowerBound) As VLF
        ilUpperBound = UBound(tlVlf)
        ilVlfRecLen = Len(tlVlf(ilLowerBound))
        'Determine effective date
        If slType = "S" Then
            tlSrchKey0.iSellCode = ilVefCode
            tlSrchKey0.iSellDay = ilDay
            tlSrchKey0.iEffDate(0) = ilEffDate0
            tlSrchKey0.iEffDate(1) = ilEffDate1
            ilEffDate0 = 0
            ilEffDate1 = 0
            tlSrchKey0.iSellTime(0) = 0
            tlSrchKey0.iSellTime(1) = 6144  '24*256
            tlSrchKey0.iSellPosNo = 32000
            ilRet = btrGetLessOrEqual(hlVlf, tlVlf(ilUpperBound), ilVlfRecLen, tlSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tlVlf(ilUpperBound).iSellCode = ilVefCode)
                ilTerminated = False
                'Check for CBS
                If (tlVlf(ilUpperBound).iTermDate(1) <> 0) Or (tlVlf(ilUpperBound).iTermDate(0) <> 0) Then
                    If (tlVlf(ilUpperBound).iTermDate(1) < tlVlf(ilUpperBound).iEffDate(1)) Or ((tlVlf(ilUpperBound).iEffDate(1) = tlVlf(ilUpperBound).iTermDate(1)) And (tlVlf(ilUpperBound).iTermDate(0) < tlVlf(ilUpperBound).iEffDate(0))) Then
                        ilTerminated = True
                    End If
                End If
                If (tlVlf(ilUpperBound).sStatus <> "P") And (tlVlf(ilUpperBound).iSellDay = ilDay) And (Not ilTerminated) Then
                    ilEffDate0 = tlVlf(ilUpperBound).iEffDate(0)
                    ilEffDate1 = tlVlf(ilUpperBound).iEffDate(1)
                    Exit Do
                End If
                ilRet = btrGetPrevious(hlVlf, tlVlf(ilUpperBound), ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            'If (ilRet = BTRV_ERR_NONE) And (tlVlf(ilUpperBound).iSellCode = ilVefCode) And (tlVlf(ilUpperBound).iSellDay = ilDay) And (tlVlf(ilUpperBound).sStatus = "C") Then
            '    ilEffDate0 = tlVlf(ilUpperBound).iEffDate(0)
            '    ilEffDate1 = tlVlf(ilUpperBound).iEffDate(1)
            'Else
            '    ilEffDate0 = 0
            '    ilEffDate1 = 0
            'End If
        Else
            tlSrchKey1.iAirCode = ilVefCode
            tlSrchKey1.iAirDay = ilDay
            tlSrchKey1.iEffDate(0) = ilEffDate0
            tlSrchKey1.iEffDate(1) = ilEffDate1
            ilEffDate0 = 0
            ilEffDate1 = 0
            tlSrchKey1.iAirTime(0) = 0
            tlSrchKey1.iAirTime(1) = 6144   '24*256
            tlSrchKey1.iAirPosNo = 32000
            ilRet = btrGetLessOrEqual(hlVlf, tlVlf(ilUpperBound), ilVlfRecLen, tlSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tlVlf(ilUpperBound).iAirCode = ilVefCode)
                ilTerminated = False
                'Check for CBS
                If (tlVlf(ilUpperBound).iTermDate(1) <> 0) Or (tlVlf(ilUpperBound).iTermDate(0) <> 0) Then
                    If (tlVlf(ilUpperBound).iTermDate(1) < tlVlf(ilUpperBound).iEffDate(1)) Or ((tlVlf(ilUpperBound).iEffDate(1) = tlVlf(ilUpperBound).iTermDate(1)) And (tlVlf(ilUpperBound).iTermDate(0) < tlVlf(ilUpperBound).iEffDate(0))) Then
                        ilTerminated = True
                    End If
                End If
                If (tlVlf(ilUpperBound).sStatus <> "P") And (tlVlf(ilUpperBound).iAirDay = ilDay) And (Not ilTerminated) Then
                    ilEffDate0 = tlVlf(ilUpperBound).iEffDate(0)
                    ilEffDate1 = tlVlf(ilUpperBound).iEffDate(1)
                    Exit Do
                End If
                ilRet = btrGetPrevious(hlVlf, tlVlf(ilUpperBound), ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            'If (ilRet = BTRV_ERR_NONE) And (tlVlf(ilUpperBound).iSellCode = ilVefCode) And (tlVlf(ilUpperBound).iSellDay = ilDay) And (tlVlf(ilUpperBound).sStatus = "C") Then
            '    ilEffDate0 = tlVlf(ilUpperBound).iEffDate(0)
            '    ilEffDate1 = tlVlf(ilUpperBound).iEffDate(1)
            'Else
            '    ilEffDate0 = 0
            '    ilEffDate1 = 0
            'End If
        End If
        If slType = "S" Then
            tlSrchKey0.iSellCode = ilVefCode
            tlSrchKey0.iSellDay = ilDay
            tlSrchKey0.iEffDate(0) = ilEffDate0
            tlSrchKey0.iEffDate(1) = ilEffDate1
            tlSrchKey0.iSellTime(0) = 0
            tlSrchKey0.iSellTime(1) = 0
            tlSrchKey0.iSellPosNo = 0
            ilRet = btrGetGreaterOrEqual(hlVlf, tlVlf(ilUpperBound), ilVlfRecLen, tlSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tlVlf(ilUpperBound).iSellCode = ilVefCode) And (tlVlf(ilUpperBound).iSellDay = ilDay)
                If tlVlf(ilUpperBound).sStatus = "C" Then
                    gUnpackDate tlVlf(ilUpperBound).iEffDate(0), tlVlf(ilUpperBound).iEffDate(1), slDate
                    llEffDate = gDateValue(slDate)
                    gUnpackDate tlVlf(ilUpperBound).iTermDate(0), tlVlf(ilUpperBound).iTermDate(1), slDate
                    If slDate = "" Then
                        slDate = "12/31/2060"
                    End If
                    llTermDate = gDateValue(slDate)
                    If (llDate >= llEffDate) And (llDate <= llTermDate) Then
                        ilUpperBound = ilUpperBound + 1
                        ReDim Preserve tlVlf(ilLowerBound To ilUpperBound) As VLF
                    Else
                        If llDate < llEffDate Then
                            Exit Do
                        End If
                    End If
                End If
                ilRet = btrGetNext(hlVlf, tlVlf(ilUpperBound), ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Else
            tlSrchKey1.iAirCode = ilVefCode
            tlSrchKey1.iAirDay = ilDay
            tlSrchKey1.iEffDate(0) = ilEffDate0
            tlSrchKey1.iEffDate(1) = ilEffDate1
            tlSrchKey1.iAirTime(0) = 0
            tlSrchKey1.iAirTime(1) = 0
            tlSrchKey1.iAirPosNo = 0
            ilRet = btrGetGreaterOrEqual(hlVlf, tlVlf(ilUpperBound), ilVlfRecLen, tlSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tlVlf(ilUpperBound).iAirCode = ilVefCode) And (tlVlf(ilUpperBound).iAirDay = ilDay)
                If tlVlf(ilUpperBound).sStatus = "C" Then
                    gUnpackDate tlVlf(ilUpperBound).iEffDate(0), tlVlf(ilUpperBound).iEffDate(1), slDate
                    llEffDate = gDateValue(slDate)
                    gUnpackDate tlVlf(ilUpperBound).iTermDate(0), tlVlf(ilUpperBound).iTermDate(1), slDate
                    If slDate = "" Then
                        slDate = "12/31/2060"
                    End If
                    llTermDate = gDateValue(slDate)
                    If (llDate >= llEffDate) And (llDate <= llTermDate) Then
                        ilUpperBound = ilUpperBound + 1
                        ReDim Preserve tlVlf(ilLowerBound To ilUpperBound) As VLF
                    Else
                        If llDate < llEffDate Then
                            Exit Do
                        End If
                    End If
                End If
                ilRet = btrGetNext(hlVlf, tlVlf(ilUpperBound), ilVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:gPreemptible                    *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Test for spot can be preempted  *
'*                                                     *
'*******************************************************
Function gPreemptible(ilSchMode As Integer, tlSpotMove() As SPOTMOVE, tlAvail As AVAILSS, tlSpotTest As CSPOTSS, ilBkQH As Integer, slInOut As String, slPreempt As String, ilPriceLevel As Integer) As Integer
    Dim ilCheck As Integer
    Dim ilSpotPriceLevel As Integer

    If sgApplyVehPreemptRule = "Y" Then
        If sgVehPreemptRule = "Y" Then
            gPreemptible = False
            Exit Function
        End If
    End If
    ilCheck = False
    If ilSchMode = 3 Then
        'If (UBound(tlSpotMove) = 1) And ((tlSpotTest.iRank And RANKMASK) <> RESERVATION) And ((tlAvail.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
        If (UBound(tlSpotMove) = LBound(tlSpotMove)) And ((tlSpotTest.iRank And RANKMASK) <> RESERVATIONRANK) And ((tlAvail.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
            ilCheck = True
        End If
    End If
    If ilSchMode = 4 Then
        'If ((UBound(tlSpotMove) = 1) Or (UBound(tlSpotMove) = 2)) And ((tlSpotTest.iRank And RANKMASK) = 2000) And ((tlAvail.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
        If ((UBound(tlSpotMove) = LBound(tlSpotMove)) Or (UBound(tlSpotMove) = LBound(tlSpotMove) + 1)) And ((tlSpotTest.iRank And RANKMASK) = 2000) And ((tlAvail.iAvInfo And SSLOCKSPOT) <> SSLOCKSPOT) Then
            ilCheck = True
        End If
    End If

    If ilCheck Then
        '6/20/06:  $ spots have a high priority then $0 and N/C and Fill spots
        ilSpotPriceLevel = (tlSpotTest.iRank And PRICELEVELMASK) / SHIFT11
        If (ilPriceLevel > ilSpotPriceLevel) And (ilSpotPriceLevel <= 1) Then
            gPreemptible = True
            Exit Function
        End If
        If (ilPriceLevel < ilSpotPriceLevel) And (ilPriceLevel <= 1) Then
            gPreemptible = False
            Exit Function
        End If
        If ((tlSpotTest.iRank And RANKMASK) > ilBkQH) Then
            gPreemptible = True
            Exit Function
        ElseIf (tlSpotTest.iRank And RANKMASK) = ilBkQH Then
            If (slInOut = "I") Or ((tlSpotTest.iRecType And SSAVAILBUY) = SSAVAILBUY) Then
                If (slInOut = "I") And ((tlSpotTest.iRecType And SSAVAILBUY) <> SSAVAILBUY) Then
                    gPreemptible = True
                    Exit Function
                End If
                If (slInOut <> "I") And ((tlSpotTest.iRecType And SSAVAILBUY) = SSAVAILBUY) Then
                    gPreemptible = False
                    Exit Function
                End If
            End If
            If (slInOut = "O") Or ((tlSpotTest.iRecType And SSEXAVAILBUY) = SSEXAVAILBUY) Then
                If (slInOut = "O") And ((tlSpotTest.iRecType And SSEXAVAILBUY) <> SSEXAVAILBUY) Then
                    gPreemptible = True
                    Exit Function
                End If
                If (slInOut <> "O") And ((tlSpotTest.iRecType And SSEXAVAILBUY) = SSEXAVAILBUY) Then
                    gPreemptible = False
                    Exit Function
                End If
            End If
            'Non-Preemptible test is higher then the Price Level.  This allows the Low/High to work as a override
            If (slPreempt <> "P") Or ((tlSpotTest.iRecType And SSPREEMPTIBLE) <> SSPREEMPTIBLE) Then
                If (slPreempt = "P") And ((tlSpotTest.iRecType And SSPREEMPTIBLE) <> SSPREEMPTIBLE) Then
                    gPreemptible = False
                    Exit Function
                End If
                If (slPreempt <> "P") And ((tlSpotTest.iRecType And SSPREEMPTIBLE) = SSPREEMPTIBLE) Then
                    gPreemptible = True
                    Exit Function
                End If
            End If
            'Price test
            If ilPriceLevel > ilSpotPriceLevel Then
                gPreemptible = True
                Exit Function
            End If
            gPreemptible = False
            Exit Function
        Else
            gPreemptible = False
            Exit Function
        End If
    Else
        gPreemptible = False
        Exit Function
    End If
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gPrgToDelete                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Add Log library to delete       *
'*                     library calendar date           *
'*                                                     *
'*******************************************************
Function gPrgToDelete(frm As Form, tlLvf As LVF, ilType As Integer) As Integer
'
'    ilRet = gPrgToDelete(MainForm, tlLvf, slType)
'    Where:
'       MainForm (I)- Name of Form to unload if error exists
'       tlLvf (I) - Library to be removed into pending calendar
'
'       ilType (I)- Type of library (0=Regular Programming; 1->NN = Sports Programming (Game Number))
'
'       tgRPrg(I)- contains dates/times to remove libraries
'
    Dim hlLcf As Integer            'Log calendar library file handle
    Dim hlDLcf As Integer            'Log calendar library file handle
    Dim hlLvf As Integer            'Log library file handle
    Dim hlLtf As Integer            'Log library file handle
    Dim tlCLcf As LCF               'LCF record image
    Dim tlDLcf As LCF               'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim tlLtf As LTF               'LTF record image
    Dim tlLtfSrchKey As INTKEY0     'LTF key record image
    Dim ilLtfRecLen As Integer         'LTF record length
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim llLatestDate As Long    'Latest date of libraries to be inserted
    Dim llPLatestDate As Long   'Pending latest date to be checked to see if pending must be extended
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilMatch As Integer
    Dim llDRecPos As Long
    Dim llCRecPos As Long
    ReDim ilStartTime(0 To 1) As Integer
    hlLcf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRemovePrgErr
    gBtrvErrorMsg ilRet, "gRemovePrg (btrOpen: Lcf.btr)", frm
    On Error GoTo 0
    hlDLcf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlDLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRemovePrgErr
    gBtrvErrorMsg ilRet, "gRemovePrg (btrOpen: Lcf.btr)", frm
    On Error GoTo 0
    hlLvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRemovePrgErr
    gBtrvErrorMsg ilRet, "gRemovePrg (btrOpen: Lvf.btr)", frm
    On Error GoTo 0
    hlLtf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlLtf, "", sgDBPath & "Ltf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRemovePrgErr
    gBtrvErrorMsg ilRet, "gRemovePrg (btrOpen: Ltf.btr)", frm
    On Error GoTo 0
    ilLcfRecLen = Len(tlCLcf)
    ilLtfRecLen = Len(tlLtf)
    llLatestDate = -1
    llPLatestDate = -1
'    If igViewType = 1 Then
'        tlLcfSrchKey.sType = "A"
'        slType = "A"
'    Else
'        tlLcfSrchKey.sType = "O"
'        slType = "O"
'    End If
    tlLtfSrchKey.iCode = tlLvf.iLtfCode
    ilRet = btrGetEqual(hlLtf, tlLtf, ilLtfRecLen, tlLtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
    On Error GoTo gRemovePrgErr
    gBtrvErrorMsg ilRet, "gRemovePrg (btrGetEqual: Ltf.btr)", frm
    On Error GoTo 0
    tlLcfSrchKey.iType = ilType
    tlLcfSrchKey.sStatus = "C"
    tlLcfSrchKey.iVefCode = tlLtf.iVefCode
    tlLcfSrchKey.iLogDate(0) = 257  'Year 1/1/2100
    tlLcfSrchKey.iLogDate(1) = 2100
    tlLcfSrchKey.iSeqNo = 1
    ilRet = btrGetLessOrEqual(hlLcf, tlCLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    If (ilRet = BTRV_ERR_NONE) And (tlCLcf.sStatus = "C") And (tlCLcf.iVefCode = tlLtf.iVefCode) And (tlCLcf.iType = ilType) Then
        gUnpackDate tlCLcf.iLogDate(0), tlCLcf.iLogDate(1), slDate
        llLatestDate = gDateValue(slDate)
    End If
    tlLcfSrchKey.sStatus = "P"
    tlLcfSrchKey.iVefCode = tlLtf.iVefCode
    tlLcfSrchKey.iLogDate(0) = 257  'Year 1/1/2100
    tlLcfSrchKey.iLogDate(1) = 2100
    tlLcfSrchKey.iSeqNo = 1
    ilRet = btrGetLessOrEqual(hlLcf, tlCLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    If (ilRet = BTRV_ERR_NONE) And (tlCLcf.sStatus = "P") And (tlCLcf.iVefCode = tlLtf.iVefCode) And (tlCLcf.iType = ilType) Then
        gUnpackDate tlCLcf.iLogDate(0), tlCLcf.iLogDate(1), slDate
        llDate = gDateValue(slDate)
        llPLatestDate = llDate
        If llDate > llLatestDate Then
            llLatestDate = llDate
        End If
    End If
    For ilLoop = LBound(tgRPrg) To UBound(tgRPrg) - 1 Step 1
        If (tgRPrg(ilLoop).sStartTime <> "") And (tgRPrg(ilLoop).sStartDate <> "") Then
            If tgRPrg(ilLoop).sEndDate <> "" Then
                llStartDate = gDateValue(tgRPrg(ilLoop).sStartDate)
'                If slType = "A" Then
'                    llEndDate = llStartDate
'                Else
                    llEndDate = gDateValue(tgRPrg(ilLoop).sEndDate)
'                End If
                'Date must exist within LCF- so extend not required
                'If (llStartDate > llPLatestDate) And (llPLatestDate <> -1) Then
                '    For llDate = llPLatestDate + 1 To llStartDate - 1 Step 1
                '        slDate = gFormatDate(Format$(llDate, "m/d/yy"))
                '        gPackDate slDate, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1)
                '        gExtendTFN hlLcf, 0, 0, 0, "P", tlLtf.iVefCode, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1)
                '    Next llDate
                'End If
                For llDate = llStartDate To llEndDate Step 1
                    If tgRPrg(ilLoop).iDay(gWeekDayLong(llDate)) = 1 Then
                        slDate = gFormatDate(Format$(llDate, "m/d/yy"))
                        gPackDate slDate, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1)
                        gPackTime tgRPrg(ilLoop).sStartTime, ilStartTime(0), ilStartTime(1)
                        ilRet = mAddPrgToDelete(hlLcf, hlLvf, tlLvf, ilType, tlLtf.iVefCode, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1), ilStartTime(0), ilStartTime(1))
                        If ilRet <> BTRV_ERR_NONE Then
                            igBtrError = gConvertErrorCode(ilRet)
                            sgErrLoc = "gPrgToDelete-mAddPrgToDelete(1)"
                            btrDestroy hlLtf
                            btrDestroy hlLvf
                            btrDestroy hlDLcf
                            btrDestroy hlLcf
                            gPrgToDelete = False
                            Exit Function
                        End If
                    End If
                Next llDate
            Else    'Add program To TFN week
                For ilIndex = 1 To 7 Step 1
                    If tgRPrg(ilLoop).iDay(ilIndex - 1) = 1 Then
                        gPackTime tgRPrg(ilLoop).sStartTime, ilStartTime(0), ilStartTime(1)
                        ilRet = mAddPrgToDelete(hlLcf, hlLvf, tlLvf, ilType, tlLtf.iVefCode, ilIndex, 0, ilStartTime(0), ilStartTime(1))
                        If ilRet <> BTRV_ERR_NONE Then
                            igBtrError = gConvertErrorCode(ilRet)
                            sgErrLoc = "gPrgToDelete-mAddPrgToDelete(1)"
                            btrDestroy hlLtf
                            btrDestroy hlLvf
                            btrDestroy hlDLcf
                            btrDestroy hlLcf
                            gPrgToDelete = False
                            Exit Function
                        End If
                    End If
                Next ilIndex
                llStartDate = gDateValue(tgRPrg(ilLoop).sStartDate)
                If llLatestDate <> -1 Then
                    llEndDate = llLatestDate
                    If llEndDate < llStartDate + 6 Then
                        llEndDate = llStartDate + 6
                    End If
                Else
                    llEndDate = llStartDate + 6
                End If
                'Date must exist within LCF- so extend not required
                'If (llStartDate > llPLatestDate) And (llPLatestDate <> -1) Then
                '    For llDate = llPLatestDate + 1 To llStartDate - 1 Step 1
                '        slDate = gFormatDate(Format$(llDate, "m/d/yy"))
                '        gPackDate slDate, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1)
                '        gExtendTFN hlLcf, 0, 0, 0, "P", tlLtf.iVefCode, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1)
                '    Next llDate
                'End If
                For llDate = llStartDate To llEndDate Step 1
                    If tgRPrg(ilLoop).iDay(gWeekDayLong(llDate)) = 1 Then
                        slDate = gFormatDate(Format$(llDate, "m/d/yy"))
                        gPackDate slDate, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1)
                        gPackTime tgRPrg(ilLoop).sStartTime, ilStartTime(0), ilStartTime(1)
                        ilRet = mAddPrgToDelete(hlLcf, hlLvf, tlLvf, ilType, tlLtf.iVefCode, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1), ilStartTime(0), ilStartTime(1))
                        If ilRet <> BTRV_ERR_NONE Then
                            igBtrError = gConvertErrorCode(ilRet)
                            sgErrLoc = "gPrgToDelete-mAddPrgToDelete(1)"
                            btrDestroy hlLtf
                            btrDestroy hlLvf
                            btrDestroy hlDLcf
                            btrDestroy hlLcf
                            gPrgToDelete = False
                            Exit Function
                        End If
                    End If
                Next llDate
            End If
        End If
    Next ilLoop
    'Compare Pending and delete records- if same remove both
    tlLcfSrchKey.iType = ilType
    tlLcfSrchKey.sStatus = "D"
    tlLcfSrchKey.iVefCode = tlLtf.iVefCode
    tlLcfSrchKey.iLogDate(0) = 1    '257  'Year 1/1/1900
    tlLcfSrchKey.iLogDate(1) = 0    '1900
    tlLcfSrchKey.iSeqNo = 1
    ilRet = btrGetGreaterOrEqual(hlDLcf, tlDLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tlDLcf.sStatus = "D") And (tlDLcf.iVefCode = tlLtf.iVefCode) And (tlDLcf.iType = ilType)
        ilRet = btrGetPosition(hlDLcf, llDRecPos)
        tlLcfSrchKey.iType = ilType
        tlLcfSrchKey.sStatus = "P"
        tlLcfSrchKey.iVefCode = tlLtf.iVefCode
        tlLcfSrchKey.iLogDate(0) = tlDLcf.iLogDate(0)  'Year 1/1/1900
        tlLcfSrchKey.iLogDate(1) = tlDLcf.iLogDate(1)
        tlLcfSrchKey.iSeqNo = 1
        ilRet = btrGetEqual(hlLcf, tlCLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        If ilRet = BTRV_ERR_NONE Then
            ilRet = btrGetPosition(hlLcf, llCRecPos)
            If (tlDLcf.iLastTime(0) = tlCLcf.iLastTime(0)) And (tlDLcf.iLastTime(1) = tlCLcf.iLastTime(1)) Then
                ilMatch = True
                'For ilLoop = 1 To 50 Step 1
                For ilLoop = LBound(tlDLcf.lLvfCode) To UBound(tlDLcf.lLvfCode) Step 1
                    If tlDLcf.lLvfCode(ilLoop) <> tlCLcf.lLvfCode(ilLoop) Then
                        ilMatch = False
                        Exit For
                    End If
                    If (tlDLcf.iTime(0, ilLoop) <> tlCLcf.iTime(0, ilLoop)) And (tlDLcf.iTime(1, ilLoop) <> tlCLcf.iTime(1, ilLoop)) Then
                        ilMatch = False
                        Exit For
                    End If
                Next ilLoop
                If ilMatch Then
                    Do
                        ilRet = btrGetDirect(hlLcf, tlCLcf, ilLcfRecLen, llCRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        'tmSRec = tlCLcf
                        'ilRet = gGetByKeyForUpdate("Lcf", hlLcf, tmSRec)
                        'tlCLcf = tmSRec
                        'If ilRet <> BTRV_ERR_NONE Then
                        '    igBtrError = gConvertErrorCode(ilRet)
                        '    sgErrLoc = "gPrgToDelete-Get by Key Lcf(3)"
                        '    btrDestroy hlLtf
                        '    btrDestroy hlLvf
                        '    btrDestroy hlDLcf
                        '    btrDestroy hlLcf
                        '    gPrgToDelete = False
                        '    Exit Function
                        'End If
                        ilRet = btrDelete(hlLcf)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    If ilRet <> BTRV_ERR_NONE Then
                        igBtrError = gConvertErrorCode(ilRet)
                        sgErrLoc = "gPrgToDelete-Delete Lcf(4)"
                        btrDestroy hlLtf
                        btrDestroy hlLvf
                        btrDestroy hlDLcf
                        btrDestroy hlLcf
                        gPrgToDelete = False
                        Exit Function
                    End If
                    Do
                        ilRet = btrGetDirect(hlDLcf, tlDLcf, ilLcfRecLen, llDRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                        'tmSRec = tlDLcf
                        'ilRet = gGetByKeyForUpdate("Lcf", hlDLcf, tmSRec)
                        'tlDLcf = tmSRec
                        'If ilRet <> BTRV_ERR_NONE Then
                        '    igBtrError = gConvertErrorCode(ilRet)
                        '    sgErrLoc = "gPrgToDelete-Get by Key Lcf(5)"
                        '    btrDestroy hlLtf
                        '    btrDestroy hlLvf
                        '    btrDestroy hlDLcf
                        '    btrDestroy hlLcf
                        '    gPrgToDelete = False
                        '    Exit Function
                        'End If
                        ilRet = btrDelete(hlDLcf)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    If ilRet <> BTRV_ERR_NONE Then
                        igBtrError = gConvertErrorCode(ilRet)
                        sgErrLoc = "gPrgToDelete-Delete Lcf(6)"
                        btrDestroy hlLtf
                        btrDestroy hlLvf
                        btrDestroy hlDLcf
                        btrDestroy hlLcf
                        gPrgToDelete = False
                        Exit Function
                    End If
                End If
            End If
        Else
            'Test if deleted record create for day that does not exist in pending or current
            '(delete thru TFN but date original lib deleted only ran one week)
            tlLcfSrchKey.iType = ilType
            tlLcfSrchKey.sStatus = "C"
            tlLcfSrchKey.iVefCode = tlLtf.iVefCode
            tlLcfSrchKey.iLogDate(0) = tlDLcf.iLogDate(0)  'Year 1/1/1900
            tlLcfSrchKey.iLogDate(1) = tlDLcf.iLogDate(1)
            tlLcfSrchKey.iSeqNo = 1
            ilRet = btrGetEqual(hlLcf, tlCLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If ilRet <> BTRV_ERR_NONE Then
                Do
                    ilRet = btrGetDirect(hlDLcf, tlDLcf, ilLcfRecLen, llDRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    'tmSRec = tlDLcf
                    'ilRet = gGetByKeyForUpdate("Lcf", hlDLcf, tmSRec)
                    'tlDLcf = tmSRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    igBtrError = gConvertErrorCode(ilRet)
                    '    sgErrLoc = "gPrgToDelete-Get by Key Lcf(7)"
                    '    btrDestroy hlLtf
                    '    btrDestroy hlLvf
                    '    btrDestroy hlDLcf
                    '    btrDestroy hlLcf
                    '    gPrgToDelete = False
                    '    Exit Function
                    'End If
                    ilRet = btrDelete(hlDLcf)
                Loop While ilRet = BTRV_ERR_CONFLICT
            End If
        End If
        ilRet = btrGetNext(hlDLcf, tlDLcf, ilLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    ilRet = btrClose(hlLtf)
    btrDestroy hlLtf
    ilRet = btrClose(hlLvf)
    btrDestroy hlLvf
    ilRet = btrClose(hlDLcf)
    btrDestroy hlDLcf
    ilRet = btrClose(hlLcf)
    btrDestroy hlLcf
    gPrgToDelete = True
    Exit Function
gRemovePrgErr:
    igBtrError = gConvertErrorCode(ilRet)
    sgErrLoc = "gPrgToDelete(8)"
    ilRet = btrClose(hlLtf)
    btrDestroy hlLtf
    ilRet = btrClose(hlLvf)
    btrDestroy hlLvf
    ilRet = btrClose(hlDLcf)
    btrDestroy hlDLcf
    ilRet = btrClose(hlLcf)
    btrDestroy hlLcf
    gDbg_HandleError "Schedule: gPrgToDelete"
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gPrgToPend                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Insert Log library into Pending *
'*                     calendar date                   *
'*                     tgPrg contains the times/dates  *
'*                     where lvf is to be moved        *
'*                                                     *
'*******************************************************
Function gPrgToPend(frm As Form, tlLvf As LVF, ilType As Integer) As Integer
'
'    ilRet = gPrgToPend(MainForm, tlLvf, slType)
'    Where:
'       MainForm (I)- Name of Form to unload if error exists
'       tlLvf (I) - Library to be inserted into pending calendar
'       ilType (I)- Type of library (0=Regular Programming; 1->NN = Sports Programming (Game Number))
'
'       tgPrg(I)- contains dates/times to schedule libraries
'
    Dim hlLcf As Integer            'Log calendar library file handle
    Dim hlLvf As Integer            'Log library file handle
    Dim hlLtf As Integer            'Log library file handle
    Dim tlCLcf As LCF               'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim tlLtf As LTF               'LTF record image
    Dim tlLtfSrchKey As INTKEY0     'LTF key record image
    Dim ilLtfRecLen As Integer         'LTF record length
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim slDate As String
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim llLatestDate As Long    'Latest date of libraries to be inserted
    Dim llPLatestDate As Long   'Pending latest date to be checked to see if pending must be extended
    Dim llStartDate As Long
    Dim llEndDate As Long
    ReDim ilStartTime(0 To 1) As Integer
    hlLcf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hlLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPrgToPendErr
    gBtrvErrorMsg ilRet, "gPrgToPend (btrOpen: Lcf.btr)", frm
    On Error GoTo 0
    hlLvf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlLvf, "", sgDBPath & "Lvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPrgToPendErr
    gBtrvErrorMsg ilRet, "gPrgToPend (btrOpen: Lvf.btr)", frm
    On Error GoTo 0
    hlLtf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hlLtf, "", sgDBPath & "Ltf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gPrgToPendErr
    gBtrvErrorMsg ilRet, "gPrgToPend (btrOpen: Ltf.btr)", frm
    On Error GoTo 0
    ilLcfRecLen = Len(tlCLcf)
    ilLtfRecLen = Len(tlLtf)
    llLatestDate = -1
    llPLatestDate = -1
'    If igViewType = 1 Then
'        tlLcfSrchKey.sType = "A"
'        slType = "A"
'    Else
'        tlLcfSrchKey.sType = "O"
'        slType = "O"
'    End If
    tlLtfSrchKey.iCode = tlLvf.iLtfCode
    ilRet = btrGetEqual(hlLtf, tlLtf, ilLtfRecLen, tlLtfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
    On Error GoTo gPrgToPendErr
    gBtrvErrorMsg ilRet, "gPrgToPend (btrGetEqual: Ltf.btr)", frm
    On Error GoTo 0
    tlLcfSrchKey.iType = ilType
    tlLcfSrchKey.sStatus = "C"
    tlLcfSrchKey.iVefCode = tlLtf.iVefCode
    tlLcfSrchKey.iLogDate(0) = 257  'Year 1/1/2100
    tlLcfSrchKey.iLogDate(1) = 2100
    tlLcfSrchKey.iSeqNo = 1
    ilRet = btrGetLessOrEqual(hlLcf, tlCLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    If (ilRet = BTRV_ERR_NONE) And (tlCLcf.sStatus = "C") And (tlCLcf.iVefCode = tlLtf.iVefCode) And (tlCLcf.iType = ilType) Then
        gUnpackDate tlCLcf.iLogDate(0), tlCLcf.iLogDate(1), slDate
        llLatestDate = gDateValue(slDate)
    End If
    tlLcfSrchKey.sStatus = "P"
    tlLcfSrchKey.iVefCode = tlLtf.iVefCode
    tlLcfSrchKey.iLogDate(0) = 257  'Year 1/1/2100
    tlLcfSrchKey.iLogDate(1) = 2100
    tlLcfSrchKey.iSeqNo = 1
    ilRet = btrGetLessOrEqual(hlLcf, tlCLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    If (ilRet = BTRV_ERR_NONE) And (tlCLcf.sStatus = "P") And (tlCLcf.iVefCode = tlLtf.iVefCode) And (tlCLcf.iType = ilType) Then
        If tlCLcf.iLogDate(1) > 1960 Then
            gUnpackDate tlCLcf.iLogDate(0), tlCLcf.iLogDate(1), slDate
            llDate = gDateValue(slDate)
            llPLatestDate = llDate
            If llDate > llLatestDate Then
                llLatestDate = llDate
            End If
        End If
    End If
    For ilLoop = LBound(tgPrg) To UBound(tgPrg) - 1 Step 1
        If (tgPrg(ilLoop).sStartTime <> "") And (tgPrg(ilLoop).sStartDate <> "") Then
            If tgPrg(ilLoop).sEndDate <> "" Then
                llStartDate = gDateValue(tgPrg(ilLoop).sStartDate)
                'If slType = "A" Then
                '    llEndDate = llStartDate
                'Else
                    llEndDate = gDateValue(tgPrg(ilLoop).sEndDate)
                'End If
                If (llStartDate > llPLatestDate) And (llPLatestDate <> -1) Then
                    For llDate = llPLatestDate + 1 To llStartDate - 1 Step 1
                        slDate = gFormatDate(Format$(llDate, "m/d/yy"))
                        gPackDate slDate, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1)
                        ilRet = gExtendTFN(hlLcf, 0, 0, 0, "P", tlLtf.iVefCode, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1), True)
                        If Not ilRet Then
                            btrDestroy hlLtf
                            btrDestroy hlLvf
                            btrDestroy hlLcf
                            gPrgToPend = False
                            Exit Function
                        End If
                    Next llDate
                End If
                For llDate = llStartDate To llEndDate Step 1
                    If tgPrg(ilLoop).iDay(gWeekDayLong(llDate)) = 1 Then
                        slDate = gFormatDate(Format$(llDate, "m/d/yy"))
                        gPackDate slDate, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1)
                        gPackTime tgPrg(ilLoop).sStartTime, ilStartTime(0), ilStartTime(1)
                        ilRet = mMovePrgToPending(hlLcf, hlLvf, tlLvf, ilType, tlLtf.iVefCode, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1), ilStartTime(0), ilStartTime(1))
                        If ilRet <> BTRV_ERR_NONE Then
                            igBtrError = gConvertErrorCode(ilRet)
                            sgErrLoc = "gPrgToPend- mMovePrgToPending(1)"
                            btrDestroy hlLtf
                            btrDestroy hlLvf
                            btrDestroy hlLcf
                            gPrgToPend = False
                            Exit Function
                        End If
                    End If
                Next llDate
            Else    'Move program into TFN week
                For ilIndex = 1 To 7 Step 1
                    If tgPrg(ilLoop).iDay(ilIndex - 1) = 1 Then
                        gPackTime tgPrg(ilLoop).sStartTime, ilStartTime(0), ilStartTime(1)
                        ilRet = mMovePrgToPending(hlLcf, hlLvf, tlLvf, ilType, tlLtf.iVefCode, ilIndex, 0, ilStartTime(0), ilStartTime(1))
                        If ilRet <> BTRV_ERR_NONE Then
                            igBtrError = gConvertErrorCode(ilRet)
                            sgErrLoc = "gPrgToPend- mMovePrgToPending(2)"
                            btrDestroy hlLtf
                            btrDestroy hlLvf
                            btrDestroy hlLcf
                            gPrgToPend = False
                            Exit Function
                        End If
                    End If
                Next ilIndex
                llStartDate = gDateValue(tgPrg(ilLoop).sStartDate)
                If llLatestDate <> -1 Then
                    llEndDate = llLatestDate
                    If llEndDate < llStartDate + 6 Then
                        llEndDate = llStartDate + 6
                    End If
                Else
                    llEndDate = llStartDate + 6
                End If
                If (llStartDate > llPLatestDate) And (llPLatestDate <> -1) Then
                    For llDate = llPLatestDate + 1 To llStartDate - 1 Step 1
                        slDate = gFormatDate(Format$(llDate, "m/d/yy"))
                        gPackDate slDate, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1)
                        ilRet = gExtendTFN(hlLcf, 0, 0, 0, "P", tlLtf.iVefCode, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1), True)
                        If Not ilRet Then
                            btrDestroy hlLtf
                            btrDestroy hlLvf
                            btrDestroy hlLcf
                            gPrgToPend = False
                            Exit Function
                        End If
                    Next llDate
                End If
                For llDate = llStartDate To llEndDate Step 1
                    If tgPrg(ilLoop).iDay(gWeekDayLong(llDate)) = 1 Then
                        slDate = gFormatDate(Format$(llDate, "m/d/yy"))
                        gPackDate slDate, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1)
                        gPackTime tgPrg(ilLoop).sStartTime, ilStartTime(0), ilStartTime(1)
                        ilRet = mMovePrgToPending(hlLcf, hlLvf, tlLvf, ilType, tlLtf.iVefCode, tlLcfSrchKey.iLogDate(0), tlLcfSrchKey.iLogDate(1), ilStartTime(0), ilStartTime(1))
                        If ilRet <> BTRV_ERR_NONE Then
                            igBtrError = gConvertErrorCode(ilRet)
                            sgErrLoc = "gPrgToPend- mMovePrgToPending(3)"
                            btrDestroy hlLtf
                            btrDestroy hlLvf
                            btrDestroy hlLcf
                            gPrgToPend = False
                            Exit Function
                        End If
                    End If
                Next llDate
            End If
        End If
    Next ilLoop
    ilRet = btrClose(hlLtf)
    btrDestroy hlLtf
    ilRet = btrClose(hlLvf)
    btrDestroy hlLvf
    ilRet = btrClose(hlLcf)
    btrDestroy hlLcf
    gPrgToPend = True
    Exit Function
gPrgToPendErr:
    igBtrError = gConvertErrorCode(ilRet)
    sgErrLoc = "gPrgToPend(4)"
    ilRet = btrClose(hlLtf)
    btrDestroy hlLtf
    ilRet = btrClose(hlLvf)
    btrDestroy hlLvf
    ilRet = btrClose(hlLcf)
    btrDestroy hlLcf
    gDbg_HandleError "Schedule: gPrgToPend"
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gRemoveSmf                      *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Remove MG Spec for a spot      *
'*                                                     *
'*******************************************************
Function gRemoveSmf(hlSmf As Integer, tlSmf As SMF, tlSdf As SDF, hlSxf As Integer) As Integer
'
'   ilRet = gRemoveSmf(hlSmf, tlSmf, tlSdf)
'   Where:
'       hlSmf(I)- Smf open handle
'       tlSmf(O) - Spot MG record if found and deleted (this is returned so specification can be retained if
'                  MG moved within speicifed boundary
'       tlSdf(I/O)- Sdf record, if smf exist, then dates within Sdf reset
'
'       ilRet = True or False
'
    Dim tlWSmf As SMF
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim llSmfRecPos As Long
    tgSxfSdf = tlSdf
    tlSmf.lChfCode = 0
    If (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O") Then
        ilRet = gSxfDelete(hlSxf, tgSxfSdf)
'        tmSmfSrchKey.lChfCode = tlSdf.lChfCode
'        tmSmfSrchKey.iLineNo = tlSdf.iLineNo
'        tmSmfSrchKey.iMissedDate(0) = 0 'sch date =tlSdf.iDate(0)
'        tmSmfSrchKey.iMissedDate(1) = 0 'sch date =tlSdf.iDate(1)
'        imSmfRecLen = Len(tlWSmf)
'        ilRet = btrGetGreaterOrEqual(hlSmf, tlWSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get current record
'        Do While (ilRet = BTRV_ERR_NONE) And (tlWSmf.lChfCode = tlSdf.lChfCode) And (tlWSmf.iLineNo = tlSdf.iLineNo)
'            'If (tlWSmf.sSchStatus = tlSdf.sSchStatus) And (tlWSmf.iActualDate(0) = tlSdf.iDate(0)) And (tlWSmf.iActualDate(1) = tlSdf.iDate(1)) And (tlWSmf.iActualTime(0) = tlSdf.iTime(0)) And (tlWSmf.iActualTime(1) = tlSdf.iTime(1)) Then
        imSmfRecLen = Len(tlWSmf)
        tmSmfSrchKey2.lCode = tlSdf.lCode
        ilRet = btrGetEqual(hlSmf, tlWSmf, imSmfRecLen, tmSmfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)
        If ilRet = BTRV_ERR_NONE Then

'            If tlWSmf.lSdfCode = tlSdf.lCode Then
                'Remove
                tlSmf = tlWSmf 'return SMF that is deleted
                'Reset schedule dates to missed date
                If (tlSmf.lMtfCode = 0) Or (lgMtfNoRecs = 0) Then
                    tlSdf.iDate(0) = tlSmf.iMissedDate(0)
                    tlSdf.iDate(1) = tlSmf.iMissedDate(1)
                    tlSdf.iTime(0) = tlSmf.iMissedTime(0)
                    tlSdf.iTime(1) = tlSmf.iMissedTime(1)
                    If tlSmf.iOrigSchVef <> 0 Then
                        tlSdf.iVefCode = tlSmf.iOrigSchVef
                    End If
                    tlSdf.iGameNo = tlSmf.iGameNo
                End If
                'In gChgSchSpot, Smf.lMtfCode is checked and is not zero, then sdf.lSmfCode set to smf.lMtfCode
                'and sdf.sTracer set to *
                tlSdf.lSmfCode = 0
                ilRet = btrGetPosition(hlSmf, llSmfRecPos)
                Do
                    ilCRet = btrGetDirect(hlSmf, tlWSmf, imSmfRecLen, llSmfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    'tmSRec = tlWSmf
                    'ilCRet = gGetByKeyForUpdate("SMF", hlSmf, tmSRec)
                    'tlWSmf = tmSRec
                    'If ilCRet <> BTRV_ERR_NONE Then
                    '    igBtrError = ilCRet
                    '    sgErrLoc = "gRemoveSmf-GetByKey Smf(1)"
                    '    gRemoveSmf = False
                    '    Exit Function
                    'End If
                    ilRet = btrDelete(hlSmf)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    igBtrError = gConvertErrorCode(ilRet)
                    sgErrLoc = "gRemoveSmf-Delete Smf(2)"
                    gRemoveSmf = False
                    Exit Function
                End If
'                Exit Do
'            End If
'            ilRet = btrGetNext(hlSmf, tlWSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
        Else
            igBtrError = gConvertErrorCode(ilRet)
            sgErrLoc = "gRemoveSmf-Delete Smf(3)"
            gRemoveSmf = False
            Exit Function
        End If
    Else
        tlSmf.lMtfCode = 0  'This is required for gChgSchSpot
    End If
    gRemoveSmf = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gUnschSpots                     *
'*                                                     *
'*             Created:7/19/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Unschedule spots               *
'*                                                     *
'*******************************************************
Function gUnschSpots(ilVehCode As Integer, llChfCode As Long, llLastLogDate As Long, slStartDate As String, slEndDate As String, slStartTime As String, slEndTime As String, ilInGameNo As Integer) As Integer
'
'   ilRet = gUnschSpots(ilVehCode, llChfCode, llLastLogDate, slStartDate, slEndDate, slStartTime, slEndTime)
'
'   Where:
'       ilVehCode (I)-Vehicle code number
'       llChfCode(I)-Contract Code Number or -1 for all
'       slStartDate (I)- Start Date that that spots are to be removed
'       slEndDate (I)- End Date that that spots are to be removed
'       slStartTime (I)- Start Time (included)
'       slEndTime (I)- End time (not included)
'
'       ilRet = True (Ok) or (False (i/o error)
'
    Dim ilType As Integer
    Dim ilRet As Integer
    Dim ilCRet As Integer
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
    Dim llSpotTime As Long
    Dim ilSpot As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilTime0 As Integer
    Dim ilTime1 As Integer
    Dim ilOrigSchVef As Integer
    Dim ilLoop As Integer
    Dim llRecPos As Long
    'Dim llSdfRecPos As Long
    Dim ilLineNo As Integer
    Dim ilChfFound As Integer
    'Spot summary
    Dim hlSsf As Integer        'Spot summary file handle
    'Spot detail record information
    Dim hlSdf As Integer        'Spot detail file handle
    Dim hlSmf As Integer
    Dim hlStf As Integer
    Dim hlGsf As Integer
    Dim hlGhf As Integer
    Dim hlSxf As Integer
    Dim slOrigSchStatus As String
    Dim slXSpotType As String
    Dim ilEvt As Integer
    Dim ilGameVeh As Integer
    Dim ilOrigGameNo As Integer
    Dim ilVef As Integer
    Dim ilGame As Integer
    Dim ilGameNo As Integer
    ReDim tlBBSdf(0 To 0) As SDF

    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    llStartTime = CLng(gTimeToCurrency(slStartTime, False))
    llEndTime = CLng(gTimeToCurrency(slEndTime, True)) - 1
    hlSsf = CBtrvTable(TWOHANDLES)        'Create SSF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gUnschSpots-Open Ssf(1)"
        ilRet = btrClose(hlSsf)
        btrDestroy hlSsf
        gUnschSpots = False
        Exit Function
    End If
    imSdfRecLen = Len(tmSdf)  'Get and save SDF record length
    hlSdf = CBtrvTable(TWOHANDLES)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gUnschSpots-Open Sdf(2)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        gUnschSpots = False
        Exit Function
    End If
    imSmfRecLen = Len(tmSmf)  'Get and save SMF record length
    hlSmf = CBtrvTable(TWOHANDLES)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gUnschSpots-Open Smf(3)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlSmf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlSmf
        gUnschSpots = False
        Exit Function
    End If
    imStfRecLen = Len(tmStf)  'Get and save SMF record length
    hlStf = CBtrvTable(TWOHANDLES)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlStf, "", sgDBPath & "Stf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gUnschSpots-Open Stf(4)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlStf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlSmf
        btrDestroy hlStf
        gUnschSpots = False
        Exit Function
    End If
    hlGsf = CBtrvTable(TWOHANDLES)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gUnschSpots-Open Gsf(4)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlStf)
        ilRet = btrClose(hlGsf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlSmf
        btrDestroy hlStf
        btrDestroy hlGsf
        gUnschSpots = False
        Exit Function
    End If
    hlGhf = CBtrvTable(TWOHANDLES)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gUnschSpots-Open Gsf(4)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlStf)
        ilRet = btrClose(hlGsf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlSmf
        btrDestroy hlStf
        btrDestroy hlGsf
        btrDestroy hlGhf
        gUnschSpots = False
        Exit Function
    End If
    
    hlSxf = CBtrvTable(TWOHANDLES)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSxf, "", sgDBPath & "Sxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gUnschSpots-Open Gsf(4)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlStf)
        ilRet = btrClose(hlGsf)
        ilRet = btrClose(hlGhf)
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlSmf
        btrDestroy hlStf
        btrDestroy hlGsf
        btrDestroy hlGhf
        gUnschSpots = False
        Exit Function
    End If
    
    ilGameVeh = False
    ilVef = gBinarySearchVef(ilVehCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).sType = "G" Then
            ilGameVeh = True
        End If
    End If
    If llChfCode < 0 Then
        For llDate = llStartDate To llEndDate Step 1
            ilType = 0
            slDate = Format$(llDate, "m/d/yy")
            ilDay = gWeekDayStr(slDate)
            gPackDate slDate, ilLogDate0, ilLogDate1
            imSsfRecLen = Len(tmSsf)
            ReDim ilGameNos(0 To 0) As Integer
            If Not ilGameVeh Then
                tmSsfSrchKey.iType = ilType
                tmSsfSrchKey.iVefCode = ilVehCode
                tmSsfSrchKey.iDate(0) = ilLogDate0
                tmSsfSrchKey.iDate(1) = ilLogDate1
                tmSsfSrchKey.iStartTime(0) = 0
                tmSsfSrchKey.iStartTime(1) = 0
                ilRet = gSSFGetGreaterOrEqual(hlSsf, tgSsf(0), imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            Else
                tmSsfSrchKey2.iVefCode = ilVehCode
                tmSsfSrchKey2.iDate(0) = ilLogDate0
                tmSsfSrchKey2.iDate(1) = ilLogDate1
                ilRet = gSSFGetEqualKey2(hlSsf, tgSsf(0), imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            
            End If
            Do While (ilRet = BTRV_ERR_NONE) And (((tgSsf(0).iType = ilType) And (Not ilGameVeh)) Or ((tgSsf(0).iType > 0) And (ilGameVeh))) And (tgSsf(0).iVefCode = ilVehCode) And (tgSsf(0).iDate(0) = ilLogDate0) And (tgSsf(0).iDate(1) = ilLogDate1)
                ilRPRet = gSSFGetPosition(hlSsf, llSsfRecPos)
                If (Not ilGameVeh) Or ((ilGameVeh) And ((ilInGameNo = -1) Or (tgSsf(0).iType = ilInGameNo))) Then
                    tmSsf = tgSsf(0)   'Move header
                    tmSsf.iCount = 0    'Remove all records
                    'Add all records except spots
                    If ilGameVeh Then
                        ilGameNos(UBound(ilGameNos)) = tgSsf(0).iType
                        ReDim Preserve ilGameNos(0 To UBound(ilGameNos) + 1) As Integer
                    End If
                    ilLoop = 1
                    Do While ilLoop <= tgSsf(0).iCount
                       LSet tmAvail = tgSsf(0).tPas(ADJSSFPASBZ + ilLoop)
                        If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then 'Contract Avail subrecord
                            'Test time-
                            gUnpackTime tmAvail.iTime(0), tmAvail.iTime(1), "A", "1", slTime
                            llAvailTime = CLng(gTimeToCurrency(slTime, False))
                            If (llAvailTime >= llStartTime) And (llAvailTime <= llEndTime) Then
                                'tmAvail.iAvInfo = tmAvail.iAvInfo And &HF3F 'Remove close and freeze
                                For ilSpot = ilLoop + 1 To ilLoop + tmAvail.iNoSpotsThis Step 1
                                   LSet tmSpot = tgSsf(0).tPas(ADJSSFPASBZ + ilSpot)
                                    tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                    ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                    If ilRet = BTRV_ERR_NONE Then
                                        'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                                        slOrigSchStatus = tmSdf.sSchStatus
                                        ilRet = gMakeTracer(hlSdf, tmSdf, 0, hlStf, llLastLogDate, "M", "U", tmSdf.iRotNo, hlGsf)
                                        If Not ilRet Then
                                            btrDestroy hlSsf
                                            btrDestroy hlSdf
                                            btrDestroy hlSmf
                                            btrDestroy hlStf
                                            btrDestroy hlGsf
                                            btrDestroy hlGhf
                                            btrDestroy hlSxf
                                            gUnschSpots = False
                                            Exit Function
                                        End If
                                        ilRet = gRemoveSmf(hlSmf, tmSmf, tmSdf, hlSxf)  'resets missed date
                                        If Not ilRet Then
                                            btrDestroy hlSsf
                                            btrDestroy hlSdf
                                            btrDestroy hlSmf
                                            btrDestroy hlStf
                                            btrDestroy hlGsf
                                            btrDestroy hlGhf
                                            btrDestroy hlSxf
                                            gUnschSpots = False
                                            Exit Function
                                        End If
                                        ilDate0 = tmSdf.iDate(0)
                                        ilDate1 = tmSdf.iDate(1)
                                        ilTime0 = tmSdf.iTime(0)
                                        ilTime1 = tmSdf.iTime(1)
                                        ilOrigSchVef = tmSdf.iVefCode
                                        ilOrigGameNo = tmSdf.iGameNo
                                        If tmSdf.sSpotType = "X" Then
                                            slXSpotType = "X"
                                            If ((tmSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                                                slXSpotType = ""
                                            End If
                                        Else
                                            slXSpotType = ""
                                        End If
                                        'If (tmSdf.sSpotType = "T") Or (tmSdf.sSpotType = "Q") Or (tmSdf.sSpotType = "S") Or (tmSdf.sSpotType = "M") Or (slXSpotType = "X") Then
                                        If ((tmSdf.sSpotType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmSdf.sSpotType = "Q") Or ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (slXSpotType = "X") Then
                                            'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                                            Do
                                                ilRet = btrDelete(hlSdf)
                                                If ilRet = BTRV_ERR_CONFLICT Then
                                                    'ilCRet = btrGetDirect(hlSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                    tmSdfSrchKey3.lCode = tmSdf.lCode
                                                    ilCRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                    'tmSRec = tmSdf
                                                    'ilCRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                                                    'tmSdf = tmSRec
                                                    'If ilCRet <> BTRV_ERR_NONE Then
                                                    '    igBtrError = ilCRet
                                                    '    sgErrLoc = "gUnschSpots-GetByKey Sdf(5)"
                                                    '    btrDestroy hlSsf
                                                    '    btrDestroy hlSdf
                                                    '    btrDestroy hlSmf
                                                    '    btrDestroy hlStf
                                                    '    gUnschSpots = False
                                                    '    Exit Function
                                                    'End If
                                                End If
                                            Loop While ilRet = BTRV_ERR_CONFLICT
                                            If ilRet <> BTRV_ERR_NONE Then
                                                igBtrError = gConvertErrorCode(ilRet)
                                                sgErrLoc = "gUnschSpots-Delete Sdf(6)"
                                                btrDestroy hlSsf
                                                btrDestroy hlSdf
                                                btrDestroy hlSmf
                                                btrDestroy hlStf
                                                btrDestroy hlGsf
                                                btrDestroy hlGhf
                                                btrDestroy hlSxf
                                                gUnschSpots = False
                                                Exit Function
                                            End If
                                        Else
                                            Do
                                                'ilRet = btrGetDirect(hlSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                tmSdfSrchKey3.lCode = tmSdf.lCode
                                                ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                'tmSRec = tmSdf
                                                'ilRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                                                'tmSdf = tmSRec
                                                'If ilRet <> BTRV_ERR_NONE Then
                                                '    igBtrError = gConvertErrorCode(ilRet)
                                                '    sgErrLoc = "gUnschSpots-GetByKey Sdf(7)"
                                                '    btrDestroy hlSsf
                                                '    btrDestroy hlSdf
                                                '    btrDestroy hlSmf
                                                '    btrDestroy hlStf
                                                '    gUnschSpots = False
                                                '    Exit Function
                                                'End If
                                                tmSdf.sSchStatus = "M"
                                                tmSdf.iMnfMissed = igMnfMissed
                                                'tmSdf.sPtType = "0"
                                                'tmSdf.iRotNo = 0
                                                tmSdf.iDate(0) = ilDate0
                                                tmSdf.iDate(1) = ilDate1
                                                tmSdf.iTime(0) = ilTime0
                                                tmSdf.iTime(1) = ilTime1
                                                tmSdf.iVefCode = ilOrigSchVef
                                                tmSdf.iGameNo = ilOrigGameNo
                                                tmSdf.lSmfCode = 0
                                                If ((tmSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                                                    tmSdf.sTracer = "*"
                                                    tmSdf.lSmfCode = tmSmf.lMtfCode
                                                End If
                                                tmSdf.sXCrossMidnight = "N"
                                                tmSdf.sWasMG = "N"
                                                tmSdf.sFromWorkArea = "N"
                                                tmSdf.iUrfCode = tgUrf(0).iCode
                                                ilRet = btrUpdate(hlSdf, tmSdf, imSdfRecLen)
                                            Loop While ilRet = BTRV_ERR_CONFLICT
                                            If ilRet <> BTRV_ERR_NONE Then
                                                igBtrError = gConvertErrorCode(ilRet)
                                                sgErrLoc = "gUnschSpots-Update Sdf(8)"
                                                btrDestroy hlSsf
                                                btrDestroy hlSdf
                                                btrDestroy hlSmf
                                                btrDestroy hlStf
                                                btrDestroy hlGsf
                                                btrDestroy hlGhf
                                                btrDestroy hlSxf
                                                gUnschSpots = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                Next ilSpot
                                ilLoop = ilLoop + tmAvail.iNoSpotsThis  'bypass spots
                                tmAvail.iNoSpotsThis = 0    'remove spot count
                            End If
                        End If
                        tmSsf.iCount = tmSsf.iCount + 1
                        tmSsf.tPas(ADJSSFPASBZ + tmSsf.iCount) = tmAvail
                        ilLoop = ilLoop + 1 'Increment to next event
                    Loop
                    imSsfRecLen = igSSFBaseLen + tmSsf.iCount * Len(tmAvail)
                    ilRet = gSSFUpdate(hlSsf, tmSsf, imSsfRecLen)
                    If ilRet <> BTRV_ERR_NONE Then
                        igBtrError = gConvertErrorCode(ilRet)
                        sgErrLoc = "gUnschSpots-Update Ssf(9)"
                        btrDestroy hlSsf
                        btrDestroy hlSdf
                        btrDestroy hlSmf
                        btrDestroy hlStf
                        btrDestroy hlGsf
                        btrDestroy hlGhf
                        btrDestroy hlSxf
                        gUnschSpots = False
                        Exit Function
                    End If
                End If
                Do
                    imSsfRecLen = Len(tmSsf)
                    If Not ilGameVeh Then
                        ilRet = gSSFGetNext(hlSsf, tgSsf(0), imSsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    Else
                        ilRet = gSSFGetNext(hlSsf, tgSsf(0), imSsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    End If
                    ilRPRet = gSSFGetPosition(hlSsf, llRecPos)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                Loop While llRecPos = llSsfRecPos
            Loop
            If tgSpf.sUsingBBs = "Y" Then
                For ilGame = 0 To UBound(ilGameNos) - 1 Step 1
                    ilGameNo = ilGameNos(ilGame)
                    ReDim tlBBSdf(0 To 0) As SDF
                    ilRet = gGetBBSpots(hlSdf, ilVehCode, ilGameNo, slDate, tlBBSdf())
                    For ilEvt = 0 To UBound(tlBBSdf) - 1 Step 1
                        Do
                            tmSdfSrchKey3.lCode = tlBBSdf(ilEvt).lCode
                            ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                            If ilRet = BTRV_ERR_NONE Then
                                ilRet = btrDelete(hlSdf)
                            End If
                        Loop While ilRet = BTRV_ERR_CONFLICT
                    Next ilEvt
                Next ilGame
            End If
        Next llDate
    Else
        'Remove scheduled spots
        ilLineNo = 0
        Do
            ilChfFound = False
            tmSdfSrchKey0.iVefCode = ilVehCode
            tmSdfSrchKey0.lChfCode = llChfCode
            tmSdfSrchKey0.iLineNo = ilLineNo
            tmSdfSrchKey0.lFsfCode = 0
            slDate = Format$(llStartDate, "m/d/yy")
            gPackDate slDate, ilDate0, ilDate1
            tmSdfSrchKey0.iDate(0) = ilDate0
            tmSdfSrchKey0.iDate(1) = ilDate1
            tmSdfSrchKey0.sSchStatus = ""
            tmSdfSrchKey0.iTime(0) = 0
            tmSdfSrchKey0.iTime(1) = 0
            ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVehCode) And (tmSdf.lChfCode = llChfCode)
                ilChfFound = True
                ilLineNo = tmSdf.iLineNo
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                If (llDate > llEndDate) Then
                    Exit Do
                End If
                If ((tmSdf.sSchStatus = "S") Or (tmSdf.sSchStatus = "O") Or (tmSdf.sSchStatus = "G")) And (llDate >= llStartDate) Then
                    If (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                        gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slTime
                        llSpotTime = CLng(gTimeToCurrency(slTime, False))
                        If (llSpotTime >= llStartTime) And (llSpotTime <= llEndTime) Then
                            'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                            If (tmSdf.sSpotType <> "O") And (tmSdf.sSpotType <> "C") Then
                                ilGameNo = tmSdf.iGameNo
                                ilRet = gMakeTracer(hlSdf, tmSdf, 0, hlStf, llLastLogDate, "M", "U", tmSdf.iRotNo, hlGsf)
                                If Not ilRet Then
                                    btrDestroy hlSsf
                                    btrDestroy hlSdf
                                    btrDestroy hlSmf
                                    btrDestroy hlStf
                                    btrDestroy hlGsf
                                    btrDestroy hlGhf
                                    btrDestroy hlSxf
                                    gUnschSpots = False
                                    Exit Function
                                End If
                                If Not gChgSchSpot("M", hlSdf, tmSdf, hlSmf, ilGameNo, tmSmf, hlSsf, tgSsf(0), lgSsfDate(0), lgSsfRecPos(0), hlSxf, hlGsf, hlGhf) Then
                                    btrDestroy hlSsf
                                    btrDestroy hlSdf
                                    btrDestroy hlSmf
                                    btrDestroy hlStf
                                    btrDestroy hlGsf
                                    btrDestroy hlGhf
                                    btrDestroy hlSxf
                                    gUnschSpots = False
                                    Exit Function
                                End If
                            Else
                                Do
                                    'ilRet = btrGetDirect(hlSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                    tmSdfSrchKey3.lCode = tmSdf.lCode
                                    ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                    ilRet = btrDelete(hlSdf)
                                Loop While ilRet = BTRV_ERR_CONFLICT
                            End If
                            tmSdfSrchKey0.iVefCode = ilVehCode
                            tmSdfSrchKey0.lChfCode = llChfCode
                            tmSdfSrchKey0.iLineNo = ilLineNo
                            tmSdfSrchKey0.lFsfCode = 0
                            slDate = Format$(llDate, "m/d/yy")
                            gPackDate slDate, ilDate0, ilDate1
                            tmSdfSrchKey0.iDate(0) = ilDate0
                            tmSdfSrchKey0.iDate(1) = ilDate1
                            tmSdfSrchKey0.sSchStatus = ""
                            tmSdfSrchKey0.iTime(0) = 0
                            tmSdfSrchKey0.iTime(1) = 0
                            ilRet = btrGetGreaterOrEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Else
                            ilRet = btrGetNext(hlSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        End If
                    Else
                        ilRet = btrGetNext(hlSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    End If
                Else
                    ilRet = btrGetNext(hlSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                End If
            Loop
            If (tmSdf.lChfCode <> llChfCode) Then
                Exit Do
            End If
            ilLineNo = ilLineNo + 1
        Loop While ilChfFound
    End If
    Erase tlBBSdf
    ilRet = btrClose(hlSsf)
    btrDestroy hlSsf
    ilRet = btrClose(hlSdf)
    btrDestroy hlSdf
    ilRet = btrClose(hlSmf)
    btrDestroy hlSmf
    ilRet = btrClose(hlStf)
    btrDestroy hlStf
    ilRet = btrClose(hlGsf)
    btrDestroy hlGsf
    btrDestroy hlGhf
    btrDestroy hlSxf
    gUnschSpots = True
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gRemovePrgFromPending           *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove library from pending log *
'*                     dates                           *
'*                                                     *
'*******************************************************
Private Function mAddPrgToDelete(hlLcf As Integer, hlLvf As Integer, tlLvf As LVF, ilType As Integer, ilVefCode As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, ilStartTime0 As Integer, ilStartTime1 As Integer) As Integer
    Dim tlDLcf As LCF               'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim llLvfCode As Long
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilSeqNo As Integer
    Dim ilTime0 As Integer
    Dim ilTime1 As Integer
    Dim llPosition As Long  'Record position incase of conflict
    Dim ilFound As Integer
    ilFound = False
    ilLcfRecLen = Len(tlDLcf)
    llLvfCode = tlLvf.lCode
    ilTime0 = ilStartTime0
    ilTime1 = ilStartTime1
    ilSeqNo = 1
    tlLcfSrchKey.iType = ilType 'on Air or Alternate
    tlLcfSrchKey.sStatus = "D"
    tlLcfSrchKey.iVefCode = ilVefCode
    tlLcfSrchKey.iLogDate(0) = ilLogDate0   '1=Monday TFN; 2=Tues,...7=Sunday TFN
    tlLcfSrchKey.iLogDate(1) = ilLogDate1
    tlLcfSrchKey.iSeqNo = ilSeqNo
    ilRet = btrGetEqual(hlLcf, tlDLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
    Do
        If (ilRet <> BTRV_ERR_NONE) Or (tlDLcf.iType <> ilType) Or (tlDLcf.sStatus <> "D") Or (tlDLcf.iVefCode <> ilVefCode) Or (tlDLcf.iLogDate(0) <> ilLogDate0) Or (tlDLcf.iLogDate(1) <> ilLogDate1) Then
            'Create date image- so library can be merged into date
            tlDLcf.iVefCode = ilVefCode
            tlDLcf.iLogDate(0) = ilLogDate0
            tlDLcf.iLogDate(1) = ilLogDate1
            tlDLcf.iSeqNo = ilSeqNo
            tlDLcf.iType = ilType
            tlDLcf.sStatus = "D"
            tlDLcf.sTiming = "N" 'Timing not started
            tlDLcf.sAffPost = "N"
            tlDLcf.iLastTime(0) = 0
            tlDLcf.iLastTime(1) = 0
            For ilIndex = LBound(tlDLcf.lLvfCode) To UBound(tlDLcf.lLvfCode) Step 1
                tlDLcf.lLvfCode(ilIndex) = 0
                tlDLcf.iTime(0, ilIndex) = 0
                tlDLcf.iTime(1, ilIndex) = 0
            Next ilIndex
            tlDLcf.lLvfCode(LBound(tlDLcf.lLvfCode)) = llLvfCode
            tlDLcf.iTime(0, LBound(tlDLcf.lLvfCode)) = ilTime0
            tlDLcf.iTime(1, LBound(tlDLcf.lLvfCode)) = ilTime1
            Do  'Loop until record updated or added
                tlDLcf.lCode = 0
                ilRet = btrInsert(hlLcf, tlDLcf, ilLcfRecLen, INDEXKEY3)
            Loop While ilRet = BTRV_ERR_CONFLICT
            mAddPrgToDelete = ilRet
            ilFound = True
            Exit Function
        Else    'Add to array
            ilRet = btrGetPosition(hlLcf, llPosition)    'Get position incase of conflict
            ilLoop = LBound(tlDLcf.lLvfCode)
            Do While ilLoop <= UBound(tlDLcf.lLvfCode)
                If tlDLcf.lLvfCode(ilLoop) = 0 Then
                    tlDLcf.lLvfCode(ilLoop) = llLvfCode
                    tlDLcf.iTime(0, ilLoop) = ilTime0
                    tlDLcf.iTime(1, ilLoop) = ilTime1
                    ilRet = btrUpdate(hlLcf, tlDLcf, ilLcfRecLen)
                    If ilRet = BTRV_ERR_NONE Then
                        ilFound = True
                        Exit Function
                    End If
                    ilRet = btrGetDirect(hlLcf, tlDLcf, ilLcfRecLen, llPosition, INDEXKEY0, BTRV_LOCK_NONE)
                    'tmSRec = tlDLcf
                    'ilRet = gGetByKeyForUpdate("Lcf", hlLcf, tmSRec)
                    'tlDLcf = tmSRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    mAddPrgToDelete = ilRet
                    '    Exit Function
                    'End If
                    ilLoop = LBound(tlDLcf.lLvfCode) - 1
                End If
                ilLoop = ilLoop + 1
            Loop
            ilSeqNo = ilSeqNo + 1
            ilRet = btrGetNext(hlLcf, tlDLcf, ilLcfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
        End If
    Loop While Not ilFound
    mAddPrgToDelete = BTRV_ERR_NONE
    Exit Function
End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gMovePrgToPending               *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Move library to pending log     *
'*                     dates                           *
'*            Note: Pending Library calendar contains  *
'*                  only altered libraries             *
'*                  Deleted libraries are retained in  *
'*                  LCf with status of "D"             *
'*                                                     *
'*******************************************************
Private Function mMovePrgToPending(hlLcf As Integer, hlLvf As Integer, tlLvf As LVF, ilType As Integer, ilVefCode As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, ilStartTime0 As Integer, ilStartTime1 As Integer) As Integer
    Dim tlPLcf As LCF               'LCF record image
    Dim tlPNLcf As LCF               'LCF record image
    Dim tlLcfSrchKey As LCFKEY0     'LCF key record image
    Dim ilLcfRecLen As Integer         'LCF record length
    Dim tlLvfSrchKey As LONGKEY0     'LVF key record image
    Dim ilLvfRecLen As Integer         'LVF record length
    Dim tlPLvf As LVF
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilSeqNo As Integer
    Dim ilTFNSeqNo As Integer   'Copy TFN into Date before merging Library
    Dim slPTime As String
    Dim slLTime As String
    Dim llDnLvfCode As Long     'LvfCode moved from end to make room
    Dim ilDnTime0 As Integer    'iTime moved from end to make room
    Dim ilDnTime1 As Integer
    Dim llLvfCode As Long
    Dim ilTime0 As Integer
    Dim ilTime1 As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim llPosition As Long  'Record position incase of conflict
    Dim llNPosition As Long  'Record position incase of conflict
    Dim slStr As String
    Dim slPEndTime As String
    Dim slLEndTime As String
    Dim ilRemove As Integer
    Dim ilRemoveState As Integer    'Indicates if testing for overlaps
    Dim ilStartLoop As Integer
    Dim slXMid As String

    ilLcfRecLen = Len(tlPLcf)
    ilLvfRecLen = Len(tlPLvf)
    llLvfCode = tlLvf.lCode
    ilTime0 = ilStartTime0
    ilTime1 = ilStartTime1
    ilRemoveState = True
    ilSeqNo = 1
    Do  'Cycle thru all seq numbers for date
        'Determine if pending date exist
        'If so, merge library into date
        'If not, create day , then add library
        gUnpackTime ilTime0, ilTime1, "A", "1", slPTime
        tlLcfSrchKey.iType = ilType 'on Air or Alternate
        tlLcfSrchKey.sStatus = "P"
        tlLcfSrchKey.iVefCode = ilVefCode
        tlLcfSrchKey.iLogDate(0) = ilLogDate0   '1=Monday TFN; 2=Tues,...7=Sunday TFN
        tlLcfSrchKey.iLogDate(1) = ilLogDate1
        tlLcfSrchKey.iSeqNo = ilSeqNo
        ilRet = btrGetEqual(hlLcf, tlPLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            If ilSeqNo = 1 Then
                'Date does not exist for Date- move pending TFN into date
                tlLcfSrchKey.iType = ilType 'on Air or Alternate
                tlLcfSrchKey.sStatus = "P"
                tlLcfSrchKey.iVefCode = ilVefCode
                'If not TFN check if TFN exist, if so move to date
                If (ilLogDate0 > 7) Or (ilLogDate1 <> 0) Then
                    'Copy TFN into Date before merging Library into pending
                    ilTFNSeqNo = 2
                    Do
                        gUnpackDate ilLogDate0, ilLogDate1, slDate
                        llDate = gDateValue(slDate)
                        tlLcfSrchKey.iLogDate(0) = gWeekDayStr(slDate) + 1 '1=Monday TFN; 2=Tues,...7=Sunday TFN
                        tlLcfSrchKey.iLogDate(1) = 0
                        tlLcfSrchKey.iSeqNo = ilTFNSeqNo
                        ilRet = btrGetEqual(hlLcf, tlPLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                        If ilRet = BTRV_ERR_NONE Then
                            tlPLcf.iLogDate(0) = ilLogDate0
                            tlPLcf.iLogDate(1) = ilLogDate1
                            tlPLcf.iUrfCode = tgUrf(0).iCode
                            Do  'Loop until record updated or added
                                tlPLcf.lCode = 0
                                ilRet = btrInsert(hlLcf, tlPLcf, ilLcfRecLen, INDEXKEY3)
                            Loop While ilRet = BTRV_ERR_CONFLICT
                            If ilRet <> BTRV_ERR_NONE Then
                                mMovePrgToPending = ilRet
                                Exit Function
                            End If
                            ilTFNSeqNo = ilTFNSeqNo + 1
                        Else
                            ilTFNSeqNo = 1
                        End If
                    Loop While ilTFNSeqNo > 1
                    gUnpackDate ilLogDate0, ilLogDate1, slDate
                    llDate = gDateValue(slDate)
                    tlLcfSrchKey.iLogDate(0) = gWeekDayStr(slDate) + 1 '1=Monday TFN; 2=Tues,...7=Sunday TFN
                    tlLcfSrchKey.iLogDate(1) = 0
                    tlLcfSrchKey.iSeqNo = ilTFNSeqNo
                    ilRet = btrGetEqual(hlLcf, tlPLcf, ilLcfRecLen, tlLcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                Else
                    ilRet = -1
                End If
            Else
                ilRet = -1
            End If
            tlPLcf.iUrfCode = tgUrf(0).iCode
            If ilRet <> BTRV_ERR_NONE Then
                'Create date image- so library can be merged into date
                tlPLcf.iVefCode = ilVefCode
                tlPLcf.iLogDate(0) = ilLogDate0
                tlPLcf.iLogDate(1) = ilLogDate1
                tlPLcf.iSeqNo = ilSeqNo
                tlPLcf.iType = ilType
                tlPLcf.sStatus = "P"
                tlPLcf.sTiming = "N" 'Timing not started
                tlPLcf.sAffPost = "N"
                tlPLcf.iLastTime(0) = 0
                tlPLcf.iLastTime(1) = 0
                For ilIndex = LBound(tlPLcf.lLvfCode) To UBound(tlPLcf.lLvfCode) Step 1
                    tlPLcf.lLvfCode(ilIndex) = 0
                    tlPLcf.iTime(0, ilIndex) = 0
                    tlPLcf.iTime(1, ilIndex) = 0
                Next ilIndex
                tlPLcf.lLvfCode(LBound(tlPLcf.lLvfCode)) = llLvfCode
                tlPLcf.iTime(0, LBound(tlPLcf.lLvfCode)) = ilTime0
                tlPLcf.iTime(1, LBound(tlPLcf.lLvfCode)) = ilTime1
                Do  'Loop until record updated or added
                    tlPLcf.lCode = 0
                    ilRet = btrInsert(hlLcf, tlPLcf, ilLcfRecLen, INDEXKEY3)
                Loop While ilRet = BTRV_ERR_CONFLICT
                mMovePrgToPending = ilRet
                Exit Function
            Else
                tlPLcf.iLogDate(0) = ilLogDate0
                tlPLcf.iLogDate(1) = ilLogDate1
                Do  'Loop until record updated or added
                    tlPLcf.lCode = 0
                    ilRet = btrInsert(hlLcf, tlPLcf, ilLcfRecLen, INDEXKEY3)
                Loop While ilRet = BTRV_ERR_CONFLICT
                If ilRet <> BTRV_ERR_NONE Then
                    mMovePrgToPending = ilRet
                    Exit Function
                End If
            End If
        End If
        'Find Location to place library
        ilRet = btrGetPosition(hlLcf, llPosition)    'Get position incase of conflict
        'First remove any that are overlapped, then insert
        If ilRemoveState Then
            For ilLoop = LBound(tlPLcf.lLvfCode) To UBound(tlPLcf.lLvfCode) Step 1
                ilRemove = False
                If tlPLcf.lLvfCode(ilLoop) = 0 Then
                    Exit For
                End If
                gUnpackTime tlPLcf.iTime(0, ilLoop), tlPLcf.iTime(1, ilLoop), "A", "1", slLTime
                If gTimeToCurrency(slPTime, False) = gTimeToCurrency(slLTime, False) Then 'Replace
                    'Remove entry
                    ilRemove = True
                Else
                    gUnpackLength tlLvf.iLen(0), tlLvf.iLen(1), "3", False, slStr
                    gAddTimeLength slPTime, slStr, "A", "1", slPEndTime, slXMid
                    If (gTimeToCurrency(slPTime, False) < gTimeToCurrency(slLTime, False)) And (gTimeToCurrency(slPEndTime, True) - 1 >= gTimeToCurrency(slLTime, False)) Then 'Replace
                        'Remove entry
                        ilRemove = True
                    End If
                End If
                If Not ilRemove Then
                    tlLvfSrchKey.lCode = tlPLcf.lLvfCode(ilLoop)
                    ilRet = btrGetEqual(hlLvf, tlPLvf, ilLvfRecLen, tlLvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    If ilRet = BTRV_ERR_NONE Then
                        gUnpackLength tlPLvf.iLen(0), tlPLvf.iLen(1), "3", False, slStr
                        gAddTimeLength slLTime, slStr, "A", "1", slLEndTime, slXMid
                        If (gTimeToCurrency(slLTime, False) < gTimeToCurrency(slPTime, False)) And (gTimeToCurrency(slLEndTime, True) - 1 >= gTimeToCurrency(slPTime, False)) Then 'Replace
                            'Remove entry
                            ilRemove = True
                        End If
                    End If
                End If
                llNPosition = llPosition
                ilStartLoop = ilLoop + 1
                Do While ilRemove
                    Do
                        ilRet = btrGetDirect(hlLcf, tlPLcf, ilLcfRecLen, llNPosition, INDEXKEY0, BTRV_LOCK_NONE)
                        ilRet = btrGetNext(hlLcf, tlPNLcf, ilLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            tlPNLcf.lLvfCode(LBound(tlPNLcf.lLvfCode)) = 0
                        End If
                        If (tlPNLcf.iType <> ilType) Or (tlPNLcf.sStatus <> "P") Or (tlPNLcf.iVefCode <> ilVefCode) Or (tlPNLcf.iLogDate(0) <> ilLogDate0) Or (tlPNLcf.iLogDate(1) <> ilLogDate1) Then
                            tlPNLcf.lLvfCode(LBound(tlPNLcf.lLvfCode)) = 0
                        End If
                        ilRet = btrGetDirect(hlLcf, tlPLcf, ilLcfRecLen, llNPosition, INDEXKEY0, BTRV_LOCK_NONE)
                        'tmSRec = tlPLcf
                        'ilRet = gGetByKeyForUpdate("Lcf", hlLcf, tmSRec)
                        'tlPLcf = tmSRec
                        'If ilRet <> BTRV_ERR_NONE Then
                        '    mMovePrgToPending = ilRet
                        '    Exit Function
                        'End If
                        For ilIndex = ilStartLoop To UBound(tlPLcf.lLvfCode) Step 1
                            tlPLcf.lLvfCode(ilIndex - 1) = tlPLcf.lLvfCode(ilIndex)
                            tlPLcf.iTime(0, ilIndex - 1) = tlPLcf.iTime(0, ilIndex)
                            tlPLcf.iTime(1, ilIndex - 1) = tlPLcf.iTime(1, ilIndex)
                        Next ilIndex
                        If tlPNLcf.lLvfCode(LBound(tlPNLcf.lLvfCode)) <> 0 Then
                            tlPLcf.lLvfCode(UBound(tlPLcf.lLvfCode)) = tlPNLcf.lLvfCode(LBound(tlPNLcf.lLvfCode))
                            tlPLcf.iTime(0, UBound(tlPLcf.lLvfCode)) = tlPNLcf.iTime(0, LBound(tlPNLcf.lLvfCode))
                            tlPLcf.iTime(1, UBound(tlPLcf.lLvfCode)) = tlPNLcf.iTime(1, LBound(tlPNLcf.lLvfCode))
                        Else
                            tlPLcf.lLvfCode(UBound(tlPLcf.lLvfCode)) = 0
                            tlPLcf.iTime(0, UBound(tlPLcf.lLvfCode)) = 0
                            tlPLcf.iTime(1, UBound(tlPLcf.lLvfCode)) = 0
                        End If
                        ilRet = btrUpdate(hlLcf, tlPLcf, ilLcfRecLen)
                    Loop While ilRet = BTRV_ERR_CONFLICT
                    If ilRet <> BTRV_ERR_NONE Then
                        mMovePrgToPending = ilRet
                        Exit Function
                    End If
                    If tlPNLcf.lLvfCode(LBound(tlPNLcf.lLvfCode)) <> 0 Then
                        ilRemove = True
                        ilRet = btrGetNext(hlLcf, tlPLcf, ilLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        ilRet = btrGetPosition(hlLcf, llNPosition)    'Get position incase of conflict
                        ilStartLoop = LBound(tlPLcf.lLvfCode) + 1
                    Else
                        ilRemove = False
                        ilRet = btrGetDirect(hlLcf, tlPLcf, ilLcfRecLen, llPosition, INDEXKEY0, BTRV_LOCK_NONE)
                    End If
                Loop
            Next ilLoop
            ilRet = btrGetDirect(hlLcf, tlPLcf, ilLcfRecLen, llPosition, INDEXKEY0, BTRV_LOCK_NONE)
        End If
        ilSeqNo = ilSeqNo + 1   'If time location not found
        For ilLoop = LBound(tlPLcf.lLvfCode) To UBound(tlPLcf.lLvfCode) Step 1
            If tlPLcf.lLvfCode(ilLoop) = 0 Then
                Do
                    ilRet = btrGetDirect(hlLcf, tlPLcf, ilLcfRecLen, llPosition, INDEXKEY0, BTRV_LOCK_NONE)
                    'tmSRec = tlPLcf
                    'ilRet = gGetByKeyForUpdate("Lcf", hlLcf, tmSRec)
                    'tlPLcf = tmSRec
                    'If ilRet <> BTRV_ERR_NONE Then
                    '    mMovePrgToPending = ilRet
                    '    Exit Function
                    'End If
                    tlPLcf.lLvfCode(ilLoop) = llLvfCode
                    tlPLcf.iTime(0, ilLoop) = ilTime0
                    tlPLcf.iTime(1, ilLoop) = ilTime1
                    ilRet = btrUpdate(hlLcf, tlPLcf, ilLcfRecLen)
                Loop While ilRet = BTRV_ERR_CONFLICT
                mMovePrgToPending = ilRet
                Exit Function
            Else
                gUnpackTime tlPLcf.iTime(0, ilLoop), tlPLcf.iTime(1, ilLoop), "A", "1", slLTime
'                If gTimeToCurrency(slPTime, False) = gTimeToCurrency(slLTime, False) Then 'Replace
'                    Do
'                        ilRet = btrGetDirect(hlLcf, tlPLcf, ilLcfRecLen, llPosition, INDEXKEY0, BTRV_LOCK_NONE)
'                        tlPLcf.lLnfCode(ilLoop) = llLnfCode
'                        tlPLcf.iTime(0, ilLoop) = ilTime0
'                        tlPLcf.iTime(1, ilLoop) = ilTime1
'                        ilRet = btrUpdate(hlLcf, tlPLcf, ilLcfRecLen)
'                    Loop While ilRet = BTRV_ERR_CONFLICT
'                    Exit Sub
'                End If
                If gTimeToCurrency(slPTime, False) < gTimeToCurrency(slLTime, False) Then
                    'Insert prior to this time
                    If tlPLcf.lLvfCode(UBound(tlPLcf.lLvfCode)) = 0 Then    'Room within this record
                        Do
                            ilRet = btrGetDirect(hlLcf, tlPLcf, ilLcfRecLen, llPosition, INDEXKEY0, BTRV_LOCK_NONE)
                            'tmSRec = tlPLcf
                            'ilRet = gGetByKeyForUpdate("Lcf", hlLcf, tmSRec)
                            'tlPLcf = tmSRec
                            'If ilRet <> BTRV_ERR_NONE Then
                            '    mMovePrgToPending = ilRet
                            '    Exit Function
                            'End If
                            For ilIndex = UBound(tlPLcf.lLvfCode) - 1 To ilLoop Step -1
                                tlPLcf.lLvfCode(ilIndex + 1) = tlPLcf.lLvfCode(ilIndex)
                                tlPLcf.iTime(0, ilIndex + 1) = tlPLcf.iTime(0, ilIndex)
                                tlPLcf.iTime(1, ilIndex + 1) = tlPLcf.iTime(1, ilIndex)
                            Next ilIndex
                            tlPLcf.lLvfCode(ilLoop) = llLvfCode
                            tlPLcf.iTime(0, ilLoop) = ilTime0
                            tlPLcf.iTime(1, ilLoop) = ilTime1
                            ilRet = btrUpdate(hlLcf, tlPLcf, ilLcfRecLen)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        mMovePrgToPending = ilRet
                        Exit Function
                    Else
                        'Move down and move last to next sequence no
                        Do
                            ilRet = btrGetDirect(hlLcf, tlPLcf, ilLcfRecLen, llPosition, INDEXKEY0, BTRV_LOCK_NONE)
                            'tmSRec = tlPLcf
                            'ilRet = gGetByKeyForUpdate("Lcf", hlLcf, tmSRec)
                            'tlPLcf = tmSRec
                            'If ilRet <> BTRV_ERR_NONE Then
                            '    mMovePrgToPending = ilRet
                            '    Exit Function
                            'End If
                            llDnLvfCode = tlPLcf.lLvfCode(UBound(tlPLcf.lLvfCode))
                            ilDnTime0 = tlPLcf.iTime(0, UBound(tlPLcf.lLvfCode))
                            ilDnTime1 = tlPLcf.iTime(1, UBound(tlPLcf.lLvfCode))
                            For ilIndex = UBound(tlPLcf.lLvfCode) - 1 To ilLoop Step -1
                                tlPLcf.lLvfCode(ilIndex + 1) = tlPLcf.lLvfCode(ilIndex)
                                tlPLcf.iTime(0, ilIndex + 1) = tlPLcf.iTime(0, ilIndex)
                                tlPLcf.iTime(1, ilIndex + 1) = tlPLcf.iTime(1, ilIndex)
                            Next ilIndex
                            tlPLcf.lLvfCode(ilLoop) = llLvfCode
                            tlPLcf.iTime(0, ilLoop) = ilTime0
                            tlPLcf.iTime(1, ilLoop) = ilTime1
                            ilRet = btrUpdate(hlLcf, tlPLcf, ilLcfRecLen)
                        Loop While ilRet = BTRV_ERR_CONFLICT
                        If ilRet <> BTRV_ERR_NONE Then
                            mMovePrgToPending = ilRet
                            Exit Function
                        End If
                        llLvfCode = llDnLvfCode
                        ilTime0 = ilDnTime0
                        ilTime1 = ilDnTime1
                        ilRemoveState = False
                    End If
                End If
            End If
        Next ilLoop
    Loop While ilSeqNo >= 0
    mMovePrgToPending = BTRV_ERR_NONE
    Exit Function
End Function

Private Function mRSFsForSpot(llSDFCode As Long) As Integer()

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mSetSepValues                   *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Set length and counts. Used by  *
'*                     gGetLineSchParameters           *
'*                                                     *
'*******************************************************
Private Sub mSetSepValues(slStartTime As String, slEndTime As String, llLength As Long, ilTHour() As Integer, ilTQH() As Integer, ilAHour() As Integer, ilAQH() As Integer)
    Dim llLnStart As Long
    Dim llLnEnd As Long
    Dim llIndex As Long
    Dim ilIndex As Integer
    llLnStart = CLng(gTimeToCurrency(slStartTime, False))
    llLnEnd = CLng(gTimeToCurrency(slEndTime, True)) - 1
    llLength = llLength + llLnEnd - llLnStart
    For llIndex = llLnStart To llLnEnd Step 3600
        'ilIndex = (llIndex \ 3600&) + 1
        ilIndex = (llIndex \ 3600&)
        ilAHour(ilIndex) = ilTHour(ilIndex)
    Next llIndex
    'Get last hour (i.e. 11:41 {42120} thru 7:22 {69720}; last hour is missed)
    'ilIndex = (llLnEnd \ 3600&) + 1
    ilIndex = (llLnEnd \ 3600&)
    ilAHour(ilIndex) = ilTHour(ilIndex)
    If llLnStart + 3600 <= llLnEnd Then
        'For ilIndex = 1 To 4 Step 1
        For ilIndex = 0 To 3 Step 1
            ilAQH(ilIndex) = ilTQH(ilIndex)
        Next ilIndex
    Else
        For llIndex = llLnStart \ 900 To llLnEnd \ 900 Step 1
            'ilIndex = (llIndex Mod 4) + 1
            ilIndex = (llIndex Mod 4)
            ilAQH(ilIndex) = ilTQH(ilIndex)
        Next llIndex
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mVehCompConflictTest            *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Competitive conflict test       *
'*                     for selling to airing vehicles  *
'*                                                     *
'*******************************************************
Private Function mVehCompConflictTest(hlSsf As Integer, tlSsf As SSF, tlSpotMove() As SPOTMOVE, ilMnfComp0 As Integer, ilMnfComp1 As Integer, tlVcf As VCF) As Integer
    Dim ilIndex As Integer
    Dim ilVehIndex As Integer
    Dim ilRet As Integer
    Dim ilSpotIndex As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilBypass As Integer
    Dim ilBypassIndex As Integer
    ilDate0 = tlSsf.iDate(0)
    ilDate1 = tlSsf.iDate(1)
    For ilVehIndex = 1 To 5 Step 1
        If tlVcf.iCSV(ilVehIndex - 1) > 0 Then
            imSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
            tmSsfSrchKey.iType = 0 'slType
            tmSsfSrchKey.iVefCode = tlVcf.iCSV(ilVehIndex - 1)
            tmSsfSrchKey.iDate(0) = ilDate0
            tmSsfSrchKey.iDate(1) = ilDate1
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetGreaterOrEqual(hlSsf, tmCTSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
            Do While (ilRet = BTRV_ERR_NONE) And (tmCTSsf.iType = 0) And (tmCTSsf.iVefCode = tlVcf.iCSV(ilVehIndex - 1)) And (tmCTSsf.iDate(0) = ilDate0) And (tmCTSsf.iDate(1) = ilDate1)
                For ilIndex = 1 To tmCTSsf.iCount Step 1
                    tmAvailTest = tmCTSsf.tPas(ADJSSFPASBZ + ilIndex)
                    'If tmAvailTest.iRecType = 2 Then
                    If (tmAvailTest.iRecType = 2) Or (tmAvailTest.iRecType = 8) Or (tmAvailTest.iRecType = 9) Then
                        If (tmAvailTest.iTime(0) = tlVcf.iCST(0, ilVehIndex - 1)) And (tmAvailTest.iTime(1) = tlVcf.iCST(1, ilVehIndex - 1)) Then
                            'Avail found- test for conflicts
                            For ilSpotIndex = ilIndex + 1 To ilIndex + tmAvailTest.iNoSpotsThis Step 1
                                LSet tmSpotTest = tmCTSsf.tPas(ADJSSFPASBZ + ilSpotIndex)
                                ilBypass = False
                                For ilBypassIndex = LBound(tlSpotMove) To UBound(tlSpotMove) - 1 Step 1
                                    If (tmSpotTest.lSdfCode = tlSpotMove(ilBypassIndex).lSdfCode) Then
                                        ilBypass = True
                                        Exit For
                                    End If
                                Next ilBypassIndex
                                If Not ilBypass Then
                                    If (ilMnfComp0 <> 0) And ((ilMnfComp0 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp0 = tmSpotTest.iMnfComp(1))) Then
                                        mVehCompConflictTest = False
                                        Exit Function
                                    End If
                                    If (ilMnfComp1 <> 0) And ((ilMnfComp1 = tmSpotTest.iMnfComp(0)) Or (ilMnfComp1 = tmSpotTest.iMnfComp(1))) Then
                                        mVehCompConflictTest = False
                                        Exit Function
                                    End If
                                End If
                            Next ilSpotIndex
                            Exit Do
                        End If
                        'Exit Do    'exit avail loop test
                    End If
                Next ilIndex
                'If (tmCTSsf.iNextTime(0) = 1) And (tmCTSsf.iNextTime(1) = 0) Then
                    Exit Do
                'Else
                '    imSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                '    ilRet = gSSFGetNext(hlSsf, tmCTSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                'End If
            Loop
        End If
    Next ilVehIndex
    mVehCompConflictTest = True
    Exit Function
End Function

Public Function gCheckAvailAttributes(ilSpotStatus As Integer) As Integer
'
'
'   ilSpotStatus(I)- Spots in conflict between todays date and Last Log Date: 0=Pre-empt; 1=Retain
'
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilUpper As Integer
    Dim ilAnf As Integer
    Dim ilVef As Integer
    Dim slNowDate As String
    Dim ilVpfIndex As Integer
    Dim llLLD As Long
    Dim ilEvtFrom As Integer
    Dim ilEvtTo As Integer
    Dim ilAvEvt As Integer
    Dim ilAvInfo As Integer
    Dim ilSpot As Integer
    Dim ilNoSpotsThis As Integer
    Dim ilAvailFd As Integer
    Dim ilSpotOK As Integer
    Dim ilSSFChanged As Integer
    Dim llSsfRecPos As Long
    Dim llSsfDate As Long
    Dim slOrigSchStatus As String
    Dim slXSpotType As String
    'Dim llSdfRecPos As Long
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilTime0 As Integer
    Dim ilTime1 As Integer
    Dim ilOrigSchVef As Integer
    Dim ilOrigGameNo As Integer

    'Avail names
    Dim hlAnf As Integer            'Avail name file handle
    'Spot summary
    Dim hlSsf As Integer        'Spot summary file handle
    Dim ilSsfRecLen As Integer  'SSF record length

    Dim hlSdf As Integer
    Dim hlSmf As Integer
    Dim hlStf As Integer
    Dim hlGsf As Integer
    Dim hlSxf As Integer

    gGetSchParameters
    gObtainMissedReasonCode
    'ReDim lgReschSdfCode(1 To 1) As Long
    ReDim lgReschSdfCode(0 To 0) As Long
    ilUpper = LBound(tgSsf)
    ilSsfRecLen = Len(tgSsf(ilUpper))  'Get and save SSF record length
    hlSsf = CBtrvTable(TWOHANDLES)        'Create SSF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gCheckAvailAttributes-Open Ssf(1)"
        ilRet = btrClose(hlSsf)
        btrDestroy hlSsf
        gCheckAvailAttributes = False
        Exit Function
    End If
    imAnfRecLen = Len(tmAnf)  'Get and save ANF record length
    hlAnf = CBtrvTable(ONEHANDLE)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gCheckAvailAttributes-Open Anf(2)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlAnf)
        btrDestroy hlSsf
        btrDestroy hlAnf
        gCheckAvailAttributes = False
        Exit Function
    End If
    imSdfRecLen = Len(tmSdf)  'Get and save ANF record length
    hlSdf = CBtrvTable(TWOHANDLES)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gCheckAvailAttributes-Open Sdf(3)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hlSdf)
        btrDestroy hlSsf
        btrDestroy hlAnf
        btrDestroy hlSdf
        gCheckAvailAttributes = False
        Exit Function
    End If
    imSmfRecLen = Len(tmSmf)  'Get and save ANF record length
    hlSmf = CBtrvTable(TWOHANDLES)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gCheckAvailAttributes-Open Smf(4)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlSmf)
        btrDestroy hlSsf
        btrDestroy hlAnf
        btrDestroy hlSdf
        btrDestroy hlSmf
        gCheckAvailAttributes = False
        Exit Function
    End If
    imStfRecLen = Len(tmStf)  'Get and save ANF record length
    hlStf = CBtrvTable(TWOHANDLES)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlStf, "", sgDBPath & "Stf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gCheckAvailAttributes-Open Smf(4)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlStf)
        btrDestroy hlSsf
        btrDestroy hlAnf
        btrDestroy hlSdf
        btrDestroy hlSmf
        btrDestroy hlStf
        gCheckAvailAttributes = False
        Exit Function
    End If
    hlGsf = CBtrvTable(TWOHANDLES)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gCheckAvailAttributes-Open Smf(4)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlStf)
        ilRet = btrClose(hlGsf)
        btrDestroy hlSsf
        btrDestroy hlAnf
        btrDestroy hlSdf
        btrDestroy hlSmf
        btrDestroy hlStf
        btrDestroy hlGsf
        gCheckAvailAttributes = False
        Exit Function
    End If
    hlSxf = CBtrvTable(TWOHANDLES)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSxf, "", sgDBPath & "Sxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        igBtrError = gConvertErrorCode(ilRet)
        sgErrLoc = "gCheckAvailAttributes-Open Smf(4)"
        ilRet = btrClose(hlSsf)
        ilRet = btrClose(hlAnf)
        ilRet = btrClose(hlSdf)
        ilRet = btrClose(hlSmf)
        ilRet = btrClose(hlStf)
        ilRet = btrClose(hlGsf)
        ilRet = btrClose(hlSxf)
        btrDestroy hlSsf
        btrDestroy hlAnf
        btrDestroy hlSdf
        btrDestroy hlSmf
        btrDestroy hlStf
        btrDestroy hlGsf
        btrDestroy hlSxf
        gCheckAvailAttributes = False
        Exit Function
    End If
    ilRet = gObtainVef()
    'Build array of avails
    ReDim tmSAnf(0 To 0) As ANF
    ilUpper = 0
    ilRet = btrGetFirst(hlAnf, tmSAnf(ilUpper), imAnfRecLen, 0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ilUpper = ilUpper + 1
        ReDim Preserve tmSAnf(0 To ilUpper) As ANF
        ilRet = btrGetNext(hlAnf, tmSAnf(ilUpper), imAnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get next record
    Loop


    'Loop on ssf from today+1 to end
    'test if attributes changed and if so are the spots Ok
    slNowDate = Format$(gNow(), "m/d/yy")
    slNowDate = gIncOneDay(slNowDate)
    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If (tgMVef(ilVef).sType = "S") Or (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "G") Then
            ilVpfIndex = gBinarySearchVpfPlus(tgMVef(ilVef).iCode)
            If ilVpfIndex <> -1 Then
                gUnpackDateLong tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), llLLD
            Else
                llLLD = 0
            End If
            'tmSsfSrchKey.iType = 0
            'tmSsfSrchKey.iVefCode = tgMVef(ilVef).iCode
            'gPackDate slNowDate, tmSsfSrchKey.iDate(0), tmSsfSrchKey.iDate(1)
            'tmSsfSrchKey.iStartTime(0) = 0
            'tmSsfSrchKey.iStartTime(1) = 0
            ilUpper = LBound(tgSsf)
            tmSsfSrchKey2.iVefCode = tgMVef(ilVef).iCode
            gPackDate slNowDate, tmSsfSrchKey2.iDate(0), tmSsfSrchKey2.iDate(1)
            imSsfRecLen = Len(tgSsf(ilUpper))
            ilRet = gSSFGetGreaterOrEqualKey2(hlSsf, tgSsf(ilUpper), imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tgSsf(ilUpper).iVefCode = tgMVef(ilVef).iCode)
                ilRet = gSSFGetPosition(hlSsf, llSsfRecPos)
                ilSSFChanged = False
                gUnpackDateLong tgSsf(ilUpper).iDate(0), tgSsf(ilUpper).iDate(1), llSsfDate
                tmSsf.iCount = 0
                tmSsf.iDate(0) = tgSsf(ilUpper).iDate(0)
                tmSsf.iDate(1) = tgSsf(ilUpper).iDate(1)
                'tmSsf.iNextTime(0) = tgSsf(ilUpper).iNextTime(0)
                'tmSsf.iNextTime(1) = tgSsf(ilUpper).iNextTime(1)
                tmSsf.iStartTime(0) = tgSsf(ilUpper).iStartTime(0)
                tmSsf.iStartTime(1) = tgSsf(ilUpper).iStartTime(1)
                tmSsf.iType = tgSsf(ilUpper).iType
                tmSsf.iVefCode = tgSsf(ilUpper).iVefCode

                ilEvtFrom = 1
                ilEvtTo = 1
                Do While ilEvtFrom <= tgSsf(ilUpper).iCount
                   LSet tmAvail = tgSsf(ilUpper).tPas(ADJSSFPASBZ + ilEvtFrom)
                    If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9) Then
                        ilNoSpotsThis = tmAvail.iNoSpotsThis
                        tmAvail.iNoSpotsThis = 0
                        tmSsf.tPas(ADJSSFPASBZ + ilEvtTo) = tmAvail
                        ilAvEvt = ilEvtTo
                        ilEvtTo = ilEvtTo + 1
                        tmSsf.iCount = tmSsf.iCount + 1
                        ilAvailFd = False
                        For ilAnf = 0 To UBound(tmSAnf) - 1 Step 1
                            If tmAvail.ianfCode = tmSAnf(ilAnf).iCode Then
                                ilAvailFd = True
                                ilAvInfo = tmAvail.iAvInfo And &H1F
                                If tmSAnf(ilAnf).sSustain = "Y" Then
                                    ilAvInfo = ilAvInfo Or SSSUSTAINING
                                End If
                                If tmSAnf(ilAnf).sSponsorship = "Y" Then
                                    ilAvInfo = ilAvInfo Or SSSPONSORSHIP
                                End If
                                If tmSAnf(ilAnf).sBookLocalFeed = "L" Then
                                    ilAvInfo = ilAvInfo Or SSLOCALONLY
                                End If
                                If tmSAnf(ilAnf).sBookLocalFeed = "F" Then
                                    ilAvInfo = ilAvInfo Or SSFEEDONLY
                                End If
                                If tmAvail.iAvInfo <> ilAvInfo Then
                                    tmAvail.iAvInfo = ilAvInfo
                                    ilSSFChanged = True
                                    For ilSpot = 1 To ilNoSpotsThis Step 1
                                       LSet tmSpot = tgSsf(ilUpper).tPas(ADJSSFPASBZ + ilEvtFrom + ilSpot)
                                        If (tmSpot.iRecType And SSAVAILBUY) = SSAVAILBUY Then
                                            ilSpotOK = True
                                        Else
                                            If (ilAvInfo And SSSUSTAINING) <> SSSUSTAINING Then
                                                ilSpotOK = False
                                                tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                'If ilRet = BTRV_ERR_NONE Then
                                                '    ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                                                'End If
                                            Else
                                                ilSpotOK = True
                                                If (tgSpf.sSystemType = "R") Then
                                                    If ((ilAvInfo And SSLOCALONLY) = SSLOCALONLY) Or ((ilAvInfo And SSFEEDONLY) = SSFEEDONLY) Then
                                                        tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                        ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            'ilRet = btrGetPosition(hlSdf, llSdfRecPos)
                                                            If tmSdf.lChfCode > 0 Then
                                                                If (ilAvInfo And SSLOCALONLY) <> SSLOCALONLY Then
                                                                    ilSpotOK = False
                                                                End If
                                                            Else
                                                                ilSpotOK = False
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                        If Not ilSpotOK Then
                                            'Pre-empt spot if after last log date
                                            If (llSsfDate <= llLLD) And (ilSpotStatus = 1) Then
                                                ilSpotOK = True
                                            End If
                                        End If
                                        If ilSpotOK Then
                                            tmSsf.tPas(ADJSSFPASBZ + ilEvtTo) = tgSsf(ilUpper).tPas(ADJSSFPASBZ + ilEvtFrom + ilSpot)
                                            tmAvail.iNoSpotsThis = tmAvail.iNoSpotsThis + 1
                                            ilEvtTo = ilEvtTo + 1
                                            tmSsf.iCount = tmSsf.iCount + 1
                                        Else
                                            ilRet = gMakeTracer(hlSdf, tmSdf, 0, hlStf, llLLD, "C", "C", tmSdf.iRotNo, hlGsf)
                                            ilRet = gRemoveSmf(hlSmf, tmSmf, tmSdf, hlSxf)
                                            If tmSdf.sSpotType = "X" Then
                                                slXSpotType = "X"
                                                If ((tmSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                                                    slXSpotType = ""
                                                End If
                                            Else
                                                slXSpotType = ""
                                            End If
                                            If ((tmSdf.sSpotType = "T") And (tgSpf.sSchdRemnant <> "Y")) Or (tmSdf.sSpotType = "Q") Or ((tmSdf.sSpotType = "S") And (tgSpf.sSchdPSA <> "Y")) Or ((tmSdf.sSpotType = "M") And (tgSpf.sSchdPromo <> "Y")) Or (slXSpotType = "X") Then
                                                Do
                                                    ilRet = btrDelete(hlSdf)
                                                    If ilRet = BTRV_ERR_CONFLICT Then
                                                        'ilCRet = btrGetDirect(hlSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                        tmSdfSrchKey3.lCode = tmSdf.lCode
                                                        ilCRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                    End If
                                                Loop While ilRet = BTRV_ERR_CONFLICT
                                            Else
                                                ilDate0 = tmSdf.iDate(0)
                                                ilDate1 = tmSdf.iDate(1)
                                                ilTime0 = tmSdf.iTime(0)
                                                ilTime1 = tmSdf.iTime(1)
                                                ilOrigSchVef = tmSdf.iVefCode
                                                ilOrigGameNo = tmSdf.iGameNo
                                                'Update Sdf record
                                                Do
                                                    'ilRet = btrGetDirect(hlSdf, tmSdf, imSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                                    tmSdfSrchKey3.lCode = tmSdf.lCode
                                                    ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                                                    tmSdf.sSchStatus = "M"
                                                    tmSdf.iMnfMissed = igMnfMissed
                                                    tmSdf.iDate(0) = ilDate0
                                                    tmSdf.iDate(1) = ilDate1
                                                    tmSdf.iTime(0) = ilTime0
                                                    tmSdf.iTime(1) = ilTime1
                                                    tmSdf.iVefCode = ilOrigSchVef
                                                    tmSdf.iGameNo = ilOrigGameNo
                                                    tmSdf.lSmfCode = 0
                                                    If ((tmSmf.lMtfCode <> 0) And (lgMtfNoRecs <> 0)) And ((slOrigSchStatus = "G") Or (slOrigSchStatus = "O")) Then
                                                        tmSdf.sTracer = "*"
                                                        tmSdf.lSmfCode = tmSmf.lMtfCode
                                                    End If
                                                    tmSdf.sXCrossMidnight = "N"
                                                    tmSdf.sWasMG = "N"
                                                    tmSdf.sFromWorkArea = "N"
                                                    tmSdf.iUrfCode = tgUrf(0).iCode
                                                    ilRet = btrUpdate(hlSdf, tmSdf, imSdfRecLen)
                                                Loop While ilRet = BTRV_ERR_CONFLICT
                                                lgReschSdfCode(UBound(lgReschSdfCode)) = tmSdf.lCode
                                                'ReDim Preserve lgReschSdfCode(1 To UBound(lgReschSdfCode) + 1) As Long
                                                ReDim Preserve lgReschSdfCode(LBound(lgReschSdfCode) To UBound(lgReschSdfCode) + 1) As Long
                                            End If
                                        End If
                                    Next ilSpot
                                Else
                                    For ilSpot = 1 To ilNoSpotsThis Step 1
                                        tmSsf.tPas(ADJSSFPASBZ + ilEvtTo) = tgSsf(ilUpper).tPas(ADJSSFPASBZ + ilEvtFrom + ilSpot)
                                        tmAvail.iNoSpotsThis = tmAvail.iNoSpotsThis + 1
                                        ilEvtTo = ilEvtTo + 1
                                        tmSsf.iCount = tmSsf.iCount + 1
                                    Next ilSpot
                                End If
                                Exit For
                            End If
                        Next ilAnf
                        If Not ilAvailFd Then
                            For ilSpot = 1 To ilNoSpotsThis Step 1
                                tmSsf.tPas(ADJSSFPASBZ + ilEvtTo) = tgSsf(ilUpper).tPas(ADJSSFPASBZ + ilEvtFrom + ilSpot)
                                tmAvail.iNoSpotsThis = tmAvail.iNoSpotsThis + 1
                                ilEvtTo = ilEvtTo + 1
                                tmSsf.iCount = tmSsf.iCount + 1
                            Next ilSpot
                        End If
                        tmSsf.tPas(ADJSSFPASBZ + ilAvEvt) = tmAvail
                        ilEvtFrom = ilEvtFrom + ilNoSpotsThis    'bypass spots
                    Else
                        tmSsf.tPas(ADJSSFPASBZ + ilEvtTo) = tmAvail
                        ilEvtTo = ilEvtTo + 1
                        tmSsf.iCount = tmSsf.iCount + 1
                    End If
                    ilEvtFrom = ilEvtFrom + 1   'Increment to next event
                Loop
                If ilSSFChanged Then
                    imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                    ilRet = gSSFGetDirect(hlSsf, tgSsf(ilUpper), imSsfRecLen, llSsfRecPos, INDEXKEY2, BTRV_LOCK_NONE)
                    imSsfRecLen = igSSFBaseLen + tmSsf.iCount * Len(tmAvail)
                    ilRet = gSSFUpdate(hlSsf, tmSsf, imSsfRecLen)
                End If
                imSsfRecLen = Len(tgSsf(ilUpper))
                ilRet = gSSFGetNext(hlSsf, tgSsf(ilUpper), imSsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
            Loop
        End If
    Next ilVef

    Erase tmSAnf
    ilRet = btrClose(hlSsf)
    ilRet = btrClose(hlAnf)
    ilRet = btrClose(hlSdf)
    ilRet = btrClose(hlSmf)
    ilRet = btrClose(hlStf)
    ilRet = btrClose(hlGsf)
    ilRet = btrClose(hlSxf)
    btrDestroy hlSsf
    btrDestroy hlAnf
    btrDestroy hlSdf
    btrDestroy hlSmf
    btrDestroy hlStf
    btrDestroy hlGsf
    btrDestroy hlSxf

    If UBound(lgReschSdfCode) > LBound(lgReschSdfCode) Then

        If gOpenSchFiles() Then
            ilRet = gReSchSpots(False, 0, "YYYYYYY", 0, 86400)
            gCloseSchFiles
        End If
    End If

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:gCreateBBSpots                  *
'*                                                     *
'*             Created:3/23/05       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Create BB spots if required     *
'*                     1. Get LLC structure for day    *
'*                     2. Get Current Spots for day    *
'*                        Including BB Spots           *
'*                     3. Loop on Regular Spots        *
'*                        Test if BB required for Spot *
'*                        Look for BB Placement in LLC *
'*                        See if Spot already exist, if*
'*                        create BB spot               *
'*                        Assign Copy to BB Spot       *
'*                     4. Remove any extra BB Spots    *
'*                                                     *
'*******************************************************
Public Function gCreateBBSpots(hlSdf As Integer, ilVefCode As Integer, slDate As String, Optional ilInGameNo As Integer = -1, Optional hlInClf As Integer = -1, Optional hlInLcf As Integer = -1, Optional ilMode As Integer = 0, Optional llChfCode As Long = -1, Optional ilLineNo As Integer = -1) As Integer
'
'   hlSdf(I)- SDF File Handle
'   ilVefCode(I)- Conventional or Selling vehicle code
'   slDate(I)- Date to create Billboard spots
'   hlInClf(I)- CLF File Handle
'   hlInLcf(I)- LCF File Handle
'   ilMode(I)- 0=Create and remove unused BB; 1=Test and Remove unused only; 2=Remove all BB only
'   llChfCode(I)- Match contract code reference; -1=all
'   ilLineNo(I)- Match line number; -1=All
'
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim ilEvtRet As Integer
    Dim ilSdf As Integer
    Dim ilSdfRet As Integer
    Dim tlSdf As SDF
    Dim ilLLC As Integer
    Dim ilFind As Integer
    Dim ilChkSdf As Integer
    Dim llAvailTime As Long
    Dim llSDFTime As Long
    Dim llChkTime As Long
    Dim llPass1AvailTime As Long
    Dim ilFound As Integer
    Dim ilStart As Integer
    Dim ilEnd As Integer
    Dim ilStep As Integer
    Dim slBBType As String
    Dim slSpotType As String
    Dim ilLen As Integer
    Dim ilPass As Integer
    Dim ilSPass As Integer
    Dim ilEPass As Integer
    Dim ilType As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilLoop As Integer
    Dim ilVefIndex As Integer
    Dim tlCreateSdf As SDF
    ReDim ilEvtType(0 To 14) As Integer
    ReDim tlRegSdf(0 To 0) As SDF
    ReDim tlBBSdf(0 To 0) As SDF

    If tgSpf.sUsingBBs <> "Y" Then
        gCreateBBSpots = True
        Exit Function
    End If
    If gDateValue(slDate) < gDateValue(Format$(gNow(), "m/d/yy")) Then
        gCreateBBSpots = True
        Exit Function
    End If

    imClfRecLen = Len(tmClf)  'Get and save CLF record length
    If hlInClf = -1 Then
        hmClf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        On Error GoTo 0
        ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            gCreateBBSpots = False
            Exit Function
        End If
    Else
        hmClf = hlInClf
    End If
    imLcfRecLen = Len(tmLcf)  'Get and save CLF record length
    If hlInLcf = -1 Then
        hmLcf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        On Error GoTo 0
        ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            If hlInClf = -1 Then
                ilRet = btrClose(hmClf)
                btrDestroy hmClf
            End If
            gCreateBBSpots = False
            Exit Function
        End If
    Else
        hmLcf = hlInLcf
    End If
        
    imSdfRecLen = Len(tlCreateSdf)
    'Build array of game numbers to be processed
    ilVefIndex = gBinarySearchVef(ilVefCode)
    If tgMVef(ilVefIndex).sType <> "G" Then
        ReDim ilGameNo(0 To 1) As Integer
        ilGameNo(0) = 0
    Else
        ReDim ilGameNo(0 To 0) As Integer
        tmLcfSrchKey2.iVefCode = ilVefCode
        gPackDate slDate, ilLogDate0, ilLogDate1
        tmLcfSrchKey2.iLogDate(0) = ilLogDate0
        tmLcfSrchKey2.iLogDate(1) = ilLogDate1
        ilRet = btrGetGreaterOrEqual(hmLcf, tmLcf, imLcfRecLen, tmLcfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
        Do While (ilRet = BTRV_ERR_NONE) And (tmLcf.iVefCode = ilVefCode) And (tmLcf.iLogDate(0) = ilLogDate0) And (tmLcf.iLogDate(1) = ilLogDate1)
            If (ilInGameNo = -1) Or (ilInGameNo = tmLcf.iType) Then
                ilGameNo(UBound(ilGameNo)) = tmLcf.iType
                ReDim Preserve ilGameNo(0 To UBound(ilGameNo) + 1) As Integer
            End If
            ilRet = btrGetNext(hmLcf, tmLcf, imLcfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    For ilLoop = 0 To UBound(ilGameNo) - 1 Step 1
        ilType = ilGameNo(ilLoop)
        ReDim tlLLC(0 To 0) As LLC  'Merged library names
        ReDim ilEvtType(0 To 14) As Integer
        ReDim tlRegSdf(0 To 0) As SDF
        ReDim tlBBSdf(0 To 0) As SDF
        For ilIndex = LBound(ilEvtType) To UBound(ilEvtType) Step 1
            ilEvtType(ilIndex) = False
        Next ilIndex
        ilEvtType(2) = True 'avail
        ilEvtType(3) = True 'Open BB
        ilEvtType(5) = True 'Close BB
        If ilMode <> 2 Then
            ilEvtRet = gBuildEventDay(ilType, "C", ilVefCode, slDate, "12M", "12M", ilEvtType(), tlLLC())
        End If
        ilSdfRet = mGetRegAndBBSpots(hlSdf, ilVefCode, ilType, slDate, llChfCode, ilLineNo, False, tlRegSdf(), tlBBSdf())
        For ilSdf = 0 To UBound(tlRegSdf) - 1 Step 1
            If (tlRegSdf(ilSdf).sAffChg = "T") Or (tlRegSdf(ilSdf).sAffChg = "A") Or (tlRegSdf(ilSdf).sAffChg = "Y") Then
                Erase tlRegSdf
                Erase tlBBSdf
                Erase ilEvtType
                If hlInLcf = -1 Then
                    ilRet = btrClose(hmLcf)
                    btrDestroy hmLcf
                End If
                If hlInClf = -1 Then
                    ilRet = btrClose(hmClf)
                    btrDestroy hmClf
                End If
                gCreateBBSpots = True
                Exit Function
            End If
        Next ilSdf
        For ilSdf = 0 To UBound(tlBBSdf) - 1 Step 1
            If (tlBBSdf(ilSdf).sAffChg = "T") Or (tlBBSdf(ilSdf).sAffChg = "A") Or (tlBBSdf(ilSdf).sAffChg = "Y") Then
                Erase tlRegSdf
                Erase tlBBSdf
                Erase ilEvtType
                If hlInLcf = -1 Then
                    ilRet = btrClose(hmLcf)
                    btrDestroy hmLcf
                End If
                If hlInClf = -1 Then
                    ilRet = btrClose(hmClf)
                    btrDestroy hmClf
                End If
                gCreateBBSpots = True
                Exit Function
            End If
        Next ilSdf
        If ilMode <> 2 Then
            For ilSdf = 0 To UBound(tlRegSdf) - 1 Step 1
                tlSdf = tlRegSdf(ilSdf)
                'Get Line of spot
                tmClfSrchKey.lChfCode = tlSdf.lChfCode
                tmClfSrchKey.iLine = tlSdf.iLineNo
                tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tlSdf.lChfCode) And (tmClf.iLine = tlSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                    ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tlSdf.lChfCode) And (tmClf.iLine = tlSdf.iLineNo) Then
                    If (tmClf.iBBOpenLen > 0) Or (tmClf.iBBCloseLen > 0) Then
                        gUnpackTimeLong tlSdf.iTime(0), tlSdf.iTime(1), False, llSDFTime
                        'Find the Avail for spot, then get search for open
                        For ilLLC = 0 To UBound(tlLLC) - 1 Step 1
                            If tlLLC(ilLLC).sType = "2" Then    'Contract Avail
                                llAvailTime = gTimeToLong(tlLLC(ilLLC).sStartTime, False)
                                If llSDFTime = llAvailTime Then
                                    ilSPass = 1
                                    ilEPass = 2
                                    If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                                        For ilPass = 1 To 2 Step 1
                                            ilStart = ilLLC - 1
                                            ilEnd = 0
                                            ilStep = -1
                                            slBBType = "3"
                                            slSpotType = "O"
                                            ilLen = tmClf.iBBOpenLen
                                            If ilPass = 2 Then
                                                ilStart = ilLLC + 1
                                                ilEnd = UBound(tlLLC) - 1
                                                ilStep = 1
                                                slBBType = "5"
                                                slSpotType = "C"
                                                'ilLen = tmClf.iBBCloseLen
                                            End If
                                            ilFound = -1
                                            If ilLen > 0 Then
                                                For ilFind = ilStart To ilEnd Step ilStep
                                                    If tlLLC(ilFind).sType = slBBType Then    'Open Avail
                                                        If ilPass = 1 Then
                                                            llPass1AvailTime = gTimeToLong(tlLLC(ilFind).sStartTime, False)
                                                        Else
                                                            llAvailTime = gTimeToLong(tlLLC(ilFind).sStartTime, False)
                                                            If (llSDFTime - llPass1AvailTime) < (llAvailTime - llSDFTime) Then
                                                                ilEPass = 1
                                                            Else
                                                                ilSPass = 2
                                                            End If
                                                        End If
                                                        Exit For
                                                    End If
                                                Next ilFind
                                            End If
                                        Next ilPass
                                    End If
                                    For ilPass = ilSPass To ilEPass Step 1
                                        ilStart = ilLLC - 1
                                        ilEnd = 0
                                        ilStep = -1
                                        slBBType = "3"
                                        slSpotType = "O"
                                        ilLen = tmClf.iBBOpenLen
                                        If ilPass = 2 Then
                                            ilStart = ilLLC + 1
                                            ilEnd = UBound(tlLLC) - 1
                                            ilStep = 1
                                            slBBType = "5"
                                            slSpotType = "C"
                                            If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                                                ilLen = tmClf.iBBOpenLen
                                            Else
                                                ilLen = tmClf.iBBCloseLen
                                            End If
                                        End If
                                        ilFound = -1
                                        If ilLen > 0 Then
                                            For ilFind = ilStart To ilEnd Step ilStep
                                                If tlLLC(ilFind).sType = slBBType Then    'Open Avail
                                                    llAvailTime = gTimeToLong(tlLLC(ilFind).sStartTime, False)
                                                    For ilChkSdf = 0 To UBound(tlBBSdf) - 1 Step 1
                                                        If (tlBBSdf(ilChkSdf).lChfCode = tlSdf.lChfCode) And (tlBBSdf(ilChkSdf).iLen = ilLen) Then
                                                            gUnpackTimeLong tlBBSdf(ilChkSdf).iTime(0), tlBBSdf(ilChkSdf).iTime(1), False, llChkTime
                                                            If llChkTime = llAvailTime Then
                                                                ilFound = ilChkSdf
                                                                Exit For
                                                            End If
                                                        End If
                                                    Next ilChkSdf
                                                    If ilFound = -1 Then
                                                        If ilMode = 0 Then
                                                            'Make spot
                                                            tlCreateSdf = tlSdf
                                                            gPackTime tlLLC(ilFind).sStartTime, tlCreateSdf.iTime(0), tlCreateSdf.iTime(1)
                                                            tlCreateSdf.sSpotType = slSpotType
                                                            tlCreateSdf.sSchStatus = "S"
                                                            tlCreateSdf.sTracer = ""
                                                            tlCreateSdf.iLen = ilLen
                                                            tlCreateSdf.iRotNo = 0
                                                            tlCreateSdf.lCopyCode = 0
                                                            tlCreateSdf.sPtType = "0"
                                                            tlCreateSdf.lFsfCode = 0
                                                            tlCreateSdf.lSmfCode = 0
                                                            tlCreateSdf.sAffChg = ""
                                                            tlCreateSdf.iUrfCode = tgUrf(0).iCode
                                                            tlCreateSdf.sXCrossMidnight = "N"
                                                            If (tlLLC(ilFind).iAvailInfo And SSXMID) = SSXMID Then
                                                                tlCreateSdf.sXCrossMidnight = "Y"
                                                            End If
                                                            tlCreateSdf.sWasMG = "N"
                                                            tlCreateSdf.sFromWorkArea = "N"
                                                            tlCreateSdf.sUnused = ""
                                                            tlCreateSdf.lCode = 0
                                                            tlCreateSdf.iUrfCode = tgUrf(0).iCode
                                                            ilRet = btrInsert(hlSdf, tlCreateSdf, imSdfRecLen, INDEXKEY3)
                                                            If ilRet = BTRV_ERR_NONE Then
                                                                tlBBSdf(UBound(tlBBSdf)) = tlCreateSdf
                                                                ilFound = UBound(tlBBSdf)
                                                                tlBBSdf(ilFound).lCode = -tlBBSdf(ilFound).lCode
                                                                ReDim Preserve tlBBSdf(0 To UBound(tlBBSdf) + 1) As SDF
                                                            End If
                                                        End If
                                                    Else
                                                        If tlBBSdf(ilFound).lCode > 0 Then
                                                            tlBBSdf(ilFound).lCode = -tlBBSdf(ilFound).lCode
                                                        End If
                                                    End If
                                                    Exit For
                                                End If
                                            Next ilFind
                                        End If
                                    Next ilPass
                                End If
                            End If
                        Next ilLLC
                    End If
                End If
            Next ilSdf
        End If
        'Remove any BB not referenced
        For ilSdf = 0 To UBound(tlBBSdf) - 1 Step 1
            tlSdf = tlBBSdf(ilSdf)
            If tlSdf.lCode > 0 Then
                Do
                    tmSdfSrchKey3.lCode = tlSdf.lCode
                    ilRet = btrGetEqual(hlSdf, tlSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                    If ilRet <> BTRV_ERR_NONE Then
                        Exit Do
                    End If
                    ilRet = btrDelete(hlSdf)
                Loop While ilRet = BTRV_ERR_CONFLICT
            End If
        Next ilSdf
    Next ilLoop
    Erase tlRegSdf
    Erase tlBBSdf
    Erase ilEvtType
    If hlInLcf = -1 Then
        ilRet = btrClose(hmLcf)
        btrDestroy hmLcf
    End If
    If hlInClf = -1 Then
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
    End If
    gCreateBBSpots = True
End Function

Public Function gMakeBBAndAssignCopy(hlSdf As Integer, hlVlf As Integer, ilInVefCode As Integer, llInStartDate As Long, llInEndDate As Long) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************

    Dim ilVef As Integer
    Dim ilFound As Integer
    Dim ilTest As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim slStartDate As String
    Dim llDate As Long
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilZoneAdj As Integer
    Dim ilVpfIndex As Integer
    Dim ilZone As Integer
    Dim llNowDate As Long
    ReDim ilVefCode(0 To 0) As Integer
    Dim tlVef As VEF

    ReDim tlBBVefInfo(0 To 0) As BBVEFINFO
    If tgSpf.sUsingBBs <> "Y" Then
        gMakeBBAndAssignCopy = True
        Exit Function
    End If
    llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
    'Determine vehicles to create Billboard spots
    ilVef = gBinarySearchVef(ilInVefCode)
    If ilVef <> -1 Then
        tlVef = tgMVef(ilVef)
        llStartDate = llInStartDate
        llEndDate = llInEndDate
        slStartDate = Format$(llInStartDate, "m/d/yy")
        ilVpfIndex = gBinarySearchVpf(tlVef.iCode)
        ilZoneAdj = 0
        If ilVpfIndex <> -1 Then
            For ilZone = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
                If Trim$(tgVpf(ilVpfIndex).sGZone(ilZone)) <> "" Then
                    ilZoneAdj = 1
                    Exit For
                End If
            Next ilZone
        End If
        If tlVef.sType = "A" Then
            gBuildLinkArray hlVlf, tlVef, slStartDate, ilVefCode() 'Build igSVefCode so that gBuildODFSpotDay can use it
        ElseIf tlVef.sType = "L" Then
            For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If (tgMVef(ilVef).sType = "C") And (tgMVef(ilVef).iVefCode = tlVef.iCode) Then
                    ilVefCode(UBound(ilVefCode)) = tgMVef(ilVef).iCode
                    ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
                End If
            Next ilVef
        Else
            ReDim ilVefCode(0 To 1) As Integer
            ilVefCode(0) = tlVef.iCode
        End If
        For ilVef = LBound(ilVefCode) To UBound(ilVefCode) - 1 Step 1
            ilFound = False
            For ilTest = 0 To UBound(tlBBVefInfo) - 1 Step 1
                If (tlBBVefInfo(ilTest).iVefCode = ilVefCode(ilVef)) And (tlBBVefInfo(ilTest).lStartDate = llStartDate) And (tlBBVefInfo(ilTest).lEndDate = llEndDate) Then
                    ilFound = True
                    Exit For
                End If
            Next ilTest
            If Not ilFound Then
                'Add one to end date so that time zone is covered
                For llDate = llStartDate To llEndDate + ilZoneAdj Step 1
                    slDate = Format$(llDate, "m/d/yy")
                    If llDate > llNowDate Then
                        ilRet = gCreateBBSpots(hlSdf, ilVefCode(ilVef), slDate)
'                        If rbcLogType(0).Value Then
'                            ilRet = gAssignCopyToSpots(2, ilVefCode(ilVef), 0, slDate, slDate, "12M", "12M")
'                        Else
                            ilRet = gAssignCopyToSpots(-1, ilVefCode(ilVef), 1, slDate, slDate, "12M", "12M")
'                        End If
                    End If
                Next llDate
                tlBBVefInfo(UBound(tlBBVefInfo)).iVefCode = ilVefCode(ilVef)
                tlBBVefInfo(UBound(tlBBVefInfo)).lStartDate = llStartDate
                tlBBVefInfo(UBound(tlBBVefInfo)).lEndDate = llEndDate
                ReDim Preserve tlBBVefInfo(0 To UBound(tlBBVefInfo) + 1) As BBVEFINFO
            End If
        Next ilVef
    End If
    Erase ilVefCode
    Erase tlBBVefInfo
    gMakeBBAndAssignCopy = True
End Function

'***************************************************************************************
'*
'*      Procedure Name:gRemoveBBSpots
'*
'*             Created:5/17/93       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments: Remove Billboard spots by Vehicle/Date
'*
'*      4-24/14  Many reports call this routine to remove BB spots in the future.
'                Due to issues where BB spots are getting removed from an already generated
'                log (export), the BB spots will remain and let log and or export re-correct them
'                if necessary.  There seemed to be a timing issue that if both reports and
'                export/log were genned at the same time, the BB would be removed when it shouldnt.
'                Since many places call this routine, it will EXIT out without processing
'***************************************************************************************
Public Function gRemoveBBSpots(hlSdf, ilVefCode As Integer, ilGameNo As Integer, slInSDate As String, slEDate As String, llChfCode As Long, ilLineNo As Integer) As Integer
    '
    '  slEDate = "": TFN
    '  llChfCode < 0 or zero: Bypass Contract and Line matching
    '  ilLineNo < 0 or zero: Bypass Line number matching
    '  ilGameNo < 0 or zero: Bypass game number matching
    '
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilSdf As Integer
    Dim tlSdf As SDF
    Dim tlSdfSrchKey1 As SDFKEY1
    Dim ilSdfRecLen As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlCharTypeBuff As POPCHARTYPE   'Type field record
    Dim slSDate As String
    Dim ilVpfIndex As Integer

    gRemoveBBSpots = True           '4-24-14   see note above
    Exit Function
    
    If tgSpf.sUsingBBs <> "Y" Then
        gRemoveBBSpots = True
        Exit Function
    End If
    ReDim tlBBSdf(0 To 0) As SDF
    gRemoveBBSpots = True
    'Don't remove BB from days in past
    ilVpfIndex = gBinarySearchVpf(ilVefCode)
    If ilVpfIndex <> -1 Then
        gUnpackDate tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), slSDate
        If slSDate = "" Then
            slSDate = Format(Now, "m/d/yy")
        Else
            If gDateValue(slSDate) < gDateValue(Format(Now, "m/d/yy")) Then
                slSDate = Format(Now, "m/d/yy")
            End If
        End If
        slSDate = gIncOneDay(slSDate)
    Else
        slSDate = gIncOneDay(Format(Now, "m/d/yy"))
    End If
    If gDateValue(slInSDate) > gDateValue(slSDate) Then
        slSDate = slInSDate
    End If
    If slEDate <> "" Then
        If gDateValue(slEDate) < gDateValue(slSDate) Then
            Exit Function
        End If
    End If
    btrExtClear hlSdf   'Clear any previous extend operation
    ilExtLen = Len(tlSdf)  'Extract operation record size
    tlSdfSrchKey1.iVefCode = ilVefCode
    gPackDate slSDate, tlSdfSrchKey1.iDate(0), tlSdfSrchKey1.iDate(1)
    tlSdfSrchKey1.iTime(0) = 0
    tlSdfSrchKey1.iTime(1) = 0
    tlSdfSrchKey1.sSchStatus = ""   'slType
    ilSdfRecLen = Len(tlSdf)
    ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    If (tlSdf.iVefCode = ilVefCode) And (ilRet <> BTRV_ERR_END_OF_FILE) Then

        ' Prepare to execute an extended operation.
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDF", "") '"EG") 'Set extract limits (all records)

        ' We only the records for the passed in vehicle code.
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, ilVefCode, 2)


        ' And on the records where the date is equal to the passed in log date
        gPackDate slSDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        If slEDate <> "" Then
            gPackDate slEDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Sdf", "SdfDate")
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
        End If

        If llChfCode > 0 Then
            ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, llChfCode, 4)
            If ilLineNo > 0 Then
                ilOffSet = gFieldOffset("Sdf", "SdfLineNo")
                ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, ilLineNo, 2)
            End If
        Else
            llChfCode = 0
            ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, llChfCode, 4)
        End If
        
        tlCharTypeBuff.sType = "O"    'Extract all matching records
        ilOffSet = gFieldOffset("Sdf", "SdfSpotType")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_OR, tlCharTypeBuff, 1)

        tlCharTypeBuff.sType = "C"    'Extract all matching records
        ilOffSet = gFieldOffset("Sdf", "SdfSpotType")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlCharTypeBuff, 1)

        ilRet = btrExtAddField(hlSdf, 0, ilExtLen) 'Extract the whole record

        ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            ilExtLen = Len(tlSdf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If (tlSdf.iGameNo = ilGameNo) Or (ilGameNo <= 0) Then
                    If (tlSdf.sSpotType = "O") Or (tlSdf.sSpotType = "C") Then
                        tlBBSdf(UBound(tlBBSdf)) = tlSdf
                        ReDim Preserve tlBBSdf(0 To UBound(tlBBSdf) + 1) As SDF
                    End If
                End If
                ilExtLen = Len(tlSdf)
                ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
                Loop
                DoEvents
            Loop
        End If
        For ilSdf = 0 To UBound(tlBBSdf) - 1 Step 1
            tmSdfSrchKey3.lCode = tlBBSdf(ilSdf).lCode
            ilRet = btrGetEqual(hlSdf, tlSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hlSdf)
            End If
        Next ilSdf
    End If

End Function
'***************************************************************************************
'*
'*      Procedure Name:mGetSpots
'*
'*             Created:5/17/93       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments:
'*
'***************************************************************************************
Private Function mGetRegAndBBSpots(hlSdf, ilVefCode As Integer, ilGameNo As Integer, slDate As String, llInChfCode As Long, ilLineNo As Integer, blIncludeFill As Boolean, tlRegSdf() As SDF, tlBBSdf() As SDF) As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim llChfCode As Long
    Dim tlSdf As SDF
    Dim tlSdfSrchKey1 As SDFKEY1
    Dim ilSdfRecLen As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    ReDim tlRegSdf(0 To 0) As SDF
    ReDim tlBBSdf(0 To 0) As SDF
    mGetRegAndBBSpots = True
    llChfCode = llInChfCode
    btrExtClear hlSdf   'Clear any previous extend operation
    ilExtLen = Len(tlSdf)  'Extract operation record size
    tlSdfSrchKey1.iVefCode = ilVefCode
    gPackDate slDate, tlSdfSrchKey1.iDate(0), tlSdfSrchKey1.iDate(1)
    tlSdfSrchKey1.iTime(0) = 0
    tlSdfSrchKey1.iTime(1) = 0
    tlSdfSrchKey1.sSchStatus = ""   'slType
    ilSdfRecLen = Len(tlSdf)
    ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    If (tlSdf.iVefCode = ilVefCode) And (ilRet <> BTRV_ERR_END_OF_FILE) Then

        ' Prepare to execute an extended operation.
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDF", "") '"EG") 'Set extract limits (all records)

        ' We only the records for the passed in vehicle code.
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, ilVefCode, 2)


        If llChfCode > 0 Then
            ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, llChfCode, 4)
            If ilLineNo > 0 Then
                ilOffSet = gFieldOffset("Sdf", "SdfLineNo")
                ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, ilLineNo, 2)
            End If
        Else
            llChfCode = 0
            ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
            ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, llChfCode, 4)
        End If

        ' And on the records where the date is equal to the passed in log date
        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hlSdf, 0, ilExtLen) 'Extract the whole record

        ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            ilExtLen = Len(tlSdf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If tlSdf.iGameNo = ilGameNo Then
                    If (tlSdf.sSpotType = "O") Or (tlSdf.sSpotType = "C") Then
                        tlBBSdf(UBound(tlBBSdf)) = tlSdf
                        ReDim Preserve tlBBSdf(0 To UBound(tlBBSdf) + 1) As SDF
                    Else
                        If (tlSdf.sSpotType <> "X") Or ((blIncludeFill) And (tlSdf.sSpotType = "X")) Then
                            If (tlSdf.sSchStatus = "S") Or (tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O") Then
                                tlRegSdf(UBound(tlRegSdf)) = tlSdf
                                ReDim Preserve tlRegSdf(0 To UBound(tlRegSdf) + 1) As SDF
                            End If
                        End If
                    End If
                End If
                ilExtLen = Len(tlSdf)
                ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
                Loop
                DoEvents
            Loop
        End If
    End If

End Function

'*******************************************************
'*                                                     *
'*      Procedure Name:mGetOtherLineSpots              *
'*                                                     *
'*             Created:4/21/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Determine if Lines are buying   *
'*                     the same days/times.  If so,    *
'*                     then include the spots in the   *
'*                     spot separation computation     *
'*                                                     *
'*           6/28/05-Changed from just testing RdfCode *
'*                   to checking airing days and times *
'*                                                     *
'*******************************************************
Private Function mGetOtherLineSpots(llDate As Long, tlClf As CLF, tlCff As CFF, ilVpfIndex As Integer, tlRdf As RDF) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  hlChf                                                                                 *
'******************************************************************************************

    Dim hlClf As Integer
    Dim hlCff As Integer
    Dim ilRet As Integer
    Dim ilClf As Integer
    Dim ilCff As Integer
    Dim ilSpots As Integer
    Dim ilDay As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llCFFStartTime As Long
    Dim llCFFEndTime As Long
    Dim llTestStartTime As Long
    Dim llTestEndTime As Long
    Dim ilLoop As Integer
    Dim ilTest As Integer
    Dim ilRdf As Integer
    Dim ilOverlapDays As Integer
    Dim ilOverlapTimes As Integer
    Dim ilClfUpper As Integer
    Dim tlXMidClf As CLF
    Dim tlXMidRdf As RDF

    mGetOtherLineSpots = 0
    If tlClf.iAdvtSepFlag = 0 Then
        Exit Function
    End If
    If tlClf.iAdvtSepFlag > 0 Then
        mGetOtherLineSpots = tlClf.iAdvtSepFlag
        Exit Function
    End If
    'hlChf = CBtrvTable(ONEHANDLE)        'Create SDF object handle
    'ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    'If ilRet <> BTRV_ERR_NONE Then
    '    btrDestroy hlChf
    '    Exit Function
    'End If
    hlClf = CBtrvTable(ONEHANDLE)        'Create SDF object handle
    ilRet = btrOpen(hlClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        'btrDestroy hlChf
        btrDestroy hlClf
        Exit Function
    End If

    ReDim tmChkClf(0 To 0) As CLF
    ilClfUpper = 0
    imClfRecLen = Len(tmChkClf(0))
    tmClfSrchKey1.lChfCode = tlClf.lChfCode
    tmClfSrchKey1.iVefCode = tlClf.iVefCode
    ilRet = btrGetEqual(hlClf, tmChkClf(ilClfUpper), imClfRecLen, tmClfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
    Do While (ilRet = BTRV_ERR_NONE) And (tmChkClf(ilClfUpper).lChfCode = tlClf.lChfCode) And (tmChkClf(ilClfUpper).iVefCode = tlClf.iVefCode)
        If tmChkClf(ilClfUpper).iLine <> tlClf.iLine Then
            ilClfUpper = ilClfUpper + 1
            ReDim Preserve tmChkClf(0 To ilClfUpper) As CLF
        End If
        ilRet = btrGetNext(hlClf, tmChkClf(ilClfUpper), imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If ilClfUpper = 0 Then
        btrDestroy hlClf
        Exit Function
    End If

    hlCff = CBtrvTable(ONEHANDLE)        'Create SDF object handle
    ilRet = btrOpen(hlCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        'btrDestroy hlChf
        btrDestroy hlClf
        btrDestroy hlCff
        Exit Function
    End If
    ilDay = gWeekDayLong(llDate)
    ilSpots = 0
    'ilRet = gObtainCntr(hlChf, hlClf, hlCff, tlClf.lChfCode, False, tmChkChf, tmChkClf(), tmChkCFF())
    For ilClf = LBound(tmChkClf) To UBound(tmChkClf) - 1 Step 1
        If tlClf.iLine <> tmChkClf(ilClf).iLine Then
            If (tmChkClf(ilClf).sType <> "O") And (tmChkClf(ilClf).sType <> "A") And (tmChkClf(ilClf).sType <> "E") Then
                If tlClf.iVefCode = tmChkClf(ilClf).iVefCode Then
                    'If tlClf.iRdfcode = tmChkClf(ilClf).iRdfcode Then
                        ilRet = gGetFlightsFromLine(hlCff, tmChkClf(ilClf), tmChkCFF())
                        For ilCff = LBound(tmChkCFF) To UBound(tmChkCFF) - 1 Step 1
                            If tmChkCFF(ilCff).iClfLine = tmChkClf(ilClf).iLine Then
                                gUnpackDateLong tmChkCFF(ilCff).iStartDate(0), tmChkCFF(ilCff).iStartDate(1), llStartDate
                                gUnpackDateLong tmChkCFF(ilCff).iEndDate(0), tmChkCFF(ilCff).iEndDate(1), llEndDate
                                If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                                    'Check if Days Overlap
                                    ilOverlapDays = False
                                    If tlCff.sDyWk <> "D" Then
                                        If (tmChkCFF(ilCff).sDyWk <> "D") Then  'Weekly
                                            For ilTest = 0 To 6 Step 1
                                                If (tlCff.iDay(ilTest) <> 0) And (tmChkCFF(ilCff).iDay(ilTest) <> 0) Then
                                                    ilOverlapDays = True
                                                    Exit For
                                                End If
                                            Next ilTest
                                        Else
                                            If (tlCff.iDay(ilDay) <> 0) And (tmChkCFF(ilCff).iDay(ilDay) > 0) Then
                                                ilOverlapDays = True
                                            End If
                                        End If
                                    Else
                                        If (tmChkCFF(ilCff).sDyWk <> "D") Then  'Weekly
                                            If (tlCff.iDay(ilDay) > 0) And (tmChkCFF(ilCff).iDay(ilDay) <> 0) Then
                                                ilOverlapDays = True
                                            End If
                                        Else
                                            If (tlCff.iDay(ilDay) > 0) And (tmChkCFF(ilCff).iDay(ilDay) > 0) Then
                                                ilOverlapDays = True
                                            End If
                                        End If
                                    End If
                                    'Check if Times overlap
                                    ilOverlapTimes = False
                                    If ilOverlapDays Then
                                        If (tlRdf.iLtfCode(0) <> 0) Or (tlRdf.iLtfCode(1) <> 0) Or (tlRdf.iLtfCode(2) <> 0) Then
                                            If tlClf.iRdfCode = tmChkClf(ilClf).iRdfCode Then
                                                ilOverlapTimes = True
                                            End If
                                        Else
                                            ilRdf = gBinarySearchRdf(tmChkClf(ilClf).iRdfCode)
                                            If ilRdf <> -1 Then
                                                gXMidClfRdfToRdf "", tmChkClf(ilClf), tgMRdf(ilRdf), tlXMidClf, tlXMidRdf
                                                If ((tlClf.iStartTime(0) = 1) And (tlClf.iStartTime(1) = 0)) Then
                                                    If (tlXMidRdf.iLtfCode(0) <> 0) Or (tlXMidRdf.iLtfCode(1) <> 0) Or (tlXMidRdf.iLtfCode(2) <> 0) Then
                                                        If tlClf.iRdfCode = tlXMidClf.iRdfCode Then
                                                            ilOverlapTimes = True
                                                        End If
                                                    Else
                                                        If ((tlXMidClf.iStartTime(0) = 1) And (tlXMidClf.iStartTime(1) = 0)) Then
                                                            For ilLoop = LBound(tlRdf.iStartTime, 2) To UBound(tlRdf.iStartTime, 2) Step 1
                                                                If (tlXMidRdf.iStartTime(0, ilLoop) <> 1) Or (tlXMidRdf.iStartTime(1, ilLoop) <> 0) Then
                                                                    gUnpackTimeLong tlRdf.iStartTime(0, ilLoop), tlRdf.iStartTime(1, ilLoop), False, llCFFStartTime
                                                                    gUnpackTimeLong tlRdf.iEndTime(0, ilLoop), tlRdf.iEndTime(1, ilLoop), True, llCFFEndTime
                                                                    For ilTest = LBound(tlXMidRdf.iStartTime, 2) To UBound(tlXMidRdf.iStartTime, 2) Step 1
                                                                        If (tlXMidRdf.iStartTime(0, ilTest) <> 1) Or (tlXMidRdf.iStartTime(1, ilTest) <> 0) Then
                                                                            gUnpackTimeLong tlXMidRdf.iStartTime(0, ilTest), tlXMidRdf.iStartTime(1, ilTest), False, llTestStartTime
                                                                            gUnpackTimeLong tlXMidRdf.iEndTime(0, ilTest), tlXMidRdf.iEndTime(1, ilTest), True, llTestEndTime
                                                                            If (llTestStartTime < llCFFEndTime) And (llTestEndTime > llCFFStartTime) Then
                                                                                ilOverlapTimes = True
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    Next ilTest
                                                                    If ilOverlapTimes Then
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            Next ilLoop
                                                        Else
                                                            gUnpackTimeLong tlXMidClf.iStartTime(0), tlXMidClf.iStartTime(1), False, llTestStartTime
                                                            gUnpackTimeLong tlXMidClf.iEndTime(0), tlXMidClf.iEndTime(1), True, llTestEndTime
                                                            For ilLoop = LBound(tlRdf.iStartTime, 2) To UBound(tlRdf.iStartTime, 2) Step 1
                                                                If (tlXMidRdf.iStartTime(0, ilLoop) <> 1) Or (tlXMidRdf.iStartTime(1, ilLoop) <> 0) Then
                                                                    gUnpackTimeLong tlRdf.iStartTime(0, ilLoop), tlRdf.iStartTime(1, ilLoop), False, llCFFStartTime
                                                                    gUnpackTimeLong tlRdf.iEndTime(0, ilLoop), tlRdf.iEndTime(1, ilLoop), True, llCFFEndTime
                                                                    If (llTestStartTime < llCFFEndTime) And (llTestEndTime > llCFFStartTime) Then
                                                                        ilOverlapTimes = True
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            Next ilLoop
                                                        End If
                                                    End If
                                                Else
                                                    gUnpackTimeLong tlClf.iStartTime(0), tlClf.iStartTime(1), False, llCFFStartTime
                                                    gUnpackTimeLong tlClf.iEndTime(0), tlClf.iEndTime(1), True, llCFFEndTime
                                                    If (tlXMidRdf.iLtfCode(0) <> 0) Or (tlXMidRdf.iLtfCode(1) <> 0) Or (tlXMidRdf.iLtfCode(2) <> 0) Then
                                                        If tlClf.iRdfCode = tlXMidClf.iRdfCode Then
                                                            ilOverlapTimes = True
                                                        End If
                                                    Else
                                                        If ((tlXMidClf.iStartTime(0) = 1) And (tlXMidClf.iStartTime(1) = 0)) Then
                                                            For ilTest = LBound(tlXMidRdf.iStartTime, 2) To UBound(tlXMidRdf.iStartTime, 2) Step 1
                                                                If (tlXMidRdf.iStartTime(0, ilTest) <> 1) Or (tlXMidRdf.iStartTime(1, ilTest) <> 0) Then
                                                                    gUnpackTimeLong tlXMidRdf.iStartTime(0, ilTest), tlXMidRdf.iStartTime(1, ilTest), False, llTestStartTime
                                                                    gUnpackTimeLong tlXMidRdf.iEndTime(0, ilTest), tlXMidRdf.iEndTime(1, ilTest), True, llTestEndTime
                                                                    If (llTestStartTime < llCFFEndTime) And (llTestEndTime > llCFFStartTime) Then
                                                                        ilOverlapTimes = True
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            Next ilTest
                                                        Else
                                                            gUnpackTimeLong tlXMidClf.iStartTime(0), tlXMidClf.iStartTime(1), False, llTestStartTime
                                                            gUnpackTimeLong tlXMidClf.iEndTime(0), tlXMidClf.iEndTime(1), True, llTestEndTime
                                                            If (llTestStartTime < llCFFEndTime) And (llTestEndTime > llCFFStartTime) Then
                                                                ilOverlapTimes = True
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    If ilOverlapTimes And ilOverlapDays Then
                                        If (tmChkCFF(ilCff).sDyWk <> "D") Then  'Weekly
                                            ilSpots = ilSpots + tmChkCFF(ilCff).iSpotsWk
                                        Else
                                            ilSpots = ilSpots + tmChkCFF(ilCff).iDay(ilDay)
                                        End If
                                    End If
                                    Exit For
                                End If
                            End If
                        Next ilCff
                    'End If
                End If
            End If
        End If
    Next ilClf
    'ilRet = btrClose(hlChf)
    'btrDestroy hlChf
    ilRet = btrClose(hlClf)
    btrDestroy hlClf
    ilRet = btrClose(hlCff)
    btrDestroy hlCff
    mGetOtherLineSpots = ilSpots

End Function

'***************************************************************************************
'*
'*      Procedure Name:mGetSpots
'*
'*             Created:5/17/93       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments:
'*
'***************************************************************************************
Public Function gGetBBSpots(hlSdf, ilVefCode As Integer, ilGameNo As Integer, slDate As String, tlBBSdf() As SDF) As Integer
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim llChfCode As Long
    Dim tlSdf As SDF
    Dim tlSdfSrchKey1 As SDFKEY1
    Dim ilSdfRecLen As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    ReDim tlBBSdf(0 To 0) As SDF
    gGetBBSpots = True
    If tgSpf.sUsingBBs <> "Y" Then
        Exit Function
    End If
    btrExtClear hlSdf   'Clear any previous extend operation
    ilExtLen = Len(tlSdf)  'Extract operation record size
    tlSdfSrchKey1.iVefCode = ilVefCode
    gPackDate slDate, tlSdfSrchKey1.iDate(0), tlSdfSrchKey1.iDate(1)
    tlSdfSrchKey1.iTime(0) = 0
    tlSdfSrchKey1.iTime(1) = 0
    tlSdfSrchKey1.sSchStatus = ""   'slType
    ilSdfRecLen = Len(tlSdf)
    ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
    If (tlSdf.iVefCode = ilVefCode) And (ilRet <> BTRV_ERR_END_OF_FILE) Then

        ' Prepare to execute an extended operation.
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDF", "") '"EG") 'Set extract limits (all records)

        ' We only the records for the passed in vehicle code.
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, ilVefCode, 2)


        ' And only records where the ChfCode = 0
        llChfCode = 0
        ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, llChfCode, 4)

        ' And on the records where the date is equal to the passed in log date
        gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
        ilOffSet = gFieldOffset("Sdf", "SdfDate")
        ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hlSdf, 0, ilExtLen) 'Extract the whole record

        ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            ilExtLen = Len(tlSdf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                If tlSdf.iGameNo = ilGameNo Then
                    If (tlSdf.sSpotType = "O") Or (tlSdf.sSpotType = "C") Then
                        tlBBSdf(UBound(tlBBSdf)) = tlSdf
                        ReDim Preserve tlBBSdf(0 To UBound(tlBBSdf) + 1) As SDF
                    End If
                End If
                ilExtLen = Len(tlSdf)
                ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSdf, tlSdf, ilExtLen, llRecPos)
                Loop
                DoEvents
            Loop
        End If
    End If

End Function

'***************************************************************************************
'*
'*      Procedure Name:gXMidClfRdfToRdf
'*
'*             Created:5/17/93       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments:  Translate Line time overrides and/or Daypart
'*                       times that cross mid night into rdf
'*
'***************************************************************************************
Public Sub gXMidClfRdfToRdf(slWkType As String, tlInClf As CLF, tlInRdf As RDF, tlOutClf As CLF, tlOutRdf As RDF)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilFirstDay                    ilLastDay                                               *
'******************************************************************************************

'   slWkType(I)- Week Type (F=First; L=Last, B=Both)

    Dim ilLoop As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilVpfIndex As Integer
    Dim ilRdfIndex As Integer
    Dim ilDayIndex As Integer
    Dim ilDay As Integer
    Dim slWkDays(0 To 6) As String

    ilVpfIndex = gBinarySearchVpf(tlInClf.iVefCode)

    tlOutClf = tlInClf
    tlOutRdf = tlInRdf
    If (tlInRdf.iLtfCode(0) = 0) And (tlInRdf.iLtfCode(1) = 0) And (tlInRdf.iLtfCode(2) = 0) Then
        If ((tlInClf.iStartTime(0) = 1) And (tlInClf.iStartTime(1) = 0)) Or (tgVpf(ilVpfIndex).sGMedium = "S") Then
            gXMidRdfToRdf slWkType, tlInRdf, tlOutRdf
        Else
            gUnpackTimeLong tlInClf.iStartTime(0), tlInClf.iStartTime(1), False, llStartTime
            gUnpackTimeLong tlInClf.iEndTime(0), tlInClf.iEndTime(1), True, llEndTime
            tlOutClf.iStartTime(0) = 1
            tlOutClf.iStartTime(1) = 0
            For ilDay = 0 To 6 Step 1
                slWkDays(ilDay) = "N"
            Next ilDay
            For ilLoop = LBound(tlInRdf.iStartTime, 2) To UBound(tlInRdf.iStartTime, 2) Step 1
                If (tlInRdf.iStartTime(0, ilLoop) <> 1) Or (tlInRdf.iStartTime(1, ilLoop) <> 0) Then
                    ilDayIndex = 0
                    For ilDay = 1 To 7 Step 1
                        If tlInRdf.sWkDays(ilLoop, ilDay - 1) = "Y" Then
                            slWkDays(ilDayIndex) = "Y"
                        End If
                        ilDayIndex = ilDayIndex + 1
                    Next ilDay
                End If
                tlOutRdf.iStartTime(0, ilLoop) = 1
                tlOutRdf.iStartTime(1, ilLoop) = 0
                For ilDay = 1 To 7 Step 1
                    tlOutRdf.sWkDays(ilLoop, ilDay - 1) = "N"
                Next ilDay
            Next ilLoop
            ilRdfIndex = UBound(tlOutRdf.iStartTime, 2)
            If (llStartTime > llEndTime) Then
                mBuildXMidSegment slWkType, slWkDays(), llStartTime, llEndTime, ilRdfIndex, tlOutRdf
            Else
                gPackTimeLong llStartTime, tlOutRdf.iStartTime(0, ilRdfIndex), tlOutRdf.iStartTime(1, ilRdfIndex)
                gPackTimeLong llEndTime, tlOutRdf.iEndTime(0, ilRdfIndex), tlOutRdf.iEndTime(1, ilRdfIndex)
                ilDayIndex = 0
                For ilDay = 1 To 7 Step 1
                    tlOutRdf.sWkDays(ilRdfIndex, ilDay - 1) = slWkDays(ilDayIndex)
                    ilDayIndex = ilDayIndex + 1
                Next ilDay
                tlOutRdf.iSpotPct(ilRdfIndex) = 100
            End If
        End If
    End If
End Sub


'***************************************************************************************
'*
'*      Procedure Name:gXMidToRdf
'*
'*             Created:5/17/93       By:D. LeVine
'*            Modified:              By:
'*
'*            Comments:  Translate Daypart
'*                       times that cross mid night into rdf
'*
'***************************************************************************************
Public Sub gXMidRdfToRdf(slWkType As String, tlInRdf As RDF, tlOutRdf As RDF)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilFirstDay                    ilLastDay                                               *
'******************************************************************************************

    Dim ilLoop As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilAnyXMid As Integer
    Dim ilRdfIndex As Integer
    Dim ilDayIndex As Integer
    Dim ilDay As Integer
    Dim slWkDays(0 To 6) As String

    tlOutRdf = tlInRdf
    If (tlInRdf.iLtfCode(0) = 0) And (tlInRdf.iLtfCode(1) = 0) And (tlInRdf.iLtfCode(2) = 0) Then

        ilAnyXMid = False
        For ilLoop = LBound(tlInRdf.iStartTime, 2) To UBound(tlInRdf.iStartTime, 2) Step 1
            If (tlInRdf.iStartTime(0, ilLoop) <> 1) Or (tlInRdf.iStartTime(1, ilLoop) <> 0) Then
                gUnpackTimeLong tlInRdf.iStartTime(0, ilLoop), tlInRdf.iStartTime(1, ilLoop), False, llStartTime
                gUnpackTimeLong tlInRdf.iEndTime(0, ilLoop), tlInRdf.iEndTime(1, ilLoop), True, llEndTime
                If llEndTime < llStartTime Then
                    ilAnyXMid = True
                End If
            End If
        Next ilLoop
        If ilAnyXMid Then
            ilRdfIndex = UBound(tlInRdf.iStartTime, 2)
            For ilLoop = UBound(tlInRdf.iStartTime, 2) To LBound(tlInRdf.iStartTime, 2) Step -1
                If (tlInRdf.iStartTime(0, ilLoop) <> 1) Or (tlInRdf.iStartTime(1, ilLoop) <> 0) Then
                    gUnpackTimeLong tlInRdf.iStartTime(0, ilLoop), tlInRdf.iStartTime(1, ilLoop), False, llStartTime
                    gUnpackTimeLong tlInRdf.iEndTime(0, ilLoop), tlInRdf.iEndTime(1, ilLoop), True, llEndTime
                    If llStartTime <= llEndTime Then
                        tlOutRdf.iStartTime(0, ilRdfIndex) = tlInRdf.iStartTime(0, ilLoop)
                        tlOutRdf.iStartTime(1, ilRdfIndex) = tlInRdf.iStartTime(1, ilLoop)
                        tlOutRdf.iEndTime(0, ilRdfIndex) = tlInRdf.iEndTime(0, ilLoop)
                        tlOutRdf.iEndTime(1, ilRdfIndex) = tlInRdf.iEndTime(1, ilLoop)
                        For ilDay = 1 To 7 Step 1
                            tlOutRdf.sWkDays(ilRdfIndex, ilDay - 1) = tlInRdf.sWkDays(ilLoop, ilDay - 1)
                        Next ilDay
                        tlOutRdf.iSpotPct(ilRdfIndex) = tlOutRdf.iSpotPct(ilLoop)
                    Else
                        ilDayIndex = 0
                        For ilDay = 1 To 7 Step 1
                            slWkDays(ilDayIndex) = tlInRdf.sWkDays(ilLoop, ilDay - 1)
                            ilDayIndex = ilDayIndex + 1
                        Next ilDay
                        mBuildXMidSegment slWkType, slWkDays(), llStartTime, llEndTime, ilRdfIndex, tlOutRdf
                    End If
                Else
                    tlOutRdf.iStartTime(0, ilRdfIndex) = 1
                    tlOutRdf.iStartTime(1, ilRdfIndex) = 0
                    For ilDay = 1 To 7 Step 1
                        tlOutRdf.sWkDays(ilRdfIndex, ilDay - 1) = "N"
                    Next ilDay
                    tlOutRdf.iSpotPct(ilRdfIndex) = 0
                End If
                ilRdfIndex = ilRdfIndex - 1
                If ilRdfIndex < LBound(tlInRdf.iStartTime, 2) Then
                    Exit For
                End If
            Next ilLoop
        End If
    End If
End Sub

Private Sub mBuildXMidSegment(slWkType As String, slWkDays() As String, llStartTime As Long, llEndTime As Long, ilRdfIndex As Integer, tlRdf As RDF)
    Dim ilDay As Integer
    Dim ilFirstDay As Integer
    Dim ilLastDay As Integer
    Dim ilDayIndex As Integer

    gPackTimeLong llStartTime, tlRdf.iStartTime(0, ilRdfIndex), tlRdf.iStartTime(1, ilRdfIndex)
    gPackTime "12M", tlRdf.iEndTime(0, ilRdfIndex), tlRdf.iEndTime(1, ilRdfIndex)
    ilDayIndex = 0
    For ilDay = 1 To 7 Step 1
        tlRdf.sWkDays(ilRdfIndex, ilDay - 1) = slWkDays(ilDayIndex)
        ilDayIndex = ilDayIndex + 1
    Next ilDay
    tlRdf.iSpotPct(ilRdfIndex) = 100
    ilRdfIndex = ilRdfIndex - 1
    If ilRdfIndex < LBound(tlRdf.iStartTime, 2) Then
        Exit Sub
    End If
    gPackTime "12M", tlRdf.iStartTime(0, ilRdfIndex), tlRdf.iStartTime(1, ilRdfIndex)
    gPackTimeLong llEndTime, tlRdf.iEndTime(0, ilRdfIndex), tlRdf.iEndTime(1, ilRdfIndex)
    ilFirstDay = 0
    ilLastDay = 0
    ilDayIndex = 0
    For ilDay = 1 To 7 Step 1
        'For now, don't rotate days in the 12m-1a segment
        'If ilDay <> 7 Then
        '    tlRdf.sWkDays(ilRdfIndex, ilDay+1) = slWkDays(ilDay))
        'Else
        '    tlRdf.sWkDays(ilRdfIndex, 1) = slWkDays(ilDay)
        'End If
        tlRdf.sWkDays(ilRdfIndex, ilDay - 1) = slWkDays(ilDayIndex)
        If slWkDays(ilDayIndex) = "Y" Then
            If ilFirstDay = 0 Then
                ilFirstDay = ilDay
                ilLastDay = ilDay
            Else
                ilLastDay = ilDay
            End If
        End If
        ilDayIndex = ilDayIndex + 1
    Next ilDay
    If ilFirstDay > 0 Then
        'If rotating days, only eliminate monday (ilLastDay = 7), use 1
        If (slWkType = "F") Or (slWkType = "B") Then
            tlRdf.sWkDays(ilRdfIndex, ilFirstDay - 1) = "N"
        End If
        'Not rotating days so last is not required
        'Eliminate day if not sunday(ilLastDay <> 7), use ilLastDay + 1
        'If (slWkType = "L") Or (slWkType = "B") Then
        '    tlRdf.sWkDays(ilRdfIndex, ilLastDay) = "N"
        'End If
    End If
    tlRdf.iSpotPct(ilRdfIndex) = (100 * llEndTime) / (86400 - llStartTime + llEndTime)
    tlRdf.iSpotPct(ilRdfIndex + 1) = 100 - tlRdf.iSpotPct(ilRdfIndex)
End Sub
Public Function gConvertErrorCode(ilRet) As Integer
    gConvertErrorCode = ilRet
    If ilRet >= 30000 Then
        gConvertErrorCode = csiHandleValue(0, 7)
    End If
End Function


Public Function gGetPrgName(llDate As Long, llTime As Long, tlPaf() As PAF) As String
    Dim ilLoop As Long
    Dim ilDay As Integer
    Dim llPafDate As Long
    Dim llPafTime As Long
    Dim ilDayOk As Integer
    
    gGetPrgName = ""
    For ilLoop = 0 To UBound(tlPaf) - 1 Step 1
        ilDayOk = False
        ilDay = gWeekDayLong(llDate)
        Select Case ilDay
            Case 0  'Monday
                If tlPaf(ilLoop).sMo = "Y" Then
                    ilDayOk = True
                End If
            Case 1  'Tuesday
                If tlPaf(ilLoop).sTu = "Y" Then
                    ilDayOk = True
                End If
            Case 2  'Wednesady
                If tlPaf(ilLoop).sWe = "Y" Then
                    ilDayOk = True
                End If
            Case 3  'Thursday
                If tlPaf(ilLoop).sTh = "Y" Then
                    ilDayOk = True
                End If
            Case 4  'Friday
                If tlPaf(ilLoop).sFr = "Y" Then
                    ilDayOk = True
                End If
            Case 5  'Saturday
                If tlPaf(ilLoop).sSa = "Y" Then
                    ilDayOk = True
                End If
            Case 6  'Sunday
                If tlPaf(ilLoop).sSu = "Y" Then
                    ilDayOk = True
                End If
        End Select
        If ilDayOk Then
            gUnpackDateLong tlPaf(ilLoop).iStartDate(0), tlPaf(ilLoop).iStartDate(1), llPafDate
            If (llDate >= llPafDate) Then
                gUnpackDateLong tlPaf(ilLoop).iEndDate(0), tlPaf(ilLoop).iEndDate(1), llPafDate
                If llDate <= llPafDate Then
                    gUnpackTimeLong tlPaf(ilLoop).iStartTime(0), tlPaf(ilLoop).iStartTime(1), False, llPafTime
                    If llTime >= llPafTime Then
                        gUnpackTimeLong tlPaf(ilLoop).iEndTime(0), tlPaf(ilLoop).iEndTime(1), True, llPafTime
                        If llTime <= llPafTime Then
                            gGetPrgName = Trim$(tlPaf(ilLoop).sName)
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next ilLoop
End Function

Public Sub gMakeLogAlert(tlSdf As SDF, slSubType As String, hlInGsf As Integer)
    Dim ilVefCode As Integer
    Dim llLogDate As Long
    Dim ilVef As Integer
    Dim ilVpf As Integer
    Dim ilLogVef As Integer
    Dim ilCombineVef As Integer
    Dim ilRet As Integer
    Dim hlVlf As Integer
    Dim ilAir As Integer
    Dim tlVef As VEF
    Dim hlGsf As Integer
    Dim ilGsfRecLen As Integer
    Dim tlGsf As GSF
    Dim tlGsfSrchKey3 As GSFKEY3
    Dim llGsfDate As Long
    ReDim ilAirVefCode(0 To 0) As Integer
    
    ilVefCode = tlSdf.iVefCode
    gUnpackDateLong tlSdf.iDate(0), tlSdf.iDate(1), llLogDate
    ilVef = gBinarySearchVef(ilVefCode)
    If ilVef <> -1 Then
        If tgMVef(ilVef).sType = "C" Then
            If tgMVef(ilVef).iVefCode > 0 Then  'Log Vehicle
                ilLogVef = gBinarySearchVef(tgMVef(ilVef).iVefCode)
                If ilLogVef <> -1 Then
                    mAddLogAlert tgMVef(ilLogVef).iCode, llLogDate, slSubType
                End If
            ElseIf tgMVef(ilVef).iCombineVefCode > 0 Then   'Multi Vehicle Log
                mAddLogAlert ilVefCode, llLogDate, slSubType
                ilCombineVef = gBinarySearchVef(tgMVef(ilVef).iCombineVefCode)
                If ilCombineVef <> -1 Then
                    mAddLogAlert tgMVef(ilCombineVef).iCode, llLogDate, slSubType
                End If
            Else
                mAddLogAlert ilVefCode, llLogDate, slSubType
            End If
        ElseIf tgMVef(ilVef).sType = "S" Then
            'The conversion is performed with gAlertVehicleReplace when Alert selected on Log screen
            'therefore leave the alert with the Selling vehicle.
            'hlVlf = CBtrvTable(ONEHANDLE)        'Create SDF object handle
            'On Error GoTo 0
            'ilRet = btrOpen(hlVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            'If ilRet = BTRV_ERR_NONE Then
            '    tlVef = tgMVef(ilVef)
            '    gBuildLinkArray hlVlf, tlVef, Format$(llLogDate, "m/d/yy"), ilAirVefCode()
            '    For ilAir = 0 To UBound(ilAirVefCode) - 1 Step 1
            '        mAddLogAlert ilAirVefCode(ilAir), llLogDate, slSubType
            '    Next ilAir
            'End If
            'ilRet = btrClose(hlVlf)
            'btrDestroy hlVlf
            mAddLogAlert ilVefCode, llLogDate, slSubType
        ElseIf tgMVef(ilVef).sType = "G" Then
            mAddLogAlert ilVefCode, llLogDate, slSubType
            If hlInGsf <= 0 Then
                hlGsf = CBtrvTable(ONEHANDLE)        'Create SDF object handle
                On Error GoTo 0
                ilRet = btrOpen(hlGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
            Else
                hlGsf = hlInGsf
                ilRet = BTRV_ERR_NONE
            End If
            If ilRet = BTRV_ERR_NONE Then
                If tgMVef(ilVef).iVefCode > 0 Then  'Log Vehicle
                    ilLogVef = gBinarySearchVef(tgMVef(ilVef).iVefCode)
                    If ilLogVef <> -1 Then
                        mAddLogAlert tgMVef(ilLogVef).iCode, llLogDate, slSubType
                    End If
                Else
                    ilGsfRecLen = Len(tlGsf)
                    'All Pre-empt vehicle
                    tlGsfSrchKey3.iVefCode = tlSdf.iVefCode
                    tlGsfSrchKey3.iGameNo = tlSdf.iGameNo
                    ilRet = btrGetEqual(hlGsf, tlGsf, ilGsfRecLen, tlGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tlGsf.iVefCode = tlSdf.iVefCode) And (tlGsf.iGameNo = tlSdf.iGameNo)
                        gUnpackDateLong tlGsf.iAirDate(0), tlGsf.iAirDate(1), llGsfDate
                        If llLogDate = llGsfDate Then
                            If tlGsf.iAirVefCode > 0 Then
                                mAddLogAlert tlGsf.iAirVefCode, llLogDate, slSubType
                            End If
                            Exit Do
                        End If
                        ilRet = btrGetNext(hlGsf, tlGsf, ilGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    Loop
                End If
            End If
            If hlInGsf <= 0 Then
                ilRet = btrClose(hlGsf)
                btrDestroy hlGsf
            End If
        End If
    End If
    

End Sub

Public Sub gMakeExportAlert(ilVefCode As Integer, llAlertDate As Long, slType As String, slSubType As String)
    Dim ilVff As Integer
    Dim ilVpf As Integer
    Dim llNowDate As Long
    Dim llLastDate As Long
    Dim ilRet As Integer
    Dim SQLQuery As String
    Dim rst_Lst As Recordset
    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
        If ilVefCode = tgVff(ilVff).iVefCode Then
            'TTP 10496 - Affiliate alerts created when log is generated even if there's no spots
            If sgLogStartDate <> "" And sgLogEndDate <> "" Then
                SQLQuery = "SELECT lstCode FROM lst WHERE (lstLogVefCode = " & ilVefCode & " AND lstsdfCode > 0 AND (lstLogDate >= '" & Format(sgLogStartDate, "yyyy-MM-DD") & "' AND lstLogDate <= '" & Format(sgLogEndDate, "yyyy-MM-DD") & "'))"
                Set rst_Lst = gSQLSelectCall(SQLQuery)
                'If rst_Lst.EOF Then "No Log Spots Generated for " & smFWkDate & "-" & smLWkDate & " " & sVehicle
                If rst_Lst.EOF Then
'Debug.Print " - gMakeExportAlert - No Log Spots for VefCode:" & ilVefCode & ", between LogDate:" & Format(sgLogStartDate, "yyyy-MM-DD") & " AND " & Format(sgLogEndDate, "yyyy-MM-DD")
                    'dont Create an Alert because this Vehicle has NO Spots
                    Exit Sub
                End If
            End If
'Debug.Print " - gMakeExportAlert for VefCode:" & ilVefCode & ",  between LogDate:" & Format(sgLogStartDate, "yyyy-MM-DD") & " AND " & Format(sgLogEndDate, "yyyy-MM-DD")
            gUnpackDateLong tgVff(ilVff).iLastAffExptDate(0), tgVff(ilVff).iLastAffExptDate(1), llLastDate
            If llLastDate = 0 Then
                ilVpf = gBinarySearchVpf(ilVefCode)
                If ilVpf <> -1 Then
                    gUnpackDateLong tgVpf(ilVpf).iLLD(0), tgVpf(ilVpf).iLLD(1), llLastDate
                Else
                    llLastDate = 0
                End If
            End If
            If llLastDate > 0 Then
                If (llAlertDate <= llLastDate) Then
                    'TTP 6550: 2/27/14
                    ilRet = gAlertAdd("R", slSubType, 0, ilVefCode, Format$(llAlertDate, "m/d/yy"))
                Else
                    If slType = "A" Then
                        ilRet = gAlertAdd("R", slSubType, 0, ilVefCode, Format$(llAlertDate, "m/d/yy"))
                    Else
                        ilRet = gAlertAdd(slType, slSubType, 0, ilVefCode, Format$(llAlertDate, "m/d/yy"))
                    End If
                End If
            End If
            Exit For
        End If
    Next ilVff

End Sub

Private Sub mAddLogAlert(ilVefCode As Integer, llLogDate As Long, slSubType As String)
    Dim ilVpf As Integer
    Dim llNowDate As Long
    Dim llLastLogDate As Long
    Dim ilRet As Integer
    
    ilVpf = gBinarySearchVpf(ilVefCode)
    If ilVpf <> -1 Then
        llNowDate = gDateValue(Format$(gNow(), "m/d/yy"))
        gUnpackDateLong tgVpf(ilVpf).iLLD(0), tgVpf(ilVpf).iLLD(1), llLastLogDate
        If (llLogDate > llNowDate) And (llLogDate <= llLastLogDate) Then
            ilRet = gAlertAdd("L", slSubType, 0, ilVefCode, Format$(llLogDate, "m/d/yy"))
        End If
    End If
End Sub

Public Function gSxfAdd(hlSxf As Integer, slType As String, tlSdf As SDF) As Integer
    'slType (I)- G=MG or Outside; W=Work Area
    Dim ilRet As Integer
    
    gSxfAdd = False
    If slType = "G" Then
        tlSdf.sWasMG = "N"
    Else
        tlSdf.sFromWorkArea = "N"
    End If
    If (tlSdf.sSpotType <> "X") And ((((tlSdf.sSchStatus = "G") Or (tlSdf.sSchStatus = "O")) And (slType = "G")) Or ((tlSdf.sSchStatus = "M") And (slType = "W"))) Then    'Create sxf
        If slType = "W" Then
            If tlSdf.sSchStatus <> "M" Then
                Exit Function
            End If
        Else
            If (tlSdf.sSchStatus <> "G") And (tlSdf.sSchStatus <> "O") Then
                Exit Function
            End If
        End If
        imSxfRecLen = Len(tmSxf)
        tmSxfSrchKey1.sType = slType
        tmSxfSrchKey1.lSdfCode = tlSdf.lCode
        ilRet = btrGetEqual(hlSxf, tmSxf, imSxfRecLen, tmSxfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get first record as starting point of extend operation
        If ilRet <> BTRV_ERR_NONE Then
            tmSxf.lCode = 0
            tmSxf.lSdfCode = tlSdf.lCode
            tmSxf.sType = slType
            tmSxf.lChfCode = tlSdf.lChfCode
            tmSxf.iLineNo = tlSdf.iLineNo
            tmSxf.iAdfCode = tlSdf.iAdfCode
            tmSxf.iMissedVefCode = tlSdf.iVefCode
            tmSxf.iMissedDate(0) = tlSdf.iDate(0)
            tmSxf.iMissedDate(1) = tlSdf.iDate(1)
            tmSxf.iMissedTime(0) = tlSdf.iTime(0)
            tmSxf.iMissedTime(1) = tlSdf.iTime(1)
            tmSxf.iMissedGameNo = tlSdf.iGameNo
            gPackDate Format(gNow(), "m/d/yy"), tmSxf.iEnteredDate(0), tmSxf.iEnteredDate(1)
            gPackTime Format(gNow(), "h:mm:ssAM/PM"), tmSxf.iEnteredTime(0), tmSxf.iEnteredTime(1)
            tmSxf.iUrfCode = tgUrf(0).iCode
            ilRet = btrInsert(hlSxf, tmSxf, imSxfRecLen, INDEXKEY0)
            If ilRet = BTRV_ERR_NONE Then
                If slType = "G" Then
                    tlSdf.sWasMG = "Y"
                Else
                    tlSdf.sFromWorkArea = "Y"
                End If
                gSxfAdd = True
            End If
        Else
            If slType = "G" Then
                tlSdf.sWasMG = "Y"
            Else
                tlSdf.sFromWorkArea = "Y"
            End If
            gSxfAdd = True
        End If
    End If
End Function

Public Function gSxfDelete(hlSxf As Integer, tlSdf As SDF, Optional blDeleteMG As Boolean = True, Optional blDeleteWA As Boolean = True) As Integer
    Dim ilRet As Integer
    
    If (tlSdf.sWasMG = "Y") Or (tlSdf.sFromWorkArea = "Y") Then
        imSxfRecLen = Len(tmSxf)
        If (tlSdf.sWasMG = "Y") And (blDeleteMG) Then
            tmSxfSrchKey1.sType = "G"
            tmSxfSrchKey1.lSdfCode = tlSdf.lCode
            ilRet = btrGetEqual(hlSxf, tmSxf, imSxfRecLen, tmSxfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hlSxf)
            End If
            tlSdf.sWasMG = "N"
        End If
        If (tlSdf.sFromWorkArea = "Y") And (blDeleteWA) Then
            tmSxfSrchKey1.sType = "W"
            tmSxfSrchKey1.lSdfCode = tlSdf.lCode
            ilRet = btrGetEqual(hlSxf, tmSxf, imSxfRecLen, tmSxfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            If ilRet = BTRV_ERR_NONE Then
                ilRet = btrDelete(hlSxf)
            End If
            tlSdf.sFromWorkArea = "N"
        End If
    Else
        gSxfDelete = BTRV_ERR_NONE
    End If

End Function

Private Sub mTestSDF(ilTestOk As Integer, ilUpdateSdf As Integer, ilRemoveSpot As Integer, ilAdfCode As Integer, llChfCode As Long, llFsfCode As Long, llCntrNo As Long, ilSplitNetworkPriRemoved As Integer, ilTestSplitNetworkLen As Integer, ilBoundIndex As Integer, ilAvailIndex As Integer, ilSpotIndex As Integer, ilTestOnly As Integer, hlGsf As Integer, hlSxf As Integer, hlSdf As Integer, hlSmf As Integer, ilSdfRecLen As Integer, slDate As String, slTime As String, slMsg As String)
    Dim ilRet As Integer
    'Dim llChfRecPos As Long
    'Dim llClfRecPos As Long
    'Dim llCffRecPos As Long
    Dim llClfCode As Long
    Dim llCffCode As Long
    Dim ilTest As Integer
    Dim tlSpot As CSPOTSS
    '7/8/18
    Dim blChkPriceRank As Boolean
    Dim slPrice As String
    
    ilTestOk = True
    ilUpdateSdf = False
    ilRemoveSpot = False
    ilAdfCode = tmSdf.iAdfCode
    llChfCode = tmSdf.lChfCode
    llFsfCode = tmSdf.lFsfCode
    llCntrNo = 0
    blChkPriceRank = False
    If (tmSdf.lChfCode = 0) And (tmSdf.lFsfCode = 0) Then
        ilRemoveSpot = True
    Else
        If tmSdf.lChfCode > 0 Then
        'Test if line exist
            tmChfSrchKey.lCode = tmSdf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                If (tmSdf.sSpotType = "X") And ((tgSpot.iRank And RANKMASK) < 1000) Then
                    If tmChf.sType = "S" Then
                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + PSARANK  '1060
                    ElseIf tmChf.sType = "M" Then
                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + PROMORANK    '1050
                    ElseIf tmChf.sType = "Q" Then
                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + PERINQUIRYRANK   '1030
                    Else
                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + EXTRARANK    '1045
                    End If
                ElseIf (tmSdf.sSpotType <> "X") Then
                    '7/8/18: Add Trade test
                    If tmChf.iPctTrade <> 100 Then
                        If tmChf.sType = "S" Then
                            tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + PSARANK  '1060
                        ElseIf tmChf.sType = "M" Then
                            tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + PROMORANK    '1050
                        ElseIf tmChf.sType = "Q" Then
                            tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + PERINQUIRYRANK   '1030
                        ElseIf tmChf.sType = "T" Then
                            tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + REMNANTRANK  '1020
                        ElseIf tmChf.sType = "R" Then
                            tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + DIRECTRESPONSERANK   '1010
                        '7/8/18: Add reservation
                        ElseIf tmChf.sType = "V" Then
                            tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + RESERVATIONRANK   '1010
                        Else
                            blChkPriceRank = True
                        End If
                    '7/8/18: Added trade setting
                    Else
                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + TRADERANK
                    End If
                End If
                ilAdfCode = tmChf.iAdfCode
                llCntrNo = tmChf.lCntrNo
                If tmChf.sDelete = "Y" Then
                    tmChfSrchKey1.lCntrNo = tmChf.lCntrNo
                    tmChfSrchKey1.iCntRevNo = 32000
                    tmChfSrchKey1.iPropVer = 32000
                    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llCntrNo)
                        If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.sDelete <> "Y") Then
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                    If (tmChf.lCntrNo = llCntrNo) And (tmChf.sDelete <> "Y") Then
                        llChfCode = tmChf.lCode
                        ilAdfCode = tmChf.iAdfCode
                        llCntrNo = tmChf.lCntrNo
                        ilUpdateSdf = True
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
                            If llChfCode <> tmChf.lCode Then
                                ilUpdateSdf = True
                            End If
                            llChfCode = tmChf.lCode
                            ilAdfCode = tmChf.iAdfCode
                            llCntrNo = tmChf.lCntrNo
                            sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Contract Delete Flag Error:" & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
                            ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                            If Not ilTestOnly Then
                                '6/4/16: Replaced GetDirect with getequal
                                'ilRet = btrGetPosition(hmChf, llChfRecPos)
                                Do
                                    '6/4/16: Replaced GetDirect with getequal
                                    'ilRet = btrGetDirect(hmChf, tmChf, imChfRecLen, llChfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                    tmChfSrchKey.lCode = llChfCode
                                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                    tmChf.sDelete = "N"
                                    ilRet = btrUpdate(hmCHF, tmChf, imCHFRecLen)
                                Loop While ilRet = BTRV_ERR_CONFLICT
                            End If
                        Else
                            ilRemoveSpot = True
                        End If
                    End If
                End If
                tmClfSrchKey.lChfCode = llChfCode
                tmClfSrchKey.iLine = tmSdf.iLineNo
                tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                If (ilRet <> BTRV_ERR_NONE) Or (tmClf.lChfCode <> llChfCode) Or (tmClf.iLine <> tmSdf.iLineNo) Then
                    ilRemoveSpot = True
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
                                ilRemoveSpot = False
                                sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Line Contract Code Error:" & slDate & " " & slTime & " Cntr #=" & str$(tmChf.lCntrNo) & " Line=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
                                ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                                If ilTestOnly Then
                                    Exit Do
                                End If
                                '6/4/16: Replaced GetDirect with getequal
                                'ilRet = btrGetPosition(hmClf, llClfRecPos)
                                llClfCode = tmClf.lCode
                                Do
                                    '6/4/16: Replaced GetDirect with getequal
                                    'ilRet = btrGetDirect(hmClf, tmClf, imClfRecLen, llClfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                    tmClfSrchKey2.lCode = llClfCode
                                    ilRet = btrGetEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
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
                                    '6/4/16: Replaced GetDirect with getequal
                                    'ilRet = btrGetPosition(hmCff, llCffRecPos)
                                    llCffCode = tmCff.lCode
                                    Do
                                        '6/4/16: Replaced GetDirect with getequal
                                        'ilRet = btrGetDirect(hmCff, tmCff, imCffRecLen, llCffRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                                        tmCffSrchKey1.lCode = llCffCode
                                        ilRet = btrGetEqual(hmCff, tmCff, imCffRecLen, tmCffSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
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
                Else
                    If (tmSdf.sSchStatus = "S") And (tmSdf.iVefCode <> tmClf.iVefCode) Then
                        ilRemoveSpot = True
                    '7/8/18: Added setting the Rank for No Charge
                    Else
                        If blChkPriceRank And (tmSdf.sSchStatus = "S") Then
                            ilRet = gGetFlightPrice(tmSdf, tmClf, hmCff, hlSmf, slPrice)
                            If ilRet = True Then
                                If tgPriceCff.sPriceType = "T" Then
                                    If tgPriceCff.lActPrice = 0 Then
                                        tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + NOCHARGERANK
                                    End If
                                Else
                                    tgSpot.iRank = (tgSpot.iRank And PRICELEVELMASK) + NOCHARGERANK
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                ilRemoveSpot = True
            End If
        Else
            tmFsfSrchKey0.lCode = tmSdf.lFsfCode
            ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet = BTRV_ERR_NONE Then
                ilAdfCode = tmFsf.iAdfCode
                llCntrNo = 0
                'Determine if this is the latest fsf
                tmFsfSrchKey4.lPrevFsfCode = tmFsf.lCode
                ilRet = btrGetEqual(hmFsf, tmTFsf, imFsfRecLen, tmFsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                Do While (ilRet = BTRV_ERR_NONE)
                    If (tmTFsf.sSchStatus <> "F") Then
                        Exit Do
                    End If
                    tmFsf = tmTFsf
                    tmFsfSrchKey4.lPrevFsfCode = tmTFsf.lCode
                    ilRet = btrGetEqual(hmFsf, tmTFsf, imFsfRecLen, tmFsfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                Loop
                If tmSdf.lFsfCode <> tmFsf.lCode Then
                    ilUpdateSdf = True
                    sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = "Feed Code Error:" & slDate & " " & slTime & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
                    ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
                End If
                ilAdfCode = tmFsf.iAdfCode
                llFsfCode = tmFsf.lCode
                'Check vehicle
                If (tmSdf.sSchStatus = "S") And (tmSdf.iVefCode <> tmFsf.iVefCode) Then
                    ilRemoveSpot = True
                End If
            Else
                ilRemoveSpot = True
            End If
        End If
    End If
    If (Not ilRemoveSpot) And (tmVef.sType = "A") Then
        ilRemoveSpot = True
    End If
    If (Not ilRemoveSpot) And (ilTestSplitNetworkLen) Then
        If (tgSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
            For ilTest = ilSpotIndex - 1 To ilAvailIndex + 1 Step -1
               LSet tlSpot = tgSsf(ilBoundIndex).tPas(ADJSSFPASBZ + ilTest)
                If (tlSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
                    If (tgSpot.iPosLen And &HFFF) <> (tlSpot.iPosLen And &HFFF) Then
                        ilRemoveSpot = True
                    End If
                    Exit For
                End If
            Next ilTest
        End If
    End If
    If ilRemoveSpot Then
        If ilTestOnly Then
            If tmVef.sType = "A" Then
                slMsg = "Sdf in Air Veh: "
            Else
                slMsg = "Sdf Data Missing: "
            End If
        Else
            If (tgSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
                ilSplitNetworkPriRemoved = True
            Else
                ilSplitNetworkPriRemoved = False
            End If
            ilRet = gMakeTracer(hlSdf, tmSdf, 0, hmStf, -1, "M", "P", tmSdf.iRotNo, hlGsf)
            ilRet = gRemoveSmf(hlSmf, tmSmf, tmSdf, hlSxf)  'resets missed date
            Do
                tmSdfSrchKey3.lCode = tmSdf.lCode
                ilRet = btrGetEqual(hlSdf, tmSdf, ilSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORWRITE)
                'tmSRec = tmSdf
                'ilCRet = gGetByKeyForUpdate("SDF", hlSdf, tmSRec)
                'tmSdf = tmSRec
                'If ilCRet <> BTRV_ERR_NONE Then
                '    igBtrError = ilCRet
                '    sgErrLoc = "gMakeSSF-Get by Key Sdf(16)"
                '    ilTestOk = False
                'End If
                ilRet = btrDelete(hlSdf)
                'If ilRet = BTRV_ERR_CONFLICT Then
                '    ilCRet = btrGetDirect(hlSdf, tmSdf, ilSdfRecLen, llSdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                'End If
            Loop While ilRet = BTRV_ERR_CONFLICT
            If ilRet <> BTRV_ERR_NONE Then
                igBtrError = gConvertErrorCode(ilRet)
                sgErrLoc = "gMakeSSF-Delete Sdf(17)"
                ilTestOk = False
            End If
            If tmVef.sType = "A" Then
                slMsg = "Sdf in Air Veh, Removed: "
            Else
                slMsg = "Sdf Data Missing, Removed: "
            End If
        End If
        If llCntrNo = 0 Then
            sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Line #=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
        Else
            sgSSFErrorMsg(UBound(sgSSFErrorMsg)) = slMsg & slDate & " " & slTime & " Cntr #=" & str$(llCntrNo) & " Line #=" & str$(tmSdf.iLineNo) & " Sdf ID=" & str$(tgSpot.lSdfCode) & " " & Trim$(tmVef.sName)
        End If
        ReDim Preserve sgSSFErrorMsg(0 To UBound(sgSSFErrorMsg) + 1) As String
    End If
End Sub
Private Sub mUnassignRegionalAsNeeded(hlCrf As Integer, llSDFCode As Long, slSSFDate As String, llSDFTime As Long, ilAVAILAnfCode As Integer, slCLFLive As String, slSDFSheduleStatus As String, ilSdfVefCode As Integer)
    Dim blKeepAssignment As Boolean
    Dim ilRet As Integer
    Dim slSql As String
    Dim rst As ADODB.Recordset
    Dim llCrfCode As Long
    Dim llRsfCode As Long
    Dim ilDay As Integer
    Dim slRSFsToDelete As String
    
    Dim tlCrfSrchKey1 As CRFKEY1
    Dim ilCrfRecLen As Integer
    Dim tlCrf As CRF
    
    Dim llSAsgnDate As Long
    Dim llEAsgnDate As Long
    Dim llSAsgnTime As Long
    Dim llEAsgnTime As Long
    Dim llSDFDate As Long
    
    Dim ilVef As Integer
    Dim rstCvfOrPvf As ADODB.Recordset
    Dim llCvfCrfCode As Long
    Dim llPvfCode As Long
    
    slRSFsToDelete = ""
    ilCrfRecLen = Len(tlCrf)
    ilDay = gWeekDayStr(slSSFDate)
    llSDFDate = gDateValue(slSSFDate)
    'What regionals are assigned?   For SDF, each RSF to get CRF code
    slSql = "Select rsfCode, rsfCrfCode from RSF_Region_Schd_Copy where rsfSdfCode = " & llSDFCode
    Set rst = gSQLSelectCall(slSql)
    Do While Not rst.EOF
        blKeepAssignment = False
        llCrfCode = rst!rsfCrfCode
        llRsfCode = rst!rsfCode
        ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, llCrfCode, INDEXKEY0, BTRV_LOCK_NONE)
        If ilRet = BTRV_ERR_NONE Then
            If mRegionalActiveStateMatches(tlCrf.sState, slCLFLive, tlCrf.sLiveCopy) Then
                'matches SSF's(also SDF) day?
                If (tlCrf.sDay(ilDay) = "Y") Then
                    gUnpackDateLong tlCrf.iStartDate(0), tlCrf.iStartDate(1), llSAsgnDate
                    gUnpackDateLong tlCrf.iEndDate(0), tlCrf.iEndDate(1), llEAsgnDate
                    'Date valid?
                    If llSDFDate >= llSAsgnDate And llSDFDate <= llEAsgnDate Then
                        gUnpackTimeLong tlCrf.iStartTime(0), tlCrf.iStartTime(1), False, llSAsgnTime
                        gUnpackTimeLong tlCrf.iEndTime(0), tlCrf.iEndTime(1), True, llEAsgnTime
                        'Time valid?
                        If (llSDFTime >= llSAsgnTime) And (llSDFTime <= llEAsgnTime) Then
                            'Check avail type  must be in/cannot be in
                            If (tlCrf.sInOut = "I") Or (tlCrf.sInOut = "O") Then
                                If tlCrf.sInOut = "I" Then
                                    If tlCrf.ianfCode = ilAVAILAnfCode Then
                                        blKeepAssignment = True
                                    End If
                                Else
                                    If tlCrf.ianfCode <> ilAVAILAnfCode Then
                                        blKeepAssignment = True
                                    End If
                                End If
                            Else
                                blKeepAssignment = True
                            End If
                        End If
                    End If
                End If 'matching various basic tests
                If blKeepAssignment Then
                    ' makegood or outside must test the vehicle is ok
                    If slSDFSheduleStatus = "G" Or slSDFSheduleStatus = "O" Then
                        blKeepAssignment = False
                        ' multiple vehicles have no crfVefCode. Stored in cvf
                        If tlCrf.iVefCode = 0 Then
                            llCvfCrfCode = tlCrf.lCode
                            'for cvf Linking to itself
                            Do While llCvfCrfCode > 0
                                slSql = "Select * from CVF_Copy_Vehicles where cvfCrfCode =" & llCvfCrfCode
                                Set rstCvfOrPvf = gSQLSelectCall(slSql)
                                If Not rstCvfOrPvf.EOF Then
                                    If mCvfVehiclesMatch(rstCvfOrPvf, ilSdfVefCode, llSDFDate) Then
                                        blKeepAssignment = True
                                        Exit Do
                                    End If
                                    'stop the recursive looping
                                    llCvfCrfCode = 0
                                    If blKeepAssignment = False And rstCvfOrPvf!cvfLkCvfCode > 0 Then
                                        llCvfCrfCode = rstCvfOrPvf!cvfLkCvfCode
                                    End If
                                End If
                            Loop
                        Else
                            'test to see if airing
                            ilVef = gBinarySearchVef(tlCrf.iVefCode)
                            If ilVef <> -1 Then
                                'airing? get selling
                                If tgMVef(ilVef).sType = "A" Then
                                    blKeepAssignment = mSellingVehiclesMatch(tlCrf.iVefCode, llSDFDate, ilSdfVefCode)
                                'packages?
                                ElseIf tgMVef(ilVef).sType = "P" Then
                                    'loop thru pvf vehicles in a single pvf record. There is a link to another pvf file for more vehicles.
                                    llPvfCode = tgMVef(ilVef).lPvfCode
                                    'Dynamic package?
                                    If llPvfCode = 0 Then
                                        blKeepAssignment = mDynamicPackagesMatch(tlCrf.lChfCode, tlCrf.iVefCode, llSDFDate, ilSdfVefCode)
                                    Else
                                        'standard package.  Let's gather the hidden lines. Recursive
                                        Do While llPvfCode > 0
                                            slSql = "Select * from PVF_Package_Vehicle where pvfcode = " & llPvfCode
                                            Set rstCvfOrPvf = gSQLSelectCall(slSql)
                                            If mPackageVehiclesMatch(rstCvfOrPvf, llSDFDate, ilSdfVefCode) Then
                                                blKeepAssignment = True
                                                Exit Do
                                            End If
                                            llPvfCode = 0
                                            If blKeepAssignment = False And rstCvfOrPvf!pvfLkPvfCode > 0 Then
                                                llPvfCode = rstCvfOrPvf!pvfLkPvfCode
                                            End If
                                        Loop
                                    End If 'standard vs dynamic package
                                'finally!  the simplest test
                                ElseIf tlCrf.iVefCode = ilSdfVefCode Then
                                    blKeepAssignment = True
                                End If 'Airing vehicle?
                            End If 'Vef ok
                        End If 'Multiple Crf vehicles?
                    End If 'difficult vehicle matching for MG/Outside only
                End If 'basic matches were good
            End If 'it's a day
        End If 'ilret ok
        If Not blKeepAssignment Then
            slRSFsToDelete = slRSFsToDelete & llRsfCode & ","
        End If
        rst.MoveNext
    Loop
    'delete outside of loop
    If Len(slRSFsToDelete) > 0 Then
        slRSFsToDelete = mLoseLastLetterIfComma(slRSFsToDelete)
        slSql = "Delete from RSF_Region_Schd_Copy where rsfcode in (" & slRSFsToDelete & ")"
        gSQLWaitNoMsgBox slSql, False
    End If
End Sub
Private Function mRegionalActiveStateMatches(slState As String, slLive As String, slLiveCopyFromCRF As String)
    Dim blRet As Boolean
    'return true if the sdf record matches the crf
    blRet = True
    'is active
    If slState <> "D" Then
        'slLive comes from clf record
        If slLive = "L" Then
            If slLiveCopyFromCRF <> "L" Then
                blRet = False
            End If
        ElseIf slLive = "M" Then
            If slLiveCopyFromCRF <> "M" Then
                blRet = False
            End If
        ElseIf slLive = "S" Then
            If slLiveCopyFromCRF <> "S" Then
                blRet = False
            End If
        ElseIf slLive = "P" Then
            If slLiveCopyFromCRF <> "P" Then
                blRet = False
            End If
        ElseIf slLive = "Q" Then
            If slLiveCopyFromCRF <> "Q" Then
                blRet = False
            End If
        Else
            If (slLiveCopyFromCRF = "L") Or (slLiveCopyFromCRF = "M") Or (slLiveCopyFromCRF = "S") Or (slLiveCopyFromCRF = "P") Or (slLiveCopyFromCRF = "Q") Then
                blRet = False
            End If
        End If
    End If
    mRegionalActiveStateMatches = blRet
End Function
Private Function mLoseLastLetterIfComma(slInput As String) As String
    Dim llLength As Long
    Dim slNewString As String
    Dim llLastLetter As Long
    
    llLength = Len(slInput)
    llLastLetter = InStrRev(slInput, ",")
    If llLength > 0 And llLastLetter = llLength Then
        slNewString = Mid(slInput, 1, llLength - 1)
    Else
        slNewString = slInput
    End If
    mLoseLastLetterIfComma = slNewString
End Function
Private Function mCvfVehiclesMatch(rstCvf As ADODB.Recordset, ilSdfVefCode As Integer, llSDFDate As Long) As Boolean
    Dim blRet As Boolean
    Dim ilCvfCode As Integer
    Dim ilIndex As Integer
    Dim ilVef As Integer
    Dim llPvfCode As Long
    Dim slSql As String
    Dim rstPvf As ADODB.Recordset
    
    blRet = False
    '3 to 102
    For ilIndex = 3 To 102 Step 1
        'if 10524 implemented, change here
        ilCvfCode = rstCvf(ilIndex).Value
        If ilCvfCode > 0 Then
            ilVef = gBinarySearchVef(ilCvfCode)
            If ilVef <> -1 Then
                'handle airing to selling
                If tgMVef(ilVef).sType = "A" Then
                    If mSellingVehiclesMatch(ilCvfCode, llSDFDate, ilSdfVefCode) Then
                        blRet = True
                        GoTo FINISH
                    End If
                'packages
                ElseIf tgMVef(ilVef).sType = "P" Then
                    'loop thru pvf vehicles in a single pvf record. There is a link to another pvf file for more vehicles.
                    llPvfCode = tgMVef(ilVef).lPvfCode
                    Do While llPvfCode > 0
                        slSql = "Select * from PVF_Package_Vehicle where pvfcode = " & llPvfCode
                        Set rstPvf = gSQLSelectCall(slSql)
                        If Not rstPvf.EOF Then
                            If mPackageVehiclesMatch(rstPvf, llSDFDate, ilSdfVefCode) Then
                                blRet = True
                                GoTo FINISH
                            End If
                        End If
                        llPvfCode = 0
                        If blRet = False And rstPvf!pvfLkPvfCode > 0 Then
                            llPvfCode = rstPvf!pvfLkPvfCode
                        End If
                    Loop
                'simplest
                Else
                    If ilCvfCode = ilSdfVefCode Then
                        blRet = True
                        GoTo FINISH
                    End If
                End If 'Airing vehicle?
            End If
        Else
            GoTo FINISH
        End If
    Next
FINISH:
    mCvfVehiclesMatch = blRet
End Function
Private Function mSellingVehiclesMatch(ilCrfVefCode As Integer, llSDFDate As Long, ilSdfVefCode As Integer) As Boolean
    Dim blRet As Boolean
    Dim hlVlf As Integer
    Dim tlVlf() As VLF
    Dim llDate As Long
    Dim ilLoop As Integer
    
    blRet = False
    hlVlf = CBtrvTable(TWOHANDLES)
    If btrOpen(hlVlf, "", sgDBPath & "Vlf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE) = BTRV_ERR_NONE Then
        ReDim tlVlf(0 To 0)
        gObtainVlf "A", hlVlf, ilCrfVefCode, llSDFDate, tlVlf
        For ilLoop = LBound(tlVlf) To UBound(tlVlf) - 1 Step 1
            If tlVlf(ilLoop).iSellCode = ilSdfVefCode Then
                blRet = True
                Exit For
            End If
        Next
    End If
    btrClose hlVlf
    btrDestroy hlVlf
    mSellingVehiclesMatch = blRet
End Function
Private Function mPackageVehiclesMatch(rstPvf As ADODB.Recordset, llSDFDate As Long, ilSdfVefCode As Integer) As Boolean
    Dim blRet As Boolean
    Dim slSql As String
    Dim ilPvfVefCode As Integer
    Dim ilIndex As Integer
    Dim ilVef As Integer
        
    blRet = False
    If Not rstPvf.EOF Then
        '3 to 27 Go through each hidden vehicle
        For ilIndex = 3 To 27 Step 1
        'if 10524 implemented, change here
            ilPvfVefCode = rstPvf(ilIndex).Value
            If ilPvfVefCode > 0 Then
                'test airing
                ilVef = gBinarySearchVef(ilPvfVefCode)
                If ilVef <> -1 Then
                    'handle airing to selling
                    If tgMVef(ilVef).sType = "A" Then
                        If mSellingVehiclesMatch(ilPvfVefCode, llSDFDate, ilSdfVefCode) Then
                            blRet = True
                            GoTo FINISH
                        End If
                    'basic
                    ElseIf ilPvfVefCode = ilSdfVefCode Then
                        blRet = True
                        GoTo FINISH
                    End If
                End If
            Else
                GoTo FINISH
            End If
        Next
        '10621 always just one record--don't need to look for another
        'rstPvf.MoveNext
    End If
FINISH:
    mPackageVehiclesMatch = blRet
End Function
Private Function mDynamicPackagesMatch(llChfCode As Long, ilCrfVefCode As Integer, llSDFDate As Long, ilSdfVefCode As Integer) As Boolean
    Dim blRet As Boolean
    Dim ilClfVef As Integer
    Dim slSql As String
    Dim rstClfDynamicPkgLines As ADODB.Recordset
    Dim rstClfVefs As ADODB.Recordset
    Dim ilLine As Integer
    Dim ilVef As Integer
    
    blRet = False
    slSql = "Select clfLine from CLF_Contract_Line where clfchfcode = " & llChfCode & " and clfvefcode = " & ilCrfVefCode
    Set rstClfDynamicPkgLines = gSQLSelectCall(slSql)
    Do While Not rstClfDynamicPkgLines.EOF
        ilLine = rstClfDynamicPkgLines!clfLine
        slSql = "Select clfVefCode from CLF_Contract_Line where clfchfcode = " & llChfCode & "  and clfpklineno = " & ilLine
        Set rstClfVefs = gSQLSelectCall(slSql)
        Do While Not rstClfVefs.EOF
            ilClfVef = rstClfVefs!clfVefCode
            'test airing
            ilVef = gBinarySearchVef(ilClfVef)
            If ilVef <> -1 Then
                'handle airing to selling
                If tgMVef(ilVef).sType = "A" Then
                    If mSellingVehiclesMatch(ilClfVef, llSDFDate, ilSdfVefCode) Then
                        blRet = True
                        GoTo FINISH
                    End If
                ElseIf ilClfVef = ilSdfVefCode Then
                    blRet = True
                    GoTo FINISH
                End If
            End If
            rstClfVefs.MoveNext
        Loop
        rstClfDynamicPkgLines.MoveNext
    Loop
FINISH:
    mDynamicPackagesMatch = blRet
End Function
