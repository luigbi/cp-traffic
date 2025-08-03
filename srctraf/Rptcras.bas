Attribute VB_Name = "RPTCRAS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcras.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
Public igmnfRated As Integer            'rated mnf code
Public igmnfNonRated As Integer         'non rated mnf code
Public igmnfSuburban As Integer         'Suburban mnf code
Public igFoundLoc As Integer               '1-9-02
Public igFoundNat As Integer            '1-9-01
Dim tmChfAdvtExt() As CHFADVTEXT
Dim lmSDFRecdPos() As Long      'SDF record positions of the MG/out that hvae been processed
Dim hmMnf As Integer            'Multi-Names file handle
Dim tmMnf As MNF                'MNF record image
Dim imMnfRecLen As Integer      'MNF record length
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmCHF As Integer            'Contract header file handle
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClfSrchKey As CLFKEY0
Dim tmClf As CLF
Dim hmSdf As Integer            'Spot file handle
Dim imSdfRecLen As Integer        'SDF record length
Dim hmSmf As Integer            'MG file handle
Dim tmSmf As SMF                'SMF record image
Dim tmSmfSrchKey As SMFKEY0     'SMF record image
Dim tmSmfSrchKey1 As LONGKEY0
Dim imSmfRecLen As Integer        'SMF record length
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim tmZeroGrf As GRF                   'GRF image
Dim tmGrf() As GRF
'*************************************************************
'*                                                           *
'*      Procedure Name:gCRAdvtSpotCounts                     *
'*                                                           *
'*             Created:10/22/98      By:D. Hosaka            *
'*            Modified:              By:                     *
'*                                                           *
''      Access all spot data from SSF between two dates      *
'       entered by the user.  Determine the contractual      *
'*      # of spots ordered by the scheduled + missed +       *
'*      makegoods.  Determine the aired spots by the MG      *
'*      and regular scheduled spots.   Show the difference.  *
'*      Break these up into rated & non-rated vehicles.      *
'                                                            *
'     3-30-05 Exclude bb spots (in mObtainspotsbydate)       *
'*************************************************************
Sub gCRAdvtSpotCounts()
'
    Dim ilRet As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim tlCntTypes As CNTTYPES
    Dim ilLoop As Integer
    Dim llStartDate As Long                     'user entered start date
    Dim llEndDate As Long                       'user entered end date
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)


    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmMnf
        btrDestroy hmCHF
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVef
        btrDestroy hmMnf
        btrDestroy hmCHF
       Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSmf
        btrDestroy hmVef
        btrDestroy hmMnf
        btrDestroy hmCHF
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)
    hmGrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
        btrDestroy hmSmf
        btrDestroy hmVef
        btrDestroy hmMnf
        btrDestroy hmCHF
        Exit Sub
    End If
    imGrfRecLen = Len(tmZeroGrf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSdf
        btrDestroy hmGrf
        btrDestroy hmSmf
        btrDestroy hmVef
        btrDestroy hmMnf
        btrDestroy hmCHF
        Exit Sub
    End If
    imSdfRecLen = Len(hmSdf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmSdf
        btrDestroy hmGrf
        btrDestroy hmSmf
        btrDestroy hmVef
        btrDestroy hmMnf
        btrDestroy hmCHF
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)
    'these filters unused for now
'    tlCntTypes.iHold = gSetCheck(RptSelAS!ckcSelC1(0).Value)
'    tlCntTypes.iOrder = gSetCheck(RptSelAS!ckcSelC1(1).Value)
'    tlCntTypes.iStandard = gSetCheck(RptSelAS!ckcSelC1(2).Value)
'    tlCntTypes.iReserv = gSetCheck(RptSelAS!ckcSelC1(3).Value)
'    tlCntTypes.iRemnant = gSetCheck(RptSelAS!ckcSelC1(4).Value)
'    tlCntTypes.iDR = gSetCheck(RptSelAS!ckcSelC1(5).Value)
'    tlCntTypes.iPI = gSetCheck(RptSelAS!ckcSelC1(6).Value)
'    tlCntTypes.iTrade = gSetCheck(RptSelAS!ckcSelC1(9).Value)
'    tlCntTypes.iMissed = gSetCheck(RptSelAS!ckcSelC1(10).Value)
'    tlCntTypes.iNC = gSetCheck(RptSelAS!ckcSelC1(11).Value)
'    tlCntTypes.iXtra = gSetCheck(RptSelAS!ckcSelC1(12).Value)
'    tlCntTypes.iPSA = gSetCheck(RptSelAS!ckcSelC1(7).Value)
'    tlCntTypes.iPromo = gSetCheck(RptSelAS!ckcSelC1(8).Value)
    tlCntTypes.iRated = False
    tlCntTypes.iNonRAted = False
    tlCntTypes.iSuburban = False
    If RptSelAS!ckcLocNatl(0).Value = vbChecked Or RptSelAS!ckcLocNatl(1).Value = vbChecked Then
        tlCntTypes.iCntrSpots = True
    Else
        tlCntTypes.iCntrSpots = False      '9-17-04 include local spots (vs network)
    End If
    tlCntTypes.iFeedSpots = gSetCheck(RptSelAS!ckcCntrFeed(1).Value)     '9-17-04 include network (feed ) spots


    'change over from Rated, Non-rated or both to Check boxes for RAted, Non-Rated & Suburban
    'Set flag to include rated if user wants it
    'If RptSelAS!rbcSelC2(0).Value Or RptSelAS!rbcSelC2(2).Value Then
    If gSetCheck(RptSelAS!ckcSelC3(0).Value) Then
        tlCntTypes.iRated = True
    End If
    'Set flag to include non-rated if user wants it
    If gSetCheck(RptSelAS!ckcSelC3(1).Value) Then
        tlCntTypes.iNonRAted = True
    End If
    If gSetCheck(RptSelAS!ckcSelC3(2).Value) Then
        tlCntTypes.iSuburban = True
    End If
    'If RptSelAS!rbcSelC2(1).value Or RptSelAS!rbcSelC2(2).value Then
    '    tlCntTypes.iNonRated = True
    'End If
'    slStartDate = RptSelAS!edcSelCFrom.Text   'Start date
    slStartDate = RptSelAS!CSI_CalFrom.Text   'Start date       9-4-19 use csi calendar control vs edit box
    llStartDate = gDateValue(slStartDate)
'    slEndDate = RptSelAS!edcSelCFrom1.Text   'End date
    slEndDate = RptSelAS!CSI_CalTo.Text   'End date
    llEndDate = gDateValue(slEndDate)

    'Gather all spots between requested dates
    ilRet = mObtainSpotByDate(RptSelAS, slStartDate, slEndDate, tlCntTypes)

    'send across rated, non rated requests for report heading
    For ilLoop = LBound(tmGrf) To UBound(tmGrf) - 1
        'tmGrf(ilLoop).iPerGenl(5) = tmGrf(ilLoop).iPerGenl(3) - tmGrf(ilLoop).iPerGenl(1)       'rated ordered difference
        tmGrf(ilLoop).iPerGenl(4) = tmGrf(ilLoop).iPerGenl(2) - tmGrf(ilLoop).iPerGenl(0)       'rated ordered difference
        'tmGrf(ilLoop).iPerGenl(6) = tmGrf(ilLoop).iPerGenl(4) - tmGrf(ilLoop).iPerGenl(2)       'non-rated ordered difference
        tmGrf(ilLoop).iPerGenl(5) = tmGrf(ilLoop).iPerGenl(3) - tmGrf(ilLoop).iPerGenl(1)       'non-rated ordered difference
        'tmGrf(ilLoop).sBktType = slStatus       'send across user request
        ilRet = btrInsert(hmGrf, tmGrf(ilLoop), imGrfRecLen, INDEXKEY0)
    Next ilLoop

    sgCntrForDateStamp = ""         'init incase of re-entering
    Erase tmGrf
    Erase lmSDFRecdPos
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmClf)
    btrDestroy hmSdf
    btrDestroy hmSmf
    btrDestroy hmVef
    btrDestroy hmCHF
    btrDestroy hmMnf
    btrDestroy hmGrf
    btrDestroy hmClf

End Sub
'
'
'           Look for entry with tlGRF array to find
'           a matching contract #.  Build data for
'           rated/non-rated by contrac #
'
'           Created: 10/23/98
'
'           <input> tlSdfExt - image of spot gathered (from sdf or smf)
'                   tlcnttypes - if selectivity of types of contracts, inclusions/exclusions of types
'                   llStartDate - User entered start date
'                   llEndDate - User entered end date
'                   ilSdfOrSmf - 0 = Sdf, 1 = SMF
'           <output> tmGrf - array of contracts information containing spot
'                   counts for rated/non-rated ordered vs scheduled
'
'       10/29/98 shadow doesnt want to see any crossover counts when they
'               move from rated to non-rated or vice versa if asking for
'               only rated or only non-rated.  Previously, this was showing
'               counts for the aired,regardless of asking for both rated or
'               non-rated, but not counted for ordered.  Now excluding it all
'               when one or the other requested.
'       11/1/98 Add option to include Suburban.  These counts are added into
'               the same bucket as Non-rated (Shadow request)
'       12-28-91 Add option to print by vehicle within contract within advt
'       8-23-04 If not using research ratings (rated/non-rated, but some research vehicle groups
'               defined, produce results for the "Local/National" version.

Sub mCreateGRFbyCnt(tlSdfExt As SDFEXT, tlCntTypes As CNTTYPES, llStartDate As Long, llEndDate As Long, ilSdforSmf As Integer)
Dim ilLoop As Integer
Dim ilFound As Integer
Dim ilSchRating As Integer       'rated/non rated mnf code from vehicle
Dim ilMissedRating As Integer   'rated/non-rated mnf code from missed part of mg
Dim ilMissedLoop As Integer
Dim ilRet As Integer
Dim llMissedDate As Long
Dim ilTempVef As Integer            '7-15-02
Dim ilOKtoSeeVeh As Integer         '11-13-03

    ilSchRating = 0
    ilMissedRating = 0
    'Determine whether the spot to process is a Rated or Non-Rated spot
    'For ilLoop = LBound(tgMVef) To UBound(tgMVef)
    '    If tgMVef(ilLoop).iCode = tlSdfExt.iVefCode Then         'scheduled vehicle
        ilLoop = gBinarySearchVef(tlSdfExt.iVefCode)
        If ilLoop <> -1 Then
            ilSchRating = tgMVef(ilLoop).iMnfVehGp5Rsch
            ilMissedRating = ilSchRating
            If igmnfRated = 0 Or igmnfNonRated = 0 Then      'research ratings (rated, nonrated, suburban doesnt exist)
                ilSchRating = 0
                ilMissedRating = 0
            End If
    '        Exit For
        End If
    'Next ilLoop
    'If coming from processing Makegoods (outsides), determine the rated Or Non-Rated status of the missed spot
    If ilSdforSmf = 1 Then                                      'coming from schedule spots , missed vehicle vs sch vehicle doesnt matter
        ilTempVef = tlSdfExt.iStatus                '7-15-02
        If tlSdfExt.iVefCode <> tlSdfExt.iStatus Then                   'missed vehicle different than mg vehicle
            'For ilLoop = LBound(tgMVef) To UBound(tgMVef)
            '    If tgMVef(ilLoop).iCode = tlSdfExt.iStatus Then      'look for missed vehicles research status (rated/non-rated)
                ilLoop = gBinarySearchVef(tlSdfExt.iStatus)
                If ilLoop <> -1 Then
                    'ilSchRating = tgMVef(ilLoop).iMnfVehGp5Rsch
                    ilMissedRating = tgMVef(ilLoop).iMnfVehGp5Rsch
            '        Exit For
                End If
            'Next ilLoop
        End If
    Else                    '7-15-02
        ilTempVef = tlSdfExt.iVefCode
    End If
    ilOKtoSeeVeh = gUserAllowedVehicle(tlSdfExt.iVefCode)   '11-13-03 see if user allowed to see spot from scheduled veh or original missed veh if mg or out

    For ilLoop = LBound(tmGrf) To UBound(tmGrf) - 1 Step 1          'determine which bucket to add count into
                                                                    'if crossing vehicle and its different rated/non rated status,
                                                                    'show it under the correct status.  Also, even if that rated/nonrated
                                                                    'status not selected, show it to highlight discrepancy
        If RptSelAS!ckcVehicle.Value = vbChecked Then               '12-28-01
            If ilSdforSmf = 0 Then          '7-15-02 coming from SDF
                ilTempVef = tlSdfExt.iVefCode
                If tmGrf(ilLoop).lChfCode = tlSdfExt.lChfCode And tmGrf(ilLoop).iVefCode = tlSdfExt.iVefCode And tmGrf(ilLoop).lCode4 = tlSdfExt.lMdDate Then
                    ilFound = True
                    Exit For
                End If
            Else                     '7-15-02 coming from smf (missed part to figure out ordered)
                ilTempVef = tlSdfExt.iStatus
                If tmGrf(ilLoop).lChfCode = tlSdfExt.lChfCode And tmGrf(ilLoop).iVefCode = tlSdfExt.iStatus And tmGrf(ilLoop).lCode4 = tlSdfExt.lMdDate Then
                    ilFound = True
                    Exit For
                End If
            End If
        Else
            If tmGrf(ilLoop).lChfCode = tlSdfExt.lChfCode And tmGrf(ilLoop).lCode4 = tlSdfExt.lMdDate Then
                ilFound = True
                Exit For
            End If
        End If
    Next ilLoop

    If Not ilFound And ilOKtoSeeVeh Then            '11-13-03
        ilLoop = UBound(tmGrf)

        '4-28-17 if outside/mg on different vehicle, reflect the correct vehicle where the origina spot came from
        If ilSdforSmf = 0 Then
            tmGrf(ilLoop).iVefCode = tlSdfExt.iVefCode
        Else
            tmGrf(ilLoop).iVefCode = tlSdfExt.iStatus
        End If
        tmGrf(ilLoop).lChfCode = tlSdfExt.lChfCode
        tmGrf(ilLoop).iAdfCode = tlSdfExt.iAdfCode
        tmGrf(ilLoop).iGenDate(0) = igNowDate(0)
        tmGrf(ilLoop).iGenDate(1) = igNowDate(1)
        tmGrf(ilLoop).lCode4 = tlSdfExt.lMdDate         'fsfcode overlaps mddate
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf(ilLoop).lGenTime = lgNowTime
        'gPackDateLong llStartDate, tmGrf(ilLoop).iDateGenl(0, 1), tmGrf(ilLoop).iDateGenl(1, 1)
        gPackDateLong llStartDate, tmGrf(ilLoop).iDateGenl(0, 0), tmGrf(ilLoop).iDateGenl(1, 0)
        'gPackDateLong llEndDate, tmGrf(ilLoop).iDateGenl(0, 2), tmGrf(ilLoop).iDateGenl(1, 2)
        gPackDateLong llEndDate, tmGrf(ilLoop).iDateGenl(0, 1), tmGrf(ilLoop).iDateGenl(1, 1)
        ReDim Preserve tmGrf(0 To UBound(tmGrf) + 1)
    End If


    'determine which bucket this spot goes into
    If ilOKtoSeeVeh Then                   '11-13-03
        If ilSdforSmf = 0 Then                                  'coming from processing SDF (all aired or missed spots)
            If tlSdfExt.sSchStatus = "S" Then                   'scheduled

                If ilSchRating = igmnfRated And tlCntTypes.iRated Then
                    If tlSdfExt.sSpotType <> "X" Then
                        'tmGrf(ilLoop).iPerGenl(1) = tmGrf(ilLoop).iPerGenl(1) + 1   'rated ordered
                        tmGrf(ilLoop).iPerGenl(0) = tmGrf(ilLoop).iPerGenl(0) + 1   'rated ordered
                    End If
                    'tmGrf(ilLoop).iPerGenl(3) = tmGrf(ilLoop).iPerGenl(3) + 1   'rated aired
                    tmGrf(ilLoop).iPerGenl(2) = tmGrf(ilLoop).iPerGenl(2) + 1   'rated aired
                Else
                    If (ilSchRating = igmnfNonRated And tlCntTypes.iNonRAted) Or (ilSchRating = igmnfSuburban And tlCntTypes.iSuburban) Then
                        If tlSdfExt.sSpotType <> "X" Then
                            'tmGrf(ilLoop).iPerGenl(2) = tmGrf(ilLoop).iPerGenl(2) + 1   'non rated ordered
                            tmGrf(ilLoop).iPerGenl(1) = tmGrf(ilLoop).iPerGenl(1) + 1   'non rated ordered
                        End If
                        'tmGrf(ilLoop).iPerGenl(4) = tmGrf(ilLoop).iPerGenl(4) + 1   'non rated aired
                        tmGrf(ilLoop).iPerGenl(3) = tmGrf(ilLoop).iPerGenl(3) + 1   'non rated aired
                    End If
                End If
                'missed, hidden or cancelled
            ElseIf tlSdfExt.sSchStatus = "M" Or tlSdfExt.sSchStatus = "H" Or tlSdfExt.sSchStatus = "C" Then
                If ilSchRating = igmnfRated And tlCntTypes.iRated Then
                    'tmGrf(ilLoop).iPerGenl(1) = tmGrf(ilLoop).iPerGenl(1) + 1   'rated ordered
                    tmGrf(ilLoop).iPerGenl(0) = tmGrf(ilLoop).iPerGenl(0) + 1   'rated ordered
                Else
                    If (ilSchRating = igmnfNonRated And tlCntTypes.iNonRAted) Or (ilSchRating = igmnfSuburban And tlCntTypes.iSuburban) Then
                        'tmGrf(ilLoop).iPerGenl(2) = tmGrf(ilLoop).iPerGenl(2) + 1   'non-rated ordered
                        tmGrf(ilLoop).iPerGenl(1) = tmGrf(ilLoop).iPerGenl(1) + 1   'non-rated ordered
                    End If
                End If

            ElseIf (tlSdfExt.sSchStatus = "O" Or tlSdfExt.sSchStatus = "G") Then
                'See if the Missed part of this MG (or OUT) has the same rated status.  Read smf from the code stored in SDF
                ilFound = True
                '10-17-01 Smf reference was picked up from incorrect field (from lmddate to lrecpos)
                'tmSmfSrchKey.lCode = tlSdfExt.lMDDate          'smf code stored in MMDate
                tmSmfSrchKey1.lCode = tlSdfExt.lRecPos

                ilRet = btrGetEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_NONE Then
                    ilFound = False
                Else
                    tmClfSrchKey.lChfCode = tmSmf.lChfCode
                    tmClfSrchKey.iLine = tmSmf.iLineNo
                    tmClfSrchKey.iPropVer = 32000
                    tmClfSrchKey.iCntRevNo = 32000
                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    If tmClf.lChfCode <> tmSmf.lChfCode And tmClf.iLine <> tmSmf.iLineNo Then
                        ilFound = False
                    Else
                        'determine vehicle status for rated vs nonrated
                        'For ilMissedLoop = LBound(tgMVef) To UBound(tgMVef)
                        '    If tgMVef(ilMissedLoop).iCode = tmClf.iVefCode Then      'look for missed vehicles research status (rated/non-rated)
                            ilMissedLoop = gBinarySearchVef(tmClf.iVefCode)
                            If ilMissedLoop <> -1 Then
                                ilMissedRating = tgMVef(ilMissedLoop).iMnfVehGp5Rsch
                        '        Exit For
                            End If
                        'Next ilMissedLoop
                    End If
                End If
                If ilFound Then
                    If (ilSchRating = igmnfRated) Then                'test for rated vehicle
                        'If (tlCnttypes.iRated Or ilSchRating <> ilMissedRating) Then   'always show the cross-overs
                        If (tlCntTypes.iRated) Then
                            'The aired count should gointo the schedule vehicle
                            'tmGrf(ilLoop).iPerGenl(3) = tmGrf(ilLoop).iPerGenl(3) + 1       ' rated aired
                            tmGrf(ilLoop).iPerGenl(2) = tmGrf(ilLoop).iPerGenl(2) + 1       ' rated aired

                        End If
                    Else                  'not rated, must be nonrated
                        If (ilSchRating = igmnfNonRated And tlCntTypes.iNonRAted) Or (ilSchRating = igmnfSuburban And tlCntTypes.iSuburban) Then
                            'If (tlCnttypes.iNonRated Or ilSchRating <> ilMissedRating) Then     'always show the cross-overs
                            'If (tlCntTypes.iNonRated) Then      'always show the cross-overs
                                'tmGrf(ilLoop).iPerGenl(4) = tmGrf(ilLoop).iPerGenl(4) + 1       'non rated aired
                                tmGrf(ilLoop).iPerGenl(3) = tmGrf(ilLoop).iPerGenl(3) + 1       'non rated aired
                            'End If
                        End If
                    End If
                    '

                    'The following is for the ordered count to follow the "Makegood or Outside" spot.
                    'Shadow wants it to reflect what the contract acttually looks like.
                    'commentout for now 10/27/98
                    'If (ilMissedRating = igmnfRated) Then              'test for rated vehicle
                    '    If (tlCnttypes.iRated Or ilSchRating <> ilMissedRating) And (tlSdfExt.sSpotType <> "X") Then
                    '        'The ordered count should go into the missed vehicle
                    '        tmGrf(ilLoop).iPerGenl(1) = tmGrf(ilLoop).iPerGenl(1) + 1       ' rated ordered
                     '    End If
                    'Else                  'not rated, must be nonrated
                    '    If (ilMissedRating = igmnfNonRated) Then
                    '        If (tlCnttypes.iNonRated Or ilSchRating <> ilMissedRating) And (tlSdfExt.sSpotType <> "X") Then
                    '            'the ordered should be from the missed vehicle rating status
                    '            tmGrf(ilLoop).iPerGenl(2) = tmGrf(ilLoop).iPerGenl(2) + 1       'non rated ordered
                    ''        End If
                    '   End If
                    'End If
                End If
            End If
        Else                    ' Missed portion for Mg or Outsides
            'Determine if the Makegood has been processed for this missed.  If so,
            'the ordered has already been counted
            ilFound = False


            'Table of associated makegood references built into memory.  Used to determine
            'if the ordered count has been accumulated with themakegood spot.
            'Comment out 10/27/98.  Show ordered count where missed is, not makegood.
            'For ilMissedLoop = 0 To UBound(lmSDFRecdPos) - 1 Step 1
            '    If lmSDFRecdPos(ilMissedLoop) = tlSdfExt.lRecPos Then
            '        ilfound = True
            '        Exit For
            '    End If
            'Next ilMissedLoop
            gUnpackDateLong tlSdfExt.iDate(0), tlSdfExt.iDate(1), llMissedDate      'convert missed date- add in only if
            'within period reporting
            If (ilMissedRating = igmnfRated) Or (igmnfRated = 0) Then           'missed portion is rated
                If tlSdfExt.sSpotType <> "X" And Not ilFound Then 'exclude fills (extras)
                    'If (tlCnttypes.iRated Or ilSchRating <> ilMissedRating) Then       'always show cross overs
                    If (tlCntTypes.iRated) Then
                        'if missed date is within the period reporting, add it in
                        If llMissedDate >= llStartDate And llMissedDate <= llEndDate Then
                            'tmGrf(ilLoop).iPerGenl(1) = tmGrf(ilLoop).iPerGenl(1) + 1   'rated ordered
                            tmGrf(ilLoop).iPerGenl(0) = tmGrf(ilLoop).iPerGenl(0) + 1   'rated ordered
                        End If
                    End If
                End If
            Else                        'missed portion must be nonrated
                If (ilMissedRating = igmnfNonRated And tlCntTypes.iNonRAted) Or (ilMissedRating = igmnfSuburban And tlCntTypes.iSuburban) Or (tlCntTypes.iNonRAted = True And igmnfNonRated = 0) Or (tlCntTypes.iSuburban = True And igmnfSuburban = 0) Then
                    If tlSdfExt.sSpotType <> "X" And Not ilFound Then
                        'If tlCnttypes.iNonRated Or ilSchRating <> ilMissedRating Then       'always show cross-overs
                        'If tlCntTypes.inonRated Then       'always show cross-overs
                            'if missed date is within the period reporting, add it in
                            If llMissedDate >= llStartDate And llMissedDate <= llEndDate Then
                                'tmGrf(ilLoop).iPerGenl(2) = tmGrf(ilLoop).iPerGenl(2) + 1   'non rated ordered
                                tmGrf(ilLoop).iPerGenl(1) = tmGrf(ilLoop).iPerGenl(1) + 1   'non rated ordered
                            End If
                        'End If
                    End If
                End If
            End If

        End If
    End If                          'ilOkToSeeVeh
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainSpotbyDate               *
'*                                                     *
'*             Created:6/13/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain spots for                *
'*                     specified dates                 *
'*      Note:  fields in tlSdfExt replaced as follows: *
'*             iStatus = original missed vehicle code  *
'*             ilen =    Contr advt code               *
'*      7-29-04 include/exclude contract/feed spot     *
'*******************************************************
Function mObtainSpotByDate(frm As Form, slStartDate As String, slEndDate As String, tlCntTypes As CNTTYPES) As Integer
'
'   ilRet = gObtainSpotbyDate (MainForm, slStartDate, slEndDate)
'   Where:
'       MainForm (I)- Name of Form to unload if error exist
'       slStartDate(I)- Start date  ("" = Start at earliest date)
'       slEndDate(I)- End date  ("" = TFN)
'       ilRet (O)- True if spots obtained OK; False if error
'
    Dim tlSdfExt As SDFEXT
    Dim hlSdf As Integer        'Sdf handle
    Dim ilSdfRecLen As Integer     'Record length
    Dim tlSdf As SDF
    Dim hlSmf As Integer        'Smf handle
    Dim ilSmfRecLen As Integer     'Record length
    Dim tlSmf As SMF
    Dim slStr As String
    Dim llRecPos As Long        'Record location
    Dim llNoRec As Long
    Dim ilRet As Integer

    Dim tlSmfSrchKey As SMFKEY0
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLTypeBuff As POPLCODE   'Type field record
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim slDate As String
    Dim llDate As Long
    Dim ilOffSet As Integer
    Dim ilExtLen As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilIncludeSpot As Integer
    Dim tlSdfSrchKey1 As SDFKEY1
    Dim tlSdfSrchKey3 As LONGKEY0
    Dim ilVef As Integer
    Dim ilChf As Integer
    ReDim tmGrf(0 To 0) As GRF
    'ReDim tlMnf(0 To 0) As MNF
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilClf As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilCount As Integer
    hlSdf = hmSdf
    hlSmf = hmSmf
    ReDim lmSDFRecdPos(0 To 0) As Long
    If slStartDate <> "" Then
        llDate = gDateValue(slStartDate)
        slDate = Trim$(str$(llDate))
        Do While Len(slDate) < 6
            slDate = "0" & slDate
        Loop
    Else
        slDate = "000000"
    End If
    slStr = slDate
    If slEndDate <> "" Then
        llDate = gDateValue(slEndDate)
        slDate = Trim$(str$(llDate))
        Do While Len(slDate) < 6
            slDate = "0" & slDate
        Loop
    Else
        slDate = "000000"
    End If
    llStartDate = gDateValue(slStartDate)
    llEndDate = gDateValue(slEndDate)
    mObtainSpotByDate = True
    ilSmfRecLen = Len(tmSmf) 'btrRecordLength(hlSmf)  'Get and save record length
    ilSdfRecLen = Len(tlSdf)
    ilRet = gObtainVef()
    'Gather the vehicle groups and detrmine rated vs non-rated
    'ilRet = gObtainMnfForType("H", slStamp, tlMnf())
    'For ilLoop = LBound(tlMnf) To UBound(tlMnf) Step 1
    '    If Val(tlMnf(ilLoop).sUnitType) = 5 Then
    '        If Trim$(tlMnf(ilLoop).sName) = "Rated" Then
    ''            ilmnfRated = tlMnf(ilLoop).icode
    '        ElseIf Trim$(tlMnf(ilLoop).sName) = "Non-Rated" Then
    '            ilmnfnonRated = tlMnf(ilLoop).icode
    '        End If
    '    End If
   'Next ilLoop


    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
        If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "S") Then
            ilCount = 0             'debug
            ilExtLen = 2 + 4 + 2 + 4 + 4 + 1 + 1 + 2 + 1 + 1 + 1 + 2 + 2 + 1 + 4 + 4 + 4 'VefCode+ChfCode+LineNo+Date+Time+Status+Trace+Length+PriceType+SpotType+Bill+Adf+GameNo+Midnight+Code+... Len(tlSdfExt(1)) - 9'Extract operation record size
            llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
            btrExtClear hlSdf   'Clear any previous extend operation
            tlSdfSrchKey1.iVefCode = tgMVef(ilVef).iCode
            If slStartDate = "" Then
                slDate = "1/1/1970"
            Else
                slDate = slStartDate
            End If
            gPackDate slDate, tlSdfSrchKey1.iDate(0), tlSdfSrchKey1.iDate(1)
            tlSdfSrchKey1.iTime(0) = 0
            tlSdfSrchKey1.iTime(1) = 0
            tlSdfSrchKey1.sSchStatus = " "
            ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_KEY_NOT_FOUND) Then
                If (ilRet <> BTRV_ERR_NONE) Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                Call btrExtSetBounds(hlSdf, llNoRec, -1, "UC", "SDFEXTPK_RPT", SDFEXTPK_RPT) 'Set extract limits (all records)

                '7-19-04 determine to include local, network (feed) or both
                tlLTypeBuff.lCode = 0
                ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                If tlCntTypes.iCntrSpots = True And tlCntTypes.iFeedSpots = False Then        'include local, exclude network (feed)
                    ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_GT, BTRV_EXT_AND, tlLTypeBuff, 4)
                ElseIf tlCntTypes.iCntrSpots = False And tlCntTypes.iFeedSpots = True Then   'exclude local, include feed
                    ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLTypeBuff, 4)
                End If
                tlIntTypeBuff.iType = tgMVef(ilVef).iCode
                ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
                ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                If slStartDate = "" Then
                    slDate = "1/1/1970"
                Else
                    slDate = slStartDate
                End If
                gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                ilOffSet = gFieldOffset("Sdf", "SdfDate")
                ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
                If slEndDate = "" Then
                    slDate = "12/31/2069"
                Else
                    slDate = slEndDate
                End If
                gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                ilOffSet = gFieldOffset("Sdf", "SdfDate")
                ilRet = btrExtAddLogicConst(hlSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 2)  'Extract Name
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                ilOffSet = gFieldOffset("Sdf", "SdfChfCode")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 4)  'Extract iCode field
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                ilOffSet = gFieldOffset("Sdf", "SdfLineNo")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 2)  'Extract Name
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                ilOffSet = gFieldOffset("Sdf", "SdfDate")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 4) 'Extract Variation
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                ilOffSet = gFieldOffset("Sdf", "SdfTime")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 4) 'Extract Variation
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                ilOffSet = gFieldOffset("Sdf", "SdfSchStatus")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 1) 'Extract Variation
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                ilOffSet = gFieldOffset("Sdf", "SdfTracer")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 1) 'Extract Variation
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If


                ilOffSet = gFieldOffset("Sdf", "SdfAdfCode")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 2) 'Extract advt code instead of length
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                'ilOffset = gFieldOffset("Sdf", "SdfLen")
                'ilRet = btrExtAddField(hlSdf, ilOffset, 2) 'Extract Variation
                'If ilRet <> BTRV_ERR_NONE Then
                '    mObtainSpotByDate = False
                '    ilRet = btrClose(hlSdf)
                '    ilRet = btrClose(hlSmf)
                '    btrDestroy hlSdf
                '    btrDestroy hlSmf
                '    Exit Function
                'End If
                ilOffSet = gFieldOffset("Sdf", "SdfPriceType")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 1) 'Extract Variation
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                ilOffSet = gFieldOffset("Sdf", "SdfSpotType")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 1) 'Extract Variation
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                ilOffSet = gFieldOffset("Sdf", "SdfBill")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 1) 'Extract Variation
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                ilOffSet = gFieldOffset("Sdf", "SdfAdfCode")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 2) 'Extract advt code instead of length
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                ilOffSet = gFieldOffset("Sdf", "SdfGameNo")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 2) 'Extract advt code instead of length
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                
                '5/22/14: Add cross midnight
                ilOffSet = gFieldOffset("Sdf", "SdfXCrossMidnight")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 1) 'Extract Variation
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                
                ilOffSet = gFieldOffset("Sdf", "SdfCode")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 4) 'Extract Variation
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If
                ilOffSet = gFieldOffset("Sdf", "SDFSMFCODE")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 4) 'Extract Variation
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If

                ilOffSet = gFieldOffset("Sdf", "SDFFSFCODE")
                ilRet = btrExtAddField(hlSdf, ilOffSet, 4) 'Extract Variation
                If ilRet <> BTRV_ERR_NONE Then
                    mObtainSpotByDate = False
                    ilRet = btrClose(hlSdf)
                    ilRet = btrClose(hlSmf)
                    btrDestroy hlSdf
                    btrDestroy hlSmf
                    Exit Function
                End If

                ilRet = btrExtGetNext(hlSdf, tlSdfExt, ilExtLen, llRecPos)
                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                    If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                        mObtainSpotByDate = False
                        ilRet = btrClose(hlSdf)
                        ilRet = btrClose(hlSmf)
                        btrDestroy hlSdf
                        btrDestroy hlSmf
                        Exit Function
                    End If
                    ilExtLen = 2 + 4 + 2 + 4 + 4 + 1 + 1 + 2 + 1 + 1 + 1 + 2 + 2 + 1 + 4 + 4 + 4 'VefCode+ChfCode+LineNo+Date+Time+Status+Tracer+Length+PriceType+SpotType+Bill+Adf+GameNo+Midnight+Code+... Len(tlSdfExt(1)) - 4'Extract operation record size
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hlSdf, tlSdfExt, ilExtLen, llRecPos)
                    Loop
                    Do While ilRet = BTRV_ERR_NONE And tlSdfExt.iVefCode = tgMVef(ilVef).iCode
                        'Test if spot previously added- if so ignore (virtual then one virtual vehicle for contract, and
                        'the virtual vehicles book into same vehicles)
                        'ilFound = False
                        'For ilTest = LBound(tlSdfEXT) To UBound(tlSdfEXT) - 1 Step 1
                        '    If tlSdfEXT(ilTest).lRecPos = llRecPos Then
                        '        ilFound = True
                        '        Exit For
                        '    End If
                        'Next ilTest
                        'Build gRf from the SDFbuffer
                        '3-30-05 always ignore all BB spots

                        If tlSdfExt.sSpotType <> "S" And tlSdfExt.sSpotType <> "M" And tlSdfExt.sSpotType <> "O" And tlSdfExt.sSpotType <> "C" Then    'always ignore psas & promos
                            'test for selective advt

                             ilFound = True
                             If Not gSetCheck(RptSelAS!ckcAll.Value) Then
                                ilFound = False
                                For ilLoop = 0 To RptSelAS!lbcSelection(0).ListCount - 1 Step 1
                                    If RptSelAS!lbcSelection(0).Selected(ilLoop) Then
                                        slNameCode = tgRptSelAdvertiserCode(ilLoop).sKey
                                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                        If Val(slCode) = tlSdfExt.iLen Then
                                            ilFound = True
                                            Exit For
                                        End If
                                    End If
                                Next ilLoop
                            End If
                            If ilFound Then
                                'Build table of mg or outsides that have been processed (store the SDF record # only)
                                If (tlSdfExt.sSchStatus = "O" Or tlSdfExt.sSchStatus = "G") Then
                                    lmSDFRecdPos(UBound(lmSDFRecdPos)) = tlSdfExt.lRecPos
                                    ReDim Preserve lmSDFRecdPos(0 To UBound(lmSDFRecdPos) + 1)
                                End If
                                ilCount = ilCount + 1
                                mCreateGRFbyCnt tlSdfExt, tlCntTypes, llStartDate, llEndDate, 0
                            End If
                        End If
                        ilRet = btrExtGetNext(hlSdf, tlSdfExt, ilExtLen, llRecPos)
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hlSdf, tlSdfExt, ilExtLen, llRecPos)
                        Loop
                    Loop                'while ilret = btrv_err_none
                End If                  'SDF btrExtGetNext- If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT)
            End If                      'SDF btrGetGreaterOrEqual-If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_KEY_NOT_FOUND)
        End If
    Next ilVef
    'Get any missed spots of MG
    If tlCntTypes.iCntrSpots = True Then
        ilRet = gObtainCntrForDate(frm, slStartDate, slEndDate, "HO", "", 1, tmChfAdvtExt())
        If ilRet = 0 Then
            For ilChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
                ilRet = gObtainChfClf(hmCHF, hmClf, tmChfAdvtExt(ilChf).lCode, False, tgChfAS, tgClfAS())
                ilExtLen = Len(tlSmf)
                llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
                btrExtClear hlSmf   'Clear any previous extend operation
                tlSmfSrchKey.lChfCode = tmChfAdvtExt(ilChf).lCode
                tmSmfSrchKey.lFsfCode = 0
                tlSmfSrchKey.iLineNo = 0
                tlSmfSrchKey.iMissedDate(0) = 0
                tlSmfSrchKey.iMissedDate(1) = 0
                ilRet = btrGetGreaterOrEqual(hlSmf, tlSmf, ilSmfRecLen, tlSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If ilRet <> BTRV_ERR_END_OF_FILE Then
                    If ilRet <> BTRV_ERR_NONE Then
                        Erase tmChfAdvtExt
                        btrDestroy hlSmf
                        btrDestroy hlSdf
                        mObtainSpotByDate = True
                        Exit Function
                    End If
                    Call btrExtSetBounds(hlSmf, llNoRec, -1, "UC", "SMF", "") 'Set extract limits (all records)
                    tlLTypeBuff.lCode = tmChfAdvtExt(ilChf).lCode
                    ilOffSet = gFieldOffset("Smf", "SmfChfCode")
                    ilRet = btrExtAddLogicConst(hlSmf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLTypeBuff, 4)
                    If slStartDate = "" Then
                        slDate = "1/1/1970"
                    Else
                        slDate = slStartDate
                    End If
                    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Smf", "SMFMISSEDDATE")
                    ilRet = btrExtAddLogicConst(hlSmf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
                    If slEndDate = "" Then
                        slDate = "12/31/2069"
                    Else
                        slDate = slEndDate
                    End If
                    gPackDate slDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
                    ilOffSet = gFieldOffset("Smf", "SMFMISSEDDATE")
                    ilRet = btrExtAddLogicConst(hlSmf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                    ilRet = btrExtAddField(hlSmf, 0, ilExtLen)  'Extract Name
                    If ilRet <> BTRV_ERR_NONE Then
                        Erase tmChfAdvtExt
                        btrDestroy hlSmf
                        btrDestroy hlSdf
                        mObtainSpotByDate = True
                        Exit Function
                    End If
                    ilRet = btrExtGetNext(hlSmf, tlSmf, ilExtLen, llRecPos)
                    If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                        If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                            Erase tmChfAdvtExt
                            btrDestroy hlSmf
                            btrDestroy hlSdf
                            mObtainSpotByDate = True
                            Exit Function
                        End If
                        ilExtLen = Len(tlSmf)
                        'ilRet = btrExtGetFirst(hlSdf, tlSdfExt, ilExtLen, llRecPos)
    
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hlSmf, tlSmf, ilExtLen, llRecPos)
                        Loop
                        Do While ilRet = BTRV_ERR_NONE
                            'examine Outsides and makegoods that are not psas or promos
'                            If ((tlSmf.sSchStatus = "O") Or (tlSmf.sSchStatus = "G")) And (tlSdfExt.sSpotType <> "S" And tlSdfExt.sSpotType <> "M") Then
                            If ((tlSmf.sSchStatus = "O") Or (tlSmf.sSchStatus = "G")) Then
                                ilIncludeSpot = True
                                tlSdfSrchKey3.lCode = tlSmf.lSdfCode
                                ilRet = btrGetEqual(hlSdf, tlSdf, ilSdfRecLen, tlSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                If ilRet <> BTRV_ERR_NONE Then
                                    ilIncludeSpot = False
                                Else        'ignore if psa, promo, open, close and the sched vehicle is the same as the orig vehicle
                                
                                    If (tlSdf.sSpotType <> "S" And tlSdf.sSpotType <> "M" And tlSdf.sSpotType <> "O" And tlSdf.sSpotType <> "C") Then
                                        ilFound = True
                                        If Not gSetCheck(RptSelAS!ckcAll.Value) Then
                                            ilFound = False
                                            If Not gSetCheck(RptSelAS!ckcAll.Value) Then
                                               For ilLoop = 0 To RptSelAS!lbcSelection(0).ListCount - 1 Step 1
                                                   If RptSelAS!lbcSelection(0).Selected(ilLoop) Then
                                                       slNameCode = tgRptSelAdvertiserCode(ilLoop).sKey
                                                       ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                                       If Val(slCode) = tlSdf.iAdfCode Then
                                                           ilFound = True
                                                       End If
                                                   End If
                                               Next ilLoop
                                            End If
                                        End If
                                        If Not ilFound Then
                                            ilIncludeSpot = False
                                        End If
                                    Else
                                        'ignore psa,promo, fill, open, close
                                        ilIncludeSpot = False
                                    End If
                                End If                  'btrv_err_none
                            Else
                                ilIncludeSpot = False
                            End If
                            If ilIncludeSpot Then
                                tlSdfExt.iVefCode = tlSdf.iVefCode     'Vehicle Code where mg schedule (missed vehicle may be different) (combos not allowed)
                                tlSdfExt.lChfCode = tlSdf.lChfCode       'Contract code
                                tlSdfExt.iLineNo = tlSdf.iLineNo       'Contract code
                                tlSdfExt.iAdfCode = tlSdf.iAdfCode       'advertisre code
                                tlSdfExt.iDate(0) = tlSmf.iMissedDate(0)    'missed Date of spot
                                tlSdfExt.iDate(1) = tlSmf.iMissedDate(1)    'missed Date of spot
                                tlSdfExt.iTime(0) = tlSmf.iMissedTime(0)    'missed time of spot
                                tlSdfExt.iTime(1) = tlSmf.iMissedTime(1)    'missed time of spot
                                tlSdfExt.sSchStatus = tlSdf.sSchStatus    'S=Scheduled, M=Missed, R=Ready to schd MG, U=Unscheduled MG,
                                tlSdfExt.sTracer = tlSdf.sTracer
                                tlSdfExt.iLen = tlSdf.iLen         'Spot length
                                tlSdfExt.sPriceType = tlSdf.sPriceType
                                tlSdfExt.sSpotType = tlSdf.sSpotType
                                  tlSdfExt.lMdDate = 0    'Missed date for MG's (used in gCntrDisp)
                                tlSdfExt.lRecPos = tlSmf.lSdfCode
                                tlSdfExt.iGameNo = tlSdf.iGameNo
                                'get the vehicle the spot was originally missed from
                                'tlSdfExt.iStatus = 0     '0=From Sdf; 1= from Smf
                                'replace .iStatus field with vehicle the spot was originally missed from
                                For ilClf = LBound(tgClfAS) To UBound(tgClfAS) - 1
                                    If tlSdf.iLineNo = tgClfAS(ilClf).ClfRec.iLine Then
                                        tlSdfExt.iStatus = tgClfAS(ilClf).ClfRec.iVefCode
                                        Exit For
                                    End If
                                Next ilClf

                                'Build gRf from the SDFbuffer
                                mCreateGRFbyCnt tlSdfExt, tlCntTypes, llStartDate, llEndDate, 1
                            End If
                            ilRet = btrExtGetNext(hlSmf, tlSmf, ilExtLen, llRecPos)
                            Do While ilRet = BTRV_ERR_REJECT_COUNT
                                ilRet = btrExtGetNext(hlSmf, tlSmf, ilExtLen, llRecPos)
                            Loop
                        Loop
                    End If
                End If
            Next ilChf
        End If                      'if tlcnttypes.icntrspots = true
    End If
    Erase tmChfAdvtExt
    'Erase tlMnf
    mObtainSpotByDate = True
    Exit Function
End Function
