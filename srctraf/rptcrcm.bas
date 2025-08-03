Attribute VB_Name = "RPTCRCM"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcros.bas on Fri 3/12/10 @ 11:00 AM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
'Public igYear As Integer                'budget year used for filtering
'Public igMonthOrQtr As Integer          'entered month or qtr
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer            'Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim tmClfSrchKey As CLFKEY0     'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey2 As SDFKEY2            'SDF record image (key 2)
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdf As SDF
Dim tmSdfSrchKey3 As LONGKEY0     'SDF record image (SDF code as keyfield)

Dim hmSmf As Integer
Dim tmSmf As SMF
Dim imSmfRecLen As Integer

Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim tmRcf As RCF
Dim imRcfCode As Integer        'selected Rate CArd
Dim hmRdf As Integer            'Dayparts file handle
Dim imRdfRecLen As Integer      'RD record length
Dim tmRdfSrchKey As INTKEY0     'RDF key image
Dim tmRdf As RDF

Dim hmFsf As Integer            'Feed file handle
Dim tmFSFSrchKey As LONGKEY0    'FSF search key
Dim imFsfRecLen As Integer      'FSF record length
Dim tmFsf As FSF                'FSF record buffer
Dim hmAnf As Integer            'Named avail file handle
Dim tmAnfSrchKey As INTKEY0    'ANF record image
Dim imAnfRecLen As Integer      'ANF record length
Dim tmAnf As ANF

'Log Calendar
Dim hmLcf As Integer            'Log Calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim imLcfRecLen As Integer        'LCF record length

Dim hmSsf As Integer            'Spot Summary file handle
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS

Dim tmAvRdf() As RDF            'array of dayparts to process
Dim tmBaseCounts() As DPCOUNTS      'competitives within base DPs
Dim tmNoBAseCounts() As DPCOUNTS    'competitives without a base dp
Dim tmSpotTypes As SPOTTYPES

Dim imInclProdCodes As Integer  'include or exclude flag
Dim imUseProdCodes() As Integer 'prod protection codes to include/exclude

Type DPCOUNTS
    sKey As String * 20         'DP key (10) | dp count(5)
    iVefCode As Integer         'vehicle code
    iRdfCode As Integer         'daypart code
    iRcfCode As Integer         'Rate CArd Code
    iSort As Integer            'sort #
    iCCMnfCode As Integer       'mnf product protection code
    'iCCPeriod(1 To 13) As Integer   'weeks 1 - 13
    iCCPeriod(0 To 13) As Integer   'weeks 1 - 13. Index zero ignored
End Type
'*******************************************************************
'*                                                                 *
'*      Procedure Name:gCRQtrlyBookSpots                           *
'*                                                                 *
'*             Created:12/29/97      By:D. Hosaka                  *
'*            Modified:              By:                           *
'*                                                                 *
'*            Comments: Generate Competitive Report                   *
'*                                                                 *
'*      3/11/98 Look at "Base DP" only, (not Report DP)            *
'*      4/12/98 Remove duplication of spots from vehicle           *
'*              These spots appeared to be moved across vehicles
'
'
'*                                                                 *
'*******************************************************************
Sub gCreateCompetitiveCats()
'
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slDate As String
    ReDim ilDate(0 To 1) As Integer
    Dim ilVehicle As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilDay As Integer
    Dim ilVpfIndex As Integer
    Dim ilUpper As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilRcf As Integer
    Dim ilFound As Integer
    Dim ilNoWks As Integer                      'user # weeks requested
    Dim tlCntTypes As CNTTYPES
    Dim ilSaveSort As Integer                   'DP or RIF field:  sort code
    Dim ilMajorSet As Integer                      'vehicle sort group
    Dim ilMinorSet As Integer                   'minor vehicle group (not used)
    Dim ilMnfMajorCode As Integer               'vehicle group mnf code
    Dim ilmnfMinorCode As Integer               'minor MNF code (not used)
    Dim ilRateCardOK As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim ilStartDate(0 To 1) As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilLoopOnCount As Integer
    Dim ilTemp As Integer
    Dim slKey As String
    Dim slStr As String
    Dim llTemp As Long
    Dim ilFirst As Integer
    Dim ilRdfCode As Integer
    
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
        btrDestroy hmCHF
    Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)
    
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imVsfRecLen = Len(tmVsf)
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmRdf
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imRdfRecLen = Len(tmRdf)

    hmFsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmFsf
        btrDestroy hmRdf
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imFsfRecLen = Len(tmFsf)

    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAnf
        btrDestroy hmFsf
        btrDestroy hmRdf
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "SSf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSsf
        btrDestroy hmAnf
        btrDestroy hmFsf
        btrDestroy hmRdf
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imSsfRecLen = Len(tmSsf)
    
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmLcf
        btrDestroy hmSsf
        btrDestroy hmAnf
        btrDestroy hmFsf
        btrDestroy hmRdf
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imLcfRecLen = Len(tmLcf)

    tlCntTypes.iHold = gSetCheck(RptSelCM!ckcCType(0).Value)
    tlCntTypes.iOrder = gSetCheck(RptSelCM!ckcCType(1).Value)
    tlCntTypes.iNetwork = gSetCheck(RptSelCM!ckcCType(2).Value)
    tlCntTypes.iStandard = gSetCheck(RptSelCM!ckcCType(3).Value)
    tlCntTypes.iReserv = gSetCheck(RptSelCM!ckcCType(4).Value)
    tlCntTypes.iRemnant = gSetCheck(RptSelCM!ckcCType(5).Value)
    tlCntTypes.iDR = gSetCheck(RptSelCM!ckcCType(6).Value)
    tlCntTypes.iPI = gSetCheck(RptSelCM!ckcCType(7).Value)
    tlCntTypes.iPSA = gSetCheck(RptSelCM!ckcCType(8).Value)
    tlCntTypes.iPromo = gSetCheck(RptSelCM!ckcCType(9).Value)
    tlCntTypes.iTrade = gSetCheck(RptSelCM!ckcCType(10).Value)
    tlCntTypes.iMissed = gSetCheck(RptSelCM!ckcSpots(0).Value)
    tlCntTypes.iCharge = gSetCheck(RptSelCM!ckcSpots(1).Value)
    tlCntTypes.iZero = gSetCheck(RptSelCM!ckcSpots(2).Value)
    tlCntTypes.iADU = gSetCheck(RptSelCM!ckcSpots(3).Value)
    tlCntTypes.iBonus = gSetCheck(RptSelCM!ckcSpots(4).Value)
    tlCntTypes.iXtra = gSetCheck(RptSelCM!ckcSpots(5).Value)
    tlCntTypes.iFill = gSetCheck(RptSelCM!ckcSpots(6).Value)
    tlCntTypes.iNC = gSetCheck(RptSelCM!ckcSpots(7).Value)
    tlCntTypes.iRecapturable = gSetCheck(RptSelCM!ckcSpots(8).Value)
    tlCntTypes.iSpinoff = gSetCheck(RptSelCM!ckcSpots(9).Value)
    tlCntTypes.iMG = gSetCheck(RptSelCM!ckcSpots(10).Value)     'MG option are MG schedule line rates & also MG/outside spot types
    'SEtup filters for the routine to access all spots from vehicle (gGetSpotsbyVefDate)
    If tlCntTypes.iMissed Then              'include Missed
        tmSpotTypes.iCancel = True
        tmSpotTypes.iMissed = True
    Else
        tmSpotTypes.iCancel = False
        tmSpotTypes.iMissed = False
    End If
    tmSpotTypes.iHidden = False             'always exclude hidden spots
    tmSpotTypes.iFill = True
    tmSpotTypes.iSched = True
    tmSpotTypes.iMG = True
    tmSpotTypes.iOutside = True
    
    If (tlCntTypes.iHold) Or (tlCntTypes.iOrder) Then        '1-26-05 set general cntr type for inclusion/exclusion if hold or ordered selected
        tlCntTypes.iCntrSpots = True
    Else
        tlCntTypes.iCntrSpots = False
    End If
    
    If (Not tlCntTypes.iMG) Then            'exclude MGs & outsides?
        tmSpotTypes.iMG = False
        tmSpotTypes.iOutside = False
    End If

    'Get the vehicle group selected for sorting
    ilRet = RptSelCM!cbcGroup.ListIndex
    ilMajorSet = gFindVehGroupInx(ilRet, tgVehicleSets1())
    ilRet = gObtainRcfRifRdf()          'get the rate cards and assoc dayparts

    'get all the dates needed to work with
    slDate = RptSelCM!edcSelCFrom.Text               'effective date entred
    llStartDate = gDateValue(slDate)
    'backup to Monday
    ilDay = gWeekDayLong(llStartDate)
    Do While ilDay <> 0
        llStartDate = llStartDate - 1
        ilDay = gWeekDayLong(llStartDate)
    Loop
    
    gPackDateLong llStartDate, ilStartDate(0), ilStartDate(1)
    ilNoWks = Val(RptSelCM!edcSelCFrom1.Text)
    llEndDate = llStartDate + ((ilNoWks) * 7) - 1    'get end of last week
    
    'get the prod prot codes selected
    gObtainCodesForMultipleLists 2, tgMnfCodeCT(), imInclProdCodes, imUseProdCodes(), RptSelCM

    tmVef.iCode = 0
    For ilVehicle = 0 To RptSelCM!lbcSelection(0).ListCount - 1 Step 1
        If (RptSelCM!lbcSelection(0).Selected(ilVehicle)) Then
            slNameCode = tgCSVNameCode(ilVehicle).sKey 'RptSelSP!lbcCSVNameCode.List(ilVehicle)
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            ilVpfIndex = -1
            ilLoop = gBinarySearchVpf(ilVefCode)
            If ilLoop <> -1 Then
                ilVpfIndex = ilLoop
            End If
            If ilVpfIndex >= 0 Then
                For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                    tmRcf = tgMRcf(ilRcf)
                    ilRateCardOK = False
                    For ilLoop = 0 To RptSelCM!lbcSelection(1).ListCount - 1 Step 1
                        slNameCode = tgRateCardCode(ilLoop).sKey
                        ilRet = gParseItem(slNameCode, 3, "\", slCode)
                        If Val(slCode) = tgMRcf(ilRcf).iCode Then
                            If (RptSelCM!lbcSelection(1).Selected(ilLoop)) Then
                                ilRateCardOK = True
                                imRcfCode = Val(slCode)
                            End If
                            Exit For
                        End If
                    Next ilLoop

                    If ilRateCardOK Then
                        ReDim tmAvRdf(0 To 0) As RDF
                        ilUpper = 0
                        'Setup the DP with rates to process
                        For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                            If tgMRif(llRif).iRcfCode = tgMRcf(ilRcf).iCode Then
                                'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                                If ilRdf <> -1 Then
                                    'determine if rate card items have sort codes, etc; otherwise use DP
                                    If tgMRif(llRif).sBase = "Y" Then
                                        If tgMRif(llRif).iSort = 0 Then
                                            ilSaveSort = tgMRdf(ilRdf).iSortCode
                                        Else
                                            ilSaveSort = tgMRif(llRif).iSort
                                        End If
                                        If tgMRdf(ilRdf).iCode = tgMRif(llRif).iRdfCode And tgMRdf(ilRdf).sState <> "D" And tgMRif(llRif).iVefCode = ilVefCode Then
                                            ilFound = False
                                            For ilLoop = LBound(tmAvRdf) To ilUpper - 1 Step 1
                                                If tmAvRdf(ilLoop).iCode = tgMRdf(ilRdf).iCode Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilLoop
                                            If Not ilFound Then
                                                tmAvRdf(ilUpper) = tgMRdf(ilRdf)
                                                tmAvRdf(ilUpper).iSortCode = ilSaveSort         'replace the RIF sort code in table for Crystal
                                                ilUpper = ilUpper + 1
                                                ReDim Preserve tmAvRdf(0 To ilUpper) As RDF
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next llRif
                        ReDim tlSdf(0 To 0) As SDF
                        mGetSpotCountsCM ilVefCode, llStartDate, llEndDate, tmAvRdf(), tlCntTypes, tlSdf()
                        
                        'if top how many is blank or 100, dont need to sort as the entire list is generated
                        If Trim$(RptSelCM!edcHowMany.Text) = "" Then
                            ilTemp = 100
                        Else
                            ilTemp = Val(RptSelCM!edcHowMany.Text)
                        End If
                        'sort the array of dayparts & product prot codes desc by prod protection within daypart (a vehicle at a time)
                        If UBound(tmBaseCounts) - 1 > 1 Or (ilTemp > 0 Or ilTemp < 100) Then        'no need to sort if requesting all categories
                            For ilLoop = LBound(tmBaseCounts) To UBound(tmBaseCounts) - 1
                                slStr = Trim$(str$(tmBaseCounts(ilLoop).iRdfCode))
                                Do While Len(slStr) < 9
                                    slStr = "0" & slStr
                                Loop
                                slKey = slStr & "|"
                                llTemp = 99999 - tmBaseCounts(ilLoop).iCCPeriod(1)    '1st period is the top down column
                                slStr = Trim$(str$(llTemp))
                                Do While Len(slStr) < 5
                                    slStr = "0" & slStr
                                Loop
                                slKey = slKey & slStr
                                tmBaseCounts(ilLoop).sKey = slKey
                            Next ilLoop
                            ArraySortTyp fnAV(tmBaseCounts(), 0), UBound(tmBaseCounts) - 1, 0, LenB(tmBaseCounts(0)), 0, LenB(tmBaseCounts(0).sKey), 0
                        End If

                        ilFirst = True
                        
'                        ilLoopOnCount = 0
'                        Do While ilLoopOnCount < UBound(tmBaseCounts)
'
'                        '      10/6/10 grfFields
'                        '      GenDate - Generation date
'                        '      GenTime - Generation Time
'                        '      VefCode - vehicle code
'                        '      rdfCode - DP Code
'                        '      sofcode - competitive mnf code
'                        '      Code2 - mnf vehicle group
'                        '      StartDate - Date of Competitive data
'                        '      PerGenl(1-13) = counts, wk 1-13
'                        'Setup the major sort factor
'                            If ilFirst Then
'                                ilFirst = False
'                                ilRdfCode = tmBaseCounts(ilLoopOnCount).iRdfcode
'                                ilLoop = ilTemp          'count for top how many
'                            End If
'                            If ilRdfCode = tmBaseCounts(ilLoopOnCount).iRdfcode Then
'                                ilLoop = ilLoop - 1
'                                If ilLoop >= 0 Then
'                                    mUpdateGRFCount ilStartDate(), ilMajorSet, tmBaseCounts(ilLoopOnCount)
'                                End If
'                            Else
'                                ilRdfCode = tmBaseCounts(ilLoopOnCount).iRdfcode
'                                ilLoop = ilTemp - 1                  'reset count for top how many
'                                mUpdateGRFCount ilStartDate(), ilMajorSet, tmBaseCounts(ilLoopOnCount)
'                            End If
'                            ilLoopOnCount = ilLoopOnCount + 1
'                        Loop
                        
                        For ilLoopOnCount = LBound(tmBaseCounts) To UBound(tmBaseCounts) - 1
                            mUpdateGRFCount ilStartDate(), ilMajorSet, tmBaseCounts(ilLoopOnCount)
'                            If ilFirst Then
'                                ilFirst = False
'                                ilRdfCode = tmBaseCounts(ilLoopOnCount).iRdfcode
'                                ilLoop = ilTemp          'count for top how many
'                            End If
'                            If ilRdfCode = tmBaseCounts(ilLoopOnCount).iRdfcode Then
'                                ilLoop = ilLoop - 1
'                                If ilLoop >= 0 Then
'                                    mUpdateGRFCount ilStartDate(), ilMajorSet, tmBaseCounts(ilLoopOnCount)
'                                End If
'                            Else
'                                ilRdfCode = tmBaseCounts(ilLoopOnCount).iRdfcode
'                                ilLoop = ilTemp - 1                  'reset count for top how many
'                                mUpdateGRFCount ilStartDate(), ilMajorSet, tmBaseCounts(ilLoopOnCount)
'                            End If
                        Next ilLoopOnCount
                       
                        'process the set of categories that dont belong in any Base DP to process
                        For ilLoopOnCount = 0 To UBound(tmNoBAseCounts) - 1 Step 1
                            
                            'For now, dont do anything with the spots not in any base dp
                            'mUpdateGRFCount ilStartDate(), ilMajorSet, tmNoBAseCounts(ilLoopOnCount)
                            
                        Next ilLoopOnCount
                       
                    End If
                    If ilRateCardOK Then
                        Exit For
                    End If
                Next ilRcf                      'next rate card (should only be 1)
            End If                              'vpfindex > 0
        End If                                  'vehicle selected
        Next ilVehicle                              'For ilvehicle = 0 To RptSelCM!lbcSelection(0).ListCount - 1
    Erase tmAvRdf, tlSdf, tmBaseCounts, tmNoBAseCounts
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmFsf)
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmLcf)
    btrDestroy hmRdf
    btrDestroy hmSdf
    btrDestroy hmVef
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmMnf
    btrDestroy hmFsf
    btrDestroy hmAnf
    btrDestroy hmSsf
    btrDestroy hmLcf
    Exit Sub
End Sub
'*****************************************************************
'*                                                               *
'*                                                               *
'*                                                               *
'*      Procedure Name:mGetSpotCounts
'       <input> ilVefCode - vehicle                              *
'               llSDate - earliest Monday start date
'               llEDate - latest Sunday end date
'               tlAvRdf() - array of base dayparts for selected RC
'               tlCntTypes - inclusion/exclusion of contr types
'               tlSpottypes - Inclusion/exclusion of spot types
'       <output> tlSdf() - array of spots for 1 vehicle
'*      Gather spot counts by base daypart and vehicle           *
'*      Created: 10-6-10       By:D. Hosaka                      *
'*                                                               *
'*****************************************************************
Sub mGetSpotCountsCM(ilVefCode As Integer, llSDate As Long, llEDate As Long, tlAvRdf() As RDF, tlCntTypes As CNTTYPES, tlSdf() As SDF)
Dim slStartDate As String
Dim slEndDate As String
Dim slDate As String
Dim llLoopOnSpot As Long
Dim ilSpotOK As Integer
Dim ilRdf As Integer
Dim ilLoop As Integer
Dim ilFoundDP As Integer
Dim ilFoundPP As Integer
Dim llTime As Long
Dim ilLoopOnCounts As Integer
Dim llSpotDate As Long
Dim ilLoopIndex As Integer
Dim ilWkInx As Integer
Dim ilUpper As Integer
Dim llStartTime As Long
Dim llEndTime As Long
Dim ilDay As Integer
Dim llDate As Long
Dim ilDate0 As Integer
Dim ilDate1 As Integer
Dim llLoopDate As Long
Dim llLatestDate As Long
Dim ilSpot As Integer
Dim ilVefIndex As Integer
Dim ilType As Integer
Dim ilRet As Integer
Dim ilWeekDay As Integer
Dim ilIndex As Integer
Dim ilEvt As Integer
Dim slKey As String
Dim slStr As String
Dim llTemp As Long
Dim ilRdfCode As Integer
Dim ilEvtType(0 To 14) As Integer

        slStartDate = Format$(llSDate, "m/d/yy")
        slEndDate = Format$(llEDate, "m/d/yy")
        ReDim tmBaseCounts(0 To 0) As DPCOUNTS 'competitive code counts for one vehicle, as many base dayparts found in rate card
        ReDim tmNoBAseCounts(0 To 0) As DPCOUNTS 'competitive code counts for one vehicle, spot doesnt fall in a base DP
        
        ilVefIndex = gBinarySearchVef(ilVefCode)
 
        llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
        'set the type of events to get fro the day (only Contract avails)
        For ilLoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
            ilEvtType(ilLoop) = False
        Next ilLoop
        ilEvtType(2) = True                 'avails only
        ilType = 0
        'Gather all spots for the vehicle
        For llLoopDate = llSDate To llEDate
            slDate = Format$(llLoopDate, "m/d/yy")
            gPackDate slDate, ilDate0, ilDate1
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            If tgMVef(ilVefIndex).sType <> "G" Then
                tmSsfSrchKey.iType = ilType
                tmSsfSrchKey.iVefCode = ilVefCode
                tmSsfSrchKey.iDate(0) = ilDate0
                tmSsfSrchKey.iDate(1) = ilDate1
                tmSsfSrchKey.iStartTime(0) = 0
                tmSsfSrchKey.iStartTime(1) = 0
                ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            Else
                '5-25-06 Change to use SSFKey2.  When there is a game, not all game SSF records are retrieved
                tmSsfSrchKey2.iVefCode = ilVefCode
                tmSsfSrchKey2.iDate(0) = ilDate0
                tmSsfSrchKey2.iDate(1) = ilDate1
                ilRet = gSSFGetGreaterOrEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
                ilType = tmSsf.iType
            End If
            '   if games, tmSsf.iType will be the game #.  this report is not driven by game; but by week
            If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iVefCode <> ilVefCode) Or ((tmSsf.iDate(0) <> ilDate0) And (tmSsf.iDate(1) = ilDate1)) Or (ilType <> tmSsf.iType And tgMVef(ilVefIndex).sType <> "G") Then
                If (llLoopDate > llLatestDate) Then
                    ReDim tlLLC(0 To 0) As LLC  'Merged library names
                    If tgMVef(ilVefIndex).sType <> "G" Then
                        ilWeekDay = gWeekDayStr(slDate) + 1 '1=Monday TFN; 2=Tues,...7=Sunday TFN
                        If ilWeekDay = 1 Then
                             ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNMO", "12M", "12M", ilEvtType(), tlLLC())
                        ElseIf ilWeekDay = 2 Then
                             ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTU", "12M", "12M", ilEvtType(), tlLLC())
                        ElseIf ilWeekDay = 3 Then
                             ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNWE", "12M", "12M", ilEvtType(), tlLLC())
                        ElseIf ilWeekDay = 4 Then
                             ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTH", "12M", "12M", ilEvtType(), tlLLC())
                        ElseIf ilWeekDay = 5 Then
                             ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNFR", "12M", "12M", ilEvtType(), tlLLC())
                        ElseIf ilWeekDay = 6 Then
                             ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSA", "12M", "12M", ilEvtType(), tlLLC())
                        ElseIf ilWeekDay = 7 Then
                             ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSU", "12M", "12M", ilEvtType(), tlLLC())
                        End If
                    End If
                    'tmSsf.sType = "O"
                    tmSsf.iType = ilType
                    tmSsf.iVefCode = ilVefCode
                    tmSsf.iDate(0) = ilDate0
                    tmSsf.iDate(1) = ilDate1
                    gPackTime tlLLC(0).sStartTime, tmSsf.iStartTime(0), tmSsf.iStartTime(1)
                    tmSsf.iCount = 0
                    'tmSsf.iNextTime(0) = 1  'Time not defined
                    'tmSsf.iNextTime(1) = 0
    
                    For ilIndex = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
    
                        tmAvail.iRecType = Val(tlLLC(ilIndex).sType)
                        gPackTime tlLLC(ilIndex).sStartTime, tmAvail.iTime(0), tmAvail.iTime(1)
                        tmAvail.iLtfCode = tlLLC(ilIndex).iLtfCode
                        tmAvail.iAvInfo = tlLLC(ilIndex).iAvailInfo Or tlLLC(ilIndex).iUnits
                        tmAvail.iLen = CInt(gLengthToCurrency(tlLLC(ilIndex).sLength))
                        tmAvail.ianfCode = Val(tlLLC(ilIndex).sName)
                        tmAvail.iNoSpotsThis = 0
                        tmAvail.iOrigUnit = 0
                        tmAvail.iOrigLen = 0
                        tmSsf.iCount = tmSsf.iCount + 1
                        tmSsf.tPas(ADJSSFPASBZ + tmSsf.iCount) = tmAvail
                    Next ilIndex
                    ilRet = BTRV_ERR_NONE
                End If
            End If
    
            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
                ilDay = gWeekDayLong(llLoopDate)
                ilEvt = 1
                Do While ilEvt <= tmSsf.iCount
                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    If (tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 2) Then 'Contract Avails only
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                        'determine the Base DP this spot belongs in
                        ilFoundDP = False
                        For ilRdf = LBound(tlAvRdf) To UBound(tlAvRdf) - 1
                            For ilLoop = LBound(tlAvRdf(ilRdf).iStartTime, 2) To UBound(tlAvRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                If (tlAvRdf(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tlAvRdf(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                                    gUnpackTimeLong tlAvRdf(ilRdf).iStartTime(0, ilLoop), tlAvRdf(ilRdf).iStartTime(1, ilLoop), False, llStartTime
                                    gUnpackTimeLong tlAvRdf(ilRdf).iEndTime(0, ilLoop), tlAvRdf(ilRdf).iEndTime(1, ilLoop), True, llEndTime
                                    'If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDay + 1) = "Y") Then
                                    If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDay) = "Y") Then
                                        'times & days passed, test avail type
                                        If tlAvRdf(ilRdf).sInOut = "I" Then   'Book into
                                            If tmAvail.ianfCode = tlAvRdf(ilRdf).ianfCode Then
                                                ilFoundDP = True
                                                Exit For
                                            End If
                                        ElseIf tlAvRdf(ilRdf).sInOut = "O" Then   'Exclude
                                            If tmAvail.ianfCode <> tlAvRdf(ilRdf).ianfCode Then
                                                ilFoundDP = True
                                                Exit For
                                            End If
                                        Else
                                            'Neither, all avails
                                            ilFoundDP = True
                                        End If
                                    End If
                                End If
                            Next ilLoop
                            If ilFoundDP Then
                                Exit For
                            End If
                        Next ilRdf
                        
                        'get the relative week index of the spot to accum count
                        ilWkInx = (llLoopDate - llSDate) \ 7 + 1
                        'get the spot
                        For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                        
                            ilSpotOK = True
                           LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                     
                            tmSdfSrchKey3.lCode = tmSpot.lSdfCode   'get the spot record
                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet <> BTRV_ERR_NONE Then
                                ilSpotOK = False                    'invalid sdf code
                            Else
                                ilSpotOK = mFilterSpot(tlCntTypes)
                            End If

                            If ilSpotOK Then            'passed spot and contract filters
                                If ilFoundDP Then       'a base daypart was found
                                    ilFoundPP = False
                                    ilRdfCode = tlAvRdf(ilRdf).iCode
                                    If RptSelCM!rbcTotals(1).Value Then         'totals by vehicle, not by DP
                                        ilRdfCode = 0
                                    End If
                                    For ilLoopOnCounts = LBound(tmBaseCounts) To UBound(tmBaseCounts) - 1
                                        'If tmBaseCounts(ilLoopOnCounts).iRdfcode = tlAvRdf(ilRdf).iCode And tmBaseCounts(ilLoopOnCounts).iCCMnfCode = tmChf.iMnfComp(0) Then
                                        If tmBaseCounts(ilLoopOnCounts).iRdfCode = ilRdfCode And tmBaseCounts(ilLoopOnCounts).iCCMnfCode = tmChf.iMnfComp(0) Then
                                            'determine week this spot count belongs in
                                            tmBaseCounts(ilLoopOnCounts).iCCPeriod(ilWkInx) = tmBaseCounts(ilLoopOnCounts).iCCPeriod(ilWkInx) + 1
                                            ilFoundPP = True
                                            Exit For
                                        End If
                                    Next ilLoopOnCounts
                                    If Not ilFoundPP Then       'Base DP Counts not created yet
                                        ilUpper = UBound(tmBaseCounts)

                                        tmBaseCounts(ilUpper).iVefCode = ilVefCode
                                        If RptSelCM!rbcTotals(1).Value Then             'no daypart subtotals
                                            tmBaseCounts(ilUpper).iRdfCode = 0
                                            tmBaseCounts(ilUpper).iSort = 0
                                        Else
                                            tmBaseCounts(ilUpper).iRdfCode = tlAvRdf(ilRdf).iCode
                                            tmBaseCounts(ilUpper).iSort = tlAvRdf(ilRdf).iSortCode
                                        End If
                                        tmBaseCounts(ilUpper).iCCMnfCode = tmChf.iMnfComp(0)
                                        'determine week this spot count belongs in
                                        tmBaseCounts(ilUpper).iCCPeriod(ilWkInx) = tmBaseCounts(ilLoopOnCounts).iCCPeriod(ilWkInx) + 1
                                        ReDim Preserve tmBaseCounts(0 To ilUpper + 1) As DPCOUNTS
                                    End If
                                Else            'no base daypart found, put the spot into orphan daypart
                                    'currently not doing anything with this array of No Base dayparts
                                    ilFoundPP = False
                                    For ilLoopOnCounts = LBound(tmNoBAseCounts) To UBound(tmNoBAseCounts) - 1
                                        If tmNoBAseCounts(ilLoopOnCounts).iCCMnfCode = tmChf.iMnfComp(0) Then
                                            ilUpper = UBound(tmNoBAseCounts)
                                            
                                            'determine week this spot count belongs in
                                            tmNoBAseCounts(ilLoopOnCounts).iCCPeriod(ilWkInx) = tmNoBAseCounts(ilLoopOnCounts).iCCPeriod(ilWkInx) + 1
                                            'ReDim Preserve tmNoBAseCounts(0 To ilUpper + 1) As DPCOUNTS
                                            ilFoundPP = True
                                            Exit For
                                        End If
                                    Next ilLoopOnCounts
                                    If Not ilFoundPP Then
                                        ilUpper = UBound(tmNoBAseCounts)
                                        tmNoBAseCounts(ilUpper).iVefCode = ilVefCode
                                        tmNoBAseCounts(ilUpper).iRdfCode = -1
                                        tmNoBAseCounts(ilUpper).iSort = 32000
                                        tmNoBAseCounts(ilUpper).iCCMnfCode = tmChf.iMnfComp(0)
                                        'determine week this spot count belongs in
                                        tmNoBAseCounts(ilUpper).iCCPeriod(ilWkInx) = 1
                                        ReDim Preserve tmNoBAseCounts(0 To ilUpper + 1) As DPCOUNTS
                                    End If
                                End If
                            End If              'ilspotok
                        Next ilSpot
                        ilEvt = ilEvt + tmAvail.iNoSpotsThis
                    End If
                    ilEvt = ilEvt + 1
                Loop
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                If tgMVef(ilVefIndex).sType = "G" Then
                    ilType = tmSsf.iType
                End If
            Loop
        Next llLoopDate           'next ssf
        Exit Sub
End Sub
'
'
'
'               mFilterSpot - Test header and line exclusions for user request.
'
'               <input> tlCntTypes - structure of inclusions/exclusions of contract types & status
'              <output>  None
'                return   ilSpotOk - true if spot is OK, else false to ignore spot
Function mFilterSpot(tlCntTypes As CNTTYPES) As Integer
Dim ilRet As Integer
Dim slPrice As String
Dim ilSpotOK As Integer
Dim ilInclProdProt As Integer

        ilSpotOK = True
        
        If tmSdf.lChfCode <> tmChf.lCode Then           'if already in mem, don't reread
            tmChfSrchKey.lCode = tmSdf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                ilSpotOK = False
            End If
        End If
        'Test header exclusions (types of contrcts and statuses)
        If tmChf.sStatus = "H" Then
            If Not tlCntTypes.iHold Then
                ilSpotOK = False
            End If
        ElseIf tmChf.sStatus = "O" Then

            If Not tlCntTypes.iOrder Then
                ilSpotOK = False
            End If
        End If

        If tmChf.sType = "C" Then               '3-16-10 wrong flag was tested for standard, test for C not S
            If Not tlCntTypes.iStandard Then       'include Standard types?
                ilSpotOK = False
            End If

        ElseIf tmChf.sType = "V" Then
            If Not tlCntTypes.iReserv Then      'include reservations ?
                ilSpotOK = False
            End If

        ElseIf tmChf.sType = "R" Then
            If Not tlCntTypes.iDR Then       'include DR?
                ilSpotOK = False
            End If
        End If
        
        If tmChf.iPctTrade = 100 And Not tlCntTypes.iTrade Then         'only test for 100% trade to exclude
            ilSpotOK = False
        End If

        ilInclProdProt = gFilterLists(tmChf.iMnfComp(0), imInclProdCodes, imUseProdCodes())
        If Not ilInclProdProt Then
            ilSpotOK = False
        End If

        If tmSdf.lChfCode <> tmChf.lChfCode Then
            tmClfSrchKey.lChfCode = tmSdf.lChfCode
            tmClfSrchKey.iLine = tmSdf.iLineNo
            tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
            tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        End If
        If (tmClf.lChfCode <> tmSdf.lChfCode) Then     'got the matching line reference?
            ilSpotOK = False
        End If
        'Retrieve spot cost from flight ; flight not returned if spot type is Extra/Fill
        'otherwise flight returned in tgPriceCff
        ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
        'look for inclusion of spot types
        If (InStr(slPrice, ".") <> 0) Then    'found spot cost
            'is it a .00?
            If gCompNumberStr(slPrice, "0.00") = 0 Then     'its a .00 spot
                If Not tlCntTypes.iZero Then
                    ilSpotOK = False
                End If
            Else
                If Not tlCntTypes.iCharge Then           'exclude charged spots
                    ilSpotOK = False
                End If
            End If
        ElseIf Trim$(slPrice) = "ADU" Then
            If Not tlCntTypes.iADU Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "Bonus" Then
            If Not tlCntTypes.iBonus Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "+ Fill" Then       '3-24-03
            If Not tlCntTypes.iXtra Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "- Fill" Then        '3-24-03
            If Not tlCntTypes.iFill Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "N/C" Then
            If Not tlCntTypes.iNC Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "Recapturable" Then
            If Not tlCntTypes.iRecapturable Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "Spinoff" Then
            If Not tlCntTypes.iSpinoff Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "MG" Then
            If Not tlCntTypes.iMG Then
                ilSpotOK = False
            End If
        End If

        'test for spot sched statuses: mg & outsides, missed/cancelled
        If (tmSdf.sSchStatus = "G" Or tmSdf.sSchStatus = "O") And (Not tlCntTypes.iMG) Then
            ilSpotOK = False
        End If
        If (tmSdf.sSchStatus = "M" Or tmSdf.sSchStatus = "C") And (Not tlCntTypes.iMissed) Then
            ilSpotOK = False
        End If
          
        mFilterSpot = ilSpotOK
End Function

Public Sub mUpdateGRFCount(ilStartDate() As Integer, ilMajorSet As Integer, tlCounts As DPCOUNTS)
Dim ilMinorSet As Integer
Dim ilmnfMinorCode As Integer
Dim ilMnfMajorCode As Integer
Dim ilLoop As Integer
Dim ilRet As Integer

        '      10/6/10 grfFields
        '      GenDate - Generation date
        '      GenTime - Generation Time
        '      VefCode - vehicle code
        '      rdfCode - DP Code
        '      sofcode - competitive mnf code
        '      Code2 - mnf vehicle group
        '      StartDate - Date of Competitive data
        '      PerGenl(1-13) = counts, wk 1-13
        '      Code4 - DP sort code

        tmGrf.iGenDate(0) = igNowDate(0)
        tmGrf.iGenDate(1) = igNowDate(1)
        tmGrf.lGenTime = lgNowTime
        tmGrf.iVefCode = tlCounts.iVefCode              'vehicle code
        tmGrf.iSofCode = tlCounts.iCCMnfCode    'competitive mnf code
        tmGrf.iRdfCode = tlCounts.iRdfCode  'daypart
        tmGrf.iSlfCode = imRcfCode          'rate card code
        gGetVehGrpSets tmGrf.iVefCode, ilMinorSet, ilMajorSet, ilmnfMinorCode, ilMnfMajorCode
        tmGrf.iCode2 = ilMnfMajorCode  'vehicle group mnf code
        tmGrf.iStartDate(0) = ilStartDate(0)     'start date of requested period
        tmGrf.iStartDate(1) = ilStartDate(1)
        tmGrf.lCode4 = tlCounts.iSort
        For ilLoop = 1 To 13
            'tmGrf.iPerGenl(ilLoop) = tlCounts.iCCPeriod(ilLoop)
            tmGrf.iPerGenl(ilLoop - 1) = tlCounts.iCCPeriod(ilLoop)
        Next ilLoop
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        Exit Sub
End Sub
