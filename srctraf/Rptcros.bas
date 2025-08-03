Attribute VB_Name = "RPTCROS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcros.bas on Wed 6/17/09 @ 12:56 PM
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
Dim hmSsf As Integer
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim hmSmf As Integer            'MG file handle
Dim tmSmf As SMF                'SMF record image
Dim imSmfRecLen As Integer        'SMF record length
'Log Calendar
Dim hmLcf As Integer            'Log Calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim imLcfRecLen As Integer        'LCF record length
Dim tmGrf() As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim tmRcf As RCF
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

Dim tmAvRdf() As RDF            'array of dayparts
'*******************************************************************
'*                                                                 *
'*      Procedure Name:gCRQtrlyBookSpots                           *
'*                                                                 *
'*             Created:12/29/97      By:D. Hosaka                  *
'*            Modified:              By:                           *
'*                                                                 *
'*            Comments: Generate Oversold Report                   *
'*                                                                 *
'*      3/11/98 Look at "Base DP" only, (not Report DP)            *
'*      4/12/98 Remove duplication of spots from vehicle           *
'*              These spots appeared to be moved across vehicles
'
'
'*                                                                 *
'*******************************************************************
Sub gCreateOS()
'
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim sldate As String
    ReDim ilDate(0 To 1) As Integer
    Dim ilVehicle As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilDay As Integer
    Dim ilVpfIndex As Integer
    Dim ilUpper As Integer
    Dim ilDateOk As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilRcf As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim llEffDate As Long
    Dim llDateEntered As Long                   'orig user date entered (backed up to a Monday)
    Dim ilNoWks As Integer                      'user # weeks requested
    Dim tlCntTypes As CNTTYPES
    Dim ilSaveSort As Integer                   'DP or RIF field:  sort code
    Dim slSaveReport As String                  'DP or RIF field:  Save on report
    Dim ilWeek As Integer
    ReDim tmGrf(0 To 0) As GRF
    Dim ilDone As Integer
    Dim ilOptionBookedUnsold As Integer           'true if Show booked & unsold
    Dim ilMajorSet As Integer                      'vehicle sort group
    Dim ilMinorSet As Integer                   'minor vehicle group (not used)
    Dim ilMnfMajorCode As Integer               'vehicle group mnf code
    Dim ilmnfMinorCode As Integer               'minor MNF code (not used)
    Dim ilHowManyHL As Integer                  'number of spot lengths to highlight (min 1, max 4)
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
    imGrfRecLen = Len(tmGrf(0))

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
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmSsf
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmLcf
        btrDestroy hmSsf
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    imLcfRecLen = Len(tmLcf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmSmf
        btrDestroy hmLcf
        btrDestroy hmSsf
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
    imAnfRecLen = Len(tmAnf)

    tlCntTypes.iHold = gSetCheck(RptSelOS!ckcCType(0).Value)
    tlCntTypes.iOrder = gSetCheck(RptSelOS!ckcCType(1).Value)
    tlCntTypes.iNetwork = gSetCheck(RptSelOS!ckcCType(2).Value)
    tlCntTypes.iStandard = gSetCheck(RptSelOS!ckcCType(3).Value)
    tlCntTypes.iReserv = gSetCheck(RptSelOS!ckcCType(4).Value)
    tlCntTypes.iRemnant = gSetCheck(RptSelOS!ckcCType(5).Value)
    tlCntTypes.iDR = gSetCheck(RptSelOS!ckcCType(6).Value)
    tlCntTypes.iPI = gSetCheck(RptSelOS!ckcCType(7).Value)
    tlCntTypes.iPSA = gSetCheck(RptSelOS!ckcCType(8).Value)
    tlCntTypes.iPromo = gSetCheck(RptSelOS!ckcCType(9).Value)
    tlCntTypes.iTrade = gSetCheck(RptSelOS!ckcCType(10).Value)
    tlCntTypes.iMissed = gSetCheck(RptSelOS!ckcSpots(0).Value)
    tlCntTypes.iCharge = gSetCheck(RptSelOS!ckcSpots(1).Value)
    tlCntTypes.iZero = gSetCheck(RptSelOS!ckcSpots(2).Value)
    tlCntTypes.iADU = gSetCheck(RptSelOS!ckcSpots(3).Value)
    tlCntTypes.iBonus = gSetCheck(RptSelOS!ckcSpots(4).Value)
    tlCntTypes.iXtra = gSetCheck(RptSelOS!ckcSpots(5).Value)
    tlCntTypes.iFill = gSetCheck(RptSelOS!ckcSpots(6).Value)
    tlCntTypes.iNC = gSetCheck(RptSelOS!ckcSpots(7).Value)
    tlCntTypes.iRecapturable = gSetCheck(RptSelOS!ckcSpots(8).Value)
    tlCntTypes.iSpinoff = gSetCheck(RptSelOS!ckcSpots(9).Value)
    tlCntTypes.iMG = gSetCheck(RptSelOS!ckcSpots(10).Value)         '10-27-10
    tlCntTypes.iFixedTime = gSetCheck(RptSelOS!ckcRank(0).Value)
    tlCntTypes.iSponsor = gSetCheck(RptSelOS!ckcRank(1).Value)
    tlCntTypes.iDP = gSetCheck(RptSelOS!ckcRank(2).Value)
    tlCntTypes.iROS = gSetCheck(RptSelOS!ckcRank(3).Value)

    If (tlCntTypes.iHold) Or (tlCntTypes.iOrder) Then        '1-26-05 set general cntr type for inclusion/exclusion if hold or ordered selected
        tlCntTypes.iCntrSpots = True
    Else
        tlCntTypes.iCntrSpots = False
    End If

    'build the spot lengths high to lo order
    For ilLoop = 0 To 3
        tlCntTypes.iLenHL(ilLoop) = Val(RptSelOS!edcLength(ilLoop))
    Next ilLoop
    ilDone = False
    Do While Not ilDone
        ilDone = True
        For ilLoop = 1 To 3
            If tlCntTypes.iLenHL(ilLoop - 1) < tlCntTypes.iLenHL(ilLoop) Then
                'swap the two
                ilRet = tlCntTypes.iLenHL(ilLoop - 1)
                tlCntTypes.iLenHL(ilLoop - 1) = tlCntTypes.iLenHL(ilLoop)
                tlCntTypes.iLenHL(ilLoop) = ilRet
                ilDone = False
            End If
        Next ilLoop
    Loop
    ilHowManyHL = 0
    For ilLoop = 0 To 3
        If tlCntTypes.iLenHL(ilLoop) > 0 Then
            ilHowManyHL = ilHowManyHL + 1
        End If
    Next ilLoop
    For ilLoop = 0 To 6
        tlCntTypes.iValidDays(ilLoop) = True
        If Not RptSelOS!ckcDays(ilLoop).Value = vbChecked Then
            tlCntTypes.iValidDays(ilLoop) = False
        End If
    Next ilLoop
    If RptSelOS!rbcShow(1).Value Then       'show sold & avail
        ilOptionBookedUnsold = True
    Else                                    'show booked & unsold
        ilOptionBookedUnsold = False
    End If
    'Get the vehicle group selected for sorting
    ilRet = RptSelOS!cbcGroup.ListIndex
    ilMajorSet = gFindVehGroupInx(ilRet, tgVehicleSets1())
    ilRet = gObtainRcfRifRdf()          'get the rate cards and assoc dayparts

    'get all the dates needed to work with
'    sldate = RptSelOS!edcSelCFrom.Text               'effective date entred
    sldate = RptSelOS!CSI_CalFrom.Text               '12-13-19 change to use csi calendar control vs edit box;effective date entred
    llEffDate = gDateValue(sldate)
    'backup to Monday
    ilDay = gWeekDayLong(llEffDate)
    Do While ilDay <> 0
        llEffDate = llEffDate - 1
        ilDay = gWeekDayLong(llEffDate)
    Loop
    llDateEntered = llEffDate               'save orig date to calculate week index

    tmVef.iCode = 0
    ilNoWks = Val(RptSelOS!edcSelCFrom1.Text)
    For ilWeek = 1 To ilNoWks Step 1
        For ilVehicle = 0 To RptSelOS!lbcSelection(0).ListCount - 1 Step 1
            If (RptSelOS!lbcSelection(0).Selected(ilVehicle)) Then
                slNameCode = tgCSVNameCode(ilVehicle).sKey 'RptSelSP!lbcCSVNameCode.List(ilVehicle)
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slName, 3, "|", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                ilVpfIndex = -1
                'For ilLoop = 0 To UBound(tgVpf) Step 1
                '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
                    ilLoop = gBinarySearchVpf(ilVefCode)
                    If ilLoop <> -1 Then
                        ilVpfIndex = ilLoop
                '        Exit For
                    End If
                'Next ilLoop
                If ilVpfIndex >= 0 Then
                    For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                        tmRcf = tgMRcf(ilRcf)
                        ilDateOk = False
                        For ilLoop = 0 To RptSelOS!lbcSelection(1).ListCount - 1 Step 1
                            slNameCode = tgRateCardCode(ilLoop).sKey
                            ilRet = gParseItem(slNameCode, 3, "\", slCode)
                            If Val(slCode) = tgMRcf(ilRcf).iCode Then
                                If (RptSelOS!lbcSelection(1).Selected(ilLoop)) Then
                                    ilDateOk = True
                                End If
                                Exit For
                            End If
                        Next ilLoop

                        If ilDateOk Then
                            ReDim tmGrf(0 To 0) As GRF
                            ReDim tmAvRdf(0 To 0) As RDF
                            ReDim tmRifRate(0 To 0) As RIF
                            ilUpper = 0
                            'Setup the DP with rates to process
                            For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                                If tgMRif(llRif).iRcfCode = tgMRcf(ilRcf).iCode Then
                                    'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                    ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                                    If ilRdf <> -1 Then
                                        'determine if rate card items have sort codes, etc; otherwise use DP
                                        If tgMRif(llRif).sRpt <> "Y" And tgMRif(llRif).sRpt <> "N" Then
                                            slSaveReport = tgMRdf(ilRdf).sReport
                                        Else
                                            slSaveReport = tgMRif(llRif).sRpt
                                        End If
                                        If tgMRif(llRif).iSort = 0 Then
                                            ilSaveSort = tgMRdf(ilRdf).iSortCode
                                        Else
                                            ilSaveSort = tgMRif(llRif).iSort
                                        End If
                                        If tgMRdf(ilRdf).iCode = tgMRif(llRif).iRdfCode And slSaveReport = "Y" And tgMRdf(ilRdf).sState <> "D" And tgMRif(llRif).iVefCode = ilVefCode Then
                                            ilFound = False
                                            For ilLoop = LBound(tmAvRdf) To ilUpper - 1 Step 1
                                                If tmAvRdf(ilLoop).iCode = tgMRdf(ilRdf).iCode Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilLoop
                                            If Not ilFound Then
                                                tmAvRdf(ilUpper) = tgMRdf(ilRdf)
                                                tmRifRate(ilUpper) = tgMRif(llRif)
                                                tmRifRate(ilUpper).iSort = ilSaveSort
                                                ilUpper = ilUpper + 1
                                                ReDim Preserve tmAvRdf(0 To ilUpper) As RDF
                                                ReDim Preserve tmRifRate(0 To ilUpper) As RIF
                                            End If
                                    '        Exit For
                                        End If
                                    End If
                                    'Next ilRdf
                                End If
                            Next llRif
                            mGetSpotCountsOS ilVefCode, ilVpfIndex, llEffDate, llEffDate + 6, tmAvRdf(), tmRifRate(), tlCntTypes
                            'Calculate week # for sorting
                            ilDay = (llEffDate - llDateEntered) / 7 + 1
                            gPackDateLong llEffDate, ilDate(0), ilDate(1)
                            For ilIndex = 0 To UBound(tmGrf) - 1 Step 1
                            '      9/26/00 grfFields
                            '      GenDate - Generation date
                            '      GenTime - Generation Time
                            '      VefCode - vehicle code
                            '      rdfCode - DP Code
                            '      adfCode - DP Sort code from RIF for vehicle
                            '      StartDate - Date of Oversold data (1 record/day)
                            '      Date      - start date of week
                            '      Year - Day of Week (0-6, M-Su)
                            '      Code2 - mnf code for vehicle group sort
                            '      lDollars(1) - Orig Inventory (seconds or units)
                            '      lDollars(2) - Booked            ""
                            '      lDollars(3) - Missed            ""
                            '                    % unsold calculated in Crystal
                            '      lDollars(4) - Unsold            ""
                            '      lDollars(5) - Fixed Time        ""
                            '      lDollars(6) - Sponsorship       ""
                            '      lDollars(7) - DP high           ""
                            '      lDollars(8) - DP LO             ""
                            '      lDollars(9) - ROS               ""
                            '      lDollars(10) - Trades            ""
                            '      lDollars(11) - Reservations     ""
                            '      lDollars(12) - Other (DR,PI,PSA,Promo,Remnant)   ""
                            '      lDollars(13) - Extra/Fills     (seconds or units)
                            '      lDollars(14) - Network          ""
                            '      PerGenl(1)   - Length #1 to highlight count
                            '      PerGenl(2)     Length #2    ""
                            '      PerGenl(3)     Length #3    ""
                            '      PerGenl(4)     Length #4    ""
                            '      PerGenl(5)     Number of spot length columns to show
                            '      PerGenl(6)   - Length #1 to highlight for report hders
                            '      PerGenl(7)     Length #2    ""
                            '      PerGenl(8)     Length #3    ""
                            '      PerGenl(9)     Length #4    ""
                            '      PerGenl(10)    Week Index #
                                'Setup the major sort factor
                                gGetVehGrpSets ilVefCode, ilMinorSet, ilMajorSet, ilmnfMinorCode, ilMnfMajorCode
                                tmGrf(ilIndex).iCode2 = ilMnfMajorCode  'vehicle group mnf code
                                'tmGrf(ilIndex).iPerGenl(5) = ilHowManyHL
                                tmGrf(ilIndex).iPerGenl(4) = ilHowManyHL
                                'Determine Unsold based on Booked & Unsold vs Sold & Available
                                If ilOptionBookedUnsold Then
                                    'tmGrf(ilIndex).lDollars(4) = tmGrf(ilIndex).lDollars(1) - tmGrf(ilIndex).lDollars(2)    'orig inv. - booked (not including missed)
                                    tmGrf(ilIndex).lDollars(3) = tmGrf(ilIndex).lDollars(0) - tmGrf(ilIndex).lDollars(1)    'orig inv. - booked (not including missed)
                                Else
                                    'tmGrf(ilIndex).lDollars(4) = tmGrf(ilIndex).lDollars(1) - (tmGrf(ilIndex).lDollars(2) + tmGrf(ilIndex).lDollars(3))   'orig inv. - (booked+ Missed), could be negative
                                    tmGrf(ilIndex).lDollars(3) = tmGrf(ilIndex).lDollars(0) - (tmGrf(ilIndex).lDollars(1) + tmGrf(ilIndex).lDollars(2))   'orig inv. - (booked+ Missed), could be negative
                                End If
                                'Pass spot lengths to highlight for headers
                                For ilLoop = 0 To 3
                                    'tmGrf(ilIndex).iPerGenl(6 + ilLoop) = tlCntTypes.iLenHL(ilLoop)
                                    tmGrf(ilIndex).iPerGenl(6 + ilLoop - 1) = tlCntTypes.iLenHL(ilLoop)
                                Next ilLoop
                                'tmGrf(ilIndex).iPerGenl(10) = ilDay     'week index
                                tmGrf(ilIndex).iPerGenl(9) = ilDay     'week index
                                tmGrf(ilIndex).iDate(0) = ilDate(0)     'start date of week
                                tmGrf(ilIndex).iDate(1) = ilDate(1)
                                ilRet = btrInsert(hmGrf, tmGrf(ilIndex), imGrfRecLen, INDEXKEY0)
                            Next ilIndex
                        End If
                    Next ilRcf                      'next rate card (should only be 1)
                End If                              'vpfindex > 0
            End If                                  'vehicle selected
        Next ilVehicle                              'For ilvehicle = 0 To RptSelOS!lbcSelection(0).ListCount - 1
        llEffDate = llEffDate + 7                   'next week
    Next ilWeek
    Erase tmAvRdf, tmGrf
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmLcf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmFsf)
    ilRet = btrClose(hmAnf)
    btrDestroy hmRdf
    btrDestroy hmSdf
    btrDestroy hmVef
    btrDestroy hmSsf
    btrDestroy hmLcf
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmMnf
    btrDestroy hmFsf
    btrDestroy hmAnf
    Exit Sub
End Sub
'*****************************************************************
'*                                                               *
'*                                                               *
'*                                                               *
'*      Procedure Name:mGetSpotCounts                            *
'*                                                               *
'*      Created:9/27/00       By:D. Hosaka                       *
'*                                                               *
'*
'*      3-24-03 change way to test fill/extra spots.  Use SDF not SSF
'*****************************************************************
Sub mGetSpotCountsOS(ilVefCode As Integer, ilVpfIndex As Integer, llSDate As Long, llEDate As Long, tlAvRdf() As RDF, tlRif() As RIF, tlCntTypes As CNTTYPES)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilMissedDate                                                                          *
'******************************************************************************************

'
'   Where:
'
'   ilVefCode (I) - vehicle code to process
'   ilVpfIndex (I) - vehicle options pointer
'   llSDate (I) - start date to begin searching Avails
'   llEDate (I) - end date to stop searching avails
'   tlAvRdf() (I) - array of Dayparts
'   tlCntTypes (I) - contract and spot types to include in search
'
'   Note: Remnants; Direct Response; per Inquiry; PSA and Promos are not
'         saved with a miss status
'         For scheduled spots the rank is used to determine if it is one
'         of the above (Direct reponse=1010; Remnant=1020; per Inquiry= 1030;
'         PSA=1060; Promo=1050.
'

    Dim slType As String
    Dim ilType As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim sldate As String
    Dim lldate As Long
    Dim ilEvt As Integer
    Dim ilRet As Integer
    Dim ilSpot As Integer
    Dim llTime As Long
    Dim ilRdf As Integer
    Dim llRif As Long
    Dim ilRifDPSort As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilRec As Integer
    Dim ilRecIndex As Integer
    Dim ilAddCategory As Integer        'length in seconds of avail or spot or number of units to accumulate based
                                        'on the selling method of vehicle
    Dim slDays As String
    Dim ilLtfCode As Integer
    Dim ilAvailOk As Integer
    Dim ilPass As Integer
    Dim ilDayIndex As Integer
    Dim ilLoopIndex As Integer
    Dim ilBucketIndex As Integer
    Dim ilSpotOK As Integer
    Dim llLoopDate As Long
    Dim ilWeekDay As Integer
    Dim llLatestDate As Long
    Dim ilIndex As Integer
    Dim ilRemLen As Integer     'time in seconds of avail, each spot length subtracted to get remaining seconds
    Dim ilRemUnits As Integer   '# of units of avail, each spot length subtracted to get remaining units
    Dim ilCTypes As Integer       'bit map of cnt types to include starting lo order bit
                                  '0 = unused, 1= unused, 2 = network, 3 = std, 4 = Reserved, 5 = remanant, 6 = DR
                                  '7 = PI, 8 = psa, 9 = promo, 10 = trade
                                  'bit 0 and 1 previously hold & order; but it conflicts with other contract types
    Dim ilSpotTypes As Integer    'bit map of spot types to include starting lo order bit:
                                  '0 = missed, 1 = charge, 2 = 0.00, 3 = adu, 4 = bonus, 5 = extra
                                  '6 = fill, 7 = n/c 8 = recapturable, 9 = spinoff, 10 = MG
    Dim ilRanks As Integer        '0 = fixed time, 1 = sponsorship, 2 = DP, 3= ROS
    Dim ilOrphanMissedLoop As Integer
    Dim ilOrphanFound As Integer
    Dim ilFilterDay As Integer
    ReDim ilEvtType(0 To 14) As Integer
    'ReDim ilRdfCodes(0 To 1) As Integer
    ReDim tmGrf(0 To 0) As GRF
    Dim slChfType As String
    Dim slChfStatus As String
    Dim ilVefIndex As Integer

    slType = "O"
    ilType = 0
    llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
    'set the type of events to get fro the day (only Contract avails)
    For ilLoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilLoop) = False
    Next ilLoop
    ilEvtType(2) = True
    imSdfRecLen = Len(tmSdf)
    imCHFRecLen = Len(tmChf)
    ilVefIndex = gBinarySearchVef(ilVefCode)
    For llLoopDate = llSDate To llEDate Step 1
        ilFilterDay = gWeekDayLong(llLoopDate)
        If tlCntTypes.iValidDays(ilFilterDay) Then      'Has this day of the week been selected?
            sldate = Format$(llLoopDate, "m/d/yy")
            gPackDate sldate, ilDate0, ilDate1
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            'tmSsfSrchKey.sType = slType
            If tgMVef(ilVefIndex).sType <> "G" Then
                tmSsfSrchKey.iType = ilType
                tmSsfSrchKey.iVefCode = ilVefCode
                tmSsfSrchKey.iDate(0) = ilDate0
                tmSsfSrchKey.iDate(1) = ilDate1
                tmSsfSrchKey.iStartTime(0) = 0
                tmSsfSrchKey.iStartTime(1) = 0
                ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
            Else
                tmSsfSrchKey2.iVefCode = ilVefCode
                tmSsfSrchKey2.iDate(0) = ilDate0
                tmSsfSrchKey2.iDate(1) = ilDate1
                ilRet = gSSFGetGreaterOrEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
                ilType = tmSsf.iType
            End If
            'If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.sType <> slType) Or (tmSsf.iVefcode <> ilVefCode Or (tmSsf.iDate(0) <> ilDate0) And (tmSsf.iDate(1) = ilDate1)) Then
            If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iType <> ilType) Or (tmSsf.iVefCode <> ilVefCode Or (tmSsf.iDate(0) <> ilDate0) And (tmSsf.iDate(1) = ilDate1)) Then
                If (llLoopDate > llLatestDate) Then
                    ReDim tlLLC(0 To 0) As LLC  'Merged library names
                    If tgMVef(ilVefIndex).sType <> "G" Then
                        ilWeekDay = gWeekDayStr(sldate) + 1 '1=Monday TFN; 2=Tues,...7=Sunday TFN
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

            'Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slType) And (tmSsf.iVefcode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
                gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), lldate
                ilBucketIndex = gWeekDayLong(lldate)        'day of week bucket index
                ilEvt = 1
                Do While ilEvt <= tmSsf.iCount
                   LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    If tmProg.iRecType = 1 Then    'Program (not working for nested prog)
                        ilLtfCode = tmProg.iLtfCode
                    ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 2) Then 'Contract Avails only
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                        'Determine which rate card program this is associated with
                        For ilRdf = LBound(tlAvRdf) To UBound(tlAvRdf) - 1 Step 1
                            ilAvailOk = False
                            If (tlAvRdf(ilRdf).iLtfCode(0) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(1) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(2) <> 0) Then
                                If (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(0)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Then
                                    ilAvailOk = False    'True- code later
                                End If
                            Else

                                For ilLoop = LBound(tlAvRdf(ilRdf).iStartTime, 2) To UBound(tlAvRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                    If (tlAvRdf(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tlAvRdf(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                                        gUnpackTimeLong tlAvRdf(ilRdf).iStartTime(0, ilLoop), tlAvRdf(ilRdf).iStartTime(1, ilLoop), False, llStartTime
                                        gUnpackTimeLong tlAvRdf(ilRdf).iEndTime(0, ilLoop), tlAvRdf(ilRdf).iEndTime(1, ilLoop), True, llEndTime
                                        'If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilBucketIndex + 1) = "Y") Then
                                        If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilBucketIndex) = "Y") Then
                                            ilAvailOk = True
                                            ilLoopIndex = ilLoop
                                            slDays = ""
                                            For ilDayIndex = 1 To 7 Step 1
                                                If (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "Y") Or (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "N") Then
                                                    slDays = slDays & tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1)
                                                Else
                                                    slDays = slDays & "N"
                                                End If
                                            Next ilDayIndex
                                            Exit For
                                        End If
                                    End If
                                Next ilLoop
                            End If


                            If ilAvailOk Then
                                'find the associated Items entry to get the DP sort code
                                ilRifDPSort = 32000         'no DP sort Code in Items record, make it last
                                For llRif = LBound(tlRif) To UBound(tlRif) - 1
                                    If tlRif(llRif).iRdfCode = tlAvRdf(ilRdf).iCode Then
                                        ilRifDPSort = tlRif(llRif).iSort
                                        Exit For
                                    End If
                                Next llRif
                                If tlAvRdf(ilRdf).sInOut = "I" Then   'Book into
                                    If tmAvail.ianfCode <> tlAvRdf(ilRdf).ianfCode Then
                                        ilAvailOk = False
                                    End If
                                ElseIf tlAvRdf(ilRdf).sInOut = "O" Then   'Exclude
                                    If tmAvail.ianfCode = tlAvRdf(ilRdf).ianfCode Then
                                        ilAvailOk = False
                                    End If
                                End If

                                 '7-19-04 the Named avail property must allow local spots to be included
                                tmAnfSrchKey.iCode = tmAvail.ianfCode
                                ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                If (ilRet = BTRV_ERR_NONE) Then
                                    If Not (tlCntTypes.iHold Or tlCntTypes.iOrder) And tmAnf.sBookLocalFeed = "L" Then      'Local avail requested to be excluded, exclude if avail type = "L"
                                        ilAvailOk = False
                                    End If
                                    If Not tlCntTypes.iNetwork And tmAnf.sBookLocalFeed = "F" Then      'Network avail requested to be excluded, exclude if avail type = "F"
                                        ilAvailOk = False
                                    End If
                                End If
                                'allow the avail to be gathered if the field doesnt have a value, indicating an original avail defined as Both
                                'allow the avail to be gathered even if the named avail code isnt found
                            End If
                            If ilAvailOk Then
                                ilRemLen = tmAvail.iLen
                                ilRemUnits = tmAvail.iAvInfo And &H1F

                                'Determine if Grf created - create 1 record per day per daypart
                                ilFound = False
                                For ilRec = 0 To UBound(tmGrf) - 1 Step 1
                                    'match on daypart code and day of week
                                    'If (ilRdfCodes(ilRec) = tlAvRdf(ilRdf).iCode) And (tmGrf(ilRec).iYear = ilBucketIndex) Then
                                    If (tmGrf(ilRec).iRdfCode = tlAvRdf(ilRdf).iCode) And (tmGrf(ilRec).iYear = ilBucketIndex) Then
                                        ilFound = True
                                        ilRecIndex = ilRec
                                        Exit For
                                    End If
                                Next ilRec
                                If Not ilFound Then
                                    ilRecIndex = UBound(tmGrf)
                                    tmGrf(ilRecIndex).iGenDate(0) = igNowDate(0)
                                    tmGrf(ilRecIndex).iGenDate(1) = igNowDate(1)
                                    'tmGrf(ilRecIndex).iGenTime(0) = igNowTime(0)
                                    'tmGrf(ilRecIndex).iGenTime(1) = igNowTime(1)
                                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                    tmGrf(ilRecIndex).lGenTime = lgNowTime
                                    tmGrf(ilRecIndex).iVefCode = ilVefCode
                                    tmGrf(ilRecIndex).iYear = ilBucketIndex        'day of week
                                    tmGrf(ilRecIndex).iStartDate(0) = ilDate0       'date of avails
                                    tmGrf(ilRecIndex).iStartDate(1) = ilDate1
                                    tmGrf(ilRecIndex).iAdfCode = ilRifDPSort      'Use sort code from Items record, not Daypart recd
                                    'ilRdfCodes(ilRecIndex) = tlAvRdf(ilRdf).iCode
                                    tmGrf(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iCode
                                    ReDim Preserve tmGrf(0 To ilRecIndex + 1) As GRF
                                    'ReDim Preserve ilRdfCodes(0 To ilRecIndex + 1)
                                End If

                                'Always gather inventory
                                If tgVpf(ilVpfIndex).sSSellOut = "B" Or tgVpf(ilVpfIndex).sSSellOut = "M" Then       'units & seconds or exact match
                                    ilAddCategory = ilRemLen
                                Else                                            'units
                                    ilAddCategory = ilRemUnits
                                End If

                                'Count of Inventory
                                'tmGrf(ilRecIndex).lDollars(1) = tmGrf(ilRecIndex).lDollars(1) + ilAddCategory
                                tmGrf(ilRecIndex).lDollars(0) = tmGrf(ilRecIndex).lDollars(0) + ilAddCategory

                                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                   LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                                    ilSpotOK = True                             'assume spot is OK to include
                                    ilSpotTypes = 0
                                    ilCTypes = 0
                                    ilRanks = 0

                                    If (tmSpot.iRank And RANKMASK) = DIRECTRESPONSERANK Then      'DR
                                        ilCTypes = &H40
                                        If Not tlCntTypes.iDR Then
                                            ilSpotOK = False
                                        End If
                                    ElseIf (tmSpot.iRank And RANKMASK) = REMNANTRANK Then
                                        ilCTypes = &H20 '&H10
                                        If Not tlCntTypes.iRemnant Then
                                            ilSpotOK = False
                                        End If

                                    ElseIf (tmSpot.iRank And RANKMASK) = PERINQUIRYRANK Then    'PI
                                        ilCTypes = &H80 ' &H800
                                        If Not tlCntTypes.iPI Then
                                            ilSpotOK = False
                                        End If

                                    ElseIf (tmSpot.iRank And RANKMASK) = TRADERANK Then  'trades
                                        ilCTypes = &H400
                                        If Not tlCntTypes.iTrade Then
                                            ilSpotOK = False
                                        End If
                                    '3-24-03 removed, see mfilterspot
                                    'ElseIf tmSpot.iRank = 1045 Then
                                    '    ilSpotTypes = &H20
                                    '    If Not tlCntTypes.iXtra Then
                                    '        ilSpotOK = False
                                    '    End If

                                    ElseIf (tmSpot.iRank And RANKMASK) = PROMORANK Then  'promo
                                        ilCTypes = &H200
                                        If Not tlCntTypes.iPromo Then
                                            ilSpotOK = False
                                        End If

                                    ElseIf (tmSpot.iRank And RANKMASK) = PSARANK Then  'psa
                                        ilCTypes = &H100
                                        If Not tlCntTypes.iPSA Then
                                            ilSpotOK = False
                                        End If
                                    End If


                                    If ilSpotOK Then                            'continue testing other filters
                                        tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)

                                        If ilRet <> BTRV_ERR_NONE Then
                                            ilSpotOK = False                    'invalid sdf code
                                        End If

                                        'Test for Feed Spot
                                        If ilRet = BTRV_ERR_NONE And tmSdf.lChfCode = 0 And ilSpotOK Then        'feed spot
                                            'obtain the network information
                                            tmFSFSrchKey.lCode = tmSdf.lFsfCode
                                            ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                            If ilRet <> BTRV_ERR_NONE Or Not tlCntTypes.iNetwork Then
                                                ilSpotOK = False                    'invalid network code
                                            End If

                                            ilCTypes = &H4          'flag as Feed spot
                                            slChfType = ""          'contract types dont apply with feed spots
                                            slChfStatus = ""       'status types dont apply with feed spots
                                        Else            'Test for contract spots
                                            If ilRet = BTRV_ERR_NONE And tmSdf.lChfCode > 0 And ilSpotOK Then
                                                'obtain contract info
                                                If tmSpot.lSdfCode = tmSdf.lCode And ilRet = BTRV_ERR_NONE And ilSpotOK = True Then
                                                    If tmSdf.lChfCode <> tmChf.lCode Then                      'if already in mem, don't reread
                                                        tmChfSrchKey.lCode = tmSdf.lChfCode
                                                        ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                                        If ilRet <> BTRV_ERR_NONE Then
                                                            ilSpotOK = False
                                                        End If
                                                    End If
                                                    slChfType = tmChf.sType
                                                    slChfStatus = tmChf.sStatus
                                                    mFilterSpot ilVefCode, tlCntTypes, ilSpotOK, ilCTypes, ilSpotTypes

                                                    If ilSpotOK Then
                                                        If (tmClf.lChfCode <> tmSdf.lChfCode) Then     'got the matching line reference?
                                                            ilSpotOK = False
                                                        Else
                                                            'Determine the Rank from the Daypart
                                                            tmRdfSrchKey.iCode = tmClf.iRdfCode
                                                            ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching r/c recd
                                                            If ilRet <> BTRV_ERR_NONE Then
                                                                ilSpotOK = False
                                                            End If
                                                            'Fixed time is <= 15M on Daypart or line Override times .  If Fixed time & sponsorship, add counts to fixed time
                                                            ilFound = False
                                                            If ((tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And ((tmClf.iEndTime(0) <> 1) Or (tmClf.iEndTime(1) <> 0)) Then
                                                                gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llStartTime
                                                                gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llEndTime
                                                            Else
                                                                'gUnpackTimeLong tmRdf.iStartTime(0, 7), tmRdf.iStartTime(1, 7), False, llStartTime
                                                                'gUnpackTimeLong tmRdf.iEndTime(0, 7), tmRdf.iEndTime(1, 7), True, llEndTime
                                                                gUnpackTimeLong tmRdf.iStartTime(0, 6), tmRdf.iStartTime(1, 6), False, llStartTime
                                                                gUnpackTimeLong tmRdf.iEndTime(0, 6), tmRdf.iEndTime(1, 6), True, llEndTime
                                                            End If
                                                            If llTime >= llStartTime And llTime <= llEndTime Then   'spot is scheduled within the override and/or DP times; if its not, consider it a DP count
                                                                If (llEndTime - llStartTime) <= 900 Then          '15M
                                                                    ilFound = True
                                                                    ilRanks = &H1
                                                                    If Not tlCntTypes.iFixedTime Then
                                                                        ilSpotOK = False
                                                                    End If
                                                                End If
                                                            Else        'spot outside its DP or override times
                                                                ilFound = True
                                                                ilRanks = &H4           'DP count
                                                                If Not tlCntTypes.iDP Then
                                                                    ilSpotOK = False
                                                                End If
                                                            End If

                                                            'test for Sponsor
                                                            If Not ilFound Then
                                                                If tmRdf.sInOut = "I" Then   'Book into
                                                                    If tmAvail.ianfCode = tmRdf.ianfCode Then
                                                                        ilFound = True
                                                                        ilRanks = &H2
                                                                        If Not tlCntTypes.iSponsor Then
                                                                            ilSpotOK = False
                                                                        End If
                                                                    Else            'sponsship spot not in matching avail, force to DP count
                                                                        ilFound = True
                                                                        ilRanks = &H4
                                                                        If Not tlCntTypes.iDP Then
                                                                            ilSpotOK = False
                                                                        End If
                                                                    End If
                                                                End If
                                                                If Not ilFound Then      'not sponsorship or fixed time
                                                                    If tlRif(0).sBase <> "Y" Then   'ROS
                                                                        ilRanks = &H8
                                                                        If Not tlCntTypes.iROS Then
                                                                            ilSpotOK = False
                                                                        End If
                                                                    Else
                                                                        ilRanks = &H4
                                                                        If Not tlCntTypes.iDP Then
                                                                            ilSpotOK = False
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If


                                                End If
                                            End If
                                        End If


                                        If ilSpotOK Then
                                            If tgVpf(ilVpfIndex).sSSellOut = "B" Or tgVpf(ilVpfIndex).sSSellOut = "M" Then       'units & seconds or exact match
                                                ilAddCategory = tmSdf.iLen              'spot length
                                            Else                                            'units
                                                ilAddCategory = 1                           'each spot worth 1 unit
                                            End If

                                            ilRemLen = ilRemLen - tmSdf.iLen   'keep running total of whats remaining in avail based on spots used in stats
                                            ilRemUnits = ilRemUnits - 1

                                            'Accumulate statistics for the spot and write day/vehicle to disk
                                            'tmGrf(ilRecIndex).lDollars(2) = tmGrf(ilRecIndex).lDollars(2) + ilAddCategory       'booked
                                            tmGrf(ilRecIndex).lDollars(1) = tmGrf(ilRecIndex).lDollars(1) + ilAddCategory       'booked
                                            If ilRanks = &H1 Then           'fixed time
                                                'tmGrf(ilRecIndex).lDollars(5) = tmGrf(ilRecIndex).lDollars(5) + ilAddCategory       'Fixed time
                                                tmGrf(ilRecIndex).lDollars(4) = tmGrf(ilRecIndex).lDollars(4) + ilAddCategory       'Fixed time
                                            ElseIf ilRanks = &H2 Then       'Sponsor
                                                'tmGrf(ilRecIndex).lDollars(6) = tmGrf(ilRecIndex).lDollars(6) + ilAddCategory       'sponsor
                                                tmGrf(ilRecIndex).lDollars(5) = tmGrf(ilRecIndex).lDollars(5) + ilAddCategory       'sponsor
                                            ElseIf ilRanks = &H4 Then       'DP
                                                'Is it hi or lo priority
                                                If tmClf.sPreempt = "N" Then    'hi priority
                                                    'tmGrf(ilRecIndex).lDollars(7) = tmGrf(ilRecIndex).lDollars(7) + ilAddCategory       'hi DP
                                                    tmGrf(ilRecIndex).lDollars(6) = tmGrf(ilRecIndex).lDollars(6) + ilAddCategory       'hi DP
                                                Else
                                                    'tmGrf(ilRecIndex).lDollars(8) = tmGrf(ilRecIndex).lDollars(8) + ilAddCategory       'lo DP
                                                    tmGrf(ilRecIndex).lDollars(7) = tmGrf(ilRecIndex).lDollars(7) + ilAddCategory       'lo DP
                                                End If
                                            Else
                                                'ROS
                                                'tmGrf(ilRecIndex).lDollars(9) = tmGrf(ilRecIndex).lDollars(9) + ilAddCategory       'ROS
                                                tmGrf(ilRecIndex).lDollars(8) = tmGrf(ilRecIndex).lDollars(8) + ilAddCategory       'ROS
                                            End If

                                            If (ilSpotTypes = &H20) Or (ilSpotTypes = &H40) Then    'Xtra or fills
                                                'tmGrf(ilRecIndex).lDollars(13) = tmGrf(ilRecIndex).lDollars(13) + ilAddCategory
                                                tmGrf(ilRecIndex).lDollars(12) = tmGrf(ilRecIndex).lDollars(12) + ilAddCategory
                                            ElseIf ilCTypes = &H4 Then      'network
                                                'tmGrf(ilRecIndex).lDollars(14) = tmGrf(ilRecIndex).lDollars(14) + ilAddCategory       'Network
                                                tmGrf(ilRecIndex).lDollars(13) = tmGrf(ilRecIndex).lDollars(13) + ilAddCategory       'Network
                                            ElseIf ilCTypes = &H400 Then
                                                'tmGrf(ilRecIndex).lDollars(10) = tmGrf(ilRecIndex).lDollars(10) + ilAddCategory       'Trades
                                                tmGrf(ilRecIndex).lDollars(9) = tmGrf(ilRecIndex).lDollars(9) + ilAddCategory       'Trades
                                            ElseIf ilCTypes = &H10 Then
                                                'tmGrf(ilRecIndex).lDollars(11) = tmGrf(ilRecIndex).lDollars(11) + ilAddCategory       'Reserv
                                                tmGrf(ilRecIndex).lDollars(10) = tmGrf(ilRecIndex).lDollars(10) + ilAddCategory       'Reserv
                                            ElseIf (ilCTypes = &H40) Or (ilCTypes = &H80) Or (ilCTypes = &H20) Or (ilCTypes = &H100) Or (ilCTypes = &H200) Then 'DR, PI, Remnant, PSa or Promo
                                                'tmGrf(ilRecIndex).lDollars(12) = tmGrf(ilRecIndex).lDollars(12) + ilAddCategory       '
                                                tmGrf(ilRecIndex).lDollars(11) = tmGrf(ilRecIndex).lDollars(11) + ilAddCategory       '
                                            End If
                                        End If                              'ilspotOK
                                    End If
                                Next ilSpot                                 'loop from ssf file for # spots in avail
                                'Determine the counts for highlighted (max 4) spot lengths
                                For ilLoop = 0 To 3
                                    Do While (ilRemLen >= tlCntTypes.iLenHL(ilLoop) And ilRemUnits > 0) And (tlCntTypes.iLenHL(ilLoop) <> 0)
                                        'tmGrf(ilRecIndex).iPerGenl(ilLoop + 1) = tmGrf(ilRecIndex).iPerGenl(ilLoop + 1) + 1  'accum count of highlighted spot lengths
                                        tmGrf(ilRecIndex).iPerGenl(ilLoop) = tmGrf(ilRecIndex).iPerGenl(ilLoop) + 1   'accum count of highlighted spot lengths
                                        ilRemLen = ilRemLen - tlCntTypes.iLenHL(ilLoop)
                                        ilRemUnits = ilRemUnits - 1
                                    Loop
                                Next ilLoop
                            End If                                          'Avail OK
                        Next ilRdf                                          'ilRdf = lBound(tlAvRdf)
                        ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                    End If
                    ilEvt = ilEvt + 1   'Increment to next event
                Loop                                                        'do while ilEvt <= tmSsf.iCount
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                If tgMVef(ilVefIndex).sType = "G" Then
                    ilType = tmSsf.iType
                End If
            Loop
        End If
    Next llLoopDate

    'Get missed
    '3/30/99 For each missed status (missed, ready, & unscheduled) there are up to 2 passes
    'for each spot.  The 1st pass looks or a daypart that matches the shedule lines DP.
    'If found, the missed spot is placed in that DP (if that DP is to be shown on the report).
    'If no DP are found that match, the 2nd pass places it in the first DP that surrounds
    'the missed spots time.
    sldate = Format$(llSDate, "m/d/yy")
    gPackDate sldate, ilDate0, ilDate1

    If (tlCntTypes.iMissed) Then
        'Key 2: VefCode; SchStatus; AdfCode; Date, Time
        For ilPass = 0 To 2 Step 1
            tmSdfSrchKey2.iVefCode = ilVefCode
            If ilPass = 0 Then
                slType = "M"
            ElseIf ilPass = 1 Then
                slType = "R"
            ElseIf ilPass = 2 Then
                slType = "U"
            End If
            tmSdfSrchKey2.sSchStatus = slType
            tmSdfSrchKey2.iAdfCode = 0
            tmSdfSrchKey2.iDate(0) = ilDate0
            tmSdfSrchKey2.iDate(1) = ilDate1
            tmSdfSrchKey2.iTime(0) = 0
            tmSdfSrchKey2.iTime(1) = 0
            ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
            'This code added as replacement for Ext operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.sSchStatus = slType)
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), lldate
                ilFilterDay = gWeekDayLong(lldate)

                If (lldate >= llSDate And lldate <= llEDate) And (tlCntTypes.iValidDays(ilFilterDay)) And ((tlCntTypes.iCntrSpots = True And tmSdf.lChfCode > 0) Or (tlCntTypes.iNetwork = True And tmSdf.lChfCode = 0)) Then
                    ilBucketIndex = gWeekDayLong(lldate)
                    ilBucketIndex = gWeekDayLong(lldate)
                    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                    For ilOrphanMissedLoop = 1 To 2
                        ilOrphanFound = False
                        For ilRdf = LBound(tlAvRdf) To UBound(tlAvRdf) - 1 Step 1
                            ilAvailOk = False
                            If (tlAvRdf(ilRdf).iLtfCode(0) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(1) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(2) <> 0) Then
                                If (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(0)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Then
                                    ilAvailOk = False    'True- code later
                                End If
                            Else
                                For ilLoop = LBound(tlAvRdf(ilRdf).iStartTime, 2) To UBound(tlAvRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                    If (tlAvRdf(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tlAvRdf(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                                        gUnpackTimeLong tlAvRdf(ilRdf).iStartTime(0, ilLoop), tlAvRdf(ilRdf).iStartTime(1, ilLoop), False, llStartTime
                                        gUnpackTimeLong tlAvRdf(ilRdf).iEndTime(0, ilLoop), tlAvRdf(ilRdf).iEndTime(1, ilLoop), True, llEndTime
                                        If UBound(tlAvRdf) - 1 = LBound(tlAvRdf) Then   'could be a conv bumped spot sched in
                                                                                    'in conven veh.  The VV has DP times different than the
                                                                                    'conven veh.
                                            llStartTime = llTime
                                            llEndTime = llTime + 1              'actual time of spot
                                        End If
                                        'Don't include the end time i.e. 10a-3p is 10a thru 2:59:59p
                                        ilLoopIndex = 1     '11-11-99 day spotmissed isnt valid for DP to be shown
                                        'If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilBucketIndex + 1) = "Y") Then
                                        If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilBucketIndex) = "Y") Then
                                            ilAvailOk = True
                                            ilLoopIndex = ilLoop
                                            slDays = ""
                                            For ilDayIndex = 1 To 7 Step 1
                                                If (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "Y") Or (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "N") Then
                                                    slDays = slDays & tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1)
                                                Else
                                                    slDays = slDays & "N"
                                                End If
                                            Next ilDayIndex
                                            Exit For
                                        End If
                                    End If
                                Next ilLoop
                            End If
                            If ilAvailOk Then   'Or ilOrphanMissedLoop = 3 Then     '6-6-00
                                'find the associated Items entry to get the DP sort code
                                ilRifDPSort = 32000         'no DP sort Code in Items record, make it last
                                For llRif = LBound(tlRif) To UBound(tlRif) - 1
                                    If tlRif(llRif).iRdfCode = tlAvRdf(ilRdf).iCode Then
                                        ilRifDPSort = tlRif(llRif).iSort
                                        Exit For
                                    End If
                                Next llRif

                                ilSpotOK = True                'assume spot is OK
                                If tmSdf.lChfCode = 0 Then
                                    'obtain the network information
                                    tmFSFSrchKey.lCode = tmSdf.lFsfCode
                                    ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                    If ilRet <> BTRV_ERR_NONE Or Not tlCntTypes.iNetwork Then
                                        ilSpotOK = False                    'invalid network code
                                    End If
                                    slChfType = ""          'contract types dont apply with feed spots
                                    slChfStatus = ""       'status types dont apply with feed spots
                                Else

                                   ilRet = BTRV_ERR_NONE
                                   tmClfSrchKey.lChfCode = tmSdf.lChfCode
                                   tmClfSrchKey.iLine = tmSdf.iLineNo
                                   tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
                                   tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                                   ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                   If (tmClf.lChfCode <> tmSdf.lChfCode) Then     'got the matching line reference?
                                       ilSpotOK = False
                                   Else
                                       If ilOrphanMissedLoop = 1 Then
                                           If tmClf.iRdfCode <> tlAvRdf(ilRdf).iCode Then
                                               ilSpotOK = False
                                           End If
                                       Else
                                           ilOrphanFound = True
                                       End If
                                   End If
                                   ilRet = BTRV_ERR_NONE
                                   If tmSdf.lChfCode <> tmChf.lCode Then               'if already in mem, don't reread
                                       tmChfSrchKey.lCode = tmSdf.lChfCode
                                       ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                       If ilRet <> BTRV_ERR_NONE Then
                                           ilSpotOK = False
                                       End If
                                   End If
                                   If tmChf.sType = "T" Then
                                       ilSpotTypes = &H10
                                       If Not tlCntTypes.iRemnant Then
                                           ilSpotOK = False
                                       End If

                                   ElseIf tmChf.sType = "Q" Then
                                       ilSpotTypes = &H800
                                       If Not tlCntTypes.iPI Then
                                           ilSpotOK = False
                                       End If

                                   ElseIf tmChf.iPctTrade = 100 Then
                                       ilSpotTypes = &H400
                                       If Not tlCntTypes.iTrade Then
                                           ilSpotOK = False
                                       End If

                                   ElseIf tmSdf.sSpotType = "X" Then
                                       ilSpotTypes = &H20
                                       If Not tlCntTypes.iXtra Then
                                           ilSpotOK = False
                                       End If

                                   ElseIf tmChf.sType = "M" Then
                                       ilSpotTypes = &H20
                                       If Not tlCntTypes.iPromo Then
                                           ilSpotOK = False
                                       End If

                                   ElseIf tmChf.sType = "S" Then
                                       ilSpotTypes = &H100
                                       If Not tlCntTypes.iPSA Then
                                           ilSpotOK = False
                                       End If
                                   End If

                                   mFilterSpot ilVefCode, tlCntTypes, ilSpotOK, ilCTypes, ilSpotTypes
                                End If

                                If ilSpotOK Then
                                    ilOrphanFound = True
                                    'Determine if Avr created
                                    ilFound = False
                                    For ilRec = 0 To UBound(tmGrf) - 1 Step 1
                                        'match on daypart and day of week
                                        'If (ilRdfCodes(ilRec) = tlAvRdf(ilRdf).iCode) And (tmGrf(ilRec).iYear = ilBucketIndex) Then
                                        If (tmGrf(ilRec).iRdfCode = tlAvRdf(ilRdf).iCode) And (tmGrf(ilRec).iYear = ilBucketIndex) Then
                                            ilFound = True
                                            ilRecIndex = ilRec
                                            Exit For
                                        End If
                                    Next ilRec
                                    If Not ilFound Then
                                        ilRecIndex = UBound(tmGrf)
                                        tmGrf(ilRecIndex).iGenDate(0) = igNowDate(0)
                                        tmGrf(ilRecIndex).iGenDate(1) = igNowDate(1)
                                        'tmGrf(ilRecIndex).iGenTime(0) = igNowTime(0)
                                        'tmGrf(ilRecIndex).iGenTime(1) = igNowTime(1)
                                        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                        tmGrf(ilRecIndex).lGenTime = lgNowTime
                                        tmGrf(ilRecIndex).iVefCode = ilVefCode
                                        tmGrf(ilRecIndex).iYear = ilBucketIndex        'day of week
                                        '7-12-01 added next 3 fields for case when no programming and only missed spots exists
                                        'show programming on date missed spots found

                                        'tmGrf(ilRecIndex).iStartDate(0) = ilDate0       'date of avails
                                        'tmGrf(ilRecIndex).iStartDate(1) = ilDate1
                                        '9-15-08  use the date of the spot where missed found, not start of week
                                        tmGrf(ilRecIndex).iStartDate(0) = tmSdf.iDate(0)
                                        tmGrf(ilRecIndex).iStartDate(1) = tmSdf.iDate(1)
                                        tmGrf(ilRecIndex).iAdfCode = ilRifDPSort      'Use sort code from Items record, not Daypart recd

                                        'ilRdfCodes(ilRecIndex) = tlAvRdf(ilRdf).iCode
                                        tmGrf(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iCode
                                        ReDim Preserve tmGrf(0 To ilRecIndex + 1) As GRF
                                        'ReDim Preserve ilRdfCodes(0 To ilRecIndex + 1)
                                    End If
                                    If tgVpf(ilVpfIndex).sSSellOut = "B" Or tgVpf(ilVpfIndex).sSSellOut = "M" Then        'units & seconds  or exact match
                                        ilAddCategory = tmSdf.iLen              'spot length
                                    Else                                            'units
                                        ilAddCategory = 1                           'each spot worth 1 unit
                                    End If

                                    If ilSpotOK Then
                                        'tmGrf(ilRecIndex).lDollars(3) = tmGrf(ilRecIndex).lDollars(3) + ilAddCategory       'Missed
                                        tmGrf(ilRecIndex).lDollars(2) = tmGrf(ilRecIndex).lDollars(2) + ilAddCategory       'Missed

                                        Exit For                'force exit on this missed if found a matching daypart
                                    End If
                                End If                      'ilSpotOK
                            End If                          'ilAvailOK
                            'If ilOrphanMissedLoop = 3 Then  '6-6-00
                            '    Exit For
                            'End If
                        Next ilRdf
                        If ilOrphanFound Then
                            Exit For
                        End If
                    Next ilOrphanMissedLoop
                End If
                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next ilPass
    End If

    Erase ilEvtType
    'Erase ilRdfCodes
    Erase tlLLC
End Sub
'
'
'
'               mFilterSpot - Test header and line exclusions for user request.
'
'               <input> ilVefCode - airing vehicle
'                       tlCntTypes - structure of inclusions/exclusions of contract types & status
'              <output> ilSpotOk - true if spot is OK, else false to ignore spot
'                       ilCTypes - set bit to 1 for the matching type
'                       ilSpotTypes - set bit to 1 for matching type
'
Sub mFilterSpot(ilVefCode As Integer, tlCntTypes As CNTTYPES, ilSpotOK As Integer, ilCTypes As Integer, ilSpotTypes As Integer)
Dim ilRet As Integer
Dim slPrice As String

    If ilSpotOK Then
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
            ilCTypes = &H8
            If Not tlCntTypes.iStandard Then       'include Standard types?
                ilSpotOK = False
            End If

        ElseIf tmChf.sType = "V" Then
            ilCTypes = &H10
            If Not tlCntTypes.iReserv Then      'include reservations ?
                ilSpotOK = False
            End If

        ElseIf tmChf.sType = "R" Then
            ilCTypes = &H20
            If Not tlCntTypes.iDR Then       'include DR?
                ilSpotOK = False
            End If
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
                ilSpotTypes = &H4
                If Not tlCntTypes.iZero Then
                    ilSpotOK = False
                End If
            Else
                ilSpotTypes = &H2
                If Not tlCntTypes.iCharge Then           'exclude charged spots
                    ilSpotOK = False
                End If
            End If
        ElseIf Trim$(slPrice) = "ADU" Then
            ilSpotTypes = &H8
            If Not tlCntTypes.iADU Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "Bonus" Then
            ilSpotTypes = &H10
            If Not tlCntTypes.iBonus Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "+ Fill" Then       '3-24-03
            ilSpotTypes = &H20
            If Not tlCntTypes.iXtra Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "- Fill" Then        '3-24-03
            ilSpotTypes = &H40
            If Not tlCntTypes.iFill Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "N/C" Then
            ilSpotTypes = &H80
            If Not tlCntTypes.iNC Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "Recapturable" Then
            ilSpotTypes = &H100
            If Not tlCntTypes.iRecapturable Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "Spinoff" Then
            ilSpotTypes = &H200
            If Not tlCntTypes.iSpinoff Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "MG" Then               '10-27-10
            ilSpotTypes = &H400
            If Not tlCntTypes.iMG Then
                ilSpotOK = False
            End If
        End If
        
        '10-27-10  if excluding MG, that includes MG rate spot types & MG/outside spots
        'test for mg/outside (cant be a bonus spot) and if MG should be included
        If ((tmSdf.sSchStatus = "G" Or tmSdf.sSchStatus = "O") And tmSdf.sSpotType <> "X") And (Not tlCntTypes.iMG) Then
            ilSpotOK = False
        End If

    End If                                  'ilspotOK
End Sub
