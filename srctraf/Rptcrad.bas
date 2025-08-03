Attribute VB_Name = "RPTCRAD"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrad.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  imUseCode                                                                             *
'******************************************************************************************

Option Explicit
Option Compare Text
'Type ACTIVECNTS
'    sKey As String * 8          'contract code left filled with zeroes for sorting
'    lChfCode As Long            'contract code
'    lStartDate As Long          'contract start date
'    lEndDate As Long            'contract end date
'    sType As String * 1         'contract type
'    lPop As Long                'population ,-1 = initialized value, 0 = pop varies across spots, else population
'    lContrCost As Long          'total contract cost from schedule lines
'    iMnfDemo As Integer         'primary demo
'    lPledged As Long            'pledged contract aud
'    lContrGrimp As Long           'audience total for all spots (sum of avg aud which = gross impressions0
'    lCharge As Long             'audience charged spots
'    lNC As Long                 'aud no charge spots
'    lFill As Long               'aud fill spots
'    lADU As Long                'aud ADU spots
'    lMissed As Long             'aud missed spots
'    lContrGrp As Long           'total grps per contract
'    lChargeGrp As Long             'audience charged spots
'    lNCGrp As Long                 'aud no charge spots
'    lFillGrp As Long               'aud fill spots
'    lADUGrp As Long                'aud ADU spots
'    lMissedGrp As Long             'aud missed spots
'    lContrSpots As Long           'total Spots per contract
'    lChargeSpots As Long          'Total charged spots
'    lNCSpots As Long              'Total no charge spots
'    lFillSpots As Long            'Total fill spots
'    lADUSpots As Long             'Total ADU spots
'    lMissedSpots As Long          'Total missed spots
'    iBookMissing As Integer     '0 = book exists for every vehicle in contract, 1= at least 1 vehicle doesnt have a book
'    iFirstPopLink As Integer     'first pointer to entry containing the population for a schdule line
'    iVaryPop As Integer         '0 = use the population found or 0 for varying pop across lines; 1 = cant product grps, varying pop within same line
'End Type
'Type POPLINKLIST
'    iLine As Integer          'line #
'    iNextLink As Integer        'pointer to next entry in list, which is another schedule line to a contract
'    lPop As Long                'pop for a schedule line; -1 initialized, 0 = varying pop within the line, non-zero = population for line
'End Type
'Type VEHICLEBOOK
'    iVefCode As Integer         'vehicle code
'    'iDnfFirstLink As Integer    'index into first book for this vehicle
'    'iDnfLastLink As Integer     'index into last book for this vehicle
'    lDnfFirstLink As Long    '1-15-08 chg to long index into first book for this vehicle
'    lDnfLastLink As Long     '1-15-08 chg to long index into last book for this vehicle
'
'End Type
'Type BOOKLIST
'    sKey As String * 5          'date in string form,left filled with zeroes for sort
'    iDnfCode As Integer         '
'    lStartDate As Long          'start date of book
'End Type
'Type DNFLINKLIST
'    idnfInx As Integer
'End Type
'Type SPOTTYPESORTAD
'    sKey As String * 80 'Office Advertiser Contract
'    tSdf As SDF
'End Type
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Public Const TA_LEFT = 0
'Public Const TA_RIGHT = 2
'Public Const TA_CENTER = 6
'Public Const TA_TOP = 0
'Public Const TA_BOTTOM = 8
'Public Const TA_BASELINE = 24
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0    'CHF record image
Dim tmChfSrchKey1 As CHFKEY1
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim tmClfSrchKey As CLFKEY0     'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmDnf As Integer            'Book Name file handle
Dim imDnfRecLen As Integer      '
Dim tmDnf As DNF
Dim tmDnfSrchKey As INTKEY0
Dim hmDrf As Integer            'Demo Research data
Dim imDrfRecLen As Integer
Dim tmDrf As DRF
Dim tmDrfSrchKey1 As DRFKEY1
Dim hmDpf As Integer            'Demo Plus Research data
Dim imDpfRecLen As Integer
Dim tmDpf As DPF
Dim hmDef As Integer            'Demo Plus Research data
Dim hmRaf As Integer
Dim tmRaf As RAF
Dim imRafRecLen As Integer
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim hmSdf As Integer            'Spots file handle
Dim imSdfRecLen As Integer      '
Dim tmSdfSrchKey1 As SDFKEY1    'retrieve by vehicle, date
Dim tmSdfSrchKey2 As SDFKEY2    'retrieve by vehicle, advt
Dim tmSdf As SDF
Dim tmPLSdf() As SPOTTYPESORTAD
Dim hmSmf As Integer            'Spots file handle
Dim imSmfRecLen As Integer      '
Dim tmSmf As SMF
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer
Dim tmVsf As VSF
Dim imVsfRecLen As Integer
Dim tmActiveCnts() As ACTIVECNTS       'Array of all active contracts, where all aud infor will be stored
Dim tmVehicleBook() As VEHICLEBOOK   'array of all vehicle and their associated books
Dim tmBookList() As BOOKLIST        'array of book names
Dim tmDnfLinkList() As DNFLINKLIST  'list of book indices associated with a vehicle
Dim imIncludeCodes As Integer
Dim imUseCodes() As Integer
Dim lmSingleCntr As Long

'********************************************************************************************
'
'       5-5-00        gCrAudDelivery - Prepass for Audience Delivery report
'
'
'       produce the pre-pass to show actual audience delivery of contracts and compare it
'       to the pledge delivery, highlighting the difference, whether under or over, between
'       what is plede and what was delivered
'
'       6-6-04 Implement Demo Estimates into Research
'********************************************************************************************
Sub gCrAudDelivery()
ReDim ilNowTime(0 To 1) As Integer    'end time of run
Dim slStr As String
Dim ilRet As Integer
'user entered parametrs
Dim llStart As Long             'active Start Date entered
Dim slStart As String           'active start date entered
Dim llEnd As Long               'active End date entered
Dim slEnd As String             'active end date entered
Dim llEarliestStart As Long     'earliest start date from all contracts to process
Dim llLatestEnd As Long         'latest end date from all contracts to process
Dim slEarliestStart As String   'earliest start date from all contracts to process
Dim slLatestEnd As String       'latest end date from all contrcts to process
Dim ilCPPCPM As Integer         '0 = CPP option, 1= CPM option
Dim ilAscDesc As Integer        '0 = ascending order, 1 = descending order
Dim ilPledge As Integer         '0 = incl pledged orders only, 1 = non-pledged, 2 = both
Dim ilBook As Integer           '0 = closest book to air dates, 1 = specific, 2 = schedule line book
Dim ilOverUnder As Integer      '0 = over only, 1 = under only, 2 = both
'end of user entered parameters
Dim ilVehicle As Integer        'loop to process spots: gather by one vehicle at a time
Dim ilVefCode As Integer
Dim llContrCode As Long
Dim ilSpotLoop As Integer
Dim ilOk As Integer
Dim ilMin As Integer
Dim ilMax As Integer
Dim ilFound As Integer
Dim ilDay As Integer
Dim ilActiveCntInx As Integer
Dim llTime As Long              'avail time
Dim llOvStartTime As Long
Dim llOvEndTime As Long
Dim llPop As Long               'pop of demo
Dim llAvgAud As Long            'aud for demo
ReDim ilInputDays(0 To 6) As Integer    'days of week flags for avgaud rtn
Dim illoop As Integer
Dim ilLink As Integer
Dim ilKeepLast As Integer
Dim ilUpperPop As Integer
Dim ilBookInx As Integer
Dim ilDnfCode As Integer        'book to use for spot obtained
Dim llDnfDate As Long
Dim llDate As Long
Dim slPrice As String
Dim llPrice As Long
Dim ilClf As Integer            'sched line processing loop
'ReDim llCntStartDates(1 To 2) As Long     'contracts start and end dates to build $ from flight
ReDim llCntStartDates(0 To 2) As Long     'contracts start and end dates to build $ from flight. Index zero ignored
'ReDim llProject(1 To 1) As Long
ReDim llProject(0 To 1) As Long     'Index zero ignored
'****** following required for gAvgAudToLnResearch, calculated for every spot
ReDim llTotalCost(0 To 0) As Long
ReDim llWklyspots(0 To 0) As Long    '1 spot for 1 week for aud routine
ReDim llWklyRates(0 To 0) As Long       'spot price per week
ReDim llWklyAvgAud(0 To 0) As Long      'avg aud perweek
ReDim llWklyPopEst(0 To 0) As Long
'Dim llContrGross As Long         'total cost (or spot cost)
Dim dlContrGross As Double         'total cost (or spot cost)'TTP 10439 - Rerate 21,000,000
Dim llTotalAvgAud As Long      'avg aud per week
ReDim ilWklyRtg(0 To 0) As Integer        'wkly weekly rating
Dim ilAVgRtg As Integer        'avg rating
ReDim llWklyGrimp(0 To 0) As Long 'weekly gross impressions
Dim llTotalGrImp As Long        'total grimps
ReDim llWklyGRP(0 To 0) As Long   'weekly gross rating points
Dim llTotalGRP As Long          'total grps
Dim llTotalCPP As Long          'total CPPS
Dim llTotalCPM As Long          'Total CPMS
Dim llSpots As Long
Dim llGuarPct As Long        '5-27-03 default to 100 if zero so that the
Dim llPopEst As Long
Dim llMinLink As Long            '1-15-08
Dim llMaxLink As Long            '1-15-08
Dim llLoop As Long               '1-15-08
Dim ilError As Integer
Dim ilListIndex As Integer
Dim ilAudFromSource As Integer
Dim llAudFromCode As Long
Dim slPopLineType As String * 1
Dim ilPopRdfCode As Integer
Dim llPopRafCode As Long

    ilListIndex = RptSelAD!lbcRptType.ListIndex      'selected report

    ilError = mOpenDelivery()           'open applicable files
    If ilError Then
        Exit Sub            'at least 1 open error
    End If

    ReDim imUseCodes(0 To 0) As Integer     'not applicable for this report, but sent to common rtn

    '7-23-01 setup global variable to determine if demo plus info exists
    lgDpfNoRecs = btrRecords(hmDpf)
    If lgDpfNoRecs = 0 Then
        lgDpfNoRecs = -1
    End If
    ilCPPCPM = 0                'assume CPP
    If RptSelAD!rbcCPPCPM(1).Value Then
        ilCPPCPM = 1            'cpm
    End If
    '0 = incl pledged orders only, 1 = non-pledged, 2 = both
    If RptSelAD!rbcPledge(0).Value Then     'pledged only
        ilPledge = 0
    ElseIf RptSelAD!rbcPledge(1).Value Then     'non pledge only
        ilPledge = 1
    Else
        ilPledge = 2
    End If
    ilBook = 0                    'use closest book to air dates
    If RptSelAD!rbcBook(1).Value Then       'default book
        ilBook = 1
    ElseIf RptSelAD!rbcBook(2).Value Then       'sched line book
        ilBook = 2
    End If

    ilAscDesc = 0           'assume ascending
    If RptSelAD!rbcSortBy(1).Value Then         'slsp option, which way to subsort?
        'by advertiser, over/under (asc), over/under(desc)
        If RptSelAD!rbcSubsort(2).Value Then
            ilAscDesc = 1                       'descending
        End If
    ElseIf RptSelAD!rbcSortBy(3).Value Then     'sort by descending
        ilAscDesc = 1
    End If
    If RptSelAD!rbcOverUnder(0).Value Then         'Include over only
        ilOverUnder = 0
    ElseIf RptSelAD!rbcOverUnder(1).Value Then  'include under only
        ilOverUnder = 1
    Else
        ilOverUnder = 2                      'include over & under
    End If
'    slStr = RptSelAD!edcSelCFrom.Text               'Active Start Date
    slStr = RptSelAD!CSI_CalFrom.Text               'Active Start Date  9-3-19 use csi cal control vs edit box
    If slStr = "" Then
        slStr = "1/1/1970"                          'get everything for the selected contracts - this is never used for "All advt"
    End If
    llStart = gDateValue(slStr)                     'gather contracts thru this date
'    slStr = RptSelAD!edcSelCTo.Text               'Active Start Date
    slStr = RptSelAD!CSI_CalTo.Text               'Active Start Date
    If slStr = "" Then
        slStr = "12/31/2069"
    End If
    llEnd = gDateValue(slStr)                     'gather contracts thru this date
    slStart = Format$(llStart, "m/d/yy")   'insure the year is in the format, may not have been entered with the date input
    slEnd = Format$(llEnd, "m/d/yy")

    ilRet = mBuildTables(slStart, slEnd, llEarliestStart, llLatestEnd, ilListIndex)  'build table of active contracts, vehicles, books and the list of associated books with each vehicle
    If ilRet <> 0 Then
        Erase tmActiveCnts, tmVehicleBook, tmDnfLinkList, tmPLSdf
        Erase tgCffAD, tgClfAD

        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmDnf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmMnf)
        btrDestroy hmRaf
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmMnf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Exit Sub
    End If
    'Link list of each schedule line/contract with their line # and populations
    ReDim tmpoplinklist(0 To 0) As POPLINKLIST
    'Loop thru each vehicle and gather all the spots required
    slEarliestStart = Format(llEarliestStart, "m/d/yy")
    slLatestEnd = Format(llLatestEnd, "m/d/yy")
    tmClf.lChfCode = 0          'initialize for first time thru and re-entrant
    For ilVehicle = 0 To UBound(tmVehicleBook) - 1
        ilVefCode = tmVehicleBook(ilVehicle).iVefCode

        ReDim tmPLSdf(0 To 0) As SPOTTYPESORTAD
        mObtainSdf ilVefCode, slEarliestStart, slLatestEnd, INDEXKEY1, imIncludeCodes, imUseCodes()
        For ilSpotLoop = 0 To UBound(tmPLSdf) - 1          '5-10-12 was sometimes counting  1 too many spots due to missing -1 on upper bound
            'filter out unwanted contract types .  Spot List (tmPlSdf is in Contract code, line & date order)
            tmSdf = tmPLSdf(ilSpotLoop).tSdf
            llContrCode = tmSdf.lChfCode
            ilMin = LBound(tmActiveCnts)
            ilMax = UBound(tmActiveCnts)
If tmSdf.lCode = 4180721 Or tmSdf.lCode = 480722 Then
ilRet = ilRet
End If
            'if same contract as previous, no need to find it, ilOK contains the acceptance value
            'If (tmClf.lChfCode <> tmSdf.lChfCode) Then
                ilOk = True
                ilActiveCntInx = mBinarySearch(llContrCode, ilMin, ilMax)    'find the matching cntr in the list to process

                If ilActiveCntInx = -1 Then     'not found
                    ilOk = False
                End If
            'End If
            If ilOk Then        'found a matching contr in the active list
                'always ignore: psa, promo, reserveration, PI, DR, and Remnants
                If (tmActiveCnts(ilActiveCntInx).sType = "S") Or (tmActiveCnts(ilActiveCntInx).sType = "M") Or (tmActiveCnts(ilActiveCntInx).sType = "V") Or (tmActiveCnts(ilActiveCntInx).sType = "T") Or (tmActiveCnts(ilActiveCntInx).sType = "R") Or (tmActiveCnts(ilActiveCntInx).sType = "Q") Then
                    ilOk = False
                End If
            End If
            If ilOk Then            'only test to reread line of contract is OK to use (not an excluded type such as PI, DR, etc)
                ilRet = 0
                If (tmClf.lChfCode <> tmSdf.lChfCode) Or (tmClf.lChfCode <> tmSdf.lChfCode Or tmClf.iLine <> tmSdf.iLineNo) Then

                    'get the correctline for this spot
                    tmClfSrchKey.lChfCode = tmSdf.lChfCode
                    tmClfSrchKey.iLine = tmSdf.iLineNo
                    tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                    tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
            End If
            If ilOk Then
                If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then

                    ilRet = gGetFlightPrice(tmSdf, tmClf, hmCff, hmSmf, slPrice)  'get spot price here so that the flight can be accessed in cff

                    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), True, llOvStartTime
                    llOvEndTime = llOvStartTime           'use the avail time for overrides to determine audience
                    For illoop = 0 To 6                 'init all days to not airing, setup for research results later
                        ilInputDays(illoop) = False
                    Next illoop
                    'set day of week aired
                    gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                    slStr = Format(llDate, "m/d/yy")
                    illoop = gWeekDayStr(slStr)     'day index

                    ilInputDays(illoop) = True     'set day of week in week pattern
                    'Determine whether to use the default book or the book closest to airing
                    'Set Default book of vehicle in case not found
                    ilDnfCode = 0
                    'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1
                    '    If ilVefCode = tgMVef(ilLoop).iCode Then
                        illoop = gBinarySearchVef(ilVefCode)
                        If illoop <> -1 Then
                            ilDnfCode = tgMVef(illoop).iDnfCode
                    '        Exit For
                        End If
                    'Next ilLoop
                    'default book has been set in case there isnt a book closest to airing found
                    
                    slPopLineType = tmClf.sType        'for fills, cant have hidden lines
                    ilPopRdfCode = tmClf.iRdfCode       'default for daypart from schedule line
                    llPopRafCode = tmClf.lRafCode       'default for region from schedule line
                    
                    If tmSdf.sSpotType = "X" Or tmSdf.sSchStatus = "G" Or tmSdf.sSchStatus = "O" Then           'fill, mg or outside spots:  use time of spot as overrides, do not use a daypart
                        'use times of spot, considered overrides
                        gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llOvStartTime
                        gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), True, llOvEndTime
                        'day of week for ilInputDays has been set above
                        'Line type (hidden, conventional, etc).  Dont want to ignore the overrides for these types of spots
                        slPopLineType = ""
                        ilPopRdfCode = 0
                        llPopRafCode = 0
                    End If
                    If ilBook = 0 Then      'use closest to airing (vs default book)
                        ilFound = False
                        '1-15-08 change the references to dnflink to long variables
                        llMinLink = tmVehicleBook(ilVehicle).lDnfFirstLink
                        llMaxLink = tmVehicleBook(ilVehicle).lDnfLastLink
                        'The LinkList points to the associated tmBookList entry of this vehicle
                        If llMinLink <> -1 Then          '-1 indicates no books found
                            For llLoop = llMaxLink To llMinLink Step -1
                            'For ilLoop = ilMin To ilMax
                                ilBookInx = tmDnfLinkList(llLoop).idnfInx
                                llDnfDate = tmBookList(ilBookInx).lStartDate
                                If llDate > llDnfDate And llDnfDate <> 0 Then
                                    ilDnfCode = tmBookList(ilBookInx).iDnfCode
                                    ilFound = True
                                    Exit For
                                End If
                            Next llLoop
                        End If
                        If Not ilFound Then
                            tmActiveCnts(ilActiveCntInx).iBookMissing = 1   'Flag error, 1 vehiclemissing an associated book
                        End If
                    ElseIf ilBook = 2 Then          'use the line book for debugging, only if one exists; else use the default book
                        If tmSdf.sSpotType = "X" Or tmSdf.sSchStatus = "G" Or tmSdf.sSchStatus = "O" Then           'fill, use sch line book if same vehicle, otherwise use default book where it is scheduled.  if none, use book closest to airing
                            'times of overrides and days, daypart have been preset for the fill/mg/out spot
                            If tmClf.iVefCode = tmSdf.iVefCode Then         'fill spot schedule in same vehicle as the order line
                                ilDnfCode = tmClf.iDnfCode
                            ElseIf ilDnfCode > 0 Then           'is there a default book  set for the aired spot?
                                ilDnfCode = ilDnfCode       'use it
                            Else                                'no default book defined, use book closest to airing
                                ilFound = False
                                '1-15-08 change the references to dnflink to long variables
                                llMinLink = tmVehicleBook(ilVehicle).lDnfFirstLink
                                llMaxLink = tmVehicleBook(ilVehicle).lDnfLastLink
                                'The LinkList points to the associated tmBookList entry of this vehicle
                                If llMinLink <> -1 Then          '-1 indicates no books found
                                    For llLoop = llMaxLink To llMinLink Step -1
                                    'For ilLoop = ilMin To ilMax
                                        ilBookInx = tmDnfLinkList(llLoop).idnfInx
                                        llDnfDate = tmBookList(ilBookInx).lStartDate
                                        If llDate > llDnfDate And llDnfDate <> 0 Then
                                            ilDnfCode = tmBookList(ilBookInx).iDnfCode
                                            ilFound = True
                                            Exit For
                                        End If
                                    Next llLoop
                                End If
                                If Not ilFound Then
                                    tmActiveCnts(ilActiveCntInx).iBookMissing = 1   'Flag error, 1 vehiclemissing an associated book
                                End If
                            End If
                        Else
                        If tmClf.iDnfCode <> 0 Then
                            ilDnfCode = tmClf.iDnfCode
                            If tmClf.iStartTime(0) = 1 And tmClf.iStartTime(1) = 0 Then
                                llOvStartTime = 0
                                llOvEndTime = 0
                            Else
                                'override times exist
                                gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llOvStartTime
                                gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llOvEndTime
                            End If

                            If tgPriceCff.sDyWk = "W" Then            'weekly
                                For ilDay = 0 To 6 Step 1
                                    If tgPriceCff.iDay(ilDay) > 0 Or tgPriceCff.sXDay(ilDay) = "1" Then
                                        ilInputDays(ilDay) = True
                                    End If
                                Next ilDay
                            Else                                        'daily
                                For ilDay = 0 To 6 Step 1
                                    If tgPriceCff.iDay(ilDay) > 0 Then
                                        ilInputDays(ilDay) = True
                                    End If
                                Next ilDay
                            End If
                        Else
                            tmActiveCnts(ilActiveCntInx).iBookMissing = 1   'Flag error, 1 vehiclemissing an associated book
                        End If
                    End If
                    End If
                    'retain slPrice to accumulate the grps & audience
                    'ilRet = gGetFlightPrice(tmSdf, tmClf, hmCff, hmSmf, slPrice)

                    llPrice = 0     'init incase a decimal number isnt in price field (its adu, nc, fill,etc.)
                    If (InStr(slPrice, ".") <> 0) Then        'found spot cost
                        llPrice = gStrDecToLong(slPrice, 2)
                    End If
                    
                    '7-18-12 For fills, mg, and outsides, book depends upon which option:
                    'if using schedule line book, it will use the default book of the vehicle the spot is scheduled in unless the spot is in the same
                    'vehicle of the ordered line the fill (mg or outside) came from.  If there is no default book, then it uses the book closest to the airing of spot.
                    '
                    ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, 0, tmActiveCnts(ilActiveCntInx).iMnfDemo, llPop)

                    '6-6-04 Get the pops only if not using Demo Estimates
                    If tgSpf.sDemoEstAllowed <> "Y" Then        'not using demo estimates
                        'Set if varying populations within/across schedule lines in contract
                        If tmActiveCnts(ilActiveCntInx).lPop < 0 Then       'never been set up, save population first time
                            If llPop <> 0 Then
                                tmActiveCnts(ilActiveCntInx).lPop = llPop
                            End If
                        ElseIf tmActiveCnts(ilActiveCntInx).lPop <> 0 Then
                            If tmActiveCnts(ilActiveCntInx).lPop <> llPop And llPop <> 0 Then
                                tmActiveCnts(ilActiveCntInx).lPop = 0
                            End If
                        End If


                        'Keep track if varying populations within same schedule line (across days or weeks)
                        'If varying within sch line, cant product grp numbers, in which case it should be zerod
                        ilUpperPop = UBound(tmpoplinklist)
                        If tmActiveCnts(ilActiveCntInx).iFirstPopLink = -1 Then      'first time thru for this contract
                            tmActiveCnts(ilActiveCntInx).iFirstPopLink = ilUpperPop
                            If llPop <> 0 Then
                                tmpoplinklist(ilUpperPop).lPop = llPop
                            End If
                            tmpoplinklist(ilUpperPop).iLine = tmClf.iLine
                            tmpoplinklist(ilUpperPop).iNextLink = -1
                            ReDim Preserve tmpoplinklist(0 To ilUpperPop + 1) As POPLINKLIST
                        Else            'already found at least one spot for this contract , see if varying pops
                            ilUpperPop = UBound(tmpoplinklist)
                            ilFound = False
                            ilLink = tmActiveCnts(ilActiveCntInx).iFirstPopLink
                            Do While ilLink <> -1
                                If tmpoplinklist(ilLink).iLine = tmClf.iLine Then
                                    ilFound = True
                                    If tmpoplinklist(ilLink).lPop < 0 Then       'never been set up, save population first time
                                        If llPop <> 0 Then
                                            tmpoplinklist(ilLink).lPop = llPop
                                        End If
                                    ElseIf tmpoplinklist(ilLink).lPop <> 0 Then
                                        If tmpoplinklist(ilLink).lPop <> llPop And llPop <> 0 Then
                                            tmpoplinklist(ilLink).lPop = 0
                                            tmActiveCnts(ilActiveCntInx).iVaryPop = 1   'cant compute cpps from grps
                                        End If
                                    End If
                                    ilLink = -1     'force to stop search
                                Else                'not matching line #
                                    ilKeepLast = ilLink
                                    ilLink = tmpoplinklist(ilLink).iNextLink
                                End If
                            Loop
                            If Not ilFound Then             'set up new entry
                                tmpoplinklist(ilKeepLast).iNextLink = ilUpperPop
                                If llPop <> 0 Then
                                    tmpoplinklist(ilUpperPop).lPop = llPop
                                End If
                                tmpoplinklist(ilUpperPop).iLine = tmClf.iLine
                                tmpoplinklist(ilUpperPop).iNextLink = -1
                                ReDim Preserve tmpoplinklist(0 To ilUpperPop + 1) As POPLINKLIST
                            End If
                        End If
                    End If

                    '************ DEBUGGING ONLY   *******************
                    'llOVStarttime = 0         use daypart times only, no overrides
                    'llOVEndTime = 0
                    'ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, tmClf.iVefCode, 0, tmActiveCnts(ilActiveCntInx).iMnfDemo, llDate, llDate, tmClf.iRdfcode, llOvStartTime, llOvEndTime, ilInputDays(), tmClf.sType, tmClf.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                    '7-18-12 For fills, mg and outsides:  use time of spot as an override. There is no DP and region code.
                    'Book determine by the option selected:  use default book, use book closest to airing, or use schedule line.  in all cases, it uses the vehicle its scheduled in to retrieve any book.
                    'if using schedule line book, use the book from line if spot scheduled in same vehicle.  If not, use the default book of vehicle, if that doesnt exist, use book closest to airing.
                    'If there are overrrides, no DP, the routine finds the "Best Fit"
                    
                    ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, ilVefCode, 0, tmActiveCnts(ilActiveCntInx).iMnfDemo, llDate, llDate, ilPopRdfCode, llOvStartTime, llOvEndTime, ilInputDays(), slPopLineType, llPopRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                    
                    'need to get grp from gAvgAudToLnResearch
                    'If llPop < 0 Then
                   '     llPop = 0
                   ' End If
                    llWklyspots(0) = 1
                    llWklyAvgAud(0) = llAvgAud
                    llWklyRates(0) = llPrice
                    llWklyPopEst(0) = llPopEst


                    'Use population of the book for the individual spot
                    '10-30-14 default to use 1 place rating regardless of agency flag
                    'gAvgAudToLnResearch "1", True, llPop, llWklyPopEst(), llWklyspots(), llWklyRates(), llWklyAvgAud(), llContrGross, llTotalAvgAud, ilWklyRtg(), ilAVgRtg, llWklyGrimp(), llTotalGrImp, llWklyGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst
                    gAvgAudToLnResearch "1", True, llPop, llWklyPopEst(), llWklyspots(), llWklyRates(), llWklyAvgAud(), dlContrGross, llTotalAvgAud, ilWklyRtg(), ilAVgRtg, llWklyGrimp(), llTotalGrImp, llWklyGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst  'TTP 10439 - Rerate 21,000,000
                    '6-6-04 Get the pops only if not using Demo Estimates
                    If tgSpf.sDemoEstAllowed = "Y" Then        'using demo estimates
                        'Set if varying populations within/across schedule lines in contract
                        If tmActiveCnts(ilActiveCntInx).lPop < 0 Then       'never been set up, save population first time
                            If llPop <> 0 Then
                                tmActiveCnts(ilActiveCntInx).lPop = llPopEst
                            End If
                        ElseIf tmActiveCnts(ilActiveCntInx).lPop <> 0 Then
                            If tmActiveCnts(ilActiveCntInx).lPop <> llPopEst And llPopEst <> 0 Then
                                tmActiveCnts(ilActiveCntInx).lPop = 0
                            End If
                        End If
                        'Keep track if varying populations within same schedule line (across days or weeks)
                        'If varying within sch line, cant product grp numbers, in which case it should be zerod
                        ilUpperPop = UBound(tmpoplinklist)
                        If tmActiveCnts(ilActiveCntInx).iFirstPopLink = -1 Then      'first time thru for this contract
                            tmActiveCnts(ilActiveCntInx).iFirstPopLink = ilUpperPop
                            If llPopEst <> 0 Then
                                tmpoplinklist(ilUpperPop).lPop = llPopEst
                            End If
                            tmpoplinklist(ilUpperPop).iLine = tmClf.iLine
                            tmpoplinklist(ilUpperPop).iNextLink = -1
                            ReDim Preserve tmpoplinklist(0 To ilUpperPop + 1) As POPLINKLIST
                        Else            'already found at least one spot for this contract , see if varying pops
                            ilUpperPop = UBound(tmpoplinklist)
                            ilFound = False
                            ilLink = tmActiveCnts(ilActiveCntInx).iFirstPopLink
                            Do While ilLink <> -1
                                If tmpoplinklist(ilLink).iLine = tmClf.iLine Then
                                    ilFound = True
                                    If tmpoplinklist(ilLink).lPop < 0 Then       'never been set up, save population first time
                                        If llPopEst <> 0 Then
                                            tmpoplinklist(ilLink).lPop = llPopEst
                                        End If
                                    ElseIf tmpoplinklist(ilLink).lPop <> 0 Then
                                        If tmpoplinklist(ilLink).lPop <> llPopEst And llPopEst <> 0 Then
                                            tmpoplinklist(ilLink).lPop = 0
                                            tmActiveCnts(ilActiveCntInx).iVaryPop = 1   'cant compute cpps from grps
                                        End If
                                    End If
                                    ilLink = -1     'force to stop search
                                Else                'not matching line #
                                    ilKeepLast = ilLink
                                    ilLink = tmpoplinklist(ilLink).iNextLink
                                End If
                            Loop
                            If Not ilFound Then             'set up new entry
                                tmpoplinklist(ilKeepLast).iNextLink = ilUpperPop
                                If llPopEst <> 0 Then
                                    tmpoplinklist(ilUpperPop).lPop = llPopEst
                                End If
                                tmpoplinklist(ilUpperPop).iLine = tmClf.iLine
                                tmpoplinklist(ilUpperPop).iNextLink = -1
                                ReDim Preserve tmpoplinklist(0 To ilUpperPop + 1) As POPLINKLIST
                            End If
                        End If
                    End If


                    'overall contract avg aud (or gross impressions)
                    tmActiveCnts(ilActiveCntInx).lContrGrimp = tmActiveCnts(ilActiveCntInx).lContrGrimp + llTotalGrImp
                    tmActiveCnts(ilActiveCntInx).lContrGrp = tmActiveCnts(ilActiveCntInx).lContrGrp + llTotalGRP
                    tmActiveCnts(ilActiveCntInx).lContrSpots = tmActiveCnts(ilActiveCntInx).lContrSpots + 1
                    If tmSdf.sSchStatus = "M" Or tmSdf.sSchStatus = "C" Or tmSdf.sSchStatus = "H" Then
                        tmActiveCnts(ilActiveCntInx).lMissed = tmActiveCnts(ilActiveCntInx).lMissed + llTotalGrImp
                        tmActiveCnts(ilActiveCntInx).lMissedGrp = tmActiveCnts(ilActiveCntInx).lMissedGrp + llTotalGRP
                        tmActiveCnts(ilActiveCntInx).lMissedSpots = tmActiveCnts(ilActiveCntInx).lMissedSpots + 1
                    Else                   'not missed, cancelled or hidden.  Accumulate for Charged, nc, adu orfills
                        If Trim$(slPrice) = ".00" Or Trim$(slPrice) = "N/C" Or Trim$(slPrice) = "Bonus" Or Trim$(slPrice) = "MG" Or Trim$(slStr) = "Spinoff" Then        'its a .00 spot, N/c or bonus
                            tmActiveCnts(ilActiveCntInx).lNC = tmActiveCnts(ilActiveCntInx).lNC + llTotalGrImp
                            tmActiveCnts(ilActiveCntInx).lNCGrp = tmActiveCnts(ilActiveCntInx).lNCGrp + llTotalGRP
                            tmActiveCnts(ilActiveCntInx).lNCSpots = tmActiveCnts(ilActiveCntInx).lNCSpots + 1
                        ElseIf Trim$(slPrice) = "ADU" Or Trim$(slPrice) = "Recapturable" Then
                                tmActiveCnts(ilActiveCntInx).lADU = tmActiveCnts(ilActiveCntInx).lADU + llTotalGrImp
                                tmActiveCnts(ilActiveCntInx).lADUGrp = tmActiveCnts(ilActiveCntInx).lADUGrp + llTotalGRP
                                tmActiveCnts(ilActiveCntInx).lADUSpots = tmActiveCnts(ilActiveCntInx).lADUSpots + 1
                        'ElseIf Trim$(slPrice) = "Extra" Or Trim$(slPrice) = "+Fill" Or Trim$(slPrice) = "-Fill" Then
                        ElseIf Trim$(slPrice) = "Extra" Or InStr(slPrice, "Fill") <> 0 Then             '3-17-12 test for any kind of fill (+/-)
                                tmActiveCnts(ilActiveCntInx).lFill = tmActiveCnts(ilActiveCntInx).lFill + llTotalGrImp
                                tmActiveCnts(ilActiveCntInx).lFillGrp = tmActiveCnts(ilActiveCntInx).lFillGrp + llTotalGRP
                                tmActiveCnts(ilActiveCntInx).lFillSpots = tmActiveCnts(ilActiveCntInx).lFillSpots + 1
                        Else        'charge spot
                            tmActiveCnts(ilActiveCntInx).lContrCost = tmActiveCnts(ilActiveCntInx).lContrCost + llPrice
                            tmActiveCnts(ilActiveCntInx).lCharge = tmActiveCnts(ilActiveCntInx).lCharge + llTotalGrImp
                            tmActiveCnts(ilActiveCntInx).lChargeGrp = tmActiveCnts(ilActiveCntInx).lChargeGrp + llTotalGRP
                            tmActiveCnts(ilActiveCntInx).lChargeSpots = tmActiveCnts(ilActiveCntInx).lChargeSpots + 1
                        End If
                    End If

                End If
            End If
        Next ilSpotLoop
    Next ilVehicle
    'loop thru all contracts to process and get their total contract $.  Cant rely on the gross $ input.
    'only when all spots gathered:  tmVCost(1 to 1) and should be total contract cost , grimps = sum of avg aud (they are the same) , sum of grps is what I have gathered already, # spots = 1
    For illoop = 0 To UBound(tmActiveCnts) - 1
        'Write out the GRF prepass records
        'tmGrf.iGenTime(0) = igNowTime(0)
        'tmGrf.iGenTime(1) = igNowTime(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf.lGenTime = lgNowTime
        tmGrf.iGenDate(0) = igNowDate(0)
        tmGrf.iGenDate(1) = igNowDate(1)
        tmGrf.lChfCode = tmActiveCnts(illoop).lChfCode

        llContrCode = tmActiveCnts(illoop).lChfCode
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfAD, tgClfAD(), tgCffAD())
        gUnpackDateLong tgChfAD.iStartDate(0), tgChfAD.iStartDate(1), llCntStartDates(1)
        gUnpackDateLong tgChfAD.iEndDate(0), tgChfAD.iEndDate(1), llCntStartDates(2)
        llProject(1) = 0
        If llCntStartDates(2) > llCntStartDates(1) Then
            For ilClf = LBound(tgClfAD) To UBound(tgClfAD) - 1
                tmClf = tgClfAD(ilClf).ClfRec
                If tmClf.sType = "S" Or tmClf.sType = "H" Then
                'determine start and end dates of contract to gather all flights and see what
                'the true $ are
                gBuildFlights ilClf, llCntStartDates(), 1, 2, llProject(), 1, tgClfAD(), tgCffAD()
                End If
            Next ilClf
            llSpots = 0
            'Calculate the total contracts Pledged CPP/CPM
            tmActiveCnts(illoop).lContrCost = llProject(1)      'cost of contract from flights
            llTotalCost(0) = tmActiveCnts(illoop).lContrCost

            llWklyGrimp(0) = tgChfAD.lGrImp                       'calc gross impressions from order entered
            'If tmActiveCnts(ilLoop).iVaryPop = 0 Then          'varying populations within same line
                llWklyGRP(0) = tgChfAD.lGRP                           'calc gross ratings from order
            'Else
            '    llWklyGrp(1) = 0
            'End If
            llPop = tmActiveCnts(illoop).lPop           'if pop < 0, it has never been set.
            If llPop < 0 Then
                llPop = 0
            End If
            
            '10-30-14 default to use 1 place rating regardless of agency flag
            'gResearchTotals "1", True, llPop, llTotalCost(), llWklyGrimp(), llWklyGRP(), llSpots, llContrGross, ilAVgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
            gResearchTotals "1", True, llPop, llTotalCost(), llWklyGrimp(), llWklyGRP(), llSpots, dlContrGross, ilAVgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud  'TTP 10439 - Rerate 21,000,000
            If ilCPPCPM = 0 Then        'CPP
                If tmActiveCnts(illoop).iVaryPop = 0 Then
                    'tmGrf.lDollars(1) = llTotalCPP
                    tmGrf.lDollars(0) = llTotalCPP
                Else
                    'tmGrf.lDollars(1) = 0
                    tmGrf.lDollars(0) = 0
                End If
            Else                        'CPM
                'tmGrf.lDollars(1) = llTotalCPM
                tmGrf.lDollars(0) = llTotalCPM
            End If
            If (Asc(tgSpf.sUsingFeatures6) And GUARBYGRIMP) = GUARBYGRIMP Then  'Using Delivery Guarantee
                'Darlene, we need to discuss how to compute the percent with Jim.  For now, we'll just put it out as
                '100%
                  If ilCPPCPM = 0 Then
                    ''tmGrf.lDollars(3) = (tgChfAD.iGuar * tgChfAD.lGRP) / 100
                    'tmGrf.lDollars(3) = tgChfAD.lGRP
                    tmGrf.lDollars(2) = tgChfAD.lGRP
                Else
                    ''tmGrf.lDollars(3) = (tgChfAD.lguar * tgChfAD.lGrImp) / 100
                    'tmGrf.lDollars(3) = tgChfAD.lGrImp
                    tmGrf.lDollars(2) = tgChfAD.lGrImp
                End If

            Else
                'calculate the Pledged amout of grimps/grps based on the gurantee % stored in the header
                llGuarPct = tgChfAD.lGuar
                If llGuarPct = 0 Then
                    llGuarPct = 100
                End If
                If ilCPPCPM = 0 Then
                    ''tmGrf.lDollars(3) = (tgChfAD.iGuar * tgChfAD.lGRP) / 100
                    'tmGrf.lDollars(3) = (llGuarPct * tgChfAD.lGRP) / 100
                    tmGrf.lDollars(2) = (llGuarPct * tgChfAD.lGRP) / 100
                Else
                    ''tmGrf.lDollars(3) = (tgChfAD.lguar * tgChfAD.lGrImp) / 100
                    'tmGrf.lDollars(3) = (llGuarPct * tgChfAD.lGrImp) / 100
                    tmGrf.lDollars(2) = (llGuarPct * tgChfAD.lGrImp) / 100
                End If
            End If
            'Calculate the total Charge grimps or grps
            tmActiveCnts(illoop).lContrCost = llTotalCost(0)
            llWklyGrimp(0) = tmActiveCnts(illoop).lCharge
            'If tmActiveCnts(ilLoop).iVaryPop = 0 Then          'varying populations within same line
                llWklyGRP(0) = tmActiveCnts(illoop).lChargeGrp
            'Else
            '    llWklyGrp(1) = 0
            'End If
            If tmActiveCnts(illoop).lChargeSpots > 0 And llWklyGrimp(0) = 0 Then  'spots exists without any gross impressions
                'tmGrf.iPerGenl(3) = 1       'indicate at least 1 spot has book missing for charged spots
                tmGrf.iPerGenl(2) = 1       'indicate at least 1 spot has book missing for charged spots
            Else
                'tmGrf.iPerGenl(3) = 0
                tmGrf.iPerGenl(2) = 0
            End If
            llPop = tmActiveCnts(illoop).lPop
            
            '10-30-14 default to use 1 place rating regardless of agency flag
            'gResearchTotals "1", True, llPop, llTotalCost(), llWklyGrimp(), llWklyGRP(), llSpots, llContrGross, ilAVgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
            gResearchTotals "1", True, llPop, llTotalCost(), llWklyGrimp(), llWklyGRP(), llSpots, dlContrGross, ilAVgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud  'TTP 10439 - Rerate 21,000,000
            If ilCPPCPM = 0 Then        'CPP
                'tmGrf.lDollars(4) = llTotalGRP
                tmGrf.lDollars(3) = llTotalGRP
            Else                        'CPM
                'tmGrf.lDollars(4) = llTotalGrImp
                tmGrf.lDollars(3) = llTotalGrImp
            End If

            'Calculate the total NC grimps/grps
            tmActiveCnts(illoop).lContrCost = llTotalCost(0)
            llWklyGrimp(0) = tmActiveCnts(illoop).lNC
            'If tmActiveCnts(ilLoop).iVaryPop = 0 Then          'varying populations within same line
                llWklyGRP(0) = tmActiveCnts(illoop).lNCGrp
            'Else
            '    llWklyGrp(1) = 0
            'End If
            If tmActiveCnts(illoop).lNCSpots > 0 And llWklyGrimp(0) = 0 Then  'spots exists without any gross impressions
                'tmGrf.iPerGenl(4) = 1       'indicate at least 1 spot has book missing for no charged spots
                tmGrf.iPerGenl(3) = 1       'indicate at least 1 spot has book missing for no charged spots
            Else
                'tmGrf.iPerGenl(4) = 0
                tmGrf.iPerGenl(3) = 0
            End If

            llPop = tmActiveCnts(illoop).lPop
            
            '10-30-14 default to use 1 place rating regardless of agency flag
            'gResearchTotals "1", True, llPop, llTotalCost(), llWklyGrimp(), llWklyGRP(), llSpots, llContrGross, ilAVgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
            gResearchTotals "1", True, llPop, llTotalCost(), llWklyGrimp(), llWklyGRP(), llSpots, dlContrGross, ilAVgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud  'TTP 10439 - Rerate 21,000,000
            If ilCPPCPM = 0 Then        'CPP
                'tmGrf.lDollars(5) = llTotalGRP
                tmGrf.lDollars(4) = llTotalGRP
            Else                        'CPM
                'tmGrf.lDollars(5) = llTotalGrImp
                tmGrf.lDollars(4) = llTotalGrImp
            End If

            'Calculate the total Fills grimps/grps
            tmActiveCnts(illoop).lContrCost = llTotalCost(0)
            llWklyGrimp(0) = tmActiveCnts(illoop).lFill
            'If tmActiveCnts(ilLoop).iVaryPop = 0 Then          'varying populations within same line
                llWklyGRP(0) = tmActiveCnts(illoop).lFillGrp
           ' Else
           '     llWklyGrp(1) = 0
           ' End If
            If tmActiveCnts(illoop).lFillSpots > 0 And llWklyGrimp(0) = 0 Then  'spots exists without any gross impressions
                'tmGrf.iPerGenl(5) = 1       'indicate at least 1 spot has book missing for fill spots
                tmGrf.iPerGenl(4) = 1       'indicate at least 1 spot has book missing for fill spots
            Else
                'tmGrf.iPerGenl(5) = 0
                tmGrf.iPerGenl(4) = 0
            End If
            llPop = tmActiveCnts(illoop).lPop
            
            '10-30-14 default to use 1 place rating regardless of agency flag
            'gResearchTotals "1", True, llPop, llTotalCost(), llWklyGrimp(), llWklyGRP(), llSpots, llContrGross, ilAVgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
            gResearchTotals "1", True, llPop, llTotalCost(), llWklyGrimp(), llWklyGRP(), llSpots, dlContrGross, ilAVgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud  'TTP 10439 - Rerate 21,000,000
            If ilCPPCPM = 0 Then        'CPP
                'tmGrf.lDollars(6) = llTotalGRP
                tmGrf.lDollars(5) = llTotalGRP
            Else                        'CPM
                'tmGrf.lDollars(6) = llTotalGrImp
                tmGrf.lDollars(5) = llTotalGrImp
            End If

            'Calculate the total ADU  grimps/grps
            tmActiveCnts(illoop).lContrCost = llTotalCost(0)
            llWklyGrimp(0) = tmActiveCnts(illoop).lADU
            'If tmActiveCnts(ilLoop).iVaryPop = 0 Then          'varying populations within same line
                llWklyGRP(0) = tmActiveCnts(illoop).lADUGrp
            'Else
            '    llWklyGrp(1) = 0
            'End If
            If tmActiveCnts(illoop).lADUSpots > 0 And llWklyGrimp(0) = 0 Then  'spots exists without any gross impressions
                'tmGrf.iPerGenl(6) = 1       'indicate at least 1 spot has book missing for ADU spots
                tmGrf.iPerGenl(5) = 1       'indicate at least 1 spot has book missing for ADU spots
            Else
                'tmGrf.iPerGenl(6) = 0
                tmGrf.iPerGenl(5) = 0
            End If
            llPop = tmActiveCnts(illoop).lPop
            '10-30-14 default to use 1 place rating regardless of agency flag
            'gResearchTotals "1", True, llPop, llTotalCost(), llWklyGrimp(), llWklyGRP(), llSpots, llContrGross, ilAVgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
            gResearchTotals "1", True, llPop, llTotalCost(), llWklyGrimp(), llWklyGRP(), llSpots, dlContrGross, ilAVgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud  'TTP 10439 - Rerate 21,000,000
            If ilCPPCPM = 0 Then        'CPP
                'tmGrf.lDollars(7) = llTotalGRP
                tmGrf.lDollars(6) = llTotalGRP
            Else                        'CPM
                'tmGrf.lDollars(7) = llTotalGrImp
                tmGrf.lDollars(6) = llTotalGrImp
            End If

            'Calculate the total Missed grimps/grps
            tmActiveCnts(illoop).lContrCost = llTotalCost(0)
            llWklyGrimp(0) = tmActiveCnts(illoop).lMissed
            'If tmActiveCnts(ilLoop).iVaryPop = 0 Then          'varying populations within same line
                llWklyGRP(0) = tmActiveCnts(illoop).lMissedGrp
            'Else
            '    llWklyGrp(1) = 0
            'End If
            If tmActiveCnts(illoop).lMissedSpots > 0 And llWklyGrimp(0) = 0 Then  'spots exists without any gross impressions
                'tmGrf.iPerGenl(7) = 1       'indicate at least 1 spot has book missing for missed spots
                tmGrf.iPerGenl(6) = 1       'indicate at least 1 spot has book missing for missed spots
            Else
                'tmGrf.iPerGenl(7) = 0
                tmGrf.iPerGenl(6) = 0
            End If
            llPop = tmActiveCnts(illoop).lPop
            '10-30-14 default to use 1 place rating regardless of agency flag
            'gResearchTotals "1", True, llPop, llTotalCost(), llWklyGrimp(), llWklyGRP(), llSpots, llContrGross, ilAVgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud
            gResearchTotals "1", True, llPop, llTotalCost(), llWklyGrimp(), llWklyGRP(), llSpots, dlContrGross, ilAVgRtg, llTotalGrImp, llTotalGRP, llTotalCPP, llTotalCPM, llTotalAvgAud  'TTP 10439 - Rerate 21,000,000
            If ilCPPCPM = 0 Then        'CPP
                'tmGrf.lDollars(8) = llTotalGRP
                tmGrf.lDollars(7) = llTotalGRP
            Else                        'CPM
                'tmGrf.lDollars(8) = llTotalGrImp
                tmGrf.lDollars(7) = llTotalGrImp
            End If
            'tmGrf.lDollars(2) = tgChfAD.lGuar         'guarantee %
            tmGrf.lDollars(1) = tgChfAD.lGuar         'guarantee %
            'Calculate the Over/Under (Accum Charged & N/c & Fills & ADU & Missed) User want to count missed as though it aired since they
            'will make it good within the original audience: amfm)
            'tmGrf.lDollars(9) = ((tmGrf.lDollars(4) + tmGrf.lDollars(5) + tmGrf.lDollars(6) + tmGrf.lDollars(7)) + tmGrf.lDollars(8)) - tmGrf.lDollars(3)
            tmGrf.lDollars(8) = ((tmGrf.lDollars(3) + tmGrf.lDollars(4) + tmGrf.lDollars(5) + tmGrf.lDollars(6)) + tmGrf.lDollars(7)) - tmGrf.lDollars(2)
            'tmGrf.lDollars(10) = tmGrf.lDollars(9)      'this field is used for sorting; if sorting ascending, no need to change it
            tmGrf.lDollars(9) = tmGrf.lDollars(8)      'this field is used for sorting; if sorting ascending, no need to change it
            'if sorting descending by over/under, reverse sign so report doesnt have to do anything special
            '0 = ascending, 1 = descending by over/under
            If ilAscDesc = 1 Then                   'sort the over/under by descending order?
                'tmGrf.lDollars(10) = -tmGrf.lDollars(10)
                tmGrf.lDollars(9) = -tmGrf.lDollars(9)
            End If
            'ldollars(1) = pledged CPP or CPM calculated from the grimps or grps stored in the header
            'ldollars(2) = pledged (guaranteed %) stored in header
            'ldollars(3) = grimps or grps pledged (calc by guaranteed % * pledged grimps or grps from header
            'ldollars(4) = delivered grimps or grps from $ spots
            'ldollars(5) = delivered grimps or grps from nc/bonus/zero/mg/spinoff spots
            'ldollars(6) = delivered grimps or grps from fill spots
            'ldollars(7) = delivered grimps or grps from adu/recapturable spots
            'ldollars(8) = delivered grimps or grps from missed spots
            'ldollars(9) = delivered over/under grimps or grps
            'ldollars(10) - used for sorting in Crystal (if sorting descending, the over/under value is reversed)
            'PerGenl(1) - 1 = book missing (for at least 1 spot in contract)
            'PerGenl(2) 1 = multiple books within line
            'PerGenl(3) - 1 = book missing for at least 1 charge spot
            'PerGenl(4) - 1 = book missing for at least 1 n/c spot
            'Pergenl(5) - 1 = book missing for at least 1 fill spot
            'Pergenl(6) - 1 = book missing for at least 1 adu spot
            'Pergenl(7) - 1 = book missing for at least 1 missed spot
            'tmGrf.iPerGenl(2) = tmActiveCnts(ilLoop).iVaryPop   'flag to indicate varying books (population) within single line
            tmGrf.iPerGenl(1) = tmActiveCnts(illoop).iVaryPop   'flag to indicate varying books (population) within single line
            If tmActiveCnts(illoop).lContrSpots <> 0 Then       'no spots on contract, must be cancel before start
                'tmGrf.iPerGenl(1) = tmActiveCnts(ilLoop).iBookMissing
                tmGrf.iPerGenl(0) = tmActiveCnts(illoop).iBookMissing
                'Test final delivery to see if should be shown
                'If tmGrf.lDollars(9) < 0 And (ilOverUnder = 1 Or ilOverUnder = 2) Then    'under, include only under or both
                If tmGrf.lDollars(8) < 0 And (ilOverUnder = 1 Or ilOverUnder = 2) Then    'under, include only under or both
                    'see if pledged vs non-pledged should be included
                    If (tgChfAD.lGuar = 0 And ilPledge > 0) Or (tgChfAD.lGuar <> 0 And ilPledge <> 1) Then        'no guarantee entered and user asing for pledged only (ignore)
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    End If
                'ElseIf tmGrf.lDollars(9) > 0 And (ilOverUnder = 0 Or ilOverUnder = 2) Then    'over, include only over or both
                ElseIf tmGrf.lDollars(8) > 0 And (ilOverUnder = 0 Or ilOverUnder = 2) Then    'over, include only over or both
                    If (tgChfAD.lGuar = 0 And ilPledge > 0) Or (tgChfAD.lGuar <> 0 And ilPledge <> 1) Then       'no guarantee entered and user asing for pledged only (ignore)
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    End If
                'ElseIf tmGrf.lDollars(9) = 0 And ilOverUnder = 2 Then                   'matches delivery, include all?
                ElseIf tmGrf.lDollars(8) = 0 And ilOverUnder = 2 Then                   'matches delivery, include all?
                    If (tgChfAD.lGuar = 0 And ilPledge > 0) Or (tgChfAD.lGuar <> 0 And ilPledge <> 1) Then        'no guarantee entered and user asing for pledged only (ignore)
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                    End If
                End If
            End If
        End If              'endif llStartDates(2) > llStartDates(1)
    Next illoop

    'debugging only for time program took to run
    slStr = Format$(gNow(), "h:mm:ssAM/PM")       'end time of run
    gPackTime slStr, ilNowTime(0), ilNowTime(1)
    gUnpackTimeLong ilNowTime(0), ilNowTime(1), False, llTime
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, llPop   'start time of run
    llPop = llPop - llTime              'time in seconds in runtime
    ilRet = gSetFormula("RunTime", llPop)  'show how long report generated

    sgCntrForDateStamp = ""     'initialize contract routine next time thru
    Erase tmActiveCnts, tmVehicleBook, tmDnfLinkList, tmBookList, tmPLSdf
    Erase tgCffAD, tgClfAD
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmDrf)
    ilRet = btrClose(hmDnf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmMnf)
    btrDestroy hmRaf
    btrDestroy hmDef
    btrDestroy hmDpf
    btrDestroy hmMnf
    btrDestroy hmVsf
    btrDestroy hmSmf
    btrDestroy hmSdf
    btrDestroy hmVef
    btrDestroy hmDrf
    btrDestroy hmDnf
    btrDestroy hmCHF
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    Exit Sub
mTerminate:
    On Error GoTo 0
    Exit Sub
End Sub
'
'
'           mBinarySearch - find the spots contract code in the list of active
'            contracts to process
'
'           <input> llChfcode - contract code to match against list
'           <output> ilmin - starting point of list
'                    ilmax - ending point of list
'                    mBinarySearch - index of matching entry
Function mBinarySearch(llChfCode As Long, ilMin As Integer, ilMax As Integer) As Integer
Dim ilMiddle As Integer
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If llChfCode = tmActiveCnts(ilMiddle).lChfCode Then
            'found the match
            mBinarySearch = ilMiddle
            Exit Function
        ElseIf llChfCode < tmActiveCnts(ilMiddle).lChfCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearch = -1
End Function
'
'
'       mBuildTAbles - build tables required to process the Delivery information
'
'       Created 5/7/00
'
'       tmActiveCnts() -Active contracts are gathered based on the users start/end dates.  All
'       contracts active during those dates are processed for its contracts start date
'       to its contract end date (all spots). Array is sorted by contract code.  A binary
'       search is used to find the contract associated with each spot.
'
'       tmVehiclebook() - all valid (active conventional and selling vehicles which hold
'       the vehicle code and first and last book name index pointers.  These are the
'       books associated with each vehicle, which point to tmBOOKLIST.
'
'       tmBookList() - array of books containing book start date and book code.  This array
'       is sorted by book start date to speed up search.  Each spot needs to find the
'       book closest to airing.
'
'       tmDnfLinkList() - array of indices that point to the tmBookList array associated with a vehicle.
'       tmVehicleBook points to this array.
'
'       <input>  slstart - user entered active start date
'                slend - user enetered active end date
'       <output> llEarliestStart - earliest start date of all contracts gathered (for spot gthering)
'                llLatestEnd - latest end date of all contracts gathered (for Spot gathering)
'
'       <return> mBuildTables - 0 = OK, 1 = error
'
Function mBuildTables(slStart As String, slEnd As String, llEarliestStart As Long, llLatestEnd As Long, ilListIndex As Integer) As Integer
Dim ilUpper As Integer
Dim illoop As Integer
Dim ilVehicle As Integer
Dim ilRet As Integer
Dim slCntrTypes As String
Dim slCntrStatus As String
Dim ilHOState As Integer
Dim slStr As String
Dim slCode As String
Dim llfirstTime As Long         '1-15-08 chg to long
Dim llChfStartDate As Long
Dim llChfEndDate As Long
Dim llEnteredStartDate As Long
Dim llEnteredEndDate As Long
Dim tlChfAdvtExt() As CHFADVTEXT
Dim llUpper As Long                     '1-15-08
Dim ilDaysInPast As Integer

    ilListIndex = RptSelAD!lbcRptType.ListIndex
         '0 = aud delivery, 1 = post buy
    ilDaysInPast = 730
    If ilListIndex = DELIVERY_POSTBUY Then
        ilDaysInPast = 100                  'go back 1 quarter in the past for safety
    End If

    mBuildTables = 0
    llEnteredStartDate = gDateValue(slStart)
    llEnteredEndDate = gDateValue(slEnd)
    ReDim tmActiveCnts(0 To 0) As ACTIVECNTS
    'Gather all contracts for previous year and current year whose effective date entered
    'is prior to the effective date that affects either previous year or current year
    slCntrTypes = gBuildCntTypes()
    slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)

    'if selective advertiser, get the contracts for that advt
    If Not (RptSelAD!ckcAll.Value = vbChecked) Then
        'Loop on the contract list for the selective contracts
        For illoop = 0 To RptSelAD!lbcSelection(0).ListCount - 1 Step 1
            If RptSelAD!lbcSelection(0).Selected(illoop) Then
                slStr = RptSelAD!lbcCntrCode.List(illoop)
                ilRet = gParseItem(slStr, 2, "\", slCode)
                tmChfSrchKey.lCode = Val(slCode)
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet <> 0 Then
                    mBuildTables = 1
                    Exit Function
                End If
                'If user date entered, the contract must span it
                gUnpackDateLong tmChf.iStartDate(0), tmChf.iStartDate(1), llChfStartDate
                gUnpackDateLong tmChf.iEndDate(0), tmChf.iEndDate(1), llChfEndDate
                If llChfEndDate >= llEnteredStartDate And llChfStartDate <= llEnteredEndDate Then
                    ilUpper = UBound(tmActiveCnts)
                    slCode = Trim$(str$(tmChf.lCode))
                    Do While Len(slCode) < 8
                        slCode = "0" & slCode
                    Loop
                    tmActiveCnts(ilUpper).sKey = slCode
                    tmActiveCnts(ilUpper).lChfCode = tmChf.lCode
                    tmActiveCnts(ilUpper).sType = tmChf.sType
                    tmActiveCnts(ilUpper).iMnfDemo = tmChf.iMnfDemo(0)
                    tmActiveCnts(ilUpper).lPop = -1
                    tmActiveCnts(ilUpper).iFirstPopLink = -1
                    tmActiveCnts(ilUpper).iVaryPop = 0          'flag to indicate varying pop within same line; across weeks or days
                    'gUnPackDateLong tmChf.iStartDate(0), tmChf.iStartDate(1), tmActiveCnts(ilUpper).lStartDate
                    'gUnPackDateLong tmChf.iEndDate(0), tmChf.iEndDate(1), tmActiveCnts(ilUpper).lEndDate
                    tmActiveCnts(ilUpper).lStartDate = llChfStartDate
                    tmActiveCnts(ilUpper).lEndDate = llChfEndDate
                    'Get the earliest and latest contract dates from all contracts to process
                    If llEarliestStart = 0 Then
                        llEarliestStart = tmActiveCnts(ilUpper).lStartDate
                    Else
                        If tmActiveCnts(ilUpper).lStartDate < llEarliestStart Then
                            llEarliestStart = tmActiveCnts(ilUpper).lStartDate
                        End If
                    End If
                    If tmActiveCnts(ilUpper).lEndDate > llLatestEnd Then
                        llLatestEnd = tmActiveCnts(ilUpper).lEndDate
                    End If
                    ReDim Preserve tmActiveCnts(0 To ilUpper + 1) As ACTIVECNTS
                End If
            End If
        Next illoop
    Else        'all advt, see if selective contract, retrieve active ones based on the dates entered
        lmSingleCntr = Val(RptSelAD!edcContract.Text)       'selective contract #
        If lmSingleCntr > 0 Then                'get contract for code
            tmChfSrchKey1.lCntrNo = lmSingleCntr
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            If ilRet = BTRV_ERR_NONE And tmChf.lCntrNo <> lmSingleCntr Then
                'contract does not exist
                MsgBox "Contract does not exist"
                mBuildTables = 1
                Exit Function
            Else
                'ReDim tlChfAdvtExt(1 To 2) As CHFADVTEXT
                ReDim tlChfAdvtExt(0 To 1) As CHFADVTEXT
                lmSingleCntr = tmChf.lCode
                'tlChfAdvtExt(1).lCode = tmChf.lCode
                tlChfAdvtExt(0).lCode = tmChf.lCode
            End If
        Else
            ilRet = gObtainCntrForDate(RptSelAD, slStart, slEnd, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
            If ilRet <> 0 Then
                mBuildTables = 1
                Exit Function
            End If
        End If
        
        For illoop = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1

            tmChfSrchKey.lCode = tlChfAdvtExt(illoop).lCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            ilUpper = UBound(tmActiveCnts)
            slCode = Trim$(str$(tmChf.lCode))
            Do While Len(slCode) < 8
                slCode = "0" & slCode
            Loop
            tmActiveCnts(ilUpper).sKey = slCode         'contract code left filled with zeroes for sorting
            tmActiveCnts(ilUpper).lChfCode = tmChf.lCode
            tmActiveCnts(ilUpper).sType = tmChf.sType
            tmActiveCnts(ilUpper).iMnfDemo = tmChf.iMnfDemo(0)
            tmActiveCnts(ilUpper).lPop = -1
            tmActiveCnts(ilUpper).iFirstPopLink = -1
            tmActiveCnts(ilUpper).iVaryPop = 0          'flag to indicate varying pop within same line; across weeks or days
            gUnpackDateLong tmChf.iStartDate(0), tmChf.iStartDate(1), tmActiveCnts(ilUpper).lStartDate
            gUnpackDateLong tmChf.iEndDate(0), tmChf.iEndDate(1), tmActiveCnts(ilUpper).lEndDate
            'Get the earliest and latest contract dates from all contracts to process
            If llEarliestStart = 0 Then
                llEarliestStart = tmActiveCnts(ilUpper).lStartDate
            Else
                If tmActiveCnts(ilUpper).lStartDate < llEarliestStart And llEarliestStart <> 0 Then
                    llEarliestStart = tmActiveCnts(ilUpper).lStartDate
                End If
            End If
            If tmActiveCnts(ilUpper).lEndDate > llLatestEnd Then
                llLatestEnd = tmActiveCnts(ilUpper).lEndDate
            End If
            ReDim Preserve tmActiveCnts(0 To ilUpper + 1) As ACTIVECNTS
        Next illoop
    End If
    If ilUpper > 0 Then
         ArraySortTyp fnAV(tmActiveCnts(), 0), ilUpper + 1, 0, LenB(tmActiveCnts(0)), 0, LenB(tmActiveCnts(0).sKey), 0
    End If

    Erase tlChfAdvtExt
    ReDim tmVehicleBook(0 To 0) As VEHICLEBOOK
    'ilRet = gObtainVef()         'vehicles have already been gathered in global array tgMVef
    For illoop = LBound(tgMVef) To UBound(tgMVef) - 1
        'look for Active (not dormant) and vehicle type Conventional or Selling vehicle, or game (added 5-06-08)
        If (tgMVef(illoop).sType = "C" Or tgMVef(illoop).sType = "S" Or tgMVef(illoop).sType = "G") And (tgMVef(illoop).sState = "A") Then
        
            ilUpper = UBound(tmVehicleBook)         '
            tmVehicleBook(ilUpper).iVefCode = tgMVef(illoop).iCode
            tmVehicleBook(ilUpper).lDnfFirstLink = -1
            tmVehicleBook(ilUpper).lDnfLastLink = 0
            ReDim Preserve tmVehicleBook(0 To ilUpper + 1) As VEHICLEBOOK
        End If
    Next illoop
    ReDim tmBookList(0 To 0) As BOOKLIST        'list of books

    For illoop = 0 To RptSelAD!cbcBook.ListCount - 1 Step 1
        ilUpper = UBound(tmBookList)

        slStr = tgBookNameCode(illoop).sKey
        ilRet = gParseItem(slStr, 2, "\", slCode)
        tmBookList(ilUpper).iDnfCode = Val(slCode)  'book code
        'get the book to store the start date
        tmDnfSrchKey.iCode = Val(slCode)
        ilRet = btrGetGreaterOrEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point
        If ilRet <> BTRV_ERR_NONE Then
            mBuildTables = 1
            Exit Function
        End If
        gUnpackDateLong tmDnf.iBookDate(0), tmDnf.iBookDate(1), tmBookList(ilUpper).lStartDate

        'If tmBookList(ilUpper).lStartDate >= (llChfStartDate - ilDaysInPast) Then            '1-15-08 only gather the books whose book date is equal/greater than the
        '8-22-14 wrong date used in gathering data, reduce amount of array entries
        If tmBookList(ilUpper).lStartDate >= (llEarliestStart - ilDaysInPast) Then            '1-15-08 only gather the books whose book date is equal/greater than the
                                                                            'earliest contract start date minus 2 years (365 * 2)
            slStr = Trim$(str$(tmBookList(ilUpper).lStartDate))
            Do While Len(slStr) < 5        'left fill zeroes for date sort
                slStr = "0" & slStr
            Loop
            tmBookList(ilUpper).sKey = slStr
            ReDim Preserve tmBookList(0 To ilUpper + 1) As BOOKLIST
        End If
    Next illoop
    If ilUpper > 0 Then    'sort by book date
         ArraySortTyp fnAV(tmBookList(), 0), ilUpper + 1, 0, LenB(tmBookList(0)), 0, LenB(tmBookList(0).sKey), 0
    End If
    'Build  list of associated books with vehicles and demo research link list
    ReDim tmDnfLinkList(0 To 0) As DNFLINKLIST
    For ilVehicle = 0 To UBound(tmVehicleBook) - 1
        llfirstTime = -1            '1-15-08 chg to long
        For illoop = 0 To UBound(tmBookList) - 1
            tmDrfSrchKey1.iDnfCode = tmBookList(illoop).iDnfCode
            tmDrfSrchKey1.sDemoDataType = "D"
            tmDrfSrchKey1.iMnfSocEco = 0
            tmDrfSrchKey1.iVefCode = tmVehicleBook(ilVehicle).iVefCode
            tmDrfSrchKey1.iStartTime(0) = 0
            tmDrfSrchKey1.iStartTime(1) = 0
            tmDrfSrchKey1.sInfoType = "D"
            ilRet = btrGetGreaterOrEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point
            If ilRet <> BTRV_ERR_NONE Then
                mBuildTables = 1
                Exit Function
            End If
            If tmDrf.iVefCode = tmVehicleBook(ilVehicle).iVefCode And tmDrf.iDnfCode = tmBookList(illoop).iDnfCode Then
                llUpper = UBound(tmDnfLinkList)     '1-15-08 chg to long
                If llfirstTime < 0 Then
                    tmVehicleBook(ilVehicle).lDnfFirstLink = llUpper
                End If
                llfirstTime = llUpper
               ' tmDnfLinkList(ilUpper).idnfInx = tmDrf.iDnfCode
               'The LinkList points to the associated tmBookList entry of this vehicle
                tmDnfLinkList(llUpper).idnfInx = illoop
                ReDim Preserve tmDnfLinkList(0 To llUpper + 1) As DNFLINKLIST
            End If
        Next illoop
        'No more books for the current vehicle, set its last index
        tmVehicleBook(ilVehicle).lDnfLastLink = llUpper
    Next ilVehicle

    Exit Function

    On Error GoTo 0
    mBuildTables = 1
    Exit Function
End Function
'****************************************************************************
'*
'*      Procedure Name:mObtainSdf
'       <input>  ilVefCode = vehicle code to search
'                slStartDate as string - earliest date to retrieve
'                slEndDate as string - latest date to retrieve
'                ilWhichKey as integer - KEYINDEX1 or KEYINDEX2
'                ilIncludeCodes - true = include codes in ilCludecodes() array, false = exclude codes in array
'                ilUseCodes() -array of advt codes to include or exclude
'*
'*             Created:10/09/93      By:D. LeVine
'*            Modified: 11/20/96     By:d.h.
'*
'*            Comments:Obtain the Sdf records to be
'*                     reported
'       5/7/00 Extracted from Spots by Date & Time;
'               and modified
'
'           3/12/98  Add time filter
'           4/6/99 For packages, test to bill as ordered
'                  when billing as aired
'           6/18/99 Included missed when requested for
'              option as aired/pkg ordered for NONE.
'              previously, package missed not included.
'           7/2/99 more problems as described on 6/18/99.
'               NONe didnt work properly, advt option
'               OK
'           6-19-03 spot key was not created properly, sort array by contract
'           3-29-05 exclude bb spots
'*****************************************************************************
Sub mObtainSdf(ilVefCode As Integer, slStartDate As String, slEndDate As String, ilWhichKey As Integer, ilIncludeCodes As Integer, ilUseCodes() As Integer)
'
'
    Dim slDate As String
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim ilOffSet As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slStr As String
    Dim ilUpper As Integer
    Dim ilOk As Integer
    Dim ilFound As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim llChfCode As Long
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record

    tmVefSrchKey.iCode = ilVefCode
    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_NONE Then
        Exit Sub
    End If
    ilUpper = UBound(tmPLSdf)
    btrExtClear hmSdf   'Clear any previous extend operation
    ilExtLen = Len(tmSdf)
    llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlSdf) 'Obtain number of records
    btrExtClear hmSdf   'Clear any previous extend operation
    If ilWhichKey = 1 Then      'by vehicle, then date
        tmSdfSrchKey1.iVefCode = ilVefCode
        gPackDate slStartDate, tmSdfSrchKey1.iDate(0), tmSdfSrchKey1.iDate(1)
        tmSdfSrchKey1.iTime(0) = 0
        tmSdfSrchKey1.iTime(1) = 0
        tmSdfSrchKey1.sSchStatus = ""
        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Else
        tmSdfSrchKey2.iAdfCode = 0
        gPackDate slStartDate, tmSdfSrchKey2.iDate(0), tmSdfSrchKey2.iDate(1)
        tmSdfSrchKey2.sSchStatus = ""
        tmSdfSrchKey2.iVefCode = ilVefCode
        tmSdfSrchKey2.iTime(0) = 0
        tmSdfSrchKey2.iTime(1) = 0
        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    End If

    If ilRet <> BTRV_ERR_END_OF_FILE Then
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") 'Set extract limits (all records)
        tlIntTypeBuff.iType = ilVefCode
        ilOffSet = gFieldOffset("Sdf", "SdfVefCode")
        If (slStartDate <> "") Or (slEndDate <> "") Then
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        Else
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        End If
        If slStartDate <> "" Then
            gPackDate slStartDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Sdf", "SdfDate")
            If slEndDate <> "" Then
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            Else
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
            End If
        End If
        If slEndDate <> "" Then
            gPackDate slEndDate, tlDateTypeBuff.iDate0, tlDateTypeBuff.iDate1
            ilOffSet = gFieldOffset("Sdf", "SdfDate")
            ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
        End If
        ilRet = btrExtAddField(hmSdf, 0, ilExtLen)  'Extract Name
        If ilRet <> BTRV_ERR_NONE Then
            Exit Sub
        End If
        ilRet = btrExtGetNext(hmSdf, tmPLSdf(ilUpper).tSdf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
                Exit Sub
            End If
            'ilRet = btrExtGetFirst(hlSdf, tlSdfExt(ilUpper), ilExtLen, llRecPos)
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmSdf, tmPLSdf(ilUpper).tSdf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                'determine if this spot should be selected based on the active list of contracts for the week entered by user
                'search tmActiveCnts array which is contract code order
                ilOk = True
                llChfCode = tmPLSdf(ilUpper).tSdf.lChfCode

                ilMin = LBound(tmActiveCnts)
                ilMax = UBound(tmActiveCnts)
                ilFound = mBinarySearch(llChfCode, ilMin, ilMax)    'find the matching cntr in the list to process
                '3-20-05 exclude spots whose contract is not found or whose spot is an open or closed bb
                If ilFound = -1 Or tmPLSdf(ilUpper).tSdf.sSpotType = "O" Or tmPLSdf(ilUpper).tSdf.sSpotType = "C" Then      'not found
                    ilOk = False
                End If

                If ilWhichKey = 2 And ilOk Then          'search for advertisers if trying keyindex2
                    If Not gFilterLists(tmPLSdf(ilUpper).tSdf.iAdfCode, ilIncludeCodes, ilUseCodes()) Then
                        ilOk = False
                    End If
                End If

                If ilOk Then
                    slStr = Trim$(str$(tmPLSdf(ilUpper).tSdf.lChfCode))
                    Do While Len(slStr) < 8
                        slStr = "0" & slStr
                    Loop
                    tmPLSdf(ilUpper).sKey = Trim$(slStr) & "|"
                    slStr = Trim$(str$(tmPLSdf(ilUpper).tSdf.iLineNo))
                    Do While Len(slStr) < 4
                        slStr = "0" & slStr
                    Loop
                    tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & Trim$(slStr) & "|"
                    gUnpackDateForSort tmPLSdf(ilUpper).tSdf.iDate(0), tmPLSdf(ilUpper).tSdf.iDate(1), slDate
                    tmPLSdf(ilUpper).sKey = Trim$(tmPLSdf(ilUpper).sKey) & Trim$(slDate)
                    ReDim Preserve tmPLSdf(0 To ilUpper + 1) As SPOTTYPESORTAD
                    ilUpper = ilUpper + 1

                End If
                ilRet = btrExtGetNext(hmSdf, tmPLSdf(ilUpper).tSdf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmSdf, tmPLSdf(ilUpper).tSdf, ilExtLen, llRecPos)
                Loop
            Loop
            If ilUpper > 0 Then    'sort by book date
                ArraySortTyp fnAV(tmPLSdf(), 0), ilUpper, 0, LenB(tmPLSdf(0)), 0, LenB(tmPLSdf(0).sKey), 0          '3-17-12 chged from ilupper+1 to ilupper
            End If
        End If
    End If
    Exit Sub

    ilRet = err.Number
    Resume Next
End Sub
'
'       gCrPostBuy - create prepass for Post Buy Analysis report
Public Sub gCrPostBuy()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slEarliestStart               slLatestEnd                   ilBook                    *
'*  llCntStartDates               llTotalCost                   llSpots                   *
'*  llGuarPct                     ilDay                                                   *
'******************************************************************************************

Dim ilError As Integer
Dim ilRet As Integer
Dim slStr As String
Dim llStart As Long         'user entered start date
Dim llEnd As Long           'user entered end date
Dim slStart As String   'user entered start date
Dim slEnd As String     'user entered end date
Dim llEarliestStart As Long
Dim llLatestEnd As Long
ReDim imUseCodes(0 To 0) As Integer     'array of advt codes to include/exclude
Dim ilListIndex As Integer
Dim ilVehicle As Integer
Dim ilVefCode As Integer
Dim ilSpotLoop As Integer
Dim llContrCode As Long
Dim ilMin As Integer
Dim ilMax As Integer
Dim ilOk As Integer
Dim ilActiveCntInx As Integer
Dim llOvEndTime As Long
Dim llOvStartTime As Long
Dim illoop As Integer
Dim llDate As Long
Dim ilDnfCode As Integer
Dim ilInputDays(0 To 6) As Integer
Dim llMinLink As Long
Dim llMaxLink As Long
Dim llLoop As Long
Dim ilBookInx As Integer
Dim llDnfDate As Long
Dim ilFound As Integer
Dim llPrice As Long
Dim slPrice As String
Dim llPop As Long
Dim ilUpperPop As Integer
Dim ilLink As Integer
Dim ilKeepLast As Integer
'****** following required for gAvgAudToLnResearch, calculated for every spot
ReDim llWklyspots(0 To 0) As Long    '1 spot for 1 week for aud routine
ReDim llWklyRates(0 To 0) As Long       'spot price per week
ReDim llWklyAvgAud(0 To 0) As Long      'avg aud perweek
ReDim llWklyPopEst(0 To 0) As Long
'Dim llContrGross As Long         'total cost (or spot cost)
Dim dlContrGross As Double       'total cost (or spot cost)'TTP 10439 - Rerate 21,000,000
Dim llTotalAvgAud As Long      'avg aud per week
ReDim ilWklyRtg(0 To 0) As Integer        'wkly weekly rating
Dim ilAVgRtg As Integer        'avg rating
ReDim llWklyGrimp(0 To 0) As Long 'weekly gross impressions
Dim llTotalGrImp As Long        'total grimps
ReDim llWklyGRP(0 To 0) As Long   'weekly gross rating points
Dim llTotalGRP As Long          'total grps
Dim llTotalCPP As Long          'total CPPS
Dim llTotalCPM As Long          'Total CPMS
Dim llAvgAud As Long
Dim llPopEst As Long
Dim ilSpotRateOK As Integer
Dim tlCntTypes As CNTTYPES
'Dim llSingleCntr As Long
Dim ilAudFromSource As Integer
Dim llAudFromCode As Long

    ilListIndex = RptSelAD!lbcRptType.ListIndex      'selected report
    ilError = mOpenDelivery()           'open applicable files
    If ilError Then
        Exit Sub            'at least 1 open error
    End If
'    slStr = RptSelAD!edcSelCFrom.Text               'Active Start Date
    slStr = RptSelAD!CSI_CalFrom.Text               'Active Start Date  9-3-19 use csi cal control vs edit box
    llStart = gDateValue(slStr)                     'gather contracts thru this date
'    slStr = RptSelAD!edcSelCTo.Text               'Active Start Date
    slStr = RptSelAD!CSI_CalTo.Text               'Active Start Date
    llEnd = gDateValue(slStr)                     'gather contracts thru this date
    slStart = Format$(llStart, "m/d/yy")   'insure the year is in the format, may not have been entered with the date input
    slEnd = Format$(llEnd, "m/d/yy")

    ilRet = mBuildTables(slStart, slEnd, llEarliestStart, llLatestEnd, ilListIndex) 'build table of active contracts, vehicles, books and the list of associated books with each vehicle
    If ilRet <> 0 Then
        MsgBox "gCrPostBuy: mBuildTables- Error in building post buy tables"
        Erase tmActiveCnts, tmVehicleBook, tmDnfLinkList, tmPLSdf
        Erase tgCffAD, tgClfAD
        mCloseFiles
        Exit Sub
    End If

    lgDpfNoRecs = btrRecords(hmDpf)     'determine if demo plus info exist
    If lgDpfNoRecs = 0 Then
        lgDpfNoRecs = -1
    End If

    tlCntTypes.iCharge = gSetCheck(RptSelAD!ckcSelC5(0).Value)
    tlCntTypes.iZero = gSetCheck(RptSelAD!ckcSelC5(1).Value)
    tlCntTypes.iADU = gSetCheck(RptSelAD!ckcSelC5(2).Value)
    tlCntTypes.iBonus = gSetCheck(RptSelAD!ckcSelC5(3).Value)
    tlCntTypes.iXtra = gSetCheck(RptSelAD!ckcSelC5(4).Value)
    tlCntTypes.iFill = gSetCheck(RptSelAD!ckcSelC5(5).Value)
    tlCntTypes.iNC = gSetCheck(RptSelAD!ckcSelC5(6).Value)
    tlCntTypes.iMG = gSetCheck(RptSelAD!ckcSelC5(7).Value)
    tlCntTypes.iRecapturable = gSetCheck(RptSelAD!ckcSelC5(8).Value)
    tlCntTypes.iSpinoff = gSetCheck(RptSelAD!ckcSelC5(9).Value)
    tlCntTypes.iMissed = gSetCheck(RptSelAD!ckcSelC5(10).Value)

'8-22-14 move to mBuildTAbles routine
'    llSingleCntr = Val(RptSelAD!edcContract.Text)       'selective contract #
'    If llSingleCntr > 0 Then                'get contract for code
'        tmChfSrchKey1.lCntrNo = llSingleCntr
'        tmChfSrchKey1.iCntRevNo = 32000
'        tmChfSrchKey1.iPropVer = 32000
'        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
'        If ilRet = BTRV_ERR_NONE And tmChf.lCntrNo <> llSingleCntr Then
'            'contract does not exist
'            MsgBox "Contract does not exist"
'            Erase tmActiveCnts, tmVehicleBook, tmDnfLinkList, tmPLSdf
'            Erase tgCffAD, tgClfAD
'            mCloseFiles
'            Exit Sub
'        Else
'            llSingleCntr = tmChf.lCode
'        End If
'    End If

    'Link list of each schedule line/contract with their line # and populations
    ReDim tmpoplinklist(0 To 0) As POPLINKLIST
    If lmSingleCntr > 0 Then            'for single contract, set the adv selection, that contract is still in memory
        imIncludeCodes = -1
        'ReDim imUseCodes(1 To 2) As Integer
        'imUseCodes(1) = tmChf.iAdfCode
        ReDim imUseCodes(0 To 1) As Integer
        imUseCodes(0) = tmChf.iAdfCode
    Else
        gObtainCodesForMultipleLists 1, tgAdvertiser(), imIncludeCodes, imUseCodes(), RptSelAD
    End If
    
    tmClf.lChfCode = 0          'initialize for first time thru and re-entrant
    For ilVehicle = 0 To UBound(tmVehicleBook) - 1
        ilVefCode = tmVehicleBook(ilVehicle).iVefCode
        ReDim tmPLSdf(0 To 0) As SPOTTYPESORTAD
        mObtainSdf ilVefCode, slStart, slEnd, INDEXKEY2, imIncludeCodes, imUseCodes()
        For ilSpotLoop = 0 To UBound(tmPLSdf) - 1
            'filter out unwanted contract types .  Spot List (tmPlSdf is in Contract code, line & date order)
            tmSdf = tmPLSdf(ilSpotLoop).tSdf
            llContrCode = tmSdf.lChfCode
            ilMin = LBound(tmActiveCnts)
            ilMax = UBound(tmActiveCnts)
            'if same contract as previous, no need to find it, ilOK contains the acceptance value
            'If (tmClf.lChfCode <> tmSdf.lChfCode) Then
                ilOk = True

                ilActiveCntInx = mBinarySearch(llContrCode, ilMin, ilMax)    'find the matching cntr in the list to process

                If ilActiveCntInx = -1 Then     'not found
                    ilOk = False
                End If
            'End If
            If ilOk Then        'found a matching contr in the active list
                'always ignore: psa, promo, reserveration, PI, DR, and Remnants
                If (tmActiveCnts(ilActiveCntInx).sType = "S") Or (tmActiveCnts(ilActiveCntInx).sType = "M") Or (tmActiveCnts(ilActiveCntInx).sType = "V") Or (tmActiveCnts(ilActiveCntInx).sType = "T") Or (tmActiveCnts(ilActiveCntInx).sType = "R") Or (tmActiveCnts(ilActiveCntInx).sType = "Q") Or (lmSingleCntr <> 0 And lmSingleCntr <> tmSdf.lChfCode) Then
                    ilOk = False
                End If
            End If
            If ilOk Then            'only test to reread line of contract is OK to use (not an excluded type such as PI, DR, etc)
                ilRet = 0
                If (tmClf.lChfCode <> tmSdf.lChfCode) Or (tmClf.lChfCode <> tmSdf.lChfCode Or tmClf.iLine <> tmSdf.iLineNo) Then

                    'get the correctline for this spot
                    tmClfSrchKey.lChfCode = tmSdf.lChfCode
                    tmClfSrchKey.iLine = tmSdf.iLineNo
                    tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                    tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
                If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
                    'retain slPrice to accumulate the grps & audience
                    ilRet = gGetFlightPrice(tmSdf, tmClf, hmCff, hmSmf, slPrice)  'get spot price here so that the flight can be accessed in cff
If tmSdf.lCode = 4180721 Or tmSdf.lCode = 4180722 Then
ilRet = ilRet
End If

                    ilSpotRateOK = gFilterSpotRateType(tlCntTypes, slPrice)   'filter out types of line spot rate types (charge, nc, adu, fills, etc)
                    If Not ilSpotRateOK Then            'ignore spot
                        ilOk = False
                    End If
                    'spot rate types tested; MG inclusion/exclusion make up 2 different types of mg flags:
                    '1:  mg spot rate on sched line, the other is the spot sched status = "G"
                    If Not tlCntTypes.iMG And tmSdf.sSchStatus = "G" Then       'exclude mg if user deselected
                        ilOk = False
                    End If

                    If Not tlCntTypes.iMissed And tmSdf.sSchStatus = "M" Then       'exclude missed spts
                        ilOk = False
                    End If
                Else
                    ilOk = False
                End If
            End If
            If ilOk Then
                'Some of the following code may not be necessary; but has been extracted from Audience Delivery report
                'This report deals with 1 spot at a time, Audience Delivery totals the spots for the contract
                gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), True, llOvStartTime
                llOvEndTime = llOvStartTime           'use the avail time for overrides to determine audience
                For illoop = 0 To 6                 'init all days to not airing, setup for research results later
                    ilInputDays(illoop) = False
                Next illoop
                'set day of week aired
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                slStr = Format(llDate, "m/d/yy")
                illoop = gWeekDayStr(slStr)     'day index

                ilInputDays(illoop) = True     'set day of week in week pattern
                'Determine whether to use the default book or the book closest to airing
                'Set Default book of vehicle in case not found
                ilDnfCode = 0
                'For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1
                '    If ilVefCode = tgMVef(ilLoop).iCode Then
                    illoop = gBinarySearchVef(ilVefCode)
                    If illoop <> -1 Then
                        ilDnfCode = tgMVef(illoop).iDnfCode
                '        Exit For
                    End If
                'Next ilLoop
                'default book has been set in case there isnt a book closest to airing found
                'use closest to airing (vs default book)
                ilFound = False
                '1-15-08 change the references to dnflink to long variables
                llMinLink = tmVehicleBook(ilVehicle).lDnfFirstLink
                llMaxLink = tmVehicleBook(ilVehicle).lDnfLastLink
                'The LinkList points to the associated tmBookList entry of this vehicle
                If llMinLink <> -1 Then          '-1 indicates no books found
                    For llLoop = llMaxLink To llMinLink Step -1
                    'For ilLoop = ilMin To ilMax
                        ilBookInx = tmDnfLinkList(llLoop).idnfInx
                        llDnfDate = tmBookList(ilBookInx).lStartDate
                        If llDate >= llDnfDate And llDnfDate <> 0 Then
                            ilDnfCode = tmBookList(ilBookInx).iDnfCode
                            ilFound = True
                            Exit For
                        End If
                    Next llLoop
                End If
                If Not ilFound Then
                    tmActiveCnts(ilActiveCntInx).iBookMissing = 1   'Flag error, 1 vehiclemissing an associated book
                End If

                llPrice = 0     'init incase a decimal number isnt in price field (its adu, nc, fill,etc.)
                If (InStr(slPrice, ".") <> 0) Then        'found spot cost
                    llPrice = gStrDecToLong(slPrice, 2)
                End If
                ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, 0, tmActiveCnts(ilActiveCntInx).iMnfDemo, llPop)

                '6-6-04 Get the pops only if not using Demo Estimates
                If tgSpf.sDemoEstAllowed <> "Y" Then        'not using demo estimates
                    'Set if varying populations within/across schedule lines in contract
                    If tmActiveCnts(ilActiveCntInx).lPop < 0 Then       'never been set up, save population first time
                        If llPop <> 0 Then
                            tmActiveCnts(ilActiveCntInx).lPop = llPop
                        End If
                    ElseIf tmActiveCnts(ilActiveCntInx).lPop <> 0 Then
                        If tmActiveCnts(ilActiveCntInx).lPop <> llPop And llPop <> 0 Then
                            tmActiveCnts(ilActiveCntInx).lPop = 0
                        End If
                    End If

                    'Keep track if varying populations within same schedule line (across days or weeks)
                    'If varying within sch line, cant product grp numbers, in which case it should be zerod
                    ilUpperPop = UBound(tmpoplinklist)
                    If tmActiveCnts(ilActiveCntInx).iFirstPopLink = -1 Then      'first time thru for this contract
                        tmActiveCnts(ilActiveCntInx).iFirstPopLink = ilUpperPop
                        If llPop <> 0 Then
                            tmpoplinklist(ilUpperPop).lPop = llPop
                        End If
                        tmpoplinklist(ilUpperPop).iLine = tmClf.iLine
                        tmpoplinklist(ilUpperPop).iNextLink = -1
                        ReDim Preserve tmpoplinklist(0 To ilUpperPop + 1) As POPLINKLIST
                    Else            'already found at least one spot for this contract , see if varying pops
                        ilUpperPop = UBound(tmpoplinklist)
                        ilFound = False
                        ilLink = tmActiveCnts(ilActiveCntInx).iFirstPopLink
                        Do While ilLink <> -1
                            If tmpoplinklist(ilLink).iLine = tmClf.iLine Then
                                ilFound = True
                                If tmpoplinklist(ilLink).lPop < 0 Then       'never been set up, save population first time
                                    If llPop <> 0 Then
                                        tmpoplinklist(ilLink).lPop = llPop
                                    End If
                                ElseIf tmpoplinklist(ilLink).lPop <> 0 Then
                                    If tmpoplinklist(ilLink).lPop <> llPop And llPop <> 0 Then
                                        tmpoplinklist(ilLink).lPop = 0
                                        tmActiveCnts(ilActiveCntInx).iVaryPop = 1   'cant compute cpps from grps
                                    End If
                                End If
                                ilLink = -1     'force to stop search
                            Else                'not matching line #
                                ilKeepLast = ilLink
                                ilLink = tmpoplinklist(ilLink).iNextLink
                            End If
                        Loop
                        If Not ilFound Then             'set up new entry
                            tmpoplinklist(ilKeepLast).iNextLink = ilUpperPop
                            If llPop <> 0 Then
                                tmpoplinklist(ilUpperPop).lPop = llPop
                            End If
                            tmpoplinklist(ilUpperPop).iLine = tmClf.iLine
                            tmpoplinklist(ilUpperPop).iNextLink = -1
                            ReDim Preserve tmpoplinklist(0 To ilUpperPop + 1) As POPLINKLIST
                        End If
                    End If
                End If

                '************ DEBUGGING ONLY   *******************
                'llOVStarttime = 0         use daypart times only, no overrides
                'llOVEndTime = 0
                ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, tmClf.iVefCode, 0, tmActiveCnts(ilActiveCntInx).iMnfDemo, llDate, llDate, tmClf.iRdfCode, llOvStartTime, llOvEndTime, ilInputDays(), tmClf.sType, tmClf.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)

                'need to get grp from gAvgAudToLnResearch
                'If llPop < 0 Then
               '     llPop = 0
               ' End If
                llWklyspots(0) = 1
                llWklyAvgAud(0) = llAvgAud
                llWklyRates(0) = llPrice
                llWklyPopEst(0) = llPopEst


                'Use population of the book for the individual spot
                '10-30-14 default to use 1 place rating regardless of agency flag
                'gAvgAudToLnResearch "1", True, llPop, llWklyPopEst(), llWklyspots(), llWklyRates(), llWklyAvgAud(), llContrGross, llTotalAvgAud, ilWklyRtg(), ilAVgRtg, llWklyGrimp(), llTotalGrImp, llWklyGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst
                gAvgAudToLnResearch "1", True, llPop, llWklyPopEst(), llWklyspots(), llWklyRates(), llWklyAvgAud(), dlContrGross, llTotalAvgAud, ilWklyRtg(), ilAVgRtg, llWklyGrimp(), llTotalGrImp, llWklyGRP(), llTotalGRP, llTotalCPP, llTotalCPM, llPopEst  'TTP 10439 - Rerate 21,000,000
                '6-6-04 Get the pops only if not using Demo Estimates
                If tgSpf.sDemoEstAllowed = "Y" Then        'using demo estimates
                    'Set if varying populations within/across schedule lines in contract
                    If tmActiveCnts(ilActiveCntInx).lPop < 0 Then       'never been set up, save population first time
                        If llPop <> 0 Then
                            tmActiveCnts(ilActiveCntInx).lPop = llPopEst
                        End If
                    ElseIf tmActiveCnts(ilActiveCntInx).lPop <> 0 Then
                        If tmActiveCnts(ilActiveCntInx).lPop <> llPopEst And llPopEst <> 0 Then
                            tmActiveCnts(ilActiveCntInx).lPop = 0
                        End If
                    End If
                    'Keep track if varying populations within same schedule line (across days or weeks)
                    'If varying within sch line, cant product grp numbers, in which case it should be zerod
                    ilUpperPop = UBound(tmpoplinklist)
                    If tmActiveCnts(ilActiveCntInx).iFirstPopLink = -1 Then      'first time thru for this contract
                        tmActiveCnts(ilActiveCntInx).iFirstPopLink = ilUpperPop
                        If llPopEst <> 0 Then
                            tmpoplinklist(ilUpperPop).lPop = llPopEst
                        End If
                        tmpoplinklist(ilUpperPop).iLine = tmClf.iLine
                        tmpoplinklist(ilUpperPop).iNextLink = -1
                        ReDim Preserve tmpoplinklist(0 To ilUpperPop + 1) As POPLINKLIST
                    Else            'already found at least one spot for this contract , see if varying pops
                        ilUpperPop = UBound(tmpoplinklist)
                        ilFound = False
                        ilLink = tmActiveCnts(ilActiveCntInx).iFirstPopLink
                        Do While ilLink <> -1
                            If tmpoplinklist(ilLink).iLine = tmClf.iLine Then
                                ilFound = True
                                If tmpoplinklist(ilLink).lPop < 0 Then       'never been set up, save population first time
                                    If llPopEst <> 0 Then
                                        tmpoplinklist(ilLink).lPop = llPopEst
                                    End If
                                ElseIf tmpoplinklist(ilLink).lPop <> 0 Then
                                    If tmpoplinklist(ilLink).lPop <> llPopEst And llPopEst <> 0 Then
                                        tmpoplinklist(ilLink).lPop = 0
                                        tmActiveCnts(ilActiveCntInx).iVaryPop = 1   'cant compute cpps from grps
                                    End If
                                End If
                                ilLink = -1     'force to stop search
                            Else                'not matching line #
                                ilKeepLast = ilLink
                                ilLink = tmpoplinklist(ilLink).iNextLink
                            End If
                        Loop
                        If Not ilFound Then             'set up new entry
                            tmpoplinklist(ilKeepLast).iNextLink = ilUpperPop
                            If llPopEst <> 0 Then
                                tmpoplinklist(ilUpperPop).lPop = llPopEst
                            End If
                            tmpoplinklist(ilUpperPop).iLine = tmClf.iLine
                            tmpoplinklist(ilUpperPop).iNextLink = -1
                            ReDim Preserve tmpoplinklist(0 To ilUpperPop + 1) As POPLINKLIST
                        End If
                    End If
                End If

                'tmGrf.lDollars(1) = 0
                'tmGrf.lDollars(2) = 0
                'tmGrf.lDollars(3) = 0
                'tmGrf.lDollars(4) = 0
                'tmGrf.lDollars(5) = 0
                tmGrf.lDollars(0) = 0
                tmGrf.lDollars(1) = 0
                tmGrf.lDollars(2) = 0
                tmGrf.lDollars(3) = 0
                tmGrf.lDollars(4) = 0
                'If tmSdf.sSchStatus = "M" Or tmSdf.sSchStatus = "C" Or tmSdf.sSchStatus = "H" Then
                '    tmGrf.lDollars(1) = llTotalGrImp
                'Else                   'not missed, cancelled or hidden.  Accumulate for Charged, nc, adu orfills
                    If Trim$(slPrice) = ".00" Or Trim$(slPrice) = "N/C" Or Trim$(slPrice) = "Bonus" Or Trim$(slPrice) = "Spinoff" Or Trim$(slPrice) = "ADU" Or Trim$(slPrice) = "Recapturable" Or Trim$(slPrice) = "Extra" Or InStr(slPrice, "Fill") <> 0 Then   '3-17-12 test for +/-Fill Trim$(slPrice) = "+ Fill" Then         'its a .00 spot, N/c or bonus
                        'tmGrf.lDollars(3) = llTotalGrImp
                        tmGrf.lDollars(2) = llTotalGrImp
                    ElseIf Trim$(slPrice) = "MG" Or tmSdf.sSpotType = "G" Then     'schedule line rate defined as MG or SDF spot as a MG
                        'tmGrf.lDollars(4) = llTotalGrImp
                        tmGrf.lDollars(3) = llTotalGrImp
                    Else        'charge spot
                        'tmGrf.lDollars(2) = llTotalGrImp
                        tmGrf.lDollars(1) = llTotalGrImp
                    End If
                'End If
                'save the region to calculate over/under
                tmGrf.lLong = tmClf.lRafCode    'region code
                If tmSdf.sPtType = 1 Then
                    'tmGrf.lDollars(5) = tmSdf.lCopyCode 'copy inventory code to access the product, isci, creative title
                    tmGrf.lDollars(4) = tmSdf.lCopyCode 'copy inventory code to access the product, isci, creative title
                ElseIf tmSdf.sPtType = 3 Then           'time zone copy
                    'N/A
                End If

                'write out the spot record
                tmGrf.lChfCode = tmSdf.lChfCode    'used for contract # and Advertiser
                tmGrf.iCode2 = tmSdf.iLineNo     'schedule line
                tmGrf.iDate(0) = tmSdf.iDate(0) 'scheduled date
                tmGrf.iDate(1) = tmSdf.iDate(1)
                tmGrf.iTime(0) = tmSdf.iTime(0) 'scheduled time
                tmGrf.iTime(1) = tmSdf.iTime(1)
                tmGrf.lCode4 = tmChf.lGuar      'guaranteed aud in %
                tmGrf.sBktType = tmSdf.sSchStatus   'sched status for missed flag

                gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                tmGrf.lGenTime = lgNowTime
                tmGrf.iGenDate(0) = igNowDate(0)
                tmGrf.iGenDate(1) = igNowDate(1)

                'prepass fields for Post Buy
                'GrfGenTime - generation time for filtering records to print
                'grfGenDate - generation date for filtering records to print
                'grfChfcode - contract code
                'grfDate(0 to 1) - spot scheduled date
                'grfCode4 - contract guaranteed %
                'grfCode2 - schedule line #

                'GrfLong - region code for % of audience to multiple the books audience with
                'grfDollars(1) - gross impressions missed
                'grfDollars(2) - gross impressions paid
                'grfDollars(3) - gross impressions bonus/free (all $0 rates except MG spot rate type)
                'grfDollars(4) - gross impressions mg
                'grfDollars(5) - copy pointer for isci, creative title, product name
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)

            End If              'if ilOK
        Next ilSpotLoop
    Next ilVehicle

    'debugging only for time program took to run
'    slStr = Format$(gNow(), "h:mm:ssAM/PM")       'end time of run
'    gPackTime slStr, ilNowTime(0), ilNowTime(1)
'    gUnpackTimeLong ilNowTime(0), ilNowTime(1), False, llTime
'    gUnpackTimeLong igNowTime(0), igNowTime(1), False, llPop   'start time of run
'    llPop = llPop - llTime              'time in seconds in runtime
'    ilRet = gSetFormula("RunTime", llPop)  'show how long report generated

    sgCntrForDateStamp = ""     'initialize contract routine next time thru
    Erase tmActiveCnts, tmVehicleBook, tmDnfLinkList, tmBookList, tmPLSdf
    Erase tgCffAD, tgClfAD
    mCloseFiles
    Exit Sub
mTerminate:
    On Error GoTo 0
    Exit Sub
End Sub
'
'           mOpenDelivery - open all applicables files for Audience Delivery
'           and Post Buy Analysis reports
'           <return>  true if some kind of I/o error
'
Public Function mOpenDelivery() As Integer
Dim ilRet As Integer
Dim slTable As String * 3
Dim ilError As Integer

    ilError = False
    On Error GoTo mOpenDeliveryErr

    slTable = "Grf"
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imGrfRecLen = Len(tmGrf)

    slTable = "Clf"
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imClfRecLen = Len(tmClf)

    slTable = "Cff"
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCffRecLen = Len(tmCff)

    slTable = "Chf"
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imCHFRecLen = Len(tmChf)

    slTable = "Dnf"
    hmDnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imDnfRecLen = Len(tmDnf)

    slTable = "Drf"
    hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imDrfRecLen = Len(tmDrf)

    slTable = "Vef"
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imVefRecLen = Len(tmVef)

    slTable = "Sdf"
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSdfRecLen = Len(tmSdf)

    slTable = "Smf"
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imSmfRecLen = Len(tmSmf)

    slTable = "Vsf"
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imVsfRecLen = Len(tmVsf)

    slTable = "Mnf"
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imMnfRecLen = Len(tmMnf)

    slTable = "Dpf"
    hmDpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imDpfRecLen = Len(tmDpf)

    slTable = "Def"
    hmDef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)

    slTable = "Raf"
    hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    imRafRecLen = Len(tmRaf)

    ReDim tgClfAD(0 To 0) As CLFLIST
    tgClfAD(0).iStatus = -1 'Not Used
    tgClfAD(0).lRecPos = 0
    tgClfAD(0).iFirstCff = -1
    ReDim tgCffAD(0 To 0) As CFFLIST
    tgCffAD(0).iStatus = -1 'Not Used
    tgCffAD(0).lRecPos = 0
    tgCffAD(0).iNextCff = -1


    If ilError Then
        ilRet = btrClose(hmRaf)
        ilRet = btrClose(hmDef)
        ilRet = btrClose(hmDpf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmDnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)

        btrDestroy hmRaf
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmMnf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault

    End If
    mOpenDelivery = ilError
    Exit Function

mOpenDeliveryErr:
    ilError = True
    gBtrvErrorMsg ilRet, "mOpenDelivery (OpenError) #" & str(ilRet) & ": " & slTable, RptSelAD
    Resume Next

End Function
Public Sub mCloseFiles()
Dim ilRet As Integer

        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmDnf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmDpf)
        ilRet = btrClose(hmDef)
        ilRet = btrClose(hmRaf)
        btrDestroy hmRaf
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmMnf
        btrDestroy hmVsf
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Exit Sub
End Sub
