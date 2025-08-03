Attribute VB_Name = "RPTCR30"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptCr30.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  imUseCode                                                                             *
'******************************************************************************************

Option Explicit
Option Compare Text

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
Dim hmRaf As Integer            'regions

Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer
Dim tmVsf As VSF
Dim imVsfRecLen As Integer
Dim tmActiveCnts() As ACTIVECNTS       'Array of all active contracts, where all aud infor will be stored

Dim bmMissingDnfCode As Boolean         '1-10-20 flag if missing a scheduleline book reference
Dim smGrossNet As String * 1          '1-28-19 gross or net option
Dim imInclAdvtCodes As Integer
Dim imUseAdvtCodes() As Integer
Dim imInclVefCodes As Integer
Dim imUsevefcodes() As Integer
Dim imInclVGCodes As Integer
Dim imUseVGCodes() As Integer
Dim imPerType As Integer                    '0 = cal (not implemented), 1 = corp (not implemented), 2 = std (only option)
Dim lmSingleCntr As Long
Dim imPeriods As Integer
Dim imSortBy As Integer                 '0 = advt, 1 = vehicle
'Dim lmStartDates(1 To 13) As Long       'max 12 months, plus start of 13th for last month
Dim lmStartDates(0 To 13) As Long       'max 12 months, plus start of 13th for last month. Index zero ignored
'Dim lmProject(1 To 13) As Long     'Not used
Dim imBook As Integer           '2 = closest book to air dates,  1 = schedule line book, 2 = vehicle default
Dim imCPPCPM As Integer         '0 = cpp, 1 = cpm
Dim imMajorSet As Integer      'vehicle group selected
Dim imDemo As Integer           '-1 : use primary demo from header, else mnfdemo code
Dim imMnfDemoCode As Integer
Dim smDP As String * 1           'Show results by vehicle:  O= use override DPs from sch line, S = use std DP from sch line.  Removed:  always show only by std DP
Dim bmIncludeDPDetail As Boolean    'For vehicle option only:  true to include dp detail by line
Dim tmCntTypes As CNTTYPES
Dim tmSpotLenRatio As SPOTLENRATIO
'Dim imWeeksPerMonth(1 To 12) As Integer '# weeks per std bdcst month, for 1 year
Dim imWeeksPerMonth(0 To 12) As Integer '# weeks per std bdcst month, for 1 year. Index zero ignored
Dim lmPopForAll() As Long
Dim imVehListForAll() As Integer
Dim lmPop() As Long                   'array of population by vehicle
Dim tmWklyInfoByDP() As WKLYINFOBYDP        'contract weekly aud and pop estimates for a contract, broken out by vehicle & DP
Dim tmMonthInfoByCnt() As MONTHINFOBYVEHICLE    'totals by line
Dim tmMonthInfoByVehicle() As MONTHINFOBYVEHICLE    'totals by vehicle
Dim tmMonthInfoFinals() As MONTHINFOBYVEHICLE       'totals by contract
Dim tmVehicle30UnitDetail() As VEHICLE30UNITDetail      'vehicle option, array of summary line items by vehicle and DP
Dim tmVehicle30UnitHdr() As VEHICLE30UnitHDR      'array of header infro by vehicle and DP, creating a linked list to tmMonthInfoByVehicleDet array
Dim tmChfAdvtExt() As CHFADVTEXT

Type WKLYINFOBYDP                               'weekly info per line of the years spots/aud/rate to send to get Line research totals
    iVefCode As Integer
    iRdfCode As Integer
    iDays(0 To 6) As Integer
    lOVStartTime As Long
    lOVEndTime As Long
    iLineNo As Integer
    'lAvgAud(1 To 53) As Long
    'lPopEst(1 To 53) As Long
    'lSpots(1 To 53) As Long
    'lRate(1 To 53) As Long
    lAvgAud(0 To 52) As Long
    lPopEst(0 To 52) As Long
    lSpots(0 To 52) As Long
    lRate(0 To 52) As Long
End Type

Type MONTHINFOBYVEHICLE
    lCntrNo As Long
    iAdfCode As Integer
    imnfDemoCode As Integer
    bmMissingBookCode As Boolean        '1-21-20
    iVefCode As Integer
    iRdfCode As Integer
    iDays(0 To 6) As Integer
    lOVStartTime As Long
    lOVEndTime As Long
    'lTotalCost(1 To 13) As Long
    'lAvgAud(1 To 13) As Long
    'iAvgRtg(1 To 13) As Integer
    'lGrImp(1 To 13) As Long
    'lGRP(1 To 13) As Long
    'lCPP(1 To 13) As Long
    'lCPM(1 To 13) As Long
    'lPopEst(1 To 13) As Long
    'lSpots(1 To 13) As Long

    'lTotalCost(0 To 12) As Long
    'fTotalCost(0 To 12) As Single
    dTotalCost(0 To 12) As Double 'TTP 10439 - Rerate 21,000,000
    lAvgAud(0 To 12) As Long
    iAvgRtg(0 To 12) As Integer
    lGrImp(0 To 12) As Long
    lGRP(0 To 12) As Long
    lCPP(0 To 12) As Long
    lCPM(0 To 12) As Long
    lPopEst(0 To 12) As Long
    lSpots(0 To 12) As Long
End Type

Type VEHICLE30UnitHDR          'CPP/CPM Vehicle Option
    lFirstIndex As Long            '-1 initalized, else the index to the next element.  Link list pointing to array of MonthInfoByVehicleDet information
    lLastIndex As Long
    iVefCode As Integer
    iRdfCode As Integer
    lPop As Long
    iDays(0 To 6) As Integer
    lOVStartTime As Long
    lOVEndTime As Long
End Type

Type VEHICLE30UNITDetail         'CPP/CPM Vehicle Option to retain all the lines summary by DP and vehicle
    lNextIndex As Long              'pointer to next index in chain.  -1 indicates no more
    'lTotalCost(1 To 13) As Long
    'lAvgAud(1 To 13) As Long
    'iAvgRtg(1 To 13) As Integer
    'lGrImp(1 To 13) As Long
    'lGRP(1 To 13) As Long
    'lCPP(1 To 13) As Long
    'lCPM(1 To 13) As Long
    'lPopEst(1 To 13) As Long
    'lSpots(1 To 13) As Long

    'lTotalCost(0 To 12) As Long
    'fTotalCost(0 To 12) As Single
    dTotalCost(0 To 12) As Double
    lAvgAud(0 To 12) As Long
    iAvgRtg(0 To 12) As Integer
    lGrImp(0 To 12) As Long
    lGRP(0 To 12) As Long
    lCPP(0 To 12) As Long
    lCPM(0 To 12) As Long
    lPopEst(0 To 12) As Long
    lSpots(0 To 12) As Long
End Type

'********************************************************************************************
'
'              gCrCP30UnitVehicle - Prepass for CPP/CPM 30"Unit Report
'
'
'       Produce prepass for cpp/cpm for 30" units by  vehicle.  The generated
'       data will be obtained from contracts, using the contract's primary demo, or a selected demo.
'       Up to 12 monthly values are obtained for std months (corp & Calendar not implemented).
'       Calc can be based on schedule line book default book.   book closest to spot air date/date
'       (this option not implemented due to retrieving from contracts, not spots).
'       For spots that are not 30 in length, user makes the rules to calculate the value of
'       spots not divisible by 30".  for example, a 15" will be 1/2 of a 30.
'
'********************************************************************************************
Sub gCrCP30UnitVehicle()
    ReDim ilNowTime(0 To 1) As Integer    'end time of run
    Dim slStr As String
    Dim ilRet As Integer
    Dim llEarliestStart As Long     'earliest start date from all contracts to process
    Dim llLatestEnd As Long         'latest end date from all contracts to process
    Dim slEarliestStart As String   'earliest start date from all contracts to process
    Dim slLatestEnd As String       'latest end date from all contrcts to process
    Dim ilVehicle As Integer        'loop to process spots: gather by one vehicle at a time
    Dim ilVefCode As Integer
    Dim ilVefIndex As Integer
    Dim ilOk As Integer
    Dim blFound As Boolean
    Dim ilDay As Integer
    Dim llTime As Long              'avail time
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim llPop As Long               'pop of demo
    Dim llAvgAud As Long            'aud for demo
    Dim illoop As Integer
    Dim ilDnfCode As Integer        'book to use for spot obtained
    Dim llDate As Long
    Dim ilClf As Integer            'sched line processing loop
    
    Dim llLoop As Long               '1-15-08
    Dim ilOpenError As Integer
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim slPopLineType As String * 1
    Dim ilPopRdfCode As Integer
    Dim llPopRafCode As Long
    Dim llPopEst As Long
    ReDim llkeyarray(0 To 0) As Long
    Dim ilLoopOnKey As Integer
    Dim llKeyCode As Long
    Dim blValidCType As Boolean
    Dim ilUpperClf As Integer
    Dim llResearchPop As Long
    'ReDim llPopByLine(1 To 1) As Long
    ReDim llPopByLine(0 To 0) As Long
    'ReDim ilVehList(1 To 1) As Integer        'list of unique vehicles
    ReDim ilVehList(0 To 0) As Integer        'list of unique vehicles
    Dim ilCff As Integer
    Dim ilfirstTime As Integer
    Dim ilSocEcoMnfCode As Integer
    Dim ilInputDays(0 To 6) As Integer          'valid days of the week (true/false), if daily, # spots/day
    Dim ilDemoAvgAudDays(0 To 6) As Integer     'valid days of the week (true/false)
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim ilSpots As Integer
    Dim llDate2 As Long
    Dim ilWeekInx As Integer
    Dim ilLoopOnInfo As Integer
    Dim blFoundWklyInfoDP As Boolean
    Dim ilUpperWklyInfoByDP As Integer
    Dim blAtLeast1FlightFound As Boolean
    Dim llLineStartDate As Long
    Dim llLineEndDate As Long
    Dim ilHowManyUnits As Integer
    Dim ilTemp As Integer
    Dim ilTemp2 As Integer
    Dim ilVehicleForAll As Integer
    Dim blValidSpotType As Boolean

    ilOpenError = mOpen30Unit()           'open applicable files
    If ilOpenError Then
        Exit Sub            'at least 1 open error
    End If

    mObtainSelectivity
    
    llEarliestStart = lmStartDates(1)
    llLatestEnd = lmStartDates(imPeriods + 1) - 1
    slEarliestStart = Format$(llEarliestStart, "m/d/yy")
    slLatestEnd = Format$(llLatestEnd, "m/d/yy")
           
    For illoop = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1
        'filter out the advertisers  or contract types not requested
        If gFilterLists(tmChfAdvtExt(illoop).iAdfCode, imInclAdvtCodes, imUseAdvtCodes()) Then
            'filtering of contract header types should have been filtered at the time of obtains the active contracts
            blValidCType = gFilterContractType(tmChf, tmCntTypes, False)         'exclude proposal type checks
            If blValidCType Then            'if valid, build into array; otherwise bypass
                llkeyarray(UBound(llkeyarray)) = tmChfAdvtExt(illoop).lCode
                ReDim Preserve llkeyarray(0 To UBound(llkeyarray) + 1) As Long
            End If
        End If
    Next illoop
    
    ilSocEcoMnfCode = 0         'not using socio-economic codes for research
    
    'ReDim tmMonthInfoFinals(1 To 1) As MONTHINFOBYVEHICLE        'required if option by vehicle
    ReDim tmMonthInfoFinals(0 To 0) As MONTHINFOBYVEHICLE        'required if option by vehicle
    'ReDim tmVehicle30UnitDetail(1 To 1) As VEHICLE30UNITDetail       'Retain vehicle list for all contracts if by vehicle option
    ReDim tmVehicle30UnitDetail(0 To 0) As VEHICLE30UNITDetail       'Retain vehicle list for all contracts if by vehicle option
    'ReDim tmVehicle30UnitHdr(1 To 1) As VEHICLE30UnitHDR
    ReDim tmVehicle30UnitHdr(0 To 0) As VEHICLE30UnitHDR
    'This is to determine if the same book used for each vehicle for ResearchTotals by vehicle
    'ReDim lmPopForAll(1 To 1) As Long                              'population by vehicle for All contracts
    ReDim lmPopForAll(0 To 0) As Long                              'population by vehicle for All contracts
    'ReDim imVehListForAll(1 To 1) As Integer              'vehicles for all contracts selected to process
    ReDim imVehListForAll(0 To 0) As Integer              'vehicles for all contracts selected to process
    
    For ilLoopOnKey = 0 To UBound(llkeyarray) - 1           'looping on chfcode or vehicle code
        llKeyCode = llkeyarray(ilLoopOnKey)
        ilUpperClf = 0                                  'no lines in research table
        llResearchPop = -1                       'pop from book if all same books across lines, else its zero

        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llKeyCode, False, tgChfCT, tgClfCT(), tgCffCT(), False)
        ReDim tmMonthInfoByCnt(0 To 0) As MONTHINFOBYVEHICLE
        ilUpperClf = UBound(tgClfCT)
        'If UBound(tgClfCT) > UBound(llPopByLine) Then                 'if no sch lines, results in error to redim 0
        If UBound(tgClfCT) - 1 > UBound(llPopByLine) Then                 'if no sch lines, results in error to redim 0
            'ReDim llPopByLine(1 To UBound(tgClfCT)) As Long
            ReDim llPopByLine(0 To UBound(tgClfCT) - 1) As Long
        End If

        'ReDim lmPop(1 To 1) As Long                              'population bh vehicle
        ReDim lmPop(0 To 0) As Long                              'population bh vehicle
        'ReDim ilVehList(1 To 1) As Integer              'vehicles in this contract
        ReDim ilVehList(0 To 0) As Integer              'vehicles in this contract
        blAtLeast1FlightFound = False
        
        For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
            tmClf = tgClfCT(ilClf).ClfRec
            ilVefIndex = gBinarySearchVef(tmClf.iVefCode)
            
            'filter vehicle selectivity
            If Not gFilterLists(tmClf.iVefCode, imInclVefCodes, imUsevefcodes()) Then
                ilVefIndex = -1         'not a selected vehicle, bypass
            Else                        'valid vehicle, is the vehicle group OK
                'Setup the major sort factor
                gGetVehGrpSets tmClf.iVefCode, 0, imMajorSet, ilTemp, ilTemp2
                'check selectivity of vehicle groups
                If (imMajorSet > 0) Then
                    If Not gFilterLists(ilTemp2, imInclVGCodes, imUseVGCodes()) Then
                        ilVefIndex = -1
                    End If
                End If
            End If
             
            gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), slStr
            llLineStartDate = gDateValue(slStr)
            gUnpackDate tmClf.iEndDate(0), tmClf.iEndDate(1), slStr
            llLineEndDate = gDateValue(slStr)

            If (tmClf.sType <> "O" And tmClf.sType <> "A" And tmClf.sType <> "E") And ilVefIndex >= 0 And llLineEndDate > llLineStartDate Then         'ignore package lines and CBS lines
                ilHowManyUnits = gDetermineSpotLenRatio(tmClf.iLen, tmSpotLenRatio)     'determine 30" unit of this spot length by the user defined table
                ilHowManyUnits = ilHowManyUnits / 10                                    'do not carry to hundreds
                '6-16-20 if spot length doesnt exist, the # of computed units (based on 30" units), is returned as a negative number.  Use that computed number
                If ilHowManyUnits < 0 Then                                              'len not found in table, use the 30 sec unit
                    ilHowManyUnits = -ilHowManyUnits
                End If
                If ilHowManyUnits = 0 Then                                              'len has to be at least 1 unit (if less than 30 is in table, make it one)
                    ilHowManyUnits = 1
                End If
                ilDnfCode = tmClf.iDnfCode      'assume using sch line book
                If imBook = 1 Then              'use vehicle default book
                    ilDnfCode = tgMVef(ilVefIndex).iDnfCode
                End If
                blFound = False
                For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1 Step 1           'this list starts over with each new contract
                    If tmClf.iVefCode = ilVehList(ilVehicle) Then
                        blFound = True
                        Exit For
                    End If
                Next ilVehicle
                If Not blFound Then
                    ilVehList(UBound(ilVehList)) = tmClf.iVefCode
                    ilVehicle = UBound(ilVehList)
                    'ReDim Preserve ilVehList(1 To UBound(ilVehList) + 1)
                    ReDim Preserve ilVehList(0 To UBound(ilVehList) + 1)
                    'ReDim Preserve lmPop(1 To UBound(lmPop) + 1)
                    ReDim Preserve lmPop(0 To UBound(lmPop) + 1)
'                        imVehListForAll(UBound(imVehListForAll)) = tmClf.iVefCode
'                        ilVehicleForAll = UBound(imVehListForAll)
'                        ReDim Preserve imVehListForAll(1 To UBound(imVehListForAll) + 1)
'                        ReDim Preserve lmPopForAll(1 To UBound(lmPopForAll) + 1)
                End If
                
                'retain vehicle list for the entire run for all contracts combined
                'For ilVehicleForAll = 1 To UBound(imVehListForAll) - 1 Step 1           'this list is for the entire run for all contracts
                For ilVehicleForAll = 0 To UBound(imVehListForAll) - 1 Step 1           'this list is for the entire run for all contracts
                    If tmClf.iVefCode = imVehListForAll(ilVehicleForAll) Then
                        blFound = True
                        Exit For
                    End If
                Next ilVehicleForAll
                If Not blFound Then
                    imVehListForAll(UBound(imVehListForAll)) = tmClf.iVefCode
                    ilVehicleForAll = UBound(imVehListForAll)
                    'ReDim Preserve imVehListForAll(1 To UBound(imVehListForAll) + 1)
                    ReDim Preserve imVehListForAll(0 To UBound(imVehListForAll) + 1)
                    'ReDim Preserve lmPopForAll(1 To UBound(lmPopForAll) + 1)
                    ReDim Preserve lmPopForAll(0 To UBound(lmPopForAll) + 1)
                End If

                'Build population table by vehicle
                If imDemo < 0 Then              '-1 indicates to use primary demo from header
                    imMnfDemoCode = tgChfCT.iMnfDemo(0)
                Else
                    imMnfDemoCode = imDemo
                End If
                ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, ilSocEcoMnfCode, imMnfDemoCode, llPop)
                'If llPop > 0 Then       '9-18-15 bypass any lines without a population
                    llOvStartTime = 0           'assume no override times
                    llOvEndTime = 0
                    'If smDP = "O" Then          'always use the override when doing the research. the orig question was how to show on reports, which has been taken out.
                        If tmClf.iStartTime(0) = 1 And tmClf.iStartTime(1) = 0 Then
                            llOvStartTime = 0
                            llOvEndTime = 0
                        Else
                            'override times exist
                            gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llOvStartTime
                            gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llOvEndTime
                        End If
                    'End If

                    If tgSpf.sDemoEstAllowed <> "Y" Then
                        'llPopByLine(ilClf + 1) = llPop    '11-23-99 save the pop by line (each linemay have different survey books)
                        llPopByLine(ilClf) = llPop    '11-23-99 save the pop by line (each linemay have different survey books)
                        If lmPop(ilVehicle) = 0 Then            'same vehicle found more than once, if pop already stored,
                                                                'dont wipe out with a possible non-population value
                            lmPop(ilVehicle) = llPop            'associate the population with the vehicle
                        End If
                        If llResearchPop = -1 And llPop <> 0 Then          'first time, llResearchPop is for the summary records (-1 first time thru, 0 = different books across vehicles)
                            llResearchPop = llPop
                        Else
                            If (llResearchPop <> 0) And (llResearchPop <> llPop) And (llPop <> 0) Then      'test to see if this pop is different that the prev one.
                                llResearchPop = 0                                           'if different pops, calculate the contract  summary different
                                If llPop <> lmPop(ilVehicle) Then
                                    lmPop(ilVehicle) = -1
                                End If
                            Else
                                'if current line has population, but there was already a different across
                                'lines in pop, dont save new one
                                If llPop <> 0 And (llResearchPop <> 0 And llResearchPop <> -1) Then
                                    llResearchPop = llPop
                                Else
                                    If lmPop(ilVehicle) <> llPop And lmPop(ilVehicle) <> -1 Then
                                        lmPop(ilVehicle) = -1
                                    End If
                                End If
                            End If
                        End If
                        
                        'keep running track of all vehicles for all contracts for varying pops.  This is for vehicle totals at end to calc ResearchTotals
                        If lmPopForAll(ilVehicleForAll) = 0 Then            'same vehicle found more than once, if pop already stored,
                                                                'dont wipe out with a possible non-population value
                            lmPopForAll(ilVehicleForAll) = llPop            'associate the population with the vehicle
                        Else                                    'pop has been set before, either it has a value or its already known to have varying populations
                            If (lmPopForAll(ilVehicleForAll) <> llPop) And (llPop <> 0) Then       'test to see if this pop is different that the prev one.
                                lmPopForAll(ilVehicleForAll) = -1                   'indicate varying pops
                            Else
                                'if current line has population, but there was already a different across
                                'lines in pop, dont save new one
                                If lmPopForAll(ilVehicleForAll) <> llPop And lmPopForAll(ilVehicleForAll) <> -1 Then
                                    lmPopForAll(ilVehicleForAll) = -1
                                End If
                            End If
                        End If
                    End If
                    
                    ReDim tmWklyInfoByDP(0 To 0) As WKLYINFOBYDP        'aud & pop estimates by unique vehicle, DP (and overrides if applicable)
                    ilUpperWklyInfoByDP = 0
                    
                    ilCff = tgClfCT(ilClf).iFirstCff
                    Do While ilCff <> -1
                        tmCff = tgCffCT(ilCff).CffRec
                       
                        blValidSpotType = mFilterSpotType()
                        If blValidSpotType Then                             'if valid spot type continue; otherwise ignore this flight
                            If tgMVef(ilVefIndex).sType = "G" Then          'sports vehicle
                                tmCff.sDyWk = "W"
                            End If
    
                            ilfirstTime = True                  'set to calc avg aud one time only for this flight
                            For illoop = 0 To 6                 'init all days to not airing, setup for research results later
                                ilInputDays(illoop) = False
                                ilDemoAvgAudDays(illoop) = False        'initalize to 0
                            Next illoop
    
                            gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
                            llFltStart = gDateValue(slStr)
                            'backup start date to Monday
                            illoop = gWeekDayLong(llFltStart)
                            Do While illoop <> 0
                                llFltStart = llFltStart - 1
                                illoop = gWeekDayLong(llFltStart)
                            Loop
                            gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
                            llFltEnd = gDateValue(slStr)
                            
                            'process the flight weeks only within the requested report period
                            If llFltStart < llEarliestStart Then        'flight starts earlier than requested period, use requested period start date as starting point
                                llFltStart = llEarliestStart
                            End If
                            If llFltEnd > llLatestEnd Then              'flight end date extends past requested period, use requested period end date as ending point
                                llFltEnd = llLatestEnd
                            End If
                            '
                            'Loop thru the flight by week and build the number of spots for each week
                            For llDate2 = llFltStart To llFltEnd Step 7
                                If ilfirstTime Then                 'only need to determine valid days & # spots once per flight entry
                                    If tmCff.sDyWk = "W" Then            'weekly
                                        ilSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                                        For ilDay = 0 To 6 Step 1
                                            If (llDate2 + ilDay >= llFltStart) And (llDate2 + ilDay <= llFltEnd) Then
                                                If tmCff.iDay(ilDay) > 0 Or tmCff.sXDay(ilDay) = "1" Then
                                                    ilInputDays(ilDay) = True
                                                    ilDemoAvgAudDays(ilDay) = True      '11-19-02 for weekly, each day is indicated by true false as a valid airing day of week
                                                End If
                                            End If
                                        Next ilDay
                                    Else                                        'daily
                                        If illoop + 6 < llFltEnd Then           'we have a whole week
                                            ilSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)    'got entire week
                                            For ilDay = 0 To 6 Step 1
                                                If tmCff.iDay(ilDay) > 0 Then
                                                    ilInputDays(ilDay) = tmCff.iDay(ilDay)
                                                    ilDemoAvgAudDays(ilDay) = True      ' for daily, each day is indicated by # spots per day as a valid airing day
                                                End If
                                            Next ilDay
                                        Else                                    'do partial week
                                            For llDate = llDate2 To llFltEnd Step 1
                                                ilDay = gWeekDayLong(llDate)
                                                ilSpots = ilSpots + tmCff.iDay(ilDay)
                                                If tmCff.iDay(ilDay) > 0 Then
                                                    ilInputDays(ilDay) = tmCff.iDay(ilDay)
                                                    ilDemoAvgAudDays(ilDay) = True      'for daily, each day is indicated by # spots per day as a valid airing day
                                                End If
                                            Next llDate
                                        End If
                                    End If
                                End If
    
                                ilWeekInx = (llDate2 - llEarliestStart) / 7 + 1
                                If ilWeekInx > 0 And ilWeekInx < 54 Then           ' has to be a valid week within requested period
     
                                    If ilfirstTime Then
                                        If tgSpf.sDemoEstAllowed <> "Y" Then
                                            ilfirstTime = False
                                        End If
                                        'Daily and weekly need the valid airing day, not the spots per day if daily (ilDemoAvgAudDays)
                                        ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, tmClf.iVefCode, ilSocEcoMnfCode, imMnfDemoCode, llDate2, llDate2, tmClf.iRdfCode, llOvStartTime, llOvEndTime, ilDemoAvgAudDays(), tmClf.sType, tmClf.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                        'returned llAvgAud and llPopEst
                                    End If
                                    blAtLeast1FlightFound = True            'flag at least one valid flight found for the contract
                                    'Build array of the weekly spots & rates for one line at a time
                                    mGetWeekInfoByDP ilWeekInx, llOvStartTime, llOvEndTime, ilDemoAvgAudDays(), llAvgAud, llPopEst, (ilSpots * ilHowManyUnits)
    
                                End If
                                
                            Next llDate2
                        End If                                      'valid spot type
                        ilCff = tgCffCT(ilCff).iNextCff               'get next flight record from mem
                    Loop                                            'while ilcff <> -1
                    'at the end of each line, get the monthly values of audience, rating, grimps, cpp, cpm for this one line
                    'mGetVehicleLineMonthInfo llPopByLine(ilClf + 1)            'get line totals
                    If ilDnfCode = 0 Then                      '1-10-20 flag this in the report as at least one line with missing demo reference
                        bmMissingDnfCode = True
                    End If
                    mGetVehicleLineMonthInfo llPopByLine(ilClf), tmClf.iDnfCode            'get line totals
                     
                'End If
            End If
         Next ilClf                 'loop on contracts
         'All contracts have to be built to get the information for vehicle totals
        mGetVehicleDPFinals       'get vehicle & DP totals
    Next ilLoopOnKey                    'loop on chfcode or vehicle code
    mInsertCntFinalInfo "D"             'write out vehicle/dp detail records
    mGetVehicleFinals                   'create the vehicle summaries
    mInsertCntFinalInfo "V"


    'debugging only for time program took to run
    slStr = Format$(gNow(), "h:mm:ssAM/PM")       'end time of run
    gPackTime slStr, ilNowTime(0), ilNowTime(1)
    gUnpackTimeLong ilNowTime(0), ilNowTime(1), False, llTime
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, llPop   'start time of run
    llPop = llPop - llTime              'time in seconds in runtime
    ilRet = gSetFormula("RunTime", llPop)  'show how long report generated

    sgCntrForDateStamp = ""     'initialize contract routine next time thru
    Erase tmChfAdvtExt
    Erase tgCffCT, tgClfCT
    Erase llkeyarray, tmWklyInfoByDP
    Erase tmVehicle30UnitDetail, tmVehicle30UnitHdr
    Erase lmPop
    Erase ilVehList
    Erase lmPopForAll, imVehListForAll

    mCloseFiles
    Exit Sub
mTerminate:
    On Error GoTo 0
    Exit Sub
End Sub

'
'           mOpen30Unit - open all applicables files for CPPCPM 30" Unit report
'           <return>  true if some kind of I/o error
'
Public Function mOpen30Unit() As Integer
    Dim ilRet As Integer
    Dim slTable As String * 3
    Dim ilError As Integer

    ilError = False
    On Error GoTo mOpen30UnitErr

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
    
    ReDim tgClfCT(0 To 0) As CLFLIST
    tgClfCT(0).iStatus = -1 'Not Used
    tgClfCT(0).lRecPos = 0
    tgClfCT(0).iFirstCff = -1
    ReDim tgCffCT(0 To 0) As CFFLIST
    tgCffCT(0).iStatus = -1 'Not Used
    tgCffCT(0).lRecPos = 0
    tgCffCT(0).iNextCff = -1


    If ilError Then
        ilRet = btrClose(hmDef)
        ilRet = btrClose(hmDpf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmVsf)

        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmDrf)
        ilRet = btrClose(hmDnf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmRaf)
        btrDestroy hmDef
        btrDestroy hmDpf
        btrDestroy hmMnf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmDrf
        btrDestroy hmDnf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmRaf
        Screen.MousePointer = vbDefault
    End If
    mOpen30Unit = ilError
    Exit Function

mOpen30UnitErr:
    ilError = True
    gBtrvErrorMsg ilRet, "mOpen30Unit (OpenError) #" & str(ilRet) & ": " & slTable, RptSel30
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
    btrDestroy hmVef
    btrDestroy hmDrf
    btrDestroy hmDnf
    btrDestroy hmCHF
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    Exit Sub
End Sub

Private Sub mObtainSelectivity()
    Dim slStart As String
    Dim ilVGSort As Integer
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim slLen(0 To 9) As String
    Dim slIndex(0 To 9) As String
    Dim slEarliestStart As String
    Dim slLatestEnd As String
    Dim slCntrTypes As String
    Dim slCntrStatus As String
    Dim ilHOState As Integer

   
    For illoop = 0 To 9
        slLen(illoop) = Trim$(RptSel30!edcLen(illoop))
        slIndex(illoop) = Trim$(RptSel30!edcIndex(illoop))
    Next illoop
    gBuildSpotLenAndIndexTable slLen(), slIndex(), tmSpotLenRatio
    
   If RptSel30!rbcMonthType(0).Value = True Then           'cal
        imPerType = 0
    ElseIf RptSel30!rbcMonthType(1).Value = True Then       'corp
        imPerType = 1
    Else                                                    'std
        imPerType = 2
    End If
       
    imPeriods = Val(RptSel30!edcNoMonths.Text)
    
    If imPerType = 2 Then   'set start dates of 12 standard periods
        slStart = str$(igMonthOrQtr) & "/15/" & str$(igYear)
        gBuildStartDates slStart, 1, 13, lmStartDates()
        'Some months may be 4 or 5 week months, determine # of weeks for each
        For illoop = 1 To 12                '12 months
            imWeeksPerMonth(illoop) = (lmStartDates(illoop + 1) - lmStartDates(illoop)) / 7
        Next illoop

    'other month type not completed in feature
    ElseIf imPerType = 1 Then   'set start dates of 12 corporate periods
        slStart = str$(igMonthOrQtr) & "/15/" & str$(igYear)
        gBuildStartDates slStart, 2, imPeriods + 1, lmStartDates()
        'Some months may be 4 or 5 week months, determine # of weeks for each
        For illoop = 1 To 12                '12 months
            imWeeksPerMonth(illoop) = (lmStartDates(illoop) + 1 - lmStartDates(illoop)) / 7
        Next illoop

    ElseIf imPerType = 0 Then  'set start dates of 12 calendar periods
        slStart = str$(igMonthOrQtr) & "/1/" & str$(igYear)
        gBuildStartDates slStart, 4, imPeriods + 1, lmStartDates()
    End If

    'Selective contract #
    lmSingleCntr = Val(RptSel30!edcContract.Text)
    
    tmCntTypes.iHold = gSetCheck(RptSel30!ckcCType(0).Value)
    tmCntTypes.iOrder = gSetCheck(RptSel30!ckcCType(1).Value)
    tmCntTypes.iStandard = gSetCheck(RptSel30!ckcCType(3).Value)
    tmCntTypes.iReserv = gSetCheck(RptSel30!ckcCType(4).Value)
    tmCntTypes.iRemnant = gSetCheck(RptSel30!ckcCType(5).Value)
    tmCntTypes.iDR = gSetCheck(RptSel30!ckcCType(6).Value)
    tmCntTypes.iPI = gSetCheck(RptSel30!ckcCType(7).Value)
    tmCntTypes.iPSA = gSetCheck(RptSel30!ckcCType(8).Value)
    tmCntTypes.iPromo = gSetCheck(RptSel30!ckcCType(9).Value)
    tmCntTypes.iTrade = gSetCheck(RptSel30!ckcCType(10).Value)
    tmCntTypes.iPolit = gSetCheck(RptSel30!ckcCType(2).Value)
    tmCntTypes.iNonPolit = gSetCheck(RptSel30!ckcCType(11).Value)
            
    'line types to include:  only chargeable for revenue; otherwise all
    tmCntTypes.iCharge = gSetCheck(RptSel30!ckcSpotType(0).Value)
    tmCntTypes.iZero = gSetCheck(RptSel30!ckcSpotType(1).Value)
    tmCntTypes.iADU = gSetCheck(RptSel30!ckcSpotType(2).Value)
    tmCntTypes.iBonus = gSetCheck(RptSel30!ckcSpotType(3).Value)
    tmCntTypes.iNC = gSetCheck(RptSel30!ckcSpotType(4).Value)
    tmCntTypes.iMG = gSetCheck(RptSel30!ckcSpotType(5).Value)
    tmCntTypes.iRecapturable = gSetCheck(RptSel30!ckcSpotType(6).Value)
    tmCntTypes.iSpinoff = gSetCheck(RptSel30!ckcSpotType(7).Value)
    'ckcCTypes(8-11) unused
    ReDim imUseAdvtCodes(0 To 0) As Integer
    ReDim imUsevefcodes(0 To 0) As Integer
    gObtainCodesForMultipleLists 2, tgCSVNameCode(), imInclVefCodes, imUsevefcodes(), RptSel30
    gObtainCodesForMultipleLists 0, tgAdvertiser(), imInclAdvtCodes, imUseAdvtCodes(), RptSel30
    
    imBook = 2                    'use closest book to air dates
    If RptSel30!rbcBook(0).Value Then       'use sched line book
        imBook = 0
    ElseIf RptSel30!rbcBook(1).Value Then       'vehicle default book
        imBook = 1
    End If

    imSortBy = 0                                'advt
    bmIncludeDPDetail = False                   'default to ignore any detail by line on output
    If RptSel30!rbcSortBy(1).Value Then         'vehicle
        imSortBy = 1
        If RptSel30!ckcDetail.Value = vbChecked Then
            bmIncludeDPDetail = True
        End If
    End If
    
     imCPPCPM = 0                'assume CPP
    If RptSel30!rbcByCPPCPM(1).Value Then
        imCPPCPM = 1            'cpm
    End If
    
    'The option has been removed but defaulted to show STD DPs on report.  Feature is not fully implemented if it needs to be re-instated.  Question:  will they need to see breakout by overrides with days of week
    smDP = "O"           'use override DP
    If (RptSel30!rbcDP(1).Value And RptSel30!rbcSortBy(1).Value) Or RptSel30!rbcSortBy(0).Value Then      'use std DP if by advt, or selecting vehicle and user selected std DP
        smDP = "S"       'use std DP
    End If
    
    smGrossNet = "G"                                '1-28-19 implement gross or net
    If RptSel30!rbcGrossNet(1).Value Then           'net selected
        smGrossNet = "N"
    End If
    
    imMajorSet = 0
    ilVGSort = RptSel30!cbcSet1.ListIndex

    If ilVGSort > 0 Then
        illoop = RptSel30!cbcSet1.ListIndex
        imMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())
        gObtainCodesForMultipleLists 3, tgSOCode(), imInclVGCodes, imUseVGCodes(), RptSel30
    End If

    If RptSel30!rbcDemo(0).Value = True Then      'use contract header primary
        imDemo = -1
    Else                                        'get the demo from listbox
        For illoop = 0 To RptSel30!lbcSelection(1).ListCount - 1 Step 1
            If RptSel30!lbcSelection(1).Selected(illoop) Then
                slNameCode = tgRptSelDemoCodeCP(illoop).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                imDemo = Val(slCode)
                Exit For
            End If
        Next illoop
    End If
    
    gSetupCntTypesForGet tmCntTypes, slCntrTypes, slCntrStatus
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)

    slEarliestStart = Format$(lmStartDates(1), "m/d/yy")
    slLatestEnd = Format$(lmStartDates(imPeriods + 1) - 1, "m/d/yy")

    If lmSingleCntr <> 0 Then                  'single contract has been selected since there is an internal code in the list, not all
        ReDim tmChfAdvtExt(0 To 1) As CHFADVTEXT
        tmChfSrchKey1.lCntrNo = lmSingleCntr
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, Len(tmChf), tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mCloseFiles
            Exit Sub
        End If
        tmChfAdvtExt(0).lCode = tmChf.lCode
        tmChfAdvtExt(0).iAdfCode = tmChf.iAdfCode
        tmChfAdvtExt(0).lCntrNo = tmChf.lCntrNo
    Else        'all contracts, retrieve active ones based on the dates entered
        ilRet = gObtainCntrForDate(RptSel30, slEarliestStart, slLatestEnd, slCntrStatus, slCntrTypes, ilHOState, tmChfAdvtExt())
        If ilRet <> 0 Then
            mCloseFiles
            Exit Sub
        End If
    End If

    bmMissingDnfCode = False            '1-10-20 assume all lines have a book reference
    
    '7-23-01 setup global variable to determine if demo plus info exists (gGetDemoAud)
    lgDpfNoRecs = btrRecords(hmDpf)
    If lgDpfNoRecs = 0 Then
        lgDpfNoRecs = -1
    End If
    Exit Sub
End Sub

'
'           mGetWklyInfoByDP - find the matching vehicle and DP entry to place the research results
'           Roll over the weekly totals into monthly totals for each vehicle/DP
'           <input> ilWeekInx - week index relative to start of period requested
'                   llOVStartTime - override start time if applicable
'                   llOVEndTime = override end time if applicable
'                   llDemoAvgAudDays - array of valid days of week
'                   llAvgAud - avg aud for week
'                   llPopEst - pop estimate
'                   ilSpots - spot count for week
'           <output>  tmWklyInfoBYDp with new entry or updated
Public Sub mGetWeekInfoByDP(ilWeekInx As Integer, llOvStartTime As Long, llOvEndTime As Long, ilDemoAvgAudDays() As Integer, llAvgAud As Long, llPopEst As Long, ilSpots As Integer)
    Dim blFoundWklyInfoDP As Boolean
    Dim ilLoopOnInfo As Integer
    Dim ilUpperWklyInfoByDP As Integer
    Dim ilDay As Integer
    Dim llRate As Long

    blFoundWklyInfoDP = False
    ilUpperWklyInfoByDP = UBound(tmWklyInfoByDP)
    For ilLoopOnInfo = LBound(tmWklyInfoByDP) To ilUpperWklyInfoByDP - 1
        
        If smDP = "O" Then      'use overrides; if by advt its defaulted to using std DP.  Feature has been Defaulted to use std dp
            'Feature has been defaulted to "S" (show std DP on report), this code will not be executed.
            If tmWklyInfoByDP(ilLoopOnInfo).iVefCode = tmClf.iVefCode And tmWklyInfoByDP(ilLoopOnInfo).iRdfCode = tmClf.iRdfCode And tmWklyInfoByDP(ilLoopOnInfo).lOVStartTime = llOvStartTime And tmWklyInfoByDP(ilLoopOnInfo).lOVEndTime = llOvEndTime Then
                'vehicle, DP, override start & end times match; test the override days of week
                For ilDay = 0 To 6
                    If tmWklyInfoByDP(ilLoopOnInfo).iDays(ilDay) <> ilDemoAvgAudDays(ilDay) Then
                        blFoundWklyInfoDP = False
                        Exit For
                    End If
                    blFoundWklyInfoDP = True        'found a valid matching DP entry
                Next ilDay
            End If
        Else                    'show results by standard DPs on vehicle report option
            If tmWklyInfoByDP(ilLoopOnInfo).iVefCode = tmClf.iVefCode And tmWklyInfoByDP(ilLoopOnInfo).iRdfCode = tmClf.iRdfCode Then
                blFoundWklyInfoDP = True
            End If
        End If
        If blFoundWklyInfoDP Then
            tmWklyInfoByDP(ilLoopOnInfo).lAvgAud(ilWeekInx - 1) = llAvgAud
            tmWklyInfoByDP(ilLoopOnInfo).lPopEst(ilWeekInx - 1) = llPopEst
            tmWklyInfoByDP(ilLoopOnInfo).lSpots(ilWeekInx - 1) = tmWklyInfoByDP(ilLoopOnInfo).lSpots(ilWeekInx - 1) + ilSpots
            llRate = gGetGrossOrNetFromRate(tmCff.lActPrice, smGrossNet, tgChfCT.iAgfCode)      '1-28-19 implement gross net option
            'tmWklyInfoByDP(ilLoopOnInfo).lRate(ilWeekInx - 1) = tmCff.lActPrice
            tmWklyInfoByDP(ilLoopOnInfo).lRate(ilWeekInx - 1) = llRate                  'gross or net option, calculated net return if net
            Exit For
        End If
    Next ilLoopOnInfo
    If Not blFoundWklyInfoDP Then
        tmWklyInfoByDP(ilUpperWklyInfoByDP).iVefCode = tmClf.iVefCode
        tmWklyInfoByDP(ilUpperWklyInfoByDP).iRdfCode = tmClf.iRdfCode
        tmWklyInfoByDP(ilUpperWklyInfoByDP).lOVStartTime = llOvStartTime
        tmWklyInfoByDP(ilUpperWklyInfoByDP).lOVEndTime = llOvEndTime
        For ilDay = 0 To 6
            tmWklyInfoByDP(ilUpperWklyInfoByDP).iDays(ilDay) = ilDemoAvgAudDays(ilDay)
        Next ilDay
        tmWklyInfoByDP(ilLoopOnInfo).lAvgAud(ilWeekInx - 1) = llAvgAud
        tmWklyInfoByDP(ilLoopOnInfo).lPopEst(ilWeekInx - 1) = llPopEst
        tmWklyInfoByDP(ilLoopOnInfo).lSpots(ilWeekInx - 1) = ilSpots
        llRate = gGetGrossOrNetFromRate(tmCff.lActPrice, smGrossNet, tgChfCT.iAgfCode)      '10-3-19 implement gross net option
        tmWklyInfoByDP(ilLoopOnInfo).lRate(ilWeekInx - 1) = llRate
'                tmWklyInfoByDP(ilLoopOnInfo).lRate(ilWeekInx - 1) = tmCff.lActPrice
        tmWklyInfoByDP(ilLoopOnInfo).iLineNo = tmClf.iLine

        ilUpperWklyInfoByDP = ilUpperWklyInfoByDP + 1
        ReDim Preserve tmWklyInfoByDP(0 To ilUpperWklyInfoByDP) As WKLYINFOBYDP
    End If
Exit Sub
End Sub

'
'                   mGetLineMonthlyInfo - obtain the monthly Research data for a single line from the weekly summary of spots, rate & audience
'                   Build array of all the line totals for each month + year total (from gAvgAudToLnResearch)
'tmWklyInfoByDP
'imWeeksPerMonth
Public Sub mGetLineMonthInfo(llPop As Long)
    Dim ilWklyInfoByDP As Integer
    Dim ilUpperMonthInfo As Integer
    Dim ilLoopOnMonth As Integer
    Dim ilWeeks As Integer                  '# weeks in period (4 or 5)
    Dim ilWeekInx As Integer                'week index 1-53
    
    
    'line totals - output
    'Dim llLineCost As Long
    Dim dlLineCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim flLineCost As Single
    Dim llLineAvgAud As Long
    Dim ilLineAvgRtg As Integer
    Dim llLineGrimp As Long
    Dim llLineGRP As Long
    Dim llLineCPP As Long
    Dim llLineCPM As Long
    Dim llLinePopEst As Long
    'end line totals-output
    
    Dim blFoundMonthInfo As Boolean
    Dim ilLoopOnInfo As Integer
    Dim ilDay As Integer
    Dim llMonthSpotCount As Long
    Dim ilLoopOnMonthMinusOne As Integer

    For ilWklyInfoByDP = LBound(tmWklyInfoByDP) To UBound(tmWklyInfoByDP) - 1     'single contracts information by vehicle, dp & week
        ilWeekInx = 0
        ilUpperMonthInfo = UBound(tmMonthInfoByCnt)
        tmMonthInfoByCnt(ilUpperMonthInfo).iVefCode = tmWklyInfoByDP(ilWklyInfoByDP).iVefCode
        tmMonthInfoByCnt(ilUpperMonthInfo).iRdfCode = tmWklyInfoByDP(ilWklyInfoByDP).iRdfCode
         
        tmMonthInfoByCnt(ilUpperMonthInfo).lOVStartTime = tmWklyInfoByDP(ilWklyInfoByDP).lOVStartTime
        tmMonthInfoByCnt(ilUpperMonthInfo).lOVEndTime = tmWklyInfoByDP(ilWklyInfoByDP).lOVEndTime
        For ilDay = 0 To 6
            tmMonthInfoByCnt(ilUpperMonthInfo).iDays(ilDay) = tmWklyInfoByDP(ilWklyInfoByDP).iDays(ilDay)
        Next ilDay

        
        For ilLoopOnMonth = 1 To imPeriods
            llMonthSpotCount = 0
            'ReDim llPopEst(1 To imWeeksPerMonth(ilLoopOnMonth)) As Long         'dimension for the number of weeks in the month (4 or 5)
            'ReDim llSpots(1 To imWeeksPerMonth(ilLoopOnMonth)) As Integer
            'ReDim llrate(1 To imWeeksPerMonth(ilLoopOnMonth)) As Long
            'ReDim llAvgAud(1 To imWeeksPerMonth(ilLoopOnMonth)) As Long
            'ReDim llGrImp(1 To imWeeksPerMonth(ilLoopOnMonth)) As Long
            'ReDim llGRP(1 To imWeeksPerMonth(ilLoopOnMonth)) As Long
            'ReDim ilAVgRtg(1 To imWeeksPerMonth(ilLoopOnMonth)) As Integer
            
            ReDim llPopEst(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Long       'dimension for the number of weeks in the month (4 or 5)
            ReDim llSpots(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Long
            ReDim llRate(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Long
            ReDim llAvgAud(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Long
            ReDim llGrImp(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Long
            ReDim llGRP(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Long
            ReDim ilAVgRtg(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Integer
            'need to make arrays exact size for the audience rtn
            For ilWeeks = 1 To imWeeksPerMonth(ilLoopOnMonth)
                llPopEst(ilWeeks - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lPopEst(ilWeekInx + ilWeeks - 1)
                llSpots(ilWeeks - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lSpots(ilWeekInx + ilWeeks - 1)
                llRate(ilWeeks - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lRate(ilWeekInx + ilWeeks - 1)
                llAvgAud(ilWeeks - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lAvgAud(ilWeekInx + ilWeeks - 1)
                llMonthSpotCount = llMonthSpotCount + llSpots(ilWeeks - 1)
            Next ilWeeks
            ilWeekInx = ilWeekInx + imWeeksPerMonth(ilLoopOnMonth)     'keep running total of the number weeks in each month
            
            '10-30-14 default to use 1 place rating regardless of agency flag
            'gAvgAudToLnResearch "1", True, llPop, llPopEst(), llSpots(), llRate(), llAvgAud(), llLineCost, llLineAvgAud, ilAVgRtg(), ilLineAvgRtg, llGrImp(), llLineGrimp, llGRP(), llLineGRP, llLineCPP, llLineCPM, llLinePopEst, flLineCost
            gAvgAudToLnResearch "1", True, llPop, llPopEst(), llSpots(), llRate(), llAvgAud(), dlLineCost, llLineAvgAud, ilAVgRtg(), ilLineAvgRtg, llGrImp(), llLineGrimp, llGRP(), llLineGRP, llLineCPP, llLineCPM, llLinePopEst, flLineCost 'TTP 10439 - Rerate 21,000,000

            'save monthly results for contract totals later
'                ilUpperMonthInfo = UBound(tmMonthInfoByCnt)
'                tmMonthInfoByCnt(ilUpperMonthInfo).iVefCode = tmWklyInfoByDP(ilLoopOnInfo).iVefCode
'                tmMonthInfoByCnt(ilUpperMonthInfo).iRdfcode = tmWklyInfoByDP(ilLoopOnInfo).iRdfcode
'
'                tmMonthInfoByCnt(ilUpperMonthInfo).lOVStartTime = tmWklyInfoByDP(ilLoopOnInfo).lOVStartTime
'                tmMonthInfoByCnt(ilUpperMonthInfo).lOVEndTime = tmWklyInfoByDP(ilLoopOnInfo).lOVEndTime
'                For ilDay = 0 To 6
'                    tmMonthInfoByCnt(ilUpperMonthInfo).iDays(ilDay) = tmWklyInfoByDP(ilLoopOnInfo).iDays(ilDay)
'                Next ilDay
'
            ilLoopOnMonthMinusOne = ilLoopOnMonth - 1
            tmMonthInfoByCnt(ilUpperMonthInfo).lAvgAud(ilLoopOnMonthMinusOne) = llLineAvgAud
            tmMonthInfoByCnt(ilUpperMonthInfo).lPopEst(ilLoopOnMonthMinusOne) = llLinePopEst
            'tmMonthInfoByCnt(ilUpperMonthInfo).lTotalCost(ilLoopOnMonthMinusOne) = llLineCost
            'tmMonthInfoByCnt(ilUpperMonthInfo).fTotalCost(ilLoopOnMonthMinusOne) = flLineCost
            tmMonthInfoByCnt(ilUpperMonthInfo).dTotalCost(ilLoopOnMonthMinusOne) = dlLineCost 'TTP 10439 - Rerate 21,000,000
            tmMonthInfoByCnt(ilUpperMonthInfo).iAvgRtg(ilLoopOnMonthMinusOne) = ilLineAvgRtg
            tmMonthInfoByCnt(ilUpperMonthInfo).lCPP(ilLoopOnMonthMinusOne) = llLineCPP
            tmMonthInfoByCnt(ilUpperMonthInfo).lCPM(ilLoopOnMonthMinusOne) = llLineCPM
            tmMonthInfoByCnt(ilUpperMonthInfo).lGrImp(ilLoopOnMonthMinusOne) = llLineGrimp
            tmMonthInfoByCnt(ilUpperMonthInfo).lGRP(ilLoopOnMonthMinusOne) = llLineGRP
            tmMonthInfoByCnt(ilUpperMonthInfo).lSpots(ilLoopOnMonthMinusOne) = llMonthSpotCount
        Next ilLoopOnMonth
        
        'Get total for the year for each line (13th month)
        ilWeeks = 0
        For ilLoopOnMonth = 1 To imPeriods
            ilWeeks = ilWeeks + imWeeksPerMonth(ilLoopOnMonth)      'determine total number of weeks for the months requested; only process that many weeks
        Next ilLoopOnMonth
        For ilLoopOnMonth = 1 To ilWeeks
            llMonthSpotCount = 0
            'ReDim llPopEst(1 To ilWeeks) As Long         'dimension for the number of weeks in the month (4 or 5)
            'ReDim llSpots(1 To ilWeeks) As Integer
            'ReDim llrate(1 To ilWeeks) As Long
            'ReDim llAvgAud(1 To ilWeeks) As Long
            'ReDim llGrImp(1 To ilWeeks) As Long
            'ReDim llGRP(1 To ilWeeks) As Long
            'ReDim ilAVgRtg(1 To ilWeeks) As Integer
            
            ReDim llPopEst(0 To ilWeeks - 1) As Long       'dimension for the number of weeks in the month (4 or 5)
            ReDim llSpots(0 To ilWeeks - 1) As Long
            ReDim llRate(0 To ilWeeks - 1) As Long
            ReDim llAvgAud(0 To ilWeeks - 1) As Long
            ReDim llGrImp(0 To ilWeeks - 1) As Long
            ReDim llGRP(0 To ilWeeks - 1) As Long
            ReDim ilAVgRtg(0 To ilWeeks - 1) As Integer
            'need to make arrays exact size for the audience rtn
            For ilWeekInx = 1 To ilWeeks
                'llPopEst(ilWeekInx) = tmWklyInfoByDP(ilWklyInfoByDP).lPopEst(ilWeekInx)
                'llSpots(ilWeekInx) = tmWklyInfoByDP(ilWklyInfoByDP).lSpots(ilWeekInx)
                'llrate(ilWeekInx) = tmWklyInfoByDP(ilWklyInfoByDP).lRate(ilWeekInx)
                'llAvgAud(ilWeekInx) = tmWklyInfoByDP(ilWklyInfoByDP).lAvgAud(ilWeekInx)
                'llMonthSpotCount = llMonthSpotCount + llSpots(ilWeekInx)
                llPopEst(ilWeekInx - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lPopEst(ilWeekInx - 1)
                llSpots(ilWeekInx - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lSpots(ilWeekInx - 1)
                llRate(ilWeekInx - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lRate(ilWeekInx - 1)
                llAvgAud(ilWeekInx - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lAvgAud(ilWeekInx - 1)
                llMonthSpotCount = llMonthSpotCount + llSpots(ilWeekInx - 1)
            Next ilWeekInx
        Next ilLoopOnMonth
        
        '10-30-14 default to use 1 place rating regardless of agency flag
        'gAvgAudToLnResearch "1", True, llPop, llPopEst(), llSpots(), llRate(), llAvgAud(), llLineCost, llLineAvgAud, ilAVgRtg(), ilLineAvgRtg, llGrImp(), llLineGrimp, llGRP(), llLineGRP, llLineCPP, llLineCPM, llLinePopEst, flLineCost
        gAvgAudToLnResearch "1", True, llPop, llPopEst(), llSpots(), llRate(), llAvgAud(), dlLineCost, llLineAvgAud, ilAVgRtg(), ilLineAvgRtg, llGrImp(), llLineGrimp, llGRP(), llLineGRP, llLineCPP, llLineCPM, llLinePopEst, flLineCost 'TTP 10439 - Rerate 21,000,000
        'year total for the line
        'tmMonthInfoByCnt(ilUpperMonthInfo).lAvgAud(13) = llLineAvgAud
        'tmMonthInfoByCnt(ilUpperMonthInfo).lPopEst(13) = llLinePopEst
        'tmMonthInfoByCnt(ilUpperMonthInfo).lTotalCost(13) = llLineCost
        'tmMonthInfoByCnt(ilUpperMonthInfo).iAvgRtg(13) = ilLineAvgRtg
        'tmMonthInfoByCnt(ilUpperMonthInfo).lCPP(13) = llLineCPP
        'tmMonthInfoByCnt(ilUpperMonthInfo).lCPM(13) = llLineCPM
        'tmMonthInfoByCnt(ilUpperMonthInfo).lGrImp(13) = llLineGrimp
        'tmMonthInfoByCnt(ilUpperMonthInfo).lGRP(13) = llLineGRP
        'tmMonthInfoByCnt(ilUpperMonthInfo).lSpots(13) = llMonthSpotCount
        tmMonthInfoByCnt(ilUpperMonthInfo).lAvgAud(12) = llLineAvgAud
        tmMonthInfoByCnt(ilUpperMonthInfo).lPopEst(12) = llLinePopEst
        'tmMonthInfoByCnt(ilUpperMonthInfo).lTotalCost(12) = llLineCost
        'tmMonthInfoByCnt(ilUpperMonthInfo).fTotalCost(12) = flLineCost
        tmMonthInfoByCnt(ilUpperMonthInfo).dTotalCost(12) = dlLineCost 'TTP 10439 - Rerate 21,000,000
        tmMonthInfoByCnt(ilUpperMonthInfo).iAvgRtg(12) = ilLineAvgRtg
        tmMonthInfoByCnt(ilUpperMonthInfo).lCPP(12) = llLineCPP
        tmMonthInfoByCnt(ilUpperMonthInfo).lCPM(12) = llLineCPM
        tmMonthInfoByCnt(ilUpperMonthInfo).lGrImp(12) = llLineGrimp
        tmMonthInfoByCnt(ilUpperMonthInfo).lGRP(12) = llLineGRP
        tmMonthInfoByCnt(ilUpperMonthInfo).lSpots(12) = llMonthSpotCount
        
        ReDim Preserve tmMonthInfoByCnt(0 To ilUpperMonthInfo + 1) As MONTHINFOBYVEHICLE
    Next ilWklyInfoByDP
    Erase llPopEst, llSpots, llRate, llAvgAud, llGrImp, llGRP, ilAVgRtg
    Exit Sub
End Sub

'
'           mGetCntMonthInfo - get all line totals for 1month at a time to obtain monthly research results
'           Roll over the Monthly totals for all lines to get the contract total for the month and year
'
Public Sub mGetCntMonthInfo(llResearchPop As Long)
    Dim ilMonthInfoByCnt As Integer
    Dim ilUpperMonthInfo As Integer
    Dim ilLoopOnMonth As Integer
    Dim ilLoopOnYear As Integer
    Dim ilDay As Integer
    Dim ilStartMonth As Integer
    Dim ilEndMonth As Integer
    Dim llSpots As Long
    'output for gResearchTotals
    'Dim llMonthCost As Long
    Dim dlMonthCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim flMonthCost As Single
    Dim ilMonthAvgRtg As Integer
    Dim llMonthGrimp As Long
    Dim llMonthGRP As Long
    Dim llMonthCPP As Long
    Dim llMonthCPM As Long
    Dim llMonthAVgAud As Long
    Dim ilLoopOnMonthMinusOne As Integer

    For ilLoopOnYear = 1 To 2                               'pass 1 = process all periods requested, pass 2 = process year (13th month)
        If ilLoopOnYear = 1 Then
            ilStartMonth = 1
            ilEndMonth = imPeriods
        Else
            ilStartMonth = 13
            ilEndMonth = 13
        End If
        For ilLoopOnMonth = ilStartMonth To ilEndMonth

            llSpots = 0
            'ReDim llPopEst(1 To UBound(tmMonthInfoByCnt)) As Long
            'ReDim llCost(1 To UBound(tmMonthInfoByCnt)) As Long
            'ReDim llAvgAud(1 To UBound(tmMonthInfoByCnt)) As Long
            'ReDim llGrImp(1 To UBound(tmMonthInfoByCnt)) As Long
            'ReDim llGRP(1 To UBound(tmMonthInfoByCnt)) As Long
            'ReDim ilAVgRtg(1 To UBound(tmMonthInfoByCnt)) As Integer
            
            ReDim llPopEst(0 To UBound(tmMonthInfoByCnt) - 1) As Long
            'ReDim llCost(0 To UBound(tmMonthInfoByCnt) - 1) As Long
            'ReDim flCost(0 To UBound(tmMonthInfoByCnt) - 1) As Single
            ReDim dlCost(0 To UBound(tmMonthInfoByCnt) - 1) As Double 'TTP 10439 - Rerate 21,000,000
            ReDim llAvgAud(0 To UBound(tmMonthInfoByCnt) - 1) As Long
            ReDim llGrImp(0 To UBound(tmMonthInfoByCnt) - 1) As Long
            ReDim llGRP(0 To UBound(tmMonthInfoByCnt) - 1) As Long
            ReDim ilAVgRtg(0 To UBound(tmMonthInfoByCnt) - 1) As Integer
            
            ilLoopOnMonthMinusOne = ilLoopOnMonth - 1
            'gather all lines monthly results and create array for the total contract monthly total
            For ilMonthInfoByCnt = LBound(tmMonthInfoByCnt) To UBound(tmMonthInfoByCnt) - 1     'single contracts information by vehicle, dp & week
                'llCost(ilMonthInfoByCnt + 1) = tmMonthInfoByCnt(ilMonthInfoByCnt).lTotalCost(ilLoopOnMonthMinusOne)
                'llGrImp(ilMonthInfoByCnt + 1) = tmMonthInfoByCnt(ilMonthInfoByCnt).lGrImp(ilLoopOnMonthMinusOne)
                'llGRP(ilMonthInfoByCnt + 1) = tmMonthInfoByCnt(ilMonthInfoByCnt).lGRP(ilLoopOnMonthMinusOne)
                'ilAVgRtg(ilMonthInfoByCnt + 1) = tmMonthInfoByCnt(ilMonthInfoByCnt).iAvgRtg(ilLoopOnMonthMinusOne)
                'llSpots = llSpots + tmMonthInfoByCnt(ilMonthInfoByCnt + 1).lSpots(ilLoopOnMonthMinusOne)
            
                'llCost(ilMonthInfoByCnt) = tmMonthInfoByCnt(ilMonthInfoByCnt).lTotalCost(ilLoopOnMonthMinusOne)
                'flCost(ilMonthInfoByCnt) = tmMonthInfoByCnt(ilMonthInfoByCnt).fTotalCost(ilLoopOnMonthMinusOne)
                dlCost(ilMonthInfoByCnt) = tmMonthInfoByCnt(ilMonthInfoByCnt).dTotalCost(ilLoopOnMonthMinusOne) 'TTP 10439 - Rerate 21,000,000
                llGrImp(ilMonthInfoByCnt) = tmMonthInfoByCnt(ilMonthInfoByCnt).lGrImp(ilLoopOnMonthMinusOne)
                llGRP(ilMonthInfoByCnt) = tmMonthInfoByCnt(ilMonthInfoByCnt).lGRP(ilLoopOnMonthMinusOne)
                ilAVgRtg(ilMonthInfoByCnt) = tmMonthInfoByCnt(ilMonthInfoByCnt).iAvgRtg(ilLoopOnMonthMinusOne)
                llSpots = llSpots + tmMonthInfoByCnt(ilMonthInfoByCnt + 1).lSpots(ilLoopOnMonthMinusOne)
            Next ilMonthInfoByCnt
            
            '10-30-14 default to use 1 place rating regardless of agency flag
            'gResearchTotalsFloat "1", True, llResearchPop, flCost(), llGrImp(), llGRP(), llSpots, llMonthCost, ilMonthAvgRtg, llMonthGrimp, llMonthGRP, llMonthCPP, llMonthCPM, llMonthAVgAud, flMonthCost
            gResearchTotalsFloat "1", True, llResearchPop, dlCost(), llGrImp(), llGRP(), llSpots, dlMonthCost, ilMonthAvgRtg, llMonthGrimp, llMonthGRP, llMonthCPP, llMonthCPM, llMonthAVgAud, flMonthCost 'TTP 10439 - Rerate 21,000,000
            'save monthly results for contract totals later
            ilUpperMonthInfo = UBound(tmMonthInfoFinals)
            tmMonthInfoFinals(ilUpperMonthInfo).lCntrNo = tgChfCT.lCntrNo
            tmMonthInfoFinals(ilUpperMonthInfo).iAdfCode = tgChfCT.iAdfCode
            tmMonthInfoFinals(ilUpperMonthInfo).imnfDemoCode = imMnfDemoCode
            tmMonthInfoFinals(ilUpperMonthInfo).bmMissingBookCode = bmMissingDnfCode            '1-21-20

            tmMonthInfoFinals(ilUpperMonthInfo).lAvgAud(ilLoopOnMonthMinusOne) = llMonthAVgAud
            'tmMonthInfoFinals(ilUpperMonthInfo).lTotalCost(ilLoopOnMonthMinusOne) = llMonthCost
            tmMonthInfoFinals(ilUpperMonthInfo).dTotalCost(ilLoopOnMonthMinusOne) = dlMonthCost 'TTP 10439 - Rerate 21,000,000
            tmMonthInfoFinals(ilUpperMonthInfo).iAvgRtg(ilLoopOnMonthMinusOne) = ilMonthAvgRtg
            tmMonthInfoFinals(ilUpperMonthInfo).lCPP(ilLoopOnMonthMinusOne) = llMonthCPP
            tmMonthInfoFinals(ilUpperMonthInfo).lCPM(ilLoopOnMonthMinusOne) = llMonthCPM
            tmMonthInfoFinals(ilUpperMonthInfo).lGrImp(ilLoopOnMonthMinusOne) = llMonthGrimp
            tmMonthInfoFinals(ilUpperMonthInfo).lGRP(ilLoopOnMonthMinusOne) = llMonthGRP
        Next ilLoopOnMonth
    Next ilLoopOnYear
    'ReDim Preserve tmMonthInfoFinals(1 To ilUpperMonthInfo + 1) As MONTHINFOBYVEHICLE
    ReDim Preserve tmMonthInfoFinals(0 To ilUpperMonthInfo + 1) As MONTHINFOBYVEHICLE
    Exit Sub
End Sub

'
'           mInsertCntFinalInfo - loop thru the Final contract array (tmMonthInfoFinals) and write records out for crystal reporting
'           <input> slRecordType flag for Vehicle option: D= detail vehicle/dp record, V = vehicle total record
'                                For Advt option, unused
Public Sub mInsertCntFinalInfo(Optional slRecordType As String = " ")
    Dim llLoopOnFinals As Long
    Dim ilLoopOnMonth As Integer
    Dim ilRet As Integer
    Dim ilTemp As Integer
    
    tmGrf.lGenTime = lgNowTime
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iPerGenl(2) = imPeriods
    'tmGrf.iPerGenl(3) = imMajorSet          'for the vehicle option only.  if no vehicle group defined for vehicle, shown Unknown if a group has been selected
    tmGrf.iPerGenl(1) = imPeriods
    tmGrf.iPerGenl(2) = imMajorSet          'for the vehicle option only.  if no vehicle group defined for vehicle, shown Unknown if a group has been selected
    
    For llLoopOnFinals = LBound(tmMonthInfoFinals) To UBound(tmMonthInfoFinals) - 1
        'If tmMonthInfoFinals(llLoopOnFinals).lCPM(13) <> 0 Or tmMonthInfoFinals(llLoopOnFinals).lCPP(13) <> 0 Then  'ignore if year total is zero
            tmGrf.sBktType = slRecordType                            'this flag is for the vehicle option only, which indicates the Detail records (vehicle/dp totals)
                                                            'vs a Vehicle total record ("V")
            tmGrf.iAdfCode = tmMonthInfoFinals(llLoopOnFinals).iAdfCode
            tmGrf.lChfCode = tmMonthInfoFinals(llLoopOnFinals).lCntrNo
            tmGrf.iRdfCode = tmMonthInfoFinals(llLoopOnFinals).iRdfCode
            tmGrf.iSofCode = tmMonthInfoFinals(llLoopOnFinals).imnfDemoCode
            tmGrf.iVefCode = tmMonthInfoFinals(llLoopOnFinals).iVefCode
            'find the vehicle group
            tmGrf.iCode2 = 0
            gGetVehGrpSets tmGrf.iVefCode, 0, imMajorSet, ilTemp, tmGrf.iCode2

            'override start/end times if applicable
            gPackTimeLong tmMonthInfoFinals(llLoopOnFinals).lOVStartTime, tmGrf.iMissedTime(0), tmGrf.iMissedTime(1)
            gPackTimeLong tmMonthInfoFinals(llLoopOnFinals).lOVEndTime, tmGrf.iTime(0), tmGrf.iTime(1)
            'Parse day string into text
            
            For ilLoopOnMonth = 1 To 13
                If imCPPCPM = 0 Then
                    tmGrf.lDollars(ilLoopOnMonth - 1) = tmMonthInfoFinals(llLoopOnFinals).lCPP(ilLoopOnMonth - 1)
                Else
                    tmGrf.lDollars(ilLoopOnMonth - 1) = tmMonthInfoFinals(llLoopOnFinals).lCPM(ilLoopOnMonth - 1)
                End If
            Next ilLoopOnMonth
            tmGrf.lCode4 = 0                '1-10-20 flag to indicate missing dnf code
            If Trim$(slRecordType) = "" Then            'its total by advt
                If tmMonthInfoFinals(llLoopOnFinals).bmMissingBookCode Then
                    tmGrf.lCode4 = 1
                End If
            Else                                        'totals by vehicle
                If bmMissingDnfCode Then
                    tmGrf.lCode4 = 1
                End If
            End If
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        'End If
    Next llLoopOnFinals
    Erase tmMonthInfoFinals
End Sub

'********************************************************************************************
'
'              gCrCP30UnitAdv - Prepass for CPP/CPM 30"Unit Report
'
'
'       Produce prepass for cpp/cpm for 30" units by  Advertiser.  The generated
'       data will be obtained from contracts, using the contract's primary demo, or a selected demo.
'       Up to 12 monthly values are obtained for std months (corp & Calendar not implemented).
'       Calc can be based on schedule line book default book.   book closest to spot air date/date
'       (this option not implemented due to retrieving from contracts, not spots).
'       For spots that are not 30 in length, user makes the rules to calculate the value of
'       spots not divisible by 30".  for example, a 15" will be 1/2 of a 30.
'
'********************************************************************************************
Public Sub gCrCP30UnitAdv()
    ReDim ilNowTime(0 To 1) As Integer    'end time of run
    Dim slStr As String
    Dim ilRet As Integer
    Dim llEarliestStart As Long     'earliest start date from all contracts to process
    Dim llLatestEnd As Long         'latest end date from all contracts to process
    Dim slEarliestStart As String   'earliest start date from all contracts to process
    Dim slLatestEnd As String       'latest end date from all contrcts to process
    Dim ilVehicle As Integer        'loop to process spots: gather by one vehicle at a time
    Dim ilVefCode As Integer
    Dim ilVefIndex As Integer
    Dim ilOk As Integer
    Dim blFound As Boolean
    Dim ilDay As Integer
    Dim llTime As Long              'avail time
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim llPop As Long               'pop of demo
    Dim llAvgAud As Long            'aud for demo
    Dim illoop As Integer
    Dim ilDnfCode As Integer        'book to use for spot obtained
    Dim llDate As Long
    Dim ilClf As Integer            'sched line processing loop
    
    Dim llLoop As Long               '1-15-08
    Dim ilOpenError As Integer
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim slPopLineType As String * 1
    Dim ilPopRdfCode As Integer
    Dim llPopRafCode As Long
    Dim llPopEst As Long
    ReDim llkeyarray(0 To 0) As Long
    Dim ilLoopOnKey As Integer
    Dim llKeyCode As Long
    Dim blValidCType As Boolean
    Dim ilUpperClf As Integer
    Dim llResearchPop As Long
    'ReDim llPopByLine(1 To 1) As Long
    ReDim llPopByLine(0 To 0) As Long
    'ReDim ilVehList(1 To 1) As Integer        'list of unique vehicles
    ReDim ilVehList(0 To 0) As Integer        'list of unique vehicles
    Dim ilCff As Integer
    Dim ilfirstTime As Integer
    Dim ilSocEcoMnfCode As Integer
    Dim ilInputDays(0 To 6) As Integer          'valid days of the week (true/false), if daily, # spots/day
    Dim ilDemoAvgAudDays(0 To 6) As Integer     'valid days of the week (true/false)
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim ilSpots As Integer
    Dim llDate2 As Long
    Dim ilWeekInx As Integer
    Dim ilLoopOnInfo As Integer
    Dim blFoundWklyInfoDP As Boolean
    Dim ilUpperWklyInfoByDP As Integer
    Dim blAtLeast1FlightFound As Boolean
    Dim llLineStartDate As Long
    Dim llLineEndDate As Long
    Dim ilHowManyUnits As Integer
    Dim ilTemp As Integer
    Dim ilTemp2 As Integer
    Dim blValidSpotType As Boolean
        

    ilOpenError = mOpen30Unit()           'open applicable files
    If ilOpenError Then
        Exit Sub            'at least 1 open error
    End If
    
    mObtainSelectivity
    
    llEarliestStart = lmStartDates(1)
    llLatestEnd = lmStartDates(imPeriods + 1) - 1
    slEarliestStart = Format$(llEarliestStart, "m/d/yy")
    slLatestEnd = Format$(llLatestEnd, "m/d/yy")
        

    For illoop = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1
        'filter out the advertisers  or contract types not requested
        If gFilterLists(tmChfAdvtExt(illoop).iAdfCode, imInclAdvtCodes, imUseAdvtCodes()) Then
            'filtering of contract header types should have been filtered at the time of obtains the active contracts
            blValidCType = gFilterContractType(tmChf, tmCntTypes, False)         'exclude proposal type checks
            If blValidCType Then            'if valid, build into array; otherwise bypass
                llkeyarray(UBound(llkeyarray)) = tmChfAdvtExt(illoop).lCode
                ReDim Preserve llkeyarray(0 To UBound(llkeyarray) + 1) As Long
            End If
        End If
    Next illoop
    ilSocEcoMnfCode = 0         'not using socio-economic codes for research
    
    'ReDim tmMonthInfoFinals(1 To 1) As MONTHINFOBYVEHICLE        'required if option by vehicle
    ReDim tmMonthInfoFinals(0 To 0) As MONTHINFOBYVEHICLE        'required if option by vehicle
    
    For ilLoopOnKey = 0 To UBound(llkeyarray) - 1           'looping on chfcode or vehicle code
        llKeyCode = llkeyarray(ilLoopOnKey)
        ilUpperClf = 0                                  'no lines in research table
        llResearchPop = -1                       'pop from book if all same books across lines, else its zero

        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llKeyCode, False, tgChfCT, tgClfCT(), tgCffCT(), False)

        ReDim tmMonthInfoByCnt(0 To 0) As MONTHINFOBYVEHICLE
        ilUpperClf = UBound(tgClfCT)
        
        bmMissingDnfCode = False
        'If UBound(tgClfCT) > UBound(llPopByLine) Then                   'if no sch lines, results in error to redim 0
        If UBound(tgClfCT) - 1 > UBound(llPopByLine) Then                 'if no sch lines, results in error to redim 0
            'ReDim llPopByLine(1 To UBound(tgClfCT)) As Long
            ReDim llPopByLine(0 To UBound(tgClfCT) - 1) As Long
        End If

        'ReDim lmPop(1 To 1) As Long                              'population bh vehicle
        ReDim lmPop(0 To 0) As Long                              'population bh vehicle
        'ReDim ilVehList(1 To 1) As Integer              'vehicles in this contract
        ReDim ilVehList(0 To 0) As Integer              'vehicles in this contract
        blAtLeast1FlightFound = False

        For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
            tmClf = tgClfCT(ilClf).ClfRec
            ilVefIndex = gBinarySearchVef(tmClf.iVefCode)
            
            'filter vehicle selectivity
            If Not gFilterLists(tmClf.iVefCode, imInclVefCodes, imUsevefcodes()) Then
                ilVefIndex = -1         'not a selected vehicle, bypass
            Else                        'valid vehicle, is the vehicle group OK
                'Setup the major sort factor
                gGetVehGrpSets tmClf.iVefCode, 0, imMajorSet, ilTemp, ilTemp2
                'check selectivity of vehicle groups
                If (imMajorSet > 0) Then
                    If Not gFilterLists(ilTemp2, imInclVGCodes, imUseVGCodes()) Then
                        ilVefIndex = -1
                    End If
                End If
            End If
             
            gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), slStr
            llLineStartDate = gDateValue(slStr)
            gUnpackDate tmClf.iEndDate(0), tmClf.iEndDate(1), slStr
            llLineEndDate = gDateValue(slStr)
                                    
            If (tmClf.sType <> "O" And tmClf.sType <> "A" And tmClf.sType <> "E") And ilVefIndex >= 0 And llLineEndDate > llLineStartDate Then         'ignore package lines and CBS lines
                ilHowManyUnits = gDetermineSpotLenRatio(tmClf.iLen, tmSpotLenRatio)     'determine 30" unit of this spot length by the user defined table
                ilHowManyUnits = ilHowManyUnits / 10                                    'do not carry to hundreds
                ilDnfCode = tmClf.iDnfCode      'assume using sch line book
                If imBook = 1 Then              'use vehicle default book
                    ilDnfCode = tgMVef(ilVefIndex).iDnfCode
                End If
                
                If ilDnfCode = 0 Then                               '1-21-20  if missing book on contract, flag it on report
                    bmMissingDnfCode = True
                End If
                
                blFound = False
                For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1 Step 1
                    If tmClf.iVefCode = ilVehList(ilVehicle) Then
                        blFound = True
                        Exit For
                    End If
                Next ilVehicle
                If Not blFound Then
                    ilVehList(UBound(ilVehList)) = tmClf.iVefCode
                    ilVehicle = UBound(ilVehList)
                    'ReDim Preserve ilVehList(1 To UBound(ilVehList) + 1)
                    ReDim Preserve ilVehList(0 To UBound(ilVehList) + 1)
                    'ReDim Preserve lmPop(1 To UBound(lmPop) + 1)
                    ReDim Preserve lmPop(0 To UBound(lmPop) + 1)
                End If

                'Build population table by vehicle
                If imDemo < 0 Then              '-1 indicates to use primary demo from header
                    imMnfDemoCode = tgChfCT.iMnfDemo(0)
                Else
                    imMnfDemoCode = imDemo
                End If
                ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, ilDnfCode, ilSocEcoMnfCode, imMnfDemoCode, llPop)
'                    If llPop > 0 Then       '9-18-15 bypass any lines without a population
                    llOvStartTime = 0           'assume no override times
                    llOvEndTime = 0
                
                    'If smDP = "O" Then          'always use the override when doing the research. the orig question was how to show on reports, which has been taken out.
                        If tmClf.iStartTime(0) = 1 And tmClf.iStartTime(1) = 0 Then
                            llOvStartTime = 0
                            llOvEndTime = 0
                        Else
                            'override times exist
                            gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llOvStartTime
                            gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llOvEndTime
                        End If
                    'End If

                    If tgSpf.sDemoEstAllowed <> "Y" Then
                        'llPopByLine(ilClf + 1) = llPop    '11-23-99 save the pop by line (each linemay have different survey books)
                        llPopByLine(ilClf) = llPop    '11-23-99 save the pop by line (each linemay have different survey books)
                        If lmPop(ilVehicle) = 0 Then            'same vehicle found more than once, if pop already stored,
                                                                'dont wipe out with a possible non-population value
                            lmPop(ilVehicle) = llPop            'associate the population with the vehicle
                        End If
                        If llResearchPop = -1 And llPop <> 0 Then          'first time, llResearchPop is for the summary records (-1 first time thru, 0 = different books across vehicles)
                            llResearchPop = llPop
                        Else
                            If (llResearchPop <> 0) And (llResearchPop <> llPop) And (llPop <> 0) Then      'test to see if this pop is different that the prev one.
                                llResearchPop = 0                                           'if different pops, calculate the contract  summary different
                                If llPop <> lmPop(ilVehicle) Then
                                    lmPop(ilVehicle) = -1
                                End If
                            Else
                                'if current line has population, but there was already a different across
                                'lines in pop, dont save new one
                                If llPop <> 0 And (llResearchPop <> 0 And llResearchPop <> -1) Then
                                    llResearchPop = llPop
                                Else
                                    If lmPop(ilVehicle) <> llPop And lmPop(ilVehicle) <> -1 Then
                                        lmPop(ilVehicle) = -1
                                    End If
                                End If
                            End If
                        End If
                       
                    End If
                    
                    ReDim tmWklyInfoByDP(0 To 0) As WKLYINFOBYDP        'aud & pop estimates by unique vehicle, DP (and overrides if applicable)
                    ilUpperWklyInfoByDP = 0
                    
                    ilCff = tgClfCT(ilClf).iFirstCff
                    Do While ilCff <> -1
                        tmCff = tgCffCT(ilCff).CffRec
                        blValidSpotType = mFilterSpotType()
                        If blValidSpotType Then                             'if valid spot type continue; otherwise ignore this flight

                            If tgMVef(ilVefIndex).sType = "G" Then          'sports vehicle
                                tmCff.sDyWk = "W"
                            End If
                            
    
                            ilfirstTime = True                  'set to calc avg aud one time only for this flight
                            For illoop = 0 To 6                 'init all days to not airing, setup for research results later
                                ilInputDays(illoop) = False
                                ilDemoAvgAudDays(illoop) = False        'initalize to 0
                            Next illoop
    
                            gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
                            llFltStart = gDateValue(slStr)
                            'backup start date to Monday
                            illoop = gWeekDayLong(llFltStart)
                            Do While illoop <> 0
                                llFltStart = llFltStart - 1
                                illoop = gWeekDayLong(llFltStart)
                            Loop
                            gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
                            llFltEnd = gDateValue(slStr)
                            
                            'process the flight weeks only within the requested report period
                            If llFltStart < llEarliestStart Then        'flight starts earlier than requested period, use requested period start date as starting point
                                llFltStart = llEarliestStart
                            End If
                            If llFltEnd > llLatestEnd Then              'flight end date extends past requested period, use requested period end date as ending point
                                llFltEnd = llLatestEnd
                            End If
                            '
                            'Loop thru the flight by week and build the number of spots for each week
                            For llDate2 = llFltStart To llFltEnd Step 7
                                If ilfirstTime Then                 'only need to determine valid days & # spots once per flight entry
                                    If tmCff.sDyWk = "W" Then            'weekly
                                        ilSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                                        For ilDay = 0 To 6 Step 1
                                            If (llDate2 + ilDay >= llFltStart) And (llDate2 + ilDay <= llFltEnd) Then
                                                If tmCff.iDay(ilDay) > 0 Or tmCff.sXDay(ilDay) = "1" Then
                                                    ilInputDays(ilDay) = True
                                                    ilDemoAvgAudDays(ilDay) = True      '11-19-02 for weekly, each day is indicated by true false as a valid airing day of week
                                                End If
                                            End If
                                        Next ilDay
                                    Else                                        'daily
                                        If illoop + 6 < llFltEnd Then           'we have a whole week
                                            ilSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)    'got entire week
                                            For ilDay = 0 To 6 Step 1
                                                If tmCff.iDay(ilDay) > 0 Then
                                                    ilInputDays(ilDay) = tmCff.iDay(ilDay)
                                                    ilDemoAvgAudDays(ilDay) = True      ' for daily, each day is indicated by # spots per day as a valid airing day
                                                End If
                                            Next ilDay
                                        Else                                    'do partial week
                                            For llDate = llDate2 To llFltEnd Step 1
                                                ilDay = gWeekDayLong(llDate)
                                                ilSpots = ilSpots + tmCff.iDay(ilDay)
                                                If tmCff.iDay(ilDay) > 0 Then
                                                    ilInputDays(ilDay) = tmCff.iDay(ilDay)
                                                    ilDemoAvgAudDays(ilDay) = True      'for daily, each day is indicated by # spots per day as a valid airing day
                                                End If
                                            Next llDate
                                        End If
                                    End If
                                End If
    
                                ilWeekInx = (llDate2 - llEarliestStart) / 7 + 1
                                If ilWeekInx > 0 And ilWeekInx < 54 Then           ' has to be a valid week within requested period
     
                                    If ilfirstTime Then
                                        If tgSpf.sDemoEstAllowed <> "Y" Then
                                            ilfirstTime = False
                                        End If
                                        'Daily and weekly need the valid airing day, not the spots per day if daily (ilDemoAvgAudDays)
                                        ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, tmClf.iVefCode, ilSocEcoMnfCode, imMnfDemoCode, llDate2, llDate2, tmClf.iRdfCode, llOvStartTime, llOvEndTime, ilDemoAvgAudDays(), tmClf.sType, tmClf.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                        'returned llAvgAud and llPopEst
                                    End If
                                    blAtLeast1FlightFound = True            'flag at least one valid flight found for the contract
                                    'Build array of the weekly spots & rates for one line at a time
                                    mGetWeekInfoByDP ilWeekInx, llOvStartTime, llOvEndTime, ilDemoAvgAudDays(), llAvgAud, llPopEst, (ilSpots * ilHowManyUnits)
    
                                End If
                                
                            Next llDate2
                        End If
                        ilCff = tgCffCT(ilCff).iNextCff               'get next flight record from mem
                    Loop                                            'while ilcff <> -1
                    'at the end of each line, get the monthly values of audience, rating, grimps, cpp, cpm for this one line
                    'mGetLineMonthInfo llPopByLine(ilClf + 1)            'get line totals; build 1 entry per line, each entry having 12 months + year values
                    mGetLineMonthInfo llPopByLine(ilClf)            'get line totals; build 1 entry per line, each entry having 12 months + year values
'                    End If     'llpop > 0
            End If
         Next ilClf                 'loop on spots by contr code or spots by vehicle
         If blAtLeast1FlightFound Then
            'finished contract, now get contract totals if by advt, otherwise format the vehicle totals
            mGetCntMonthInfo llResearchPop      'get contract totals
'                If tmClf.iDnfCode = 0 Then                      '1-10-20 flag this in the report as at least one line with missing demo reference
'                    bmMissingDnfCode = True
'                End If
            'bmMissingDnfCode indicate missing book or not
        End If
    Next ilLoopOnKey                    'loop on chfcode or vehicle code
    
    mInsertCntFinalInfo


    'debugging only for time program took to run
    slStr = Format$(gNow(), "h:mm:ssAM/PM")       'end time of run
    gPackTime slStr, ilNowTime(0), ilNowTime(1)
    gUnpackTimeLong ilNowTime(0), ilNowTime(1), False, llTime
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, llPop   'start time of run
    llPop = llPop - llTime              'time in seconds in runtime
    ilRet = gSetFormula("RunTime", llPop)  'show how long report generated

    sgCntrForDateStamp = ""     'initialize contract routine next time thru
    Erase tmChfAdvtExt
    Erase tgCffCT, tgClfCT
    Erase llkeyarray, tmWklyInfoByDP, tmMonthInfoByCnt, tmMonthInfoByVehicle
    Erase lmPop
    Erase ilVehList

    mCloseFiles
    Exit Sub
mTerminate:
    On Error GoTo 0
    Exit Sub
End Sub

'
'                   mGetVehicleLineMonthInfo - Sort by Vehicle option.
'                   All weekly information has been built into Wkly arrays,  get the
'                   Line totals to retain for all vehicles
'
Public Sub mGetVehicleLineMonthInfo(llPop As Long, ilClfDnfCode As Integer)
    Dim ilWklyInfoByDP As Integer
    Dim ilUpperMonthInfo As Integer
    Dim ilLoopOnMonth As Integer
    Dim ilWeeks As Integer                  '# weeks in period (4 or 5)
    Dim ilWeekInx As Integer                'week index 1-53
        
    'line totals - output
    'Dim llLineCost As Long
    Dim dlLineCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim flLineCost As Single 'TTP 10439 - Rerate 21,000,000
    Dim llLineAvgAud As Long
    Dim ilLineAvgRtg As Integer
    Dim llLineGrimp As Long
    Dim llLineGRP As Long
    Dim llLineCPP As Long
    Dim llLineCPM As Long
    Dim llLinePopEst As Long
    'end line totals-output
    
    Dim blFoundMonthInfo As Boolean
    Dim ilLoopOnInfo As Integer
    Dim ilDay As Integer
    Dim llMonthSpotCount As Long
    Dim llLoopOnHdr As Long
    Dim blFoundHdr As Boolean
    Dim llUpperHdr As Long
    Dim llLastIndex As Long
    Dim llUpperDetail As Long
    Dim ilLoopOnMonthMinusOne As Integer

    For ilWklyInfoByDP = LBound(tmWklyInfoByDP) To UBound(tmWklyInfoByDP) - 1     'single contracts information by vehicle, dp & week
        'Determine if the vehicle has ever been processed before
        blFoundHdr = False
        For llLoopOnHdr = LBound(tmVehicle30UnitHdr) To UBound(tmVehicle30UnitHdr) - 1
            If tmVehicle30UnitHdr(llLoopOnHdr).iVefCode = tmWklyInfoByDP(ilWklyInfoByDP).iVefCode And tmVehicle30UnitHdr(llLoopOnHdr).iRdfCode = tmWklyInfoByDP(ilWklyInfoByDP).iRdfCode Then   'And tmVehicle30UnitHdr(llLoopOnHdr).lOVStartTime = tmWklyInfoByDP(ilWklyInfoByDP).lOVStartTime And tmVehicle30UnitHdr(llLoopOnHdr).lOVEndTime = tmWklyInfoByDP(ilWklyInfoByDP).lOVEndTime Then
                blFoundHdr = True
                Exit For
            End If
        Next llLoopOnHdr
        If blFoundHdr Then
            'create another detail line summary entry, obtain the last detail entry created for this vehicle and dp
            llLastIndex = tmVehicle30UnitHdr(llLoopOnHdr).lLastIndex            'determine the last index for the existing vehicle & dp summary entries so that
                                                                                'it can be changed to point to the next one in the link
            llUpperDetail = UBound(tmVehicle30UnitDetail)                       'make room for next entry
            ReDim Preserve tmVehicle30UnitDetail(LBound(tmVehicle30UnitDetail) To llUpperDetail + 1) As VEHICLE30UNITDetail 'make room for next entry
            
            'update the hdr to point to the new entry
            tmVehicle30UnitHdr(llLoopOnHdr).lLastIndex = llUpperDetail           'update header with the last entry in chain
            If tmVehicle30UnitHdr(llLoopOnHdr).lPop = -1 And llPop <> 0 Then     'if -1, the pop has never been set; but dont set if the one coming in is 0
                tmVehicle30UnitHdr(llLoopOnHdr).lPop = llPop
            End If
            If tmVehicle30UnitHdr(llLoopOnHdr).lPop <> 0 And tmVehicle30UnitHdr(llLoopOnHdr).lPop <> llPop Then
                tmVehicle30UnitHdr(llLoopOnHdr).lPop = 0
            End If
            
            tmVehicle30UnitDetail(llLastIndex).lNextIndex = llUpperDetail        'point to the next detail entry in chain
            tmVehicle30UnitDetail(llUpperDetail).lNextIndex = -1                 'current detail entry processing has no next, flag it last in the chain
        Else
            '1st time for this vehicle & dp.  Create a header entry array and the beginning of a linked list detail array of line summary entries
            llUpperHdr = UBound(tmVehicle30UnitHdr)
            tmVehicle30UnitHdr(llUpperHdr).lFirstIndex = -1
            tmVehicle30UnitHdr(llUpperHdr).lLastIndex = -1
            tmVehicle30UnitHdr(llUpperHdr).iVefCode = tmWklyInfoByDP(ilWklyInfoByDP).iVefCode
            tmVehicle30UnitHdr(llUpperHdr).iRdfCode = tmWklyInfoByDP(ilWklyInfoByDP).iRdfCode
            tmVehicle30UnitHdr(llUpperHdr).lOVStartTime = tmWklyInfoByDP(ilWklyInfoByDP).lOVStartTime
            tmVehicle30UnitHdr(llUpperHdr).lOVEndTime = tmWklyInfoByDP(ilWklyInfoByDP).lOVEndTime
            If llPop > 0 Then                                       'dont set pop if coming is a 0
                tmVehicle30UnitHdr(llUpperHdr).lPop = llPop
            Else
                tmVehicle30UnitHdr(llUpperHdr).lPop = -1
            End If
            ReDim Preserve tmVehicle30UnitHdr(LBound(tmVehicle30UnitHdr) To llUpperHdr + 1) As VEHICLE30UnitHDR
            
            llUpperDetail = UBound(tmVehicle30UnitDetail)                       'make room for next entry
            ReDim Preserve tmVehicle30UnitDetail(LBound(tmVehicle30UnitDetail) To llUpperDetail + 1) As VEHICLE30UNITDetail 'make room for next entry
            'update the hdr to point to the new entry
            tmVehicle30UnitHdr(llLoopOnHdr).lFirstIndex = llUpperDetail           'update header with the last entry in chain
            tmVehicle30UnitHdr(llLoopOnHdr).lLastIndex = llUpperDetail           'update header with the last entry in chain
            tmVehicle30UnitDetail(llUpperDetail).lNextIndex = -1        'point to the next detail entry in chain
            llLoopOnHdr = llUpperHdr
            llLastIndex = llUpperDetail
        End If
        
        ilWeekInx = 0
        
        For ilLoopOnMonth = 1 To imPeriods
            ilLoopOnMonthMinusOne = ilLoopOnMonth - 1
            llMonthSpotCount = 0
            'ReDim llPopEst(1 To imWeeksPerMonth(ilLoopOnMonth)) As Long         'dimension for the number of weeks in the month (4 or 5)
            'ReDim llSpots(1 To imWeeksPerMonth(ilLoopOnMonth)) As Integer
            'ReDim llrate(1 To imWeeksPerMonth(ilLoopOnMonth)) As Long
            'ReDim llAvgAud(1 To imWeeksPerMonth(ilLoopOnMonth)) As Long
            'ReDim llGrImp(1 To imWeeksPerMonth(ilLoopOnMonth)) As Long
            'ReDim llGRP(1 To imWeeksPerMonth(ilLoopOnMonth)) As Long
            'ReDim ilAVgRtg(1 To imWeeksPerMonth(ilLoopOnMonth)) As Integer
            
            ReDim llPopEst(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Long       'dimension for the number of weeks in the month (4 or 5)
            ReDim llSpots(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Long
            ReDim llRate(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Long
            ReDim llAvgAud(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Long
            ReDim llGrImp(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Long
            ReDim llGRP(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Long
            ReDim ilAVgRtg(0 To imWeeksPerMonth(ilLoopOnMonth) - 1) As Integer
            'need to make arrays exact size for the audience rtn
            For ilWeeks = 1 To imWeeksPerMonth(ilLoopOnMonth)
                llPopEst(ilWeeks - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lPopEst(ilWeekInx + ilWeeks - 1)
                llSpots(ilWeeks - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lSpots(ilWeekInx + ilWeeks - 1)
                llRate(ilWeeks - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lRate(ilWeekInx + ilWeeks - 1)
                llAvgAud(ilWeeks - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lAvgAud(ilWeekInx + ilWeeks - 1)
                llMonthSpotCount = llMonthSpotCount + llSpots(ilWeeks - 1)
            Next ilWeeks
            ilWeekInx = ilWeekInx + imWeeksPerMonth(ilLoopOnMonth)     'keep running total of the number weeks in each month
            
            '10-30-14 default to use 1 place rating regardless of agency flag
            'gAvgAudToLnResearch "1", True, llPop, llPopEst(), llSpots(), llRate(), llAvgAud(), llLineCost, llLineAvgAud, ilAVgRtg(), ilLineAvgRtg, llGrImp(), llLineGrimp, llGRP(), llLineGRP, llLineCPP, llLineCPM, llLinePopEst, flLineCost
            gAvgAudToLnResearch "1", True, llPop, llPopEst(), llSpots(), llRate(), llAvgAud(), dlLineCost, llLineAvgAud, ilAVgRtg(), ilLineAvgRtg, llGrImp(), llLineGrimp, llGRP(), llLineGRP, llLineCPP, llLineCPM, llLinePopEst, flLineCost 'TTP 10439 - Rerate 21,000,000
            

            tmVehicle30UnitDetail(llUpperDetail).lAvgAud(ilLoopOnMonthMinusOne) = llLineAvgAud
            tmVehicle30UnitDetail(llUpperDetail).lPopEst(ilLoopOnMonthMinusOne) = llLinePopEst
            'tmVehicle30UnitDetail(llUpperDetail).lTotalCost(ilLoopOnMonthMinusOne) = llLineCost
            'mVehicle30UnitDetail(llUpperDetail).fTotalCost(ilLoopOnMonthMinusOne) = flLineCost
            tmVehicle30UnitDetail(llUpperDetail).dTotalCost(ilLoopOnMonthMinusOne) = dlLineCost 'TTP 10439 - Rerate 21,000,000
            tmVehicle30UnitDetail(llUpperDetail).iAvgRtg(ilLoopOnMonthMinusOne) = ilLineAvgRtg
            tmVehicle30UnitDetail(llUpperDetail).lCPP(ilLoopOnMonthMinusOne) = llLineCPP
            tmVehicle30UnitDetail(llUpperDetail).lCPM(ilLoopOnMonthMinusOne) = llLineCPM
            tmVehicle30UnitDetail(llUpperDetail).lGrImp(ilLoopOnMonthMinusOne) = llLineGrimp
            tmVehicle30UnitDetail(llUpperDetail).lGRP(ilLoopOnMonthMinusOne) = llLineGRP
            tmVehicle30UnitDetail(llUpperDetail).lSpots(ilLoopOnMonthMinusOne) = llMonthSpotCount
        Next ilLoopOnMonth
        
        
        'Get total for the year for each line (13th month)
        ilWeeks = 0
        For ilLoopOnMonth = 1 To imPeriods
            ilWeeks = ilWeeks + imWeeksPerMonth(ilLoopOnMonth)      'determine total number of weeks for the months requested; only process that many weeks
        Next ilLoopOnMonth
        For ilLoopOnMonth = 1 To ilWeeks
            llMonthSpotCount = 0
            'ReDim llPopEst(1 To ilWeeks) As Long         'dimension for the number of weeks in the month (4 or 5)
            'ReDim llSpots(1 To ilWeeks) As Integer
            'ReDim llrate(1 To ilWeeks) As Long
            'ReDim llAvgAud(1 To ilWeeks) As Long
            'ReDim llGrImp(1 To ilWeeks) As Long
            'ReDim llGRP(1 To ilWeeks) As Long
            'ReDim ilAVgRtg(1 To ilWeeks) As Integer
            
            ReDim llPopEst(0 To ilWeeks - 1) As Long       'dimension for the number of weeks in the month (4 or 5)
            ReDim llSpots(0 To ilWeeks - 1) As Long
            ReDim llRate(0 To ilWeeks - 1) As Long
            ReDim llAvgAud(0 To ilWeeks - 1) As Long
            ReDim llGrImp(0 To ilWeeks - 1) As Long
            ReDim llGRP(0 To ilWeeks - 1) As Long
            ReDim ilAVgRtg(0 To ilWeeks - 1) As Integer
            'need to make arrays exact size for the audience rtn
            For ilWeekInx = 1 To ilWeeks
                'llPopEst(ilWeekInx) = tmWklyInfoByDP(ilWklyInfoByDP).lPopEst(ilWeekInx - 1)
                'llSpots(ilWeekInx) = tmWklyInfoByDP(ilWklyInfoByDP).lSpots(ilWeekInx - 1)
                'llrate(ilWeekInx) = tmWklyInfoByDP(ilWklyInfoByDP).lRate(ilWeekInx - 1)
                'llAvgAud(ilWeekInx) = tmWklyInfoByDP(ilWklyInfoByDP).lAvgAud(ilWeekInx - 1)
                'llMonthSpotCount = llMonthSpotCount + llSpots(ilWeekInx)
            
                llPopEst(ilWeekInx - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lPopEst(ilWeekInx - 1)
                llSpots(ilWeekInx - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lSpots(ilWeekInx - 1)
                llRate(ilWeekInx - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lRate(ilWeekInx - 1)
                llAvgAud(ilWeekInx - 1) = tmWklyInfoByDP(ilWklyInfoByDP).lAvgAud(ilWeekInx - 1)
                llMonthSpotCount = llMonthSpotCount + llSpots(ilWeekInx - 1)
            Next ilWeekInx
        Next ilLoopOnMonth
        
        '10-30-14 default to use 1 place rating regardless of agency flag
        'gAvgAudToLnResearch "1", True, llPop, llPopEst(), llSpots(), llRate(), llAvgAud(), llLineCost, llLineAvgAud, ilAVgRtg(), ilLineAvgRtg, llGrImp(), llLineGrimp, llGRP(), llLineGRP, llLineCPP, llLineCPM, llLinePopEst, flLineCost
        gAvgAudToLnResearch "1", True, llPop, llPopEst(), llSpots(), llRate(), llAvgAud(), dlLineCost, llLineAvgAud, ilAVgRtg(), ilLineAvgRtg, llGrImp(), llLineGrimp, llGRP(), llLineGRP, llLineCPP, llLineCPM, llLinePopEst, flLineCost 'TTP 10439 - Rerate 21,000,000
        'year total for the line
        'tmVehicle30UnitDetail(llUpperDetail).lAvgAud(13) = llLineAvgAud
        'tmVehicle30UnitDetail(llUpperDetail).lPopEst(13) = llLinePopEst
        'tmVehicle30UnitDetail(llUpperDetail).lTotalCost(13) = llLineCost
        'tmVehicle30UnitDetail(llUpperDetail).iAvgRtg(13) = ilLineAvgRtg
        'tmVehicle30UnitDetail(llUpperDetail).lCPP(13) = llLineCPP
        'tmVehicle30UnitDetail(llUpperDetail).lCPM(13) = llLineCPM
        'tmVehicle30UnitDetail(llUpperDetail).lGrImp(13) = llLineGrimp
        'tmVehicle30UnitDetail(llUpperDetail).lGRP(13) = llLineGRP
        'tmVehicle30UnitDetail(llUpperDetail).lSpots(13) = llMonthSpotCount
        tmVehicle30UnitDetail(llUpperDetail).lAvgAud(12) = llLineAvgAud
        tmVehicle30UnitDetail(llUpperDetail).lPopEst(12) = llLinePopEst
        'tmVehicle30UnitDetail(llUpperDetail).lTotalCost(12) = llLineCost
        'tmVehicle30UnitDetail(llUpperDetail).fTotalCost(12) = flLineCost
        tmVehicle30UnitDetail(llUpperDetail).dTotalCost(12) = dlLineCost 'TTP 10439 - Rerate 21,000,000
        tmVehicle30UnitDetail(llUpperDetail).iAvgRtg(12) = ilLineAvgRtg
        tmVehicle30UnitDetail(llUpperDetail).lCPP(12) = llLineCPP
        tmVehicle30UnitDetail(llUpperDetail).lCPM(12) = llLineCPM
        tmVehicle30UnitDetail(llUpperDetail).lGrImp(12) = llLineGrimp
        tmVehicle30UnitDetail(llUpperDetail).lGRP(12) = llLineGRP
        tmVehicle30UnitDetail(llUpperDetail).lSpots(12) = llMonthSpotCount
        
        If bmIncludeDPDetail Then       'show the detail by line for proofing purposes
            For ilLoopOnMonth = 1 To 13 'imPeriods
                If imCPPCPM = 0 Then
                    tmGrf.lDollars(ilLoopOnMonth - 1) = tmVehicle30UnitDetail(llUpperDetail).lCPP(ilLoopOnMonth - 1)
                Else
                    tmGrf.lDollars(ilLoopOnMonth - 1) = tmVehicle30UnitDetail(llUpperDetail).lCPM(ilLoopOnMonth - 1)
                End If
            Next ilLoopOnMonth
            tmGrf.iRdfCode = tmVehicle30UnitHdr(llLoopOnHdr).iRdfCode
            tmGrf.iAdfCode = tgChfCT.iAdfCode
            tmGrf.lChfCode = tgChfCT.lCntrNo
            'tmGrf.iPerGenl(4) = tmWklyInfoByDP(ilWklyInfoByDP).iLineNo
            tmGrf.iPerGenl(3) = tmWklyInfoByDP(ilWklyInfoByDP).iLineNo
            tmGrf.lGenTime = lgNowTime
            tmGrf.iGenDate(0) = igNowDate(0)
            tmGrf.iGenDate(1) = igNowDate(1)
            tmGrf.sBktType = "C"            'contract detail flag for crystal
            tmGrf.iVefCode = tmVehicle30UnitHdr(llLoopOnHdr).iVefCode
            tmGrf.lCode4 = 0                    '1-10-20 assume all lines have books
            If ilClfDnfCode = 0 Then
                tmGrf.lCode4 = 1                'no book line
            End If
            ilDay = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
    Next ilWklyInfoByDP
    
    Erase llPopEst, llSpots, llRate, llAvgAud, llGrImp, llGRP, ilAVgRtg
    Exit Sub
End Sub

'
'               mGetVehicleDPFinals - All lines have been created and summarized with LineResearch values.
'               Arrays exists for all unique vehicle/dp.  This points to a linked list of elements in another array (tmvehicle30unitdetail)
'               that has all its vehicle/dp line research values.  Like vehicle/dp values will be gathered by month and year,
'               to create its Research Total (gResearchTotalsFloat).
'
Public Sub mGetVehicleDPFinals()
    Dim llVehicleHdr As Long
    Dim llUpperMonthInfo As Long
    Dim ilLoopOnMonth As Integer
    Dim ilLoopOnYear As Integer
    Dim ilDay As Integer
    Dim ilStartMonth As Integer
    Dim ilEndMonth As Integer
    Dim llSpots As Long
    Dim llVehicleDetail As Long
    Dim llResearchPop As Long
    Dim llUpper As Long
    'output for gResearchTotals
    'Dim llMonthCost As Long
    Dim dlMonthCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim flMonthCost As Single
    Dim ilMonthAvgRtg As Integer
    Dim llMonthGrimp As Long
    Dim llMonthGRP As Long
    Dim llMonthCPP As Long
    Dim llMonthCPM As Long
    Dim llMonthAVgAud As Long
    Dim ilLoopOnMonthMinusOne As Integer
    'the amount of tmVehicle30UnitHdr entries is the number of vehicle & dayparts to create
    ReDim tmMonthInfoFinals(LBound(tmVehicle30UnitHdr) To UBound(tmVehicle30UnitHdr)) As MONTHINFOBYVEHICLE

    'gather all lines monthly results and create array for the total contract monthly total
    For llVehicleHdr = LBound(tmVehicle30UnitHdr) To UBound(tmVehicle30UnitHdr) - 1     'single contracts information by vehicle, dp & week
        'tmMonthInfoFinals are the vehicle/dp totals that gets written to GRF for crystal
        tmMonthInfoFinals(llVehicleHdr).iVefCode = tmVehicle30UnitHdr(llVehicleHdr).iVefCode
        tmMonthInfoFinals(llVehicleHdr).iRdfCode = tmVehicle30UnitHdr(llVehicleHdr).iRdfCode
        tmMonthInfoFinals(llVehicleHdr).lOVStartTime = tmVehicle30UnitHdr(llVehicleHdr).lOVStartTime
        tmMonthInfoFinals(llVehicleHdr).lOVEndTime = tmVehicle30UnitHdr(llVehicleHdr).lOVEndTime
        llResearchPop = tmVehicle30UnitHdr(llVehicleHdr).lPop
        llVehicleDetail = tmVehicle30UnitHdr(llVehicleHdr).lFirstIndex
        For ilLoopOnYear = 1 To 2
            If ilLoopOnYear = 1 Then
                ilStartMonth = 1
                ilEndMonth = imPeriods
            Else
                ilStartMonth = 13
                ilEndMonth = 13
            End If
            
            For ilLoopOnMonth = ilStartMonth To ilEndMonth
                ilLoopOnMonthMinusOne = ilLoopOnMonth - 1
                llSpots = 0
                'ReDim llPopEst(1 To 1) As Long
                'ReDim llCost(1 To 1) As Long
                'ReDim llAvgAud(1 To 1) As Long
                'ReDim llGrImp(1 To 1) As Long
                'ReDim llGRP(1 To 1) As Long
                'ReDim ilAVgRtg(1 To 1) As Integer
                
                ReDim llPopEst(0 To 0) As Long
                'ReDim llCost(0 To 0) As Long
                'ReDim flCost(0 To 0) As Single
                ReDim dlCost(0 To 0) As Double 'TTP 10439 - Rerate 21,000,000
                ReDim llAvgAud(0 To 0) As Long
                ReDim llGrImp(0 To 0) As Long
                ReDim llGRP(0 To 0) As Long
                ReDim ilAVgRtg(0 To 0) As Integer
                Do While llVehicleDetail >= 0
                    llUpper = UBound(llPopEst)
                    
                    llPopEst(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).lPopEst(ilLoopOnMonthMinusOne)
                    'llCost(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).lTotalCost(ilLoopOnMonthMinusOne)
                    'flCost(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).fTotalCost(ilLoopOnMonthMinusOne)
                    dlCost(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).dTotalCost(ilLoopOnMonthMinusOne) 'TTP 10439 - Rerate 21,000,000
                    llGrImp(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).lGrImp(ilLoopOnMonthMinusOne)
                    llGRP(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).lGRP(ilLoopOnMonthMinusOne)
                    ilAVgRtg(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).iAvgRtg(ilLoopOnMonthMinusOne)
                    llSpots = llSpots + tmVehicle30UnitDetail(llVehicleDetail).lSpots(ilLoopOnMonthMinusOne)
                    'ReDim Preserve llPopEst(1 To UBound(llPopEst) + 1) As Long
                    'ReDim Preserve llCost(1 To UBound(llCost) + 1) As Long
                    'ReDim Preserve llAvgAud(1 To UBound(llAvgAud) + 1) As Long
                    'ReDim Preserve llGrImp(1 To UBound(llGrImp) + 1) As Long
                    'ReDim Preserve llGRP(1 To UBound(llGRP) + 1) As Long
                    'ReDim Preserve ilAVgRtg(1 To UBound(ilAVgRtg) + 1) As Integer
                    ReDim Preserve llPopEst(0 To UBound(llPopEst) + 1) As Long
                    'ReDim Preserve llCost(0 To UBound(llCost) + 1) As Long
                    'ReDim Preserve flCost(0 To UBound(flCost) + 1) As Single
                    ReDim Preserve dlCost(0 To UBound(dlCost) + 1) As Double 'TTP 10439 - Rerate 21,000,000
                    ReDim Preserve llAvgAud(0 To UBound(llAvgAud) + 1) As Long
                    ReDim Preserve llGrImp(0 To UBound(llGrImp) + 1) As Long
                    ReDim Preserve llGRP(0 To UBound(llGRP) + 1) As Long
                    ReDim Preserve ilAVgRtg(0 To UBound(ilAVgRtg) + 1) As Integer
                
                    llVehicleDetail = tmVehicle30UnitDetail(llVehicleDetail).lNextIndex
                Loop
               
                'If UBound(llPopEst) > 1 Then
                If UBound(llPopEst) > 0 Then
                'adjust size of array, must be exact ize
                    'ReDim Preserve llPopEst(1 To UBound(llPopEst) - 1) As Long
                    'ReDim Preserve llCost(1 To UBound(llCost) - 1) As Long
                    'ReDim Preserve llAvgAud(1 To UBound(llAvgAud) - 1) As Long
                    'ReDim Preserve llGrImp(1 To UBound(llGrImp) - 1) As Long
                    'ReDim Preserve llGRP(1 To UBound(llGRP) - 1) As Long
                    'ReDim Preserve ilAVgRtg(1 To UBound(ilAVgRtg) - 1) As Integer
                   
                    ReDim Preserve llPopEst(0 To UBound(llPopEst) - 1) As Long
                    'ReDim Preserve llCost(0 To UBound(llCost) - 1) As Long
                    'ReDim Preserve flCost(0 To UBound(flCost) - 1) As Single
                    ReDim Preserve dlCost(0 To UBound(dlCost) - 1) As Double 'TTP 10439 - Rerate 21,000,000
                    ReDim Preserve llAvgAud(0 To UBound(llAvgAud) - 1) As Long
                    ReDim Preserve llGrImp(0 To UBound(llGrImp) - 1) As Long
                    ReDim Preserve llGRP(0 To UBound(llGRP) - 1) As Long
                    ReDim Preserve ilAVgRtg(0 To UBound(ilAVgRtg) - 1) As Integer
                   
                    'done reading all vehicle/dp for one month
                    '10-30-14 default to use 1 place rating regardless of agency flag
                    'gResearchTotalsFloat "1", True, llResearchPop, flCost(), llGrImp(), llGRP(), llSpots, llMonthCost, ilMonthAvgRtg, llMonthGrimp, llMonthGRP, llMonthCPP, llMonthCPM, llMonthAVgAud, flMonthCost
                    gResearchTotalsFloat "1", True, llResearchPop, dlCost(), llGrImp(), llGRP(), llSpots, dlMonthCost, ilMonthAvgRtg, llMonthGrimp, llMonthGRP, llMonthCPP, llMonthCPM, llMonthAVgAud, flMonthCost 'TTP 10439 - Rerate 21,000,000
                    'save monthly results for contract totals later
                    ilLoopOnMonthMinusOne = ilLoopOnMonth - 1
                    tmMonthInfoFinals(llVehicleHdr).lAvgAud(ilLoopOnMonthMinusOne) = llMonthAVgAud
                    'tmMonthInfoFinals(llVehicleHdr).lTotalCost(ilLoopOnMonthMinusOne) = llMonthCost
                    'tmMonthInfoFinals(llVehicleHdr).fTotalCost(ilLoopOnMonthMinusOne) = flMonthCost
                    tmMonthInfoFinals(llVehicleHdr).dTotalCost(ilLoopOnMonthMinusOne) = dlMonthCost 'TTP 10439 - Rerate 21,000,000
                    tmMonthInfoFinals(llVehicleHdr).iAvgRtg(ilLoopOnMonthMinusOne) = ilMonthAvgRtg
                    tmMonthInfoFinals(llVehicleHdr).lCPP(ilLoopOnMonthMinusOne) = llMonthCPP
                    tmMonthInfoFinals(llVehicleHdr).lCPM(ilLoopOnMonthMinusOne) = llMonthCPM
                    tmMonthInfoFinals(llVehicleHdr).lGrImp(ilLoopOnMonthMinusOne) = llMonthGrimp
                    tmMonthInfoFinals(llVehicleHdr).lGRP(ilLoopOnMonthMinusOne) = llMonthGRP
                End If
                llVehicleDetail = tmVehicle30UnitHdr(llVehicleHdr).lFirstIndex      're-establish the start of the vehicle detail dp summaries
            Next ilLoopOnMonth
        Next ilLoopOnYear
    Next llVehicleHdr
    Exit Sub
End Sub

'
'                   mFilterSpotType - filter spot type selectivity based on schedule line defined rate
'                   slPriceType -  N=No Charge; M=MG Line; B=Bonus; S=Spinoff; R=Recapturable; A=Audience Deficiency Unit (adu)
'                   <input/output> none
'                   Return = true for valid spot type to include; else false
Public Function mFilterSpotType() As Boolean
    Dim blOKtoInclude As Boolean

    blOKtoInclude = True
    If tmCff.lActPrice > 0 Then
        If Not tmCntTypes.iCharge Then
            blOKtoInclude = False
        End If
    Else
        '$0, determine which kind
        If Not tmCntTypes.iZero Then
            blOKtoInclude = False
        End If
        If (tmCff.sPriceType = "N") And Not (tmCntTypes.iNC) Then
            blOKtoInclude = False
        End If
         If (tmCff.sPriceType = "M") And Not (tmCntTypes.iMG) Then
            blOKtoInclude = False
        End If
        If (tmCff.sPriceType = "B") And Not (tmCntTypes.iBonus) Then
            blOKtoInclude = False
        End If
        If (tmCff.sPriceType = "S") And Not (tmCntTypes.iSpinoff) Then
            blOKtoInclude = False
        End If
        If (tmCff.sPriceType = "R") And Not (tmCntTypes.iRecapturable) Then
            blOKtoInclude = False
        End If
        If (tmCff.sPriceType = "A") And Not (tmCntTypes.iADU) Then
            blOKtoInclude = False
        End If
    End If
    
    mFilterSpotType = blOKtoInclude
    Exit Function
End Function

Public Sub mGetVehicleFinals()
    Dim llVehicleHdr As Long
    Dim llUpperMonthInfo As Long
    Dim ilLoopOnMonth As Integer
    Dim ilLoopOnYear As Integer
    Dim ilDay As Integer
    Dim ilStartMonth As Integer
    Dim ilEndMonth As Integer
    Dim llSpots As Long
    Dim llVehicleDetail As Long
    Dim llResearchPop As Long
    Dim llUpper As Long
    'output for gResearchTotals
    'Dim llMonthCost As Long
    Dim dlMonthCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim flMonthCost As Single
    Dim ilMonthAvgRtg As Integer
    Dim llMonthGrimp As Long
    Dim llMonthGRP As Long
    Dim llMonthCPP As Long
    Dim llMonthCPM As Long
    Dim llMonthAVgAud As Long
    Dim ilVefCode As Integer
    Dim ilVehicle As Integer
    Dim ilLoopOnMonthMinusOne As Integer

    'ReDim tmMonthInfoFinals(1 To UBound(imVehListForAll)) As MONTHINFOBYVEHICLE
    ReDim tmMonthInfoFinals(0 To UBound(imVehListForAll)) As MONTHINFOBYVEHICLE

    For ilVehicle = LBound(imVehListForAll) To UBound(imVehListForAll) - 1
        ilVefCode = imVehListForAll(ilVehicle)

        For ilLoopOnYear = 1 To 2
            If ilLoopOnYear = 1 Then
                ilStartMonth = 1
                ilEndMonth = imPeriods
            Else
                ilStartMonth = 13
                ilEndMonth = 13
            End If
            For ilLoopOnMonth = ilStartMonth To ilEndMonth
                ilLoopOnMonthMinusOne = ilLoopOnMonth - 1
                llSpots = 0
                'ReDim llPopEst(1 To 1) As Long
                'ReDim llCost(1 To 1) As Long
                'ReDim llAvgAud(1 To 1) As Long
                'ReDim llGrImp(1 To 1) As Long
                'ReDim llGRP(1 To 1) As Long
                'ReDim ilAVgRtg(1 To 1) As Integer
                
                ReDim llPopEst(0 To 0) As Long
                'ReDim llCost(0 To 0) As Long
                'ReDim flCost(0 To 0) As Single
                ReDim dlCost(0 To 0) As Double 'TTP 10439 - Rerate 21,000,000
                ReDim llAvgAud(0 To 0) As Long
                ReDim llGrImp(0 To 0) As Long
                ReDim llGRP(0 To 0) As Long
                ReDim ilAVgRtg(0 To 0) As Integer
                For llVehicleHdr = LBound(tmVehicle30UnitHdr) To UBound(tmVehicle30UnitHdr) - 1     'single contracts information by vehicle, dp & week
                    If tmVehicle30UnitHdr(llVehicleHdr).iVefCode = ilVefCode Then
                        'tmMonthInfoFinals are the vehicle/dp totals that gets written to GRF for crystal
                        tmMonthInfoFinals(ilVehicle).iVefCode = tmVehicle30UnitHdr(llVehicleHdr).iVefCode
                        tmMonthInfoFinals(ilVehicle).iRdfCode = tmVehicle30UnitHdr(llVehicleHdr).iRdfCode
                        tmMonthInfoFinals(ilVehicle).lOVStartTime = tmVehicle30UnitHdr(llVehicleHdr).lOVStartTime
                        tmMonthInfoFinals(ilVehicle).lOVEndTime = tmVehicle30UnitHdr(llVehicleHdr).lOVEndTime
                        llResearchPop = tmVehicle30UnitHdr(llVehicleHdr).lPop
                        llVehicleDetail = tmVehicle30UnitHdr(llVehicleHdr).lFirstIndex
                        'gather all lines monthly results and create array for the total contract monthly total

                        Do While llVehicleDetail >= 0
                            llUpper = UBound(llPopEst)
                            
                            llPopEst(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).lPopEst(ilLoopOnMonthMinusOne)
                            'llCost(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).lTotalCost(ilLoopOnMonthMinusOne)
                            'flCost(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).fTotalCost(ilLoopOnMonthMinusOne)
                            dlCost(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).dTotalCost(ilLoopOnMonthMinusOne) 'TTP 10439 - Rerate 21,000,000
                            llGrImp(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).lGrImp(ilLoopOnMonthMinusOne)
                            llGRP(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).lGRP(ilLoopOnMonthMinusOne)
                            ilAVgRtg(llUpper) = tmVehicle30UnitDetail(llVehicleDetail).iAvgRtg(ilLoopOnMonthMinusOne)
                            llSpots = llSpots + tmVehicle30UnitDetail(llVehicleDetail).lSpots(ilLoopOnMonthMinusOne)
                            
                            'ReDim Preserve llPopEst(1 To UBound(llPopEst) + 1) As Long
                            'ReDim Preserve llCost(1 To UBound(llCost) + 1) As Long
                            'ReDim Preserve llAvgAud(1 To UBound(llAvgAud) + 1) As Long
                            'ReDim Preserve llGrImp(1 To UBound(llGrImp) + 1) As Long
                            'ReDim Preserve llGRP(1 To UBound(llGRP) + 1) As Long
                            'ReDim Preserve ilAVgRtg(1 To UBound(ilAVgRtg) + 1) As Integer
                            ReDim Preserve llPopEst(0 To UBound(llPopEst) + 1) As Long
                            'ReDim Preserve llCost(0 To UBound(llCost) + 1) As Long
                            'ReDim Preserve flCost(0 To UBound(flCost) + 1) As Single
                            ReDim Preserve dlCost(0 To UBound(dlCost) + 1) As Double 'TTP 10439 - Rerate 21,000,000
                            ReDim Preserve llAvgAud(0 To UBound(llAvgAud) + 1) As Long
                            ReDim Preserve llGrImp(0 To UBound(llGrImp) + 1) As Long
                            ReDim Preserve llGRP(0 To UBound(llGRP) + 1) As Long
                            ReDim Preserve ilAVgRtg(0 To UBound(ilAVgRtg) + 1) As Integer
                        
                            llVehicleDetail = tmVehicle30UnitDetail(llVehicleDetail).lNextIndex
                        Loop
                    End If
                Next llVehicleHdr
                'If UBound(llPopEst) > 1 Then
                If UBound(llPopEst) > 0 Then
                     'adjust size of array, must be exact ize
                    'ReDim Preserve llPopEst(1 To UBound(llPopEst) - 1) As Long
                    'ReDim Preserve llCost(1 To UBound(llCost) - 1) As Long
                    'ReDim Preserve llAvgAud(1 To UBound(llAvgAud) - 1) As Long
                    'ReDim Preserve llGrImp(1 To UBound(llGrImp) - 1) As Long
                    'ReDim Preserve llGRP(1 To UBound(llGRP) - 1) As Long
                    'ReDim Preserve ilAVgRtg(1 To UBound(ilAVgRtg) - 1) As Integer
                    ReDim Preserve llPopEst(0 To UBound(llPopEst) - 1) As Long
                    'ReDim Preserve llCost(0 To UBound(llCost) - 1) As Long
                    'ReDim Preserve flCost(0 To UBound(flCost) - 1) As Single
                    ReDim Preserve dlCost(0 To UBound(dlCost) - 1) As Double 'TTP 10439 - Rerate 21,000,000
                    ReDim Preserve llAvgAud(0 To UBound(llAvgAud) - 1) As Long
                    ReDim Preserve llGrImp(0 To UBound(llGrImp) - 1) As Long
                    ReDim Preserve llGRP(0 To UBound(llGRP) - 1) As Long
                    ReDim Preserve ilAVgRtg(0 To UBound(ilAVgRtg) - 1) As Integer

                    'done reading all vehicle/dp for one month
                    '10-30-14 default to use 1 place rating regardless of agency flag
                    'gResearchTotalsFloat "1", True, llResearchPop, flCost(), llGrImp(), llGRP(), llSpots, llMonthCost, ilMonthAvgRtg, llMonthGrimp, llMonthGRP, llMonthCPP, llMonthCPM, llMonthAVgAud, flMonthCost
                    gResearchTotalsFloat "1", True, llResearchPop, dlCost(), llGrImp(), llGRP(), llSpots, dlMonthCost, ilMonthAvgRtg, llMonthGrimp, llMonthGRP, llMonthCPP, llMonthCPM, llMonthAVgAud, flMonthCost 'TTP 10439 - Rerate 21,000,000

                    'save monthly results for contract totals later
                    tmMonthInfoFinals(ilVehicle).lAvgAud(ilLoopOnMonthMinusOne) = llMonthAVgAud
                    'tmMonthInfoFinals(ilVehicle).lTotalCost(ilLoopOnMonthMinusOne) = llMonthCost
                    'tmMonthInfoFinals(ilVehicle).fTotalCost(ilLoopOnMonthMinusOne) = flMonthCost
                    tmMonthInfoFinals(ilVehicle).dTotalCost(ilLoopOnMonthMinusOne) = dlMonthCost 'TTP 10439 - Rerate 21,000,000
                    tmMonthInfoFinals(ilVehicle).iAvgRtg(ilLoopOnMonthMinusOne) = ilMonthAvgRtg
                    tmMonthInfoFinals(ilVehicle).lCPP(ilLoopOnMonthMinusOne) = llMonthCPP
                    tmMonthInfoFinals(ilVehicle).lCPM(ilLoopOnMonthMinusOne) = llMonthCPM
                    tmMonthInfoFinals(ilVehicle).lGrImp(ilLoopOnMonthMinusOne) = llMonthGrimp
                    tmMonthInfoFinals(ilVehicle).lGRP(ilLoopOnMonthMinusOne) = llMonthGRP
                End If
                llVehicleDetail = tmVehicle30UnitHdr(llVehicleHdr).lFirstIndex      're-establish the start of the vehicle detail dp summaries
            Next ilLoopOnMonth
        Next ilLoopOnYear
           
    Next ilVehicle
    Exit Sub
End Sub
