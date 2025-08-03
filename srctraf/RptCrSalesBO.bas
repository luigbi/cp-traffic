Attribute VB_Name = "RptCrSalesBO"
Option Explicit
Option Compare Text
Dim tmSdfSrchKey1 As SDFKEY1            'SDF by vehicle, date, time, schstatus
Dim tmSdfSrchKey4 As SDFKEY4            'sdf  by date, chfcode

Dim hmSdf As Integer
Dim tmSdf As SDF
Dim imSdfRecLen As Integer

Dim hmSmf As Integer
Dim tmSmf As SMF                'Spot Makegood
Dim imSmfRecLen As Integer

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length

Dim hmVsf As Integer            'Vehicle file handle
Dim tmVsf As VSF                'VEF record image
Dim imVsfRecLen As Integer       'VEF record length

Dim hmCHF As Integer            'Contract Header file handle
Dim tmChf As CHF                'CHF record image
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChfSrchKey1 As CHFKEY1    'key by contract#
Dim tmChfSrchKey As LONGKEY0

Dim hmClf As Integer            'Contract Line file handle
Dim tmClf As CLF                'CLF record image
Dim imClfRecLen As Integer      'CLF record length
Dim tmClfSrchKey As CLFKEY0

Dim hmCff As Integer            'Contract Flight file handle
Dim tmCff As CFF                'CFF record image
Dim imCffRecLen As Integer      'CFF record length

Dim hmAdf As Integer            'Advertiser  file handle
Dim tmAdf As ADF                'ADF record image
Dim imAdfRecLen As Integer      'ADF record length
Dim tmAdfSrchKey0 As INTKEY0

Dim hmAgf As Integer            'Agency  file handle
Dim tmAgf As ADF                'AGF record image
Dim imAgfRecLen As Integer      'AGF record length

Dim hmGrf As Integer            'Temp  file handle
Dim tmGrf As GRF                'Temp file record image
Dim imGrfRecLen As Integer      'Temp file record length

Dim hmMnf As Integer            'Multi-list file handle
Dim tmMnf As MNF                '
Dim imMnfRecLen As Integer

Dim hmRvf As Integer            'Receivables  file handle
Dim tmRvf As RVF                'RVF file record image
Dim imRvfRecLen As Integer      'RVF file record length
Dim hmPhf As Integer            'History  file handle

Dim hmSbf As Integer            'Special Billing  file handle
Dim tmSbf As SBF                'SBF file record image
Dim imSbfRecLen As Integer      'SBF file record length

Dim imUsevefcodes() As Integer        'array of vehicle codes to include/exclude
Dim imInclVefCodes As Integer               'flag to incl or exclude vehicle codes
Dim imUseAdvtCodes() As Integer        'array of advt codes to include/exclude
Dim imInclAdvtCodes As Integer               'flag to incl or exclude advt codes

Dim smCntrType As String            'contract types to include for last year retrieval of contracts
Dim smCntrStatus As String          'contract statuses to include for last year retrieval of contracts
Dim smMonthsInYear As String * 36    'JanFebMar.....or the corp months starting with the fiscal start month (i.e. OctNovDec...)

Dim imPeriods As Integer            '# periods to generate for current year
Dim imMnthsOffForNew As Integer     '# months off before considered new (from site)
Dim imMnthsNewIsNew As Integer      '# months new remains new (site)
Dim smNewBusYearType As String * 1  'R = rolling, C = calendar (site)
Dim imSlspSplit As Integer          'do slsp splits
Dim imGrossNetSpot As Integer       '1 = gross, 2 = net, 3 = spot count
Dim imPerType As Integer            '1=week (not implemented), 2= std, 3 = corp, 4= cal
'Dim lmStartDates(1 To 15) As Long   'weekly is max 14 weeks (some 14 week qtrs).  Weekly currently not implemented in this report.
Dim lmStartDates(0 To 15) As Long   'weekly is max 14 weeks (some 14 week qtrs).  Weekly currently not implemented in this report.
                                    'much code extracted/copied from Revenue on the Books, which had weekly option. Index zero ignored
Dim lmLYStartDates() As Long     'last year start dates (# months inactivity for New + # months New is New)
Dim lmCalAmt() As Long              'calendar calcs
Dim lmLYCalAmt() As Long            'Last year calendar active flags (no $ accumulated, just indication of activity or not)
Dim lmCalAcqAmt() As Long           'calendar calcs
Dim lmLYCalAcqAmt() As Long         'last year calendar acq costs
'Dim lmProject(1 To 14) As Long             'calendar calcs
Dim lmProject(0 To 14) As Long             'calendar calcs. Index zero ignored
'Dim lmAcquisition(1 To 14) As Long         'calendar calcs
Dim lmAcquisition(0 To 14) As Long         'calendar calcs. Index zero ignored
'Dim lmProjectSpots(1 To 14) As Long        'calendar spot counts
Dim lmProjectSpots(0 To 14) As Long        'calendar spot counts. Index zero ignored
'Dim lmLYProjectSpots(1 To 14) As Long       'last year calendar spot counts. Not used
Dim lmCalSpots() As Long
Dim lmLYCalSpots() As Long
Dim lmSingleCntr As Long
Dim tmSdfInfo() As SDFSORTBYLINE
Dim tmSbfList() As SBF


Dim imSlspSplitCodes(0 To 9) As Integer
Dim lmSlspSplitPct(0 To 9) As Long

Dim tmCntTypes As CNTTYPES
Dim tmTranTypes As TRANTYPES
Dim tmNTRTypes As SBFTypes
Dim tmSpotTypes As SPOTTYPES
Dim tmSpotAndRev() As SPOTBBSTATS       'array of unique contracts/line # for a vehicle in a given month or week
Dim tmPriceTypes As PRICETYPES
Dim tmNewAdvtBus() As NEWADVTBUS

Dim tmChfAdvtExt() As CHFADVTEXT            'contracts obtained for previous year
Dim imDormantVehicles() As Integer
Type NEWADVTBUS
    sKey As String * 5          'advt internal code
    iAdfCode As Integer
    'iLYIndexMonthLastAired As Integer   'last years last month index of advt airing (based on the start month of report)
                                        'ie.  jun - may : jun will be index 1, may = index 12
    lLYMonthLastAired As Long   'last years last month advt airing (start date of the last month for the type of period requested)
    'iTYIndexMonthFirstAired As Integer  'this years first month index of advt airing
    lTYMonthFirstAired As Long         'this years earliest date of advt airing (start date of the first month for the type of period requested)
    iLastMonthYearNew(0 To 1) As Integer  'from advt file, last month/year considered new
    iUpdateAdvt As Integer              'true to update advt with first time detected new with year & month
End Type

Type SPOTBBSTATS
    lChfCode As Long            'contrct code
    iVefCode As Integer         'vehicle code
    iLineNo As Integer          'line #
    lCntrNo As Long             'contract #
    iAdfCode As Integer         'advertiser code
    iAgfCode As Integer         'agency code
    iMnfComp As Integer         'product protection code
    iMnfRevSet As Integer       'bus category code
    iPctTrade As Integer        'pct of trade
    iAgyCommPct As Integer      'comm % from agy
    iIsItHardCost As Integer    'hard cost or not
    iNTRMnfType As Integer      'NTR type
    iSlfCode(0 To 9) As Integer 'slsp splits (codes)
    lComm(0 To 9) As Long       'slsp splits commission %
    sTradeComm As String * 1    'trade commissionable
    sIsNTR As String * 1        'NTR (Y/N)
    sCashTrade As String * 1    'Cash or Trade  (required for receivables transactions)
    iVG As Integer              'vehicle group if applicable
    'lSpots(1 To 14) As Long     '# spots gathered from vehicle for contract/line
    lSpots(0 To 13) As Long     '# spots gathered from vehicle for contract/line
    'lRev(1 To 14) As Long       '$ gathered from vehicle for contract/line
    lRev(0 To 13) As Long       '$ gathered from vehicle for contract/line
    iNew As Integer             '1 = new, 0 = renewal
    iIsItPolitical As Integer   'political advt flag (true/false)
    iGame As Integer
    iDate(0 To 1) As Integer    'event date
    lghfcode As Long
End Type
'
'           Open files required for Spot Business Booked
'           Return - error flag = true for open error
'
Private Function mOpenSalesBOFiles() As Integer
Dim ilRet As Integer
Dim slTable As String * 3
Dim ilError As Integer

    ilError = False
    On Error GoTo mOpenSalesBOFilesErr

    slTable = "Chf"
    hmCHF = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Function
    End If
    imCHFRecLen = Len(tmChf)
    
    slTable = "Clf"
    hmClf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        Exit Function
    End If
    imClfRecLen = Len(tmClf)

    slTable = "Cff"
    hmCff = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        Exit Function
    End If
    imCffRecLen = Len(tmCff)
    
    slTable = "Grf"
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Exit Function
    End If
    imGrfRecLen = Len(tmGrf)

    slTable = "Mnf"
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        Exit Function
    End If
    imMnfRecLen = Len(tmMnf)
    
    slTable = "Vef"
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Function
    End If
    imVefRecLen = Len(tmVef)
    
    slTable = "Vsf"
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
        Exit Function
    End If
    imVsfRecLen = Len(tmVsf)
   
    slTable = "Sdf"
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        Exit Function
    End If
    imSdfRecLen = Len(tmSdf)
        
    slTable = "Smf"
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmSmf)
        btrDestroy hmSmf
        Exit Function
    End If
    imSmfRecLen = Len(tmSmf)
    
    slTable = "Adf"
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmAdf)
        btrDestroy hmAdf
        Exit Function
    End If
    imAdfRecLen = Len(tmAdf)
    
    slTable = "Agf"
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmAgf)
        btrDestroy hmAgf
        Exit Function
    End If
    imAgfRecLen = Len(tmAgf)
        
    slTable = "Rvf"
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        Exit Function
    End If
    imRvfRecLen = Len(tmRvf)

    slTable = "Phf"
    hmPhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmPhf)
        btrDestroy hmPhf
        Exit Function
    End If
    imRvfRecLen = Len(tmRvf)
    
    slTable = "Sbf"
    hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSalesBOFiles = True
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf
        Exit Function
    End If
    imSbfRecLen = Len(tmSbf)


    Exit Function
    
mOpenSalesBOFilesErr:
    ilError = Err.Number
    gBtrvErrorMsg ilRet, "mOpenSalesBOFiles (OpenError) #" & str(ilError) & ": " & slTable, RptSelSN

    Resume Next
End Function
'
'       Generate prepass for Spot Business Booked
'       for spots within a span of dates for selective advertisers,
'       and vehicles.  Determine what advertisers are new based
'       on a Site question for the # of months off.  A new advertiser
'       is also considered new for "x" months which is based on a Site question
'       Data is gathered for 2 years (last year, this year) to determine which
'       advertisers are new
'
Public Sub gGenSalesBO()

Dim ilError As Integer
Dim ilVefCode As Integer
Dim llStart As Long
Dim slEarliestDate As String
Dim slLatestDate As String
Dim llEnd As Long
Dim slPerStartDate As String
Dim slPerEndDate As String
Dim slLYEarliestDate As String
Dim slLYLatestDate As String
Dim ilWhichKey As Integer
Dim ilRet As Integer
Dim llLoopOnSpots As Long
Dim ilOk As Integer
Dim ilLoop As Integer
Dim slType As String
Dim slNameCode As String
Dim slCode As String
Dim ilFoundCntr As Integer
Dim llLoopOnKey As Long
Dim ilHowManyPer As Integer
Dim ilIncludeContract As Integer
Dim ilIncludeVehicle As Integer
Dim ilByCodeOrNumber As Integer
Dim llChfCode As Long
'ReDim ilVehiclesToProcess(1 To 1) As Integer
ReDim ilVehiclesToProcess(0 To 0) As Integer

Dim llLoopOnStats As Long
Dim llSpotRate As Long
Dim slSpotRate As String
Dim ilIndex As Integer
Dim ilTemp As Integer
Dim ilfirstTime As Integer
Dim ilHOState As Integer
Dim ilCurrentRecd As Integer
Dim llLYContrCode As Long
Dim ilFound As Integer
Dim ilClf As Integer
Dim ilFoundVeh As Integer
Dim ilThisYear As Integer      'true, else false for last year
Dim ilDoRep As Integer          'true if process rep
Dim ilDoAirTime As Integer      'true if process air time
Dim ilIsItPolitical As Integer

        ilError = mOpenSalesBOFiles()
        If ilError Then
            Exit Sub            'at least 1 open error
        End If
                
        mObtainSelectivity
        llChfCode = mSingleContract()       'get selective contract header
        If llChfCode < 0 Then               'illegal contract
            Exit Sub
        End If
            
        ilfirstTime = True
        
        mGetCodesFromList ilVehiclesToProcess()               'setup array of codes to include or exclude, which is less for speed
        
        llStart = lmStartDates(1)       'earliest date
        slEarliestDate = Format$(llStart, "m/d/yy")
        llEnd = lmStartDates(imPeriods + 1) - 1 'latest date
        slLatestDate = Format$(llEnd, "m/d/yy")
        
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf.lGenTime = lgNowTime
        tmGrf.iGenDate(0) = igNowDate(0)
        tmGrf.iGenDate(1) = igNowDate(1)
        
        slLYEarliestDate = Format$(lmLYStartDates(1), "m/d/yy")
        slLYLatestDate = Format$(lmLYStartDates(imMnthsOffForNew + 1) - 1, "m/d/yy")
           
        ReDim tmNewAdvtBus(0 To 0) As NEWADVTBUS
        'All types of orders (airtime,ntr,rep) are gathered to determine business LAST YEAR
        'Create array of advertisers saving the last time $ were recorded
        'Gathering the information for which advertisers aired last year/this year will retrieve holds, orders, unsch & sch
        ilThisYear = False     'last year
        mGatherAdvtAiringFromContract llChfCode, slLYEarliestDate, slLYLatestDate, lmLYStartDates(), ilThisYear, imMnthsOffForNew
        ilThisYear = True
        mGatherAdvtAiringFromContract llChfCode, slEarliestDate, slLatestDate, lmStartDates(), ilThisYear, imPeriods
        
        ilThisYear = False
        mGatherAdvtForNTR llChfCode, slLYEarliestDate, slLYLatestDate, lmLYStartDates(), ilThisYear, imMnthsOffForNew
        ilThisYear = True
        mGatherAdvtForNTR llChfCode, slEarliestDate, slLatestDate, lmStartDates(), ilThisYear, imMnthsOffForNew
        'Sort the advertisers that had activity in the previous year so it can quickly found
        'when creating the current years data.  Need to flag the created record in grf with new/renewal flag
        If UBound(tmNewAdvtBus) - 1 > 1 Then
            ArraySortTyp fnAV(tmNewAdvtBus(), 0), UBound(tmNewAdvtBus), 0, LenB(tmNewAdvtBus(0)), 0, LenB(tmNewAdvtBus(0).sKey), 0
        End If
                
        If lmSingleCntr > 0 Then
            ilWhichKey = INDEXKEY0      'search vef, cntr
        Else
            ilWhichKey = INDEXKEY1      'search sdf by vef, date
        End If
        If tmCntTypes.iAirTime Then             'include air time
            'loop on vehicles by month or week
            'If All  Vehicles selected, insure that the dormant vehicles are processed since the vehicle name
            'was notin the listof vehicless (dormant state not shown)
            ilLoop = UBound(ilVehiclesToProcess)
            For llLoopOnKey = LBound(imDormantVehicles) To UBound(imDormantVehicles) - 1
                ilVehiclesToProcess(ilLoop) = imDormantVehicles(llLoopOnKey)
                ilLoop = ilLoop + 1
                ReDim Preserve ilVehiclesToProcess(LBound(ilVehiclesToProcess) To ilLoop) As Integer
            Next llLoopOnKey
            'gather the spots in memory for the period and create another array that contains all the
            'necessary header and $ information to write to prepass once the vehicle/period has been processed
            For llLoopOnKey = LBound(ilVehiclesToProcess) To UBound(ilVehiclesToProcess) - 1
                ReDim tmSpotAndRev(0 To 0) As SPOTBBSTATS           'init array of contract information
                ilVefCode = ilVehiclesToProcess(llLoopOnKey)
                ilIncludeVehicle = mFilterVehicle(ilVefCode)
                If ilIncludeVehicle Then
                    'build spots for one period at a time.  Up to 1 year too many.
                    For ilHowManyPer = 1 To imPeriods
                        ReDim tmSdfInfo(0 To 0) As SDFSORTBYLINE            'init array of spots
                        slPerStartDate = Format$(lmStartDates(ilHowManyPer), "m/d/yy")
                        slPerEndDate = Format$(lmStartDates(ilHowManyPer + 1) - 1, "m/d/yy")
                        ilRet = gGetSpotsbyVefDateAndSort(hmSdf, ilWhichKey, ilVefCode, llChfCode, slPerStartDate, slPerEndDate, tmSpotTypes, tmSdfInfo())
                        'return array of spots sorted by chfcode, line#, date (for 1 period which is a month or week)
                        For llLoopOnSpots = LBound(tmSdfInfo) To UBound(tmSdfInfo) - 1      'loop on spots to get the contract info and store into array (tmSpotAndRev)
                            tmSdf = tmSdfInfo(llLoopOnSpots).tSdf
                            'filter out selections
                            If tmChf.lCode <> tmSdf.lChfCode Or ilfirstTime Then
                                ilByCodeOrNumber = 0            'reference by chf code
                                ilOk = mFilterContract(ilByCodeOrNumber, tmSdf.lChfCode, ilfirstTime, ilIsItPolitical)
                                ilfirstTime = False
                                If ilOk Then
                                    If Not gFilterLists(tmChf.iAdfCode, imInclAdvtCodes, imUseAdvtCodes()) Then
                                        ilOk = False
                                    End If
                                End If
                                ilIncludeContract = ilOk        'save the results of the filtering so it has to be done only once per contract
                            End If
                            If ilIncludeContract Then
                                'Get spot rates, only reread sched line if different
                                'reduce amt of times to read the schedule line.
                                llSpotRate = 0
                                If (tmClf.lChfCode <> tmSdf.lChfCode) Or (tmSdf.iLineNo <> tmClf.iLine) Then
                                    tmClfSrchKey.lChfCode = tmSdf.lChfCode
                                    tmClfSrchKey.iLine = tmSdf.iLineNo
                                    tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                                    tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                    If (ilRet <> BTRV_ERR_NONE) Or (tmClf.lChfCode <> tmChf.lCode) Or (tmClf.iLine <> tmSdf.iLineNo) Then
                                        llSpotRate = -1     'no line found
                                    End If
                                End If
                                If llSpotRate >= 0 Then
                                    llSpotRate = mGetRateAndAddToArray(ilHowManyPer)
                                End If          'spotrate >= 0
                            End If              'include contract = true
                        Next llLoopOnSpots  'next spot
                    Next ilHowManyPer           'loop on # periods to process
                    'all periods gathered for the 1 vehicle and stored in array by contract & Line; create records to temporary file
                    mUpdateGRFForAll

                End If
            Next llLoopOnKey
            Erase tmSdfInfo
            Erase tmSpotAndRev
        End If
        
        'Gather NTR or Hard Cost, dont process if spot counts requested
        If ((tmCntTypes.iNTR) Or (tmCntTypes.iHardCost)) And imGrossNetSpot <> 3 Then
            'mProcessNTR llChfCode, slEarliestDate, slLatestDate
            ilThisYear = True         'current year
            ilDoRep = True          'process REP contracts for current year
            ilDoAirTime = False     'air time for current year obtained from spots
            mGatherCurrentYearNTR llChfCode, slEarliestDate, slLatestDate, lmStartDates(), imPeriods
        End If
        
        
        'Gather Adjustments, dont process if spot counts requested
        If tmTranTypes.iAdj And imGrossNetSpot <> 3 Then
            mProcessAdj slEarliestDate, slLatestDate
        End If
        
        'Gather REP
        If tmCntTypes.iRep Then
            ilThisYear = True         'current year
            ilDoRep = True          'process REP contracts for current year
            ilDoAirTime = False     'air time for current year obtained from spots
            mGatherCurrentYearREP llChfCode, slEarliestDate, slLatestDate, lmStartDates(), imPeriods
        End If
       
        'udpate the Advertiser file for all new Advt
        For llLoopOnKey = LBound(tmNewAdvtBus) To UBound(tmNewAdvtBus) - 1
            If tmNewAdvtBus(llLoopOnKey).iUpdateAdvt = True Then
                tmAdfSrchKey0.iCode = tmNewAdvtBus(llLoopOnKey).iAdfCode
                ilRet = btrGetEqual(hmAdf, tmAdf, Len(tmAdf), tmAdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching adf recd
                If ilRet = BTRV_ERR_NONE Then
                    tmAdf.iLastMonthNew = tmNewAdvtBus(llLoopOnKey).iLastMonthYearNew(0)        'Month found as New
                    tmAdf.iLastYearNew = tmNewAdvtBus(llLoopOnKey).iLastMonthYearNew(1)         'year found as new
                    ilRet = btrUpdate(hmAdf, tmAdf, imAdfRecLen)
                End If
            End If
        Next llLoopOnKey
        
        'close all files
        mCloseSalesBOFiles
        Erase ilVehiclesToProcess, tmChfAdvtExt, imDormantVehicles
        Erase tmNewAdvtBus
        Exit Sub
End Sub
'           mObtainSelectivity - gather all selectivity entered and place
'           in common variables
'           'Determine this year/last year date spans
Private Sub mObtainSelectivity()
Dim slStart As String
Dim llStart As Long
Dim ilDay As Integer
Dim slStamp As String
Dim ilRet As Integer

        
        'what type of periods:  week, std, corporate, calendar months
        If RptSelBO!rbcPerType(0).Value = True Then     'week
            imPerType = 1
        ElseIf RptSelBO!rbcPerType(1).Value = True Then 'std
            imPerType = 2
        ElseIf RptSelBO!rbcPerType(2).Value = True Then 'corp
            imPerType = 3
        Else
            imPerType = 4                                   'cal
        End If
        
        imPeriods = Val(RptSelBO!edcPeriods.Text)
        
        'setup defaults if questions are not answered
        If tgSpf.iNoMnthNewBus = 0 Then     '# months off considered new
            imMnthsOffForNew = 13
        Else
            imMnthsOffForNew = tgSpf.iNoMnthNewBus
        End If
        If tgSpf.iNoMnthNewIsNew = 0 Then           '"# months new after becoming new
            imMnthsNewIsNew = 6
        Else
            imMnthsNewIsNew = tgSpf.iNoMnthNewIsNew
        End If
        If Trim(tgSpf.sNewBusYearType) = "" Then
            smNewBusYearType = "R"          'rolling year
        Else
            smNewBusYearType = tgSpf.sNewBusYearType
        End If
        If smNewBusYearType = "C" Then          'calendar year
            imMnthsOffForNew = 12          'hard coded :  if using Calendar year doesnt make sense to allow for # of months off to
                                            'vary, as its no business for the entire previous year
            imMnthsNewIsNew = 12            'hard coded:  anything in the entire year is considered New if not aired the previous year
        End If
        
        'weekly not implemented in this version of report
        If imPerType = 1 Then  'set start dates of Weekly periods.
            
        ElseIf imPerType = 2 Then   'set start dates of 12 standard periods
            slStart = str$(igMonthOrQtr) & "/15/" & str$(igYear)
            gBuildStartDates slStart, 1, imPeriods + 1, lmStartDates()
            'backup "x" months for last years data, determine to go back by the calendar year of rolling months
            mBuildLYStartDates imPerType, igYear, lmStartDates(1)
        ElseIf imPerType = 3 Then   'set start dates of 12 corporate periods
            slStart = str$(igMonthOrQtr) & "/15/" & str$(igYear)
            gBuildStartDates slStart, 2, imPeriods + 1, lmStartDates()
            mBuildLYStartDates imPerType, igYear, lmStartDates(1)
        ElseIf imPerType = 4 Then  'set start dates of 12 calendar periods
            slStart = str$(igMonthOrQtr) & "/1/" & str$(igYear)
            gBuildStartDates slStart, 4, imPeriods + 1, lmStartDates()
            mBuildLYStartDates imPerType, igYear, lmStartDates(1)
        End If
            
        'gross, net or spot counts
        If RptSelBO!rbcGrossNet(0).Value = True Then        'gross
            imGrossNetSpot = 1
        ElseIf RptSelBO!rbcGrossNet(1).Value = True Then        'net
            imGrossNetSpot = 2
        Else
            imGrossNetSpot = 3
        End If

        'Selective contract #
        lmSingleCntr = Val(RptSelBO!edcContract.Text)
               
        tmCntTypes.iHold = gSetCheck(RptSelBO!ckcAllTypes(0).Value)
        tmCntTypes.iOrder = gSetCheck(RptSelBO!ckcAllTypes(1).Value)
        tmCntTypes.iStandard = gSetCheck(RptSelBO!ckcAllTypes(3).Value)
        tmCntTypes.iReserv = gSetCheck(RptSelBO!ckcAllTypes(4).Value)
        tmCntTypes.iRemnant = gSetCheck(RptSelBO!ckcAllTypes(5).Value)
        tmCntTypes.iDR = gSetCheck(RptSelBO!ckcAllTypes(6).Value)
        tmCntTypes.iPI = gSetCheck(RptSelBO!ckcAllTypes(7).Value)
        tmCntTypes.iPSA = gSetCheck(RptSelBO!ckcAllTypes(8).Value)
        tmCntTypes.iPromo = gSetCheck(RptSelBO!ckcAllTypes(9).Value)
        tmCntTypes.iTrade = gSetCheck(RptSelBO!ckcAllTypes(10).Value)
        tmCntTypes.iAirTime = gSetCheck(RptSelBO!ckcAllTypes(11).Value)
        tmCntTypes.iRep = gSetCheck(RptSelBO!ckcAllTypes(12).Value)
        tmCntTypes.iNTR = gSetCheck(RptSelBO!ckcAllTypes(13).Value)
        tmCntTypes.iHardCost = gSetCheck(RptSelBO!ckcAllTypes(14).Value)
        tmCntTypes.iPolit = gSetCheck(RptSelBO!ckcAllTypes(15).Value)
        tmCntTypes.iNonPolit = gSetCheck(RptSelBO!ckcAllTypes(16).Value)
        tmCntTypes.iMissed = gSetCheck(RptSelBO!ckcAllTypes(17).Value)      'spot type inclusion/exclusion uses tmSpotTypes structure
        tmCntTypes.iCancelled = gSetCheck(RptSelBO!ckcAllTypes(18).Value)   'spot type inclusion/exclusion uses tmSpotTypes structure
        tmCntTypes.iXtra = gSetCheck(RptSelBO!ckcAllTypes(19).Value)            'fill (not used)
        tmCntTypes.iCash = True                'always include cash
        
        smCntrStatus = ""                 'statuses: hold, order, unsch hold, uns order
        If tmCntTypes.iHold Then                  'exclude holds and uns holds
            smCntrStatus = "HG"             'include orders and uns orders
        End If
        
        If tmCntTypes.iOrder Then                  'exclude holds and uns holds
            smCntrStatus = smCntrStatus & "ON"             'include orders and uns orders
        End If
        smCntrType = ""
        If tmCntTypes.iStandard Then
            smCntrType = "C"
        End If
        If tmCntTypes.iReserv Then
            smCntrType = smCntrType & "V"
        End If
        If tmCntTypes.iRemnant Then
            smCntrType = smCntrType & "T"
        End If
        If tmCntTypes.iDR Then
            smCntrType = smCntrType & "R"
        End If
        If tmCntTypes.iPI Then
            smCntrType = smCntrType & "Q"
        End If
        If tmCntTypes.iPSA Then
            smCntrType = smCntrType & "S"
        End If
        If tmCntTypes.iPromo Then
            smCntrType = smCntrType & "M"
        End If
        If smCntrType = "CVTRQSM" Then          'all types: PI, DR, etc.  except PSA(p) and Promo(m)
            smCntrType = ""                     'blank out string for "All"
        End If
        
        'spot types for inclusion/exclusion
        tmSpotTypes.iMG = True
        tmSpotTypes.iSched = True
        tmSpotTypes.iOutside = True
        'open,close & fills dont have $,can ignore for gross or net option
        tmSpotTypes.iClose = False
        tmSpotTypes.iOpen = False
        tmSpotTypes.iFill = False
        tmSpotTypes.iMissed = tmCntTypes.iMissed
        tmSpotTypes.iCancel = tmCntTypes.iCancelled
        
        'line types to include:  only chargeable for revenue; otherwise all
        tmPriceTypes.iCharge = True     'Chargeable lines
        tmPriceTypes.iZero = False      '.00 lines
        tmPriceTypes.iADU = False     'adu lines
        tmPriceTypes.iBonus = False          'bonus lines
        tmPriceTypes.iNC = False          'N/C lines
        tmPriceTypes.iRecap = False        'recapturable
        tmPriceTypes.iSpinoff = False      'spinoff
        If imGrossNetSpot = 3 Then          'spot counts
            'include all zero $ (no options for N/C)
            tmPriceTypes.iZero = True      '.00 lines
            tmPriceTypes.iADU = True     'adu lines
            tmPriceTypes.iBonus = True          'bonus lines
            tmPriceTypes.iNC = True          'N/C lines
            tmPriceTypes.iRecap = True        'recapturable
            tmPriceTypes.iSpinoff = True      'spinoff
            'only option for spot counts is to exclude fills and billboards
            If tmCntTypes.iXtra = True Then     'include fills
                tmSpotTypes.iFill = True
            End If
            If RptSelBO!ckcAllTypes(20).Value = vbChecked Then  'include billboards
                tmSpotTypes.iClose = True
                tmSpotTypes.iOpen = True
            End If
        End If
    


        'If including adjustments from receivables, look for AN transaction types for NTR, Air Time or Hard Cost (also by option)
       'If including adjustments from receivables, look for AN transaction types for NTR, Air Time or Hard Cost (also by option)
        'tmTranTypes.iAdj = gSetCheck(RptSelSpotBB!ckcAdj(0).Value)     'incl rep adjustments
        If RptSelBO!ckcAdj(0).Value = vbChecked Or RptSelBO!ckcAdj(1).Value = vbChecked Then
            tmTranTypes.iAdj = True
        Else
            tmTranTypes.iAdj = False
        End If
        tmTranTypes.iInv = False
        tmTranTypes.iWriteOff = False
        tmTranTypes.iPymt = False
        tmTranTypes.iNTR = tmCntTypes.iNTR
        '12-1-16
'        If tmCntTypes.iRep Then
'            tmTranTypes.iAirTime = True         'need to do further testing to filter out the airtime; adjustments for rep included (not scheduled spots)
'        Else
'            tmTranTypes.iAirTime = False
'        End If
        
        If tmTranTypes.iAdj Then                '12-1-16 air time designates to go thru receivables if any adj required
            tmTranTypes.iAirTime = True
        End If
        tmTranTypes.iHardCost = tmCntTypes.iHardCost
        tmTranTypes.iCash = True
        tmTranTypes.iMerch = False
        tmTranTypes.iPromo = False
        tmTranTypes.iTrade = gSetCheck(RptSelBO!ckcAllTypes(10).Value)
        
        'NTR types; SBF file has installment records and import records.  Only NTR (Hardcost) of interest
        tmNTRTypes.iImport = False
        tmNTRTypes.iInstallment = False
        If tmCntTypes.iNTR = True Or tmCntTypes.iHardCost = True Then
            tmNTRTypes.iNTR = True
            ilRet = gObtainMnfForType("I", slStamp, tgNTRMnf())        'NTR Item types to check for hard cost
        Else
            tmNTRTypes.iNTR = False
        End If
        
End Sub
Private Sub mCloseSalesBOFiles()
Dim ilRet As Integer

        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        ilRet = btrClose(hmSmf)
        btrDestroy hmSmf
        ilRet = btrClose(hmAdf)
        btrDestroy hmAdf
        ilRet = btrClose(hmAgf)
        btrDestroy hmAgf
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        ilRet = btrClose(hmPhf)
        btrDestroy hmPhf
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf
        
        Erase imUsevefcodes
        Erase imUseAdvtCodes
        Erase lmStartDates
        Erase tmSpotAndRev
        Erase tmSdfInfo
        Erase tmSbfList
        Erase lmCalSpots, lmCalAmt, lmCalAcqAmt
   
    Exit Sub
    
End Sub
'
'               Obtain contract and filter selectivity
'               <input> ilByCodeOrNUmber: 0 = use code to retrieve contract
'                                         1 = use contract # (receivables dont have chfcodes)
'                       llContractKey: contract code or Number
'                       ilIsItPolitical - true if Political advt, else false
Private Function mFilterContract(ilByCodeOrNumber As Integer, llContractKey As Long, ilfirstTime As Integer, ilIsItPolitical As Integer) As Integer
Dim ilOk As Integer
Dim ilFoundCntr As Integer
Dim ilRet As Integer

        ilOk = True
        If ilByCodeOrNumber = 0 Then            'get the contract by code
            If llContractKey <> tmChf.lCode Then
                tmChfSrchKey.lCode = llContractKey
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                If tmChf.lCntrNo = 1161 Then
                    ilRet = ilRet
                End If
            End If
            '4-8-15 ignore contracts that are not orders and not scheduled
            '8-22-16 chg test from chfType to chfStatus
            If (tmChf.sDelete = "Y") Or (tmChf.sStatus <> "H" And tmChf.sStatus <> "O") And (tmChf.sSchStatus <> "F") Then       'deleted header, not an order and not scheduled, contract shouldnt be used
'            If (tmChf.sDelete = "Y") Or (tmChf.sType <> "H" And tmChf.sType <> "O") And (tmChf.sSchStatus <> "F") Then       'deleted header, not an order and not scheduled, contract shouldnt be used
                ilOk = False
            End If
        Else                                    'get the contract by #
            If llContractKey <> tmChf.lCntrNo Then
                tmChfSrchKey1.lCntrNo = llContractKey
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llContractKey)
                    '4-8-15 ignore contracts that are not fully scheduled and not orders
'                    If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.sDelete <> "Y") And (tmChf.sType = "H" Or tmChf.sType = "O") Then
                    '12-2-16 wrong field tested in chf:  Type vs Status
                    If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.sDelete <> "Y") And (tmChf.sStatus = "H" Or tmChf.sStatus = "O") Then
                        ilFoundCntr = True
                        Exit Do
                    End If
                    ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If Not ilFoundCntr Then
                    ilOk = False
                End If
            End If
        End If
        'contract found
        If (ilOk) Or ilfirstTime Then
        
            'scheduled/unsch holds
            If (tmChf.sStatus = "H" Or tmChf.sType = "G") And Not (tmCntTypes.iHold) Then   '8-22-16 chg test from type to status
                ilOk = False
            End If
            'scheduled/unsch orders
            If (tmChf.sStatus = "O" Or tmChf.sType = "N") And Not (tmCntTypes.iOrder) Then      '8-22-16 chg test from type to status
                ilOk = False
            End If
            'standard orders
            If tmChf.sType = "C" And Not (tmCntTypes.iStandard) Then        '8-22-16 chg test from status to type
                ilOk = False
            End If
            'reserved
            If tmChf.sType = "V" And Not (tmCntTypes.iReserv) Then          '8-22-16 chg test from status to type
                ilOk = False
            End If
            'Remnant
            If tmChf.sType = "T" And Not (tmCntTypes.iRemnant) Then         '8-22-16 chg test from status to type
                ilOk = False
            End If
            'Direct Response
            If tmChf.sType = "R" And Not (tmCntTypes.iDR) Then              '8-22-16 chg test from status to type
                ilOk = False
            End If
            'Per Inquiry
            If tmChf.sType = "Q" And Not (tmCntTypes.iPI) Then              '8-22-16 chg test from status to type
                ilOk = False
            End If
            'PSAs
            If tmChf.sType = "S" And Not (tmCntTypes.iPSA) Then             '8-22-16 chg test from status to type
                ilOk = False
            End If
            'Promo
            If tmChf.sType = "M" And Not (tmCntTypes.iPromo) Then           '8-22-16 chg test from status to type
                ilOk = False
            End If
            
            'include partial trades
            If tmChf.iPctTrade = 100 And Not (tmCntTypes.iTrade) Then
                ilOk = False
            End If
            
            'Political
            ilIsItPolitical = gIsItPolitical(tmChf.iAdfCode)           'its a political, include this contract?
            If ilIsItPolitical Then
                If Not (tmCntTypes.iPolit) Then                   'its a political
                    ilOk = False
                End If
            Else
                If Not tmCntTypes.iNonPolit Then              'include non politicals?
                    ilOk = False           'no, exclude them
                End If
            End If

            
        End If

        mFilterContract = ilOk
        Exit Function
End Function
'
'               mFilterVehicle - filter vehicle and vehicle group for selection
'               <input> ilVefCode : vehicle code
'               <return> true if valid, else false to ignore
Private Function mFilterVehicle(ilVefCode As Integer) As Integer
Dim ilOk As Integer
Dim ilTemp As Integer

        ilOk = True
        If Not gFilterLists(ilVefCode, imInclVefCodes, imUsevefcodes()) Then
            ilOk = False
'            Else                'valid vehicle, see if there is a vehicle group for filtering
'                gGetVehGrpSets ilVefCode, 0, imMajorSet, ilTemp, imVG   'ilTemp = minor sort code(unused), ilMajorVehGrp = major sort code
'                If imVG > 0 Then
'                    If Not gFilterLists(imVG, imInclVGCodes, imUseVGCodes()) Then
'                        ilOk = False
'                    End If
'                End If
        End If
        mFilterVehicle = ilOk
End Function
'
'               mSingleContract - determine if single contract # entered
'                                 & retrieve it
'               <return> Single selection contract code
Private Function mSingleContract() As Long
Dim llChfCode As Long
Dim ilFoundCntr As Integer
Dim ilRet As Integer

        llChfCode = -1
        'determine if there is a single contract to retrieve
        ilFoundCntr = False
        If lmSingleCntr > 0 Then            'get the contracts internal code
            tmChfSrchKey1.lCntrNo = lmSingleCntr
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = lmSingleCntr)
                If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.sDelete <> "Y") Then
                    ilFoundCntr = True
                    llChfCode = tmChf.lCode
                    Exit Do
                End If
                ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If Not ilFoundCntr Then
                mCloseSalesBOFiles
                mSingleContract = llChfCode
                Exit Function
            End If
        Else
            llChfCode = 0
        End If
        mSingleContract = llChfCode
End Function
'
'                   mGetCodesFromList - build array of each of the list boxes
'                   for faster testing
'                   Codes are built in arrays for advertisers, agencies, vehicles,
'                   salespeople, product protection codes, business categories, & vehicle groups
Private Sub mGetCodesFromList(ilVehiclesToProcess() As Integer)
Dim ilLoop As Integer
Dim slNameCode As String
Dim slCode As String
Dim ilRet As Integer
Dim ilIndex As Integer

        'setup array of codes to include or exclude, which is less for speed
        gObtainCodesForMultipleLists 1, tgAdvertiser(), imInclAdvtCodes, imUseAdvtCodes(), RptSelBO
        
        'ReDim ilVehiclesToProcess(1 To 1) As Integer
        ReDim ilVehiclesToProcess(0 To 0) As Integer
        'build array of vehicles to include or exclude
        gObtainCodesForMultipleLists 0, tgVehicle(), imInclVefCodes, imUsevefcodes(), RptSelBO
        'this array is for air time and rep vehicles
        For ilLoop = 0 To RptSelBO!lbcSelection(0).ListCount - 1 Step 1
            slNameCode = tgVehicle(ilLoop).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            'need to process all vehicles to insure that the revenue is for all advertisers
            'to determine if the current years advt is new business
            'i.e. last year advt last aired in Feb
            'this year advt aired in vehicle A in Jan and VEhicle B in Aug.
            'if only vehicle B were selected, it might considere the advt new; but including all
            'vehicles would see that Vehicle A had airing in Jan; thus not making that a new advt
            'If RptSelBO!lbcSelection(0).Selected(ilLoop) Then               'selected ?
                ilIndex = gBinarySearchVef(Val(slCode))
                If ilIndex <> -1 Then
                    If tgMVef(ilIndex).sType = "C" Or tgMVef(ilIndex).sType = "G" Or tgMVef(ilIndex).sType = "S" Then
                        ilVehiclesToProcess(UBound(ilVehiclesToProcess)) = Val(slCode)
                        'ReDim Preserve ilVehiclesToProcess(1 To UBound(ilVehiclesToProcess) + 1)
                        ReDim Preserve ilVehiclesToProcess(0 To UBound(ilVehiclesToProcess) + 1)
                    End If
                End If
            'End If
        Next ilLoop
        
        Exit Sub
End Sub

'
'                   mCalcSplits - calculate the cash/trade split, split slsp and net amts
'                   if applicable
'                   Write prepass record to GRF
'       Grf Prepass variables
'       grfGenDate - generation date for filter
'       grfGenTime - generation time for filter
'       grfvefCode - vehicle code
'       grfslfcode - salesperson code (to get office & sales source)
'       grfrdfcode - agency code (to determine direct or commissionable)
'       grfadfcode - advt code
'       grfSofCode = NTR Item type (mnf)
'       grfChfCode - contract # not code
'       grfCode2   - 1 = new, 2 = renewal business
'       grfdatetype - C (cash), T (trade), Z = Hard Cost to keep it separated
'       grfPerGenl(1) - 0 = direct, 1 = agy, 2 = NTR, 3 = polit, 4 = H/C
'       grfDollars(1 - 14) 12 months
'       grfDollars(15) - total year or quarter
Private Sub mUpdateGRFForAll()
'Dim llAmt(1 To 14) As Long
Dim llAmt(0 To 14) As Long  'Index zero ignored
Dim ilLoop As Integer
Dim slCashAgyComm As String
Dim slStr As String
Dim ilLoopOnPer As Integer
Dim slAmount As String
Dim slSharePct As String
Dim slPctTrade As String
Dim ilCorT As Integer
Dim ilCash As Integer
Dim ilTrade As Integer
Dim slCode As String
Dim slDollar As String
Dim slNet As String
Dim ilLo As Integer
Dim ilHi As Integer
Dim ilLoopOnSlsp As Integer
Dim llTotalAll As Long
Dim ilRet As Integer
Dim llLoopOnStats
Dim ilMonthsOff As Integer
Dim ilTemp As Integer
Dim ilAdfInx As Integer
Dim llDate As Long
Dim slDate As String
Dim slMonth As String
Dim slDay As String
Dim slYear As String

Dim ilMonthsOffIsNew As Integer
Dim ilMonthsNewIsNew As Integer

        'tmSpotAndRev array contains vehicle $ information Air Time, NTR and Rep if requested
        For llLoopOnStats = LBound(tmSpotAndRev) To UBound(tmSpotAndRev) - 1
            tmGrf.iVefCode = tmSpotAndRev(llLoopOnStats).iVefCode
            tmGrf.iAdfCode = tmSpotAndRev(llLoopOnStats).iAdfCode

            ilAdfInx = gBinarySearchNewBus(tmGrf.iAdfCode)      'search array of advt aired for this year/last year
            '                                                   to determine if its new or renewal
            If ilAdfInx <> -1 Then                               'did this advertiser ever air whether last year or this year?
                'found an advertiser
                 If tmNewAdvtBus(ilAdfInx).lLYMonthLastAired = 0 Then        'not aired within the previous year, no $
                    tmGrf.iCode2 = 1                                'no $ last year, its got to be new
                    tmNewAdvtBus(ilAdfInx).iUpdateAdvt = True        'need to update this advt at end of processing
                    slDate = Format$(tmNewAdvtBus(ilAdfInx).lTYMonthFirstAired + 15, "m/d/yy")
                    gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                    tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(0) = Val(slMonth)          ' NEW business for advt
                    tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(1) = Val(slYear)
                Else        'something aired last year, but was it not airing long enough to make it new?
If tmGrf.iAdfCode = 739 Then
 ilTemp = ilTemp
 End If
                    'something found last year, if using Calendar year vs Rolling year,
                    'its automatically a renewal
                    If smNewBusYearType = "C" Then          'calendar year
                        tmGrf.iCode2 = 2                    'renewal, not within the # of months off
                    Else
                       slDate = Format(tmNewAdvtBus(ilAdfInx).lLYMonthLastAired, "m/d/yy")
                       'determine what date this years airing has to be in order to be New
                       For ilMonthsOffIsNew = 1 To imMnthsOffForNew + 1
                           'last date of airing in NewBus table is the start date of the last month advt aired last year
                           'previously converted to std, cal, or corp date
                           llDate = mGetPerTypeStartDatesForNew(slDate)
                       Next ilMonthsOffIsNew
                       If llDate < tmNewAdvtBus(ilAdfInx).lTYMonthFirstAired Then       'this years earliest airing is beyond the calculated # of months off, so its New
                           'New only if its never been flagged as new, or its still New based on # months New is New (site)
                           If tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(0) <> 0 Then     'has the advertiser ever been set to New (setting month & year)
                               slDate = str$(tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(0)) + "/15/" & Trim$(str(tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(1)))
                               For ilMonthsNewIsNew = 1 To imMnthsNewIsNew
                                   'last date of airing in NewBus table is the start date of the last month advt aired last year
                                   'previously converted to std, cal, or corp date
                                   llDate = mGetPerTypeStartDatesForNew(slDate)
                               Next ilMonthsNewIsNew
                               'if this years first date of airing found is < than the # of months New is New (based on first time it was defined as New in Advt)
                               'its still new
                               If tmNewAdvtBus(ilAdfInx).lTYMonthFirstAired < llDate Then        'still new
                                   tmGrf.iCode2 = 1
                               Else
                                   tmGrf.iCode2 = 2                    'renewal, not within the # of months off
                               End If
                           Else            '# months off considered new
                               tmGrf.iCode2 = 1
                               tmNewAdvtBus(ilAdfInx).iUpdateAdvt = True     'need to update advt with year/month found new at end of processing
                               'set the Advertiser array for first time advt detected as New
                               'need to get the true month and year
                               'get the # of first aired month and convert to year/month
                               'convert to month/day/year from the middle of that month to get the correct month #
                               'lldate is the start dateof the following first month of airing; need to go back one month
                               llDate = llDate - 15           'get to themiddle of the previous month so the actual month # can be obtained
                               slDate = Format$(llDate, "m/d/yy")
                               gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
                               tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(0) = Val(slMonth)
                               tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(1) = Val(slYear)
                           End If
                       Else        '# of months off not within "New" status, but could still be new based on # months New is New (site)
                           'first determine if it was ever flagged as New, if not, still not new
    If tmGrf.iAdfCode = 958 Then
    ilTemp = ilTemp
    End If
                           If tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(0) <> 0 Then     'has the advertiser ever been set to New (setting month & year)
                               slDate = str$(tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(0)) + "/15/" & Trim$(str(tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(1)))
                               For ilMonthsNewIsNew = 1 To imMnthsNewIsNew
                                   'last date of airing in NewBus table is the start date of the last month advt aired last year
                                   'previously converted to std, cal, or corp date
                                   llDate = mGetPerTypeStartDatesForNew(slDate)
                               Next ilMonthsNewIsNew
                               'if this years first date of airing found is < than the # of months New is New (based on first time it was defined as New in Advt)
                               'its still new
                               If tmNewAdvtBus(ilAdfInx).lTYMonthFirstAired < llDate Then        'still new
                                   tmGrf.iCode2 = 1
                               Else
                                   tmGrf.iCode2 = 2                    'renewal, not within the # of months off
                               End If
                           Else
                               tmGrf.iCode2 = 2                        'renewal
                           End If                      'tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(0) <> 0
                       End If                          'llDate < tmNewAdvtBus(ilAdfInx).lTYMonthFirstAired
                    End If                          'smNewBusYearType = "C"
                End If                              'tmNewAdvtBus(ilAdfInx).lLYMonthLastAired = 0
            Else    'no advt found, strange case since All advt both past & current year should have been created in NewBus Array
                    'consider this advt new since it was not found
                
                tmGrf.iCode2 = 1            'flag as new business
                'create entry for the advt so that the last month/year considered new can be updated into the ADF table when all is done
                ilAdfInx = UBound(tmNewAdvtBus)
                tmNewAdvtBus(ilAdfInx).iAdfCode = tmSpotAndRev(llLoopOnStats).iAdfCode
                'find the adv to get the last month/year considered New
                ilTemp = gBinarySearchAdf(tmSpotAndRev(llLoopOnStats).iAdfCode)
                If ilTemp <> -1 Then
               
'                    tmNewAdvtBus(ilTemp).iLastMonthYearNew(0) = tgCommAdf(ilTemp).iLastMonthNew
'                    tmNewAdvtBus(ilTemp).iLastMonthYearNew(1) = tgCommAdf(ilTemp).iLastYearNew
                    tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(0) = tgCommAdf(ilTemp).iLastMonthNew
                    tmNewAdvtBus(ilAdfInx).iLastMonthYearNew(1) = tgCommAdf(ilTemp).iLastYearNew

                Else
                    'no advt found
                    
                End If
                
                tmNewAdvtBus(ilAdfInx).iUpdateAdvt = True     'need to update advt with year/month found new at end of processing

                slStr = Trim$(str$(tmSpotAndRev(llLoopOnStats).iAdfCode))
                Do While Len(slStr) < 5
                    slStr = "0" & slStr
                Loop
                tmNewAdvtBus(ilAdfInx).sKey = slStr
                tmNewAdvtBus(ilAdfInx).lLYMonthLastAired = 0
                tmNewAdvtBus(ilAdfInx).lTYMonthFirstAired = 0
                'update first month airing index of $ aired for current year
                For ilLoop = 1 To imPeriods
                    If tmSpotAndRev(llLoopOnStats).lRev(ilLoop - 1) > 0 Then         'keep track of the earliest month advt aired
                        If tmNewAdvtBus(ilAdfInx).lTYMonthFirstAired < lmStartDates(ilLoop) Or tmNewAdvtBus(ilAdfInx).lTYMonthFirstAired = 0 Then
                            tmNewAdvtBus(ilAdfInx).lTYMonthFirstAired = lmStartDates(ilLoop)
                        End If
                        Exit For
                    End If
                Next ilLoop
                ReDim Preserve tmNewAdvtBus(0 To ilAdfInx + 1) As NEWADVTBUS
            End If
            
            'create the prepass record
            tmGrf.iRdfCode = tmSpotAndRev(llLoopOnStats).iAgfCode
            tmGrf.lChfCode = tmSpotAndRev(llLoopOnStats).lCntrNo
        
            ilLo = 1
            If tmCntTypes.iTrade Then
                ilHi = 2            'include trades
            Else
                ilLo = 1            'only cash
            End If
            
            slCashAgyComm = gIntToStrDec(tmSpotAndRev(llLoopOnStats).iAgyCommPct, 2)
            slPctTrade = gIntToStrDec(tmSpotAndRev(llLoopOnStats).iPctTrade, 0)
            
            For ilLoop = 0 To 9
                imSlspSplitCodes(ilLoop) = tmSpotAndRev(llLoopOnStats).iSlfCode(ilLoop)
                lmSlspSplitPct(ilLoop) = tmSpotAndRev(llLoopOnStats).lComm(ilLoop)
            Next ilLoop
    
            For ilLoopOnSlsp = 0 To 9
                slSharePct = gLongToStrDec(lmSlspSplitPct(ilLoopOnSlsp), 4)       'slsp share
                If imSlspSplitCodes(ilLoopOnSlsp) > 0 Then
                    For ilCorT = ilLo To ilHi                   'loop for cash & trade (if applicable)
                        For ilLoopOnPer = 1 To imPeriods          'init the $ table
                            llAmt(ilLoopOnPer) = 0
                        Next ilLoopOnPer
                        llTotalAll = 0
                        For ilLoopOnPer = 1 To imPeriods
                            If imGrossNetSpot = 3 Then          'spots not fully implemented
                                slAmount = gLongToStrDec(tmSpotAndRev(llLoopOnStats).lSpots(ilLoopOnPer - 1), 0) 'spot count gathered from spots
                            Else
                                slAmount = gLongToStrDec(tmSpotAndRev(llLoopOnStats).lRev(ilLoopOnPer - 1), 2) '$ gathered from spots
                            End If
                            slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
                            slStr = gRoundStr(slStr, "1", 0)
                
                            If ilCorT = 1 Then
                                slCode = gSubStr("100.", slPctTrade)
                                slDollar = gDivStr(gMulStr(slStr, slCode), "100")              'slsp gross
                                slDollar = gRoundStr(slDollar, "1", 0)
                                slNet = gRoundStr(gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "100.00"), ".01", 0)
                                If imGrossNetSpot = 1 Or imGrossNetSpot = 3 Then          'gross or spot counts
                                    llAmt(ilLoopOnPer) = Val(slDollar)
                                    llTotalAll = llTotalAll + Val(slDollar)
                                ElseIf imGrossNetSpot = 2 Then      'net
                                    llAmt(ilLoopOnPer) = Val(slNet)
                                    llTotalAll = llTotalAll + Val(slNet)
                                End If
                                tmGrf.sDateType = "C"
                            Else
                                If ilCorT = 2 Then                'at least cash is commissionable
                                    slCode = gIntToStrDec(tmSpotAndRev(llLoopOnStats).iPctTrade, 0)
                                    slDollar = gDivStr(gMulStr(slStr, slCode), "100")
                                    slDollar = gRoundStr(slDollar, "1", 0)
                                    If tmSpotAndRev(llLoopOnStats).iAgfCode > 0 And tmSpotAndRev(llLoopOnStats).sTradeComm = "Y" Then
                                        slNet = gRoundStr(gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "100.00"), "1", 0)
                                    Else
                                        slNet = slDollar    'no commission , net is same as gross
                                    End If
                                    If imGrossNetSpot = 1 Or imGrossNetSpot = 3 Then          'gross or spot counts
                                        llAmt(ilLoopOnPer) = Val(slDollar)
                                        llTotalAll = llTotalAll + Val(slDollar)
                                    Else
                                        llAmt(ilLoopOnPer) = Val(slNet)
                                        llTotalAll = llTotalAll + Val(slNet)
                                    End If
                                    tmGrf.sDateType = "T"
                                End If
                            End If
                        Next ilLoopOnPer
                        If llTotalAll <> 0 Then         'dont create a prepass record if no value
                            'tmGrf.lDollars(15) = llTotalAll
                            tmGrf.lDollars(14) = llTotalAll
                            For ilLoop = 1 To imPeriods
                                tmGrf.lDollars(ilLoop - 1) = llAmt(ilLoop)
                            Next ilLoop
                            If tmSpotAndRev(llLoopOnStats).sIsNTR = "Y" And tmSpotAndRev(llLoopOnStats).iIsItHardCost = True Then
                                tmGrf.sDateType = "Z"       'sort to end
                            End If
                            tmGrf.iSlfCode = imSlspSplitCodes(ilLoopOnSlsp)        'assume no splits, use 1st slsp
                            'tmGrf.iPerGenl(1) = 1       'assume agy commissionable
                            tmGrf.iPerGenl(0) = 1       'assume agy commissionable
                            If tmGrf.iRdfCode = 0 Then      'is there an agency
                                'tmGrf.iPerGenl(1) = 0   'direct
                                tmGrf.iPerGenl(0) = 0   'direct
                            End If
                            If tmSpotAndRev(llLoopOnStats).sIsNTR = "Y" Then           'its NTR
                                If tmSpotAndRev(llLoopOnStats).iIsItHardCost = True Then   'Hard Cost
                                    'tmGrf.iPerGenl(1) = 4
                                    tmGrf.iPerGenl(0) = 4
                                Else
                                    'tmGrf.iPerGenl(1) = 2                       'NTR
                                    tmGrf.iPerGenl(0) = 2                       'NTR
                                End If
                            End If
                            If tmSpotAndRev(llLoopOnStats).iIsItPolitical = True Then
                                'tmGrf.iPerGenl(1) = 3                            'political, separate it out
                                tmGrf.iPerGenl(0) = 3                            'political, separate it out
                            End If
                            
                            tmGrf.iSofCode = tmSpotAndRev(llLoopOnStats).iNTRMnfType
                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                            If ilRet <> BTRV_ERR_NONE Then
                                igBtrError = gConvertErrorCode(ilRet)
                                sgErrLoc = "mCalcSplits-Insert GRF"
                                Exit Sub
                            End If
                        End If
                    Next ilCorT
                    
                Else
                    Exit For
                End If              'imSlspSplitCodes(ilLoopOnSlsp) > 0
            Next ilLoopOnSlsp       'For ilLoopOnSlsp = 0 To 9
            
        Next llLoopOnStats          'llLoopOnStats = LBound(tmSpotAndRev) To UBound(tmSpotAndRev) - 1
     
        Exit Sub
End Sub
'
'                   mGetRateAndAddToArray - get flight spot rate then add to
'                   the contracts entry in tmSpotAndRev array
'                   <Input>  index to period processing
'                   Return - Spot Rate, -1 if some error and flight not found
Private Function mGetRateAndAddToArray(ilDateInx As Integer) As Long
Dim slSpotRate As String
Dim llSpotRate As Long
Dim ilFoundCntr As Integer
Dim llLoopOnStats As Long
Dim ilIndex As Integer
Dim ilRet As Integer
Dim ilUpper As Integer
Dim ilLoop As Integer
Dim ilIsItPolitical As Integer

        'flight could be different so need to get the spot rate for every spot
        If imGrossNetSpot < 3 Then          'gross or net, need $
            ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slSpotRate)
            If (InStr(slSpotRate, ".") <> 0) Then        'found spot cost
                'is it a .00?
                If gCompNumberStr(slSpotRate, "0.00") = 0 Then       'its a .00 spot
                    llSpotRate = 0
                Else
                    llSpotRate = gStrDecToLong(slSpotRate, 2)
                End If
            Else
                'its a bonus, recap, n/c, etc. which is still $0
                llSpotRate = 0
            End If
            
        End If
        ilFoundCntr = False
          
        If imGrossNetSpot = 3 Or llSpotRate <> 0 Then       'if doing spot counts (imgrossnetsdpot = 3) or spot rate is non-zero, update the arrays
            'accumulate $ in table with matching contract & line
            For llLoopOnStats = LBound(tmSpotAndRev) To UBound(tmSpotAndRev) - 1
                If tmSdf.lChfCode = tmSpotAndRev(llLoopOnStats).lChfCode And tmSdf.iLineNo = tmSpotAndRev(llLoopOnStats).iLineNo Then
                    tmSpotAndRev(llLoopOnStats).lSpots(ilDateInx - 1) = tmSpotAndRev(llLoopOnStats).lSpots(ilDateInx - 1) + 1
                    tmSpotAndRev(llLoopOnStats).lRev(ilDateInx - 1) = tmSpotAndRev(llLoopOnStats).lRev(ilDateInx - 1) + llSpotRate
                    ilFoundCntr = True
                    Exit For
                End If
            Next llLoopOnStats
            If Not ilFoundCntr Then     'contract/line not found, place in table with spot count
                ilUpper = UBound(tmSpotAndRev)
                tmSpotAndRev(ilUpper).iAgyCommPct = 0      'direct, no comm
                If tmChf.iAgfCode > 0 Then
                    ilIndex = gBinarySearchAgf(tmChf.iAgfCode)
                    If ilIndex <> -1 Then
                        'tmSpotAndRev(ilUpper).iAgyCommPct = 1500
                        tmSpotAndRev(ilUpper).iAgyCommPct = tgCommAgf(ilIndex).iCommPct
                     End If
                End If
                tmSpotAndRev(ilUpper).lChfCode = tmSdf.lChfCode
                tmSpotAndRev(ilUpper).lCntrNo = tmChf.lCntrNo
                tmSpotAndRev(ilUpper).iVefCode = tmSdf.iVefCode
                tmSpotAndRev(ilUpper).iAdfCode = tmChf.iAdfCode
                tmSpotAndRev(ilUpper).iAgfCode = tmChf.iAgfCode
                tmSpotAndRev(ilUpper).iPctTrade = tmChf.iPctTrade
                tmSpotAndRev(ilUpper).sTradeComm = tmChf.sAgyCTrade       'agy commissionable for trades
                If tmChf.iPctTrade = 0 Then
                    tmSpotAndRev(ilUpper).sCashTrade = "C"
                ElseIf tmChf.iPctTrade = 100 Then
                    tmSpotAndRev(ilUpper).sCashTrade = "T"
                Else
                    tmSpotAndRev(ilUpper).sCashTrade = "S"       'split cash trade
                End If
                tmSpotAndRev(ilUpper).iMnfComp = tmChf.iMnfComp(0)      'agy commissionable for trades
                tmSpotAndRev(ilUpper).iMnfRevSet = tmChf.iMnfRevSet(0)       'agy commissionable for trades
                tmSpotAndRev(ilUpper).iLineNo = tmSdf.iLineNo
                tmSpotAndRev(ilUpper).sIsNTR = "N"                      'not NTR
                'Political 10-18-12
                ilIsItPolitical = gIsItPolitical(tmChf.iAdfCode)           'its a political, include this contract?
                If ilIsItPolitical Then
                    tmSpotAndRev(ilUpper).iIsItPolitical = True
                Else
                    tmSpotAndRev(ilUpper).iIsItPolitical = False
                End If
                
                tmSpotAndRev(ilUpper).lSpots(ilDateInx - 1) = 1
                tmSpotAndRev(ilUpper).lRev(ilDateInx - 1) = llSpotRate
                mCreateSlspSplitTable                   'build only selected slsp into split table
                For ilLoop = 0 To 9
                    tmSpotAndRev(ilUpper).lComm(ilLoop) = lmSlspSplitPct(ilLoop)
                    tmSpotAndRev(ilUpper).iSlfCode(ilLoop) = imSlspSplitCodes(ilLoop)
                Next ilLoop
                
                ReDim Preserve tmSpotAndRev(LBound(tmSpotAndRev) To ilUpper + 1) As SPOTBBSTATS
            End If
            
            mGetRateAndAddToArray = llSpotRate
        End If
        Exit Function
End Function
'
'                   mSlspSplitTable -setup the slsp split table based on the selected slsp
'                   or if not splitting, everything goes to first slsp
'                   Assume tmChf contract in memory
'
Private Sub mCreateSlspSplitTable()
Dim ilLoop As Integer
Dim ilIndex As Integer
Dim ilValidCount As Integer

        For ilLoop = 0 To 9
            imSlspSplitCodes(ilLoop) = 0
            lmSlspSplitPct(ilLoop) = 0
        Next ilLoop
        ilValidCount = 0
        
        'If Not imSlspSplit Then
            imSlspSplitCodes(0) = tmChf.iSlfCode(0)
            lmSlspSplitPct(0) = 1000000         '100.0000
'        Else
'            For ilLoop = 0 To 9             'loop thru the contract header slsp and determine which ones to include based on selectivity
'                If gFilterLists(tmChf.iSlfCode(ilLoop), imInclSlfCodes, imUseSlfCodes()) Then
'                    'found valid one
'                    imSlspSplitCodes(ilValidCount) = tmChf.iSlfCode(ilLoop)
'                    lmSlspSplitPct(ilValidCount) = tmChf.lComm(ilLoop)
'                    ilValidCount = ilValidCount + 1
'                End If
'            Next ilLoop
'        End If
        Exit Sub
End Sub
'
'            Process rep and/or AirTime for this year/last year (separate passes)
'
'           <input> llChfCode - single contract # or 0 if all
'                   slEarliestDate - earliest data to gather
'                   slLatestDate - latest data to gather
'                   llStartDates - array of period start dates for as many periods requested
'                   ilThisYear   -true if current year, else false for last year
'                   ilPeriods - # periods to process, last year may be different than current year
'
'           Do not filter any contracts (air time/rep) to determine what aired in previous year so
'           the New status can be determined based on what aired
Private Sub mGatherAdvtAiringFromContract(llChfCode As Long, slEarliestDate As String, slLatestDate As String, llStartDates() As Long, ilThisYear As Integer, ilPeriods As Integer)

Dim ilLoop As Integer
ReDim ilRepVehicles(0 To 0) As Integer
Dim ilHOState As Integer
Dim ilAdjustDays As Integer
Dim ilValidDays(0 To 6) As Integer
Dim ilLoopOnKey As Integer
Dim ilRet As Integer
Dim llContrCode As Long
Dim ilClf As Integer
Dim ilTemp As Integer
Dim ilIncludeVehicle As Integer
Dim ilByCodeOrNumber As Integer
Dim ilOk As Integer
Dim ilfirstTime As Integer
Dim ilFoundVef As Integer
Dim ilIndex As Integer
Dim ilWeekOrMonth As Integer
Dim llTotalGross As Long
Dim slStr As String
Dim ilIsItPolitical As Integer
'ReDim tmChfAdvtExt(1 To 1) As CHFADVTEXT
ReDim tmChfAdvtExt(0 To 0) As CHFADVTEXT
'ReDim imDormantVehicles(1 To 1) As Integer          'only needed for this year processing to see if any of the vehicle lines are dormant
ReDim imDormantVehicles(0 To 0) As Integer          'only needed for this year processing to see if any of the vehicle lines are dormant
                                                    'may need to process them for to get spot data

Dim ilWhichRate As Integer
Dim slCntrStatus As String
Dim slCntrType As String

            slCntrStatus = "HOGN"                   'get sch/uns holds and orders to determine what aired in the year
            slCntrType = ""                         'get all types (standard, reservation, PI, Promo, etc ; excl psa/promo)
            If llChfCode > 0 Then                   'single contract entered?
                'tmChfAdvtExt(1).lCode = tmChf.lCode
                'ReDim Preserve tmChfAdvtExt(1 To 2) As CHFADVTEXT
                tmChfAdvtExt(0).lCode = tmChf.lCode
                ReDim Preserve tmChfAdvtExt(0 To 1) As CHFADVTEXT
            Else                                    'gather all contracts for period (last year or this year)
                ilHOState = 2
                ilRet = gObtainCntrForDate(RptSelBO, slEarliestDate, slLatestDate, slCntrStatus, slCntrType, ilHOState, tmChfAdvtExt())
            End If
            
            ilAdjustDays = (llStartDates(ilPeriods + 1) - llStartDates(1)) + 1
            ReDim lmCalSpots(0 To ilAdjustDays) As Long        'init buckets for daily calendar values (spots unused in this report)
            ReDim lmCalAmt(0 To ilAdjustDays) As Long
            ReDim lmCalAcqAmt(0 To ilAdjustDays) As Long            'acq unused in this reprot
            ReDim lmAcquistion(0 To ilAdjustDays) As Long
            For ilLoop = 0 To 6                         'days of the week
                ilValidDays(ilLoop) = True              'force alldays as valid
            Next ilLoop

            ilfirstTime = True
            For ilLoopOnKey = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1
                llContrCode = tmChfAdvtExt(ilLoopOnKey).lCode
                'obtain contract header and flights
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tmChf, tgClfCT(), tgCffCT())

                mCreateSlspSplitTable           'determine if any split slsp
                For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                    tmClf = tgClfCT(ilClf).ClfRec
                    
                    'test required when air time and rep are allowed on the same contract
                    ilRet = gBinarySearchVef(tmClf.iVefCode)
                    If ilRet <> -1 Then
                        ilByCodeOrNumber = 0            'reference by chf code
                        ilOk = mFilterContract(ilByCodeOrNumber, tmChf.lCode, ilfirstTime, ilIsItPolitical)
                        ilfirstTime = False
                        If ilOk Then
                            If Not gFilterLists(tmChf.iAdfCode, imInclAdvtCodes, imUseAdvtCodes()) Then
                                ilOk = False
                            End If
                        End If
                        If ilOk Then
                            If ilThisYear Then          'only need to gather dormant vehicles for this year to get spot data if necessary
                                If tgMVef(ilRet).sState = "D" Then          'ilret is index into global vehicle array (from gBinarySearchVef)
                                    gAddDormantVehicle tmClf.iVefCode, imDormantVehicles
                                End If
                            End If
                            For ilTemp = 1 To ilPeriods Step 1 'init projection $ each time
                                lmProject(ilTemp) = 0
                                lmAcquisition(ilTemp) = 0
                                lmProjectSpots(ilTemp) = 0
                            Next ilTemp
                             'init the cal buckets, if used
                            For ilTemp = 0 To UBound(lmCalSpots) - 1
                                lmCalSpots(ilTemp) = 0        'init buckets for daily calendar values
                                lmCalAmt(ilTemp) = 0
                                lmCalAcqAmt(ilTemp) = 0
                            Next ilTemp
                            'use hidden lines
                            If tmClf.sType <> "A" And tmClf.sType <> "O" And tmClf.sType <> "E" Then
                                'see if now a dormant vehicle;  if so, and All Vehicles checked, need to include it in spot processing
                                
                                If imPerType = 2 Or imPerType = 3 Or imPerType = 1 Then       'std or corporate, or weekly
                                    ilWeekOrMonth = 1
                                    If imPerType = 1 Then           'weekly
                                        ilWeekOrMonth = 2
                                    End If
                                    'gBuildFlights ilClf, llStdStartDates(), 1, imPeriods + 1, lmProject(), ilWeekOrMonth, tgClfCT(), tgCffCT()
                                    ilWhichRate = 0     '0 = use spot rate, 1 = use acq rate, 2 = if acq non-zero, use it.  otherwise if 0, default to line rate
                                    gBuildFlightSpotsAndRevenue ilClf, llStartDates(), 1, ilPeriods + 1, lmProject(), lmProjectSpots(), ilWeekOrMonth, ilWhichRate, tgClfCT(), tgCffCT()

                                ElseIf imPerType = 4 Then                   'calendar
                                    gCalendarFlights tgClfCT(ilClf), tgCffCT(), llStartDates(1), llStartDates(ilPeriods + 1), ilValidDays(), True, lmCalAmt(), lmCalSpots(), lmCalAcqAmt(), tmPriceTypes
                                    gAccumCalFromDays llStartDates(), lmCalAmt(), lmCalAcqAmt(), False, lmProject(), lmAcquisition(), ilPeriods
                                    gAccumCalSpotsFromDays llStartDates(), lmCalSpots(), lmProjectSpots(), ilPeriods
                                End If
                            End If
                            
                            mCreateAdvtAiredArrayForAll ilThisYear, ilPeriods    'setup table of what advt aired for last year/this year
                
                        End If              'veftype = "R"
                    End If                  'ilREt >= 0
                Next ilClf
            Next ilLoopOnKey
            sgCntrForDateStamp = ""         'reset for another run
 
            Erase tmChfAdvtExt
            Erase lmCalSpots, lmCalAmt, lmCalAcqAmt, lmAcquisition
        Exit Sub
End Sub
'
'                   mGatherAdvtForNTR - find all NTR for last year/this year and
'                   continue building array of what aired
'                   <input> llChfCode - single contract # or 0 if all
'                       slEarliestDate - earliest data to gather
'                       slLatestDate - latest data to gather
'                       llStartDates - array of period start dates for as many periods requested
'                       ilThisYear   -true if current year, else false for last year
'                       ilPeriods - # periods to process, last year may be different than current year
'
'           Do not filter any contracts/vehicles to determine what aired in previous year so
'           the New status can be determined based on what aired
'
Private Sub mGatherAdvtForNTR(llChfCode As Long, slEarliestDate As String, slLatestDate As String, llStartDates() As Long, ilThisYear As Integer, ilPeriods As Integer)
Dim slDate As String
Dim llDate As Long
Dim ilMonthInx As Integer
Dim ilFoundMonth As Integer
Dim ilFoundVef As Integer
Dim ilTemp As Integer
Dim llSBFLoop As Long               '12-2-16 chg from int to long
Dim ilIsItHardCost As Integer
Dim ilFoundOption As Integer
Dim ilWhichKey As Integer
Dim llEarliestDate As Long
Dim llLatestDate As Long
Dim ilAgyComm As Integer
Dim ilIndex As Integer
Dim ilRet As Integer
Dim ilByCodeOrNumber As Integer
Dim ilOk As Integer
Dim ilIncludeVehicle As Integer
Dim ilUpper As Integer
Dim ilfirstTime As Integer
Dim ilLoop As Integer
Dim ilIsItPolitical As Integer
ReDim tmSbfList(0 To 0) As SBF


        llEarliestDate = gDateValue(slEarliestDate)
        llLatestDate = gDateValue(slLatestDate)
        If lmSingleCntr > 0 Then
            ilWhichKey = 0
        Else
            ilWhichKey = 2          'trantype, date
        End If

        ilRet = gObtainSBF(RptSelBO, hmSbf, llChfCode, slEarliestDate, slLatestDate, tmNTRTypes, tmSbfList(), ilWhichKey)

        ilfirstTime = True
        For llSBFLoop = LBound(tmSbfList) To UBound(tmSbfList) - 1
            tmSbf = tmSbfList(llSBFLoop)
            gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), slDate
            llDate = gDateValue(slDate)
                
            ilFoundOption = True
            ilIsItHardCost = gIsItHardCost(tmSbf.iMnfItem, tgNTRMnf())
    
            If ilIsItHardCost Then              'hard cost item
                If Not tmCntTypes.iHardCost Then
                    ilFoundOption = False
                End If
            Else                                'normal NTR
                If Not tmCntTypes.iNTR Then
                    ilFoundOption = False
                End If
            End If
    
            If ilFoundOption Then
                ilFoundMonth = False
                For ilMonthInx = 1 To ilPeriods Step 1         'loop thru months to find the match
                    If llDate >= llStartDates(ilMonthInx) And llDate < llStartDates(ilMonthInx + 1) Then
                        ilFoundMonth = True
                        Exit For
                    End If
                Next ilMonthInx
    
                If ilFoundMonth Then
                    'filter out the type of contract
                    ilByCodeOrNumber = 0            'reference by chf code
                    ilOk = mFilterContract(ilByCodeOrNumber, tmSbf.lChfCode, ilfirstTime, ilIsItPolitical)
                    ilfirstTime = False
                    If ilOk Then
                        If Not gFilterLists(tmChf.iAdfCode, imInclAdvtCodes, imUseAdvtCodes()) Then
                            ilOk = False
                        End If
                    End If
                    If ilOk Then
                        ilFoundVef = False
                        'setup vehicle that spot was moved to
                        'determine agency commission on the individual item
                        ilAgyComm = 0
                        If tmSbf.sAgyComm = "Y" Then
                            'determine the amt of agy commission; can vary per agy
                            If tmChf.iAgfCode > 0 Then
                                ilIndex = gBinarySearchAgf(tmChf.iAgfCode)
                                If ilIndex <> -1 Then
                                    'ilAgyComm = 1500
                                    ilAgyComm = tgCommAgf(ilIndex).iCommPct
                                End If
                            End If
                        End If
                        
                        For ilLoop = 1 To ilPeriods
                            lmProject(ilLoop) = 0
                        Next ilLoop
                        lmProject(ilMonthInx) = (tmSbf.lGross * tmSbf.iNoItems)
                        mCreateAdvtAiredArrayForAll ilThisYear, ilPeriods    'setup table of what advt aired last year (NTR)
                   
                    Else
                        ilTemp = ilTemp
                    End If              'ilok
                End If
            End If
        Next llSBFLoop
        
        Erase tmSbfList
        Exit Sub
End Sub
'
'                   Process History and Receivables Adjustments
'                   Include Trantype = AN only, between all dates requested
'
Private Sub mProcessAdj(slEarliestDate As String, slLatestDate As String)
Dim ilRet As Integer
Dim ilVehicle As Integer
ReDim tlRvf(0 To 0) As RVF
Dim llRvf As Long
Dim ilOk As Integer
Dim slSpotOrNTR As String * 1
Dim ilIndex As Integer
Dim llDate As Long
Dim llAmt As Long
Dim ilNetPct As Integer
Dim llTemp As Long
Dim ilfirstTime As Integer
Dim ilByCodeOrNumber As Integer
Dim ilAgyComm As Integer
Dim ilDateInx As Integer
Dim slCashTrade As String * 1
Dim ilPctTrade As Integer
Dim ilLoop As Integer
Dim ilIncludeVehicle As Integer
Dim ilTemp As Integer
Dim ilFoundVef As Integer
Dim ilVefIndex As Integer
Dim ilIsItPolitical As Integer


        ReDim tmSpotAndRev(0 To 0) As SPOTBBSTATS
        ilRet = gObtainPhfRvf(RptSelBO, slEarliestDate, slLatestDate, tmTranTypes, tlRvf(), 0)
        ilfirstTime = True
        For llRvf = LBound(tlRvf) To UBound(tlRvf) - 1
            tmRvf = tlRvf(llRvf)
            'filter out the type of contract
            ilByCodeOrNumber = 1            'reference by Contract #
            ilOk = mFilterContract(ilByCodeOrNumber, tmRvf.lCntrNo, ilfirstTime, ilIsItPolitical)
            ilfirstTime = False
            If lmSingleCntr > 0 And lmSingleCntr <> tmRvf.lCntrNo Then
                ilOk = False
            End If
            If ilOk Then
                If Not gFilterLists(tmRvf.iAdfCode, imInclAdvtCodes, imUseAdvtCodes()) Then
                    ilOk = False
                End If

                If Trim$(tmRvf.sType) <> "" And tmRvf.sType <> "A" Then     'ANs can only be revenue records, no installment types
                    ilOk = False
                End If
                If tmRvf.iMnfItem > 0 And (Not tmCntTypes.iNTR) Then        'filter out NTR adjustments by user input
                    ilOk = False
                End If
                
                '12-2-16 determine adjustments by type of vehicle (rep and/or non-rep(all other types)
                'REP adjustments are only considered when including adjustments for Air Time types
                'Scheduled spots should not make adjustments
'                If tmRvf.iMnfItem = 0 And (tmCntTypes.iAirTime) Then      'filter out Air Time adjustments by user input, may need to include REP
'                    'determine the type of vehicle this is:  ignore scheduled spots vehicles (not REP)
'                    ilVefIndex = gBinarySearchVef(tmRvf.iAirVefCode)
'                    If ilVefIndex = -1 Then         '8-22-16 valid index?
'                        ilOk = False
'                    Else
'                        If tgMVef(ilVefIndex).sType <> "R" Then      'rep, OK to include; otherwise ignore it
'                            ilOk = False
'                        End If
'                    End If
'                End If

                '12-1-16  Include adj for either Rep or Air Time
                'determine the type of vehicle
                If ilOk Then
                    ilVefIndex = gBinarySearchVef(tmRvf.iAirVefCode)
                    If ilVefIndex = -1 Then
                        ilOk = False
                    Else
                        If tgMVef(ilVefIndex).sType = "R" Then
                            If (tmCntTypes.iRep = False Or RptSelBO!ckcAdj(0).Value = vbUnchecked) Then        'rep vehicle and rep adj excluded
                                ilOk = False
                            End If
                        Else
                            If (tmCntTypes.iAirTime = False Or RptSelBO!ckcAdj(1).Value = vbUnchecked) Then     'air time vehicle and adj excluded
                                ilOk = False
                            End If
                        End If
                    End If
                    ilIncludeVehicle = mFilterVehicle(tmRvf.iAirVefCode)
                End If
            End If
            
            If (ilOk) And (ilIncludeVehicle) Then

                If ilOk Then
                    gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate
                    'Determine bucket
                    For ilDateInx = 1 To imPeriods Step 1
                        If (llDate >= lmStartDates(ilDateInx)) And (llDate <= lmStartDates(ilDateInx + 1)) Then
                            'determine if agy commissionable
                            If tmChf.iAgfCode > 0 Then
                                ilIndex = gBinarySearchAgf(tmChf.iAgfCode)
                                If ilIndex <> -1 Then
                                    'ilAgyComm = 1500
                                    ilAgyComm = tgCommAgf(ilIndex).iCommPct
                                End If
                            Else
                                ilAgyComm = 0
                            End If
                            
                            If tmRvf.sCashTrade = "C" Then
                                ilPctTrade = 0      'this is cash, dont split for trades
                                slCashTrade = "C"     'this is all cash transaction
                            Else
                                ilPctTrade = 100      'this is a trade record, dont split for cash
                                slCashTrade = "T"      'this is all Trade transaction
                            End If
                            
                            gPDNToLong tmRvf.sGross, llAmt
                            'if AN gross is 0, backcompute from net
                            If ilAgyComm = 0 Then          'direct
                                gPDNToLong tmRvf.sNet, llAmt
                            Else
                                If llAmt = 0 Then                     'gross may be 0
                                    If imGrossNetSpot = 2 Then        'if net, need to get a gross amt so it will be processed in the update routine
                                                                    'because the update rtn checked to see if gross is non-zero or not to process
                                                                    'if gross and zero, leave it that way as its intended not to have a gross
                                        'back compute gross
                                        gPDNToLong tmRvf.sNet, llTemp
                                        ilNetPct = 8500         'default 85% net
                                        If ilRet = BTRV_ERR_NONE Then
                                            ilNetPct = (10000 - ilAgyComm)
                                        End If          'ilret = btrv_err_none
                                        llAmt = CDbl(llTemp) * 10000 / ilNetPct
                                    End If
                                End If
                            End If
                            For ilTemp = LBound(tmSpotAndRev) To UBound(tmSpotAndRev) - 1 Step 1
                                If tmSpotAndRev(ilTemp).iVefCode = tmRvf.iAirVefCode And tmSpotAndRev(ilTemp).iAgyCommPct = ilAgyComm And tmSpotAndRev(ilTemp).lChfCode = tmRvf.lCntrNo Then
                                    tmSpotAndRev(ilTemp).lRev(ilDateInx - 1) = tmSpotAndRev(ilTemp).lRev(ilDateInx - 1) + llAmt
                                    ilFoundVef = True
                                    Exit For
                                End If
                            Next ilTemp
                            If Not (ilFoundVef) Then
                                ilTemp = UBound(tmSpotAndRev)
                                tmSpotAndRev(ilTemp).lChfCode = tmChf.lCode
                                tmSpotAndRev(ilTemp).lCntrNo = tmChf.lCntrNo
                                tmSpotAndRev(ilTemp).iAdfCode = tmChf.iAdfCode
                                tmSpotAndRev(ilTemp).iAgfCode = tmChf.iAgfCode
                                tmSpotAndRev(ilTemp).iVefCode = tmRvf.iAirVefCode
                                'Transactions have already been split during invoicing, force so that no splits occur
                                'when creating the prepass record
                                tmSpotAndRev(ilTemp).sCashTrade = slCashTrade
                                tmSpotAndRev(ilTemp).iPctTrade = ilPctTrade
                                
                                tmSpotAndRev(ilTemp).sTradeComm = "N"       'no commissions determined for trades, its already been determined from the transaction
                                tmSpotAndRev(ilTemp).iAgyCommPct = ilAgyComm
                                
                                If tmRvf.iMnfItem > 0 Then          'NTR
                                    tmSpotAndRev(ilTemp).sIsNTR = "Y"                             'flag to indicate its an NTR item
                                    tmSpotAndRev(ilTemp).iIsItHardCost = gIsItHardCost(tmRvf.iMnfItem, tgNTRMnf())  'flag to indicate if hard cost or not
                                Else
                                    tmSpotAndRev(ilTemp).sIsNTR = "N"                             'flag to indicate its an NTR item
                                End If
                                tmSpotAndRev(ilTemp).iMnfComp = tmChf.iMnfComp(0)      'agy commissionable for trades
                                tmSpotAndRev(ilTemp).iMnfRevSet = tmChf.iMnfRevSet(0)       'agy commissionable for trades
                                tmSpotAndRev(ilTemp).iNTRMnfType = tmRvf.iMnfItem
                                mCreateSlspSplitTable                   'build only selected slsp into split table
                                For ilLoop = 0 To 9
                                    tmSpotAndRev(ilTemp).lComm(ilLoop) = lmSlspSplitPct(ilLoop)
                                    tmSpotAndRev(ilTemp).iSlfCode(ilLoop) = imSlspSplitCodes(ilLoop)
                                Next ilLoop

                
                                tmSpotAndRev(ilTemp).lRev(ilDateInx - 1) = llAmt
                                ReDim Preserve tmSpotAndRev(0 To ilTemp + 1) As SPOTBBSTATS
                                Exit For
                            End If      'if not ilfoundvef
                        End If          'if (llDate >= lmStartDates(ilDateInx)) And (llDate <= lmStartDates(ilDateInx))
                    Next ilDateInx
                End If                  'ilok
            End If                      ' If (ilOk) And (ilIncludeVehicle) Then
        Next llRvf
        mUpdateGRFForAll
        Erase tlRvf, tmSpotAndRev
    Exit Sub
End Sub
'
'               Build array of start dates for as many months that is determined
'               to decide how far back is new business
'               <input> ilPerType:  1 = weekly (not implemented), 2 = std, 3 = corp, 4 = cal
'                       ilStartYear = Current year
'                       llStartDates(1) - start date of current year
'               <output> lmLYStartDates() - array of start dates for previous year
Private Sub mBuildLYStartDates(ilPerType As Integer, ilStartYear As Integer, llStartDate As Long)
Dim ilLoop As Integer
Dim ilNoMnthNewBus As Integer
Dim ilNoMnthNewIsNew As Integer
Dim slNewBusYearType As String * 1
Dim llDate As Long
Dim slStart As String
Dim ilTemp As Integer
Dim slDate As String

        
        smMonthsInYear = "JanFebMarAprMayJunJulAugSepOctNovDec"
        
        llDate = llStartDate
    
        If ilPerType = 1 Then           'weekly, not implemented
        ElseIf ilPerType = 2 Then       'std
            If smNewBusYearType = "R" Then          'rolling year, need enough months to adjust for the period indicates # Months New is New
                'ReDim lmLYStartDates(1 To imMnthsOffForNew + 1) As Long
                ReDim lmLYStartDates(0 To imMnthsOffForNew + 1) As Long 'Index zero ignored
                For ilLoop = UBound(lmLYStartDates) To 1 Step -1
                    'llDate = llDate - 1
                    slDate = Format$(llDate, "m/d/yy")
                    slDate = gObtainStartStd(slDate)
                    lmLYStartDates(ilLoop) = gDateValue(slDate)
                    llDate = lmLYStartDates(ilLoop)
                    llDate = llDate - 1
                 Next ilLoop
            Else        'start with beginning of last year calendar
                'ReDim lmLYStartDates(1 To 13) As Long
                ReDim lmLYStartDates(0 To 13) As Long   'Index zero ignored
                slStart = "1/15/" & str$(igYear - 1)
                gBuildStartDates slStart, 1, 13, lmLYStartDates()
            End If
        ElseIf ilPerType = 3 Then       'corp
            smMonthsInYear = gGetCorpMonthString(igYear)     'get the string for the months of the corp cal:  i.e. OctNoveDecJan....Sep
            If smNewBusYearType = "R" Then                  'rolling months Off for new
                'ReDim lmLYStartDates(1 To (imMnthsOffForNew + 1)) As Long
                ReDim lmLYStartDates(0 To (imMnthsOffForNew + 1)) As Long   'Index zero ignored
                slStart = str$(igMonthOrQtr) & "/15/" & str$(igYear - 1)
                gBuildStartDates slStart, 2, UBound(lmLYStartDates), lmLYStartDates()
            Else        'start with beginning of last year calendar
                'ReDim lmLYStartDates(1 To 13) As Long
                ReDim lmLYStartDates(0 To 13) As Long   'Index zero ignored
                ilTemp = gGetCorpCalIndex(igYear - 1)
                'If ilTemp > 0 Then
                If ilTemp >= 0 Then
                    slStart = str$(tgMCof(ilTemp).iStartMnthNo) & "/15/" & str$(igYear - 1)
                    gBuildStartDates slStart, 2, 13, lmLYStartDates()
                Else
                    MsgBox "Calendar year " & str$(igYear - 1) & " not defined"
                    Exit Sub
                End If
            End If
        Else                            'cal
            If smNewBusYearType = "R" Then              'rolling months off for New
                'ReDim lmLYStartDates(1 To (imMnthsOffForNew + 1)) As Long
                ReDim lmLYStartDates(0 To (imMnthsOffForNew + 1)) As Long   'Index zero ignored
                For ilLoop = UBound(lmLYStartDates) To 1 Step -1
                    'llDate = llDate - 1
                    slDate = Format$(llDate, "m/d/yy")
                    slDate = gObtainStartCal(slDate)
                    lmLYStartDates(ilLoop) = gDateValue(slDate)
                    llDate = lmLYStartDates(ilLoop)
                    llDate = llDate - 1
                 Next ilLoop
            Else        'start with beginning of last year calendar
                'ReDim lmLYStartDates(1 To 13) As Long
                ReDim lmLYStartDates(0 To 13) As Long   'Index zero ignored
                slStart = "1/1/" & str$(igYear - 1)
                gBuildStartDates slStart, 4, 13, lmLYStartDates()
            End If
        End If
        Exit Sub
End Sub
'
'               Search the array of advertisers (tmNewAdvtBus) that aired in this year/last year
'               <input> ilAdfCode = advertiser code
'               Return : -1 if not found, else index to the advertiser aired entry
Private Function gBinarySearchNewBus(ilAdfCode As Integer)
Dim ilMiddle As Integer
Dim ilMin As Integer
Dim ilMax As Integer
    ilMin = LBound(tmNewAdvtBus)
    ilMax = UBound(tmNewAdvtBus) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilAdfCode = tmNewAdvtBus(ilMiddle).iAdfCode Then
            'found the match
            gBinarySearchNewBus = ilMiddle
            Exit Function
        ElseIf ilAdfCode < tmNewAdvtBus(ilMiddle).iAdfCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gBinarySearchNewBus = -1
    
End Function
'
'               Build array  of what advertisers aired/airing in previous/current year
'               mCreateAdvtAiredArrayAll
'               <input> ilThisYear - true for this years processing, else false indicating last year processing
'                       ilPeriods - # periods to process for prev year
Public Sub mCreateAdvtAiredArrayForAll(ilThisYear As Integer, ilPeriods As Integer)
Dim ilFoundVef As Integer
Dim llTotalGross As Long
Dim ilTemp As Integer
Dim ilLoop As Integer
Dim slStr As String
Dim ilAdfInx As Integer
Dim llTempStartDates() As Long

        'ReDim llTempStartDates(1 To ilPeriods + 1)
        ReDim llTempStartDates(0 To ilPeriods + 1)  'Index zero ignored
        ilFoundVef = False
        llTotalGross = 0
        For ilTemp = 1 To ilPeriods
            llTotalGross = llTotalGross + lmProject(ilTemp)
        Next ilTemp
        
        If ilThisYear Then
            For ilLoop = 1 To ilPeriods + 1
                llTempStartDates(ilLoop) = lmStartDates(ilLoop)
            Next ilLoop
        Else
         For ilLoop = 1 To ilPeriods + 1
            llTempStartDates(ilLoop) = lmLYStartDates(ilLoop)
            Next ilLoop
        End If

        'build all advertisers airing into array, even tho $0 business
        For ilTemp = LBound(tmNewAdvtBus) To UBound(tmNewAdvtBus) - 1
            If tmNewAdvtBus(ilTemp).iAdfCode = tmChf.iAdfCode Then
                If llTotalGross <> 0 Then
                    If ilThisYear Then
                        'this years data, determine first month aired
                        For ilLoop = 1 To ilPeriods Step 1
                            If lmProject(ilLoop) <> 0 Then           'keep track of the earliest month advt aired
                                If tmNewAdvtBus(ilTemp).lTYMonthFirstAired > llTempStartDates(ilLoop) Or tmNewAdvtBus(ilTemp).lTYMonthFirstAired = 0 Then
                                    tmNewAdvtBus(ilTemp).lTYMonthFirstAired = llTempStartDates(ilLoop)
                                End If
                                Exit For
                            End If
                        Next ilLoop
                    Else                    'last year, determine last month aired
                    'update last month index of $ aired
                        For ilLoop = ilPeriods To 1 Step -1
                            If lmProject(ilLoop) > 0 Then           'keep track of the latest month advt aired
                                If tmNewAdvtBus(ilTemp).lLYMonthLastAired < llTempStartDates(ilLoop) Or tmNewAdvtBus(ilTemp).lLYMonthLastAired = 0 Then
                                    tmNewAdvtBus(ilTemp).lLYMonthLastAired = llTempStartDates(ilLoop)
                                End If
                                Exit For
                            End If
                        Next ilLoop
                    End If
                End If
                ilFoundVef = True
                tmNewAdvtBus(ilTemp).iUpdateAdvt = False        'flag to indidcate to update the advt with new Advt found status with month/year info
                Exit For
            End If
        Next ilTemp
        If Not ilFoundVef Then
            ilTemp = UBound(tmNewAdvtBus)
            tmNewAdvtBus(ilTemp).iAdfCode = tmChf.iAdfCode
            'find the adv to get the last year/month considered New
            ilAdfInx = gBinarySearchAdf(tmChf.iAdfCode)
            If ilAdfInx <> -1 Then
                tmNewAdvtBus(ilTemp).iLastMonthYearNew(0) = tgCommAdf(ilAdfInx).iLastMonthNew
                tmNewAdvtBus(ilTemp).iLastMonthYearNew(1) = tgCommAdf(ilAdfInx).iLastYearNew
            Else
                'no advt found
                tmNewAdvtBus(ilTemp).iLastMonthYearNew(0) = 0
                tmNewAdvtBus(ilTemp).iLastMonthYearNew(1) = 0
            End If
            slStr = Trim$(str$(tmChf.iAdfCode))
            Do While Len(slStr) < 5
                slStr = "0" & slStr
            Loop
            tmNewAdvtBus(ilTemp).sKey = slStr
            tmNewAdvtBus(ilTemp).lLYMonthLastAired = 0
            tmNewAdvtBus(ilTemp).lTYMonthFirstAired = 0
            tmNewAdvtBus(ilTemp).iUpdateAdvt = False        'flag to indidcate to update the advt with new Advt found status with month/year info
            
            'If llTotalGross <> 0 Then
                If ilThisYear Then
                    'this years data, determine first month aired
                    For ilLoop = 1 To ilPeriods Step 1
                        If lmProject(ilLoop) <> 0 Then           'keep track of the earliest month advt aired
                            If tmNewAdvtBus(ilTemp).lTYMonthFirstAired > llTempStartDates(ilLoop) Or tmNewAdvtBus(ilTemp).lTYMonthFirstAired = 0 Then
                                tmNewAdvtBus(ilTemp).lTYMonthFirstAired = llTempStartDates(ilLoop)
                            End If
                            Exit For
                        End If
                    Next ilLoop
                Else                    'last year, determine last month aired
                'update last month index of $ aired
                    For ilLoop = ilPeriods To 1 Step -1
                        If lmProject(ilLoop) <> 0 Then          'keep track of the latest month advt aired
                            If tmNewAdvtBus(ilTemp).lLYMonthLastAired < llTempStartDates(ilLoop) Or tmNewAdvtBus(ilTemp).lLYMonthLastAired = 0 Then
                                tmNewAdvtBus(ilTemp).lLYMonthLastAired = llTempStartDates(ilLoop)
                            End If
                            Exit For
                        End If
                    Next ilLoop
                End If
            'End If
            ReDim Preserve tmNewAdvtBus(0 To ilTemp + 1) As NEWADVTBUS
        End If      'not ilFoundVef
        'End If          'totalgross <> 0
        Exit Sub
End Sub
'
'            process rep contracts for the current week, or
'            Process rep and/or AirTime for Last year
'
'           <input> llChfCode - single contract # or 0 if all
'                   slEarliestDate - earliest NTR to gather
'                   slLatestDate - latest NTR to gather
'                   ilPeriods - # periods to process, last year may be different than current year
'                   ilDoRep - true to process rep contracts
'                   ilDoAirTime = true to process air time contracts
Public Sub mGatherCurrentYearREP(llChfCode As Long, slEarliestDate As String, slLatestDate As String, llStartDates() As Long, ilPeriods As Integer)
Dim ilLoop As Integer
ReDim ilRepVehicles(0 To 0) As Integer
Dim ilHOState As Integer
Dim ilAdjustDays As Integer
Dim ilValidDays(0 To 6) As Integer
Dim ilLoopOnKey As Integer
Dim ilRet As Integer
Dim llContrCode As Long
Dim ilClf As Integer
Dim ilTemp As Integer
Dim ilIncludeVehicle As Integer
Dim ilByCodeOrNumber As Integer
Dim ilOk As Integer
Dim ilfirstTime As Integer
Dim ilFoundVef As Integer
Dim ilIndex As Integer
Dim ilWeekOrMonth As Integer
Dim llTotalGross As Long
Dim slStr As String
Dim ilIsItPolitical As Integer
'ReDim tmChfAdvtExt(1 To 1) As CHFADVTEXT
ReDim tmChfAdvtExt(0 To 0) As CHFADVTEXT
Dim ilWhichRate As Integer

            'build array of rep and/or airtime vehicles
            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1
                If (tgMVef(ilLoop).sType = "R" And tgMVef(ilLoop).sState = "A") Then
                    ilIncludeVehicle = mFilterVehicle(tgMVef(ilLoop).iCode)
                    If ilIncludeVehicle Then
                        ilRepVehicles(UBound(ilRepVehicles)) = tgMVef(ilLoop).iCode
                        ReDim Preserve ilRepVehicles(0 To UBound(ilRepVehicles) + 1) As Integer
                    End If
                End If
            Next ilLoop
     
            If llChfCode > 0 Then                   'single contract entered?
                'tmChfAdvtExt(1).lCode = tmChf.lCode
                'ReDim Preserve tmChfAdvtExt(1 To 2) As CHFADVTEXT
                tmChfAdvtExt(0).lCode = tmChf.lCode
                ReDim Preserve tmChfAdvtExt(0 To 1) As CHFADVTEXT
            Else
                ilHOState = 2
                ilRet = gObtainCntrForDate(RptSelBO, slEarliestDate, slLatestDate, smCntrStatus, smCntrType, ilHOState, tmChfAdvtExt())
            End If
            
            ilAdjustDays = (llStartDates(ilPeriods + 1) - llStartDates(1)) + 1
            ReDim lmCalSpots(0 To ilAdjustDays) As Long        'init buckets for daily calendar values (spots unused in this report)
            ReDim lmCalAmt(0 To ilAdjustDays) As Long
            ReDim lmCalAcqAmt(0 To ilAdjustDays) As Long            'acq unused in this reprot
            ReDim lmAcquistion(0 To ilAdjustDays) As Long
            For ilLoop = 0 To 6                         'days of the week
                ilValidDays(ilLoop) = True              'force alldays as valid
            Next ilLoop

            ilfirstTime = True
            For ilLoopOnKey = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1
                llContrCode = tmChfAdvtExt(ilLoopOnKey).lCode
                    'obtain contract header and flights
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tmChf, tgClfCT(), tgCffCT())

                ilRet = gIsCntrRep(tmChf.lVefCode, hmVsf, ilRepVehicles())
                
                'look for rep contracts only
                If ilRet Then           'process contr if at least vehicle is rep
                    
'                    llContrCode = tmChfAdvtExt(ilLoopOnKey).lCode
'                    'obtain contract header and flights
'                    ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tmChf, tgClfCT(), tgCffCT())

                    ReDim tmSpotAndRev(0 To 0) As SPOTBBSTATS
                    mCreateSlspSplitTable
                    For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                        tmClf = tgClfCT(ilClf).ClfRec
                        
                        'test required when air time and rep are allowed on the same contract
                        ilRet = gBinarySearchVef(tmClf.iVefCode)
                        If ilRet <> -1 Then
                            ilByCodeOrNumber = 0            'reference by chf code
                            ilOk = mFilterContract(ilByCodeOrNumber, tmChf.lCode, ilfirstTime, ilIsItPolitical)
                            ilfirstTime = False
                            If ilOk Then
                                If Not gFilterLists(tmChf.iAdfCode, imInclAdvtCodes, imUseAdvtCodes()) Then
                                    ilOk = False
                                End If
                                If tgMVef(ilRet).sType <> "R" Then          '12-2-16 drill down to lines and it must be a rep vehicle to include
                                    ilOk = False
                                End If
                            End If
                            
                            If (ilOk = True) Then
                                For ilTemp = 1 To 14 Step 1 'init projection $ each time
                                    lmProject(ilTemp) = 0
                                    lmAcquisition(ilTemp) = 0
                                    lmProjectSpots(ilTemp) = 0
                                Next ilTemp
                                 'init the cal buckets, if used
                                For ilTemp = 0 To UBound(lmCalSpots) - 1
                                    lmCalSpots(ilTemp) = 0        'init buckets for daily calendar values
                                    lmCalAmt(ilTemp) = 0
                                    lmCalAcqAmt(ilTemp) = 0
                                Next ilTemp
                                'use hidden lines
                                If tmClf.sType <> "A" And tmClf.sType <> "O" And tmClf.sType <> "E" Then
                                    If imPerType = 2 Or imPerType = 3 Or imPerType = 1 Then       'std or corporate, or weekly
                                        ilWeekOrMonth = 1
                                        If imPerType = 1 Then           'weekly
                                            ilWeekOrMonth = 2
                                        End If
                                        'gBuildFlights ilClf, llStdStartDates(), 1, imPeriods + 1, lmProject(), ilWeekOrMonth, tgClfCT(), tgCffCT()
                                        ilWhichRate = 0     '0 = use spot rate, 1 = use acq rate, 2 = if acq non-zero, use it.  otherwise if 0, default to line rate
                                        gBuildFlightSpotsAndRevenue ilClf, llStartDates(), 1, ilPeriods + 1, lmProject(), lmProjectSpots(), ilWeekOrMonth, ilWhichRate, tgClfCT(), tgCffCT()

                                    ElseIf imPerType = 4 Then                   'calendar
                                        gCalendarFlights tgClfCT(ilClf), tgCffCT(), llStartDates(1), llStartDates(ilPeriods + 1), ilValidDays(), True, lmCalAmt(), lmCalSpots(), lmCalAcqAmt(), tmPriceTypes
                                        gAccumCalFromDays llStartDates(), lmCalAmt(), lmCalAcqAmt(), False, lmProject(), lmAcquisition(), ilPeriods
                                        gAccumCalSpotsFromDays llStartDates(), lmCalSpots(), lmProjectSpots(), ilPeriods
                                    End If
                                End If
                                
                                ilFoundVef = False
                                'Build $ array new each contract
                                For ilTemp = LBound(tmSpotAndRev) To UBound(tmSpotAndRev) - 1 Step 1
                                    If tmSpotAndRev(ilTemp).iVefCode = tmClf.iVefCode And tmSpotAndRev(ilTemp).lChfCode And tmChf.lCode Then
                                        For ilLoop = 1 To 14
                                            tmSpotAndRev(ilTemp).lRev(ilLoop - 1) = tmSpotAndRev(ilTemp).lRev(ilLoop - 1) + lmProject(ilLoop)
                                            tmSpotAndRev(ilTemp).lSpots(ilLoop - 1) = tmSpotAndRev(ilTemp).lSpots(ilLoop - 1) + lmProjectSpots(ilLoop)
                                        Next ilLoop
                                        ilFoundVef = True
                                        Exit For
                                    End If
                                Next ilTemp
                                If Not (ilFoundVef) Then
                                    ilTemp = UBound(tmSpotAndRev)
                                    tmSpotAndRev(ilTemp).lChfCode = tmChf.lCode
                                    tmSpotAndRev(ilTemp).lCntrNo = tmChf.lCntrNo
                                    tmSpotAndRev(ilTemp).iAdfCode = tmChf.iAdfCode
                                    tmSpotAndRev(ilTemp).iAgfCode = tmChf.iAgfCode
                                    tmSpotAndRev(ilTemp).iVefCode = tmClf.iVefCode
                                    tmSpotAndRev(ilTemp).iPctTrade = tmChf.iPctTrade
                                    tmSpotAndRev(ilTemp).sTradeComm = tmChf.sAgyCTrade       'agy commissionable for trades
                                    tmSpotAndRev(ilTemp).iAgyCommPct = 0      'direct, no comm
                                    If tmChf.iAgfCode > 0 Then
                                        ilIndex = gBinarySearchAgf(tmChf.iAgfCode)
                                        If ilIndex <> -1 Then
                                            'tmSpotAndRev(ilTemp).iAgyCommPct = 1500       tgCommAgf(ilIndex).
                                            tmSpotAndRev(ilTemp).iAgyCommPct = tgCommAgf(ilIndex).iCommPct
                                         End If
                                    End If
                                    
                                    If tmChf.iPctTrade = 0 Then
                                        tmSpotAndRev(ilTemp).sCashTrade = "C"
                                    ElseIf tmChf.iPctTrade = 100 Then
                                        tmSpotAndRev(ilTemp).sCashTrade = "T"
                                    Else
                                        tmSpotAndRev(ilTemp).sCashTrade = "S"       'split cash trade
                                    End If
                                    tmSpotAndRev(ilTemp).iIsItHardCost = False            'flag to indicate if hard cost or not
                                    tmSpotAndRev(ilTemp).sIsNTR = "N"                              'flag to indicate its an NTR item
                                    tmSpotAndRev(ilTemp).iMnfComp = tmChf.iMnfComp(0)      'agy commissionable for trades
                                    tmSpotAndRev(ilTemp).iMnfRevSet = tmChf.iMnfRevSet(0)       'agy commissionable for trades
                                    tmSpotAndRev(ilTemp).iNTRMnfType = 0
                                    For ilLoop = 0 To 9
                                        tmSpotAndRev(ilTemp).lComm(ilLoop) = lmSlspSplitPct(ilLoop)
                                        tmSpotAndRev(ilTemp).iSlfCode(ilLoop) = imSlspSplitCodes(ilLoop)
                                    Next ilLoop
        
                                    For ilLoop = 1 To 14
                                        tmSpotAndRev(ilTemp).lRev(ilLoop - 1) = lmProject(ilLoop)
                                        tmSpotAndRev(ilTemp).lSpots(ilLoop - 1) = tmSpotAndRev(ilTemp).lSpots(ilLoop - 1) + lmProjectSpots(ilLoop)
                                    Next ilLoop
                                    ReDim Preserve tmSpotAndRev(0 To ilTemp + 1) As SPOTBBSTATS
                                End If
  
                            End If              'veftype = "R"
                        End If                  'ilREt >= 0
                    Next ilClf
                    
                    mUpdateGRFForAll
                End If
            Next ilLoopOnKey
            sgCntrForDateStamp = ""         'reset for another run
 
            Erase tmSpotAndRev
            Erase tmChfAdvtExt
            Erase lmCalSpots, lmCalAmt, lmCalAcqAmt, lmAcquisition
        Exit Sub
End Sub
'
'                   mGatherCurrentYearNTR - find all NTR for last year/this year
'                   <input> llChfCode - single contract # or 0 if all
'                       slEarliestDate - earliest NTR to gather
'                       slLatestDate - latest NTR to gather
'                       llStartDates - array of period start dates for as many periods requested
'                       ilPeriods - # periods to process, last year may be different than current year
Private Sub mGatherCurrentYearNTR(llChfCode As Long, slEarliestDate As String, slLatestDate As String, llStartDates() As Long, ilPeriods As Integer)
Dim slDate As String
Dim llDate As Long
Dim ilMonthInx As Integer
Dim ilFoundMonth As Integer
Dim ilFoundVef As Integer
Dim ilTemp As Integer
Dim llSBFLoop As Long                   '12-2-16 chg from int to long
Dim ilIsItHardCost As Integer
Dim ilFoundOption As Integer
Dim ilWhichKey As Integer
Dim llEarliestDate As Long
Dim llLatestDate As Long
Dim ilAgyComm As Integer
Dim ilIndex As Integer
Dim ilRet As Integer
Dim ilByCodeOrNumber As Integer
Dim ilOk As Integer
Dim ilIncludeVehicle As Integer
Dim ilUpper As Integer
Dim ilfirstTime As Integer
Dim ilLoop As Integer
Dim ilIsItPolitical As Integer
ReDim tmSbfList(0 To 0) As SBF


        llEarliestDate = gDateValue(slEarliestDate)
        llEarliestDate = gDateValue(slEarliestDate)
        llLatestDate = gDateValue(slLatestDate)
        If lmSingleCntr > 0 Then
            ilWhichKey = 0
        Else
            ilWhichKey = 2          'trantype, date
        End If

        ilRet = gObtainSBF(RptSelBO, hmSbf, llChfCode, slEarliestDate, slLatestDate, tmNTRTypes, tmSbfList(), ilWhichKey)
        ReDim tmSpotAndRev(0 To 0) As SPOTBBSTATS

        ilfirstTime = True
        For llSBFLoop = LBound(tmSbfList) To UBound(tmSbfList) - 1
            tmSbf = tmSbfList(llSBFLoop)
            gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), slDate
            llDate = gDateValue(slDate)
                
            ilFoundOption = True
            ilIsItHardCost = gIsItHardCost(tmSbf.iMnfItem, tgNTRMnf())
    
            If ilIsItHardCost Then              'hard cost item
                If Not tmCntTypes.iHardCost Then
                    ilFoundOption = False
                End If
            Else                                'normal NTR
                If Not tmCntTypes.iNTR Then
                    ilFoundOption = False
                End If
            End If
    
            If ilFoundOption Then
                ilFoundMonth = False
                For ilMonthInx = 1 To ilPeriods Step 1         'loop thru months to find the match
                    If llDate >= llStartDates(ilMonthInx) And llDate < llStartDates(ilMonthInx + 1) Then
                        ilFoundMonth = True
                        Exit For
                    End If
                Next ilMonthInx
    
                If ilFoundMonth Then
                    'filter out the type of contract
                    ilByCodeOrNumber = 0            'reference by chf code
                    ilOk = mFilterContract(ilByCodeOrNumber, tmSbf.lChfCode, ilfirstTime, ilIsItPolitical)
                    ilfirstTime = False
                    If ilOk Then
                        If Not gFilterLists(tmChf.iAdfCode, imInclAdvtCodes, imUseAdvtCodes()) Then
                            ilOk = False
                        End If
                    End If
                    ilIncludeVehicle = mFilterVehicle(tmSbf.iBillVefCode)
                    If (ilOk) And (ilIncludeVehicle) Then
                        ilFoundVef = False
                        'setup vehicle that spot was moved to
                        'determine agency commission on the individual item
                        ilAgyComm = 0
                        If tmSbf.sAgyComm = "Y" Then
                            'determine the amt of agy commission; can vary per agy
                            If tmChf.iAgfCode > 0 Then
                                ilIndex = gBinarySearchAgf(tmChf.iAgfCode)
                                If ilIndex <> -1 Then
                                    'ilAgyComm = 1500
                                    ilAgyComm = tgCommAgf(ilIndex).iCommPct
                                End If
                            End If
                        End If
                        
                        For ilTemp = LBound(tmSpotAndRev) To UBound(tmSpotAndRev) - 1 Step 1
                            If tmSpotAndRev(ilTemp).iVefCode = tmSbf.iBillVefCode And tmSpotAndRev(ilTemp).iAgyCommPct = ilAgyComm And tmSpotAndRev(ilTemp).lChfCode = tmSbf.lChfCode Then
                                tmSpotAndRev(ilTemp).lRev(ilMonthInx - 1) = tmSpotAndRev(ilTemp).lRev(ilMonthInx - 1) + (tmSbf.lGross * tmSbf.iNoItems)
                                ilFoundVef = True
                                Exit For
                            End If
                        Next ilTemp
                        If Not (ilFoundVef) Then
                            ilTemp = UBound(tmSpotAndRev)
                            tmSpotAndRev(ilTemp).lChfCode = tmSbf.lChfCode
                            tmSpotAndRev(ilTemp).lCntrNo = tmChf.lCntrNo
                            tmSpotAndRev(ilTemp).iAdfCode = tmChf.iAdfCode
                            tmSpotAndRev(ilTemp).iAgfCode = tmChf.iAgfCode
                            tmSpotAndRev(ilTemp).iVefCode = tmSbf.iBillVefCode
                            tmSpotAndRev(ilTemp).iPctTrade = tmChf.iPctTrade
                            tmSpotAndRev(ilTemp).sTradeComm = tmChf.sAgyCTrade       'agy commissionable for trades
                            tmSpotAndRev(ilTemp).iAgyCommPct = ilAgyComm
                            If tmChf.iPctTrade = 0 Then
                                tmSpotAndRev(ilTemp).sCashTrade = "C"
                            ElseIf tmChf.iPctTrade = 100 Then
                                tmSpotAndRev(ilTemp).sCashTrade = "T"
                            Else
                                tmSpotAndRev(ilTemp).sCashTrade = "S"       'split cash trade
                            End If
                            tmSpotAndRev(ilTemp).iIsItHardCost = ilIsItHardCost            'flag to indicate if hard cost or not
                            tmSpotAndRev(ilTemp).sIsNTR = "Y"                              'flag to indicate its an NTR item
                            tmSpotAndRev(ilTemp).iMnfComp = tmChf.iMnfComp(0)      'agy commissionable for trades
                            tmSpotAndRev(ilTemp).iMnfRevSet = tmChf.iMnfRevSet(0)       'agy commissionable for trades
                            tmSpotAndRev(ilTemp).iNTRMnfType = tmSbf.iMnfItem        '8-10-06
                            mCreateSlspSplitTable                   'build only selected slsp into split table
                            For ilLoop = 0 To 9
                                tmSpotAndRev(ilTemp).lComm(ilLoop) = lmSlspSplitPct(ilLoop)
                                tmSpotAndRev(ilTemp).iSlfCode(ilLoop) = imSlspSplitCodes(ilLoop)
                            Next ilLoop

                            tmSpotAndRev(ilTemp).lRev(ilMonthInx - 1) = (tmSbf.lGross * tmSbf.iNoItems)
                            ReDim Preserve tmSpotAndRev(0 To ilTemp + 1) As SPOTBBSTATS
                        End If
                        
                    Else
                        ilTemp = ilTemp
                    End If              'ilok and ilincludevehicle
                End If
            End If
        Next llSBFLoop
        
        mUpdateGRFForAll
        Erase tmSbfList, tmSpotAndRev
        Exit Sub
End Sub
'
'           mGetPerTypeStartDatesForNew - obtain the start date of a std, cal or corp month
'           any date
'           <input> slDate - date to obtain the start date of month
'           Return - Start date of a given month
Private Function mGetPerTypeStartDatesForNew(slDate As String) As Long
Dim llDate As Long

        mGetPerTypeStartDatesForNew = 0
        If imPerType = 2 Then           'std
            slDate = gObtainEndStd(slDate)
            llDate = gDateValue(slDate) + 1
            'increment for next month
            slDate = Format$(llDate, "m/d/yy")
            llDate = gDateValue(slDate)
        ElseIf imPerType = 3 Then           'corp
            slDate = gObtainEndCorp(slDate, False)
            llDate = gDateValue(slDate) + 1
            'increment for next month
            slDate = Format$(llDate, "m/d/yy")
            llDate = gDateValue(slDate)
        ElseIf imPerType = 4 Then           'cal
            slDate = gObtainEndCal(slDate)
            llDate = gDateValue(slDate) + 1
            'increment for next month
            slDate = Format$(llDate, "m/d/yy")
            llDate = gDateValue(slDate)
        End If
        mGetPerTypeStartDatesForNew = llDate            'return the new start date of month
End Function
