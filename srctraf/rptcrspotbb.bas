Attribute VB_Name = "RptCrSpotBB"
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

Dim hmAgf As Integer            'Agency  file handle
Dim tmAgf As ADF                'AGF record image
Dim imAgfRecLen As Integer      'AGF record length

Dim hmGrf As Integer            'Temp  file handle
Dim tmGrf As GRF                'Temp file record image
Dim imGrfRecLen As Integer      'Temp file record length

Dim hmGhf As Integer            'Temp  file handle
Dim tmGhf As GHF                'Temp file record image
Dim imGhfRecLen As Integer      'Temp file record length
Dim tmGhfSrchKey1 As GHFKEY1

Dim hmGsf As Integer            'Temp  file handle
Dim tmGsf As GSF                'Temp file record image
Dim imGsfRecLen As Integer      'Temp file record length
Dim tmGsfSrchKey1 As GSFKEY1

Dim hmMnf As Integer            'Multi-list file handle
Dim tmMnf As MNF                '
Dim imMnfRecLen As Integer

Dim hmSlf As Integer            'Salesperson  file handle
Dim tmSlf As SLF                'SLF file record image
Dim imSlfRecLen As Integer      'SLF file record length

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
Dim imUseAgfCodes() As Integer        'array of agy codes to include/exclude
Dim imInclAgfCodes As Integer               'flag to incl or exclude agy codes
Dim imUseCatCodes() As Integer        'array of bus category codes to include/exclude
Dim imInclCatCodes As Integer               'flag to incl or exclude bus category codes
Dim imUseProdCodes() As Integer        'array of prod prot codes to include/exclude
Dim imInclProdCodes As Integer               'flag to incl or exclude prod prot codes
Dim imUseSlfCodes() As Integer        'array of slsp codes to include/exclude
Dim imInclSlfCodes As Integer
Dim imUseVGCodes() As Integer           'vehicle group items
Dim imInclVGCodes As Integer

Dim imSort1 As Integer
Dim imSort2 As Integer
Dim imSort3 As Integer
Dim imSort4 As Integer

Dim imSepEventsForVehicle As Boolean        'true to show subtotals by game

Dim imMajorSet As Integer               'vehicle group selected
Dim imVG As Integer                 'vehicle group mnf code for vehicle
Dim imPeriods As Integer            '# periods to generate
Dim imSlspSplit As Integer          'do slsp splits
Dim imGrossNetSpot As Integer       '1 = gross, 2 = net, 3 = spot count
Dim imPerType As Integer            '1=week, 2= std, 3 = corp, 4= cal
'Dim lmStartDates(1 To 15) As Long       'weekly is max 14 weeks (some 14 week qtrs)
Dim lmStartDates(0 To 15) As Long       'weekly is max 14 weeks (some 14 week qtrs). Index zero ignored
Dim lmCalAmt() As Long              'calendar calcs
Dim lmCalAcqAmt() As Long           'calendar calcs
'Dim lmProject(1 To 14) As Long             'calendar calcs
Dim lmProject(0 To 14) As Long             'calendar calcs. Index zero ignored
'Dim lmAcquisition(1 To 14) As Long         'calendar calcs
Dim lmAcquisition(0 To 14) As Long         'calendar calcs. Index zero ignored
'Dim lmProjectSpots(1 To 14) As Long        'calendar spot counts
Dim lmProjectSpots(0 To 14) As Long        'calendar spot counts. Index zero ignored
Dim lmCalSpots() As Long
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
    iGameNo As Integer
    iDate(0 To 1) As Integer
    lghfcode As Long
    sGenDesc As String * 20
End Type

'If adding or changing order of sort/selection list boxes, change these constants and also
'see rptvfyspotbb for any further tests.
Const SORT_ADVT = 1
Const SORT_AGY = 2
Const SORT_BUSCAT = 3
Const SORT_PRODPROT = 4
Const SORT_SLSP = 5
Const SORT_VEHICLE = 6
Const SORT_VG = 7

'
'           Open files required for Spot Business Booked
'           Return - error flag = true for open error
'
Private Function mOpenSpotBBFiles() As Integer
Dim ilRet As Integer
Dim slTable As String * 3
Dim ilError As Integer

    ilError = False
    On Error GoTo mOpenSpotBBFilesErr

    slTable = "Chf"
    hmCHF = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Function
    End If
    imCHFRecLen = Len(tmChf)
    
    slTable = "Clf"
    hmClf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        Exit Function
    End If
    imClfRecLen = Len(tmClf)

    slTable = "Cff"
    hmCff = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        Exit Function
    End If
    imCffRecLen = Len(tmCff)
    
    slTable = "Grf"
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Exit Function
    End If
    imGrfRecLen = Len(tmGrf)

    slTable = "Ghf"
    hmGhf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmGhf)
        btrDestroy hmGhf
        Exit Function
    End If
    imGhfRecLen = Len(tmGhf)


    slTable = "Gsf"
    hmGsf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmGsf)
        btrDestroy hmGsf
        Exit Function
    End If
    imGsfRecLen = Len(tmGsf)
    
    slTable = "Mnf"
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        Exit Function
    End If
    imMnfRecLen = Len(tmMnf)
    
    slTable = "Vef"
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Function
    End If
    imVefRecLen = Len(tmVef)
    
    slTable = "Vsf"
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
        Exit Function
    End If
    imVsfRecLen = Len(tmVsf)
   
    slTable = "Sdf"
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        Exit Function
    End If
    imSdfRecLen = Len(tmSdf)
        
    slTable = "Smf"
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmSmf)
        btrDestroy hmSmf
        Exit Function
    End If
    imSmfRecLen = Len(tmSmf)
    
    slTable = "Adf"
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmAdf)
        btrDestroy hmAdf
        Exit Function
    End If
    imAdfRecLen = Len(tmAdf)
    
    slTable = "Agf"
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmAgf)
        btrDestroy hmAgf
        Exit Function
    End If
    imAgfRecLen = Len(tmAgf)
        
    slTable = "Slf"
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmSlf)
        btrDestroy hmSlf
        Exit Function
    End If
    imSlfRecLen = Len(tmSlf)
    
    slTable = "Rvf"
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        Exit Function
    End If
    imRvfRecLen = Len(tmRvf)

    slTable = "Phf"
    hmPhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmPhf)
        btrDestroy hmPhf
        Exit Function
    End If
    imRvfRecLen = Len(tmRvf)
    
    slTable = "Sbf"
    hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenSpotBBFiles = True
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf
        Exit Function
    End If
    imSbfRecLen = Len(tmSbf)


    Exit Function
    
mOpenSpotBBFilesErr:
    ilError = err.Number
    gBtrvErrorMsg ilRet, "mOpenSpotBBFiles (OpenError) #" & str(ilError) & ": " & slTable, RptSelSN

    Resume Next
End Function
'
'       Generate prepass for Revenue on the Books (was Spot Business Booked)
'       for spots within a span of dates for selective advertiseres,
'       contracts, and vehicles, salespeople, business categories,
'       product protection and agencies
'
'
Public Sub gGenSpotBB()

Dim ilError As Integer
Dim ilVefCode As Integer
Dim llStart As Long
Dim slEarliestDate As String
Dim slLatestDate As String
Dim llEnd As Long
Dim slPerStartDate As String
Dim slPerEndDate As String
Dim ilWhichKey As Integer
Dim ilRet As Integer
Dim llLoopOnSpots As Long
Dim ilOk As Integer
Dim illoop As Integer
Dim slType As String
Dim slNameCode As String
Dim slCode As String
Dim ilFoundCntr As Integer
Dim llLoopOnKey As Long
Dim slCntrStatus As String
Dim slCntrType As String
Dim ilHOState As Integer
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
Dim ilOKVG As Integer
Dim ilTemp As Integer
Dim ilfirstTime As Integer


        ilError = mOpenSpotBBFiles()
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
                                                              'get all the vehicles to loop thru SDF, ignore rep
        
        llStart = lmStartDates(1)       'earliest date
        slEarliestDate = Format$(llStart, "m/d/yy")
        llEnd = lmStartDates(imPeriods + 1) - 1 'latest date
        slLatestDate = Format$(llEnd, "m/d/yy")
        
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf.lGenTime = lgNowTime
        tmGrf.iGenDate(0) = igNowDate(0)
        tmGrf.iGenDate(1) = igNowDate(1)
        
        If lmSingleCntr > 0 Then
            ilWhichKey = INDEXKEY0      'search vef, cntr
        Else
            ilWhichKey = INDEXKEY1      'serach sdf by vef, date
        End If
        If tmCntTypes.iAirTime Then             'include air time
            'loop on vehicles by month or week
            'gather the spots in memory for the period and create another array that contains all the
            'necessary header and $ information to write to prepass once the vehicle/period has been processed
            For llLoopOnKey = LBound(ilVehiclesToProcess) To UBound(ilVehiclesToProcess) - 1
                ReDim tmSpotAndRev(0 To 0) As SPOTBBSTATS           'init array of contract information
                ilVefCode = ilVehiclesToProcess(llLoopOnKey)
                
                'determine if this vehicle should be processed based on vehicle group selectivity
                ilOKVG = True
                gGetVehGrpSets ilVefCode, 0, imMajorSet, ilTemp, imVG   'ilTemp = minor sort code(unused), ilMajorVehGrp = major sort code
                If imVG > 0 Then
                    If Not gFilterLists(imVG, imInclVGCodes, imUseVGCodes()) Then
                        ilOKVG = False
                    End If
                End If
                
                If ilOKVG Then          'vehicle group ok, or not using them
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
                                ilOk = mFilterContract(ilByCodeOrNumber, tmSdf.lChfCode, ilfirstTime)
                                ilfirstTime = False
                                If ilOk Then
                                    ilOk = mFilterAllLists(tmChf.iAdfCode, tmChf.iAgfCode, tmChf.iMnfBus, tmChf.iMnfComp(0), tmChf.iSlfCode())
                                    If ilOk Then
                                        'contract & header info tested, test vehicle
                                        ilIncludeVehicle = mFilterVehicle(ilVefCode)
                                    End If
                                End If
                                ilIncludeContract = ilOk        'save the results of the filtering so it has to be done only once per contract
                            End If
                            If ilIncludeContract = True And ilIncludeVehicle = True Then
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
                    'all periods gathered for the 1 vehicle; create records to temporary file
                    mUpdateGRFForAll
                End If
            Next llLoopOnKey
            Erase tmSdfInfo
            Erase tmSpotAndRev
        
        End If
        
        'Gather NTR or Hard Cost, dont process if spot counts requested
        If ((tmCntTypes.iNTR) Or (tmCntTypes.iHardCost)) And imGrossNetSpot <> 3 Then
            mProcessNTR llChfCode, slEarliestDate, slLatestDate
        End If
        
        
        'Gather Adjustments, dont process if spot counts requested
        If tmTranTypes.iAdj And imGrossNetSpot <> 3 Then
            mProcessAdj slEarliestDate, slLatestDate
        End If
        
        'Gather REP
        If tmCntTypes.iRep Then
            mProcessREPS llChfCode, slEarliestDate, slLatestDate
        End If
       
        
        'close all files
        mCloseSpotBBFiles
        Erase ilVehiclesToProcess
        Exit Sub
End Sub
'           mObtainSelectivity - gather all selectivity entered and place
'           in common variables
'
Private Sub mObtainSelectivity()
Dim slStart As String
Dim llStart As Long
Dim ilDay As Integer
Dim slStamp As String
Dim ilRet As Integer

        imSort1 = (RptSelSpotBB!cbcSort1.ListIndex) + 1     '0 will indicate no sort for other levels of sort
        If imSort1 = SORT_VEHICLE And RptSelSpotBB!rbcGameSubTotal(0).Value Then        'combine events for game (1 total by vehicle)
            imSepEventsForVehicle = False
        Else
            imSepEventsForVehicle = True                    'show subtotals by each event within the vehicle.  this option applies to only first sort by vehicle, no other subsorts
        End If
        imSort2 = RptSelSpotBB!cbcSort2.ListIndex
        imSort3 = RptSelSpotBB!cbcSort3.ListIndex
        imSort4 = RptSelSpotBB!cbcSortVG.ListIndex
        
        'what type of periods:  week, std, corporate, calendar months
        If RptSelSpotBB!rbcPerType(0).Value = True Then     'week
            imPerType = 1
        ElseIf RptSelSpotBB!rbcPerType(1).Value = True Then 'std
            imPerType = 2
        ElseIf RptSelSpotBB!rbcPerType(2).Value = True Then 'corp
            imPerType = 3
        Else
            imPerType = 4                                   'cal
        End If
        
        imPeriods = Val(RptSelSpotBB!edcPeriods.Text)
        
        If imPerType = 1 Then  'set start dates of Weekly periods
            slStart = RptSelSpotBB!edcStart.Text        'date entered
            llStart = gDateValue(slStart)
            'backup to Monday
            ilDay = gWeekDayLong(llStart)
            Do While ilDay <> 0
                llStart = llStart - 1
                ilDay = gWeekDayLong(llStart)
            Loop
            slStart = Format$(llStart, "m/d/yy")
            gBuildStartDates slStart, 3, imPeriods + 1, lmStartDates()
        ElseIf imPerType = 2 Then   'set start dates of 12 standard periods
            slStart = str$(igMonthOrQtr) & "/15/" & str$(igYear)
            gBuildStartDates slStart, 1, imPeriods + 1, lmStartDates()

        ElseIf imPerType = 3 Then   'set start dates of 12 corporate periods
            slStart = str$(igMonthOrQtr) & "/15/" & str$(igYear)
            gBuildStartDates slStart, 2, imPeriods + 1, lmStartDates()

        ElseIf imPerType = 4 Then  'set start dates of 12 calendar periods
            slStart = str$(igMonthOrQtr) & "/1/" & str$(igYear)
            gBuildStartDates slStart, 4, imPeriods + 1, lmStartDates()
        End If
            
        'gross, net or spot counts
        If RptSelSpotBB!rbcGrossNet(0).Value = True Then        'gross
            imGrossNetSpot = 1
        ElseIf RptSelSpotBB!rbcGrossNet(1).Value = True Then        'net
            imGrossNetSpot = 2
        Else
            imGrossNetSpot = 3
        End If

        'Selective contract #
        lmSingleCntr = Val(RptSelSpotBB!edcContract.Text)
        
        tmCntTypes.iHold = gSetCheck(RptSelSpotBB!ckcAllTypes(0).Value)
        tmCntTypes.iOrder = gSetCheck(RptSelSpotBB!ckcAllTypes(1).Value)
        tmCntTypes.iStandard = gSetCheck(RptSelSpotBB!ckcAllTypes(3).Value)
        tmCntTypes.iReserv = gSetCheck(RptSelSpotBB!ckcAllTypes(4).Value)
        tmCntTypes.iRemnant = gSetCheck(RptSelSpotBB!ckcAllTypes(5).Value)
        tmCntTypes.iDR = gSetCheck(RptSelSpotBB!ckcAllTypes(6).Value)
        tmCntTypes.iPI = gSetCheck(RptSelSpotBB!ckcAllTypes(7).Value)
        tmCntTypes.iPSA = gSetCheck(RptSelSpotBB!ckcAllTypes(8).Value)
        tmCntTypes.iPromo = gSetCheck(RptSelSpotBB!ckcAllTypes(9).Value)
        tmCntTypes.iTrade = gSetCheck(RptSelSpotBB!ckcAllTypes(10).Value)
        tmCntTypes.iAirTime = gSetCheck(RptSelSpotBB!ckcAllTypes(11).Value)
        tmCntTypes.iRep = gSetCheck(RptSelSpotBB!ckcAllTypes(12).Value)
        tmCntTypes.iNTR = gSetCheck(RptSelSpotBB!ckcAllTypes(13).Value)
        tmCntTypes.iHardCost = gSetCheck(RptSelSpotBB!ckcAllTypes(14).Value)
        tmCntTypes.iPolit = gSetCheck(RptSelSpotBB!ckcAllTypes(15).Value)
        tmCntTypes.iNonPolit = gSetCheck(RptSelSpotBB!ckcAllTypes(16).Value)
        tmCntTypes.iMissed = gSetCheck(RptSelSpotBB!ckcAllTypes(17).Value)      'spot type inclusion/exclusion uses tmSpotTypes structure
        tmCntTypes.iCancelled = gSetCheck(RptSelSpotBB!ckcAllTypes(18).Value)   'spot type inclusion/exclusion uses tmSpotTypes structure
        tmCntTypes.iXtra = gSetCheck(RptSelSpotBB!ckcAllTypes(19).Value)
'        tmCntTypes.iCash = True                'always include cash
        tmCntTypes.iCash = gSetCheck(RptSelSpotBB!ckcAllTypes(21).Value)            '8-7-19 option to include cash or trade separately
        
        'spot types for inclusion/exclusion
        'tmSpotTypes.iMissed = gSetCheck(RptSelSpotBB!ckcAllTypes(16).Value)
        'tmSpotTypes.iCancel = gSetCheck(RptSelSpotBB!ckcAllTypes(17).Value)
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
            If RptSelSpotBB!ckcAllTypes(20).Value = vbChecked Then  'include billboards
                tmSpotTypes.iClose = True
                tmSpotTypes.iOpen = True
            End If
        End If
    


        'If including adjustments from receivables, look for AN transaction types for NTR, Air Time or Hard Cost (also by option)
        'tmTranTypes.iAdj = gSetCheck(RptSelSpotBB!ckcAdj(0).Value)     'incl rep adjustments
        If RptSelSpotBB!ckcAdj(0).Value = vbChecked Or RptSelSpotBB!ckcAdj(1).Value = vbChecked Then
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
        'tmTranTypes.iCash = True
        tmTranTypes.iCash = gSetCheck(RptSelSpotBB!ckcAllTypes(21).Value)    '8-7-19 option to include cash or trade separately
        tmTranTypes.iMerch = False
        tmTranTypes.iPromo = False
        tmTranTypes.iTrade = gSetCheck(RptSelSpotBB!ckcAllTypes(10).Value)
        
        'NTR types; SBF file has installment records and import records.  Only NTR (Hardcost) of interest
        tmNTRTypes.iImport = False
        tmNTRTypes.iInstallment = False
        If tmCntTypes.iNTR = True Or tmCntTypes.iHardCost = True Then
            tmNTRTypes.iNTR = True
            ilRet = gObtainMnfForType("I", slStamp, tgNTRMnf())        'NTR Item types to check for hard cost
        Else
            tmNTRTypes.iNTR = False
        End If
        
        'Show Slsp Splits and test to split slsp revenue if slsp sort selected
        If imSort1 = SORT_SLSP Or imSort2 = SORT_SLSP Or imSort3 = SORT_SLSP Then
            imSlspSplit = gSetCheck(RptSelSpotBB!ckcShowSlspSplit.Value)
        Else
            imSlspSplit = False         'no slsp sort selected, no revenue splits
        End If
End Sub
Private Sub mCloseSpotBBFiles()
Dim ilRet As Integer

        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        ilRet = btrClose(hmGhf)
        btrDestroy hmGhf
        ilRet = btrClose(hmGsf)
        btrDestroy hmGsf
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
        ilRet = btrClose(hmSlf)
        btrDestroy hmSlf
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        ilRet = btrClose(hmPhf)
        btrDestroy hmPhf
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf
        
        Erase imUsevefcodes
        Erase imUseAdvtCodes
        Erase imUseAgfCodes
        Erase imUseCatCodes
        Erase imUseProdCodes
        Erase imUseSlfCodes
        Erase lmStartDates
        Erase tmSpotAndRev
        Erase tmSdfInfo
        Erase tmSbfList
        Erase lmCalSpots, lmCalAmt, lmCalAcqAmt
   
    Exit Sub
    
End Sub
'                       mFilterAllLists - filter user selections (list boxes) from header
'                       <input> ilAdfCode - advertiser code
'                               ilAgfCode - agency code
'                               ilCatCode-  Bus category code
'                               ilProdCode - primary product protection code
'                               ilSlfCode() - array of 10 slsp codes
Private Function mFilterAllLists(ilAdfCode As Integer, ilAgfCode As Integer, ilCatCode As Integer, ilProdCode As Integer, ilSlfCode() As Integer) As Integer
Dim ilOk As Integer
Dim illoop As Integer
Dim ilSlspOK As Integer

        ilOk = True
        If imSort1 = SORT_ADVT Or imSort2 = SORT_ADVT Or imSort3 = SORT_ADVT Then
            If Not gFilterLists(ilAdfCode, imInclAdvtCodes, imUseAdvtCodes()) Then
                ilOk = False
            End If
        End If
        If (imSort1 = SORT_AGY Or imSort2 = SORT_AGY Or imSort3 = SORT_AGY) Then    '12-15-11 And ilAgfCode > 0 Then
            If Not gFilterLists(ilAgfCode, imInclAgfCodes, imUseAgfCodes()) Then
                ilOk = False
            End If
        End If
        If imSort1 = SORT_BUSCAT Or imSort2 = SORT_BUSCAT Or imSort3 = SORT_BUSCAT Then
            If Not gFilterLists(ilCatCode, imInclCatCodes, imUseCatCodes()) Then
                ilOk = False
            End If
        End If
        If imSort1 = SORT_PRODPROT Or imSort2 = SORT_PRODPROT Or imSort3 = SORT_PRODPROT Then
            If Not gFilterLists(ilProdCode, imInclProdCodes, imUseProdCodes()) Then
                ilOk = False
            End If
        End If
        If imSort1 = SORT_SLSP Or imSort2 = SORT_SLSP Or imSort3 = SORT_SLSP Then
            ilSlspOK = False
            For illoop = LBound(ilSlfCode) To UBound(ilSlfCode)
                If ilSlfCode(illoop) > 0 Then
                    If gFilterLists(ilSlfCode(illoop), imInclSlfCodes, imUseSlfCodes()) Then
                        'found valid one
                        ilSlspOK = True
                        Exit For
                    End If
                End If
            Next illoop
            If Not ilSlspOK Then
                ilOk = False
            End If
        End If
        
        mFilterAllLists = ilOk
        Exit Function
End Function
'
'               Obtain contract and filter selectivity
'               <input> ilByCodeOrNUmber: 0 = use code to retrieve contract
'                                         1 = use contract # (receivables dont have chfcodes)
'                       llContractKey: contract code or Number
Private Function mFilterContract(ilByCodeOrNumber As Integer, llContractKey As Long, ilfirstTime As Integer) As Integer
Dim ilOk As Integer
Dim ilFoundCntr As Integer
Dim ilIsItPolitical As Integer
Dim ilRet As Integer
Dim ilOKForUser As Integer

        ilOk = True
        If ilByCodeOrNumber = 0 Then            'get the contract by code
            If llContractKey <> tmChf.lCode Then
                tmChfSrchKey.lCode = llContractKey
                ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            End If
            '4-8-15 ignore contracts that are not orders and not scheduled
            If (tmChf.sDelete = "Y") Or (tmChf.sStatus <> "H" And tmChf.sStatus <> "O") And (tmChf.sSchStatus <> "F") Then       'deleted header, not an order and not scheduled, contract shouldnt be used
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
                    '12-1-16 wrong field tested for hold or order
                    'If ((tmChf.sSchStatus = "F") Or (tmChf.sSchStatus = "M")) And (tmChf.sDelete <> "Y") And (tmChf.sType = "H" Or tmChf.sType = "O") Then
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
                '7-25-12 Status was tested instead of type fields (they were incorrectly interchanged for testing)
            'scheduled/unsch holds
            If (tmChf.sStatus = "H" Or tmChf.sStatus = "G") And Not (tmCntTypes.iHold) Then
                ilOk = False
            End If
            'scheduled/unsch orders
            If (tmChf.sStatus = "O" Or tmChf.sStatus = "N") And Not (tmCntTypes.iOrder) Then
                ilOk = False
            End If
            'standard orders
            If tmChf.sType = "C" And Not (tmCntTypes.iStandard) Then
                ilOk = False
            End If
            'reserved
            If tmChf.sType = "V" And Not (tmCntTypes.iReserv) Then
                ilOk = False
            End If
            'Remnant
            If tmChf.sType = "T" And Not (tmCntTypes.iRemnant) Then
                ilOk = False
            End If
            'Direct Response
            If tmChf.sType = "R" And Not (tmCntTypes.iDR) Then
                ilOk = False
            End If
            'Per Inquiry
            If tmChf.sType = "Q" And Not (tmCntTypes.iPI) Then
                ilOk = False
            End If
            'PSAs
            If tmChf.sType = "S" And Not (tmCntTypes.iPSA) Then
                ilOk = False
            End If
            'Promo
            If tmChf.sType = "M" And Not (tmCntTypes.iPromo) Then
                ilOk = False
            End If
            
            If tmChf.sStatus = "C" Or tmChf.sStatus = "W" Or tmChf.sStatus = "I" Then           '7-3-13 gnore the proposals
                ilOk = False
            End If
                
            
            'include partial trades
            If tmChf.iPctTrade = 100 And Not (tmCntTypes.iTrade) Then
                ilOk = False
            End If
            If tmChf.iPctTrade = 0 And Not (tmCntTypes.iCash) Then      '8-7-19 option to include cash or trade separately
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

            '7-9-13 allow user to see only whats valid for him
            ilOKForUser = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode())
            If Not ilOKForUser Then
                ilOk = False
            End If
            
        End If

        mFilterContract = ilOk
        Exit Function
End Function
Private Sub mForceIncludeAll(ilIndex As Integer, lbcListBox() As SORTCODE, ilIncludeCodes As Integer, ilUseCodes() As Integer, Form As Form)
Dim ilHowManyDefined As Integer
Dim ilHowMany As Integer
Dim slNameCode As String
Dim illoop As Integer
Dim slCode As String
Dim ilRet As Integer
'ReDim ilUseCodes(1 To 1) As Integer
ReDim ilUseCodes(0 To 0) As Integer

    ilIncludeCodes = False

    For illoop = 0 To Form!lbcSelection(ilIndex).ListCount - 1 Step 1
        slNameCode = lbcListBox(illoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If Form!lbcSelection(ilIndex).Selected(illoop) And ilIncludeCodes Then               'selected ?
            ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
            ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
        Else        'exclude these
            If (Not Form!lbcSelection(ilIndex).Selected(illoop)) And (Not ilIncludeCodes) Then
                ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
                ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
            End If
        End If
    Next illoop
    Exit Sub
End Sub
'
'               mFilterVehicle - filter vehicle and vehicle group for selection
'               <input> ilVefCode : vehicle code
'               <return> true if valid, else false to ignore
Private Function mFilterVehicle(ilVefCode As Integer) As Integer
Dim ilOk As Integer
Dim ilTemp As Integer

        ilOk = True
        If imSort1 = SORT_VEHICLE Or imSort2 = SORT_VEHICLE Or imSort3 = SORT_VEHICLE Then
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
                mCloseSpotBBFiles
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
Dim illoop As Integer
Dim slNameCode As String
Dim slCode As String
Dim ilRet As Integer
Dim ilIndex As Integer

        'setup array of codes to include or exclude, which is less for speed
        If imSort1 = SORT_ADVT Or imSort2 = SORT_ADVT Or imSort3 = SORT_ADVT Then
            gObtainCodesForMultipleLists 0, tgAdvertiser(), imInclAdvtCodes, imUseAdvtCodes(), RptSelSpotBB
        Else
            'force to exclude NONE (include All)
            imInclAdvtCodes = False
            'ReDim imUseAdvtCodes(1 To 1) As Integer
            ReDim imUseAdvtCodes(0 To 0) As Integer
        End If
        
        'ReDim ilVehiclesToProcess(1 To 1) As Integer
        ReDim ilVehiclesToProcess(0 To 0) As Integer
        If imSort1 = SORT_VEHICLE Or imSort2 = SORT_VEHICLE Or imSort3 = SORT_VEHICLE Then
            'build array of vehicles to include or exclude
            gObtainCodesForMultipleLists 5, tgVehicle(), imInclVefCodes, imUsevefcodes(), RptSelSpotBB
            'this array is for air time only, ignore rep vehicles
            For illoop = 0 To RptSelSpotBB!lbcSelection(5).ListCount - 1 Step 1
                slNameCode = tgVehicle(illoop).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If RptSelSpotBB!lbcSelection(5).Selected(illoop) Then               'selected ?
                    ilIndex = gBinarySearchVef(Val(slCode))
                    If ilIndex >= 0 Then
                        If tgMVef(ilIndex).sType = "C" Or tgMVef(ilIndex).sType = "G" Or tgMVef(ilIndex).sType = "S" Then
                            ilVehiclesToProcess(UBound(ilVehiclesToProcess)) = Val(slCode)
                            'ReDim Preserve ilVehiclesToProcess(1 To UBound(ilVehiclesToProcess) + 1)
                            ReDim Preserve ilVehiclesToProcess(0 To UBound(ilVehiclesToProcess) + 1)
                        End If
                    End If
                End If
            Next illoop
            If RptSelSpotBB!lbcSelection(5).ListCount = RptSelSpotBB!lbcSelection(5).SelCount Then       'all selected
                gBuildDormantVehicles ilVehiclesToProcess()            'gather all the selling, conventional and game dormant vehicles
            End If
            'vehicle group items in this list box
            'determine if any group selected
            If imSort4 > 0 Then
                gObtainCodesForMultipleLists 6, tgSOCode(), imInclVGCodes, imUseVGCodes(), RptSelSpotBB
            End If
        Else
            'force to exclude NONE (include All)
            imInclVefCodes = False
            'ReDim imUsevefcodes(1 To 1) As Integer
            ReDim imUsevefcodes(0 To 0) As Integer
            'get all the vehicles
            For illoop = 0 To RptSelSpotBB!lbcSelection(5).ListCount - 1 Step 1
                slNameCode = tgVehicle(illoop).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilIndex = gBinarySearchVef(Val(slCode))
                If ilIndex >= 0 Then
                    If tgMVef(ilIndex).sType = "C" Or tgMVef(ilIndex).sType = "G" Or tgMVef(ilIndex).sType = "S" Then
                        ilVehiclesToProcess(UBound(ilVehiclesToProcess)) = Val(slCode)
                        'ReDim Preserve ilVehiclesToProcess(1 To UBound(ilVehiclesToProcess) + 1)
                        ReDim Preserve ilVehiclesToProcess(0 To UBound(ilVehiclesToProcess) + 1)
                    End If
                End If
            Next illoop
                gBuildDormantVehicles ilVehiclesToProcess()            'gather all the selling, conventional and game dormant vehicles
        End If
        
        If imSort1 = SORT_AGY Or imSort2 = SORT_AGY Or imSort3 = SORT_AGY Then
            'get the agy codes selected
            gObtainCodesForMultipleLists 1, tgAgency(), imInclAgfCodes, imUseAgfCodes(), RptSelSpotBB
        Else
            'force to exclude NONE (include All)
            imInclAgfCodes = False
            'ReDim imUseAgfCodes(1 To 1) As Integer
            ReDim imUseAgfCodes(0 To 0) As Integer
        End If
        
        If imSort1 = SORT_BUSCAT Or imSort2 = SORT_BUSCAT Or imSort3 = SORT_BUSCAT Then
            'get the bus cat codes selected
            gObtainCodesForMultipleLists 2, tgMnfCodeCT(), imInclCatCodes, imUseCatCodes(), RptSelSpotBB
        Else
            'force to exclude NONE (include All)
            imInclCatCodes = False
            'ReDim imUseCatCodes(1 To 1) As Integer
            ReDim imUseCatCodes(0 To 0) As Integer
        End If
        
        If imSort1 = SORT_PRODPROT Or imSort2 = SORT_PRODPROT Or imSort3 = SORT_PRODPROT Then
            'get the prod prot codes selected
            gObtainCodesForMultipleLists 3, tgMNFCodeRpt(), imInclProdCodes, imUseProdCodes(), RptSelSpotBB
        Else
            'force to exclude NONE (include All)
            imInclProdCodes = False
            'ReDim imUseProdCodes(1 To 1) As Integer
            ReDim imUseProdCodes(0 To 0) As Integer
        End If
        
        If imSort1 = SORT_SLSP Or imSort2 = SORT_SLSP Or imSort3 = SORT_SLSP Then
            'get the slsp codes selected
            gObtainCodesForMultipleLists 4, tgSalesperson(), imInclSlfCodes, imUseSlfCodes(), RptSelSpotBB
        Else
            'force to exclude NONE (include All)
            imInclSlfCodes = False
            'ReDim imUseSlfCodes(1 To 1) As Integer
            ReDim imUseSlfCodes(0 To 0) As Integer
        End If
        
        imMajorSet = 0
        If imSort4 > 0 Then
            illoop = RptSelSpotBB!cbcSortVG.ListIndex
            imMajorSet = gFindVehGroupInx(illoop, tgVehicleSets1())
        End If

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
'       grfslfcode - salesperson code
'       grfrdfcode - agency code
'       grfadfcode - advt code
'       grfChfCode - contract # not code
'       grfdatetype - C (cash), T (trade), Z = Hard Cost to keep it separated
'       grfCode2 - vehicle group mnf code
'       grfCode4 - Game Schedule topick up event teams if sort1 is vehicle and separating events for subtotals
'       grfPerGenl(1) - product protection mnf code
'       grfPerGenl(2) - bus category
'       grfPerGenl(3) = game # if primary sort is vehicle, and events within the sport vehicle should be sub-totalled
'       grfDollars(1 - 14) 12 months or 14 weeks (some quarters have 14 weeks)
'       grfDollars(15) - total year or quarter
'       grfGenDesc  - contr product name
Private Sub mUpdateGRFForAll()
'Dim llAmt(1 To 14) As Long
Dim llAmt(0 To 14) As Long  'Index zero ignored
Dim illoop As Integer
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
Dim llLoopOnStats As Long
Dim llSeasonStart As Long
Dim llSeasonEnd As Long
Dim llSchedDate As Long
Dim llInvRevNetAmt As Long
        
        For llLoopOnStats = LBound(tmSpotAndRev) To UBound(tmSpotAndRev) - 1
            tmGrf.iVefCode = tmSpotAndRev(llLoopOnStats).iVefCode
            tmGrf.iAdfCode = tmSpotAndRev(llLoopOnStats).iAdfCode
            tmGrf.iRdfCode = tmSpotAndRev(llLoopOnStats).iAgfCode
            tmGrf.lChfCode = tmSpotAndRev(llLoopOnStats).lCntrNo
            tmGrf.sGenDesc = tmSpotAndRev(llLoopOnStats).sGenDesc           '3-30-18
            tmGrf.iCode2 = tmSpotAndRev(llLoopOnStats).iVG                  'vehicle group
            'tmGrf.iPerGenl(1) = tmSpotAndRev(llLoopOnStats).iMnfComp       'product protection (competitive)
            'tmGrf.iPerGenl(2) = tmSpotAndRev(llLoopOnStats).iMnfRevSet      'busines category
            'tmGrf.iPerGenl(3) = tmSpotAndRev(llLoopOnStats).iGameNo
            tmGrf.iPerGenl(0) = tmSpotAndRev(llLoopOnStats).iMnfComp       'product protection (competitive)
            tmGrf.iPerGenl(1) = tmSpotAndRev(llLoopOnStats).iMnfRevSet      'busines category
            tmGrf.iPerGenl(2) = tmSpotAndRev(llLoopOnStats).iGameNo
            tmGrf.lCode4 = 0

            'If tmGrf.iPerGenl(3) > 0 Then      'got a game #, must be primary sort by vehicle using separate event subtotals
            If tmGrf.iPerGenl(2) > 0 Then      'got a game #, must be primary sort by vehicle using separate event subtotals
'                'find the game header to determine season dates
'                gUnpackDateLong tmSpotAndRev(llLoopOnStats).iDate(0), tmSpotAndRev(llLoopOnStats).iDate(1), llSchedDate
'                tmGhfSrchKey1.iVefCode = tmGrf.iVefCode
'                ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
'
'                Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.iVefCode = tmGrf.iVefCode)
'                    gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llSeasonStart
'                    gUnpackDateLong tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), llSeasonEnd
'                    If llSchedDate >= llSeasonStart And llSchedDate <= llSeasonEnd Then         'found the season
                        'tmGsfSrchKey1.lghfcode = tmGhf.lCode        'header internal code
                        tmGsfSrchKey1.lghfcode = tmSpotAndRev(llLoopOnStats).lghfcode
                        'tmGsfSrchKey1.iGameNo = tmGrf.iPerGenl(3)       'game #
                        tmGsfSrchKey1.iGameNo = tmGrf.iPerGenl(2)       'game #
                        ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet = BTRV_ERR_NONE Then
                            If tmGsf.sGameStatus <> "C" Then            'not cancelled game
                                tmGrf.lCode4 = tmGsf.lCode
                            End If
                        End If
'                    End If
'                    ilRet = btrGetNext(hmGhf, tmGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                Loop
            End If
            
'            ilLo = 1
            If (tmCntTypes.iTrade) And (tmCntTypes.iCash) Then
                ilLo = 1
                ilHi = 2            'include trades
            End If
            
            If (tmCntTypes.iTrade) And Not (tmCntTypes.iCash) Then
                ilLo = 2
                ilHi = 2            'only trade
            End If
            If Not (tmCntTypes.iTrade) And (tmCntTypes.iCash) Then
                ilLo = 1
                ilHi = 1            'only cash
            End If
            
            slCashAgyComm = gIntToStrDec(tmSpotAndRev(llLoopOnStats).iAgyCommPct, 2)
            slPctTrade = gIntToStrDec(tmSpotAndRev(llLoopOnStats).iPctTrade, 0)
            
            For illoop = 0 To 9
                imSlspSplitCodes(illoop) = tmSpotAndRev(llLoopOnStats).iSlfCode(illoop)
                lmSlspSplitPct(illoop) = tmSpotAndRev(llLoopOnStats).lComm(illoop)
            Next illoop
    
            For ilLoopOnSlsp = 0 To 9
                slSharePct = gLongToStrDec(lmSlspSplitPct(ilLoopOnSlsp), 4)       'slsp share
                If imSlspSplitCodes(ilLoopOnSlsp) > 0 Then
                    For ilCorT = ilLo To ilHi                   'loop for cash & trade (if applicable)
                        For ilLoopOnPer = 1 To 14          'init the $ table
                            llAmt(ilLoopOnPer) = 0
                        Next ilLoopOnPer
                        llTotalAll = 0
                        For ilLoopOnPer = 1 To 14
                            If imGrossNetSpot = 3 Then
                                slAmount = gLongToStrDec(tmSpotAndRev(llLoopOnStats).lSpots(ilLoopOnPer - 1), 0) '$ gathered from spots
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
                                    '11-4-14 always save the net value for the Spot Revenue Reigster report which shows gross, net & comm
                                    'this applies to 1 month only processedfor Spot Revenue Reigster report
                                    If ilLoopOnPer = 1 Then
                                        llInvRevNetAmt = Val(slNet)
                                    End If
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
                                        '11-4-14 always save the net value for the Spot Revenue Reigster report which shows gross, net & comm
                                        'this applies to the first and only month processed for Spot Revenue Reigster report
                                        If ilLoopOnPer = 1 Then
                                            llInvRevNetAmt = Val(slNet)
                                        End If
                                    Else
                                        llAmt(ilLoopOnPer) = Val(slNet)
                                        llTotalAll = llTotalAll + Val(slNet)
                                    End If
                                    tmGrf.sDateType = "T"
                                End If
                            End If
                        Next ilLoopOnPer
                        If llTotalAll <> 0 Then         'dont create a prepass record if no value
                            'tmGrf.lDollars(15) = llTotalAll     'net value for Spot Revenue Reigster report
                            tmGrf.lDollars(14) = llTotalAll     'net value for Spot Revenue Reigster report
                            For illoop = 1 To 14
                                tmGrf.lDollars(illoop - 1) = llAmt(illoop)
                            Next illoop
                            'tmGrf.lDollars(16) = llInvRevNetAmt
                            tmGrf.lDollars(15) = llInvRevNetAmt
                            If tmSpotAndRev(llLoopOnStats).sIsNTR = "Y" And tmSpotAndRev(llLoopOnStats).iIsItHardCost = True Then
                                tmGrf.sDateType = "Z"       'sort to end
                            End If
                            tmGrf.iSlfCode = imSlspSplitCodes(ilLoopOnSlsp)        'assume no splits, use 1st slsp
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
                End If
            Next ilLoopOnSlsp
            
        Next llLoopOnStats
     
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
Dim illoop As Integer
Dim ilEvent As Integer
Dim llSchedDate As Long
Dim llSeasonStart As Long
Dim llSeasonEnd As Long
Dim llGhfCode As Long

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
          
        If imSort1 = SORT_VEHICLE And imSepEventsForVehicle And tmSdf.iGameNo > 0 Then            'primary sort is by vehicle, and subsorts required by event
            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSchedDate
            ilEvent = tmSdf.iGameNo                                         'keep games apart for subtotals
            If tmSdf.iVefCode = tmGhf.iVefCode And tmGhf.iVefCode > 0 Then
                'see if correct season for this vehicle
                gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llSeasonStart
                gUnpackDateLong tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), llSeasonEnd
                If llSchedDate >= llSeasonStart And llSchedDate <= llSeasonEnd Then
                    'game header in memory is the same
                    llGhfCode = tmGhf.lCode
                Else
                    'find the game header to determine season dates
                    tmGhfSrchKey1.iVefCode = tmSdf.iVefCode
                    ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                    Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.iVefCode = tmSdf.iVefCode)
                        gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llSeasonStart
                        gUnpackDateLong tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), llSeasonEnd
                        If llSchedDate >= llSeasonStart And llSchedDate <= llSeasonEnd Then         'found the season
                            llGhfCode = tmGhf.lCode
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmGhf, tmGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
            Else
                'find the game header to determine season dates
                tmGhfSrchKey1.iVefCode = tmSdf.iVefCode
                ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.iVefCode = tmSdf.iVefCode)
                    gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llSeasonStart
                    gUnpackDateLong tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), llSeasonEnd
                    If llSchedDate >= llSeasonStart And llSchedDate <= llSeasonEnd Then         'found the season
                        llGhfCode = tmGhf.lCode
                        Exit Do
                    End If
                    ilRet = btrGetNext(hmGhf, tmGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
        Else                                                                'if primary sort by vehicle, assume to combine all events within the vehicle (for sports vehicle)
            ilEvent = 0
            llGhfCode = 0
        End If
        If imGrossNetSpot = 3 Or llSpotRate <> 0 Then       'if doing spot counts (imgrossnetsdpot = 3) or spot rate is non-zero, update the arrays
            'accumulate $ in table with matching contract & line
            For llLoopOnStats = LBound(tmSpotAndRev) To UBound(tmSpotAndRev) - 1
                If tmSdf.lChfCode = tmSpotAndRev(llLoopOnStats).lChfCode And tmSdf.iLineNo = tmSpotAndRev(llLoopOnStats).iLineNo And tmSpotAndRev(llLoopOnStats).iGameNo = ilEvent And tmSpotAndRev(llLoopOnStats).lghfcode = llGhfCode Then
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
                    If ilIndex >= 0 Then
                        'tmSpotAndRev(ilUpper).iAgyCommPct = 1500
                        tmSpotAndRev(ilUpper).iAgyCommPct = tgCommAgf(ilIndex).iCommPct
                     End If
                End If
                tmSpotAndRev(ilUpper).lChfCode = tmSdf.lChfCode
                tmSpotAndRev(ilUpper).lCntrNo = tmChf.lCntrNo
                tmSpotAndRev(ilUpper).sGenDesc = Trim$(tmChf.sProduct)           '3-30-18
                tmSpotAndRev(ilUpper).iVefCode = tmSdf.iVefCode
                tmSpotAndRev(ilUpper).iVG = imVG
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
                'tmSpotAndRev(ilUpper).iMnfRevSet = tmChf.iMnfRevSet(0)       'agy commissionable for trades
                tmSpotAndRev(ilUpper).iMnfRevSet = tmChf.iMnfBus         '2-8-13 should be bus cat, not rev set
                tmSpotAndRev(ilUpper).iLineNo = tmSdf.iLineNo
                tmSpotAndRev(ilUpper).sIsNTR = "N"                      'not NTR
                tmSpotAndRev(ilUpper).lSpots(ilDateInx - 1) = 1
                tmSpotAndRev(ilUpper).lRev(ilDateInx - 1) = llSpotRate
                tmSpotAndRev(ilUpper).iGameNo = ilEvent                 'game # (0 if combining sports events within vehicle, else its the game #)
                If imSort1 = SORT_VEHICLE Then                          'store date table to be able to find the event name later
                    tmSpotAndRev(ilUpper).iDate(0) = tmSdf.iDate(0)
                    tmSpotAndRev(ilUpper).iDate(1) = tmSdf.iDate(1)
                    tmSpotAndRev(ilUpper).lghfcode = llGhfCode
                Else
                    tmSpotAndRev(ilUpper).iDate(0) = 0
                    tmSpotAndRev(ilUpper).iDate(1) = 0
                    tmSpotAndRev(ilUpper).lghfcode = 0                  'season refernce doesn't apply for non-vehicle sort
                End If
                mCreateSlspSplitTable                   'build only selected slsp into split table
                For illoop = 0 To 9
                    tmSpotAndRev(ilUpper).lComm(illoop) = lmSlspSplitPct(illoop)
                    tmSpotAndRev(ilUpper).iSlfCode(illoop) = imSlspSplitCodes(illoop)
                Next illoop
                
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
Dim illoop As Integer
Dim ilIndex As Integer
Dim ilValidCount As Integer

        For illoop = 0 To 9
            imSlspSplitCodes(illoop) = 0
            lmSlspSplitPct(illoop) = 0
        Next illoop
        ilValidCount = 0
        
        If Not imSlspSplit Then
            imSlspSplitCodes(0) = tmChf.iSlfCode(0)
            lmSlspSplitPct(0) = 1000000         '100.0000
        Else
            For illoop = 0 To 9             'loop thru the contract header slsp and determine which ones to include based on selectivity
                If gFilterLists(tmChf.iSlfCode(illoop), imInclSlfCodes, imUseSlfCodes()) Then
                    'found valid one
                    imSlspSplitCodes(ilValidCount) = tmChf.iSlfCode(illoop)
                    lmSlspSplitPct(ilValidCount) = tmChf.lComm(illoop)
                    ilValidCount = ilValidCount + 1
                End If
            Next illoop
        End If
        Exit Sub
End Sub
'
'                       process rep contracts
'
Public Sub mProcessREPS(llChfCode As Long, slEarliestDate As String, slLatestDate As String)
Dim illoop As Integer
'ReDim ilRepVehicles(1 To 1) As Integer
ReDim ilRepVehicles(0 To 0) As Integer
Dim slCntrStatus As String
Dim slCntrType As String
Dim ilHOState As Integer
'ReDim tlChfAdvtExt(1 To 1) As CHFADVTEXT
ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
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
Dim ilOKVG As Integer
Dim ilWeekOrMonth As Integer
Dim ilWhichRate As Integer

            'build array of rep vehicles
            For illoop = LBound(tgMVef) To UBound(tgMVef) - 1
                If tgMVef(illoop).sType = "R" And tgMVef(illoop).sState = "A" Then
                    ilRepVehicles(UBound(ilRepVehicles)) = tgMVef(illoop).iCode
                    'ReDim Preserve ilRepVehicles(1 To UBound(ilRepVehicles) + 1) As Integer
                    ReDim Preserve ilRepVehicles(LBound(ilRepVehicles) To UBound(ilRepVehicles) + 1) As Integer
                End If
            Next illoop
            
            slCntrStatus = ""
            If tmCntTypes.iHold Then
                slCntrStatus = "HG"
            End If
            If tmCntTypes.iOrder Then
                slCntrStatus = slCntrStatus & "ON"
            End If
             
            slCntrType = ""
            If tmCntTypes.iStandard Then
                slCntrType = "C"
            End If
            If tmCntTypes.iReserv Then
                slCntrType = Trim$(slCntrType) & "V"
            End If
            If tmCntTypes.iRemnant Then
                slCntrType = Trim$(slCntrType) & "T"
            End If
            If tmCntTypes.iDR Then
                slCntrType = Trim$(slCntrType) & "R"
            End If
            If tmCntTypes.iPI Then
                slCntrType = Trim$(slCntrType) & "Q"
            End If
            If tmCntTypes.iPSA Then
                slCntrType = Trim$(slCntrType) & "S"
            End If
            If tmCntTypes.iPromo Then
                slCntrType = Trim$(slCntrType) & "M"
            End If
            
            ilHOState = 2                       'get latest orders & revisions  (HOGN plus any revised orders WCI)
     
            If llChfCode > 0 Then                   'single contract entered?
                'tlChfAdvtExt(1).lCode = tmChf.lCode
                'tlChfAdvtExt(1).lVefCode = tmChf.lVefCode
                'ReDim Preserve tlChfAdvtExt(1 To 2) As CHFADVTEXT
                tlChfAdvtExt(0).lCode = tmChf.lCode
                tlChfAdvtExt(0).lVefCode = tmChf.lVefCode
                ReDim Preserve tlChfAdvtExt(0 To 1) As CHFADVTEXT
            Else
                ilRet = gObtainCntrForDate(RptSelSpotBB, slEarliestDate, slLatestDate, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())
            End If
            
            ilAdjustDays = (lmStartDates(imPeriods + 1) - lmStartDates(1)) + 1
            ReDim lmCalSpots(0 To ilAdjustDays) As Long        'init buckets for daily calendar values (spots unused in this report)
            ReDim lmCalAmt(0 To ilAdjustDays) As Long
            ReDim lmCalAcqAmt(0 To ilAdjustDays) As Long            'acq unused in this reprot
            ReDim lmAcquistion(0 To ilAdjustDays) As Long
            For illoop = 0 To 6                         'days of the week
                ilValidDays(illoop) = True              'force alldays as valid
            Next illoop

            ilfirstTime = True
            For ilLoopOnKey = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1
                 ilRet = gIsCntrRep(tlChfAdvtExt(ilLoopOnKey).lVefCode, hmVsf, ilRepVehicles())
                'look for rep contracts only
                If ilRet Then           'process contr if at least vehicle is rep
                    llContrCode = tlChfAdvtExt(ilLoopOnKey).lCode
                    'obtain contract header and flights
                    ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tmChf, tgClfCT(), tgCffCT())
                    ReDim tmSpotAndRev(0 To 0) As SPOTBBSTATS
                    mCreateSlspSplitTable
                    For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                        tmClf = tgClfCT(ilClf).ClfRec
                        
                        'test required when air time and rep are allowed on the same contract
                        ilRet = gBinarySearchVef(tmClf.iVefCode)
                        If ilRet >= 0 Then
                            ilByCodeOrNumber = 0            'reference by chf code
                            ilOk = mFilterContract(ilByCodeOrNumber, tmChf.lCode, ilfirstTime)
                            ilfirstTime = False
                            If ilOk Then
                                ilOk = mFilterAllLists(tmChf.iAdfCode, tmChf.iAgfCode, tmChf.iMnfBus, tmChf.iMnfComp(0), tmChf.iSlfCode())
                                ilIncludeVehicle = mFilterVehicle(tmClf.iVefCode)
                                'determine if this vehicle should be processed based on vehicle group selectivity
                                ilOKVG = True
                                gGetVehGrpSets tmClf.iVefCode, 0, imMajorSet, ilTemp, imVG   'ilTemp = minor sort code(unused), ilMajorVehGrp = major sort code
                                If imVG > 0 Then
                                    If Not gFilterLists(imVG, imInclVGCodes, imUseVGCodes()) Then
                                        ilOKVG = False
                                    End If
                                End If
                                
                            End If
                            'must be rep, contract valid in case single one selected with line vehicle selected
                            If (tgMVef(ilRet).sType = "R") And (ilIncludeVehicle = True) And (ilOk = True) And (ilOKVG = True) Then       'must be a rep line
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
                                        gBuildFlightSpotsAndRevenue ilClf, lmStartDates(), 1, imPeriods + 1, lmProject(), lmProjectSpots(), ilWeekOrMonth, ilWhichRate, tgClfCT(), tgCffCT()

                                    ElseIf imPerType = 4 Then                   'calendar
                                        gCalendarFlights tgClfCT(ilClf), tgCffCT(), lmStartDates(1), lmStartDates(imPeriods + 1), ilValidDays(), True, lmCalAmt(), lmCalSpots(), lmCalAcqAmt(), tmPriceTypes
                                        gAccumCalFromDays lmStartDates(), lmCalAmt(), lmCalAcqAmt(), False, lmProject(), lmAcquisition(), imPeriods
                                        gAccumCalSpotsFromDays lmStartDates(), lmCalSpots(), lmProjectSpots(), imPeriods
                                    End If
                                End If
                                ilFoundVef = False
                                'Build $ array new each contract
                                For ilTemp = LBound(tmSpotAndRev) To UBound(tmSpotAndRev) - 1 Step 1
                                    If tmSpotAndRev(ilTemp).iVefCode = tmClf.iVefCode And tmSpotAndRev(ilTemp).lChfCode And tmChf.lCode Then
                                        For illoop = 1 To 14
                                            tmSpotAndRev(ilTemp).lRev(illoop - 1) = tmSpotAndRev(ilTemp).lRev(illoop - 1) + lmProject(illoop)
                                            tmSpotAndRev(ilTemp).lSpots(illoop - 1) = tmSpotAndRev(ilTemp).lSpots(illoop - 1) + lmProjectSpots(illoop)
                                        Next illoop
                                        ilFoundVef = True
                                        Exit For
                                    End If
                                Next ilTemp
                                If Not (ilFoundVef) Then
                                    ilTemp = UBound(tmSpotAndRev)
                                    tmSpotAndRev(ilTemp).lChfCode = tmChf.lCode
                                    tmSpotAndRev(ilTemp).lCntrNo = tmChf.lCntrNo
                                    tmSpotAndRev(ilTemp).sGenDesc = Trim$(tmChf.sProduct)               '3-30-18
                                    tmSpotAndRev(ilTemp).iAdfCode = tmChf.iAdfCode
                                    tmSpotAndRev(ilTemp).iAgfCode = tmChf.iAgfCode
                                    tmSpotAndRev(ilTemp).iVefCode = tmClf.iVefCode
                                    tmSpotAndRev(ilTemp).iVG = imVG                         'vehicle group if applicable
                                    tmSpotAndRev(ilTemp).iPctTrade = tmChf.iPctTrade
                                    tmSpotAndRev(ilTemp).sTradeComm = tmChf.sAgyCTrade       'agy commissionable for trades
                                    tmSpotAndRev(ilTemp).iAgyCommPct = 0      'direct, no comm
                                    If tmChf.iAgfCode > 0 Then
                                        ilIndex = gBinarySearchAgf(tmChf.iAgfCode)
                                        If ilIndex >= 0 Then
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
                                    'tmSpotAndRev(ilTemp).iMnfRevSet = tmChf.iMnfRevSet(0)       'agy commissionable for trades
                                    tmSpotAndRev(ilTemp).iMnfRevSet = tmChf.iMnfBus       '2-8-13 should be bus cat, not rev sets
                                    tmSpotAndRev(ilTemp).iNTRMnfType = 0
                                    For illoop = 0 To 9
                                        tmSpotAndRev(ilTemp).lComm(illoop) = lmSlspSplitPct(illoop)
                                        tmSpotAndRev(ilTemp).iSlfCode(illoop) = imSlspSplitCodes(illoop)
                                    Next illoop
        
                                    For illoop = 1 To 14
                                        tmSpotAndRev(ilTemp).lRev(illoop - 1) = lmProject(illoop)
                                        tmSpotAndRev(ilTemp).lSpots(illoop - 1) = tmSpotAndRev(ilTemp).lSpots(illoop - 1) + lmProjectSpots(illoop)
                                    Next illoop
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
            Erase tlChfAdvtExt
            Erase lmCalSpots, lmCalAmt, lmCalAcqAmt, lmAcquisition
        Exit Sub
End Sub
Public Sub mProcessNTR(llChfCode As Long, slEarliestDate As String, slLatestDate As String)
    Dim slDate As String
    Dim llDate As Long
    Dim ilMonthInx As Integer
    Dim ilFoundMonth As Integer
    Dim ilFoundVef As Integer
    Dim ilTemp As Integer
    'TTP 10853 - Revenue on the Books: overflow error in mProcessNTR
    'Dim ilSBFLoop As Integer
    Dim llSBFLoop As Long
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
    Dim illoop As Integer
    ReDim tmSbfList(0 To 0) As SBF
    Dim ilOKVG As Integer

    llEarliestDate = gDateValue(slEarliestDate)
    llLatestDate = gDateValue(slLatestDate)
    If lmSingleCntr > 0 Then
        ilWhichKey = 0
    Else
        ilWhichKey = 2          'trantype, date
    End If

    ilRet = gObtainSBF(RptSelSpotBB, hmSbf, llChfCode, slEarliestDate, slLatestDate, tmNTRTypes, tmSbfList(), ilWhichKey)
    ReDim tmSpotAndRev(0 To 0) As SPOTBBSTATS

    ilfirstTime = True
    'TTP 10853 - Revenue on the Books: overflow error in mProcessNTR
    'For ilSBFLoop = LBound(tmSbfList) To UBound(tmSbfList) - 1
    For llSBFLoop = LBound(tmSbfList) To UBound(tmSbfList) - 1
        'tmSbf = tmSbfList(ilSBFLoop)
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
            For ilMonthInx = 1 To imPeriods Step 1         'loop thru months to find the match
                If llDate >= lmStartDates(ilMonthInx) And llDate < lmStartDates(ilMonthInx + 1) Then
                    ilFoundMonth = True
                    Exit For
                End If
            Next ilMonthInx

            If ilFoundMonth Then
                'filter out the type of contract
                ilByCodeOrNumber = 0            'reference by chf code
                ilOk = mFilterContract(ilByCodeOrNumber, tmSbf.lChfCode, ilfirstTime)
                ilfirstTime = False
                If ilOk Then
                    ilOk = mFilterAllLists(tmChf.iAdfCode, tmChf.iAgfCode, tmChf.iMnfBus, tmChf.iMnfComp(0), tmChf.iSlfCode())
                    ilIncludeVehicle = mFilterVehicle(tmSbf.iBillVefCode)
                     'determine if this vehicle should be processed based on vehicle group selectivity
                    ilOKVG = True
                    gGetVehGrpSets tmSbf.iBillVefCode, 0, imMajorSet, ilTemp, imVG   'ilTemp = minor sort code(unused), ilMajorVehGrp = major sort code
                    If imVG > 0 Then
                        If Not gFilterLists(imVG, imInclVGCodes, imUseVGCodes()) Then
                            ilOKVG = False
                        End If
                    End If
            
                End If
                If (ilOk) And (ilIncludeVehicle) And (ilOKVG) Then
                    ilFoundVef = False
                    'setup vehicle that spot was moved to
                    'determine agency commission on the individual item
                    ilAgyComm = 0
                    If tmSbf.sAgyComm = "Y" Then
                        'determine the amt of agy commission; can vary per agy
                        If tmChf.iAgfCode > 0 Then
                            ilIndex = gBinarySearchAgf(tmChf.iAgfCode)
                            If ilIndex >= 0 Then
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
                        tmSpotAndRev(ilTemp).sGenDesc = Trim$(tmChf.sProduct)           '3-30-18
                        tmSpotAndRev(ilTemp).iAdfCode = tmChf.iAdfCode
                        tmSpotAndRev(ilTemp).iAgfCode = tmChf.iAgfCode
                        tmSpotAndRev(ilTemp).iVefCode = tmSbf.iBillVefCode
                        tmSpotAndRev(ilTemp).iVG = imVG                         'vehicle group if applicable
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
                        'tmSpotAndRev(ilTemp).iMnfRevSet = tmChf.iMnfRevSet(0)       'agy commissionable for trades
                        tmSpotAndRev(ilTemp).iMnfRevSet = tmChf.iMnfBus             '2-8-13 should be bus cat, not rev sets
                        tmSpotAndRev(ilTemp).iNTRMnfType = tmSbf.iMnfItem        '8-10-06
                        mCreateSlspSplitTable                   'build only selected slsp into split table
                        For illoop = 0 To 9
                            tmSpotAndRev(ilTemp).lComm(illoop) = lmSlspSplitPct(illoop)
                            tmSpotAndRev(ilTemp).iSlfCode(illoop) = imSlspSplitCodes(illoop)
                        Next illoop

                        tmSpotAndRev(ilTemp).lRev(ilMonthInx - 1) = (tmSbf.lGross * tmSbf.iNoItems)
                        ReDim Preserve tmSpotAndRev(0 To ilTemp + 1) As SPOTBBSTATS
                    End If
                Else
                    ilTemp = ilTemp
                End If              'ilok and ilincludevehicle
            End If
        End If
    'Next ilSBFLoop
    Next llSBFLoop
    mUpdateGRFForAll
    Erase tmSbfList, tmSpotAndRev
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
Dim illoop As Integer
Dim ilIncludeVehicle As Integer
Dim ilTemp As Integer
Dim ilFoundVef As Integer
Dim ilVefIndex As Integer
Dim ilOKVG As Integer

        ReDim tmSpotAndRev(0 To 0) As SPOTBBSTATS
        ilRet = gObtainPhfRvf(RptSelSpotBB, slEarliestDate, slLatestDate, tmTranTypes, tlRvf(), 0)
        ilfirstTime = True
        For llRvf = LBound(tlRvf) To UBound(tlRvf) - 1
            tmRvf = tlRvf(llRvf)
            'filter out the type of contract
            ilByCodeOrNumber = 1            'reference by Contract #
            ilOk = mFilterContract(ilByCodeOrNumber, tmRvf.lCntrNo, ilfirstTime)
            ilfirstTime = False
            If lmSingleCntr > 0 And lmSingleCntr <> tmRvf.lCntrNo Then
                ilOk = False
            End If
            If ilOk Then
                ilOk = mFilterAllLists(tmChf.iAdfCode, tmChf.iAgfCode, tmChf.iMnfBus, tmChf.iMnfComp(0), tmChf.iSlfCode())
                If Trim$(tmRvf.sType) <> "" And tmRvf.sType <> "A" Then     'ANs can only be revenue records, no installment types
                    ilOk = False
                End If
                If tmRvf.iMnfItem > 0 And (Not tmCntTypes.iNTR) Then        'filter out NTR adjustments by user input
                    ilOk = False
                End If
                'REP adjustments are only considered when including adjustments for Air Time types
                'Scheduled spots should not make adjustments
'                If tmRvf.iMnfItem = 0 And (tmCntTypes.iAirTime) Then      'filter out Air Time adjustments by user input, may need to include REP
'                    'determine the type of vehicle this is:  ignore scheduled spots vehicles (not REP)
'                    ilVefIndex = gBinarySearchVef(tmRvf.iAirVefCode)
'                    If ilVefIndex < 0 Then
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
                    If ilVefIndex < 0 Then
                        ilOk = False
                    Else
                        If tgMVef(ilVefIndex).sType = "R" Then
                            If (tmCntTypes.iRep = False Or RptSelSpotBB!ckcAdj(0).Value = vbUnchecked) Then        'rep vehicle and rep adj excluded
                                ilOk = False
                            End If
                        Else
                            If (tmCntTypes.iAirTime = False Or RptSelSpotBB!ckcAdj(1).Value = vbUnchecked) Then     'air time vehicle and adj excluded
                                ilOk = False
                            End If
                        End If
                    End If
                    
                    ilIncludeVehicle = mFilterVehicle(tmRvf.iAirVefCode)
                    'determine if this vehicle should be processed based on vehicle group selectivity
                    ilOKVG = True
                    gGetVehGrpSets tmRvf.iAirVefCode, 0, imMajorSet, ilTemp, imVG   'ilTemp = minor sort code(unused), ilMajorVehGrp = major sort code
                    If imVG > 0 Then
                        If Not gFilterLists(imVG, imInclVGCodes, imUseVGCodes()) Then
                            ilOKVG = False
                        End If
                    End If
                End If

            End If
            If (ilOk) And (ilIncludeVehicle) And (ilOKVG) Then

                If ilOk Then
                    gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate
                    'Determine bucket
                    For ilDateInx = 1 To imPeriods Step 1
                        If (llDate >= lmStartDates(ilDateInx)) And (llDate <= lmStartDates(ilDateInx + 1)) Then
                            'determine if agy commissionable
                            If tmChf.iAgfCode > 0 Then
                                ilIndex = gBinarySearchAgf(tmChf.iAgfCode)
                                If ilIndex >= 0 Then
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
                                tmSpotAndRev(ilTemp).sGenDesc = Trim$(tmChf.sProduct)
                                tmSpotAndRev(ilTemp).iAdfCode = tmChf.iAdfCode
                                tmSpotAndRev(ilTemp).iAgfCode = tmChf.iAgfCode
                                tmSpotAndRev(ilTemp).iVefCode = tmRvf.iAirVefCode
                                tmSpotAndRev(ilTemp).iVG = imVG
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
                                'tmSpotAndRev(ilTemp).iMnfRevSet = tmChf.iMnfRevSet(0)       'agy commissionable for trades
                                tmSpotAndRev(ilTemp).iMnfRevSet = tmChf.iMnfBus       '2-8-13 should be bus cat, not rev sets
                                tmSpotAndRev(ilTemp).iNTRMnfType = tmRvf.iMnfItem
                                mCreateSlspSplitTable                   'build only selected slsp into split table
                                For illoop = 0 To 9
                                    tmSpotAndRev(ilTemp).lComm(illoop) = lmSlspSplitPct(illoop)
                                    tmSpotAndRev(ilTemp).iSlfCode(illoop) = imSlspSplitCodes(illoop)
                                Next illoop

                
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
