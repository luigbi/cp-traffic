Attribute VB_Name = "Rptcred"
Option Explicit
Option Compare Text

Dim hmAgf As Integer            'Agency file handle
Dim imAgfRecLen As Integer      'AGF record length
Dim tmAgf As AGF
Dim tmAgfSrchKey As INTKEY0
Dim hmGrf As Integer            'User file handle
Dim imGrfRecLen As Integer      'GRF record length
Dim tmGrf As GRF
Dim hmChf As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imChfRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmIsr As Integer            'Inv temp file handle
Dim imIsrRecLen As Integer      'ISR record length
Dim tmIsr As ISR
Dim tmIsrSrchKey As ISRKEY0
Dim hmSlf As Integer            'Slsp file handle
Dim imSlfRecLen As Integer      'SLF record length
Dim tmSlf As SLF
Dim tmSlfSrchKey As INTKEY0
Dim hmSof As Integer            'sales office file handle
Dim imSofRecLen As Integer      'SOF record length
Dim tmSof As SOF
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer      'VEF record length
'  Receivables File
Dim tmRvf As RVF
Dim hmSdf As Integer            'Sdf file handle
Dim imSdfRecLen As Integer      'SDF record length
Dim tmSdf As SDF
Dim tmSdfSrchKey0 As SDFKEY0    'veh, cntr, line date, status, time
Dim tmSdfSrchKey1 As SDFKEY1    'veh, date, time, status
Dim hmSmf As Integer            'Smf file handle
Dim imSmfRecLen As Integer      'SMF record length
Dim tmSmf As SMF
Dim imConsolidate As Integer        '4-4-02 true if no splits, false if participant splits
Dim smGrossOrNet As String * 1  'G = gross, N = net
Dim smCashAgyComm As String     'xx.xx as agy comm
Dim imMatchSSCode As Integer    'sales source for current contract
Dim imLastBilled(0 To 1) As Integer        'last billed date, btrieve format
Dim imCkcAll As Integer         'all participants selected
Dim imMnfCodes() As Integer     'array of valid participants selected
Dim imYrOrCnt As Integer
'Overall contract package & airing $
'Dim tmPkRvf() As RVF_VEF
'Dim tmAirRvf() As RVF_VEF
'Package Cash & Trade $
Dim tmPkRvfCash() As RVF_VEF
Dim tmPkRvfTrade() As RVF_VEF
'Airing Cash & Trade $
Dim tmAirRvfCash() As RVF_VEF
Dim tmAirRvfTrade() As RVF_VEF
Dim tmRvfSort() As RVFSORT      'all Cash & Trade transactions from RVF/phf
Dim tmSofList() As SOFLIST      'array of sales offices and associated sales sources
Dim tmSdfList() As SDFSORTLIST
'
'
'
'           Generate Earned Distribution Report.
'           This report is only for those clients that Bill by "Line" (vs week)
'           This shows how much money the producers earn in the past, as well
'           as in the future.
'           All receivables (prior to last billing period is obtained), plus
'           the spot data in the future is used to calulate future earnings.
'
'           This report is intended to be run when the Billing method is Show Aired, update aired,
'           and balance contract by line (vs week)
'
'
'           1. Array of start dates are built into 15 buckets.  The first date is the earliest date possible (1/1/1970).
''             Buckets 2-14 contain the start date for the 12-month period requested.  Bucket 15 is the latest
''             possible date (12/31/2026).
'           2. All Receivables are gathered from Phf/RVf for the 12 month period, in addition to everything billed
'              prior to the 12 month period. The phf/rvf records are stored in array TMRVFSORT.  All transactions
'              are sorted by contract #.
'           3. The active contracts are gathered based on the period requested, going back
'              1 month for spots booked prior to the contract start (mgs).
'           4. The entire SDF is read finding (from the earliest start date of all the active contracts gathered.
'              Each spot is written to the ISR file so it can be sorted by Contract & Line.  We choose not to gather
'              the spots by contract and vehicle because it took too long.
'              As each contract is processed, a table containing a start & end index (tlStartEndInx) of each schedule line
'              from array tmSDFList where spots are stored for each contract.
'           5.  The contracts are processed one at a time. The $ invoiced (from TMRVFSORT) are built into
'               the package (tmPkRVFCash & tmPkRvfTrade) & airing arrays (tmAirRVFCash & tmAirRvfTrade).
'               The ordered information is obtained from the  schedule lines and built into same arrays.
'               Cash & Trade are required because the report shows cash & trade distributions by vehicle.
'           6.  For each line, the SDF is obtained by contract for the entire SDF and aired $ are built into tmAirRvf
'           7.  calculations are made to compute each airing vehicles billed $, and written to GRF
'
''
'           D.Hosaka 8-28-01
'
'
'           3-7-02 avoid subscript out of range to check RVF that havent been processed (ilmissingchf = 1 pass)
'                  Select contract to balance to order didnt gather spots when everything was in the future and
'                  the end date was only in the current year
'           4-9-02 setup proper commission when split cash/trade

Sub gCREarnedDistr()
    ReDim ilNowTime(0 To 1) As Integer  'end time of report
    ReDim ilNowDate(0 To 1) As Integer
    Dim llETime As Long
    Dim llSTime As Long
    Dim ilRet As Integer
    Dim ilLoopChf  As Integer           'loop index by contract
    Dim ilClf As Integer                'loop index for sched lines
    Dim ilFound As Integer
    Dim ilPass As Integer               'outer loop to process contracts is 2 passes:  1st to find all packages (ordered) lines,
                                        'pass 2 is to process all std and hidden lines
    Dim ilPkRvf As Integer              'loop to go thru array of vehicle $
    Dim llContrCode As Long
    Dim ilLoop As Integer
    Dim ilSdfIndex As Integer
    Dim llSingleCntr As Long
    Dim ilQtr As Integer                'user quarter requested
    Dim ilYear As Integer               'user year requested
    ReDim llTempStdDates(1 To 13) As Long 'start dates of the std 12 month period
    ReDim llStdStartDates(1 To 15) As Long    'array of std start month dates:  1 = earliest date possible for all past prior to requested 12-month period (1/1/1970),
                                        '2-14 = std start months of 12 month period
                                        '15 = latest date possible (12/31/2026) for all $ in the future after 12 month requested period
    ReDim llProject(1 To 15) As Long    '$ from package line
    Dim llLastBilled As Long        'last date invoiced to dtermine past/future
    Dim ilLastBilledInx As Integer  'last month inx invoiced (adjusted by +1 due to 1st bucket of array is past: then next 12 is each month)
    Dim llLBPlus1Month As Long      'last billed plus 1 std month to account for mgs
    Dim slTempStart As String       '1st std broadcast month start date calculated from qtr & year requested
    Dim slTempEnd As String         'last std broadcast month to gather active contrcts
    Dim slCntrStatus As String          'statues to include for contract processing
    Dim slCntrTypes As String        'contract types to include for processing
    Dim ilHOState As Integer
    Dim ilStartSearch As Integer
    Dim ilEndSearch As Integer
    Dim llSDate As Long
    Dim llEDate As Long
    Dim slStr As String
    Dim slTemp As String
    Dim llTemp As Long
    Dim ilWhichKey As Integer
    Dim ilUpperSdf As Integer
    Dim llStartofRpt As Long
    Dim llEndOfRpt As Long
    Dim ilMissingChf As Integer
    ReDim tlChfAdvtExt(1 To 1) As CHFADVTEXT  'array of active contract for the requested 12 month period

    lgTotal_ISRRecs = 0     '3-28-02init # records created in ISR.  When rept completed, if # records created match number records in file,
    'the file will be overlayed by a blank image.  Otherwise, records are removed by deletion.


    slTempStart = RptSelED!edcQtr.Text
    ilQtr = Val(slTempStart)            'get the value of the qtr requested
    slTempStart = RptSelED!edcSelCFrom.Text
    ilYear = Val(slTempStart)           'get the value of the year requested

    slTempStart = RptSelED!edcContract  'single contract # requested
    llSingleCntr = Val(slTempStart)

    imConsolidate = True          '4-4-02 assume no participant splits
    If RptSelED!rbcSelC4(0).Value Or RptSelED!rbcSelC4(1).Value Then  '4-4-02 Participant gross or net splits
        imConsolidate = False      'do the participant splits
    End If

    smGrossOrNet = "G"

    If RptSelED!rbcSelC4(1).Value Or RptSelED!rbcSelC4(3).Value Then  '4-3-02 Participant Net or Consolidated Net
        smGrossOrNet = "N"
    End If
    imCkcAll = True
    If RptSelED!ckcAll.Value = vbUnchecked Then
        imCkcAll = False
    End If
    imYrOrCnt = 1           '1=show by year (12-month period only), 2 = show by contract (show past&future)
    If RptSelED!rbcEarnCnt(1).Value Then
        imYrOrCnt = 2
    End If
    'open all applicable files, and setup array of monthly start dates
    '
    If mOpenProdFiles() = 0 Then
        gGetMonthsForYr 2, ilQtr, ilYear, llTempStdDates(), llLastBilled, ilLastBilledInx    'build array of corp start & end dates
        'place 12 month period dates into array that contains the earliest and latest dates for past & future
        llStdStartDates(1) = gDateValue("1/1/1970")
        For ilLoop = 1 To 13
            llStdStartDates(ilLoop + 1) = llTempStdDates(ilLoop)
        Next ilLoop
        llStdStartDates(15) = gDateValue("12/31/2026")

        'ilLastBilledInx = ilLastBilledInx '+ 1       'adjust for the 1st bucket that is the months prior to the year requested

        For ilLoop = 1 To 14 Step 1
            If llLastBilled > llStdStartDates(ilLoop) And llLastBilled < llStdStartDates(ilLoop + 1) Then
                ilLastBilledInx = ilLoop
                Exit For
            End If
        Next ilLoop


        'convert last billed to btrieve format
        gPackDateLong llLastBilled, imLastBilled(0), imLastBilled(1)
        llLBPlus1Month = llStdStartDates(ilLastBilledInx + 1)     'used to determine if line has expired prior to last billing period;
                                                                 'but extend it to one month later because of makegoods
        'build array of selling office codes and their sales sources.
        ilLoop = 0
        ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            ReDim Preserve tmSofList(0 To ilLoop) As SOFLIST
            tmSofList(ilLoop).iSofCode = tmSof.iCode
            tmSofList(ilLoop).iMnfSSCode = tmSof.iMnfSSCode
            ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            ilLoop = ilLoop + 1
        Loop
    Else
        On Error GoTo mOpenProdErr:
    End If
    If Not imCkcAll Then           'build array of selected participants
        ReDim imMnfCodes(0 To 0) As Integer
        For ilPass = 0 To RptSelED!lbcSelection(0).ListCount - 1 Step 1
            If RptSelED!lbcSelection(0).Selected(ilPass) Then              'selected participant
                slTemp = tgVehicle(ilPass).sKey
                ilRet = gParseItem(slTemp, 2, "\", slStr)
                imMnfCodes(UBound(imMnfCodes)) = Val(slStr)
                ReDim Preserve imMnfCodes(0 To UBound(imMnfCodes) + 1)
            End If
        Next ilPass
    End If

    ' Find the contracts to process - either single or all active contracts for the 12-month period
    '
    If llSingleCntr > 0 Then          'get only contract # requested
        ReDim tlChfAdvtExt(1 To 2) As CHFADVTEXT   'fake out the array so its common code
        tmChfSrchKey1.lCntrNo = llSingleCntr
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If tmChf.lCntrNo <> llSingleCntr Then
            Exit Sub
        Else                           'fake out an entry in tlchfadvtext  for the single contract
            tlChfAdvtExt(1).lCntrNo = llSingleCntr
            tlChfAdvtExt(1).lCode = tmChf.lCode
            tlChfAdvtExt(1).iStartDate(0) = tmChf.iStartDate(0)
            tlChfAdvtExt(1).iStartDate(1) = tmChf.iStartDate(1)
            tlChfAdvtExt(1).iEndDate(0) = tmChf.iEndDate(0)
            tlChfAdvtExt(1).iEndDate(1) = tmChf.iEndDate(1)
        End If
    Else
        'setup parameters for inclusion of contracts to process
        slCntrTypes = ""            'all contract types
        slTempStart = Format$((llStdStartDates(2)) - 35, "m/d/yy")
        slTempEnd = Format$((llStdStartDates(14)) - 1, "m/d/yy")
        slCntrStatus = "HOGN"             'include orders and uns orders
        ilHOState = 2
        'Gather the contracts to process for the year
        'build table (into tlchfadvtext) of all contracts that fall within the dates required
        'Back up the start date for gathering of contracts by one month to include makegoods in the
        'prior month
        ilRet = gObtainCntrForDate(RptSelED, slTempStart, slTempEnd, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
    End If

    'Find all application transactions (IN & AN) and have them sorted by contract #
    '
    ilRet = mGatherRVF(llSingleCntr, llStdStartDates())   'tmRvfSort array is returned
    If ilRet <> 0 Then
        MsgBox "Error in Creating Phf/Rvf", vbOkOnly + vbCritical + vbApplicationModal, "gCrEarnedDist"
        Exit Sub
    End If

    'Create an array that contains the contract and the starting and ending index within the RVF list
    'so that every time it goes thru the list of transactions it doesnt have to go thru all of them
    ReDim tlRvfInx(1 To 1) As STARTENDINX
    ilFound = False
    ilPass = 1
    For ilSdfIndex = LBound(tmRvfSort) To UBound(tmRvfSort) - 1
        If Not ilFound Then       'first time thru
            tlRvfInx(ilPass).lLineNo = tmRvfSort(ilSdfIndex).tlRvfRec.lCntrNo
            tlRvfInx(ilPass).iStartInx = ilSdfIndex
            tlRvfInx(ilPass).iEndInx = ilSdfIndex
            tlRvfInx(ilPass).iProcessed = 0
            ilFound = True
        Else
            If tmRvfSort(ilSdfIndex).tlRvfRec.lCntrNo <> tlRvfInx(ilPass).lLineNo Then
                tlRvfInx(ilPass).iEndInx = ilSdfIndex - 1
                ReDim Preserve tlRvfInx(1 To UBound(tlRvfInx) + 1) As STARTENDINX
                ilPass = UBound(tlRvfInx)
                tlRvfInx(ilPass).lLineNo = tmRvfSort(ilSdfIndex).tlRvfRec.lCntrNo
                tlRvfInx(ilPass).iStartInx = ilSdfIndex
                tlRvfInx(ilPass).iEndInx = ilSdfIndex
                tlRvfInx(ilPass).iProcessed = 0
            End If
        End If
    Next ilSdfIndex
    tlRvfInx(ilPass).iEndInx = UBound(tmRvfSort) - 1


    'common data for prepass record only, following fields dont change
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)
    'tmGrf.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    tmGrf.iDate(0) = imLastBilled(0)
    tmGrf.iDate(1) = imLastBilled(1)
    tmGrf.sBktType = smGrossOrNet
    'find the earliest & latest contract header dates from all the active contracts gathered.  Thats how far back & how far forward to read SDF
    llSDate = 0
    llEDate = 0
    For ilLoopChf = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1
        gUnpackDateLong tlChfAdvtExt(ilLoopChf).iStartDate(0), tlChfAdvtExt(ilLoopChf).iStartDate(1), llTemp
        If llSDate = 0 Then
           llSDate = llTemp
        ElseIf llTemp < llSDate And llTemp <> 0 Then
            llSDate = llTemp
        End If
        gUnpackDateLong tlChfAdvtExt(ilLoopChf).iEndDate(0), tlChfAdvtExt(ilLoopChf).iEndDate(1), llTemp
        If llEDate = 0 Then
           llEDate = llTemp
        ElseIf llTemp > llEDate And llTemp <> 0 Then
            llEDate = llTemp
        End If
    Next ilLoopChf
    llSDate = llSDate - 35          'backup for possible makegoods sch prior to start of order
    llEDate = llEDate + 35          'move out 1 month for possible mgs after end of order

    ilWhichKey = 1                          'assume key by date
    If llSingleCntr > 0 Then
        ilWhichKey = 0                      'key by contract
        llContrCode = tlChfAdvtExt(1).lCode
    End If


    If imYrOrCnt = 1 Then               'report by year
        If llStdStartDates(14) - 1 > llLastBilled Then       'if the end date of the year requested is greater than the last date billed-- SDF needs to be
                                                             'retrieved to calculate the remainder months to be invoiced
            ilRet = mBuildISRfromSDF(llSDate, llEDate, ilWhichKey, llContrCode)     'read SDF by vehicle using date, & time key; create records (excluding fills & bonus) into ISR
        End If
    Else                                'balance to contract

        If llEDate > llStdStartDates(14) - 1 Or llStdStartDates(14) - 1 > llLastBilled Then      'if end date of all contracts to process exceeds the end of requested year;
                                    'and the requested year end date is greater than the last month billed, need to read SDF to calculate remainder of months billing

            ilRet = mBuildISRfromSDF(llSDate, llEDate, ilWhichKey, llContrCode)     'read SDF by vehicle using date, & time key; create records (excluding fills & bonus) into ISR
        End If
    End If


    'debugging only for time spots took to gather
    'slStr = Format$(gNow(), "h:mm:ssAM/PM")       'end time of run
    'llSDFTime = gTimeToLong(slStr, False)
    'gUnpackTimeLong igNowTime(0), igNowTime(1), False, llSTime   'start time of run
    'llSDFTime = llSDFTime - llSTime              'time in seconds in gather SDF

    For ilMissingChf = 1 To 2           'pass 1 goes thru all contracts active for the year,
                                        'pass 2 goes thru all list of expired contracts with transactions
        'loop through all the active contracts and gather spots scheduled by each line
        For ilLoopChf = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1
            'Get the entire contract
            llContrCode = tlChfAdvtExt(ilLoopChf).lCode
            ilRet = gObtainCntr(hmChf, hmClf, hmCff, llContrCode, False, tgChfED, tgClfED(), tgCffED())

            '3-27-02 check for contracts that dont exist from RVF/PHF with invalid contr reference #
            If ilRet = False Then      'contract doesnt exist
                tgChfED.lCode = 0                 'invalid cntr code
                tgChfED.lCntrNo = tlChfAdvtExt(ilLoopChf).lCntrNo
                tgChfED.iAgfCode = tlChfAdvtExt(ilLoopChf).iAgfCode
                tgChfED.iSlfCode(0) = tlChfAdvtExt(ilLoopChf).iSlfCode(0)
                'make assumptions on transaction that it cash acct
                tgChfED.sAgyCTrade = "N"
                tgChfED.iPctTrade = 0
            End If

            'obtain agency for commission
            If tgChfED.iAgfCode > 0 Then
                tmAgfSrchKey.iCode = tgChfED.iAgfCode
                ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching agy recd
                If ilRet <> BTRV_ERR_NONE Then
                    tmAgf.iComm = 0
                End If
            Else                'direct, no agency commission
                tmAgf.iComm = 0
            End If              'iagfcode > 0
            '4-9-02 dont set up agency comm here because of split cash/trade accts
            'If tgChfED.iPctTrade > 0 And tgChfED.sAgyCTrade = "N" Then  'this trade is not commissionable
            '    tmAgf.iComm = 0
            'End If
            'smCashAgyComm = gIntToStrDec(tmAgf.iComm, 2)       'net amt is determined in mWriteGRF
            'retrieve the primary slsp for the sales source
            tmSlfSrchKey.iCode = tgChfED.iSlfCode(0)
            ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            For ilLoop = LBound(tmSofList) To UBound(tmSofList)
                If tmSofList(ilLoop).iSofCode = tmSlf.iSofCode Then
                    imMatchSSCode = tmSofList(ilLoop).iMnfSSCode          'Sales source
                    Exit For
                End If
            Next ilLoop
            'Gather the spots for this contract from ISR, which are keyed by contract
            ilRet = mGatherISRbyCnt()
            'Now sort the array by the line #
            ilUpperSdf = UBound(tmSdfList) - 1
            ArraySortTyp fnAV(tmSdfList(), 1), ilUpperSdf, 0, LenB(tmSdfList(1)), 0, LenB(tmSdfList(1).sKey), 0
            'Build the overall receivables total billed into Package & Airing arrays for this contract
            'tmPkRvf & tmAirRvf will determine the overall billing for future months (cash & trade combined)
            'The Cash & Trade $ must be maintained because the report needs to show the distribution of cash vs trade
            'ReDim tmPkRvf(0 To 0)  As RVF_VEF            'package information for overall contract
            'tmPkRvf contains $ from package and conventional lines ordered & invoiced:
            'ivefCode = package or conv. vehicle code
            'ipklineno = package line # (0 if conv)
            'lTotalOrd(1 to 14) - package or conv.$ ordered by month (from clf)
            'lTotalGross(1 to 14) - package or conv $ billed by month (phf/rvf) Cash only
            'lTotalVefDollars - total package or conv $ ordered (clf)
            'lTotalVefBilledDollars - total package or conv $ billed (phf/rvf)
            'ReDim tmAirRvf(0 To 0) As RVF_VEF            'airing information for overall contract
            'tmAirRvf contains $ from all airing spots & airing spots invoiced:
            'ivefCode = airing vehicle code
            'ipklineno = package line # reference
            'lTotalOrd(1 to 14) - $ aired by month (sdf)
            'lTotalGross(1 to 14) - airing $ billed by month (phf/rvf) Cash only
            'lTotalVefDollars - total airing $ (sdf)
            'lTotalVefBilledDollars - total airing $ billed (phf/rvf)
            ReDim tmPkRvfCash(0 To 0) As RVF_VEF         'Cash portion of package $
            ReDim tmPkRvfTrade(0 To 0) As RVF_VEF        'Trade portion of package $
            ReDim tmAirRvfCash(0 To 0) As RVF_VEF        'cash portion of airing $
            ReDim tmAirRvfTrade(0 To 0) As RVF_VEF       'trade portion of airing $

            'find the starting and ending index to the transactions to process for this contract
            ilStartSearch = 1
            ilEndSearch = 1
            For ilPkRvf = LBound(tlRvfInx) To UBound(tlRvfInx)
                If tgChfED.lCntrNo = tlRvfInx(ilPkRvf).lLineNo Then
                    ilStartSearch = tlRvfInx(ilPkRvf).iStartInx
                    ilEndSearch = tlRvfInx(ilPkRvf).iEndInx
                    tlRvfInx(ilPkRvf).iProcessed = 1        'set this contract as processed from RVF/PHF
                    Exit For
                End If
            Next ilPkRvf

            mBuildRVFSummary ilStartSearch, ilEndSearch, llStdStartDates()

            'Create an array that contains the line # and the starting and ending index within the spot list
            'so that every time it goes thru the list of contracts it doesnt have to go thru all the spots
            ReDim tlStartEndInx(1 To 1) As STARTENDINX
            ilFound = False
            ilPass = 1
            For ilSdfIndex = LBound(tmSdfList) To UBound(tmSdfList) - 1
                If Not ilFound Then       'first time thru
                    tlStartEndInx(ilPass).lLineNo = tmSdfList(ilSdfIndex).tSdf.iLineNo
                    tlStartEndInx(ilPass).iStartInx = ilSdfIndex
                    tlStartEndInx(ilPass).iEndInx = ilSdfIndex
                    ilFound = True
                Else
                    If tmSdfList(ilSdfIndex).tSdf.iLineNo <> tlStartEndInx(ilPass).lLineNo Then
                        tlStartEndInx(ilPass).iEndInx = ilSdfIndex - 1
                        ReDim Preserve tlStartEndInx(1 To UBound(tlStartEndInx) + 1) As STARTENDINX
                        ilPass = UBound(tlStartEndInx)
                        tlStartEndInx(ilPass).lLineNo = tmSdfList(ilSdfIndex).tSdf.iLineNo
                        tlStartEndInx(ilPass).iStartInx = ilSdfIndex
                        tlStartEndInx(ilPass).iEndInx = ilSdfIndex
                    End If
                End If
            Next ilSdfIndex
            tlStartEndInx(ilPass).iEndInx = UBound(tmSdfList) - 1
            'For ilPass = 1 To 2
                For ilClf = LBound(tgClfED) To UBound(tgClfED) - 1
                    tmClf = tgClfED(ilClf).ClfRec
                    gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llSDate
                    gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llEDate
                    'If ilPass = 1 Then          'pass 1 goes thru all schedule lines looking for the ordered $ and conventional ordered
                                                'to build a separate array using tmPkRvf (hidden line $ are built in array tmAirRvf)
                        If (tmClf.sType = "O" Or tmClf.sType = "S") And (llSDate <= llEDate) Then        'package or std line & not CBS, generate the $ ordered, bypass all other line types
                            mGetOrderedDollars ilClf, llStdStartDates()
                        End If
                    'Else        'pass 2 builds all the hidden/std line $ from sdf
                        If (tmClf.sType = "S" Or tmClf.sType = "H") And (llSDate <= llEDate) Then '3-25-02 And llEDate + 35 > llLastBilled Then       'standard or hidden line
                            'Loop thru the spot index table finding the starting and ending indices to use within the array of spots
                            For ilSdfIndex = 1 To UBound(tlStartEndInx)
                                If tmClf.iLine = tlStartEndInx(ilSdfIndex).lLineNo Then
                                    mGetAiredDollars llStdStartDates(), tlStartEndInx(ilSdfIndex).iStartInx, tlStartEndInx(ilSdfIndex).iEndInx
                                    Exit For
                                End If
                            Next ilSdfIndex
                            'mGetAiredDollars llStdStartDates()
                        End If              'endif tmClf.sType = "S" or  "H"
                    'End If                  'endif ilpass =1
                Next ilClf                  'For ilClf LBound(tgClf) to UBound(tgClf)
            'Next ilPass                     'for ilPass = 1 to 2

            'Create one record per vehicle & participant & cash/Trade for this contract
            If tgChfED.iPctTrade = 0 Or tgChfED.iPctTrade <> 100 Then       'all cash or split cash & trade
                '4-9-02 determine comm for cash portion
                If tgChfED.iAgfCode > 0 Then      'if there is an associated agy, use the commission from agy
                    smCashAgyComm = gIntToStrDec(tmAgf.iComm, 2)       'net amt is determined in mWriteGRF
                Else                'direct, no agency commission
                    smCashAgyComm = "0.00"
                End If              'iagfcode > 0
                mWriteCashOrTrade ilLastBilledInx, llStdStartDates(), "C", tmAirRvfCash(), tmPkRvfCash()
            End If
            If tgChfED.iPctTrade = 100 Or tgChfED.iPctTrade <> 100 Then   'all trade, or split cash & trade
                '4-9-02 determine comm for trade portion
                If tgChfED.iPctTrade > 0 And tgChfED.sAgyCTrade = "N" Then  'this trade is not commissionable
                    smCashAgyComm = "0.00"
                Else
                    smCashAgyComm = gIntToStrDec(tmAgf.iComm, 2)       'net amt is determined in mWriteGRF
                End If
                mWriteCashOrTrade ilLastBilledInx, llStdStartDates(), "T", tmAirRvfTrade(), tmPkRvfTrade()
            End If
        Next ilLoopChf                      'for ilLoopChf = LBound(tgChfAdvtExt) to UBound(tgChfAdvtExt)


        If ilMissingChf = 1 Then            '1st pass completed, see if any transactions didn't get processed due
                        'to the contract already expired and not picked up
            'See if there were any RVF/PHF transactions that didn't get processed
            ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
            ilLoop = LBound(tlChfAdvtExt)
            For ilPkRvf = LBound(tlRvfInx) To UBound(tlRvfInx)
                If tlRvfInx(ilPkRvf).iProcessed = 0 Then        '0=not processed
                    'Cycle thru the transactions of the contract not processed to see if the tran date is with
                    'the year requested.  If not, the contract has expired so forget it
                    ilStartSearch = tlRvfInx(ilPkRvf).iStartInx
                    ilEndSearch = tlRvfInx(ilPkRvf).iEndInx
                    ilFound = False
                    If ilStartSearch >= LBound(tmRvfSort) Then      '3-7-02 avoid subscript out of range

                    For ilYear = ilStartSearch To ilEndSearch
                        'Examine the transaction date
                        tmRvf = tmRvfSort(ilYear).tlRvfRec
                        gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llTemp
                        'determine if in year requested
                        If llTemp >= llStdStartDates(2) And llTemp < (llStdStartDates(14)) Then
                            ilFound = True
                            'Exit For
                        End If

                        If ilFound Then
                            'Build the contract info into tlChfAdvtExt
                            tmChfSrchKey1.lCntrNo = tlRvfInx(ilPkRvf).lLineNo
                            tmChfSrchKey1.iCntRevNo = 32000
                            tmChfSrchKey1.iPropVer = 32000
                            ilRet = btrGetGreaterOrEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd

                            Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tlRvfInx(ilPkRvf).lLineNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")
                                ilRet = btrGetNext(hmChf, tmChf, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                            If tmChf.lCntrNo = tlRvfInx(ilPkRvf).lLineNo Then
                                ilYear = ilEndSearch       'force to get out of loop, found at least once tran that belongs in requested year that has not been processed
                                tlChfAdvtExt(ilLoop).lCode = tmChf.lCode
                                ilLoop = ilLoop + 1
                                ReDim Preserve tlChfAdvtExt(ilLoop) As CHFADVTEXT
                            Else                        '3-27-02 invalid contract #, contract not found.  Fake out an entry so that the unprocessed transactions
                                'are picked up (rvf entered with invalid cntr #)
                                ilYear = ilEndSearch            'force to get out of loop
                                tlChfAdvtExt(ilLoop).lCode = 0
                                tlChfAdvtExt(ilLoop).lCntrNo = tmRvf.lCntrNo
                                tlChfAdvtExt(ilLoop).iSlfCode(0) = tmRvf.iSlfCode
                                tlChfAdvtExt(ilLoop).iAgfCode = tmRvf.iAgfCode
                                ilLoop = ilLoop + 1
                                ReDim Preserve tlChfAdvtExt(ilLoop) As CHFADVTEXT
                            End If
                        End If
                    Next ilYear
                    End If
                End If
            Next ilPkRvf
        End If
    Next ilMissingChf
    'Erase tmPkRvf, tmPkRvfCash, tmPkRvfTrade
    'Erase tmAirRvf, tmAirRvfCash, tmAirRvfTrade
    Erase tmPkRvfCash, tmPkRvfTrade
    Erase tmAirRvfCash, tmAirRvfTrade
    Erase tlChfAdvtExt, tmSofList, imMnfCodes
    Erase llProject
    Erase llStdStartDates
    Erase llTempStdDates
    Erase tmSdfList

    Erase tlRvfInx
    Erase tlStartEndInx
    Erase tmRvfSort

    Erase tgClfED, tgCffED
    mCloseProdFiles
    'debugging only for time program took to run
    slStr = Format$(gNow(), "h:mm:ssAM/PM")       'end time of run
    gPackTime slStr, ilNowTime(0), ilNowTime(1)   'time report ended
    slStr = Format$(gNow(), "m/d/yy")
    gPackDate slStr, ilNowDate(0), ilNowDate(1)

    gUnpackDateLong igNowDate(0), igNowDate(1), llStartofRpt
    gUnpackDateLong ilNowDate(0), ilNowDate(1), llEndOfRpt
    gUnpackTimeLong ilNowTime(0), ilNowTime(1), False, llETime
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, llSTime   'start time of run
    If llStartofRpt = llEndOfRpt Then
        llSTime = llETime - llSTime
    Else
        llSTime = (86400 - llSTime) + llETime
    End If
    ilRet = gSetFormula("RunTime", llSTime)  'show how long report generated

    Exit Sub
mOpenProdErr:
    gBtrvErrorMsg ilRet, "gCREarnedDist (btrOpen): RptCrEd", RptSelED
    mCloseProdFiles
    Exit Sub
        
    mCloseProdFiles
    Exit Sub
End Sub
'
'
'
Sub mAddRvfAir(ilVefCode As Integer, ilPkLineNo As Integer, llAmt As Long, ilMonthInx As Integer, tlAirRvf() As RVF_VEF)
Dim ilIndex As Integer
Dim ilPkRvf As Integer
Dim ilLoop As Integer
    ilIndex = -1
    For ilPkRvf = LBound(tlAirRvf) To UBound(tlAirRvf) - 1 Step 1
        If (tlAirRvf(ilPkRvf).iVefCode = ilVefCode And tlAirRvf(ilPkRvf).iPkLineNo = ilPkLineNo) Then
            ilIndex = ilPkRvf
            Exit For
        End If
    Next ilPkRvf

    If ilIndex = -1 Then
        ilPkRvf = UBound(tlAirRvf)
        tlAirRvf(ilPkRvf).iVefCode = ilVefCode
        tlAirRvf(ilPkRvf).iPkLineNo = ilPkLineNo
        tlAirRvf(ilPkRvf).lTotalVefBilledDollars = 0
        For ilLoop = 1 To 14
            tlAirRvf(ilPkRvf).lTotalGross(ilLoop) = 0
        Next ilLoop
        ReDim Preserve tlAirRvf(0 To ilPkRvf + 1) As RVF_VEF
    Else
        ilPkRvf = ilIndex
    End If
    tlAirRvf(ilPkRvf).lTotalVefDollars = tlAirRvf(ilPkRvf).lTotalVefDollars + llAmt
    tlAirRvf(ilPkRvf).lTotalOrd(ilMonthInx) = tlAirRvf(ilPkRvf).lTotalOrd(ilMonthInx) + llAmt
End Sub
'
'
'               Accumulate Invoiced amount into the overall contract arrays, or the
'               Cash or Trade arrays
'
'               mAddRvf: <input>  ilVefCode - billing or airing vehicle code
'                                 ilPkLineNo - package line # (or 0 if conventional)
'                                 llAmt - gross or net amount
'                                 ilMonthInx - month index to add amount
'                        <output> tlRvfVef() - updated array
Sub mAddRvfInv(ilVefCode As Integer, ilPkLineNo As Integer, llAmt As Long, ilMonthInx As Integer, tlRvfVef() As RVF_VEF)
Dim ilFound As Integer
Dim ilRvfByVef As Integer
Dim ilLoop As Integer
    ilFound = False
    For ilRvfByVef = LBound(tlRvfVef) To UBound(tlRvfVef) - 1   'array containing vehicle $ totals for packages(from a PHF/RVF)
        'find matching pkg veh & line
        If ilVefCode = tlRvfVef(ilRvfByVef).iVefCode And ilPkLineNo = tlRvfVef(ilRvfByVef).iPkLineNo Then
            ilFound = True
            tlRvfVef(ilRvfByVef).lTotalVefBilledDollars = tlRvfVef(ilRvfByVef).lTotalVefBilledDollars + llAmt   'all $ invoiced so far
            tlRvfVef(ilRvfByVef).lTotalGross(ilMonthInx) = tlRvfVef(ilRvfByVef).lTotalGross(ilMonthInx) + llAmt   'Cash: get each months billing
            Exit For
        End If
    Next ilRvfByVef
    If Not ilFound Then
        'first time for this vehicle & package line #
        'accumulate the $ into this entry
        tlRvfVef(ilRvfByVef).iVefCode = ilVefCode       'billing or airing vehicle
        tlRvfVef(ilRvfByVef).lTotalVefBilledDollars = llAmt    ' $ invoiced
        tlRvfVef(ilRvfByVef).iPkLineNo = ilPkLineNo
        For ilLoop = 1 To 14
            tlRvfVef(ilRvfByVef).lTotalGross(ilLoop) = 0
        Next ilLoop
        tlRvfVef(ilRvfByVef).lTotalGross(ilMonthInx) = llAmt    'get each months billing
        ReDim Preserve tlRvfVef(0 To UBound(tlRvfVef) + 1)
    End If
End Sub
'
'
'               Accumulate ordered amount into the overall contract arrays, or the
'               Cash or Trade arrays
'
'               mAddRvf: <input>  ilVefCode - billing or airing vehicle code
'                                 ilPkLineNo - package line # (or 0 if conventional)
'                                 llAmt() - gross or net amount for past, current year, and future
'                        <output> tlRvfVef() - updated array
Sub mAddRvfOrd(ilVefCode As Integer, ilPkLineNo As Integer, llProject() As Long, tlRvfVef() As RVF_VEF)
Dim ilIndex As Integer
Dim ilPkRvf As Integer
Dim ilLoop As Integer
    ilIndex = -1
    For ilPkRvf = LBound(tlRvfVef) To UBound(tlRvfVef) - 1 Step 1
        If (tlRvfVef(ilPkRvf).iVefCode = ilVefCode And tlRvfVef(ilPkRvf).iPkLineNo = ilPkLineNo) Then
            ilIndex = ilPkRvf
            Exit For
        End If
    Next ilPkRvf
    If ilIndex = -1 Then             'first time for this package or std line
        'initialize the variables
        ilPkRvf = UBound(tlRvfVef)
        tlRvfVef(ilPkRvf).iVefCode = ilVefCode
        tlRvfVef(ilPkRvf).iPkLineNo = ilPkLineNo
        tlRvfVef(ilPkRvf).lTotalVefBilledDollars = 0
        tlRvfVef(ilPkRvf).lTotalVefDollars = 0
        For ilLoop = 1 To 14                        'store the ordered $ for past, months 1-12 & future
            tlRvfVef(ilPkRvf).lTotalOrd(ilLoop) = 0
        Next ilLoop
        ReDim Preserve tlRvfVef(0 To ilPkRvf + 1) As RVF_VEF
    Else
        ilPkRvf = ilIndex
    End If
    'Accumulate total $  for vehicle
    For ilLoop = 1 To 14                        'store the ordered $ for past, months 1-12 & future
        tlRvfVef(ilPkRvf).lTotalOrd(ilLoop) = tlRvfVef(ilPkRvf).lTotalOrd(ilLoop) + llProject(ilLoop)
        tlRvfVef(ilPkRvf).lTotalVefDollars = tlRvfVef(ilPkRvf).lTotalVefDollars + llProject(ilLoop)  'total $ ordered
    Next ilLoop

End Sub
'
'
'           mBuildISRFromSDF - cycle thru the SDF file by vehicle, using the earliest date
'           gathered from all the active contracts to process.  Ignore Fills & Bonus spots since
'           they are all $0.  Create the ISR so that they can be sorted by contract code, then line #
'
'           <input> llSDate - earliest date to obtain SDF
'                   llEDate - latest date to search SDF
'                   ilWhichKey - 0 = key 0 for selective , else 1 for key 1 by date
'                   llContrCode - Contract Code: only applies to ilWhichKey=0 (selective contract)
'           <output> ISR (prepass file sorted by contract code & line #
'
Function mBuildISRfromSDF(llSDate As Long, llEDate As Long, ilWhichKey As Integer, llSelChfCode As Long) As Integer
Dim ilVefCode As Integer
Dim ilLoopVef As Integer
Dim ilRet As Integer
Dim ilExtLen As Integer
Dim llNoRec As Long
Dim ilOffset As Integer
Dim slStr As String
Dim slTemp As String
Dim llContrCode As Long
Dim llRecPos As Long
ReDim ilEarliestSDF(0 To 1) As Integer
ReDim ilLatestSDF(0 To 1) As Integer
Dim tlDateTypeBuff As POPDATETYPE   'Type field record
Dim tlIntTypeBuff As INTKEY0
Dim tlLongTypeBuff As LONGKEY0
Dim ilKeyFound As Integer
    mBuildISRfromSDF = BTRV_ERR_NONE
    gPackDateLong llSDate, ilEarliestSDF(0), ilEarliestSDF(1)
    gPackDateLong llEDate, ilLatestSDF(0), ilLatestSDF(1)
    For ilLoopVef = LBound(tgMVef) To UBound(tgMVef) - 1
        If tgMVef(ilLoopVef).sType = "S" Or tgMVef(ilLoopVef).sType = "C" Then      'selling or conventional vehicles would have spots sched
            ilVefCode = tgMVef(ilLoopVef).iCode
            'Loop thru each selling and conventional vehicles for all spot data based on the earliest date
            'of all the contracts gathered
            btrExtClear hmSdf   'Clear any previous extend operation
            imSdfRecLen = Len(tmSdf)
            ilKeyFound = False
            'Setup key for first time this vehicle
            If ilWhichKey = 0 Then              'selective cnt
                tmSdfSrchKey0.iVefCode = ilVefCode
                tmSdfSrchKey0.lChfCode = llSelChfCode
                tmSdfSrchKey0.iLineNo = 0
                tmSdfSrchKey0.lFsfCode = 0
                tmSdfSrchKey0.iDate(0) = 0
                tmSdfSrchKey0.iDate(1) = 0
                tmSdfSrchKey0.sSchStatus = " "
                tmSdfSrchKey0.iTime(0) = 0
                tmSdfSrchKey0.iTime(1) = 0
                'access by vheicle, date, time, sch status
                ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                If tmSdf.lChfCode = llSelChfCode Then
                    ilKeyFound = True
                End If
            Else
                tmSdfSrchKey1.iVefCode = ilVefCode
                tmSdfSrchKey1.iDate(0) = ilEarliestSDF(0)
                tmSdfSrchKey1.iDate(1) = ilEarliestSDF(1)
                tmSdfSrchKey1.iTime(0) = 0
                tmSdfSrchKey1.iTime(1) = 0
                tmSdfSrchKey1.sSchStatus = ""
                'access by vheicle, date, time, sch status
                ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                ilKeyFound = True   'if any spots in this vehicle, they must be used within year requested
            End If

            If ilRet <> BTRV_ERR_END_OF_FILE And ilKeyFound Then
                ilExtLen = Len(tmSdf)  'Extract operation record size
                llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
                Call btrExtSetBounds(hmSdf, llNoRec, -1, "UC", "SDF", "") '"EG") 'Set extract limits (0 = max skipped, )

                tlIntTypeBuff.iCode = ilVefCode
                ilOffset = gFieldOffset("Sdf", "SdfvefCode")
                ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffset, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)

                'tlStrTypeBuff.sType = "X"          'filter the fills & bonus spots later, dont want any rejects
                'ilOffSet = gFieldOffset("Sdf", "SdfSpotType")
                'ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_NOT_EQUAL, BTRV_EXT_AND, tlStrTypeBuff, 1)

                If ilWhichKey = 0 Then                           'selective cnt, keyed by cnt code
                    tlLongTypeBuff.lCode = llSelChfCode
                    ilOffset = gFieldOffset("Sdf", "SdfchfCode")
                    ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlLongTypeBuff, 4)
                Else
                    tlDateTypeBuff.iDate0 = ilLatestSDF(0)                       'dont go past this end date
                    tlDateTypeBuff.iDate1 = ilLatestSDF(1)
                    ilOffset = gFieldOffset("Sdf", "SdfDate")
                    ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)

                    tlDateTypeBuff.iDate0 = ilEarliestSDF(0)                       'retrieve past  projection records
                    tlDateTypeBuff.iDate1 = ilEarliestSDF(1)
                    ilOffset = gFieldOffset("Sdf", "SdfDate")
                    ilRet = btrExtAddLogicConst(hmSdf, BTRV_KT_DATE, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                End If
                ilRet = btrExtAddField(hmSdf, 0, ilExtLen)  'Extract the whole record
                On Error GoTo mBuildIsrErr
                gBtrvErrorMsg ilRet, "gObtainSdf (btrExtAddField):" & "Sdf.Btr", RptSelED
                On Error GoTo 0
                ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
                    On Error GoTo mBuildIsrErr
                    gBtrvErrorMsg ilRet, "gObtainSdf (btrExtGetNextExt):" & "Sdf.Btr", RptSelED
                    On Error GoTo 0
                    ilExtLen = Len(tmSdf)  'Extract operation record size
                    Do While ilRet = BTRV_ERR_REJECT_COUNT
                        ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                    Loop
                    Do While ilRet = BTRV_ERR_NONE
                        If tmSdf.sSpotType <> "X" Then
                            llContrCode = tmSdf.lChfCode
                            slStr = Trim$(Str$(llContrCode))
                            Do While Len(slStr) < 10
                                slStr = "0" & slStr
                            Loop

                            slTemp = Trim$(Str$(tmSdf.iLineNo))
                            Do While Len(slTemp) < 5
                                slTemp = "0" & slTemp
                            Loop

                            'Create the ISR temporary file containing spots,which will be keyed by contract code & line #
                            tmIsr.iGenDate(0) = igNowDate(0)
                            tmIsr.iGenDate(1) = igNowDate(1)
                            tmIsr.iGenTime(0) = igNowTime(0)
                            tmIsr.iGenTime(1) = igNowTime(1)
                            'gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                            'tmIsr.lGenTime = lgNowTime
                            tmIsr.sKey = slStr      '& "|" & slTemp       'contr code & line #
                            tmIsr.lChfCode = llContrCode
                            tmIsr.iLineNo = tmSdf.iLineNo
                            tmIsr.iVefCode = tmSdf.iVefCode
                            tmIsr.lCode = tmSdf.lCode
                            tmIsr.sSchStatus = tmSdf.sSchStatus
                            tmIsr.sPriceType = tmSdf.sPriceType      'Spot type "x" have been excluded
                            tmIsr.sType = tmSdf.sSpotType
                            gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), tmIsr.lCntrNo   'use the contr # field for the date


                            ilRet = btrInsert(hmIsr, tmIsr, imIsrRecLen, INDEXKEY0)
                            lgTotal_ISRRecs = lgTotal_ISRRecs + 1           '3-28-02
                            If ilRet <> BTRV_ERR_NONE Then
                                On Error GoTo mBuildIsrErr
                                gBtrvErrorMsg ilRet, "gObtainSdf (btrExtGetNextExt):" & "Sdf.Btr", RptSelED
                                On Error GoTo 0
                            End If
                        End If
                        ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                        Do While ilRet = BTRV_ERR_REJECT_COUNT
                            ilRet = btrExtGetNext(hmSdf, tmSdf, ilExtLen, llRecPos)
                        Loop
                    Loop
                End If
            End If
        End If                  'stype = "S" or stype = "C"
    Next ilLoopVef              'for LBound(tgMVef) to uBound(tgMvef)
    Exit Function
mBuildIsrErr:
    gBtrvErrorMsg ilRet, "mBuildISRFromSdf : RptCrEd", RptSelED
    mBuildISRfromSDF = ilRet
End Function
'
'
'           mBuildRVFSummary -
'           <input>  ilStartSearch - beginning index to start searching the RVF transaction array
'                    ilEndSearch - ending index to stop searching rvf trans array
'           <output> tmPkRvf() array
'           All transactions for the requested 12 months have been read & sorted into array tmRVFSORT.
'           Find the transactions that match the contract currently being procesed and
'           accumulate all $ invoiced so far
'
'           using:  tmRvfSort() - array of PHF/RVF (IN & AN) for requested 12 months stored by contract #
'
Sub mBuildRVFSummary(ilStartSearch As Integer, ilEndSearch As Integer, llStdStartDates() As Long)
Dim ilRvf As Integer
Dim llAmt As Long
Dim llSDate As Long
Dim ilLoop As Integer
Dim ilMonthInx As Integer
Dim tlRvf As RVF


    For ilRvf = ilStartSearch To ilEndSearch
        tlRvf = tmRvfSort(ilRvf).tlRvfRec
        If tlRvf.lCntrNo = tgChfED.lCntrNo Then
            gUnpackDateLong tlRvf.iTranDate(0), tlRvf.iTranDate(1), llSDate
            'determine the month this spot belongs in
            For ilLoop = 1 To 14            'it has to go into one since there are buckets to store prior and future spots (to the requested period)
                If llSDate >= llStdStartDates(ilLoop) And llSDate < (llStdStartDates(ilLoop + 1)) Then
                    ilMonthInx = ilLoop
                    Exit For
                End If
            Next ilLoop
            'accumulate the $ into this entry
            gPDNToLong tlRvf.sGross, llAmt   'always use the gross and adjust before writing prepass
            ilStartSearch = ilRvf + 1
            'Accumulate the package billing for the overall contract
            'mAddRvfInv tlRvf.iBillVefCode, tlRvf.iPkLineNo, llAmt, ilMonthInx, tmPkRvf()
            'accumulate the airing billing for the overall contract
            'mAddRvfInv tlRvf.iAirVefCode, tlRvf.iPkLineNo, llAmt, ilMonthInx, tmAirRvf()
            'accumulate the package & airing for the cash portion
            If tlRvf.sCashTrade = "C" Then
                mAddRvfInv tlRvf.iBillVefCode, tlRvf.iPkLineNo, llAmt, ilMonthInx, tmPkRvfCash()
                mAddRvfInv tlRvf.iAirVefCode, tlRvf.iPkLineNo, llAmt, ilMonthInx, tmAirRvfCash()
            'accumulate the package & airing for the trade portion
            Else
                mAddRvfInv tlRvf.iBillVefCode, tlRvf.iPkLineNo, llAmt, ilMonthInx, tmPkRvfTrade()
                mAddRvfInv tlRvf.iAirVefCode, tlRvf.iPkLineNo, llAmt, ilMonthInx, tmAirRvfTrade()
            End If
        End If
    Next ilRvf


End Sub
'
'
'           mCloseFiles - Close all applicable files for
'
Sub mCloseProdFiles()
Dim ilRet As Integer
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmChf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmIsr)
    btrDestroy hmGrf
    btrDestroy hmChf
    btrDestroy hmVef
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmSdf
    btrDestroy hmSmf
    btrDestroy hmAgf
    btrDestroy hmSlf
    btrDestroy hmSof
    btrDestroy hmIsr
End Sub
'
'
'
'                   mGatherISRbyCnt - find all records from ISR whose key matchines the
'                   current contract code
Function mGatherISRbyCnt() As Integer
Dim slStr As String
Dim llContrCode As Long
Dim llNoRec As Long
Dim ilExtLen As Integer
Dim ilRet As Integer
Dim llRecPos As Long
Dim ilChfCodeOffSet As Integer
Dim ilGenDateOffSet As Integer
Dim ilGenTimeOffSet As Integer
Dim tlDateTypeBuff As POPDATETYPE
Dim tlKeyTypeBuff As SORTCODE
ReDim tmSdfList(1 To 1) As SDFSORTLIST
    mGatherISRbyCnt = BTRV_ERR_NONE
    btrExtClear hmIsr   'Clear any previous extend operation
    imIsrRecLen = Len(tmIsr)
    'Setup key for first time this vehicle, by contract and lowest line id
    llContrCode = tgChfED.lCode
    slStr = Trim$(Str$(llContrCode))
    Do While Len(slStr) < 10
        slStr = "0" & slStr
    Loop

    tmIsrSrchKey.sKey = slStr   '& "|00001"
    tmIsrSrchKey.iGenDate(0) = igNowDate(0)
    tmIsrSrchKey.iGenDate(1) = igNowDate(1)
    tmIsrSrchKey.iGenTime(0) = igNowTime(0)
    tmIsrSrchKey.iGenTime(1) = igNowTime(1)
    'gather the offsets for extended btrieve tests
    'ilChfCodeOffSet = GetOffSetForInt(tmIsr, tmIsr.lChfCode)
    ilChfCodeOffSet = gFieldOffset("Isr", "IsrChfCode")

    ilChfCodeOffSet = 0             'offset to tmIsr.skey
    'ilGenDateOffSet = GetOffSetForInt(tmIsr, tmIsr.iGenDate(0))
    ilGenDateOffSet = gFieldOffset("Isr", "IsrGenDate")
    'ilGenTimeOffSet = GetOffSetForInt(tmIsr, tmIsr.iGenTime(0))
    ilGenTimeOffSet = gFieldOffset("Isr", "IsrGenTime")
    'access by contract code
    ilRet = btrGetGreaterOrEqual(hmIsr, tmIsr, imIsrRecLen, tmIsrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE And llContrCode = tmIsr.lChfCode Then    'make sure something matches first time thru
        ilExtLen = Len(tmIsr)  'Extract operation record size
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hmIsr, llNoRec, -1, "UC", "Isr", "") '"EG") 'Set extract limits (all records)

        'tlLongTypeBuff.lCode = llContrCode
        'ilRet = btrExtAddLogicConst(hmIsr, BTRV_KT_INT, ilChfCodeOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlLongTypeBuff, 4)

        tlKeyTypeBuff.sKey = slStr
        ilRet = btrExtAddLogicConst(hmIsr, BTRV_KT_STRING, ilChfCodeOffSet, 10, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlKeyTypeBuff, 10)
        tlDateTypeBuff.iDate0 = igNowTime(0)                       'retrieve past  projection records
        tlDateTypeBuff.iDate1 = igNowTime(1)
        ilRet = btrExtAddLogicConst(hmIsr, BTRV_KT_DATE, ilGenTimeOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlDateTypeBuff, 4)

        tlDateTypeBuff.iDate0 = igNowDate(0)                       'retrieve past  projection records
        tlDateTypeBuff.iDate1 = igNowDate(1)
        ilRet = btrExtAddLogicConst(hmIsr, BTRV_KT_DATE, ilGenDateOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)

        ilRet = btrExtAddField(hmIsr, 0, ilExtLen)  'Extract the whole record
        On Error GoTo mBuildSpotsErr
        gBtrvErrorMsg ilRet, "gObtainIsr (btrExtAddField):" & "Isr.Btr", RptSelED
        On Error GoTo 0
        ilRet = btrExtGetNext(hmIsr, tmIsr, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mBuildSpotsErr
            gBtrvErrorMsg ilRet, "gObtainIsr (btrExtGetNextExt):" & "Isr.Btr", RptSelED
            On Error GoTo 0
            ilExtLen = Len(tmIsr)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmIsr, tmIsr, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                tmSdf.lChfCode = llContrCode
                tmSdf.iLineNo = tmIsr.iLineNo
                tmSdf.lCode = tmIsr.lCode
                tmSdf.iVefCode = tmIsr.iVefCode
                tmSdf.sSchStatus = tmIsr.sSchStatus
                tmSdf.sPriceType = tmIsr.sPriceType     'price type, spot type "x" have been excluded
                tmSdf.sSpotType = tmIsr.sType           'Sdf spot type (should not be an "x")
                gPackDateLong tmIsr.lCntrNo, tmSdf.iDate(0), tmSdf.iDate(1)


                slStr = Trim$(Str$(tmSdf.iLineNo))
                Do While Len(slStr) < 5
                    slStr = "0" & slStr
                Loop

                tmSdfList(UBound(tmSdfList)).sKey = slStr      'key for sorting  by line #
                tmSdfList(UBound(tmSdfList)).tSdf = tmSdf
                ReDim Preserve tmSdfList(1 To UBound(tmSdfList) + 1) As SDFSORTLIST
                ilRet = btrExtGetNext(hmIsr, tmIsr, ilExtLen, llRecPos)
                If ilRet <> BTRV_ERR_NONE Then
                    ilRet = ilRet
                End If
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmIsr, tmIsr, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    Exit Function
mBuildSpotsErr:
    gBtrvErrorMsg ilRet, "mGatherISRByCnt : RptCrEd", RptSelED
    mGatherISRbyCnt = ilRet
End Function
'
'
'           mBuildPast - Find all "IN" & "AN" transactions from Phf/Rvf
'           and build an entry for each unique contract.  Within that,
'           build entry with contract package totals and airing vehicle totals
'
'           <input>  llSingleCntr as long
'                    llstdStartDates() as long
'           <output> in module array:
'           tmChf_Sums - this contains the contract # & Code, plus package totals & details
'   2-11-05 Chg array of PHF/RVF array from integer to long (prevent overflow)

Function mGatherRVF(llSingleCntr As Long, llStdStartDates() As Long)
Dim slEarliestDate As String
Dim slLatestDate As String
Dim ilLoop  As Integer
Dim llUpper As Long                 '2-11-05 chg to long
Dim ilRet As Integer
Dim slStr As String
Dim llCntrNo As Long
Dim llLastBilledDate As Long
Dim tlTranType As TRANTYPES
ReDim tlRvf(0 To 0) As RVF
Dim llRvfLoop As Long               '2-11-05 chg to long


    'include only tran types "IN" & "AN" for both cash & trade
    tlTranType.iInv = True
    tlTranType.iAdj = True
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = True
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iNTR = False         '9-17-02

    slEarliestDate = Format$(llStdStartDates(2) - 60, "m/d/yy")
    gUnpackDateLong imLastBilled(0), imLastBilled(1), llLastBilledDate
    'slLatestDate = Format$(llLastBilledDate, "m/d/yy")
    slLatestDate = Format$(llStdStartDates(14) + 60, "m/d/yy")   'go 2 months past year to make sure everything billed for contracts processing
    ilRet = gObtainPhfRvf(RptSelED, slEarliestDate, slLatestDate, tlTranType, tlRvf(), 0)
    'Generate the sort key to sort all transactions by contract #
    ReDim tmRvfSort(1 To 1) As RVFSORT
    For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1
        If (llSingleCntr <> 0 And llSingleCntr = tlRvf(llRvfLoop).lCntrNo) Or (llSingleCntr = 0) Then
            llCntrNo = tlRvf(llRvfLoop).lCntrNo
            slStr = Trim$(Str$(llCntrNo))
            Do While Len(slStr) < 10
                slStr = "0" & slStr
            Loop
            llUpper = UBound(tmRvfSort)
            tmRvfSort(llUpper).sKey = Trim$(slStr)
            tmRvfSort(llUpper).tlRvfRec = tlRvf(llRvfLoop)
            ReDim Preserve tmRvfSort(1 To llUpper + 1)
        End If
    Next llRvfLoop
    llUpper = UBound(tmRvfSort) - 1
    If llUpper > 1 Then   'sort the transactions by contract # to build Summary table faster
        ArraySortTyp fnAV(tmRvfSort(), 1), llUpper, 0, LenB(tmRvfSort(1)), 0, LenB(tmRvfSort(1).sKey), 0
    End If
    Erase tlRvf
End Function
'
'
'
'           mGetAiredDollars - obtain aired dollars for the airing vehicles
'           Build into array tmAirRvf
'
'          <input> llSTdStartDates() = array of start month dates to determine where
'                                       the flight $ belong
'
Sub mGetAiredDollars(llStdStartDates() As Long, ilStartInx As Integer, ilEndInx As Integer)
Dim ilSdfIndex As Integer
Dim ilRet As Integer
Dim ilFound As Integer
Dim ilVefCode As Integer
Dim llAmt As Long
Dim llSDate As Long
Dim ilLoop As Integer
Dim slAiredPrice As String
Dim ilMonthInx As Integer
Dim ilCorT As Integer
Dim slPctTrade As String
Dim slPct As String
Dim slSplitCT As String
    'loop through spots gathered and find the ones that belong to the matching package line # or conventional line
    'For ilSdfIndex = LBound(tgLnSdfExt) To UBound(tgLnSdfExt) - 1 Step 1

    'For ilSdfIndex = LBound(tmSdfList) To UBound(tmSdfList) - 1 Step 1
    For ilSdfIndex = ilStartInx To ilEndInx
        'If (tgLnSdfExt(ilSdfIndex).iLineNo = tmClf.iLine) Then
        If (tmSdfList(ilSdfIndex).tSdf.iLineNo = tmClf.iLine) Then
            tmSdf = tmSdfList(ilSdfIndex).tSdf
            'tmSdfSrchKey3.lCode = tgLnSdfExt(ilSdfIndex).lCode  'retrieve by spot code key index
            'ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            ilFound = True
            If tgSpf.sInvAirOrder <> "S" Then   'S=Update/Update Ordered Vehicle
                ilVefCode = tmSdf.iVefCode
                'If (tgLnSdfExt(ilSdfIndex).sSchStatus = "H") Or (tgLnSdfExt(ilSdfIndex).sSchStatus = "C") Then   'hidden or cancelled spot
                If (tmSdf.sSchStatus = "H") Or (tmSdf.sSchStatus = "C") Then   'hidden or cancelled spot
                    ilFound = False
                End If
            Else
                ilVefCode = tmClf.iVefCode
            End If
            If ilFound Then
                'ilret = mGetRate(tmSdf, tmclf, hmCff, tmCff, slAiredPrice)
                ilRet = gGetFlightPrice(tmSdf, tmClf, hmCff, hmSmf, slAiredPrice)
                If (InStr(slAiredPrice, ".") = 0) Then        'found spot cost, see if its Fill, NC, Extra, etc
                    slAiredPrice = ".00"
                End If
                llAmt = gStrDecToLong(slAiredPrice, 2)
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSDate
                'determine the month this spot belongs in
                For ilLoop = 1 To 14            'it has to go into one since there are buckets to store prior and future spots (to the requested period)
                    If llSDate >= llStdStartDates(ilLoop) And llSDate < (llStdStartDates(ilLoop + 1)) Then
                        ilMonthInx = ilLoop
                        Exit For
                    End If
                Next ilLoop


                'Add spots aired $ to overall contract totals
                'mAddRvfAir ilVefCode, tmClf.iPkLineNo, llAmt, ilMonthInx, tmAirRvf()

                'Add spots aired $ to cash or trade totals
                slPctTrade = gIntToStrDec(tgChfED.iPctTrade, 0)
                For ilCorT = 1 To 2
                    If ilCorT = 1 Then                 'all cash commissionable
                        slPct = gSubStr("100.", slPctTrade)
                        slSplitCT = gDivStr(gMulStr(slAiredPrice, slPct), "100")
                        llAmt = gStrDecToLong(slSplitCT, 2)
                        'Add spots aired $ to cash contract totals
                        mAddRvfAir ilVefCode, tmClf.iPkLineNo, llAmt, ilMonthInx, tmAirRvfCash()
                    Else
                        If ilCorT = 2 Then                'at least cash is commissionable
                            slSplitCT = gDivStr(gMulStr(slAiredPrice, slPctTrade), "100")
                        End If
                        llAmt = gStrDecToLong(slSplitCT, 2)
                        'Add spots aired $ to trade contract totals
                        mAddRvfAir ilVefCode, tmClf.iPkLineNo, llAmt, ilMonthInx, tmAirRvfTrade()
                     End If
                Next ilCorT
            End If
        End If
    Next ilSdfIndex
End Sub
'
'
'
'           mGetOrderedDollars - obtain ordered dollars for a Package or Conventional line.
'           Build into array tmPkRvf
'
'           <input> ilclf - index into the schedule line to be processed from tgClf
'                   llSTdStartDates() = array of start month dates to determine where
'                                       the flight $ belong
'
Sub mGetOrderedDollars(ilClf As Integer, llStdStartDates() As Long)
Dim ilLoop As Integer
ReDim llProject(1 To 14) As Long
ReDim llSplitCT(1 To 14) As Long
Dim slAmount As String
Dim slPctTrade As String
Dim slPct As String
Dim slSplitCT As String
Dim ilCorT As Integer
Dim ilLineNo As Integer
    If tmClf.sType = "O" Then   'if its an  package line, use line # as reference
        ilLineNo = tmClf.iLine
    Else                         'other no package line # applies
        ilLineNo = tmClf.iPkLineNo
    End If
    'gBuildFlights ilClf, llStdStartDates(), ilLastBilledInx + 1, 15, llProject(), 1  'obtain gross $
    gBuildFlights ilClf, llStdStartDates(), 1, 15, llProject(), 1, tgClfED(), tgCffED() 'obtain gross $
    'Alter to net if needed
    'If smGrossOrNet = "N" Then
    '    For ilLoop = 1 To 14        'calculate the net for all periods
    '        slDollar = gLongToStrDec(llProject(ilLoop), 2)
    '        slAmount = gRoundStr(gDivStr(gMulStr(slDollar, gSubStr("100.00", smCashAgyComm)), "100.00"), ".01", 2)
    '        llProject(ilLoop) = gStrDecToLong(slAmount, 2)
    '    Next ilLoop
    'End If
    'Accumulate $ ordered in overall contract totals
    'mAddRvfOrd tmClf.ivefCode, tmClf.iPkLineNo, llProject(), tmPkRvf()
    'Add Ordered $ to cash or trade totals
    slPctTrade = gIntToStrDec(tgChfED.iPctTrade, 0)
    For ilCorT = 1 To 2
        If ilCorT = 1 Then                 'all cash commissionable
            slPct = gSubStr("100.", slPctTrade)
            For ilLoop = 1 To 14
                slAmount = gLongToStrDec(llProject(ilLoop), 2)
                slSplitCT = gDivStr(gMulStr(slAmount, slPct), "100")
                llSplitCT(ilLoop) = gStrDecToLong(slSplitCT, 2)
            Next ilLoop
            'Add spots aired $ to cash contract totals

            mAddRvfOrd tmClf.iVefCode, ilLineNo, llSplitCT(), tmPkRvfCash()
        Else
            If ilCorT = 2 Then                'at least cash is commissionable
                For ilLoop = 1 To 14
                    slAmount = gLongToStrDec(llProject(ilLoop), 2)
                    slSplitCT = gDivStr(gMulStr(slAmount, slPctTrade), "100")
                    llSplitCT(ilLoop) = gStrDecToLong(slSplitCT, 2)
                Next ilLoop
            End If
            'Add spots aired $ to trade contract totals
            mAddRvfOrd tmClf.iVefCode, ilLineNo, llSplitCT(), tmPkRvfTrade()
            End If
    Next ilCorT
    For ilLoop = 1 To 14                        'reinitialize the array that builds ordered $ from sch lines
        llProject(ilLoop) = 0
    Next ilLoop
End Sub


'
'
'           mBobOpenFiles - open files applicable to Producers Earned Distribution Report
'                           (This report is similar to the Billed and Booked, except
'                           it determines the future based on Bill as Aired using
'                           Balance across all schedule lines rather than week.
'
'
'
Function mOpenProdFiles() As Integer
Dim ilRet As Integer
Dim ilError As Integer
    ilError = False
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imGrfRecLen = Len(tmGrf)
    hmChf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imChfRecLen = Len(tmChf)

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imVefRecLen = Len(tmVef)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imCffRecLen = Len(tmCff)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSdfRecLen = Len(tmSdf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSmfRecLen = Len(tmSmf)
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "AGf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imAgfRecLen = Len(tmAgf)
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSlfRecLen = Len(tmSlf)
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSofRecLen = Len(tmSof)
    hmIsr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmIsr, "", sgDBPath & "Isr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imIsrRecLen = Len(tmIsr)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmIsr)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmChf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmSof)
        btrDestroy hmIsr
        btrDestroy hmGrf
        btrDestroy hmChf
        btrDestroy hmVef
        btrDestroy hmClf
        btrDestroy hmCff
        btrDestroy hmSdf
        btrDestroy hmSmf
        btrDestroy hmAgf
        btrDestroy hmSlf
        btrDestroy hmSof
        mOpenProdFiles = ilError
        Exit Function
    End If
End Function
'
'
'                   mWRiteCashOrTrade - Loop thru cash or trade airing & package arrays
'                   to create the current years billing by participant & vehicle
'
'                   <input> ilLastBilledInx - first month considered future
'                           llStdStartDates - array of start dates (2-13 = year requested start months)
'                           slCOrT - C = cash, T = trade
'                           tlAirRvfVefCT - cash or trade array for airing vehicle
'                           tlPkRvfVefCT - cash or trade array for package vehicle
'
'
'               Formula to calculate the Airing vehicles monthly $:
'
'               Total month to be billed $ * (Total veh $ aired, sdf - Total veh $ billed, rvf)
'                                  - - - divided by - - -
'                    Pkg line Total $ (ordered line) - Total Pkg $ already billed (RVF)
'
Sub mWriteCashOrTrade(ilLastBilledInx As Integer, llStdStartDates() As Long, slCOrT As String, tlAirRvfVefCT() As RVF_VEF, tlPkRvfVefCT() As RVF_VEF)
Dim ilAirVehLoop As Integer
Dim ilPkVehLoop As Integer
Dim ilRet As Integer
Dim ilMonthLoop As Integer
Dim ilOwnerLoop As Integer
Dim ilVeh As Integer
Dim flMonthOrdered As Single      'total package ordered for one month
Dim flVefAired As Single          'total veh $ aired (all sdf)
Dim flMonthBilled As Single       'total veh $ billed (all so far)
Dim flPkgOrdered As Single        'total package ordered entire line
Dim flPkgBilled As Single         'total package billed entire line
Dim llAdjustedAmt As Long       'adjusted monthly inv amount
Dim slDollar As String
Dim slAmount As String
Dim slSharePct As String
Dim ilProcessIt As Integer
Dim ilProdLoop As Integer
Dim llGross As Long
ReDim ilProdPct(1 To 8) As Integer  'If sOwnRep = R, then producer's %  (xx.xx)
'Dim ilDebug As Integer
    'ilDebug = False         'for debugging only, prevent from splitting by participant
    'Loop thru all the airing vehicles.  For each airing vehicle find the associated package vehicle and calc the months billing
    'Process one month at a time for all vehicles, then roll over the billed $ into the past
    For ilMonthLoop = ilLastBilledInx + 1 To 14    'loop one month at a time
        For ilAirVehLoop = LBound(tlAirRvfVefCT) To UBound(tlAirRvfVefCT) - 1     'loop thru all the hidden vehicles
            If tlAirRvfVefCT(ilAirVehLoop).iPkLineNo = 0 Then          'conventional line, whats aired is what is billed
                tlAirRvfVefCT(ilAirVehLoop).lTotalGross(ilMonthLoop) = tlAirRvfVefCT(ilAirVehLoop).lTotalOrd(ilMonthLoop)
            Else                                       'package line processing,
                'find the matching package info for the hidden line
                For ilPkVehLoop = LBound(tlPkRvfVefCT) To UBound(tlPkRvfVefCT) - 1
                    If tlPkRvfVefCT(ilPkVehLoop).iPkLineNo = tlAirRvfVefCT(ilAirVehLoop).iPkLineNo Then  'find the matching package line for ordered info

                        flMonthOrdered = tlPkRvfVefCT(ilPkVehLoop).lTotalOrd(ilMonthLoop)    'ordered for month
                        flPkgOrdered = tlPkRvfVefCT(ilPkVehLoop).lTotalVefDollars    'package ordered entire line
                        flPkgBilled = tlPkRvfVefCT(ilPkVehLoop).lTotalVefBilledDollars  'total package billed so far
                        flMonthBilled = tlAirRvfVefCT(ilAirVehLoop).lTotalVefBilledDollars  'total aired veh billed so far
                        flVefAired = tlAirRvfVefCT(ilAirVehLoop).lTotalVefDollars       'total spots aired entire line
                        If (flPkgOrdered - flPkgBilled) <> 0 Then
                            llAdjustedAmt = (flMonthOrdered * (flVefAired - flMonthBilled)) / (flPkgOrdered - flPkgBilled)
                        Else
                            llAdjustedAmt = 0
                    End If
                    'airing vehicle
                    tlAirRvfVefCT(ilAirVehLoop).lTotalGross(ilMonthLoop) = llAdjustedAmt
                    'tlAirRvfVefCT(ilAirVehLoop).lTotalVefBilledDollars = tlAirRvfVefCT(ilPkVehLoop).lTotalVefBilledDollars + llAdjustedAmt

                    End If
                Next ilPkVehLoop
            End If
        Next ilAirVehLoop
        'once all the airing vehicles are invoiced for the month,roll over the invoiced $ into the past
        For ilAirVehLoop = LBound(tlAirRvfVefCT) To UBound(tlAirRvfVefCT) - 1
            If tlAirRvfVefCT(ilAirVehLoop).iPkLineNo = 0 Then          'conventional line, whats aired is what is billed
                tlAirRvfVefCT(ilAirVehLoop).lTotalGross(ilMonthLoop) = tlAirRvfVefCT(ilAirVehLoop).lTotalOrd(ilMonthLoop)  'for the std vehicles, TotalOrd field contains the SDF $
            Else
                'find the matching package info for the hidden line
                For ilPkVehLoop = LBound(tlPkRvfVefCT) To UBound(tlPkRvfVefCT) - 1
                    If tlPkRvfVefCT(ilPkVehLoop).iPkLineNo = tlAirRvfVefCT(ilAirVehLoop).iPkLineNo Then  'find the matching package line for ordered info
                        'roll over the hidden $ for the month into the total pkg billed
                        tlPkRvfVefCT(ilPkVehLoop).lTotalVefBilledDollars = tlPkRvfVefCT(ilPkVehLoop).lTotalVefBilledDollars + tlAirRvfVefCT(ilAirVehLoop).lTotalGross(ilMonthLoop)
                        'roll over the hidden $ for the month into the monthly total pkg billed
                        tlPkRvfVefCT(ilPkVehLoop).lTotalGross(ilMonthLoop) = tlPkRvfVefCT(ilPkVehLoop).lTotalGross(ilMonthLoop) + tlAirRvfVefCT(ilAirVehLoop).lTotalGross(ilMonthLoop)
                        'roll over the hidden $ invoiced for the monthl into the total billed for this veh
                        tlAirRvfVefCT(ilAirVehLoop).lTotalVefBilledDollars = tlAirRvfVefCT(ilAirVehLoop).lTotalVefBilledDollars + tlAirRvfVefCT(ilAirVehLoop).lTotalGross(ilMonthLoop)
                        Exit For
                    End If
                Next ilPkVehLoop
            End If
        Next ilAirVehLoop
    Next ilMonthLoop
    'Each future months billing have been calculated and stored back into array, now create a
    'record for each participant associated  with the vehicle
    tmGrf.lChfCode = tgChfED.lCode
    For ilAirVehLoop = LBound(tlAirRvfVefCT) To UBound(tlAirRvfVefCT) - 1
        llGross = 0
        'accumulate totals to see if it should be written to disk (ignore $0)
        For ilVeh = 2 To 13
            llGross = llGross + tlAirRvfVefCT(ilAirVehLoop).lTotalGross(ilVeh)
        Next ilVeh

        '3-8-02 dont include the past & future values in order to show on report
        'If imYrOrCnt = 2 Then           '1 = show by yr(12 month only(, 2=show by cnt (past & future)
        '    'include the past or future if balancing by contract and the contract falls within requested year
        '    gUnpackDateLong tgChfED.iStartDate(0), tgChfED.iStartDate(1), llSDate
        '    gUnpackDateLong tgChfED.iEndDate(0), tgChfED.iEndDate(1), llEDate
        '    If llSDate <= llStdStartDates(14) - 1 And llEDate >= llStdStartDates(1) Then      'check if contract start/end dates intersects the requested year
        '        llGross = llGross + tlAirRvfVefCT(ilAirVehLoop).lTotalGross(1) + tlAirRvfVefCT(ilAirVehLoop).lTotalGross(14)
        '    End If
        'End If

        'find the matching vehicle so the participants involved can be split
        'For ilVeh = LBound(tgMVef) To UBound(tgMVef) - 1
        '    If tgMVef(ilVeh).iCode = tlAirRvfVefCT(ilAirVehLoop).iVefCode Then
        '        Exit For
        '    End If
        'Next ilVeh
        ilVeh = gBinarySearchVef(tlAirRvfVefCT(ilAirVehLoop).iVefCode)
        If ilVeh <> -1 Then
            'cycle through all the particpants processing only those that match the sales source from the contract (via slsp)
            For ilOwnerLoop = 1 To 8
                ilProdPct(ilOwnerLoop) = tgMVef(ilVeh).iProdPct(ilOwnerLoop) 'If sOwnRep = R, then producer's %  (xx.xx)
            Next ilOwnerLoop

            For ilOwnerLoop = 1 To 8
                If tgMVef(ilVeh).iMnfSSCode(ilOwnerLoop) = imMatchSSCode Then
                    'determine if selective participants
                    ilProcessIt = True
                    If Not imCkcAll Then
                        ilProcessIt = False
                        For ilProdLoop = LBound(imMnfCodes) To UBound(imMnfCodes) - 1
                            If imMnfCodes(ilProdLoop) = tgMVef(ilVeh).iMnfGroup(ilOwnerLoop) Then
                                ilProcessIt = True
                                Exit For
                            End If
                        Next ilProdLoop
                    End If
                    If ilProcessIt And llGross <> 0 Then
                        If imConsolidate Then   '4-4-02 if no splits, use 100%
                            'tgMVef(ilVeh).iProdPct(ilOwnerLoop) = 10000
                            ilProdPct(ilOwnerLoop) = 10000  'dont change contents of the global array of vef
                        End If
                        For ilMonthLoop = 1 To 14
                            slDollar = gLongToStrDec(tlAirRvfVefCT(ilAirVehLoop).lTotalGross(ilMonthLoop), 2)
                            'slSharePct = gIntToStrDec(tgMVef(ilVeh).iProdPct(ilOwnerLoop), 4)
                            slSharePct = gIntToStrDec(ilProdPct(ilOwnerLoop), 4)    '4-5-02
                            slAmount = gMulStr(slSharePct, slDollar)
                            If smGrossOrNet = "N" Then
                                slAmount = gRoundStr(gDivStr(gMulStr(slAmount, gSubStr("100.00", smCashAgyComm)), "100.00"), ".01", 2)
                            End If
                            tmGrf.lDollars(ilMonthLoop) = gStrDecToLong(slAmount, 2)
                        Next ilMonthLoop
                        tmGrf.iVefCode = tlAirRvfVefCT(ilAirVehLoop).iVefCode    'airing vehicle code
                        tmGrf.iCode2 = imMatchSSCode                        'sales source
                        tmGrf.iSofCode = tgMVef(ilVeh).iMnfGroup(ilOwnerLoop)   'participant
                        tmGrf.sDateType = slCOrT                'C = cash, T = trade
                        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                        If imConsolidate Then   '4-4-02 if no splits, 100% considered and now get out
                            ilOwnerLoop = 8         'force exit
                        End If
                    End If
                End If
            Next ilOwnerLoop
        End If
    Next ilAirVehLoop
End Sub
