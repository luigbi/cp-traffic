Attribute VB_Name = "RptCrmIV"
Option Explicit
Dim imAdvt As Integer
'true if advt option
Dim imSlsp As Integer                   'true if slsp option
Dim imVehicle As Integer                'true if vehicle option
Dim imAirVeh As Integer
Dim imBillVeh As Integer
Dim imOwner As Integer                  'true if owner option
Dim imAgency As Integer                 'true if agency option
Dim imInvoice As Integer                'true if invoice option
Dim hmChf As Integer            'Contract header file handle
Dim hmTChf As Integer           'secondary contr header handle, so get next is not destroyed
Dim tmChfSrchKey As LONGKEY0    'CHF record image
Dim tmChfSrchKey1 As CHFKEY1    'CHF record image
Dim imChfRecLen As Integer      'CHF record length
Dim tmChf As CHF
Dim tlChfAdvtExt() As CHFADVTEXT
Dim hmClf As Integer            'Contract line file handle
Dim tmClfSrchKey As CLFKEY0     'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmCtf As Integer            'Contract summary file handle
Dim imCtfRecLen As Integer        'CTF record length
Dim tmCtf As CTF
Dim tlCtfExt() As CTFLIST
Dim hmAdf As Integer            'Advertisr file handle
Dim imAdfRecLen As Integer      'ADF record length
Dim tmAdfSrchKey As INTKEY0     'ADF key image
Dim tmAdf As ADF
                                     
Dim hmAgf As Integer            'Agency file handle
Dim imAgfRecLen As Integer      'AGF record length
Dim tmAgfSrchKey As INTKEY0     'AGF key image
Dim tmAgf As AGF
Dim hmSof As Integer            'Office file handle
Dim imSofRecLen As Integer      'SOF record length
Dim tmSof As SOF
Dim hmSlf As Integer            'Salesperson file handle
Dim imSlfReclen As Integer      'SLF record length
Dim tmSlfSrchKey As INTKEY0     'SLF key image
Dim tmSlf As SLF
Dim hmUrf As Integer            'User file handle
Dim imUrfRecLen As Integer      'URF record length
Dim tmUrf As URF
Dim hmRdf As Integer            'Dayparts file handle
Dim imRdfRecLen As Integer      'RD record length
Dim tmRdfSrchKey As INTKEY0     'RDF key image
Dim tmRdf As RDF
Dim tmAvRdf() As RDF            'array of dayparts
Dim hmDnf As Integer            'Demo file handle
Dim imDnfRecLen As Integer      'DNF record length
Dim tmDnfSrchKey As INTKEY0     'DNF key image
Dim tmDnf As DNF
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnfSrchKey As INTKEY0
Dim tmMnf As MNF
Dim tlMMnf() As MNF                    'array of MNF records for specific type
Dim hmCxf As Integer            'Comment file handle
Dim imCxfRecLen As Integer      'CXF record length
Dim tmCxfSrchKey As LONGKEY0    'CXF key record image
Dim tmCxf As CXF
Dim hmCbf As Integer            'Contract BR file handle
Dim imCbfRecLen As Integer      'CBF record length
Dim tmCbfSrchKey As CBFKEY0     'Gen date and time
Dim tmCbf As CBF
Dim tmZeroCbf As CBF
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer            'Virtual Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim tmVsfSrchKey As LONGKEY0            'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
Dim hmVlf As Integer            'Vehicle Link file handle
Dim tmVlf As VLF                'VLF record image
Dim tmVlfSrchKey0 As VLFKEY0            'VLF by selling vehicle record image
Dim tmVlfSrchKey1 As VLFKEY1            'VLF by airing vehicle record image
Dim imVlfRecLen As Integer        'VLF record length
Dim hmSsf As Integer            'Spot Summary file handle
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim imSsfRecLen As Integer
Dim hmSdf As Integer            'Spot detail file handle
Dim imSdfRecLen As Integer
Dim hmSmf As Integer            'Mg/outside file handle
Dim imSmfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmBBSpot As BBSPOTSS
Dim tmProgTest As PROGRAMSS
Dim tmAvailTest As AVAILSS
Dim tmSpotTest As CSPOTSS
Dim tmBBSpotTest As BBSPOTSS
Dim tmRcf As RCF
Dim hmRcf As Integer            'Rate Card file handle
Dim tmRcfSrchKey As INTKEY0     'RCF record image
Dim imRcfRecLen As Integer      'RCF record length
Dim tmRif As RIF
Dim hmRif As Integer            'Rate Card items file handle
Dim tmRifSrchKey As INTKEY0     'RIF record image
Dim imRifRecLen As Integer      'RIF record length
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim tmGrfSrchKey As GRFKEY0       'Gen date and time
Dim imGrfRecLen As Integer        'GPF record length
Dim tmZeroGrf As GRF              'initialized Generic recd
Dim tmAnr As ANR                    'prepass analysis file
Dim hmAnr As Integer
Dim tmAnrSrchKey As GRFKEY0       'Gen date and time
Dim imAnrRecLen As Integer        'ANR record length
Dim tmZeroAnr As ANR              'initialized Analysis recd
Dim tmBvf As BVF                  'Budgets by office & vehicle
Dim hmBvf As Integer
Dim tmBvfSrchKey As BVFKEY0       'Gen date and time
Dim imBvfRecLen As Integer        'BVF record length
Dim tmBvfVeh() As BVF               'Budget by vehicle
Dim tmPjf As PJF                  'Slsp Projections
Dim hmPjf As Integer
Dim tmPjfSrchKey As PJFKEY0       'Gen date and time
Dim imPjfRecLen As Integer        'PJF record length
'Quarterly Avails
Dim hmAvr As Integer            'Quarterly Avails file handle
Dim tmAvr() As AVR                'AVR record image
Dim tmAvrSrchKey As AVRKEY0            'AVR record image
Dim imAvrRecLen As Integer        'AVR record length
Dim lmSAvailsDates(1 To 13) As Long   'Start Dates of avail week
Dim lmEAvailsDates(1 To 13) As Long   'End dates of avail week
Dim smBucketType As String 'I=Inventory; A=Avail; S=Sold
Dim imMissed As Integer 'True = Include Missed
Dim imStandard As Integer
Dim imRemnant As Integer    'True=Include Remnant
Dim imReserv As Integer  'true = include reservations
Dim imDR As Integer     'True =Include Direct Response
Dim imPI As Integer     'True=Include per Inquiry
Dim imPSA As Integer    'True=Include PSA
Dim imPromo As Integer  'True=Include Promo
Dim imXtra As Integer   'true = include xtra bonus spots
Dim imNC As Integer     'true = include NC spots
Dim imHold As Integer   'true = include hold contracts
Dim imTrade As Integer  'true = include trade contracts
Dim imCash As Integer
Dim imOrder As Integer  'true = include Complete order contracts
'Log Calendar
Dim hmLcf As Integer            'Log Calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim tmLcfSrchKey As LCFKEY0      'LCF record image
Dim imLcfRecLen As Integer        'LCF record length
'Copy Report
Dim hmCpr As Integer            'Copy Report file handle
Dim tmCpr() As CPR                'CPR record image
Dim tmCprSrchKey As CPRKEY0            'CPR record image
Dim imCprRecLen As Integer        'CPR record length
'Copy inventory
Dim hmCif As Integer        'Copy inventory file handle
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer     'CIF record length
' Copy Combo Inventory File
Dim hmCcf As Integer        'Copy Combo Inventory file handle
Dim tmCcf As CCF            'CCF record image
Dim tmCcfSrchKey As INTKEY0 'CCF key record image
Dim imCcfRecLen As Integer     'CCF record length
'  Copy Product/Agency File
Dim hmCpf As Integer        'Copy Product/Agency file handle
Dim tmCpf As CPF            'CPF record image
Dim tmCpfSrchKey As LONGKEY0 'CPF key record image
Dim imCpfRecLen As Integer     'CPF record length
' Time Zone Copy FIle
Dim hmTzf As Integer        'Time Zone Copy file handle
Dim tmTzf As TZF            'TZF record image
Dim tmTzfSrchKey As LONGKEY0 'TZF key record image
Dim imTzfReclen As Integer     'TZF record length
'  Media code File
Dim hmMcf As Integer        'Media file handle
Dim tmMcf As MCF            'MCF record image
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim imMcfRecLen As Integer     'MCF record length
'  Rating Book File
Dim hmDrf As Integer        'Rating book file handle
Dim tmDrf As DRF            'DRF record image
Dim tmDrfSrchKey As DRFKEY0 'DRF key record image
Dim imDrfRecLen As Integer  'DRF record length
'  Receivables File
Dim hmRvf As Integer        'receivables file handle
Dim tmRvf As RVF            'RVF record image
Dim tmRvfSrchKey As INTKEY0 'RVF key record image
Dim imRvfRecLen As Integer  'RVF record length
'  Receivables Report File
Dim hmRvr As Integer        'receivables report file handle
Dim tmRvr As RVR            'RVR record image
Dim tmRvrSrchKey As RVRKEY0   'RVR key record image
Dim imRvrRecLen As Integer  'RVR record length
Dim tlSofList() As SOFLIST    'list of selling office codes and sales sourcecodes
Dim tlActList() As ACTLIST      'Sales Activity list containing advt, potential code, $
Dim tlSlsList() As SLSLIST      'Sales Analysis Summary
'
'                   mBuildFlights - Loop through the flights of the schedule line
'                                   and build the projections dollars into lmprojmonths array
'                   <input> ilclf = sched line index into tgClf
'                           llStdStartDates() - 13 std month start dates
'                           ilFirstProjInx - index of 1st month to start projecting
'                           ilMaxInx - max # of buckets to loop thru
'                   <output> llProject() = array of max 12 months data corresponding to
'                                           12 std start months
'
'                   General routine to build flight $ into week, month, qtr buckets
'
Sub gBuildFlights(ilClf As Integer, llStdStartDates() As Long, ilFirstProjInx As Integer, ilMaxInx As Integer, llProject() As Long)
Dim ilCff As Integer
Dim slStr As String
Dim llFltStart As Long
Dim llFltEnd As Long
Dim ilLoop As Integer
Dim llDate As Long
Dim llDate2 As Long
Dim llSpots As Long
Dim ilTemp As Integer
Dim llStdStart As Long
Dim llStdEnd As Long
Dim ilMonthInx As Integer
    llStdStart = llStdStartDates(ilFirstProjInx)
    llStdEnd = llStdStartDates(ilMaxInx)
    ilCff = tgClf(ilClf).iFirstCff
    Do While ilCff <> -1
    tmCff = tgCff(ilCff).CffRec
    
    gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
    llFltStart = gDateValue(slStr)
    'backup start date to Monday
    ilLoop = gWeekDayLong(llFltStart)
    Do While ilLoop <> 0
        llFltStart = llFltStart - 1
        ilLoop = gWeekDayLong(llFltStart)
    Loop
    gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
    llFltEnd = gDateValue(slStr)
    'the flight dates must be within the start and end of the projection periods,
    'not be a CAncel before start flight, and have a cost > 0
    If (llFltStart < llStdEnd And llFltEnd >= llStdStart) And (llFltEnd >= llFltStart And tmCff.lActPrice > 0) Then
        'only retrieve for projections, anything in the past has already
        'been invoiced and has been retrieved from history or receiv files
        'adjust the gather dates from flights: use flight start date or requested start date, whichever is later
        If llStdStart > llFltStart Then
            llFltStart = llStdStart
        End If
        'use flight end date or requsted end date, whichever is lesser
        If llStdEnd < llFltEnd Then
            llFltEnd = llStdEnd
        End If
        
        For llDate = llFltStart To llFltEnd Step 7
            'Loop on the number of weeks in this flight
            'calc week into of this flight to accum the spot count
            If tmCff.sDyWk = "W" Then            'weekly
                llSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
            Else                                        'daily
                If ilLoop + 6 < llFltEnd Then           'we have a whole week
                    llSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)
                Else
                    llFltEnd = llDate + 6
                    If llDate > llFltEnd Then
                        llFltEnd = llFltEnd       'this flight isn't 7 days
                    End If
                    For llDate2 = llDate To llFltEnd Step 1
                        ilTemp = gWeekDayLong(llDate2)
                        llSpots = llSpots + tmCff.iDay(ilTemp)
                    Next llDate2
                End If
            End If
            'determine month that this week belongs in, then accumulate the gross and net $
            'currently, the projections are based on STandard bdcst
            For ilMonthInx = ilFirstProjInx To ilMaxInx - 1 Step 1       'loop thru months to find the match
                If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                    llProject(ilMonthInx) = llProject(ilMonthInx) + (llSpots * tmCff.lActPrice)
                    Exit For
                End If
            Next ilMonthInx
        Next llDate                                     'for llDate = llFltStart To llFltEnd
    End If                                          '
    ilCff = tgCff(ilCff).iNextCff                   'get next flight record from mem
    Loop                                            'while ilcff <> -1
End Sub
Sub gCrAnrClear()
'*******************************************************
'*                                                     *
'*      Procedure Name:gCrAnrClear    Clear            *
'*                                                     *
'*             Created:07/13/97      By:D. Hosaka      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Clear Pre-pass Analysis file    *
'*                     for Crystal report              *
'*                                                     *
'*******************************************************
    Dim ilRet As Integer
    hmAnr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnr, "", sgDBPath & "Anr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnr)
        btrDestroy hmAnr
        Exit Sub
    End If
    imAnrRecLen = Len(tmAnr)
    tmAnrSrchKey.iGenDate(0) = igNowDate(0)
    tmAnrSrchKey.iGenDate(1) = igNowDate(1)
    tmAnrSrchKey.iGenTime(0) = igNowTime(0)
    tmAnrSrchKey.iGenTime(1) = igNowTime(1)
    ilRet = btrGetGreaterOrEqual(hmAnr, tmAnr, imAnrRecLen, tmAnrSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmAnr.iGenDate(0) = igNowDate(0)) And (tmAnr.iGenDate(1) = igNowDate(1)) And (tmAnr.iGenTime(0) = igNowTime(0)) And (tmAnr.iGenTime(1) = igNowTime(1))
        ilRet = btrDelete(hmAnr)
        ilRet = btrGetNext(hmAnr, tmAnr, imAnrRecLen, BTRV_LOCK_NONE, SETFORWRITE)
    Loop
    ilRet = btrClose(hmAnr)
    btrDestroy hmAnr
End Sub
'
'
'         gCrCumeAct - Create prepass file for Cumulative
'           Activity Report.  Produce a report of all new and
'           modified activity, including contracts whose status
'           is Hold or Order.  Modifications are reflected by
'           increases and decreases from the previous week.  The
'           effetive date entered filters the contracts for the
'           current week (which is always a Monday date).  Increases/
'           decreases are compared against the previous rev #.
'           12 months of contract data is gathered.
'
'           d.hosaka 6 / 8 / 97
'
Sub gCrCumeAct(llStdStartDates() As Long)
Dim ilRet As Integer                    'return flag from all I/O
Dim slAirOrder As String                'billed as aired or ordered (from site pref)
Dim llEarliestEntry As Long             'start date of effective week
Dim llLatestEntry As Long               'end date of effective week
Dim slStr As String                     'temp string variable
Dim ilTemp As Integer                   'temp integer variable
Dim ilTemp2 As Integer                  'temp integer variable
Dim ilSlfCode As Integer                'slsp requesting report, slsp can only see his own stuff
                                        'CSI and guide always sees everything
Dim llEnterDate As Long                 'date entered from contract header
Dim ilFirstTime As Integer              'first time going thru contract header flag
Dim llRecPosition As Long               'current position of contract header for btrieve reads
Dim llGross As Long                     'total Gross $ of contract processing
Dim ilProcessCnt As Integer             'process contract flag (true or false)
Dim ilSelect As Integer                 'index into lbcSelection array for selective adv, agy, demo, vehicle
Dim llPrevCntr As Long
Dim slNameCode As String                'Parsing temporary string
Dim slCode As String                    'Parsing temporary string
Dim llContrCode As Long                 'Current contract's internal code #
Dim ilClf As Integer                    'For loop variable to loop thru sch lines
Dim ilFound As Integer                  'flag if found a selective vehicle when All not checked
ReDim llProject(1 To 12) As Long               '$ for each lines projection
Dim ilUpperVef As Integer               'max vehicles for current contract
Dim ilFoundAgain As Integer             'found veh in memory table to store $
Dim ilModOrNew As Integer               'flag to determine whether contract is mod or new
                                        'so that $ are added or subtractd
                                        '1 = previous data only (decrease), 2 = current data only (new), 3 = both (difference)
Dim ilVefIndex As Integer               'index to vehicle found in memory
Dim ilLoop As Integer                   'temp
Dim llSingleCntr As Long                'single contract # (entred by user)
Dim slGrossOrNet As String              'G = gross, N = net
Dim ilPct As Integer                    '% of cash, % of trade
Dim ilAgyComm As Integer                '100% (gorss) or 85% for net
Dim llTemp As Long                      'temp long variable
hmChf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmChf)
    btrDestroy hmChf
    Exit Sub
End If
imChfRecLen = Len(tmChf)
hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmChf)
    btrDestroy hmGrf
    btrDestroy hmChf
    Exit Sub
End If
imGrfRecLen = Len(tmGrf)
hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmChf)
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmChf
    Exit Sub
End If
imClfRecLen = Len(tmClf)
hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmChf)
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmChf
    Exit Sub
End If
imCffRecLen = Len(tmCff)
hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmChf)
    btrDestroy hmSof
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmChf
    Exit Sub
End If
imSofRecLen = Len(tmSof)
hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmChf)
    btrDestroy hmSlf
    btrDestroy hmSof
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmChf
    Exit Sub
End If
imSlfReclen = Len(tmSlf)
If RptSelIv!rbcSelCInclude(0).value Then                'adv
    ilSelect = 5
ElseIf RptSelIv!rbcSelCInclude(1).value Then            'agy
    ilSelect = 1
ElseIf RptSelIv!rbcSelCInclude(2).value Then            'demo
    ilSelect = 11
Else                                                    'vehicles
    ilSelect = 6
End If
    
slStr = RptSelIv!edcSelCFrom1.Text              'single cntr #
If slStr = "" Then
    llSingleCntr = 0
Else
    llSingleCntr = CLng(slStr)
End If
If RptSelIv!rbcSelC7(0).value Then
    slGrossOrNet = "G"
Else
    slGrossOrNet = "N"
End If
tmGrf = tmZeroGrf                'initialize new record
'Determine contracts to process based on their entered and modified dates
slStr = RptSelIv!edcSelCFrom.Text
'insure its a Monday
llEarliestEntry = gDateValue(slStr)
ilTemp = gWeekDayLong(llEarliestEntry)
Do While ilTemp <> 0
    llEarliestEntry = llEarliestEntry - 1
    ilTemp = gWeekDayLong(llEarliestEntry)
Loop
llEarliestEntry = llEarliestEntry               'effective week (start & end dates)
llLatestEntry = llEarliestEntry + 6
'build array of selling office codes and their sales sources.  This is the most major sort
'in the Business Booked reports
ilTemp = 0
ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
Do While ilRet = BTRV_ERR_NONE
    ReDim Preserve tlSofList(0 To ilTemp) As SOFLIST
    tlSofList(ilTemp).iSofCode = tmSof.iCode
    tlSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
    ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    ilTemp = ilTemp + 1
Loop
If tgUrf(0).iCode = 1 Or tgUrf(0).iCode = 2 Then    'guide or counterpoint password
    ilSlfCode = 0                   'allow guide & CSI to get all stuff
Else
    ilSlfCode = tgUrf(0).iSlfCode   'slsp gets to see only his own stuff
End If
ilFirstTime = True
slAirOrder = tgSpf.sInvAirOrder     'inv all contracts as aired or ordered
ilRet = btrGetFirst(hmChf, tmChf, imChfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)  'get contracts by external contr # (rev #)
Do While ilRet = BTRV_ERR_NONE
    ilRet = btrGetPosition(hmChf, llRecPosition)
    gUnpackDate tmChf.iOHDDate(0), tmChf.iOHDDate(1), slStr
    llEnterDate = gDateValue(slStr)
    If ilFirstTime Then
        ilFirstTime = False
        llPrevCntr = tmChf.lCntrNo
        tmGrf = tmZeroGrf                'initialize new record
        ReDim Preserve tmVefDollars(0 To 0) As ADJUSTLIST            'prepare list of mgs
        ilUpperVef = 1
        ilModOrNew = 0
        'date and time genned need only be set the first time - remains the same
        tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
        tmGrf.iGenDate(1) = igNowDate(1)
        tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
        tmGrf.iGenTime(1) = igNowTime(1)
    End If
    If llPrevCntr <> tmChf.lCntrNo Then
        If ilModOrNew > 1 Then       'dont write anything to disk if  only previous week data
            'build GRF records from vehicle tables built in memory
            tmGrf.lChfCode = tgChf.lCode        'internal contract code
            For ilTemp = LBound(tlSofList) To UBound(tlSofList)
                If tlSofList(ilTemp).iSofCode = tgChf.iSlfCode(0) Then
                    tmGrf.iSofCode = tlSofList(ilTemp).iMnfSSCode          'Sales source
                    Exit For
                End If
            Next ilTemp
            'Create the GRF record for each vehicle in the order
            mCumeInsert ilUpperVef, tmVefDollars(), tmGrf, slGrossOrNet
        End If
        llPrevCntr = tgChf.lCntrNo
        ReDim Preserve tmVefDollars(0 To 0) As ADJUSTLIST            'prepare list of mgs
        ilUpperVef = 1
        ilModOrNew = 0
    End If
    ilProcessCnt = False
    If ilSelect <> 6 And Not RptSelIv!ckcAll Then                     'for advt, agy or demo selectivity, filter out before going to lines
        If ilSelect = 5 Then                              'advt option
            For ilTemp = 0 To RptSelIv!lbcSelection(5).ListCount - 1 Step 1
                If RptSelIv!lbcSelection(5).Selected(ilTemp) Then              'selected slsp
                    slNameCode = tgAdvertiser(ilTemp).sKey 'Traffic!lbcAdvertiser.List(ilTemp)         'pick up slsp code
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) = tmChf.iAdfCode Then
                        ilProcessCnt = True
                        Exit For
                    End If
                End If
            Next ilTemp
        ElseIf ilSelect = 11 Then      'demo  option
            For ilTemp = 0 To RptSelIv!lbcSelection(11).ListCount - 1 Step 1
                If RptSelIv!lbcSelection(11).Selected(ilTemp) Then              'selected slsp
                    slNameCode = tgRptSelDemoCode(ilTemp).sKey    'RptSelIv!lbcCSVNameCode.List(ilTemp)         'pick up slsp code
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) = tmChf.iMnfDemo(0) Then
                        ilProcessCnt = True
                        Exit For
                    End If
                End If
            Next ilTemp
        Else                                'agy
            For ilTemp = 0 To RptSelIv!lbcSelection(1).ListCount - 1 Step 1
                If RptSelIv!lbcSelection(1).Selected(ilTemp) Then
                    slNameCode = tgAgency(ilTemp).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) = tmChf.iagfCode Then
                        ilProcessCnt = True
                        Exit For
                    End If
                End If
            Next ilTemp
        End If
    Else
        ilProcessCnt = True
    End If
    If llEnterDate >= llEarliestEntry And llEnterDate <= llLatestEntry Then   'current weeks contract
        If ilModOrNew >= 2 Then   'dont process if cnt in current week already processed
            ilProcessCnt = False        'multiple rev # within current week, bypass all except most recent
        End If
    ElseIf llEnterDate < llEarliestEntry Then             'past week, has it already been processed?
        If ilModOrNew = 1 Or ilModOrNew = 3 Then        '1 = previous exists, 3 = both exists
            ilProcessCnt = False
        End If
    Else
        ilProcessCnt = False                'date in future
    End If
    If (llEnterDate <= llLatestEntry) And (tmChf.iSlfCode(0) = ilSlfCode Or ilSlfCode = 0) And (tmChf.sStatus = "H" Or tmChf.sStatus = "O" Or tmChf.sStatus = "G" Or tmChf.sStatus = "N") Then
        'if llenterdate is less than llearliest entry, then it needs to be processed.  It could be
        'a contract that did not get carried over (in which case its a decrease)
            
        If (ilProcessCnt) And (llSingleCntr = 0 Or llSingleCntr = tmChf.lCntrNo) Then
            llPrevCntr = tmChf.lCntrNo
            llContrCode = tmChf.lCode
            'get entire contract with schedule line & flights
            ilRet = gObtainCntr(hmChf, hmClf, hmCff, llContrCode, False, tgChf, tgClf(), tgCff())
            If llEnterDate < llEarliestEntry Then      'previous entry
                If ilModOrNew = 0 Then                  'previous weeks data only, subtract$
                    ilModOrNew = 1
                Else                                    'previous & current exist, get difference
                    ilModOrNew = 3
                End If
            Else                                        'current weeks data
                If ilModOrNew = 0 Then                  'did previous exist?
                    ilModOrNew = 2                      'new only
                Else
                    ilModOrNew = 3                      'both exist, get diff
                End If
            End If
            For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                tmClf = tgClf(ilClf).ClfRec
                ilFound = True
                If ilSelect = 6 And Not RptSelIv!ckcAll Then
                    ilFound = False
                    For ilTemp = 0 To RptSelIv!lbcSelection(6).ListCount - 1 Step 1
                        If RptSelIv!lbcSelection(6).Selected(ilTemp) Then              'selected slsp
                            slNameCode = tgCSVNameCode(ilTemp).sKey    'RptSelIv!lbcCSVNameCode.List(ilTemp)         'pick up slsp code
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If Val(slCode) = tmClf.iVefCode Then
                                ilFound = True
                                Exit For
                            End If
                        End If
                    Next ilTemp
                End If
                If ilFound Then                 'all vehicles of selective one found
                    For ilTemp = 1 To 12 Step 1 'init projection $ each time
                        llProject(ilTemp) = 0
                    Next ilTemp
                    If slAirOrder = "O" Then                'invoice all contracts as ordered
                        If tmClf.sType <> "H" Then          'ignore all hidden lines for ordered billing
                            gBuildFlights ilClf, llStdStartDates(), 1, 12, llProject()
                        End If
                    Else                                    'inv all contracts as aired
                        If tmClf.sType = "H" Then             'but if from pkg and hidden line, ignore hidd
                            'if hidden, will project if assoc. package is set to invoice as aired (real)
                            For ilTemp = LBound(tgClf) To UBound(tgClf) - 1    'find the assoc. pkg line for these hidden
                                If tmClf.iPkLineNo = tgClf(ilTemp).ClfRec.iLine Then
                                    If tgClf(ilTemp).ClfRec.sType = "A" Then        'does the pkg line reflect bill as aired?
                                        gBuildFlights ilClf, llStdStartDates(), 1, 12, llProject()  'pkg bills as aired, project the hidden line
                                    End If
                                    Exit For
                                End If
                            Next ilTemp
                        Else                            'conventional, VV, or Pkg line
                            If tmClf.sType <> "A" Then  'if this package line to be invoiced aired (real times),
                                                        'it has already been projected above with the hidden line
                                gBuildFlights ilClf, llStdStartDates(), 1, 12, llProject()
                            End If
                        End If
                    End If
                    'llproject(1-12) contains $ for this line
                    'vehicles are build into memory with its 12 $ buckets
                    ilFoundAgain = False
                    For ilTemp = 0 To ilUpperVef - 1 Step 1
                        If tmVefDollars(ilTemp).iVefCode = tmClf.iVefCode Then
                            ilFoundAgain = True
                            ilVefIndex = ilTemp
                        End If
                    Next ilTemp
                    If Not (ilFoundAgain) Then
                        ReDim Preserve tmVefDollars(0 To ilUpperVef) As ADJUSTLIST
                        tmVefDollars(ilUpperVef).iVefCode = tmClf.iVefCode
                        ilVefIndex = ilUpperVef
                        ilUpperVef = ilUpperVef + 1             'next new vef to store
                    End If
                    'now add or subtract depending if this contract is a new or mod
                    If llEnterDate < llEarliestEntry Then      'previous entry
                        For ilTemp = 1 To 12 Step 1
                            tmVefDollars(ilVefIndex).lProject(ilTemp) = tmVefDollars(ilVefIndex).lProject(ilTemp) - llProject(ilTemp)
                        Next ilTemp
                    Else                                        'current weeks data
                        For ilTemp = 1 To 12 Step 1
                            tmVefDollars(ilVefIndex).lProject(ilTemp) = tmVefDollars(ilVefIndex).lProject(ilTemp) + llProject(ilTemp)
                        Next ilTemp
                    End If
                End If                                  'ilfound
            Next ilClf                                  'for ilclf = lbound(tgclf) - ubound(tgclf)
            'all schedule lines complete, continue to see if there's another same contract # to process before writing to disk
        End If              'ilprocesscnt
    End If                  'llenter date <= lllatesentry
    'reposition back to correct contract header, then read the nextone
    ilRet = btrGetDirect(hmChf, tmChf, imChfRecLen, llRecPosition, INDEXKEY1, BTRV_LOCK_NONE)
    ilRet = btrGetNext(hmChf, tmChf, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
Loop                                    'do while BTRV_ERR_NONE
If ilModOrNew > 1 Then           '<>0 indicates something already processed for this contract
    If tmGrf.lDollars(1) <> 0 Then       'dont write anything to disk if amount is 0
        tmGrf.lChfCode = tgChf.lCntrNo        'internal contract code
        For ilTemp = LBound(tlSofList) To UBound(tlSofList)
            If tlSofList(ilTemp).iSofCode = tgChf.iSlfCode(0) Then
                tmGrf.iSofCode = tlSofList(ilTemp).iMnfSSCode          'Sales source
                Exit For
            End If
        Next ilTemp
        'Create the GRF record for each vehicle in the order
        mCumeInsert ilUpperVef, tmVefDollars(), tmGrf, slGrossOrNet
    End If
End If
Erase tlSofList, llProject, tmVefDollars, llStdStartDates
ilRet = btrClose(hmSlf)
ilRet = btrClose(hmSof)
ilRet = btrClose(hmCff)
ilRet = btrClose(hmClf)
ilRet = btrClose(hmGrf)
ilRet = btrClose(hmChf)
End Sub
'
'
'           gCrMakePlan - Prepass toCalculate Price Needed to
'                         Make Plan by Daypart for each vehicle
'                         Calculate budgets by daypart, gather
'                         inventory, spots sold by daypart .
'                         Vehicles will be compared against the
'                         selected budget (plan or forecast) for
'                         the same year's active rate card (which
'                         is also selected).  If the rate card year
'                         doesn't exist, no data is produced.  That
'                         is, if an old rate is on file, that is
'                         the effective one for contract input, but
'                         for this purpose, that year's rate card
'                         must exist on file.
'
'           Created: D Hosaka   7/11/97
Sub gCrMakePlan()
    Dim ilRet As Integer
    Dim slNameCode As String            'parsing temp string
    Dim slNameYear As String            'parsing temp string
    Dim slYear As String                'year to process budgets
    Dim slCode As String                'paring temp string
    Dim slBdMnfName As String           'budget name
    Dim ilBdYear As Integer             'budget year
    Dim imBSelectedIndex As Integer
    Dim ilBdMnfCode As Integer          'budget code
    Dim ilRCCode As Integer             'rc code
    Dim ilRif As Integer
    Dim llYearStart As Long             'years std start date
    Dim llYearEnd As Long               'years std end date
    Dim llStdInputStart As Long          'std start date requested
    Dim llStdInputEnd As Long           'std end date requested (not past end of year)
    ReDim ilStdInputStart(0 To 1) As Integer  'btrieve form of std start date requested (to put into ANR)
    Dim ilLoop As Integer               'temp
    Dim ilTemp As Integer               'temp
    Dim slStr As String
    Dim slStart As String               'start date of input
    Dim slEnd As String                 'end of std month of first month requested
    Dim llDate As Long                  'temp date
    ReDim ilStartWeeks(1 To 14) As Integer  'start week index for each period
                                        'i.e. for Weekly report the elements will be 1, 2, 3, etc.
                                        'for monthly, the elements will be the start week of the qtr, - 1,5,10,14,18,23...
    Dim ilLoopWks As Integer
    Dim ilVeh As Integer                '
    Dim llAvail As Long
    Dim ilProcessWk As Integer
    Dim ilStartOfPer As Integer
    Dim ilEndOfPer As Integer
    hmChf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmChf)
        btrDestroy hmChf
        Exit Sub
    End If
    imChfRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmBvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmBvf)
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imBvfRecLen = Len(tmBvf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imSdfRecLen = Len(hmSdf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imSmfRecLen = Len(hmSmf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imVefRecLen = Len(hmVef)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imVsfRecLen = Len(hmVsf)
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imSsfRecLen = Len(tmSsf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imMnfRecLen = Len(hmMnf)
    hmRcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRcf, "", sgDBPath & "Rcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRcf)
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imRcfRecLen = Len(hmRcf)
    hmRif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRif, "", sgDBPath & "Rif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRif)
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imRifRecLen = Len(hmRif)
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRdf)
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imRdfRecLen = Len(tmRdf)
    hmAnr = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnr, "", sgDBPath & "Anr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnr)
        btrDestroy hmAnr
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmSsf
        btrDestroy hmVsf
        btrDestroy hmVef
        btrDestroy hmSmf
        btrDestroy hmSdf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmChf
        Exit Sub
    End If
    imAnrRecLen = Len(tmAnr)
    slNameCode = tgRptSelBudgetCode(igBSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 1, "\", slNameYear)
    ilRet = gParseItem(slNameYear, 1, "\", slYear)
    slYear = gSubStr("9999", slYear)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilBdMnfCode = Val(slCode)
    ilBdYear = Val(slYear)
    slNameCode = tgRateCardCode(igRCSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 3, "\", slCode)
    ilRCCode = Val(slCode)
    ReDim tmMRif(1 To 1) As RIF
    'Build array (tmMRif) of all valid Rates for each Vehicle's daypart
    For ilRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
        If (ilRCCode = tgMRif(ilRif).iRcfCode) And ilBdYear = tgMRif(ilRif).iYear Then
            'test for selective vehicle
            For ilLoop = 0 To RptSelIv!lbcSelection(3).ListCount - 1 Step 1
                If (RptSelIv!lbcSelection(3).Selected(ilLoop)) Then
                    slNameCode = tgVehicle(ilLoop).sKey 'Traffic!lbcVehicle.List(ilVehicle)
                    ilRet = gParseItem(slNameCode, 1, "\", slStr)
                    ilRet = gParseItem(slStr, 3, "|", slStr)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) = tgMRif(ilRif).iVefCode Then
                        tmMRif(UBound(tmMRif)) = tgMRif(ilRif)
                        ReDim Preserve tmMRif(1 To UBound(tmMRif) + 1) As RIF
                        ilLoop = RptSelIv!lbcSelection(3).ListCount 'stop the loop
                    End If
                End If
            Next ilLoop
        End If
    Next ilRif
    slNameYear = "1/15/" & Trim$(RptSelIv!edcSelCFrom)      '12/15/year entered
    slNameYear = gObtainStartStd(slNameYear)              'get the stnd years start date
    llYearStart = gDateValue(slNameYear)
    slNameYear = "12/15/" & Trim$(RptSelIv!edcSelCFrom)      '12/15/year entered
    slNameYear = gObtainEndStd(slNameYear)              'get the stnd years end date
    llYearEnd = gDateValue(slNameYear)
    slStr = RptSelIv!edcSelCTo.Text         '#of qtrs to gather
    ilTemp = Val(slStr) * 3                 '# of months to gather
    
    'no budgets can cross over years
    ilLoop = (igMonthOrQtr - 1) * 3 + 1
    slStart = Trim$(Str$(ilLoop)) & "/15/" & Trim$(Str$(ilBdYear))      'format xx/xx/xxxx
    slStart = gObtainStartStd(slStart)
    llStdInputStart = gDateValue(slStart)                 'start of std qtr requested
    slStr = slStart
    slEnd = gObtainEndStd(slStart)
    For ilLoop = 1 To 14
        ilStartWeeks(ilLoop) = 0
    Next ilLoop
    ilStartWeeks(1) = (gDateValue(slStr) - llYearStart) / 7 + 1     'first week to start gathering data
     
    'Obtain the latest date to gather data.  Also, setup week index array of designating the start week of each period (ignore for weekly option)
    For ilLoop = 1 To ilTemp                'loop for # months for all qtrs requested
        'ilstartweeks is array indicating first week of the period (for weekly request, each element will be incremented by one.
        'if quarter, each element is the start of a std brdcst month
        ilStartWeeks(ilLoop) = (gDateValue(slStr) - llYearStart) / 7 + 1     'first week to start gathering data
        slStr = gObtainStartStd(slStr)
        slEnd = gObtainEndStd(slStr)
        llDate = gDateValue(slEnd) + 1      'get to next month
        If llDate > llYearEnd Then
            slStr = Format$(llDate, "m/d/yy")
            ilLoop = ilLoop + 1
            Exit For
        End If
        slStr = Format$(llDate, "m/d/yy")
    Next ilLoop
    llStdInputStart = gDateValue(slStart)
    'convert to btrieve for Crystal
    gPackDate slStart, ilStdInputStart(0), ilStdInputStart(1)
    llStdInputEnd = gDateValue(slEnd)
    ilStartWeeks(ilLoop) = (gDateValue(slStr) - llYearStart) / 7 + 1     'first week to start gathering data
    If RptSelIv!rbcSelC4(0).value Then          'weekly option (vs quarters)
        ilRif = 1
        ilTemp = (llStdInputStart - llYearStart) / 7 + 1
        For ilLoop = ilTemp To 13 + ilTemp
            ilStartWeeks(ilRif) = ilLoop     'Only accum 1 week at a time for weekly option
            ilRif = ilRif + 1
        Next ilLoop
    End If
    mBdGetBudgetDollars hmChf, hmClf, hmCff, hmSdf, hmSmf, hmVef, hmVsf, hmSsf, hmBvf, ilBdMnfCode, ilBdYear, llStdInputStart, llStdInputEnd, tmMRif(), tgMRdf(), 0
    'Create ANR pre-pass from weekly vehicle dayparts (tgDollarRec)
    For ilVeh = LBound(tgImpactRec) To UBound(tgImpactRec) - 1 Step 1  'create a record for each daypart
        tmAnr = tmZeroAnr
        tmAnr.iGenDate(0) = igNowDate(0)
        tmAnr.iGenDate(1) = igNowDate(1)
        tmAnr.iGenTime(0) = igNowTime(0)
        tmAnr.iGenTime(1) = igNowTime(1)
        tmAnr.iVefCode = tgImpactRec(ilVeh).iVefCode           'vehicle
        tmAnr.iRdfCode = tgImpactRec(ilVeh).iRdfCode           'daypart
        tmAnr.imnfBudget = ilBdMnfCode                  'budget code
        tmAnr.iEffectiveDate(0) = ilStdInputStart(0)            'start date of requested period (used for weekly hdr dates)
        tmAnr.iEffectiveDate(1) = ilStdInputStart(1)
        For ilTemp = 1 To 13 Step 1                     'Loop thru array of weeks to accumulate - each entry is a start week for the period
            llAvail = 0
            ilProcessWk = False
            'Accum the # of weeks for each period
            ilLoop = (llStdInputStart - llYearStart) / 7 + 1
            ilStartOfPer = ilStartWeeks(ilTemp) - ilLoop + 1
            ilEndOfPer = ilStartWeeks(ilTemp + 1) - ilLoop
            'For ilLoopWks = ilStartWeeks(ilTemp) To ilStartWeeks(ilTemp + 1) - 1 Step 1
            For ilLoopWks = ilStartOfPer To ilEndOfPer Step 1
                ilProcessWk = True                      'found a week to accumulate
                tmAnr.lBudget(ilTemp) = tmAnr.lBudget(ilTemp) + tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).lBudget         'total budget for period
                tmAnr.lSold(ilTemp) = tmAnr.lSold(ilTemp) + tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).lDollarSold         'total $ sold this period
                tmAnr.lInv(ilTemp) = tmAnr.lInv(ilTemp) + (tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).l30Inv - tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).l30Sold)    'totl availabilty this period
                'llAvail = llAvail + tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).l30Inv - tgDollarRec(ilLoopWks, tgImpactRec(ilVeh).iPtDollarRec).l30Sold
            Next ilLoopWks
            If ilProcessWk Then                 'found a week to process, see if there's any available
                If tmAnr.lInv(ilTemp) < 1 Then             'already sold out
                    tmAnr.iPctSellout(ilTemp) = 1   'flag to denote sold out
                Else                            'calc price needed
                    'tmAnr.lPriceNeeded(ilTemp) = (tmAnr.lBudget(ilTemp) - tmAnr.lSold(ilTemp)) / llAvail
                    tmAnr.lPriceNeeded(ilTemp) = (tmAnr.lBudget(ilTemp) - tmAnr.lSold(ilTemp)) / tmAnr.lInv(ilTemp)
                End If
            End If
        Next ilTemp
        ilRet = btrInsert(hmAnr, tmAnr, imAnrRecLen, INDEXKEY0)
    Next ilVeh
    Erase ilStartWeeks, tgDollarRec, tgImpactRec
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmRif)
    ilRet = btrClose(hmRcf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmBvf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmChf)
    ilRet = btrClose(hmAnr)
End Sub
'**********************************************************************
'
'
'       Pre-pass to Sales Activity Report. produce a report
'       of all new and modified activity, including contracts
'       whose status is Hold or Order, plus Salesperson
'       projection data.  Modifications are reflected by increases
'       and decreases from the previous week.  Any potential
'       business entered as projections are detailed and totaled
'       by their projection category.  The effective dte entered
'       filters the contracts and projections for the current week,
'       using the rollover date.  The date is always backed up
'       to a Monday date, which denotes the current week.
'       Increases/decreases are compared agains the previous rev#.
'       The starting qtr and year entred gathers dollars affectving
'       the requested quarter.
'
'       Created:  D.hosaka  12/96
'
'***************************************************************************
'
'
'
Sub gCrSalesAct()
Dim ilRet As Integer
Dim llEarliestEntry As Long           'start date of all contracts modified or entered new
Dim llLatestEntry As Long             'end date of all contracts modified or entered new
Dim slStr As String                     'temp
Dim llEnterDate As Long                 'contract header entred date
Dim llDate As Long                      'temp
Dim ilFound As Integer
Dim llContrCode As Long                 'contract code to retrieve contr with lines
Dim ilClf As Integer
Dim slAirOrder As String * 1             'O = bill as ordered, A = bill as aired
Dim slStartQtr As String                'active dates of contrct
Dim slEndQtr As String
ReDim llProject(1 To 2) As Long
ReDim llStdStartDates(1 To 2) As Long       'only doing 1 qtr, need start date of 2nd (last) qtr
Dim ilTemp As Integer
Dim ilWhichGrf As Integer
Dim ilFirstTime As Integer
Dim llPrevCntr As Long
Dim ilProcessCnt As Integer
Dim ilUpperAct As Integer               'total # of potential advt
Dim ilCalType As Integer                '0 = std, 1 = cal. month, 4 = corp
Dim llAmount As Long
Dim slTimeStamp As String
Dim ilWkNo As Integer
Dim ilYear As Integer
Dim llRecPosition As Long               'location of contract header for get next reads
Dim ilSlfCode As Integer                'slsp processing this report
hmChf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmChf)
    btrDestroy hmChf
    Exit Sub
End If
imChfRecLen = Len(tmChf)
hmPjf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmPjf, "", sgDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmChf)
    btrDestroy hmPjf
    btrDestroy hmChf
    Exit Sub
End If
imPjfRecLen = Len(tmPjf)
hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmChf)
    btrDestroy hmGrf
    btrDestroy hmPjf
    btrDestroy hmChf
    Exit Sub
End If
imGrfRecLen = Len(tmGrf)
hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmChf)
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmPjf
    btrDestroy hmChf
    Exit Sub
End If
imClfRecLen = Len(tmClf)
hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmChf)
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmPjf
    btrDestroy hmChf
    Exit Sub
End If
imCffRecLen = Len(tmCff)
hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmChf)
    btrDestroy hmSof
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmPjf
    btrDestroy hmChf
    Exit Sub
End If
imSofRecLen = Len(tmSof)
hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmChf)
    btrDestroy hmSlf
    btrDestroy hmSof
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmPjf
    btrDestroy hmChf
    Exit Sub
End If
imSlfReclen = Len(tmSlf)
'Determine contracts to process based on their entered and modified dates
slStr = RptSelIv!edcSelCFrom.Text
'insure its a Monday
llEarliestEntry = gDateValue(slStr)
ilTemp = gWeekDayLong(llEarliestEntry)
Do While ilTemp <> 0
    llEarliestEntry = llEarliestEntry - 1
    ilTemp = gWeekDayLong(llEarliestEntry)
Loop
llEarliestEntry = llEarliestEntry
llLatestEntry = llEarliestEntry + 6
'Determine start and end dates of $ to gather
llStdStartDates(1) = lgOrigCntrNo                  'start date of qtr
llStdStartDates(2) = llStdStartDates(1) + 90       'get start of next qtr
ilYear = Val(RptSelIv!edcSelCTo.Text)
slStartQtr = Format(llStdStartDates(1), "m/d/yy")
slEndQtr = Format(llStdStartDates(2) - 1, "m/d/yy")
'build array of selling office codes and their sales sources.  This is the most major sort
'in the Business Booked reports
ilTemp = 0
ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
Do While ilRet = BTRV_ERR_NONE
    ReDim Preserve tlSofList(0 To ilTemp) As SOFLIST
    tlSofList(ilTemp).iSofCode = tmSof.iCode
    tlSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
    ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    ilTemp = ilTemp + 1
Loop
'Populate the salespeople.  Only salesp running report can see his own stuff
ilRet = gObtainSalesperson()        'populated slsp list in tgMSlf
If tgUrf(0).iCode = 1 Or tgUrf(0).iCode = 2 Then    'guide or counterpoint password
    ilSlfCode = 0                   'allow guide & CSI to get all stuff
Else
    ilSlfCode = tgUrf(0).iSlfCode   'slsp gets to see only his own stuff
End If
ilFirstTime = True
slAirOrder = tgSpf.sInvAirOrder     'inv all contracts as aired or ordered
tmGrf = tmZeroGrf                'initialize new record
ilRet = btrGetFirst(hmChf, tmChf, imChfRecLen, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)  'get contracts by external contr # (rev #)
'for each Contract process 2 entries in table of GRF records.
'If only 1 record created in previous week, its a decrease.
'if only 1 record created in current week, its an increase (New).
'If both records created for previous and current week, show difference.
'Write out 1 record into GRF file containing the final result of previous to current.
Do While ilRet = BTRV_ERR_NONE
    ilRet = btrGetPosition(hmChf, llRecPosition)
    gUnpackDate tmChf.iOHDDate(0), tmChf.iOHDDate(1), slStr
    llEnterDate = gDateValue(slStr)
    If ilFirstTime Then
        ilFirstTime = False
        llPrevCntr = tmChf.lCntrNo
    End If
    If llPrevCntr <> tmChf.lCntrNo Then
        If tmGrf.iPerGenl(3) <> 0 Then           '<>0 indicates something already processed for this contract
            tmGrf.iPerGenl(1) = 1                'assume modification
            If tmGrf.iPerGenl(3) = 2 Then        'flag 2 denotes something current already found
                tmGrf.iPerGenl(1) = 0            'send to crystal flag for new only
            End If
        End If
        If tmGrf.lDollars(1) <> 0 And tmGrf.iPerGenl(3) <> 1 Then       'dont write anything to disk if amount is 0 and only previous week data
            tmGrf.lDollars(1) = tmGrf.lDollars(1) / 100
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
        tmGrf = tmZeroGrf                'initialize new record
    End If
    ilProcessCnt = True
    'If llEnterDate >= llEarliestEntry Then   'current weeks contract
    If llEnterDate >= llEarliestEntry And llEnterDate <= llLatestEntry Then   'current weeks contract
        If tmGrf.iPerGenl(3) >= 2 Then   'dont process if cnt in current week already processed
            ilProcessCnt = False        'multiple rev # within current week, bypass all except most recent
        End If
    ElseIf llEnterDate < llEarliestEntry Then             'past week, has it already been processed?
        If tmGrf.iPerGenl(3) = 1 Or tmGrf.iPerGenl(3) = 3 Then  '1 = previous exists, 3 = both exists
            ilProcessCnt = False
        End If
    Else
        ilProcessCnt = False                'date in future
    End If
    If (llEnterDate <= llLatestEntry) And (tmChf.iSlfCode(0) = ilSlfCode Or ilSlfCode = 0) And (tmChf.sStatus = "H" Or tmChf.sStatus = "O" Or tmChf.sStatus = "G" Or tmChf.sStatus = "N") Then
        'if llenterdate is less than llearliest entry, then it needs to be processed.  It could be
        'a contract that did not get carried over (in which case its a decrease)
            
        If (ilProcessCnt) Then
            llContrCode = tmChf.lCode
            ilRet = gObtainCntr(hmChf, hmClf, hmCff, llContrCode, False, tgChf, tgClf(), tgCff())
            For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                tmClf = tgClf(ilClf).ClfRec
                If slAirOrder = "O" Then                'invoice all contracts as ordered
                    If tmClf.sType <> "H" Then          'ignore all hidden lines for ordered billing, should be Pkg or conventional lines
                        gBuildFlights ilClf, llStdStartDates(), 1, 2, llProject()
                    End If
                Else                                    'inv all contracts as aired
                    If tmClf.sType = "H" Then             'but if from pkg and hidden line, ignore hidd
                        'if hidden, will project if assoc. package is set to invoice as aired (real)
                        For ilTemp = LBound(tgClf) To UBound(tgClf) - 1    'find the assoc. pkg line for these hidden
                            If tmClf.iPkLineNo = tgClf(ilTemp).ClfRec.iLine Then
                                If tgClf(ilTemp).ClfRec.sType = "A" Then        'does the pkg line reflect bill as aired?
                                    gBuildFlights ilClf, llStdStartDates(), 1, 2, llProject()  'pkg bills as aired, project the hidden line
                                End If
                                Exit For
                            End If
                        Next ilTemp
                    Else                            'conventional, VV, or Pkg line
                        If tmClf.sType <> "A" Then  'if this package line to be invoiced aired (real times),
                                                    'it has already been projected above with the hidden line
                            gBuildFlights ilClf, llStdStartDates(), 1, 2, llProject()
                        End If
                    End If
                End If
            Next ilClf
            'all schedule lines complete, write out record
            tmGrf.sBktType = "O"                  'orders flag (vs P = project), for sorting
            tmGrf.iPerGenl(2) = tgChf.iCntRevNo
            tmGrf.sDateType = "O"               'indicate this is Orders vs A,B,C projections for sorting
            If llEnterDate < llEarliestEntry Then               'previous entry
                If tmGrf.iPerGenl(3) = 0 Then
                    tmGrf.iPerGenl(3) = 1                   'previous only
                    tmGrf.lDollars(1) = tmGrf.lDollars(1) - llProject(1)
                Else                                        'previous already exists, this has both prev & current for difference
                    tmGrf.iPerGenl(3) = 3
                    tmGrf.lDollars(1) = tmGrf.lDollars(1) - llProject(1)
                End If
            Else            'current week
                If tmGrf.iPerGenl(3) = 0 Then
                    tmGrf.iPerGenl(3) = 2              'current entry only (new)
                    tmGrf.lDollars(1) = tmGrf.lDollars(1) + llProject(1)
                Else
                    tmGrf.iPerGenl(3) = 3               'previous & current exist
                    tmGrf.lDollars(1) = tmGrf.lDollars(1) + llProject(1)
                End If
            End If
            mWriteSlsAct            'format common fields in record
            llProject(1) = 0
            llPrevCntr = tgChf.lCntrNo
        End If
    End If
    ilRet = btrGetDirect(hmChf, tmChf, imChfRecLen, llRecPosition, INDEXKEY1, BTRV_LOCK_NONE)
    ilRet = btrGetNext(hmChf, tmChf, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
Loop                                    'do while BTRV_ERR_NONE
'See if last contract needs to be created on disk
If tmGrf.iPerGenl(3) <> 0 Then           '<>0 indicates something already processed for this contract
    tmGrf.iPerGenl(1) = 1                'assume modification
    If tmGrf.iPerGenl(3) = 2 Then        'flag 2 denotes something current already found
        tmGrf.iPerGenl(1) = 0            'send to crystal flag for new only
    End If
    If tmGrf.lDollars(1) <> 0 Then       'dont write anything to disk if amount is 0
        tmGrf.lDollars(1) = tmGrf.lDollars(1) / 100
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
    End If
End If
'Process the projections for the same weeks (effective date)
'Build table of all valid projections by advt, potential code and slsp.
'Need to determine if projection is new, increase or decrease.
'Write out 1 record into GRF file containing the final result of previous to current.
ilCalType = 0                       'retrieve for std month
ilUpperAct = 0
ilRet = gObtainMnfForType("P", slTimeStamp, tlMMnf())   'populate Potential types (A,B,C)
ReDim Preserve tlActList(0 To 0) As ACTLIST
For ilClf = LBound(tgMSlf) To UBound(tgMSlf) - 1 Step 1
    tmPjfSrchKey.iSlfCode = tgMSlf(ilClf).iCode
    tmPjfSrchKey.iRolloverDate(0) = 0               'find all for this slsp
    tmPjfSrchKey.iRolloverDate(1) = 0
    ilRet = btrGetGreaterOrEqual(hmPjf, tmPjf, imPjfRecLen, tmPjfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'get all projection records
    Do While (ilRet = BTRV_ERR_NONE And tmPjf.iSlfCode = tgMSlf(ilClf).iCode) And (tmPjf.iSlfCode = ilSlfCode Or ilSlfCode = 0)
        gUnpackDate tmPjf.iRolloverDate(0), tmPjf.iRolloverDate(1), slStr                 'effec date of projected record
        llEnterDate = gDateValue(slStr)
        If llEnterDate >= llEarliestEntry - 7 And llEnterDate <= llLatestEntry And ilYear = tmPjf.iYear Then 'within previous and currnet weeks?
            ilFound = False
            llProject(1) = 0
            For llDate = llStdStartDates(1) To (llStdStartDates(2) - 1) Step 7
                slStr = Format$(llDate, "m/d/yy")
                llAmount = gGetWkDollars(ilCalType, slStr, tmPjf.lGross())
                llProject(1) = llProject(1) + llAmount
            Next llDate
            For ilTemp = 0 To ilUpperAct - 1 Step 1
                If tlActList(ilTemp).iAdfCode = tmPjf.iAdfCode And tlActList(ilTemp).iPotnCode = tmPjf.iMnfBus And tlActList(ilTemp).iSlfCode = tmPjf.iSlfCode Then
                    ilFound = True
                    Exit For
                End If
            Next ilTemp
            If Not (ilFound) Then                   'create new entry
                ReDim Preserve tlActList(0 To ilUpperAct) As ACTLIST
                tlActList(ilUpperAct).iAdfCode = tmPjf.iAdfCode
                tlActList(ilUpperAct).iPotnCode = tmPjf.iMnfBus
                tlActList(ilUpperAct).iSlfCode = tmPjf.iSlfCode
                tlActList(ilUpperAct).lcxfChgR = tmPjf.lcxfChgR
                If llEnterDate < llEarliestEntry Then               'previous entry
                    tlActList(ilUpperAct).iWeekFlag = 1
                    tlActList(ilUpperAct).lAmount = tlActList(ilUpperAct).lAmount - llProject(1)
                Else
                    tlActList(ilUpperAct).iWeekFlag = 2             'current entry only
                    tlActList(ilUpperAct).lAmount = tlActList(ilUpperAct).lAmount + llProject(1)
                End If
           
                ilUpperAct = ilUpperAct + 1         'increment for next new record
            Else                                    'update existing entry, must be increase or decrease
                tlActList(ilTemp).iWeekFlag = 3             'both prev & current exit
                If llEnterDate < llEarliestEntry Then               'previous entry
                    tlActList(ilTemp).lAmount = tlActList(ilTemp).lAmount - llProject(1)
                Else
                    tlActList(ilTemp).lAmount = tlActList(ilTemp).lAmount + llProject(1)
                End If
            End If
        End If
        ilRet = btrGetNext(hmPjf, tmPjf, imPjfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
Next ilClf                                          'next slsp
tgChf.lCntrNo = 0                                   'contract #s do not apply for the projections
'All projections for past and current weeks activity are built into memory.  Now  write them
'out to disk.
For ilTemp = 0 To ilUpperAct - 1 Step 1
    tmGrf.sBktType = "P"                  'orders flag (vs P = project), for sorting
    tmGrf.iPerGenl(1) = 1                'assume modification to projection
    If tlActList(ilTemp).iWeekFlag = 2 Then           'flag 2 denotes new only
        tmGrf.iPerGenl(1) = 0
    End If
    tmGrf.iPerGenl(2) = 0
    tmGrf.sDateType = " "
    tmGrf.lCode4 = tlActList(ilTemp).lcxfChgR
    For ilFound = LBound(tlMMnf) To UBound(tlMMnf) - 1 Step 1
        If tlMMnf(ilFound).iCode = tlActList(ilTemp).iPotnCode Then
            tmGrf.sDateType = tlMMnf(ilFound).sName     'indicate Potential code A,B,C
            Exit For
        End If
    Next ilFound
    tgChf.iAdfCode = tlActList(ilTemp).iAdfCode         'common routine assumes adv is in cntr buffer
    tgChf.iSlfCode(0) = tlActList(ilTemp).iSlfCode
    tmGrf.lDollars(1) = tlActList(ilTemp).lAmount
    mWriteSlsAct          'format common fields in record
    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
Next ilTemp
Erase tlActList, tlSofList, llProject, llStdStartDates
ilRet = btrClose(hmSlf)
ilRet = btrClose(hmSof)
ilRet = btrClose(hmCff)
ilRet = btrClose(hmClf)
ilRet = btrClose(hmGrf)
ilRet = btrClose(hmPjf)
ilRet = btrClose(hmChf)
End Sub
'
'
'                   Create Sales Analysis Summary prepass file
'                   Generate GRF file by vehicle.  Each record  contains the vehicle,
'                   plan $, Business on Books for current years Qtr, OOB w/ holds for
'                   current years qtr, slsp projection $(pjf), Last years same week OOB (chf),
'                   and last years Actual $ (from contracts)
Sub gCrSalesAna(llCurrStart As Long, llCurrEnd As Long, llPrevStart As Long, llPrevEnd As Long)
Dim slMnfStamp As String
Dim slAirOrder As String * 1                'from site pref - bill as air or ordered
ReDim ilLikePct(1 To 3) As Integer             'most likely percentage from potential code A, B & C
ReDim ilLikeCode(1 To 3) As Integer           'mnf most likely auto increment code for A, B, C
Dim ilPotnInx As Integer                    'index to ilLikePct (which % to use)
Dim ilLoop As Integer
Dim ilTemp As Integer
Dim ilSlsLoop As Integer
Dim ilRet As Integer
ReDim ilRODate(0 To 1) As Integer           'Effective Date to match retrieval of Projection record
Dim slDate As String
Dim slStr As String
Dim ilMonth As Integer
Dim ilYear As Integer
Dim llEnterFrom As Long                       'gather cnts whose entered date falls within llEnterFrom and llEnterTo
Dim llEnterTo As Long
ReDim llProject(1 To 2) As Long               'projected $, only using 1 bucket, common rtn needs assumes array
ReDim llLYDates(1 To 2) As Long               'range of  qtr dates for contract retrieval (this year)
ReDim llTYDates(1 To 2) As Long               'range of qtr dates for contract retrieval (last year)
ReDim llStartDates(1 To 2) As Long            'temp array for last year vs this years range of dates
Dim llLYGetFrom As Long                       'range of dates for contract access (this year)
Dim llLYGetTo As Long
Dim llTYGetFrom As Long                       'range of dates for contract access (last year)
Dim llTYGetTo As Long
Dim slStartDate As String                       'llLYGetFrom or llTYGetFrom converted to string
Dim slEndDate As String                         'llLYGetTo or LLTYGetTo converted to string
Dim ilBdMnfCode As Integer                      'budget name to get
Dim ilBdYear As Integer                         'budget year to get
Dim slNameCode As String
Dim slYear As String
Dim slCode As String
Dim slCntrTypes As String                       'valid contract types to access
Dim slCntrStatus As String                      'valid status (holds, orders, working, etc) to access
Dim ilHOState As Integer
Dim ilFound As Integer
Dim ilStartWk As Integer                        'starting week index to gather budget data
Dim ilEndWk As Integer                          'ending week index to gather budgets
Dim ilFirstWk As Integer                        'true if week 0 needs to be added when start wk = 1
Dim ilLastWk As Integer                         'true if week 53 needs to be added when end wk = 52
Dim llContrCode As Long                         'contr code from gObtainCntrforDate
Dim ilCurrentRecd As Integer
Dim ilPastFut As Integer                        'loop to process past contracts, then current contracts
Dim ilClf As Integer
Dim llAdjust As Long                            'Adjusted gross using the potential codes most likely %
Dim ilWeekDay As Integer
Dim ilCorpStd As Integer            '1 = corp, 2 = std
    hmChf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmChf)
        btrDestroy hmChf
        Exit Sub
    End If
    imChfRecLen = Len(tmChf)
    
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmChf)
        btrDestroy hmGrf
        btrDestroy hmChf
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmChf)
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmChf
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)
    
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmChf)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmChf
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmBvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmBvf, "", sgDBPath & "Bvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmBvf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmChf)
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmChf
        Exit Sub
    End If
    imBvfRecLen = Len(tmBvf)
    
    hmPjf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPjf, "", sgDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmBvf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmChf)
        btrDestroy hmPjf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmChf
        Exit Sub
    End If
    imPjfRecLen = Len(tmPjf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmBvf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmChf)
        btrDestroy hmMnf
        btrDestroy hmPjf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmChf
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    
    slAirOrder = tgSpf.sInvAirOrder     'inv all contracts as aired or ordered
    ilCorpStd = 2                       'force standard
    If RptSelIv!rbcSelCInclude(0).value Then     'corporate selected
        ilCorpStd = 1
    End If
    ReDim tlMMnf(1 To 1) As MNF
    'get all the Potential codes from MNF
    ilRet = gObtainMnfForType("P", slMnfStamp, tlMMnf())
    For ilLoop = 1 To UBound(tlMMnf) - 1 Step 1
        If Trim$(tlMMnf(ilLoop).sName) = "A" Then
            ilLikePct(1) = Val(tlMMnf(ilLoop).sUnitType)            'most likely percentage from potential code "A"
            ilLikeCode(1) = tlMMnf(ilLoop).iCode
        ElseIf Trim$(tlMMnf(ilLoop).sName) = "B" Then
                ilLikePct(2) = Val(tlMMnf(ilLoop).sUnitType)            'most likely percentage from potential code "B"
                ilLikeCode(2) = tlMMnf(ilLoop).iCode
        ElseIf Trim$(tlMMnf(ilLoop).sName) = "C" Then
                ilLikePct(3) = Val(tlMMnf(ilLoop).sUnitType)            'most likely percentage from potential code "C"
                ilLikeCode(3) = tlMMnf(ilLoop).iCode
        End If
    Next ilLoop
    'get all the dates needed to work with
    slDate = RptSelIv!edcSelCFrom.Text               'effective date entred
    'obtain the entered dates year based on the std month
    llTYGetTo = gDateValue(slDate)                     'gather contracts thru this date
    'setup Projection rollover date
    gPackDate slDate, ilRODate(0), ilRODate(1)
    slStr = gObtainEndStd(Format$(llTYGetTo, "m/d/yy"))
    gObtainMonthYear 0, slDate, ilMonth, ilYear           'get year  of effective date (to figure out the beginning of std year)
    slStr = "1/15/" & Trim$(Str$(ilYear))                 'Jan of std year effective dat entered
    If ilCorpStd = 1 Then
        llTYGetFrom = gDateValue(gObtainStartCorp(slStr, True))  'gather contracts from this date thru effective entered date
    Else
        llTYGetFrom = gDateValue(gObtainStartStd(slStr))  'gather contracts from this date thru effective entered date
    End If
    ilYear = Val(RptSelIv!edcSelCTo.Text)           'year requested
    'Determine this years quarter span
    ilLoop = (igMonthOrQtr - 1) * 3 + 1             'determine starting month based on qtr entred
    slStr = Trim$(Str$(ilLoop)) & "/15/" & Trim$(RptSelIv!edcSelCTo.Text)
    If ilCorpStd = 1 Then                       'corp
        llTYDates(1) = gDateValue(gObtainStartCorp(slStr, True))
    Else                                        'std
        llTYDates(1) = gDateValue(gObtainStartStd(slStr))
    End If
    llTYDates(2) = llTYDates(1) + 90                'end date of this year's quarter requested
    
    'Determine last years quarter span
    slStr = Trim$(Str$(ilLoop)) & "/15/" & Trim$(Str$(ilYear - 1))
    If ilCorpStd = 1 Then                               'corp
        llLYDates(1) = gDateValue(gObtainStartCorp(slStr, True)) 'start date of last years qtr
    Else
        llLYDates(1) = gDateValue(gObtainStartStd(slStr))  'start date of last years qtr
    End If
    llLYDates(2) = llLYDates(1) + 90                   'end date of last years qtr
    
    'Determine last years effective week
    slStr = "1/15/" & Trim$(Str$(ilYear - 1))               'Jan of std year effective dat entered for last yera
    If ilCorpStd = 1 Then                                   'std
        llLYGetFrom = gDateValue(gObtainStartCorp(slStr, True))       'gather contracts from last years std start for same # of days for this year
    Else
        llLYGetFrom = gDateValue(gObtainStartStd(slStr))        'gather contracts from last years std start for same # of days for this year
    End If
    llLYGetTo = llLYGetFrom + (llTYGetTo - llTYGetFrom)
    
    'Determine the Budget name selected
    slNameCode = tgRptSelBudgetCode(igBSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 1, "\", slStr)
    ilRet = gParseItem(slStr, 1, "\", slYear)
    slYear = gSubStr("9999", slYear)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilBdMnfCode = Val(slCode)
    ilBdYear = Val(slYear)
    
    ReDim tlSlsList(1 To 1) As SLSLIST          'array of vehicles and their sales
    'gather all budget records by vehicle for the requested year, totaling by quarter
    If Not mReadBvfRec(hmBvf, ilBdMnfCode, ilBdYear, tmBvfVeh()) Then
        Exit Sub
    End If
    
    slStartDate = Format$(llTYDates(1), "m/d/yy")
    slEndDate = Format$(llTYDates(2), "m/d/yy")
    'use startwk & endwk to gather budgets and slsp projections
    gObtainWkNo 0, slStartDate, ilStartWk, ilFirstWk          'determine the first week inx to accum (current year)
    gObtainWkNo 0, slEndDate, ilEndWk, ilLastWk               'determine the last week inx to accum  (current year)
    If ilCorpStd = 2 And ilFirstWk = 1 Then                   'if std and week 1 is start, always add week 0
        ilFirstWk = True
    End If
    ilFound = False
    For ilLoop = LBound(tmBvfVeh) To UBound(tmBvfVeh) - 1 Step 1
        For ilSlsLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1
            If tmBvfVeh(ilLoop).iVefCode = tlSlsList(ilSlsLoop).iVefCode Then
                ilFound = True
                Exit For
            End If
        Next ilSlsLoop
        If Not ilFound Then
            tlSlsList(UBound(tlSlsList)).iVefCode = tmBvfVeh(ilLoop).iVefCode
            ReDim Preserve tlSlsList(1 To UBound(tlSlsList) + 1)
        End If
        'ilSlsLoop contains index to the correct vehicle
        For ilTemp = ilStartWk To ilEndWk - 1 Step 1
            tlSlsList(ilSlsLoop).lPlan = tlSlsList(ilSlsLoop).lPlan + tmBvfVeh(ilLoop).lGross(ilTemp)
        Next ilTemp
        If ilFirstWk Then       'adjust for the partial weeks at the beginning or end of the year
                                'due to corp or calendar months
            tlSlsList(ilSlsLoop).lPlan = tlSlsList(ilSlsLoop).lPlan + tmBvfVeh(ilLoop).lGross(0)
        End If
        If ilLastWk Then
            tlSlsList(ilSlsLoop).lPlan = tlSlsList(ilSlsLoop).lPlan + tmBvfVeh(ilLoop).lGross(53)
        End If
    Next ilLoop
    
    'gather all Slsp projection records for the matching rollover date (exclude current records)
    ReDim tmTPjf(0 To 0) As PJF
    ilRet = gObtainPjf(hmPjf, ilRODate(), tmTPjf())                 'Read all applicable Projection records into memory
    'Build slsp projection $ just gathered into vehicle buckets
    For ilLoop = LBound(tmTPjf) To UBound(tmTPjf) Step 1
        ilPotnInx = 0
        For ilFound = 1 To 3 Step 1
            If tmTPjf(ilLoop).iMnfBus = ilLikeCode(ilFound) Then
                ilPotnInx = ilFound
                Exit For
            End If
        Next ilFound
        If ilPotnInx > 0 Then           'potential code exists
            For ilSlsLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1
                If tlSlsList(ilSlsLoop).iVefCode = tmTPjf(ilLoop).iVefCode Then
                    llAdjust = 0
                    For ilTemp = ilStartWk To ilEndWk - 1 Step 1
                        llAdjust = llAdjust + tmTPjf(ilLoop).lGross(ilTemp)
                    Next ilTemp
                    If ilFirstWk Then       'adjust for the partial weeks at the beginning or end of the year
                                            'due to corp or calendar months
                        'llAdjust = llAdjust + tlSlsList(ilSlsLoop).lProj + tmTPjf(ilLoop).lGross(0)
                        llAdjust = llAdjust + tmTPjf(ilLoop).lGross(0)
                    End If
                    If ilLastWk Then
                        llAdjust = llAdjust + tmTPjf(ilLoop).lGross(53)
                    End If
                    llAdjust = (llAdjust * ilLikePct(ilPotnInx)) \ 100  'adjust the gross based on the potential codes most likely %
                    tlSlsList(ilSlsLoop).lProj = tlSlsList(ilSlsLoop).lProj + llAdjust
                    Exit For
                End If
            Next ilSlsLoop
        End If
    Next ilLoop
    'gather all contracts whose entered date is equal or prior to the requested date (gather from beginning of std year to
    'input date
    slCntrTypes = gBuildCntTypes()
    slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)
    For ilPastFut = 1 To 2 Step 1
        If ilPastFut = 1 Then                       'past
            slStartDate = Format$(llLYDates(1), "m/d/yy")       'gather all cntrs whose start/end dates fall within requested qtr (last year)
            slEndDate = Format$(llLYDates(2), "m/d/yy")
            llStartDates(1) = llLYDates(1)
            llStartDates(2) = llLYDates(2)
            llEnterFrom = llLYGetFrom                           'gather all cntrs whose entered date falls within these dates
            llEnterTo = llLYGetTo
       Else                                         'current
            slStartDate = Format$(llTYDates(1), "m/d/yy")        'gather all cntrs whose start/end dates fall within requested qtr (this year)
            slEndDate = Format$(llTYDates(2), "m/d/yy")
            llStartDates(1) = llTYDates(1)
            llStartDates(2) = llTYDates(2)
            llEnterFrom = llTYGetFrom           'gather cnts whose entered date falls within these dates
            llEnterTo = llTYGetTo
        End If
        ilRet = gObtainCntrForDate(RptSelIv, slStartDate, slEndDate, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
        For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
            'project the $
            llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
            ilRet = gObtainCntr(hmChf, hmClf, hmCff, llContrCode, False, tgChf, tgClf(), tgCff())
            ilFound = False
            If ilPastFut = 2 Then       'if current, need to test entered date against the requested effective
                gUnpackDateLong tgChf.iOHDDate(0), tgChf.iOHDDate(1), llAdjust
                If llAdjust <= llEnterTo Then       'entered date must be entered thru effectve date
                    ilFound = True
                End If
            Else                        'Past
                ilFound = True          'past get all cnts affecting the qtr to get actuals as well as same wee last year
            End If
            If ilFound Then
                For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                    llProject(1) = 0                'init bkts to accum qtr $ for this line
                    tmClf = tgClf(ilClf).ClfRec
                    If slAirOrder = "O" Then                'invoice all contracts as ordered
                        If tmClf.sType <> "H" Then          'ignore all hidden lines for ordered billing, should be Pkg or conventional lines
                            gBuildFlights ilClf, llStartDates(), 1, 2, llProject()
                        End If
                    Else                                    'inv all contracts as aired
                        If tmClf.sType = "H" Then             'but if from pkg and hidden line, ignore hidd
                            'if hidden, will project if assoc. package is set to invoice as aired (real)
                            For ilTemp = LBound(tgClf) To UBound(tgClf) - 1    'find the assoc. pkg line for these hidden
                                If tmClf.iPkLineNo = tgClf(ilTemp).ClfRec.iLine Then
                                    If tgClf(ilTemp).ClfRec.sType = "A" Then        'does the pkg line reflect bill as aired?
                                        gBuildFlights ilClf, llStartDates(), 1, 2, llProject()  'pkg bills as aired, project the hidden line
                                    End If
                                    Exit For
                                End If
                            Next ilTemp
                        Else                            'conventional, VV, or Pkg line
                            If tmClf.sType <> "A" Then  'if this package line to be invoiced aired (real times),
                                                        'it has already been projected above with the hidden line
                                gBuildFlights ilClf, llStartDates(), 1, 2, llProject()
                            End If
                        End If
                    End If
                    'Accumulate the $ projected into the vehicles buckets
                    If llProject(1) > 0 Then
                        llProject(1) = llProject(1) \ 100               'drop pennies
                        For ilSlsLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1
                            If tlSlsList(ilSlsLoop).iVefCode = tmClf.iVefCode Then
                                If ilPastFut = 1 Then               'past year  (holds & orders are combined)
                                    'if outside the effective date, it's also actuals for the qtr
                                    If llAdjust > llEnterTo Then    'lladjust is the contracts entred date.  If the entered date is greater than
                                                                    'the effective date, it belongs in the actual last year column
                                        tlSlsList(ilSlsLoop).lLYAct = tlSlsList(ilSlsLoop).lLYAct + llProject(1)
                                        tlSlsList(ilSlsLoop).lLYAct = tlSlsList(ilSlsLoop).lLYAct + llProject(1)
                                    Else
                                        tlSlsList(ilSlsLoop).lLYWeek = tlSlsList(ilSlsLoop).lLYWeek + llProject(1)
                                        tlSlsList(ilSlsLoop).lLYAct = tlSlsList(ilSlsLoop).lLYAct + llProject(1)
                                    End If
                                    Exit For
                                Else                                'current year, holds and orders are added together
                                    If tgChf.sStatus = "H" Or tgChf.sStatus = "G" Then    'hold or unsch hold
                                        tlSlsList(ilSlsLoop).lTYActHold = tlSlsList(ilSlsLoop).lTYActHold + llProject(1)
                                        Exit For
                                    Else                            'order or unsch order
                                        tlSlsList(ilSlsLoop).lTYAct = tlSlsList(ilSlsLoop).lTYAct + llProject(1)
                                        Exit For
                                    End If
                                End If
                            End If
                        Next ilSlsLoop
                    End If
                Next ilClf                      'process nextline
            End If                              'llAdjust falls within requested dates
        Next ilCurrentRecd
    Next ilPastFut
    'Setup last year's qtr column heading
    'ilYear contains starting year
    ilMonth = RptSelIv!edcSelCTo1.Text              'month
    slDate = Trim$(Str$(((ilMonth - 1) * 3 + 1))) & "/15/" & Trim$(Str$(ilYear))
    slDate = gObtainStartStd(slDate)
    If ilMonth = 1 Then
        tmGrf.sGenDesc = "1st"
    ElseIf ilMonth = 2 Then
        tmGrf.sGenDesc = "2nd"
    ElseIf ilMonth = 3 Then
        tmGrf.sGenDesc = "3rd"
    Else
        tmGrf.sGenDesc = "4th"
    End If
    tmGrf.sGenDesc = Trim$(tmGrf.sGenDesc) & " Qtr" & Str$(ilYear - 1)    'add Year
    
    slDate = Format$(llLYGetTo, "m/d/yy")
    gPackDate slDate, ilMonth, ilYear
    tmGrf.iDate(0) = ilMonth                'last year's week (for last years column heading)
    tmGrf.iDate(1) = ilYear
    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
    tmGrf.iGenTime(1) = igNowTime(1)
    tmGrf.iStartDate(0) = ilRODate(0)             'effective date entered
    tmGrf.iStartDate(1) = ilRODate(1)
    tmGrf.iCode2 = ilBdMnfCode                          'budget name
    tmGrf.iPerGenl(1) = ilCorpStd             '1 = corp, 2 = std
    For ilLoop = 1 To UBound(tlSlsList) - 1 Step 1         'write a record per vehicle
        If tlSlsList(ilLoop).lPlan + tlSlsList(ilLoop).lTYAct + tlSlsList(ilLoop).lProj + tlSlsList(ilLoop).lLYWeek + tlSlsList(ilLoop).lLYAct <> 0 Then
            tmGrf.iVefCode = tlSlsList(ilLoop).iVefCode
            tmGrf.lDollars(1) = tlSlsList(ilLoop).lPlan         'current year, plan $
            tmGrf.lDollars(2) = tlSlsList(ilLoop).lTYAct        'current year, orders
            tmGrf.lDollars(3) = tlSlsList(ilLoop).lTYActHold    'current year, holds
            tmGrf.lDollars(4) = tlSlsList(ilLoop).lProj         'current rollover
            tmGrf.lDollars(5) = tlSlsList(ilLoop).lLYWeek       'last years, same week
            tmGrf.lDollars(6) = tlSlsList(ilLoop).lLYAct        'last years, actuals
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
    Next ilLoop
    Erase tlSlsList, tlMMnf
    Erase tmTPjf, tlChfAdvtExt, tmBvfVeh
    ilRet = btrClose(hmBvf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmChf)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmMnf)
End Sub
'
'
'                           Sales commission or Billed & Booked
'                           Salesperson Commissions
'
Sub gCrSalesComm()
    Dim illistindex As Integer              'report option
    Dim ilRet As Integer
    Dim ilLoop  As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilFound As Integer
    Dim ilLoopOnFile As Integer             '2 passes, 1 for History, then Receivables
    Dim slStr As String
    Dim llAmt As Long
    Dim ilMonthNo As Integer
    Dim llCalEndDate As Long                'jan 1 of current year
    Dim llCalStartDate As Long              'cal month end date of month requested
    Dim llDate As Long
    Dim ilTemp As Integer
    Dim ilFoundMonth As Integer
    Dim slStartCal As String                'start of calendar year to gather transactions-dynamic (used in projections)
    Dim slEndCal As String                  'end date of cal month to gathe trans -dynamic  (used in projections)
    Dim llTempStart As Long
    Dim llTempEnd As Long
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim slAmount As String
    Dim slDollar As String
    Dim slPct As String
    Dim llDollar As Long
    Dim llGrossDollar As Long
    Dim llNetDollar As Long
    Dim slGrossOrNet As String              'Base commission on G = Gross or N = Net
    Dim ilFoundSlsp As Integer
    Dim ilLoopSlsp As Integer
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRvf)
        btrDestroy hmRvf
        btrDestroy hmGrf
        Exit Sub
    End If
    imRvfRecLen = Len(tmRvf)
    hmChf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmChf)
        btrDestroy hmChf
        btrDestroy hmRvf
        btrDestroy hmGrf
        Exit Sub
    End If
    imChfRecLen = Len(tmChf)
    
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        btrDestroy hmChf
        btrDestroy hmSlf
        btrDestroy hmRvf
        btrDestroy hmGrf
        Exit Sub
    End If
    imSlfReclen = Len(tmSlf)
    slGrossOrNet = "N"                  'force comm on net for now
    'Determine calendar month requested, and retrieve all History and Receivables
    'records that fall within the beginning of the cal year and end of calendar month requested
     illistindex = RptSelIv!lbcRptType.ListIndex
'     'If ilListIndex = COMM_SALESCOMM Then
'        slStr = RptSelIv!edcSelCFrom.Text             'month in text form (jan..dec)
'        gGetMonthNoFromString slStr, ilMonthNo      'getmonth #
'        slStr = Trim$(Str$(ilMonthNo)) & "/1/" & Trim$(RptSelIv!edcSelCFrom1.Text)
'        slStr = gObtainEndCal(slStr)               'obtain cal month for end date to gather
'        llCalEndDate = gDateValue(slStr)
'        slStr = "1/1/" & Trim$(RptSelIv!edcSelCFrom1.Text)
'        llCalStartDate = gDateValue(slStr)
'    Else                                        'projections
'        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr       'convert last bdcst billing date to string
'        gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
'        ilLoop = Val(RptSelIv!edcSelCFrom.Text)            'year user entered
'        If Val(slYear) = ilLoop Then
'            slStr = gObtainEndCal(slStr)
'            llCalEndDate = gDateValue(slStr)
'            slStr = "1/1/" & Trim$(RptSelIv!edcSelCFrom.Text)
'            llCalStartDate = gDateValue(slStr)
'        ElseIf Val(slYear) > ilLoop Then            'actuals only, last date billed is later than year requested
'            slStr = "12/31/" & Trim$(RptSelIv!edcSelCFrom.Text) 'force end of cal year
'            llCalEndDate = gDateValue(slStr)
'            slStr = "1/1/" & Trim$(RptSelIv!edcSelCFrom.Text)
'            llCalStartDate = gDateValue(slStr)
'        Else                                        'projections only, last date billed is prior to year requested
'            'continue to build file with contract projections
'            ilRet = mBuildCommProj()
'            ilRet = btrClose(hmGrf)
'            ilRet = btrClose(hmRvf)
'            ilRet = btrClose(hmChf)
'            ilRet = btrClose(hmSlf)
'            Exit Sub
'        End If
'    End If
'    For ilLoopOnFile = 1 To 2 Step 1                 '2 passes, first History, then Receivables
'        'handles and buffers for PHF and RVF will be the same
'
'        ilRet = btrGetFirst(hmRvf, tmRvf, imRvfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'
'        Do While ilRet = BTRV_ERR_NONE
'            'ilFound = False
'            'If RptSel!ckcAll Then
'            '    ilFound = True
'            'Else
'            '    For ilLoop = 0 To RptSel!lbcSelection(2).ListCount - 1 Step 1
'            '        If RptSel!lbcSelection(2).Selected(ilLoop) Then              'selected slsp
'            '            slNameCode = Traffic!lbcSalesperson.List(ilLoop)         'pick up slsp code
'            '            ilRet = gParseItem(slNameCode, 2, "\", slCode)
'            '            If Val(slCode) = tmRvf.iSlfCode Then
'            '               ilFound = True
'            '                Exit For
'            '            End If
'            '         End If
'            '    Next ilLoop
'            'End If
'            gPDNToLong tmRvf.sNet, llAmt
'            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStr
'            llDate = gDateValue(slStr)                  'convert trans date to test if within requested limits
'            'valid record must be an "Invoice" or "Adjustment" type, non-zero amount, and transaction date within the start date of the
'            'cal year and end date of the current cal month requested
'            If ((Left$(tmRvf.sTranType, 1) = "I" Or Left$(tmRvf.sTranType, 1) = "A") And llAmt <> 0 And llDate >= llCalStartDate And llDate <= llCalEndDate And tmRvf.sCashTrade = "C") Then
'                'get contract from history or rec file
'                tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
'                tmChfSrchKey1.iCntRevNo = 32000
'                tmChfSrchKey1.iPropVer = 32000
'                ilRet = btrGetGreaterOrEqual(hmChf, tmChf, imChfRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
'                'Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo <> tmRvf.lCntrNo Or (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M"))
'                Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M")
'                     ilRet = btrGetNext(hmChf, tmChf, imChfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'                Loop
'                If ((ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus = "F" Or tmChf.sSchStatus = "M")) Then
'                    'format remainder of record
'                    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
'                    tmGrf.iGenDate(1) = igNowDate(1)
'                    tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
'                    tmGrf.iGenTime(1) = igNowTime(1)
'                    tmGrf.lChfCode = tmChf.lCode           'contr internal code
'                    tmGrf.iDateGenl(0, 1) = tmRvf.iTranDate(0)    'date billed or paid
'                    tmGrf.iDateGenl(1, 1) = tmRvf.iTranDate(1)
'                    If ilLoopOnFile = 1 Then
'                        tmGrf.sBktType = "H"                    'let crystal know these records are histroy/receivables (vs contracts)
'                    Else
'                        tmGrf.sBktType = "R"
'                    End If
'                    ilTemp = 0
'                    For ilLoop = 0 To 9 Step 1
'                        If tmChf.islfCode(ilLoop) > 0 Then
'                            ilTemp = ilTemp + 1
'                        End If
'                    Next ilLoop
'                    If ilTemp = 1 Then                      'only 1 slsp, force to 100% (no splits)
'                        tmChf.lComm(0) = 1000000             'force 100.0000%
'                    End If
'                    If ilListIndex = COMM_SALESCOMM Then  'format sales comm report different than comm projections
'                        tmGrf.lDollars(1) = tmRvf.lInvNo          'Invoice #
'                        For ilLoop = 0 To 9 Step 1          'see if there are any split commissions
'                            If tmChf.lComm(ilLoop) > 0 Then
'
'
'                                ilFoundSlsp = False
'                                If RptSelIv!ckcAll Then
'                                    ilFoundSlsp = True
'                                Else
'                                    For ilLoopSlsp = 0 To RptSelIv!lbcSelection(2).ListCount - 1 Step 1
'                                        If RptSelIv!lbcSelection(2).Selected(ilLoopSlsp) Then              'selected slsp
'                                            slNameCode = tgSalesperson(ilLoopSlsp).sKey      'pick up slsp code
'                                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
'                                            If Val(slCode) = tmChf.islfCode(ilLoop) Then
'                                                ilFoundSlsp = True
''                                                Exit For
 '                                           End If
 '                                       End If
 '                                   Next ilLoopSlsp
 '                               End If
 '                               If ilFoundSlsp Then
 '                                   tmGrf.islfCode = tmChf.islfCode(ilLoop) 'slsp code
 '
 '                                   tmGrf.lDollars(7) = tmChf.lComm(ilLoop)         'slsp share in % (xxx.xxxx)
 '                                   slPct = gLongToStrDec(tmChf.lComm(ilLoop), 4)
 '                                   gPDNToStr tmRvf.sGross, 2, slAmount
 ''                                   slDollar = gMulStr(slAmount, slPct)
 '                                   tmGrf.lDollars(2) = Val(gRoundStr(slDollar, "01.", 0))
 '
 '
 '                                   gPDNToStr tmRvf.sNet, 2, slAmount
'                                    slDollar = gMulStr(slAmount, slPct)
'                                    tmGrf.lDollars(3) = Val(gRoundStr(slDollar, "01.", 0))
'
'
 '                                   tmGrf.lDollars(4) = 0                           'merchandising amount, init for now
 '                                   tmGrf.lDollars(5) = 0                           'promotions amount, init for now
 ''                                   tmSlfSrchKey.iCode = tmChf.islfCode(ilLoop)    'find slsp record for comm
 '                                   ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)'get matching slsp recd
 '                                   If ilRet = BTRV_ERR_NONE Then
 '                                       slPct = gLongToStrDec(tmSlf.lUnderComm, 4)      'convert slsp comm % to packed dec.
 '                                       'llDollar = tmGrs.lDollars(2) - tmGrf.lDollars(4) - tmGrf.lDollars(5)
 '                                       If slGrossOrNet = "G" Then
 '                                           llDollar = tmGrf.lDollars(2) - tmGrf.lDollars(4) - tmGrf.lDollars(5)
 '                                       Else
 '                                           llDollar = tmGrf.lDollars(3) - tmGrf.lDollars(4) - tmGrf.lDollars(5)
 '                                       End If
 '                                       slAmount = gLongToStrDec(llDollar, 2)           'adjusted rate to calc slsp comm. from
 '                                       slAmount = gMulStr(slAmount, slPct)
 '                                       tmGrf.lDollars(6) = Val(gRoundStr(slAmount, "01.", 0))
 '
 '                                       tmGrf.isofCode = tmSlf.isofCode                 'office code
 '                                   End If
 '                                   'tmgrf.ldollars(1) = inv#
 '                                   'tmgrf.ldollars(2) = gross
 '                                   'tmgrs.ldollars(3) = net
 '                                   'tmgrs.ldollars(4) = merch $ (currently 0)
 '                                   'tmgrs.ldollars(5) = promo $ (currently 0)
 '                                   'tmgrs.ldollars(6) = slsp comm (calc from gross or net minus merch minus promo)
 '                                   'tmgrs.ldollars(7) = slsp split %
 '                                   ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
 '                               End If                          ' if ilFoundSlsp
 '                           End If                              'lcomm(ilLoop) > 0
 '                       Next ilLoop                             'loop for 10 slsp possible splits
 '                   Else                                        'projections
 '                       'determine the month that this transaction falls within
 '                       ilFoundMonth = False
 '                       llTempStart = llCalStartDate
 '                       slStartCal = Format$(llTempStart, "m/d/yy")
 '                       slEndCal = gObtainEndCal(slStartCal)
 '                       llTempEnd = gDateValue(slEndCal)
 '                       gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slCode
 '                           llDate = gDateValue(slCode)
 '                       For ilMonthNo = 0 To 11 Step 1         'loop thru cal months
 '                           If llDate >= llTempStart And llDate <= llTempEnd Then
 '                               ilFoundMonth = True
 '                               Exit For
 '                           Else
 '                               slStartCal = Format$(llTempStart, "m/d/yy")
 '                               gObtainYearMonthDayStr slStartCal, True, slYear, slMonth, slDay
 '                               slMonth = Str$(Val(slMonth) + 1)
 '                               slStartCal = Trim$(slMonth) & "/" & Trim$(slDay) & "/" & Trim$(slYear)
 '                               llTempStart = gDateValue(slStartCal)
 '                               slEndCal = gObtainEndCal(slStartCal)
 '                               llTempEnd = gDateValue(slEndCal)
 '                               If llTempEnd > llCalEndDate Then           'past the last billing period
 '                                   Exit For
 '                               End If
 '                           End If
'                        Next ilMonthNo
'                        If ilFoundMonth Then
'                            ilMonthNo = ilMonthNo + 1           'adjust for index into buckets
'                            For ilLoop = 0 To 9 Step 1          'see if there are any split commissions
 '                               For ilTemp = 1 To 14 Step 1     'init the years $ buckets
 '                                   tmGrf.lDollars(ilTemp) = 0
 '                               Next ilTemp
 '                               If tmChf.lComm(ilLoop) > 0 Then
 '
 '
 '                                   ilFoundSlsp = False
 '                                   If RptSelIv!ckcAll Then
 '                                       ilFoundSlsp = True
 '                                   Else
 '                                       For ilLoopSlsp = 0 To RptSelIv!lbcSelection(2).ListCount - 1 Step 1
 '                                           If RptSelIv!lbcSelection(2).Selected(ilLoopSlsp) Then              'selected slsp
 '                                               slNameCode = tgSalesperson(ilLoopSlsp).sKey        'pick up slsp code
 '                                               ilRet = gParseItem(slNameCode, 2, "\", slCode)
 '                                               If Val(slCode) = tmChf.islfCode(ilLoop) Then
 '                                                   ilFoundSlsp = True
 '                                                   Exit For
 '                                               End If
 '                                           End If
 '                                       Next ilLoopSlsp
 '                                   End If
 '
 '                                   If ilFoundSlsp Then
 '                                       tmGrf.islfCode = tmChf.islfCode(ilLoop) 'slsp code
 '                                       slPct = gLongToStrDec(tmChf.lComm(ilLoop), 4)           'slsp split share in %
 '
 '                                       gPDNToStr tmRvf.sGross, 2, slAmount
 '                                       slDollar = gMulStr(slPct, slAmount)                 'slsp gross portion of possible split
 '                                       tmGrf.lDollars(13) = Val(gRoundStr(slDollar, "01.", 0))
 '                                       llGrossDollar = tmGrf.lDollars(13)
 ''                                       gPDNToStr tmRvf.sNet, 2, slAmount
 '                                       slDollar = gMulStr(slPct, slAmount)                 'slsp net portion of possible split
 '                                       llNetDollar = Val(gRoundStr(slDollar, "01.", 0))
 '
 '                                       'obtain slsp record for % of commission off the net
 '                                       tmSlfSrchKey.iCode = tmChf.islfCode(ilLoop)    'find slsp record for comm
 '                                       ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)'get matching slsp recd
 '                                       If ilRet = BTRV_ERR_NONE Then
  '                                          'slsp comm calc from gross minus merch minus promotion values.  Adjust value when
  '                                          'promo & merch has been determined.  (currently 0)
  '                                          slPct = gLongToStrDec(tmSlf.lUnderComm, 4)      'convert slsp comm % to packed dec.
  '                                          If slGrossOrNet = "G" Then
  '                                              slAmount = gLongToStrDec(llGrossDollar, 2)           'get slsp split gross share
  '                                          Else
  '                                              slAmount = gLongToStrDec(llNetDollar, 2)             'get slsp split gross share
  '                                          End If
  ''                                          slDollar = gMulStr(slAmount, slPct)                  'calc slsp comm.
  '                                          tmGrf.lDollars(ilMonthNo) = Val(gRoundStr(slDollar, "01.", 0))
   '
   '                                         tmGrf.isofCode = tmSlf.isofCode                 'office code
   '                                     End If
   '                                    'Bucket 13 contains total gross for slsp
   '                                    'bucket 1-12 contains slsp comm amount calc from his gross allocation
   '                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
   '                                 End If                              'lcomm(ilLoop) > 0
   '                             End If
   '                         Next ilLoop                             'loop for 10 slsp possible splits
   '                     End If                                      'if foundmonth
   '                 End If                                      'sales comm or projections
   '             End If
   '         End If                                          'contr # doesnt match or not a fully sched contr
   '         ilRet = btrGetNext(hmRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
   '     Loop
   '     ilRet = btrClose(hmRvf)
   '
   '     hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
   '     ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
   '     If ilRet <> BTRV_ERR_NONE Then
   '         ilRet = btrClose(hmRvf)
   '         btrDestroy hmRvf
   '         btrDestroy hmChf
   '         btrDestroy hmGrf
   '         Exit Sub
   '     End If
   '     imRvfRecLen = Len(tmRvf)
   ' Next ilLoopOnFile                                   '2 passes, first History, then Receivbles
   '
   ' 'if Commissions Projections, continue to build file with contracts
   ' If ilListIndex = COMM_PROJECTION Then
   '     ilRet = mBuildCommProj()
   ' End If
   '
   ' ilRet = btrClose(hmGrf)
   ' ilRet = btrClose(hmRvf)
   ' ilRet = btrClose(hmChf)
   ' ilRet = btrClose(hmSlf)
End Sub
Sub gCrVehCPPCPM()
Dim ilRif As Integer                'loop variable for DP Rates
Dim ilTest As Integer
Dim ilFound As Integer
Dim ilFoundAgain As Integer
Dim ilDay As Integer
ReDim ilValidDays(0 To 6) As Integer
Dim llTPrice As Long
Dim ilNoWks As Integer
Dim llDate As Long
Dim slDate As String
Dim llWkPrice As Long
Dim ilRCCode As Integer
Dim ilRdf As Integer                'loop variable for DP
Dim ilDnfCode As Integer
Dim ilmnfDemo As Integer          'demo name code into mnf
Dim ilDemoLoop As Integer            'Index into Demo processing
Dim ilBookInx As Integer            'loop to create ANR records form BOOKGEN array
Dim ilRet As Integer
Dim slName As String
Dim ilEffYear As Integer             'year of rate card
Dim llEffDate As Long               'Effectve date entered
ReDim ilEffDate(0 To 1) As Integer    'effective date entered format for Crystal
Dim llRCStartDate As Long           'Rate Card Start Date
Dim llEffEndDate As Long            'effective end date (currently only 1 week span)
Dim slStr As String
Dim ilLoop As Integer
Dim ilRCWkNo As Integer             'week processing to gather rates from rif (currently only 1 week to process)
Dim slCode As String
Dim ilVeh As Integer              'loop variable for the vehicles to process
Dim ilSaveVeh As Integer            'vehicle code processing
ReDim ilVehicles(1 To 1) As Integer 'array of valid vehicles to process
ReDim ilDemoList(0 To 0) As Integer  'array of demo categories to process
ReDim tmMRif(1 To 1) As RIF
ReDim tmBookGen(0 To 0) As BOOKGEN
ReDim tmTAnr(0 To 0) As ANR
ReDim ilDPList(1 To 1) As Integer       'list of dayparts generated for vehicle processing
'the following variables are for the routines to retrieve the rating data
''ggetdemoAvgAud and gAvgAudToLnResearch
ReDim ilWkSpotCount(1 To 1) As Integer
ReDim llWkActPrice(1 To 1) As Long
ReDim llWkAvgAud(1 To 1) As Long
Dim llLnCost As Long
Dim llAvgAudAvg As Long
ReDim ilWkRating(1 To 1) As Integer
Dim ilLnAvgRating As Integer
ReDim llWkGrImp(1 To 1) As Long
ReDim llWkGRP(1 To 1) As Long
Dim llLnGRP As Long
Dim llLnGrImp As Long
Dim llCPP As Long
Dim llCPM As Long
Dim llPop As Long
Dim ilMnfSocEco As Integer
Dim llOvStartTime As Long
Dim llOvEndtime As Long
Dim llAvgAud As Long
'**** end of varaibles required for Research routines
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Sub
    End If
    imVefRecLen = Len(hmVef)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imMnfRecLen = Len(hmMnf)
    hmRcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRcf, "", sgDBPath & "Rcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRcf)
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imRcfRecLen = Len(hmRcf)
    hmRif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRif, "", sgDBPath & "Rif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRif)
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imRifRecLen = Len(hmRif)
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmRdf)
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imRdfRecLen = Len(tmRdf)
    hmDrf = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmDrf)
        btrDestroy hmDrf
        btrDestroy hmAnr
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imDrfRecLen = Len(tmDrf)
    hmAnr = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnr, "", sgDBPath & "Anr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnr)
        btrDestroy hmAnr
        btrDestroy hmRdf
        btrDestroy hmRif
        btrDestroy hmRcf
        btrDestroy hmMnf
        btrDestroy hmVef
        Exit Sub
    End If
    imAnrRecLen = Len(tmAnr)
    
    slStr = RptSelIv!edcSelCFrom.Text           'effective date
    'insure its a Monday
    llEffDate = gDateValue(slStr)
    ilDay = gWeekDayLong(llEffDate)
    Do While ilDay <> 0
        llEffDate = llEffDate - 1
        ilDay = gWeekDayLong(llEffDate)
    Loop
    slStr = Format$(llEffDate, "m/d/yy")
    gPackDate slStr, ilEffDate(0), ilEffDate(1)
    
    'get sunday date for the R/C year
    llEffEndDate = llEffDate + 6
    slStr = Format$(llEffEndDate, "m/d/yy")              'Sunday will always be the correct year since we're dealing with Standard month
    gPackDate slStr, ilDay, ilEffYear
    ilRet = gObtainRcfRifRdf()                'bring in 3 files in global arrays
    ilRCCode = 0                                'no r/c definition yet
    
    'find matching R/c from list box to the one entered.  If found a match, retrieve the RC Code
    For ilLoop = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
        slStr = RptSelIv!edcSelCFrom1.Text
        If Trim$(slStr) = Trim$(tgMRcf(ilLoop).sName) Then
            ilRCCode = tgMRcf(ilLoop).iCode
            gUnpackDate tgMRcf(ilLoop).iStartDate(0), tgMRcf(ilLoop).iStartDate(1), slName
            llRCStartDate = gDateValue(slName)
            Exit For
        End If
    Next ilLoop
    'get Book pointer from the Book selected
    For ilLoop = 0 To RptSelIv!lbcSelection(4).ListCount - 1 Step 1
        If (RptSelIv!lbcSelection(4).Selected(ilLoop)) Then
            slName = tgBookName(ilLoop).sKey 'Traffic!lbcVehicle.List(ilVehicle)
            ilRet = gParseItem(slName, 1, "\", slStr)
            ilRet = gParseItem(slStr, 3, "|", slStr)
            ilRet = gParseItem(slName, 2, "\", slCode)
            ilDnfCode = Val(slCode)             'Book name pointer
            Exit For
        End If
    Next ilLoop
    'Build array (tmMRif) of all valid Rates for each Vehicle's daypart to cut down on amount of processing
    For ilRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
        llTPrice = 0
        For ilTest = 0 To 53 Step 1
            llTPrice = llTPrice + tgMRif(ilRif).lRate(ilTest)
        Next ilTest
        If (ilRCCode = tgMRif(ilRif).iRcfCode) And (ilEffYear = tgMRif(ilRif).iYear) And (llTPrice > 0) Then
            'test for selective vehicle
            For ilLoop = 0 To RptSelIv!lbcSelection(3).ListCount - 1 Step 1
                If (RptSelIv!lbcSelection(3).Selected(ilLoop)) Then
                    slName = tgVehicle(ilLoop).sKey 'Traffic!lbcVehicle.List(ilVehicle)
                    ilRet = gParseItem(slName, 1, "\", slStr)
                    ilRet = gParseItem(slStr, 3, "|", slStr)
                    ilRet = gParseItem(slName, 2, "\", slCode)
                    'Build vehicle table of ones to process
                    ilFound = False
                    For ilTest = LBound(ilVehicles) To UBound(ilVehicles) - 1 Step 1
                        If ilVehicles(ilTest) = Val(slCode) Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilTest
                    If Not ilFound Then
                        ilVehicles(UBound(ilVehicles)) = Val(slCode)
                        ReDim Preserve ilVehicles(1 To UBound(ilVehicles) + 1)
                    End If
                    'If vehicle code matches daypart rates vehicle code, save the rate image
                    If Val(slCode) = tgMRif(ilRif).iVefCode Then
                        tmMRif(UBound(tmMRif)) = tgMRif(ilRif)
                        ReDim Preserve tmMRif(1 To UBound(tmMRif) + 1) As RIF
                        ilLoop = RptSelIv!lbcSelection(3).ListCount 'stop the loop
                    End If
                End If
            Next ilLoop
        End If
    Next ilRif
    'Build table of all demos that will be obtained.  This list corresponds to the
    '13 set of buckets sent in ANR (ie. the 1st 13 demos will always be in the same
    'relative index of anr.ipctsellout; the next 13 will always be in the same
    'relative index of anr.ipctsellout; etc.)
    For ilDemoLoop = 0 To RptSelIv!lbcSelection(2).ListCount - 1 Step 1
        If (RptSelIv!lbcSelection(2).Selected(ilDemoLoop)) Then
            slName = tgRptSelDemoCode(ilDemoLoop).sKey
            ilRet = gParseItem(slName, 2, "\", slCode)
            ilDemoList(UBound(ilDemoList)) = Val(slCode)                     'mnf code to Demo name
            ReDim Preserve ilDemoList(0 To UBound(ilDemoList) + 1)
        End If
    Next ilDemoLoop
    'Process 1 vehicle at a time - get cpp/cpm for all dayparts and all demos rquested.
    'Then build 1 record for each set of 13 demos by daypart and vehicle.
    For ilVeh = 1 To UBound(ilVehicles) Step 1
        ilSaveVeh = ilVehicles(ilVeh)
        For ilDemoLoop = 0 To UBound(ilDemoList) - 1 Step 1
        ilmnfDemo = ilDemoList(ilDemoLoop)
        'get population for each demo
        ilRet = gGetDemoPop(hmDrf, hmMnf, ilDnfCode, ilMnfSocEco, ilmnfDemo, llPop)
            For ilRif = LBound(tmMRif) To UBound(tmMRif) - 1 Step 1
                If (tmMRif(ilRif).iVefCode = ilSaveVeh) And (tmMRif(ilRif).iRcfCode = ilRCCode) Then
                    For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                    If tmMRif(ilRif).iRdfCode = tgMRdf(ilRdf).iCode Then
                        ilFound = False
                        For ilTest = LBound(tmBookGen) To UBound(tmBookGen) - 1 Step 1
                            If (tmBookGen(ilTest).iRdfCode = tmMRif(ilRif).iRdfCode) And (tmBookGen(ilTest).iVefCode = ilSaveVeh) And (tmBookGen(ilTest).iMnfDemo = ilmnfDemo) Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilTest
                        If Not ilFound Then
                            'Build record into tmBookGen
                            tmBookGen(UBound(tmBookGen)).lPop = llPop
                            tmBookGen(UBound(tmBookGen)).iMnfDemo = ilmnfDemo
                            tmBookGen(UBound(tmBookGen)).iRdfCode = tmMRif(ilRif).iRdfCode
                            tmBookGen(UBound(tmBookGen)).iVefCode = ilSaveVeh
                            'Get price
                            llTPrice = 0
                            ilNoWks = 0
                            For llDate = llEffDate To llEffEndDate Step 7
                                slDate = Format$(llDate, "m/d/yy")
                                ilRCWkNo = (llEffDate - llRCStartDate) \ 7 + 1
                                If ilRCWkNo = 1 Then
                                    llWkPrice = tmMRif(ilRif).lRate(0) + tmMRif(ilRif).lRate(1)
                                ElseIf ilRCWkNo = 52 Then
                                    llWkPrice = tmMRif(ilRif).lRate(52) + tmMRif(ilRif).lRate(53)
                                Else
                                    llWkPrice = tmMRif(ilRif).lRate(ilRCWkNo)
                                End If
                                llTPrice = llTPrice + llWkPrice         'no pennies
                                ilNoWks = ilNoWks + 1
                            Next llDate
                            If ilNoWks > 0 Then
                                tmBookGen(UBound(tmBookGen)).lAvgPrice = llTPrice / ilNoWks
                            End If
    
                            'generate valid days this daypart is airing (pass to Audience routines)
                            For ilLoop = 1 To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                                If (tgMRdf(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                                    For ilDay = 0 To 6 Step 1
                                        If (tgMRdf(ilRdf).sWkDays(ilLoop, ilDay + 1) = "Y") Then
                                            ilValidDays(ilDay) = True
                                        End If
                                    Next ilDay
                                End If
                            Next ilLoop
                            
                            If (ilDnfCode > 0) Then             'book exists
                                ilRet = gGetDemoAvgAud(hmDrf, hmMnf, ilDnfCode, ilSaveVeh, ilMnfSocEco, ilmnfDemo, tmMRif(ilRif).iRdfCode, llOvStartTime, llOvEndtime, ilValidDays(), llAvgAud)
                                tmBookGen(UBound(tmBookGen)).lAvgAud = llAvgAud
                                'Get Rating, avg audience , cpp, cpm
                                ilWkSpotCount(1) = 1
                                llWkActPrice(1) = tmBookGen(UBound(tmBookGen)).lAvgPrice
                                llWkAvgAud(1) = llAvgAud
                                'gAvgAudToLnResearch llPop, ilWkSpotCount(), llWkActPrice(), llWkAvgAud(), llLnCost, llAvgAudAvg, ilWkRating(), ilLnAvgRating, llWkGrImp(), llLnGrImp, llWkGRP(), llLnGRP, llCPP, llCPM
                                'tmBookGen(UBound(tmBookGen)).iAvgRating = ilLnAvgRating
                                'tmBookGen(UBound(tmBookGen)).lCPP = llCPP
                                'tmBookGen(UBound(tmBookGen)).lCPM = llCPM
                                If (tmBookGen(UBound(tmBookGen)).lAvgAud > 0) And (tmBookGen(UBound(tmBookGen)).lAvgPrice > 0) Then
                                    ilFoundAgain = False
                                    For ilLoop = 1 To UBound(ilDPList) - 1 Step 1
                                        If ilDPList(ilLoop) = tmBookGen(UBound(tmBookGen)).iRdfCode Then
                                            ilFoundAgain = True
                                            Exit For
                                        End If
                                    Next ilLoop
                                    If Not ilFoundAgain Then
                                        ilDPList(UBound(ilDPList)) = tmBookGen(UBound(tmBookGen)).iRdfCode
                                        ReDim Preserve ilDPList(1 To UBound(ilDPList) + 1) As Integer
                                    End If
                                    ReDim Preserve tmBookGen(0 To UBound(tmBookGen) + 1) As BOOKGEN
                                End If
                            End If
                        End If
                    End If
                    Next ilRdf                  'next DP
                End If
            Next ilRif                          'next DP rate
        Next ilDemoLoop
        'Vehicle is complete - Create ANR file from tmBookGen array
        'Outer loop - DAypart table - loop thru the unique dayparts for the vehicle and find  all the demos for that DP.
        'Create 1 record/daypart/vehicle for a max of 13 demos per record.
        'Create array of ANR records all built in memory for all demos for this 1 daypart before writing to disk.
        For ilRdf = 1 To UBound(ilDPList) - 1 Step 1
            For ilBookInx = 0 To UBound(tmBookGen) - 1 Step 1           'loop thru all the demos for each daypart for this vehicle to create 1 record for a set of 13 demos
                If tmBookGen(ilBookInx).iRdfCode = ilDPList(ilRdf) Then
                    For ilLoop = 0 To UBound(ilDemoList) - 1 Step 1
                        If ilDemoList(ilLoop) = tmBookGen(ilBookInx).iMnfDemo Then
                            Exit For                    'obtain the index to this demo in demo array.  The demo from the current tmBookGen
                                                        'record must be the the same index across all dayparts
                        End If
                    Next ilLoop
                    If tmBookGen(ilBookInx).iMnfDemo = ilDemoList(ilLoop) Then
                        ilDay = ilLoop \ 13 + 1            'relative ANR record(+1), may need multiples images if more than 13 demos
                        If ilDay > UBound(tmTAnr) Then  'need to allocate the record
                            tmAnr = tmZeroAnr
                            tmTAnr(ilDay - 1).iGenDate(0) = igNowDate(0)
                            tmTAnr(ilDay - 1).iGenDate(1) = igNowDate(1)
                            tmTAnr(ilDay - 1).iGenTime(0) = igNowTime(0)
                            tmTAnr(ilDay - 1).iGenTime(1) = igNowTime(1)
                            tmTAnr(ilDay - 1).iEffectiveDate(0) = ilEffDate(0)
                            tmTAnr(ilDay - 1).iEffectiveDate(1) = ilEffDate(1)
                            tmTAnr(ilDay - 1).iVefCode = tmBookGen(ilBookInx).iVefCode
                            tmTAnr(ilDay - 1).iRdfCode = tmBookGen(ilBookInx).iRdfCode
                            tmTAnr(ilDay - 1).lRCPrice(1) = tmBookGen(ilBookInx).lAvgPrice
                            ReDim Preserve tmTAnr(0 To UBound(tmTAnr) + 1) As ANR
                        End If
                        ilFound = (ilLoop Mod 13) + 1    'get remainder to determine which index this demo (1-13) will be placed in
                        tmTAnr(ilDay - 1).lUpfPrice(ilFound) = tmBookGen(ilBookInx).lAvgAud
                        tmTAnr(ilDay - 1).lMinPrice(ilFound) = tmBookGen(ilBookInx).lCPP
                        tmTAnr(ilDay - 1).lMaxPrice(ilFound) = tmBookGen(ilBookInx).lCPM
                        tmTAnr(ilDay - 1).lScatPrice(ilFound) = tmBookGen(ilBookInx).lPop
                        'ilLoop = the mnfDemo to be used
                        ilTest = (ilLoop \ 13)       'find the set (of 13) within all demos selected
                        tmTAnr(ilDay - 1).iPctSellout(ilFound) = ilDemoList((ilTest * 13) + ilFound - 1)
                        tmTAnr(ilDay - 1).iPctSellout(1) = ilDemoList(ilTest * 13)      'always place the 1st group of 13 demos in the first demo name -
                    End If
                End If
            Next ilBookInx
            For ilLoop = 0 To UBound(tmTAnr) - 1 Step 1
                ilRet = btrInsert(hmAnr, tmTAnr(ilLoop), imAnrRecLen, INDEXKEY0)
            Next ilLoop
            ReDim Preserve tmTAnr(0 To 0) As ANR
        Next ilRdf                          'creat next 13 demos for DP
        ReDim tmBookGen(0 To 0) As BOOKGEN   'initialize for the next demo
        ReDim Preserve ilDPList(1 To 1) As Integer      'initialize for next demo
    Next ilVeh                            'next vehicle for next daypart
    Erase tmBookGen, tmMRif, ilVehicles, ilDPList
    ilRet = btrClose(hmAnr)
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmRif)
    ilRet = btrClose(hmRcf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmVef)
End Sub
'**********************************************************************
'
'            mBuildCommProj
'            Loop thru contracts within date last std bdcst billing
'            through the end of the year and build up to 12 periods
'
'**********************************************************************
Function mBuildCommProj() As Integer
Dim ilRet As Integer
Dim slStr As String
Dim slMonth As String
Dim slDay As String
Dim slYear As String
Dim slStdStart As String            'start date to gather (std start)
Dim slStdEnd As String              'end date to gather (end of std year)
Dim slCntrStatus As String          'list of contract status to gather (working, order, hold, etc)
Dim slCntrType As String            'list of contract types to gather (Per inq, Direct Response, Remnants, etc)
Dim ilHOState As Integer            'which type of HO cntr states to include (whether revisions should be included)
Dim llContrCode As Long
Dim ilCurrentRecd As Integer
Dim ilFoundCnt As Integer
Dim ilLoop As Integer
Dim ilClf As Integer                'loop count for lines
Dim ilCff As Integer                'loop count for flights
Dim slNameCode As String
Dim slCode As String
Dim ilTemp As Integer
Dim llStdStart As Long              'requested start date to gather (serial date)
Dim llStdEnd As Long                'requested end date to gather (serial date)
Dim llFirstStdProjEnd               'std end date of 1st month of projection
Dim llFltStart As Long              'flight date's serial start date, altered for each flight
Dim llFltEnd As Long                'flight date's serial end date
Dim llFltWeek As Long               'flight end week
Dim llDate As Long
Dim llDate2 As Long
Dim llSpots As Long              'Total spots / week
Dim ilMonthInx As Integer           'index into 12 month buckets
Dim llTempStdStart As Long
Dim llTempStdEnd As Long
Dim ilFoundMonth As Integer
Dim slAmount As String
Dim slDollar As String
Dim slSharePct As String            'slsp share of contract xx.xxxx
Dim slSlsPct As String              'slsp comm % xx.xxxx
Dim ilCorT As Integer
Dim ilStartCorT As Integer
Dim ilEndCorT As Integer
Dim slCashAgyComm As String
Dim slTradeAgyComm As String
Dim slPctTrade As String
Dim slNet As String
Dim ilAdjust As Integer             'adjustment factor - index into accumulated buckets for future
Dim llDollar As Long
Dim ilFoundSlsp As Integer
ReDim lmProjMonths(1 To 13) As Long    'jan-dec + total all months
Dim slGrossOrNet As String             'commissions calc from G = Gross or N = net
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        mBuildCommProj = -1
        Exit Function
    End If
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        btrDestroy hmCff
        mBuildCommProj = -1
        Exit Function
    End If
    imCffRecLen = Len(tmCff)
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        btrDestroy hmAgf
        btrDestroy hmClf
        btrDestroy hmCff
        btrDestroy hmCff
        mBuildCommProj = -1
        Exit Function
    End If
    imAgfRecLen = Len(tmAgf)
    slGrossOrNet = "N"                  'force to calc comissions from net amount
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slStr       'convert last bdcst billing date to string
    gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
    ilLoop = Val(RptSelIv!edcSelCFrom.Text)            'year user entered
    If Val(slYear) = ilLoop Then
        If Val(slMonth) = 12 Then
            'no projections, all history for the year which has already been done
            mBuildCommProj = -1                 'no dates for future
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            Exit Function
        Else
            ilAdjust = Val(slMonth)             'adjustment factor (this is the offset to start
                                                'accumulating into future buckets
            slMonth = Str$(Val(slMonth) + 1)
            slStdStart = gObtainStartStd(slMonth & "/" & "15" & "/" & slYear)
            slStdEnd = gObtainEndStd("12/15/" & slYear)     'always do to end of year
            llStdStart = gDateValue(slStdStart)             'convert for comparision testing
            slStr = gObtainEndStd(slStdStart)
            llFirstStdProjEnd = gDateValue(slStr)
            llStdEnd = gDateValue(slStdEnd)
        End If
    ElseIf Val(slYear) > ilLoop Then            'actuals only, last date billed is later than year requested, Exit
        Exit Function
    Else                                        'projections only, last date billed is prior to year requested
        ilAdjust = 0                            'no adjustment factor
        slStr = "1/15/" & Trim$(Str$(ilLoop))    'start at beginning of user requested year
        slStdStart = gObtainStartStd(slStr)
        slStr = "12/15/" & Trim$(Str$(ilLoop))
        slStdEnd = gObtainEndStd(slStr)
        llStdStart = gDateValue(slStdStart)
        llStdEnd = gDateValue(slStdEnd)
        slStr = gObtainEndStd(slStdStart)
        llFirstStdProjEnd = gDateValue(slStr)         ' save the first std start month
    End If
    
    
    slCntrStatus = ""                   'all statuses: working order, hold, etc.
    slCntrType = "CVTRQ"                     'all types: PI, DR, etc.  except PSA(p) and Promo(m)
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)
    'build table (into tlchfadvtext) of all contracts that fall within the dates required
    ilRet = gObtainCntrForDate(RptSelIv, slStdStart, slStdEnd, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) Step 1                                            'loop while llCurrentRecd < llRecsRemaining
        'check for valid slsp selection
        If (RptSelIv!ckcAll) Then
            ilFoundCnt = True
        Else
            ilFoundCnt = False                             'assume nothing found until match in demo selection table
            For ilLoop = 0 To RptSelIv!lbcSelection(2).ListCount - 1 Step 1
                If RptSelIv!lbcSelection(2).Selected(ilLoop) Then
                    slNameCode = tgSalesperson(ilLoop).sKey
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                    For ilTemp = 0 To 9 Step 1
                        If Val(slCode) = tlChfAdvtExt(ilCurrentRecd).iSlfCode(ilTemp) Then
                            ilFoundCnt = True
                            Exit For
                        End If
                    Next ilTemp
                End If
            Next ilLoop
        End If
        If (ilFoundCnt) Then
            llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
            ilRet = gObtainCntr(hmChf, hmClf, hmCff, llContrCode, False, tgChf, tgClf(), tgCff())
            'obtain agency for commission
            If tgChf.iagfCode > 0 Then
                tmAgfSrchKey.iCode = tgChf.iagfCode
                ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching agy recd
                If ilRet = BTRV_ERR_NONE Then
                    slCashAgyComm = gIntToStrDec(tmAgf.iComm, 2)
                End If          'ilret = btrv_err_none
            Else
                slCashAgyComm = ".00"
            End If              'iagfcode > 0
   
            For ilLoop = 1 To 13 Step 1
                lmProjMonths(ilLoop) = 0
            Next ilLoop
            For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                tmClf = tgClf(ilClf).ClfRec
                ilCff = tgClf(ilClf).iFirstCff
                Do While ilCff <> -1
                    tmCff = tgCff(ilCff).CffRec
                    
                    gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
                    llFltStart = gDateValue(slStr)
                    'backup start date to Monday
                    ilLoop = gWeekDayLong(llFltStart)
                    Do While ilLoop <> 0
                        llFltStart = llFltStart - 1
                        ilLoop = gWeekDayLong(llFltStart)
                    Loop
                    gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
                    llFltEnd = gDateValue(slStr)
                    'the flight dates must be within the start and end of the projection periods,
                    'not be a CAncel before start flight, and have a cost > 0
                    If (llFltStart <= llStdEnd And llFltEnd >= llStdStart) And (llFltEnd >= llFltStart And tmCff.lActPrice > 0) Then
                        'only retrieve for projections, anything in the past has already
                        'been invoiced and has been retrieved from history or receiv files
                        'adjust the gather dates from flights: use flight start date or requested start date, whichever is later
                        If llStdStart > llFltStart Then
                            llFltStart = llStdStart
                        End If
                        'use flight end date or requsted end date, whichever is lesser
                        If llStdEnd < llFltEnd Then
                            llFltEnd = llStdEnd
                        End If
                        
                        For llDate = llFltStart To llFltEnd Step 7
                            'Loop on the number of weeks in this flight
                            'calc week into of this flight to accum the spot count
                            If tmCff.sDyWk = "W" Then            'weekly
                                llSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                            Else                                        'daily
                                If ilLoop + 6 < llFltEnd Then           'we have a whole week
                                    llSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)
                                Else
                                    llFltEnd = llDate + 6
                                    If llDate > llFltEnd Then
                                        llFltEnd = llFltEnd       'this flight isn't 7 days
                                    End If
                                    For llDate2 = llDate To llFltEnd Step 1
                                        ilTemp = gWeekDayLong(llDate2)
                                        llSpots = llSpots + tmCff.iDay(ilTemp)
                                    Next llDate2
                                End If
                            End If
                            'determine month that this week belongs in, then accumulate the gross and net $
                            'currently, the projections are based on STandard bdcst
                            llTempStdStart = llStdStart
                            llTempStdEnd = llFirstStdProjEnd
                            ilFoundMonth = False
                            For ilMonthInx = 0 To 11 Step 1
                                If llDate >= llTempStdStart And llDate <= llTempStdEnd Then
                                    ilFoundMonth = True
                                    Exit For
                                Else
                                    'get the next month's date
                                    llTempStdStart = llTempStdEnd + 1     'find next month
                                    slStr = Format$(llTempStdStart, "m/d/yy")
                                    slStr = gObtainEndStd(slStr)
                                    llTempStdEnd = gDateValue(slStr)
                                    If llTempStdStart > llStdEnd Then     'past end of requested date
                                        Exit For
                                    End If
                                End If
                            Next ilMonthInx
                            If ilFoundMonth Then
                                ilMonthInx = ilMonthInx + ilAdjust + 1             'adjust the month inx for the array of buckets
                                                                                'iladjust = # of months in the past,
                                                                                '+1 is the adjustment due to loop starting
                                                                                'at zero
                                lmProjMonths(ilMonthInx) = lmProjMonths(ilMonthInx) + (llSpots * tmCff.lActPrice)
                            End If
                        Next llDate                                     'for llDate = llFltStart To llFltEnd
                    End If                                          '
                    ilCff = tgCff(ilCff).iNextCff                   'get next flight record from mem
                Loop                                            'while ilcff <> -1
            Next ilClf                                          'for ilclf = lbound(tmclf) to ubound(tmclf)
            slPctTrade = gIntToStrDec(tgChf.iPctTrade, 0)
            ilTemp = 0
            'insure that the first slsp is 100 if only one exists
            For ilLoop = 0 To 9 Step 1
                If tgChf.iSlfCode(ilLoop) > 0 Then
                    ilTemp = ilTemp + 1
                End If
            Next ilLoop
            If ilTemp = 1 Then                      'only 1 slsp, force to 100% (no splits)
                tgChf.lComm(0) = 1000000             'force 100.0000%
            End If
            For ilLoop = 0 To 9 Step 1
                
                If tgChf.lComm(ilLoop) > 0 Then
                    
                    For ilTemp = 1 To 13 Step 1             'init the years $ buckets for the contract
                        tmGrf.lDollars(ilTemp) = 0
                    Next ilTemp
                  'if split slsp comm exist and selective slsp requested, don't show all slsp on this order
                    ilFoundSlsp = False                             'assume nothing found until match in demo selection table
                    For ilCorT = 0 To RptSelIv!lbcSelection(2).ListCount - 1 Step 1
                        If RptSelIv!lbcSelection(2).Selected(ilCorT) Then
                            slNameCode = tgSalesperson(ilCorT).sKey
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)    'Get application name
                            
                            If Val(slCode) = tgChf.iSlfCode(ilLoop) Then
                                ilFoundSlsp = True
                                Exit For
                            End If
                        End If
                        If ilFoundSlsp Then
                            Exit For
                        End If
                    Next ilCorT
        
                    tmSlfSrchKey.iCode = tgChf.iSlfCode(ilLoop)    'find slsp record for comm
                    ilRet = btrGetEqual(hmSlf, tmSlf, imSlfReclen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching slsp recd
                    If ilRet = BTRV_ERR_NONE And ilFoundSlsp Then
                        If tgChf.iPctTrade = 0 Then
                            ilStartCorT = 1
                            ilEndCorT = 1
                        ElseIf tgChf.iPctTrade = 100 Then
                            ilStartCorT = 2
                            ilEndCorT = 2
                        Else
                            ilStartCorT = 1
                            ilEndCorT = 2
                        End If
                        For ilCorT = ilStartCorT To ilEndCorT Step 1
                            For ilTemp = 1 To 12 Step 1
                                slAmount = gLongToStrDec(lmProjMonths(ilTemp), 2)
                                slSlsPct = gLongToStrDec(tmSlf.lUnderComm, 4)              'convert slsp comm % to packed dec.
                                slSharePct = gLongToStrDec(tgChf.lComm(ilLoop), 4)           'slsp split share in %
                                slStr = gMulStr(slSharePct, slAmount)                       'slsp gross portion of possible split
                                slStr = gRoundStr(slStr, "01.", 0)
                                If ilCorT = 1 Then                 'all cash commissionable
                                    If tgChf.iagfCode > 0 Then
                                        slCode = gSubStr("100.", slPctTrade)
                                        slDollar = gDivStr(gMulStr(slStr, slCode), "100")              'slsp gross
                                        slDollar = gRoundStr(slDollar, "01.", 0)
                                        tmGrf.lDollars(13) = tmGrf.lDollars(13) + Val(slDollar)
                                        slNet = gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "10000.00")
                                    End If
                                    tmGrf.sDateType = "C"       'cash flag for sorting
                                Else
                                    If ilCorT = 2 Then                'at least cash is commissionable
                                        slCode = gIntToStrDec(tgChf.iPctTrade, 0)
                                        slDollar = gDivStr(gMulStr(slStr, slCode), "100")
                                        slDollar = gRoundStr(slDollar, "01.", 0)
                                        tmGrf.lDollars(13) = tmGrf.lDollars(13) + Val(slDollar)
                                        slNet = slDollar                'assume no commissions on trade
                                        If tgChf.iagfCode > 0 And tgChf.sAgyCTrade = "Y" Then
                                            slNet = gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "10000.00")
                                        End If
                                        tmGrf.sDateType = "T"   'trade flag for sorting
                                    End If
                                End If
                                'commission take from gross for now, use slNet when comm needs to be obtained from Net value
                                If slGrossOrNet = "G" Then
                                    slAmount = gMulStr(slDollar, slSlsPct)         'calc slsp comm.
                                Else
                                    slAmount = gMulStr(slNet, slSlsPct)         'calc slsp comm.
                                End If
                                slAmount = gRoundStr(slAmount, "01.", 0)
                                tmGrf.lDollars(ilTemp) = tmGrf.lDollars(ilTemp) + Val(slAmount)
                            Next ilTemp                         'process next month
                            'contract complete for cash and or trade values, write out contract
                            tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for retrieval/removal of records
                            tmGrf.iGenDate(1) = igNowDate(1)
                            tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for retrieval/removal of records
                            tmGrf.iGenTime(1) = igNowTime(1)
                            tmGrf.sBktType = "C"                    'flag this as a cntr  recd (vs history)
                            tmGrf.lChfCode = tgChf.lCode            'contr internal code
                            tmGrf.iSlfCode = tgChf.iSlfCode(ilLoop)     'slsp code
                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                        Next ilCorT                             'process cash or trade portion
                    End If                          'btrv_err_none and ilFoundslsp
                End If                              'lcomm(ilLoop) > 0
            Next ilLoop                             'loop for 10 slsp possible splits
        End If                                                  'foundcnt
    Next ilCurrentRecd
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
End Function
'
'
'                   mCumeInsert - build GRF record for
'                   cumulative Activity Report.  Table of vehicles
'                   built in memory containing vehicle code and
'                   12 monthly buckets.  this routine loops through
'                   the table and creates 1 record per vehicle to disk.
'                   Zero $ records are not created.  Gross and Net $ are
'                   calculated based on user input.  Also, if contract
'                   is part cash, part trade, each % is taken in account
'                   to see if it is commissionable when net requested.
'                   <input>  ilUpperVef - # of vehicles to create records for
'                            tmVefDollars - table in memory of the vehicles gathred
'                                   for the contract (all gross $)
'                            slGrossOrNet - G = gross, N = net
'                   <output> tmGrf - GRF record created to disk
'
'                   6/9/97 d.Hosaka
Sub mCumeInsert(ilUpperVef As Integer, tmVefDollars() As ADJUSTLIST, tmGrf As GRF, slGrossOrNet)
Dim ilTemp As Integer               'Loop variable for number of vehicles to process
Dim llGross As Long                 'total of all 12 months each vehicle, if zero record not written
Dim ilPct As Integer                '% of cash or trade portion
Dim ilAgyComm As Integer            '% due station (minus the 15% agy comm), either 100 or 85
Dim ilTemp2 As Integer              'Cash/trade loop variable
Dim ilLoop As Integer               'loop variable for 12 months
Dim llTemp As Long                  'temp long variable for math, rounding, etc.
Dim ilRet As Integer                'error return from btrieve
Dim llNoPenny As Long               'project $ without pennies
    For ilTemp = 0 To ilUpperVef - 1 Step 1
        tmGrf.iVefCode = tmVefDollars(ilTemp).iVefCode
        llGross = 0
        For ilLoop = 1 To 12
            For ilTemp2 = 1 To 2                        'loop all vehicles for cash & trade (one order splits cash & trade)
                If ilTemp2 = 1 Then                     'loop to calc cash $, then trade $
                    ilPct = 100 - tgChf.iPctTrade       'get cash portion or order
                    ilAgyComm = 100                     'Assume gross requested
                    If slGrossOrNet = "N" Then          '
                        If tgChf.iagfCode > 0 Then      'agency exists,  net- take out commission
                            ilAgyComm = 85
                        End If
                    End If
                Else                                'trade portion
                    ilPct = tgChf.iPctTrade         'trade portion of order
                    ilAgyComm = 100                 'assume no commissionable on trade
                    If tgChf.sAgyCTrade = "Y" Then  'trade portion is commissionable, is it Gross orNet requested
                        If slGrossOrNet = "N" Then
                            ilAgyComm = 85
                        End If
                    End If
                End If
                llNoPenny = tmVefDollars(ilTemp).lProject(ilLoop) / 100    'drop pennies
                llTemp = llNoPenny * ilPct / 100    'calc cash vs trade
                llTemp = llTemp * ilAgyComm / 100                               'calc agy comm
                tmGrf.lDollars(ilLoop) = tmGrf.lDollars(ilLoop) + llTemp
                llGross = llGross + tmGrf.lDollars(ilLoop)
            Next ilTemp2
        Next ilLoop
        If llGross <> 0 Then        'dont write out zero $ records
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
        For ilTemp2 = 1 To 12
            tmGrf.lDollars(ilTemp2) = 0
        Next ilTemp2
    Next ilTemp
End Sub
Sub mWriteSlsAct()
Dim ilRet As Integer
Dim ilTemp As Integer
        tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
        tmGrf.iGenDate(1) = igNowDate(1)
        tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
        tmGrf.iGenTime(1) = igNowTime(1)
        tmGrf.lChfCode = tgChf.lCntrNo            'contract #
        tmGrf.iAdfCode = tgChf.iAdfCode         'advertiser code
        tmGrf.iPerGenl(4) = 0                   'assume Order (vs hold)
        If tgChf.sStatus = "H" Or tgChf.sStatus = "G" Then      'if hold or unsch hold, set flag for Crystal
            tmGrf.iPerGenl(4) = 1
        End If
        If tgChf.iSlfCode(0) <> tmSlf.iCode Then        'only read slsp recd if not in mem already
            tmSlfSrchKey.iCode = tgChf.iSlfCode(0)         'find the slsp to obtain the sales source code
            ilRet = btrGetEqual(hmSlf, tmSlf, imSlfReclen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        End If                                          'table of selling offices built into memory with its
        For ilTemp = LBound(tlSofList) To UBound(tlSofList)
            If tlSofList(ilTemp).iSofCode = tmSlf.iSofCode Then
                tmGrf.iCode2 = tlSofList(ilTemp).iMnfSSCode          'Sales source
                Exit For
            End If
        Next ilTemp
End Sub
