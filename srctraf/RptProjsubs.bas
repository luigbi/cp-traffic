Attribute VB_Name = "RPTPROJSUBS"
' Copyright 1993 Counterpoint Software®. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: AcqCommSubs.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Population subs and functions for acquisition Commission
'   and Projection flight gathering
'
Option Explicit
Option Compare Text

Dim tmVbf As VBF
Dim hmVbf As Integer
Dim imVBfRecLen As Integer
Dim tmVbfSrchKey1 As VBFKEY1
Public tgAcqComm() As ACQCOMM
Public tgAcqCommInx() As ACQCOMMINX

Type ACQCOMM
        sKey As String * 15                     '5 char vehicle internal|effective date (5)
        iVefCode As Integer
        lEffStartDate As Long
        lEndDate As Long
        iAcqCommPct As Integer          'xxx.xx
End Type
'
Type ACQCOMMINX
        iVefCode As Integer
        iLoInx As Integer                       'index into sorted array, lo index starting point to vehicle info
        iHiInx As Integer                       'index into sorted array, hi index ending point to vehicle info
End Type
'
Type ACQPctByPeriod
    iVefCode As Integer
    'lStdStartDates(1 To 13) As Long
    'iAcqCommPct(1 To 12) As Integer
    lStdStartDates(0 To 13) As Long 'Index zero ignored
    iAcqCommPct(0 To 12) As Integer 'Index zero ignored
End Type

'       gAccumDaysFromCal - Accum each of the days calculated from the calendar months requested
'           and add into the monthly buckets.  The Calendar $ have been gathered by day in llCalAmt()
'       <input>  llStdStartDates() - array of calendar start dates for 1 year
'                llCalAmt()- array of days containing the $ distributed
'                llCalAcqAmt() - array of days containing acquisition amt
'                ilHowManyPer - # of months requested
'       <return> lmproject - calendar months $ projected from contracts
'       Created:  3/5/10
Public Sub gAccumCalFromDays(llStdStartDates() As Long, llCalAmt() As Long, llCalAcqAmt() As Long, ilUseAcquisitionCost As Integer, llProject() As Long, llAcquisition() As Long, ilHowManyPer As Integer)
    Dim ilLoop As Integer
    Dim ilMonthInx As Integer
    Dim llTempDate As Long
    Dim ilMonthTest As Integer

    llTempDate = llStdStartDates(1)
    For ilLoop = LBound(llCalAmt) To UBound(llCalAmt)
        ilMonthInx = -1
        'TTP 10895 - RAB Cal Contract: when run for a long date span, may not get all data for an air time contract
        'For ilMonthTest = 1 To 12
        For ilMonthTest = 1 To UBound(llStdStartDates) - 1
            If llTempDate >= llStdStartDates(ilMonthTest) And llTempDate < llStdStartDates(ilMonthTest + 1) Then
                ilMonthInx = ilMonthTest
                Exit For
            End If
        Next ilMonthTest
       
        If ilMonthInx > 0 And ilMonthInx <= ilHowManyPer Then       'does the month fall within requested period?
            If ilUseAcquisitionCost Then
                llProject(ilMonthInx) = llProject(ilMonthInx) + llCalAcqAmt(ilLoop)
            Else
                llProject(ilMonthInx) = llProject(ilMonthInx) + llCalAmt(ilLoop)
            End If
            llAcquisition(ilMonthInx) = llAcquisition(ilMonthInx) + llCalAcqAmt(ilLoop)
        ElseIf ilMonthInx > ilHowManyPer Then        'exceed dates to gather, exit
            Exit For
        End If
        llTempDate = llTempDate + 1
    Next ilLoop
    Exit Sub
End Sub

'       gAccumDaysFromCal - Accum each of the days calculated from the calendar months requested
'           and add into the monthly buckets.  The Calendar $ have been gathered by day in llCalAmt()
'       <input>  llStdStartDates() - array of calendar start dates for x Number of months
'                llCalAmt()- array of days containing the $ distributed
'                llCalAcqAmt() - array of days containing acquisition amt
'                llCalAcqNetAmt() - array of days containing acquisition net amt if varying acq comm applied
'                ilAdjustAcquisition - true if tnet option
'                ilUseAcquisitionCost = true if showing Acq cost vs Spot count
'                ilHowManyPer - # of months requested
'       <return> lmproject - calendar months $ projected from contracts (from spots costs or if using acquisition cost its the acq costs)
'                llAcquisition() array of projected acq costs to be subtracted if tnet
'       Created:  3/5/10
'TTP 10665 - RAB Cal Contract: digital/ad server contract not appearing when lines are for Jan and the start month is Jan
'Public Sub gAccumCalFromDaysWithAcqNet(llStdStartDates() As Long, llCalAmt() As Long, llCalAcqAmt() As Long, llCalAcqNetAmt() As Long, ilAdjustAcquisition As Integer, ilUseAcquisitionCost As Integer, slGrossOrNet As String, llProject() As Long, llAcquisition() As Long, ilHowManyPer As Integer)
Public Sub gAccumCalFromDaysWithAcqNet(llStdStartDates() As Long, llCalAmt() As Long, llCalAcqAmt() As Long, llCalAcqNetAmt() As Long, ilAdjustAcquisition As Integer, ilUseAcquisitionCost As Integer, slGrossOrNet As String, llProject() As Long, llAcquisition() As Long, ilHowManyPer As Integer, ilFirstProjIdx As Integer)
    Dim ilLoop As Integer
    Dim ilMonthInx As Integer
    Dim llTempDate As Long
    Dim ilMonthTest As Integer
    'llProject - gross actual spot cost or gross acquisition costs
    'llAcquisition - gross acquisition costs
    'llAcquisitionNet - net acquisition costs if varying acq commissions; otherwise 0
    
    'TTP 10665
    'llTempDate = llStdStartDates(1)
    llTempDate = llStdStartDates(ilFirstProjIdx)
    For ilLoop = LBound(llCalAmt) To UBound(llCalAmt)
        ilMonthInx = -1
        'TTP 10895 - RAB Cal Contract: when run for a long date span, may not get all data for an air time contract
        'For ilMonthTest = 1 To 12
        For ilMonthTest = 1 To UBound(llStdStartDates) - 1
            If llTempDate >= llStdStartDates(ilMonthTest) And llTempDate < llStdStartDates(ilMonthTest + 1) Then
                ilMonthInx = ilMonthTest
                Exit For
            End If
        Next ilMonthTest

        If ilMonthInx > 0 And ilMonthInx <= ilHowManyPer Then       'does the month fall within requested period?
            If ilAdjustAcquisition Then
                If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then
                    If slGrossOrNet <> "G" Then   'g=gross, n = net, t = tnet
                        llAcquisition(ilMonthInx) = llAcquisition(ilMonthInx) + llCalAcqNetAmt(ilLoop)      'adjusting acq costs, use the net amt calc from acq comm.
                    End If
                Else
                    llAcquisition(ilMonthInx) = llAcquisition(ilMonthInx) + llCalAcqAmt(ilLoop)             'no acq comm; use the acq amt entered on line
                End If
            Else
                llAcquisition(ilMonthInx) = llAcquisition(ilMonthInx) + llCalAcqAmt(ilLoop)                 'no acq adjustments
            End If
            If ilUseAcquisitionCost Then
                If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE And slGrossOrNet <> "G" Then
                    llProject(ilMonthInx) = llProject(ilMonthInx) + llCalAcqNetAmt(ilLoop)
                Else
                    llProject(ilMonthInx) = llProject(ilMonthInx) + llCalAcqAmt(ilLoop)
                End If
            Else
                llProject(ilMonthInx) = llProject(ilMonthInx) + llCalAmt(ilLoop)
'If llCalAmt(ilLoop) <> 0 Then
'    Debug.Print " gAccumCalFromDaysAcqNet, MonthInx:" & ilMonthInx & ", Amt:" & llCalAmt(ilLoop) & ", Day:" & Format(llTempDate, "ddddd")
'End If
            End If
            'llAcquisition(ilMonthInx) = llAcquisition(ilMonthInx) + llCalAcqAmt(ilLoop)     'need for non acq comm
        ElseIf ilMonthInx > ilHowManyPer Then        'exceed dates to gather, exit
            Exit For
        End If
        llTempDate = llTempDate + 1
    Next ilLoop
    Exit Sub
End Sub

'       gAccumCalSpotsFromDays - Accum each of the days calculated from the calendar months requested
'           and add into the monthly buckets.  The Calendar spot counts have been gathered by day in llCalAmt()
'       <input>  llStdStartDates() - array of calendar start dates for 1 year
'                llCalSpots()- array of days containing the spots distributed
'                ilHowManyPer - # of months requested
'       <return> llproject - calendar months spot count projected from contracts
Public Sub gAccumCalSpotsFromDays(llStdStartDates() As Long, llCalSpots() As Long, llProject() As Long, ilHowManyPer As Integer)
    Dim ilLoop As Integer
    Dim ilMonthInx As Integer
    Dim llTempDate As Long
    Dim ilMonthTest As Integer

    llTempDate = llStdStartDates(1)
    For ilLoop = LBound(llCalSpots) To UBound(llCalSpots)
        ilMonthInx = -1
        'TTP 10895 - RAB Cal Contract: when run for a long date span, may not get all data for an air time contract
        'For ilMonthTest = 1 To 12
        For ilMonthTest = 1 To UBound(llStdStartDates) - 1
            If llTempDate >= llStdStartDates(ilMonthTest) And llTempDate < llStdStartDates(ilMonthTest + 1) Then
                ilMonthInx = ilMonthTest
                Exit For
            End If
        Next ilMonthTest
       
        If ilMonthInx > 0 And ilMonthInx <= ilHowManyPer Then       'does the month fall within requested period?
            llProject(ilMonthInx) = llProject(ilMonthInx) + llCalSpots(ilLoop)
        ElseIf ilMonthInx > ilHowManyPer Then        'exceed dates to gather, exit
            Exit For
        End If
        llTempDate = llTempDate + 1
    Next ilLoop
    Exit Sub
End Sub

'           gAccumSpotsbyFlight - Determine what flight a spot belongs in and gets its $.
'           Accumulate $ into monthly bucket
'           <input>
'                   llStdStartDates() - array of  month/week start dates
'                   ilFirstProjInx - first month to project
'                   ilMaxInx - # periods to retrieve
'                   llAcquisition() - array of the acqusition costs for period if net net (need both llproject & acq costs to adjust)
'                   ilWkOrMonth 0 = monthly buckets , 1 = weekly buckets
'                   tlClfIP - schedule line to process
'                   tlCffIp - array of flights for schedule line
'                   tlSdfInfo - array of spots found for schedule line (sorted by sch line, date)
'                   llLoInx - starting inx into tlSdfInfo for sched line to process
'                   llHiInx - ending inx into tlSdfInfo for sched line to process
'           <output>
'                   llProject() - array of spot costs accumulated for period
'                   llAcquisitionNet() - array of acq costs net down if varying commissions applicable
'           1-24-11 when makegoods are scheduled beyond the contract expiration date, the correct
'                   $ were not found.  Need to use the orig missed date to get the flight rate,
'                   then use the sched mg date to project the $
'
Public Sub gAccumSpotsbyFlight(llStdStartDates() As Long, ilFirstProjInx As Integer, ilMaxInx As Integer, llProject() As Long, llAcquisition() As Long, llAcquisitionNet() As Long, ilWkOrMonth As Integer, tlClfInp As CLFLIST, tlCffInp() As CFFLIST, tlSdfInfo() As SDFSORTBYLINE, llLoInx As Long, llHiInx As Long, hlSmf As Integer, Optional ilUseAcquisitionCost As Integer = False)
    '******************************************************************************************e
    '* Note: VBC id'd the following unreferenced items and handled them as described:         *
    '*                                                                                        *
    '* Local Variables (Removed)                                                              *
    '*  llDate2                       llSpots                       ilTemp                    *
    '*                                                                                        *
    '******************************************************************************************
    
    Dim ilCff As Integer
    Dim slStr As String
    
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim llStdStart As Long
    Dim llStdEnd As Long
    Dim ilMonthInx As Integer
    Dim ilWkInx As Integer
    Dim tlCff As CFF
    Dim llInx As Long
    Dim tlSrchSmfKey2 As LONGKEY0
    Dim tlSmf As SMF
    Dim ilRet As Integer
    Dim llOrigDate As Long
    Dim ilFoundFlight As Integer
    Dim ilAcqLoInx As Integer
    Dim ilAcqHiInx As Integer
    Dim blAcqOK As Boolean
    Dim llAcqNet As Long
    Dim llAcqComm As Long
    Dim ilAcqCommPct As Integer

    llStdStart = llStdStartDates(ilFirstProjInx)
    llStdEnd = llStdStartDates(ilMaxInx)
    ilCff = tlClfInp.iFirstCff
    llInx = llLoInx                'set starting index into tlSdfInfo array of spots
    
    For llInx = llLoInx To llHiInx
        gUnpackDateLong tlSdfInfo(llInx).tSdf.iDate(0), tlSdfInfo(llInx).tSdf.iDate(1), llDate
        llOrigDate = llDate
        If (tlSdfInfo(llInx).tSdf.sSchStatus = "G" Or tlSdfInfo(llInx).tSdf.sSchStatus = "O") And tlSdfInfo(llInx).tSdf.sSpotType = "A" Then
            tlSrchSmfKey2.lCode = tlSdfInfo(llInx).tSdf.lCode       'sdf code to find the orig missed date for rate
            ilRet = btrGetEqual(hlSmf, tlSmf, Len(tlSmf), tlSrchSmfKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get current record
            If ilRet <> BTRV_ERR_NONE Then
                Exit Sub
            Else
                gUnpackDateLong tlSmf.iMissedDate(0), tlSmf.iMissedDate(1), llOrigDate
            End If
        End If
        ilCff = tlClfInp.iFirstCff
        Do While ilCff <> -1
            tlCff = tlCffInp(ilCff).CffRec
    
            gUnpackDate tlCff.iStartDate(0), tlCff.iStartDate(1), slStr
            llFltStart = gDateValue(slStr)
    
            gUnpackDate tlCff.iEndDate(0), tlCff.iEndDate(1), slStr
            llFltEnd = gDateValue(slStr)
            
'Debug.Print " AccumSpotbyFlt, Code:" & tlCff.lCode & " ,Start:" & Format(llFltStart, "ddddd") & " ,End:" & Format(llFltEnd, "ddddd")

            'the flight dates must be within the start and end of the projection periods,
            'not be a CAncel before start flight, and have a cost > 0
            'If (llFltStart < llStdEnd And llFltEnd >= llStdStart) And (llFltEnd >= llFltStart) Then
                'backup start date to Monday
            If (llFltStart <= llFltEnd) Then        'ok, not cancel before start
                ilLoop = gWeekDayLong(llFltStart)
                Do While ilLoop <> 0
                    llFltStart = llFltStart - 1
                    ilLoop = gWeekDayLong(llFltStart)
                Loop
                
                'only retrieve $ for projected periods.  If makegood, need to do all flights to get the correct $
                'if scheduled outside its flight date.  need to use the orig missed date
                If (tlSdfInfo(llInx).tSdf.sSchStatus = "G" Or tlSdfInfo(llInx).tSdf.sSchStatus = "O") And tlSdfInfo(llInx).tSdf.sSpotType = "A" Then
                '3-5-15 need to get the correct flight for rate retrieval for the outside or mg, which flight could be outside the report parameters, but sch spot is within parameters
'                    If llStdStart > llFltStart Then
'                        llFltStart = llStdStart
'                    End If
'                    'use flight end date or requsted end date, whichever is lesser
'                    If llStdEnd < llFltEnd Then
'                        llFltEnd = llStdEnd
'                    End If
                End If
                ilFoundFlight = False
                If llOrigDate >= llFltStart And llOrigDate <= llFltEnd And tlCff.iClfLine = tlSdfInfo(llInx).tSdf.iLineNo Then
                   If ilWkOrMonth = 1 Then                     'monthly buckets
                        'determine month that this week belongs in, then accumulate the gross and net $
                        'currently, the projections are based on STandard bdcst
                        For ilMonthInx = ilFirstProjInx To ilMaxInx - 1 Step 1       'loop thru months to find the match
                            If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                                If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then
                                    ilAcqCommPct = 0
                                    blAcqOK = gGetAcqCommInfoByVehicle(tlClfInp.ClfRec.iVefCode, ilAcqLoInx, ilAcqHiInx)
                                    ilAcqCommPct = gGetEffectiveAcqComm(llDate, ilAcqLoInx, ilAcqHiInx)
                                    gCalcAcqComm ilAcqCommPct, tlClfInp.ClfRec.lAcquisitionCost, llAcqNet, llAcqComm
                                    llAcquisitionNet(ilMonthInx) = llAcquisitionNet(ilMonthInx) + llAcqNet          'acq net, may be needed to tnet
                                    'lmAcquisition gets set with the gross acquisition amts
                                End If
                               
                                 If ilUseAcquisitionCost Then
                                     llProject(ilMonthInx) = llProject(ilMonthInx) + (tlClfInp.ClfRec.lAcquisitionCost)     'gross acq cost
                                 Else
                                     llProject(ilMonthInx) = llProject(ilMonthInx) + (tlCff.lActPrice)         'gross actual spot price
                                 End If
                            
                                llAcquisition(ilMonthInx) = llAcquisition(ilMonthInx) + (tlClfInp.ClfRec.lAcquisitionCost)    'acq gross
                                ilFoundFlight = True
                                Exit For
                            End If
                        Next ilMonthInx
                    Else                                    'weekly buckets
                        ilWkInx = (llDate - llStdStartDates(1)) \ 7 + 1
                        ''4-3-07 make sure the data isnt gathered beyond the period requested
                        'If ilWkInx > 0 And llDate >= llStdStartDates(LBound(llStdStartDates)) And llDate < llStdStartDates(UBound(llStdStartDates)) Then
                        If ilWkInx > 0 And llDate >= llStdStartDates(1) And llDate < llStdStartDates(UBound(llStdStartDates)) Then
                            If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then
                                    ilAcqCommPct = 0
                                    blAcqOK = gGetAcqCommInfoByVehicle(tlClfInp.ClfRec.iVefCode, ilAcqLoInx, ilAcqHiInx)
                                    ilAcqCommPct = gGetEffectiveAcqComm(llDate, ilAcqLoInx, ilAcqHiInx)
                                    gCalcAcqComm ilAcqCommPct, tlClfInp.ClfRec.lAcquisitionCost, llAcqNet, llAcqComm
                                    llAcquisitionNet(ilWkInx) = llAcquisitionNet(ilWkInx) + llAcqNet          'acq net, may be needed to tnet
                                    'lmAcquisition gets set with the gross acquisition amts
                                End If
                               
                                 If ilUseAcquisitionCost Then
                                     llProject(ilWkInx) = llProject(ilWkInx) + (tlClfInp.ClfRec.lAcquisitionCost)     'gross acq cost
                                 Else
                                     llProject(ilWkInx) = llProject(ilWkInx) + (tlCff.lActPrice)         'gross actual spot price
                                 End If
                            
                                llAcquisition(ilWkInx) = llAcquisition(ilWkInx) + (tlClfInp.ClfRec.lAcquisitionCost)    'acq gross
                            ilFoundFlight = True
                        End If
                    End If
                End If                               ' while llDate >= llFltStart And llDate <= llFltEnd
                ilCff = tlCffInp(ilCff).iNextCff                   'get next flight record from mem
                If ilFoundFlight Then
                    'found the flight, process next spot
                    ilCff = -1          'end flight search
                End If
            Else
                ilCff = tlCffInp(ilCff).iNextCff                   'get next flight record from mem
            End If                                          '
            'ilCff = tlCffInp(ilCff).iNextCff                   'get next flight record from mem
        Loop
    Next llInx
    Exit Sub
End Sub

'                   gBuildFlights - Loop through the flights of the schedule line
'                                   and build the projections dollars into lmprojmonths array
'                   <input> ilclf = sched line index into tlClfInp
'                           llStdStartDates() - array of dates to build $ from flights
'                           ilFirstProjInx - index of 1st month/week to start projecting
'                           ilMaxInx - max # of buckets to loop thru
'                           ilWkOrMonth - 1 = Month, 2 = Week
'                   <output> llProject() = array of $ buckets corresponding to array of dates
'
'
'                   General routine to build flight $ into week, month, qtr buckets
'
Sub gBuildFlights(ilClf As Integer, llStdStartDates() As Long, ilFirstProjInx As Integer, ilMaxInx As Integer, llProject() As Long, ilWkOrMonth As Integer, tlClfInp() As CLFLIST, tlCffInp() As CFFLIST)
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
    Dim ilWkInx As Integer
    Dim tlCff As CFF
    llStdStart = llStdStartDates(ilFirstProjInx)
    llStdEnd = llStdStartDates(ilMaxInx)
    ilCff = tlClfInp(ilClf).iFirstCff
    Do While ilCff <> -1
    tlCff = tlCffInp(ilCff).CffRec
    gUnpackDate tlCff.iStartDate(0), tlCff.iStartDate(1), slStr
    llFltStart = gDateValue(slStr)
    'backup start date to Monday
    'ilLoop = gWeekDayLong(llFltStart)
    'Do While ilLoop <> 0
    '    llFltStart = llFltStart - 1
    '    ilLoop = gWeekDayLong(llFltStart)
    'Loop
    gUnpackDate tlCff.iEndDate(0), tlCff.iEndDate(1), slStr
    llFltEnd = gDateValue(slStr)
    'the flight dates must be within the start and end of the projection periods,
    'not be a CAncel before start flight, and have a cost > 0
    If (llFltStart < llStdEnd And llFltEnd >= llStdStart) And (llFltEnd >= llFltStart And tlCff.lActPrice > 0) Then
        'backup start date to Monday
        ilLoop = gWeekDayLong(llFltStart)
        Do While ilLoop <> 0
            llFltStart = llFltStart - 1
            ilLoop = gWeekDayLong(llFltStart)
        Loop
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
            If tlCff.sDyWk = "W" Then            'weekly
                llSpots = tlCff.iSpotsWk + tlCff.iXSpotsWk
            Else                                        'daily
                If ilLoop + 6 < llFltEnd Then           'we have a whole week
                    llSpots = tlCff.iDay(0) + tlCff.iDay(1) + tlCff.iDay(2) + tlCff.iDay(3) + tlCff.iDay(4) + tlCff.iDay(5) + tlCff.iDay(6)
                Else
                    llFltEnd = llDate + 6
                    If llDate > llFltEnd Then
                        llFltEnd = llFltEnd       'this flight isn't 7 days
                    End If
                    For llDate2 = llDate To llFltEnd Step 1
                        ilTemp = gWeekDayLong(llDate2)
                        llSpots = llSpots + tlCff.iDay(ilTemp)
                    Next llDate2
                End If
            End If
            If ilWkOrMonth = 1 Then                     'monthly buckets
                'determine month that this week belongs in, then accumulate the gross and net $
                'currently, the projections are based on STandard bdcst
                For ilMonthInx = ilFirstProjInx To ilMaxInx - 1 Step 1       'loop thru months to find the match
                    If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                        llProject(ilMonthInx) = llProject(ilMonthInx) + (llSpots * tlCff.lActPrice)
                        Exit For
                    End If
                Next ilMonthInx
            Else                                    'weekly buckets
                ilWkInx = (llDate - llStdStartDates(1)) \ 7 + 1
                ''4-3-07 make sure the data isnt gathered beyond the period requested
                'If ilWkInx > 0 And llDate >= llStdStartDates(LBound(llStdStartDates)) And llDate < llStdStartDates(UBound(llStdStartDates)) Then
                If ilWkInx > 0 And llDate >= llStdStartDates(1) And llDate < llStdStartDates(UBound(llStdStartDates)) Then
                    llProject(ilWkInx) = llProject(ilWkInx) + (llSpots * tlCff.lActPrice)
                End If
            End If
        Next llDate                                     'for llDate = llFltStart To llFltEnd
    End If                                          '
    ilCff = tlCffInp(ilCff).iNextCff                   'get next flight record from mem
    Loop                                            'while ilcff <> -1
End Sub

'                   gBuildFlightSpots - Loop through the flights of the schedule line
'                                   and build the flight spot counts into llproject array
'                                   This is a copy of guildFlighs except it builds spot counts, not $
'                   <input> ilclf = sched line index into tlClfInp
'                           llStdStartDates() - array of dates to build spot count from flights
'                           ilFirstProjInx - index of 1st month/week to start projecting
'                           ilMaxInx - max # of buckets to loop thru
'                           ilWkOrMonth - 1 = Month, 2 = Week
'                   <output> llProject() = array of spot count buckets corresponding to array of dates
'
'
'                   General routine to build flight $ into week, month, qtr buckets
'
Sub gBuildFlightSpots(ilClf As Integer, llStdStartDates() As Long, ilFirstProjInx As Integer, ilMaxInx As Integer, llProject() As Long, ilWkOrMonth As Integer, tlClfInp() As CLFLIST, tlCffInp() As CFFLIST)
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
    Dim ilWkInx As Integer
    Dim tlCff As CFF
    
    llStdStart = llStdStartDates(ilFirstProjInx)
    llStdEnd = llStdStartDates(ilMaxInx)
    ilCff = tlClfInp(ilClf).iFirstCff
    Do While ilCff <> -1
    tlCff = tlCffInp(ilCff).CffRec
    gUnpackDate tlCff.iStartDate(0), tlCff.iStartDate(1), slStr
    llFltStart = gDateValue(slStr)
    'backup start date to Monday
    'ilLoop = gWeekDayLong(llFltStart)
    'Do While ilLoop <> 0
    '    llFltStart = llFltStart - 1
    '    ilLoop = gWeekDayLong(llFltStart)
    'Loop
    gUnpackDate tlCff.iEndDate(0), tlCff.iEndDate(1), slStr
    llFltEnd = gDateValue(slStr)
    'the flight dates must be within the start and end of the projection periods,
    'not be a CAncel before start flight, and have a cost > 0
    If (llFltStart < llStdEnd And llFltEnd >= llStdStart) And (llFltEnd >= llFltStart) Then
        'backup start date to Monday
        ilLoop = gWeekDayLong(llFltStart)
        Do While ilLoop <> 0
            llFltStart = llFltStart - 1
            ilLoop = gWeekDayLong(llFltStart)
        Loop
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
            If tlCff.sDyWk = "W" Then            'weekly
                llSpots = tlCff.iSpotsWk + tlCff.iXSpotsWk
            Else                                        'daily
                If ilLoop + 6 < llFltEnd Then           'we have a whole week
                    llSpots = tlCff.iDay(0) + tlCff.iDay(1) + tlCff.iDay(2) + tlCff.iDay(3) + tlCff.iDay(4) + tlCff.iDay(5) + tlCff.iDay(6)
                Else
                    llFltEnd = llDate + 6
                    If llDate > llFltEnd Then
                        llFltEnd = llFltEnd       'this flight isn't 7 days
                    End If
                    For llDate2 = llDate To llFltEnd Step 1
                        ilTemp = gWeekDayLong(llDate2)
                        llSpots = llSpots + tlCff.iDay(ilTemp)
                    Next llDate2
                End If
            End If
            If ilWkOrMonth = 1 Then                     'monthly buckets
                'determine month that this week belongs in, then accumulate the gross and net $
                'currently, the projections are based on STandard bdcst
                For ilMonthInx = ilFirstProjInx To ilMaxInx - 1 Step 1       'loop thru months to find the match
                    If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                        llProject(ilMonthInx) = llProject(ilMonthInx) + llSpots
                        Exit For
                    End If
                Next ilMonthInx
            Else                                    'weekly buckets
                ilWkInx = (llDate - llStdStartDates(1)) \ 7 + 1
                ''4-3-07 make sure the data isnt gathered beyond the period requested
                'If ilWkInx > 0 And llDate >= llStdStartDates(LBound(llStdStartDates)) And llDate < llStdStartDates(ilMaxInx) Then
                If ilWkInx > 0 And llDate >= llStdStartDates(1) And llDate < llStdStartDates(ilMaxInx) Then
                    llProject(ilWkInx) = llProject(ilWkInx) + llSpots
                End If
            End If
        Next llDate                                     'for llDate = llFltStart To llFltEnd
    End If                                          '
    ilCff = tlCffInp(ilCff).iNextCff                   'get next flight record from mem
    Loop                                            'while ilcff <> -1
End Sub

'                   gBuildFlightSpotsandRevenue - Loop through the flights of the schedule line
'                           and build the projections dollars into llproject array,
'                           and build projection # of spots into llprojectspots array
'                   <input> ilclf = sched line index into tlClfInp
'                           llStdStartDates() - array of dates to build $ from flights
'                           ilFirstProjInx - index of 1st month/week to start projecting
'                           ilMaxInx - max # of buckets to loop thru
'                           ilWkOrMonth - 1 = Month, 2 = Week
'                           ilUseWhichRate - 0 = use true line rate, 1 = use acquisition rate, 2 =use acq rate if non0, otherwise use linerate
'                           slGrossOrNet - G = Gross , N = Net (default to Net).  USed to acquisition costs computation if using Acq commissions
'                   <output> llProject() = array of $ buckets corresponding to array of dates
'                           llProjectSpots() array of spot count buckets corresponding to array of dates
'
'                   General routine to build flight $/cpot count into week, month, qtr buckets
'            Created : 7-12-05
'
Public Sub gBuildFlightSpotsAndRevenue(ilClf As Integer, llStdStartDates() As Long, ilFirstProjInx As Integer, ilMaxInx As Integer, llProject() As Long, llProjectSpots() As Long, ilWkOrMonth As Integer, ilUseWhichRate As Integer, tlClfInp() As CLFLIST, tlCffInp() As CFFLIST, Optional slGrossOrNet As String = "N")
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
    Dim ilWkInx As Integer
    Dim llWhichRate As Long
    Dim tlCff As CFF
    
    Dim ilAcqCommPct As Integer
    Dim blAcqOK As Boolean
    Dim ilAcqLoInx As Integer
    Dim ilAcqHiInx As Integer
    Dim llAcqNet As Long
    Dim llAcqComm As Long

    llStdStart = llStdStartDates(ilFirstProjInx)
    llStdEnd = llStdStartDates(ilMaxInx)
    ilCff = tlClfInp(ilClf).iFirstCff
    Do While ilCff <> -1
        tlCff = tlCffInp(ilCff).CffRec

        
        gUnpackDate tlCff.iStartDate(0), tlCff.iStartDate(1), slStr
        llFltStart = gDateValue(slStr)
        'backup start date to Monday
        'ilLoop = gWeekDayLong(llFltStart)
        'Do While ilLoop <> 0
        '    llFltStart = llFltStart - 1
        '    ilLoop = gWeekDayLong(llFltStart)
        'Loop
        gUnpackDate tlCff.iEndDate(0), tlCff.iEndDate(1), slStr
        llFltEnd = gDateValue(slStr)
        
        'the flight dates must be within the start and end of the projection periods,
        'not be a CAncel before start flight, and have a cost > 0
        If (llFltStart < llStdEnd And llFltEnd >= llStdStart) And (llFltEnd >= llFltStart) Then
            'backup start date to Monday
            ilLoop = gWeekDayLong(llFltStart)
            Do While ilLoop <> 0
                llFltStart = llFltStart - 1
                ilLoop = gWeekDayLong(llFltStart)
            Loop
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
            
                If ilUseWhichRate = 0 Then             'always use linerate
                    llWhichRate = tlCff.lActPrice
                ElseIf (ilUseWhichRate = 1) Or (ilUseWhichRate = 2 And tlClfInp(ilClf).ClfRec.lAcquisitionCost <> 0) Then          'always use acquisition rate
                    'Determine net commission if applicable
                    If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then
                        llWhichRate = tlClfInp(ilClf).ClfRec.lAcquisitionCost
                        If slGrossOrNet = "N" Then
                            ilAcqCommPct = 0
                            blAcqOK = gGetAcqCommInfoByVehicle(tlClfInp(ilClf).ClfRec.iVefCode, ilAcqLoInx, ilAcqHiInx)
                            ilAcqCommPct = gGetEffectiveAcqComm(llDate, ilAcqLoInx, ilAcqHiInx)
                            gCalcAcqComm ilAcqCommPct, llWhichRate, llAcqNet, llAcqComm
                            llWhichRate = llAcqNet
                        End If
                    Else
                        llWhichRate = tlClfInp(ilClf).ClfRec.lAcquisitionCost
                    End If
                Else                                'acq rate is 0, use the line rate
                    'llWhichRate = tlClfInp(ilClf).ClfRec.lAcquisitionCost
                    'If llWhichRate = 0 Then
                        llWhichRate = tlCff.lActPrice
                   'End If
                End If
                
                'Loop on the number of weeks in this flight
                'calc week into of this flight to accum the spot count
                If tlCff.sDyWk = "W" Then            'weekly
                    llSpots = tlCff.iSpotsWk + tlCff.iXSpotsWk
                Else                                        'daily
                    If ilLoop + 6 < llFltEnd Then           'we have a whole week
                        llSpots = tlCff.iDay(0) + tlCff.iDay(1) + tlCff.iDay(2) + tlCff.iDay(3) + tlCff.iDay(4) + tlCff.iDay(5) + tlCff.iDay(6)
                    Else
                        llFltEnd = llDate + 6
                        If llDate > llFltEnd Then
                            llFltEnd = llFltEnd       'this flight isn't 7 days
                        End If
                        For llDate2 = llDate To llFltEnd Step 1
                            ilTemp = gWeekDayLong(llDate2)
                            llSpots = llSpots + tlCff.iDay(ilTemp)
                        Next llDate2
                    End If
                End If
                If ilWkOrMonth = 1 Then                     'monthly buckets
                    'determine month that this week belongs in, then accumulate the gross and net $
                    'currently, the projections are based on STandard bdcst
                    For ilMonthInx = ilFirstProjInx To ilMaxInx - 1 Step 1       'loop thru months to find the match
                        If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                            llProject(ilMonthInx) = llProject(ilMonthInx) + (llSpots * llWhichRate)
                            llProjectSpots(ilMonthInx) = llProjectSpots(ilMonthInx) + llSpots
                            Exit For
                        End If
                    Next ilMonthInx
                Else                                    'weekly buckets
                    ilWkInx = (llDate - llStdStartDates(1)) \ 7 + 1
                    ''4-3-07 make sure the data isnt gathered beyond the period requested
                    'If ilWkInx > 0 And llDate >= llStdStartDates(LBound(llStdStartDates)) And llDate < llStdStartDates(ilMaxInx) Then   '1-24-08(UBound(llStdStartDates)) Then
                    If ilWkInx > 0 And llDate >= llStdStartDates(1) And llDate < llStdStartDates(ilMaxInx) Then   '1-24-08(UBound(llStdStartDates)) Then
                        llProject(ilWkInx) = llProject(ilWkInx) + (llSpots * llWhichRate)
                        llProjectSpots(ilWkInx) = llProjectSpots(ilWkInx) + llSpots
                    End If
                End If
            Next llDate                                     'for llDate = llFltStart To llFltEnd
        End If                                          '
        ilCff = tlCffInp(ilCff).iNextCff                   'get next flight record from mem
    Loop                                            'while ilcff <> -1
End Sub

'           gCalFlights - calculate the $ amt and spot count per day for the length of the date span
'           Distribute the $ and spots by day across the valid airing days of the week
'           Weekly buys:  determine the avg $ for the total weeks spots and allocate to each valid airing day of the week
'                         determine the avg spot counts for # spots and allocate it on each valid airing day of the week
'                         (this could end up as fractional spots)
'           Daily buys - each # spots/$ are allocated on its day defined
'
'           <input> tlClfList - line & flight info
'                   tlCFFList() array of flights
'                   llStartDate - requested date to gather averages
'                   llEndDate - requested end date to gather averages
'                   ilDayOfWk(0 to 6) - array of valid airing days (true/false)
'                   ilSpotCount - true if spot counts, else 30" unit count
'                   tlPriceTypes - inclusion/exlusion of different flight price types (charge, .00, n/c, recap, spinoff, bonus)
'           <output> llAmt - array of $ by date (from llstartdate thru llenddate)
'                   llSpots - array of spot counts (from llstartdate thru llenddate)
'
Sub gCalendarFlights(tlClfList As CLFLIST, tlCffList() As CFFLIST, llCalStartDate As Long, llCalEndDate As Long, ilDaysOfWk() As Integer, ilSpotCount As Integer, llAmt() As Long, llSpots() As Long, llAcquisition() As Long, tlPriceTypes As PRICETYPES)
    '******************************************************************************************
    '* Note: VBC id'd the following unreferenced items and handled them as described:         *
    '*                                                                                        *
    '* Local Variables (Removed)                                                              *
    '*  ilTemp                                                                                *
    '******************************************************************************************
    
    Dim ilCff As Integer
    Dim slStr As String
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim ilLoop As Integer
    Dim llTempSpots As Long
    Dim llDate As Long
    Dim llTempCalStart As Long
    Dim llTempCalEnd As Long
    Dim llGross As Long
    Dim llAcquisitionGross As Long
    Dim llAvgAmt As Long
    Dim llAvgAcqAmt As Long
    Dim llAvgSpots As Long
    Dim ilNumberDays As Integer
    Dim llAmtByDay(0 To 6) As Long
    Dim llSpotsByDay(0 To 6) As Long
    Dim llAcquisitionAmtByDay(0 To 6) As Long
    Dim ilDateInx As Integer
    Dim ilXFactor As Integer
    Dim llCalcRemainderAmt As Long
    Dim llCalcRemainderSpots As Long
    Dim llCalcAcqRemainderAmt As Long
    Dim ilIncludePriceType As Integer
    ReDim tlTempAmt(0 To 0) As Long
    ReDim tlTempSpots(0 To 0) As Long
    ReDim llTempAcquisition(0 To 0) As Long
    Dim tlCff As CFF

    'Need to always work with a complete week
    'backup the requested start date to a Monday
    llTempCalStart = llCalStartDate
    ilLoop = gWeekDayLong(llTempCalStart)
    Do While ilLoop <> 0
        llTempCalStart = llTempCalStart - 1
        ilLoop = gWeekDayLong(llTempCalStart)
    Loop
    'default the requested end date to a sunday
    llTempCalEnd = llCalEndDate
    ilLoop = gWeekDayLong(llTempCalEnd)
    Do While ilLoop <> 6
        llTempCalEnd = llTempCalEnd + 1
        ilLoop = gWeekDayLong(llTempCalEnd)
    Loop

    'Create arrays for the length of the requested period
    ReDim tlTempAmt(0 To ((llTempCalEnd - llTempCalStart) + 1)) As Long
    ReDim tlTempSpots(0 To ((llTempCalEnd - llTempCalStart) + 1)) As Long
    ReDim llTempAcquisition(0 To ((llTempCalEnd - llTempCalStart) + 1)) As Long

    ilCff = tlClfList.iFirstCff
    Do While ilCff <> -1
        tlCff = tlCffList(ilCff).CffRec

        'first decide if its Cancel Before Start
        gUnpackDate tlCff.iStartDate(0), tlCff.iStartDate(1), slStr
        llFltStart = gDateValue(slStr)
        gUnpackDate tlCff.iEndDate(0), tlCff.iEndDate(1), slStr
        llFltEnd = gDateValue(slStr)
        If llFltEnd < llFltStart Then
            For ilLoop = LBound(llAmt) To UBound(llAmt)      '7-26-07 init the values that have been projected due to CBS
                llAmt(ilLoop) = 0
                llSpots(ilLoop) = 0
                llAcquisition(ilLoop) = 0
            Next ilLoop
            Exit Sub
        End If
        'backup start date to Monday
        ilLoop = gWeekDayLong(llFltStart)
        Do While ilLoop <> 0
            llFltStart = llFltStart - 1
            ilLoop = gWeekDayLong(llFltStart)
        Loop
        'the flight dates must be within the start and end of the projection periods,
        'not be a CAncel before start flight, and have a cost > 0
        If (llFltStart < llTempCalEnd And llFltEnd >= llTempCalStart) And (llFltEnd >= llFltStart) Then
            'adjust the gather dates from flights: use flight start date or requested start date, whichever is later
            If llTempCalStart > llFltStart Then
                llFltStart = llTempCalStart
            End If
            'use flight end date or requsted end date, whichever is lesser
            If llTempCalEnd < llFltEnd Then
                llFltEnd = llTempCalEnd
            End If

            'determine if price type should be included
            ilIncludePriceType = False
            Select Case tlCff.sPriceType
                Case "N"                'no charge
                    If tlPriceTypes.iNC Then
                        ilIncludePriceType = True
                    End If
                Case "B"                'bonus
                    If tlPriceTypes.iBonus Then
                        ilIncludePriceType = True
                    End If
                Case "R"                'recapturable
                    If tlPriceTypes.iRecap Then
                        ilIncludePriceType = True
                    End If
                Case "A"                'adu
                    If tlPriceTypes.iADU Then
                        ilIncludePriceType = True
                    End If
                Case "S"                'spinoff
                    If tlPriceTypes.iSpinoff Then
                        ilIncludePriceType = True
                    End If
                Case "M"                'mg rates
                    If tlPriceTypes.iMG Then
                        ilIncludePriceType = True
                    End If
                Case Else
                    If tlCff.lActPrice = 0 Then
                        If tlPriceTypes.iZero Then
                            ilIncludePriceType = True
                        End If
                    Else
                        If tlPriceTypes.iCharge Then
                            ilIncludePriceType = True
                        End If
                    End If
            End Select
            If ilIncludePriceType Then
                'determine spot count by 30" units or spot counts.  Anything less than 30" is counted as 1 unit
                If ilSpotCount Then        'spot count(vs 30"unit counts)
                    ilXFactor = 1
                Else                        'unit count , calc # 30"units
                    ilXFactor = 0
                    ilLoop = tlClfList.ClfRec.iLen
                    Do While ilLoop >= 30
                        ilLoop = ilLoop - 30
                        ilXFactor = ilXFactor + 1
                    Loop
                    If ilLoop > 0 And ilLoop < 30 Then         'round up
                        ilXFactor = ilXFactor + 1
                    End If
                End If

                'loop thru the flights, 1 week at a time.  Always process a full week
                For llDate = llFltStart To llFltEnd Step 7
                    For ilLoop = 0 To 6         'initialize the weeks spot & $ buckets
                        llSpotsByDay(ilLoop) = 0
                        llAmtByDay(ilLoop) = 0
                        llAcquisitionAmtByDay(ilLoop) = 0
                    Next ilLoop

                    'Loop on the number of weeks in this flight
                    'calc week into of this flight to accum the spot count

                    If tlCff.sDyWk = "W" Then            'weekly
                        llTempSpots = (tlCff.iSpotsWk + tlCff.iXSpotsWk)      'need to keep fractional spots
                        'determine valid days of week
                        ilNumberDays = 0
                        For ilLoop = 0 To 6
                            If tlCff.iDay(ilLoop) > 0 Then
                                ilNumberDays = ilNumberDays + 1
                            End If
                        Next ilLoop
                        If ilNumberDays <> 0 Then
                            llGross = llTempSpots * tlCff.lActPrice     'get the weeks total $
                            'determine avg $ / day based on total $ of ordered # spots against the # of valid airing days
                            llAvgAmt = llGross / ilNumberDays       'gross contains pennies
                            llAvgSpots = (llTempSpots * ilXFactor) * 100 / ilNumberDays

                            llAcquisitionGross = llTempSpots * tlClfList.ClfRec.lAcquisitionCost
                            llAvgAcqAmt = llAcquisitionGross / ilNumberDays

                            llCalcAcqRemainderAmt = 0

                            'add any remainder to the first day
                            llCalcRemainderAmt = 0
                            llCalcRemainderSpots = 0
                            For ilLoop = 0 To 6
                                If tlCff.iDay(ilLoop) > 0 Then          'its a valid airing day, place the averages on that day
                                    llAmtByDay(ilLoop) = llAvgAmt
                                    llSpotsByDay(ilLoop) = llAvgSpots
                                    'make sure all pennies and spots add up to whole.  any remaining gets placed on first valid airing day
                                    llCalcRemainderAmt = llCalcRemainderAmt + llAvgAmt
                                    llCalcRemainderSpots = llCalcRemainderSpots + llAvgSpots

                                    'acquistion if applicable
                                    llAcquisitionAmtByDay(ilLoop) = llAvgAcqAmt
                                    llCalcAcqRemainderAmt = llCalcAcqRemainderAmt + llAvgAcqAmt
                                End If
                            Next ilLoop

                            'any remaining pennies or fraction of spots get placed on first valid airing day
                            For ilLoop = 0 To 6
                                If tlCff.iDay(ilLoop) > 0 Then
                                    llAmtByDay(ilLoop) = llAmtByDay(ilLoop) + llGross - llCalcRemainderAmt
                                    llSpotsByDay(ilLoop) = llSpotsByDay(ilLoop) + (llTempSpots * 100) - llCalcRemainderSpots
                                    llAcquisitionAmtByDay(ilLoop) = llAcquisitionAmtByDay(ilLoop) + llAcquisitionGross - llCalcAcqRemainderAmt
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                    Else                                        'daily
                        For ilLoop = 0 To 6
                          llAmtByDay(ilLoop) = tlCff.iDay(ilLoop) * tlCff.lActPrice
                          llSpotsByDay(ilLoop) = (tlCff.iDay(ilLoop) * 100) * ilXFactor 'carry to hundreds because of weekly fractional spots
                          llAcquisitionAmtByDay(ilLoop) = tlCff.iDay(ilLoop) * tlClfList.ClfRec.lAcquisitionCost
                        Next ilLoop
                    End If

                    'determine which date this week belongs in,
                    'llDate is the week being processed, llTempCalStart is the Monday of the requested start date
                    ilDateInx = llDate - llTempCalStart
                    For ilLoop = 0 To 6         'move the weeks information into the buckets for the entire requested period
                        If ilDaysOfWk(ilLoop) = True Then              'requested day is valid (from user)
                            tlTempAmt(ilDateInx + ilLoop) = llAmtByDay(ilLoop)
                            tlTempSpots(ilDateInx + ilLoop) = llSpotsByDay(ilLoop)
                            llTempAcquisition(ilDateInx + ilLoop) = llAcquisitionAmtByDay(ilLoop)
                        End If
                    Next ilLoop
                Next llDate                                 'for llDate = llFltStart To llFltEnd, go process next week
            End If                                  'ilIncludePriceType
        End If                                          '
        ilCff = tlCffList(ilCff).iNextCff            'get next flight record from mem
    Loop                                            'while ilcff <> -1
    'All flights processed and averages have been distributed by date for requested span of dates, which
    'always was processed starting on Monday and ending on Sunday.
    'Send back only the dates requested if other than not starting on Monday, and not ending on Sunday
    ReDim llSpots(0 To ((llCalEndDate - llCalStartDate))) As Long
    ReDim llAmt(0 To ((llCalEndDate - llCalStartDate))) As Long
    ReDim llAcquisition(0 To ((llCalEndDate - llCalStartDate))) As Long
    ilDateInx = llCalStartDate - llTempCalStart        'determine how many days past Monday user has requested
    For ilLoop = 0 To (llCalEndDate - llCalStartDate)
        llSpots(ilLoop) = tlTempSpots(ilLoop + ilDateInx)
        llAmt(ilLoop) = tlTempAmt(ilLoop + ilDateInx)
        llAcquisition(ilLoop) = llTempAcquisition(ilLoop + ilDateInx)
    Next ilLoop
    ReDim Preserve llAmt(0 To UBound(llAmt) + 1)
    ReDim Preserve llSpots(0 To UBound(llSpots) + 1)
    ReDim Preserve llAcquisition(0 To UBound(llAcquisition) + 1)
    Exit Sub
End Sub

'
'           gCalFlightsWithNetAcq - calculate the $ amt and spot count per day for the length of the date span
'           Distribute the $ and spots by day across the valid airing days of the week
'           Weekly buys:  determine the avg $ for the total weeks spots and allocate to each valid airing day of the week
'                         determine the avg spot counts for # spots and allocate it on each valid airing day of the week
'                         (this could end up as fractional spots)
'           Daily buys - each # spots/$ are allocated on its day defined
'
'           <input> tlClfList - line & flight info
'                   tlCFFList() array of flights
'                   llStartDate - requested date to gather averages
'                   llEndDate - requested end date to gather averages
'                   ilDayOfWk(0 to 6) - array of valid airing days (true/false)
'                   ilSpotCount - true if spot counts, else 30" unit count
'                   ilUseAcquisitionCost - true if show acq cost vs spot cost
'                   tlPriceTypes - inclusion/exlusion of different flight price types (charge, .00, n/c, recap, spinoff, bonus)
'           <output> llAmt - array of $ by date (from llstartdate thru llenddate)
'                   llAcquisition() array of gross Acq $ (if using Acq only), or Acq $ if tnet
'                   llSpots - array of spot counts (from llstartdate thru llenddate)
'
Sub gCalendarFlightsWithNetAcq(tlClfList As CLFLIST, tlCffList() As CFFLIST, llCalStartDate As Long, llCalEndDate As Long, ilDaysOfWk() As Integer, ilSpotCount As Integer, llAmt() As Long, llSpots() As Long, llAcquisition() As Long, llAcquisitionNet() As Long, ilUseAcquisitionCost As Integer, tlPriceTypes As PRICETYPES)
    '******************************************************************************************
    '* Note: VBC id'd the following unreferenced items and handled them as described:         *
    '*                                                                                        *
    '* Local Variables (Removed)                                                              *
    '*  ilTemp                                                                                *
    '******************************************************************************************
    Dim ilCff As Integer
    Dim slStr As String
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim ilLoop As Integer
    Dim llTempSpots As Long
    Dim llDate As Long
    Dim llTempCalStart As Long
    Dim llTempCalEnd As Long
    Dim llGross As Long
    Dim llAcquisitionGross As Long
    Dim llAvgAmt As Long
    Dim llAvgAcqAmt As Long
    Dim llAvgSpots As Long
    Dim ilNumberDays As Integer
    Dim llAmtByDay(0 To 6) As Long
    Dim llSpotsByDay(0 To 6) As Long
    Dim llAcquisitionAmtByDay(0 To 6) As Long
    Dim ilDateInx As Integer
    Dim ilXFactor As Integer
    Dim llCalcRemainderAmt As Long
    Dim llCalcRemainderSpots As Long
    Dim llCalcAcqRemainderAmt As Long
    Dim ilIncludePriceType As Integer
    ReDim tlTempAmt(0 To 0) As Long
    ReDim tlTempSpots(0 To 0) As Long
    ReDim llTempAcquisition(0 To 0) As Long
    ReDim llTempAcquisitionNet(0 To 0) As Long
    Dim ilAcqCommPct As Integer
    Dim blAcqOK As Boolean
    Dim ilAcqLoInx As Integer
    Dim ilAcqHiInx As Integer
    Dim llAcqNet As Long
    Dim llAcqComm As Long
    Dim tlCff As CFF

    'Need to always work with a complete week
    'backup the requested start date to a Monday
    llTempCalStart = llCalStartDate
    ilLoop = gWeekDayLong(llTempCalStart)
    Do While ilLoop <> 0
        llTempCalStart = llTempCalStart - 1
        ilLoop = gWeekDayLong(llTempCalStart)
    Loop
    'default the requested end date to a sunday
    llTempCalEnd = llCalEndDate
    ilLoop = gWeekDayLong(llTempCalEnd)
    Do While ilLoop <> 6
        llTempCalEnd = llTempCalEnd + 1
        ilLoop = gWeekDayLong(llTempCalEnd)
    Loop

    'Create arrays for the length of the requested period
    ReDim tlTempAmt(0 To ((llTempCalEnd - llTempCalStart) + 1)) As Long
    ReDim tlTempSpots(0 To ((llTempCalEnd - llTempCalStart) + 1)) As Long
    ReDim llTempAcquisition(0 To ((llTempCalEnd - llTempCalStart) + 1)) As Long
    ReDim llTempAcquisitionNet(0 To ((llTempCalEnd - llTempCalStart) + 1)) As Long

    ilCff = tlClfList.iFirstCff
    Do While ilCff <> -1
        tlCff = tlCffList(ilCff).CffRec

        'first decide if its Cancel Before Start
        gUnpackDate tlCff.iStartDate(0), tlCff.iStartDate(1), slStr
        llFltStart = gDateValue(slStr)
        gUnpackDate tlCff.iEndDate(0), tlCff.iEndDate(1), slStr
        llFltEnd = gDateValue(slStr)
        If llFltEnd < llFltStart Then
            For ilLoop = LBound(llAmt) To UBound(llAmt)      '7-26-07 init the values that have been projected due to CBS
                llAmt(ilLoop) = 0
                llSpots(ilLoop) = 0
                llAcquisition(ilLoop) = 0
                llAcquisitionNet(ilLoop) = 0
            Next ilLoop
            Exit Sub
        End If
        'backup start date to Monday
        ilLoop = gWeekDayLong(llFltStart)
        Do While ilLoop <> 0
            llFltStart = llFltStart - 1
            ilLoop = gWeekDayLong(llFltStart)
        Loop
        'the flight dates must be within the start and end of the projection periods,
        'not be a CAncel before start flight, and have a cost > 0
        If (llFltStart < llTempCalEnd And llFltEnd >= llTempCalStart) And (llFltEnd >= llFltStart) Then
            'adjust the gather dates from flights: use flight start date or requested start date, whichever is later
            If llTempCalStart > llFltStart Then
                llFltStart = llTempCalStart
            End If
            'use flight end date or requsted end date, whichever is lesser
            If llTempCalEnd < llFltEnd Then
                llFltEnd = llTempCalEnd
            End If

            'determine if price type should be included
            ilIncludePriceType = False
            Select Case tlCff.sPriceType
                Case "N"                'no charge
                    If tlPriceTypes.iNC Then
                        ilIncludePriceType = True
                    End If
                Case "B"                'bonus
                    If tlPriceTypes.iBonus Then
                        ilIncludePriceType = True
                    End If
                Case "R"                'recapturable
                    If tlPriceTypes.iRecap Then
                        ilIncludePriceType = True
                    End If
                Case "A"                'adu
                    If tlPriceTypes.iADU Then
                        ilIncludePriceType = True
                    End If
                Case "S"                'spinoff
                    If tlPriceTypes.iSpinoff Then
                        ilIncludePriceType = True
                    End If
                Case "M"                'mg rates
                    If tlPriceTypes.iMG Then
                        ilIncludePriceType = True
                    End If
                Case Else
                    If tlCff.lActPrice = 0 Then
                        If tlPriceTypes.iZero Then
                            ilIncludePriceType = True
                        End If
                    Else
                        If tlPriceTypes.iCharge Then
                            ilIncludePriceType = True
                        End If
                    End If
            End Select
            If ilIncludePriceType Then
                'determine spot count by 30" units or spot counts.  Anything less than 30" is counted as 1 unit
                If ilSpotCount Then        'spot count(vs 30"unit counts)
                    ilXFactor = 1
                Else                        'unit count , calc # 30"units
                    ilXFactor = 0
                    ilLoop = tlClfList.ClfRec.iLen
                    Do While ilLoop >= 30
                        ilLoop = ilLoop - 30
                        ilXFactor = ilXFactor + 1
                    Loop
                    If ilLoop > 0 And ilLoop < 30 Then         'round up
                        ilXFactor = ilXFactor + 1
                    End If
                End If

                'loop thru the flights, 1 week at a time.  Always process a full week
                For llDate = llFltStart To llFltEnd Step 7
                    For ilLoop = 0 To 6         'initialize the weeks spot & $ buckets
                        llSpotsByDay(ilLoop) = 0
                        llAmtByDay(ilLoop) = 0
                        llAcquisitionAmtByDay(ilLoop) = 0
                    Next ilLoop

                    'Loop on the number of weeks in this flight
                    'calc week into of this flight to accum the spot count

                    If tlCff.sDyWk = "W" Then            'weekly
                        llTempSpots = (tlCff.iSpotsWk + tlCff.iXSpotsWk)      'need to keep fractional spots
                        'determine valid days of week
                        ilNumberDays = 0
                        For ilLoop = 0 To 6
                            If tlCff.iDay(ilLoop) > 0 Then
                                ilNumberDays = ilNumberDays + 1
                            End If
                        Next ilLoop
                        If ilNumberDays <> 0 Then
                            llGross = llTempSpots * tlCff.lActPrice     'get the weeks total $
                            'determine avg $ / day based on total $ of ordered # spots against the # of valid airing days
                            llAvgAmt = llGross / ilNumberDays       'gross contains pennies
                            llAvgSpots = (llTempSpots * ilXFactor) * 100 / ilNumberDays

                            llAcquisitionGross = llTempSpots * tlClfList.ClfRec.lAcquisitionCost
                            llAvgAcqAmt = llAcquisitionGross / ilNumberDays

                            llCalcAcqRemainderAmt = 0

                            'add any remainder to the first day
                            llCalcRemainderAmt = 0
                            llCalcRemainderSpots = 0
                            For ilLoop = 0 To 6
                                If tlCff.iDay(ilLoop) > 0 Then          'its a valid airing day, place the averages on that day
                                    llAmtByDay(ilLoop) = llAvgAmt
                                    llSpotsByDay(ilLoop) = llAvgSpots
                                    'make sure all pennies and spots add up to whole.  any remaining gets placed on first valid airing day
                                    llCalcRemainderAmt = llCalcRemainderAmt + llAvgAmt
                                    llCalcRemainderSpots = llCalcRemainderSpots + llAvgSpots

                                    'acquistion if applicable
                                    llAcquisitionAmtByDay(ilLoop) = llAvgAcqAmt
                                    llCalcAcqRemainderAmt = llCalcAcqRemainderAmt + llAvgAcqAmt
                                End If
                            Next ilLoop

                            'any remaining pennies or fraction of spots get placed on first valid airing day
                            For ilLoop = 0 To 6
                                If tlCff.iDay(ilLoop) > 0 Then
                                    llAmtByDay(ilLoop) = llAmtByDay(ilLoop) + llGross - llCalcRemainderAmt
                                    llSpotsByDay(ilLoop) = llSpotsByDay(ilLoop) + (llTempSpots * 100) - llCalcRemainderSpots
                                    llAcquisitionAmtByDay(ilLoop) = llAcquisitionAmtByDay(ilLoop) + llAcquisitionGross - llCalcAcqRemainderAmt
                                    Exit For
                                End If
                            Next ilLoop
                        End If
                    Else                                        'daily
                        For ilLoop = 0 To 6
                          llAmtByDay(ilLoop) = tlCff.iDay(ilLoop) * tlCff.lActPrice
                          llSpotsByDay(ilLoop) = (tlCff.iDay(ilLoop) * 100) * ilXFactor 'carry to hundreds because of weekly fractional spots
                          llAcquisitionAmtByDay(ilLoop) = tlCff.iDay(ilLoop) * tlClfList.ClfRec.lAcquisitionCost
                        Next ilLoop
                    End If

                    'determine which date this week belongs in,
                    'llDate is the week being processed, llTempCalStart is the Monday of the requested start date
                    ilDateInx = llDate - llTempCalStart
                    For ilLoop = 0 To 6         'move the weeks information into the buckets for the entire requested period
                        If ilDaysOfWk(ilLoop) = True Then              'requested day is valid (from user)
                            tlTempAmt(ilDateInx + ilLoop) = llAmtByDay(ilLoop)
                            tlTempSpots(ilDateInx + ilLoop) = llSpotsByDay(ilLoop)
                            llTempAcquisition(ilDateInx + ilLoop) = llAcquisitionAmtByDay(ilLoop)
     
                            'Determine net commission if applicable
                            If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then
                                ilAcqCommPct = 0
                                blAcqOK = gGetAcqCommInfoByVehicle(tlClfList.ClfRec.iVefCode, ilAcqLoInx, ilAcqHiInx)
                                ilAcqCommPct = gGetEffectiveAcqComm(llDate + ilLoop, ilAcqLoInx, ilAcqHiInx)
                                gCalcAcqComm ilAcqCommPct, llAcquisitionAmtByDay(ilLoop), llAcqNet, llAcqComm
                                llTempAcquisitionNet(ilDateInx + ilLoop) = llAcqNet       'acq net, may be needed to tnet
                            End If
                           
'                            If ilUseAcquisitionCost Then
'                                llProject(ilMonthInx) = llProject(ilMonthInx) + (tlClfList.ClfRec.lAcquisitionCost)     'gross acq cost
'                            Else
'                                llProject(ilMonthInx) = llProject(ilMonthInx) + (tlCff.lActPrice)         'gross actual spot price
'                            End If
'
'                            llAcquisition(ilMonthInx) = llAcquisition(ilMonthInx) + (tlClfList.ClfRec.lAcquisitionCost)    'acq gross
                            
                        End If
                    Next ilLoop
                Next llDate                                 'for llDate = llFltStart To llFltEnd, go process next week
            End If                                  'ilIncludePriceType
        End If                                          '
        ilCff = tlCffList(ilCff).iNextCff            'get next flight record from mem
    Loop                                            'while ilcff <> -1
    'All flights processed and averages have been distributed by date for requested span of dates, which
    'always was processed starting on Monday and ending on Sunday.
    'Send back only the dates requested if other than not starting on Monday, and not ending on Sunday
    ReDim llSpots(0 To ((llCalEndDate - llCalStartDate))) As Long
    ReDim llAmt(0 To ((llCalEndDate - llCalStartDate))) As Long
    ReDim llAcquisition(0 To ((llCalEndDate - llCalStartDate))) As Long
    ReDim llAcquisitionNet(0 To ((llCalEndDate - llCalStartDate))) As Long
    ilDateInx = llCalStartDate - llTempCalStart        'determine how many days past Monday user has requested
    For ilLoop = 0 To (llCalEndDate - llCalStartDate)
        llSpots(ilLoop) = tlTempSpots(ilLoop + ilDateInx)
        llAmt(ilLoop) = tlTempAmt(ilLoop + ilDateInx)
        llAcquisition(ilLoop) = llTempAcquisition(ilLoop + ilDateInx)
        llAcquisitionNet(ilLoop) = llTempAcquisitionNet(ilLoop + ilDateInx)
    Next ilLoop
    ReDim Preserve llAmt(0 To UBound(llAmt) + 1)
    ReDim Preserve llSpots(0 To UBound(llSpots) + 1)
    ReDim Preserve llAcquisition(0 To UBound(llAcquisition) + 1)
    ReDim Preserve llAcquisitionNet(0 To UBound(llAcquisitionNet) + 1)
    Exit Sub
End Sub

'***************************************************************************************
'*
'*      Procedure Name:  gBuildAcqCommInfo - Build all information from VBF for the
'               Acquistion Commission.  Build array of vehicles and effectives dates of
'               acq. commission percents.  Sort by vehicle and effective date.
'               Create another array that points to the sorted array to maintain the lowest
'               and highest indices for the vehicles information.
'*
'***************************************************************************************
Public Function gBuildAcqCommInfo(RptForm As Form) As Boolean
    Dim ilRet As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim slStr As String
    Dim ilUpper As Long
    Dim llDate As Long
    Dim ilLoInx As Integer
    Dim ilHiInx As Integer
    Dim blFirstTime As Boolean
    Dim ilUpperAcqCommInx As Integer
    Dim ilPrevVefCode As Integer
    Dim ilLoop As Integer
        
    gBuildAcqCommInfo = True
    'ReDim tgAcqComm(1 To 1) As ACQCOMM
    'ReDim tgAcqCommInx(1 To 1) As ACQCOMMINX
    ReDim tgAcqComm(0 To 0) As ACQCOMM
    ReDim tgAcqCommInx(0 To 0) As ACQCOMMINX
       
    If Not (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then 'Acquisition Commissionable then
        Exit Function
    End If

    hmVbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVbf, "", sgDBPath & "Vbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVbf)
        btrDestroy hmVbf
        gBuildAcqCommInfo = False
        Exit Function
    End If
    
    imVBfRecLen = Len(tmVbf)
    btrExtClear hmVbf   'Clear any previous extend operation
    ilExtLen = Len(tmVbf)  'Extract operation record size
    
    ilUpper = UBound(tgAcqComm)
    ilRet = btrGetFirst(hmVbf, tmVbf, imVBfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
        Call btrExtSetBounds(hmVbf, llNoRec, -1, "UC", "VBF", "")
        ilRet = btrExtAddField(hmVbf, 0, ilExtLen)  'Extract the whole record
        On Error GoTo mBuildAcqCommInfoErr
        gBtrvErrorMsg ilRet, "gBuildAcqCommInfo (btrExtAddField):" & "Vbf.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hmVbf, tmVbf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mBuildAcqCommInfoErr
            gBtrvErrorMsg ilRet, "mBuildAcqCommInfo (btrExtGetNextExt):" & "Vbf.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmVbf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hmVbf, tmVbf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                '10-16-15 ignore all the records except aquisition commission records
                If tmVbf.sMethod = "N" Then     '10-16-15 take all the acq percent table
                    'create key for sorting:  vehicle name & effective date, descending
                    'ilVefInx = gBinarySearchVef(tmVbf.ivefcode)
                    'if ilVefInx > 0 then
                    'slStr = Trim$(tgMVef(ilVefInx).sName)
                    slStr = Trim$(str(tmVbf.iVefCode))
                    Do While Len(slStr) < 6                 'left fill with zeroes
                            slStr = "0" & slStr
                    Loop
                    tgAcqComm(ilUpper).sKey = slStr & "|"
                    
                    gUnpackDateLong tmVbf.iStartDate(0), tmVbf.iStartDate(1), llDate
                    slStr = Trim$(str$(llDate))
                    Do While Len(slStr) < 6
                            slStr = "0" & slStr
                    Loop
                    
                    tgAcqComm(ilUpper).sKey = Trim$(tgAcqComm(ilUpper).sKey) & slStr
                    
                    tgAcqComm(ilUpper).iVefCode = tmVbf.iVefCode
                    tgAcqComm(ilUpper).iAcqCommPct = tmVbf.iAcqCommPct
                    gUnpackDateLong tmVbf.iStartDate(0), tmVbf.iStartDate(1), tgAcqComm(ilUpper).lEffStartDate
                    If tmVbf.iEndDate(0) <= 7 And tmVbf.iEndDate(1) = 0 Then                'tfn
                        tgAcqComm(ilUpper).lEndDate = gDateValue("12/31/2060")
                    Else
                        gUnpackDateLong tmVbf.iEndDate(0), tmVbf.iEndDate(1), tgAcqComm(ilUpper).lEndDate
                    End If
                     
                    'ReDim Preserve tgAcqComm(1 To UBound(tgAcqComm) + 1) As ACQCOMM
                    ReDim Preserve tgAcqComm(LBound(tgAcqComm) To UBound(tgAcqComm) + 1) As ACQCOMM
                    ilUpper = ilUpper + 1
                    'endif
                End If
                
                ilRet = btrExtGetNext(hmVbf, tmVbf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hmVbf, tmVbf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    
    'sort the array
    If UBound(tgAcqComm) - 1 > 0 Then
            'ArraySortTyp fnAV(tgAcqComm(), 1), UBound(tgAcqComm) - 1, 0, LenB(tgAcqComm(0)), 0, LenB(tgAcqComm(0).sKey), 0
            ArraySortTyp fnAV(tgAcqComm(), 0), UBound(tgAcqComm), 0, LenB(tgAcqComm(0)), 0, LenB(tgAcqComm(0).sKey), 0
    End If

    'create array sorted by vehicle with indices of where each vehicles acq commission Pc5 info starts and ends
    ilPrevVefCode = -1
    ilLoInx = LBound(tgAcqComm)
    ilHiInx = LBound(tgAcqComm)
    ilUpperAcqCommInx = LBound(tgAcqCommInx)
    blFirstTime = True
    For ilLoop = LBound(tgAcqComm) To UBound(tgAcqComm)
            If blFirstTime Then
                    tgAcqCommInx(ilUpperAcqCommInx).iVefCode = tgAcqComm(ilLoop).iVefCode
                    tgAcqCommInx(ilUpperAcqCommInx).iLoInx = ilLoInx
                    tgAcqCommInx(ilUpperAcqCommInx).iHiInx = ilHiInx
                    ilPrevVefCode = tgAcqComm(ilLoop).iVefCode
                    blFirstTime = False
            End If
            If ilPrevVefCode = tgAcqComm(ilLoop).iVefCode Then
                ilHiInx = ilLoop
            Else            'change in vehicle
                tgAcqCommInx(ilUpperAcqCommInx).iLoInx = ilLoInx
                tgAcqCommInx(ilUpperAcqCommInx).iHiInx = ilHiInx
                ilLoInx = ilLoop
                ilHiInx = ilLoInx
                ReDim Preserve tgAcqCommInx(LBound(tgAcqCommInx) To ilUpperAcqCommInx + 1) As ACQCOMMINX
                ilUpperAcqCommInx = ilUpperAcqCommInx + 1
                tgAcqCommInx(ilUpperAcqCommInx).iVefCode = tgAcqComm(ilLoop).iVefCode
                tgAcqCommInx(ilUpperAcqCommInx).iLoInx = ilLoInx
                tgAcqCommInx(ilUpperAcqCommInx).iHiInx = ilHiInx
                ilPrevVefCode = tgAcqComm(ilLoop).iVefCode
            End If
    Next ilLoop
    ilRet = btrClose(hmVbf)
    btrDestroy hmVbf
    Exit Function
                
mBuildAcqCommInfoErr:
    On Error GoTo 0
    gBuildAcqCommInfo = False
    Exit Function
End Function

'           gGetCommbyVehicle - An array exists by calling gBuildAcqCommInfo which contains all vehicles with start and
'           end indices to another array containing all effective commission dates for the vehicle.
'
'           By indexing into this array, obtain the lo and hi indices of where the matching vehicles commission rates
'           are stored (tgAcqComm contains the effectives dates & commission rates).
'           tgAcqCommInx contains the start/end indices pointing to tgAcqComm for the matching vehicle
'           <input> vehicle code
'           <output>  ilLoInx - starting index of matching vehicles commission info
'                     ilHiInx - ending index of matching vehicles commission info
'           <return>  true - valid vehicle found
'                     false - no vehicle found
Public Function gGetAcqCommInfoByVehicle(ilVefCode As Integer, ilLoInx As Integer, ilHiInx As Integer) As Boolean
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer

    gGetAcqCommInfoByVehicle = False
    ilLoInx = -1
    ilHiInx = -1
    ilMin = LBound(tgAcqCommInx)
    ilMax = UBound(tgAcqCommInx) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilVefCode = tgAcqCommInx(ilMiddle).iVefCode Then
            'found the match
            gGetAcqCommInfoByVehicle = True
            ilLoInx = tgAcqCommInx(ilMiddle).iLoInx
            ilHiInx = tgAcqCommInx(ilMiddle).iHiInx
            Exit Function
        ElseIf ilVefCode < tgAcqCommInx(ilMiddle).iVefCode Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    gGetAcqCommInfoByVehicle = -1
    Exit Function
End Function

'           gGetEffectiveAcqComm - loop into the table of effective acquisition commission rates (tgAcqComm) and return back
'           the effective commission to be used
'           <input> ilLoInx - starting point to find the effective comm
'                   ilHiInx - ending point to find the effective comm
'           <return> ilAcqPct - effective percent
'
Public Function gGetEffectiveAcqComm(llDate As Long, ilLoInx As Integer, ilHiInx As Integer) As Integer
    Dim ilLoopOnInx As Integer

    gGetEffectiveAcqComm = 0
    If ilLoInx < 0 Or ilHiInx < 0 Then
        Exit Function
    End If
    
    For ilLoopOnInx = ilLoInx To ilHiInx
        If llDate >= tgAcqComm(ilLoopOnInx).lEffStartDate And llDate <= tgAcqComm(ilLoopOnInx).lEndDate Then
            gGetEffectiveAcqComm = tgAcqComm(ilLoopOnInx).iAcqCommPct
            Exit For
        End If
    Next ilLoopOnInx
    Exit Function
End Function

'           gCalcAcqComm -calculate amount of commission from gross $
'           <input>  llGross - Gross $ (pennies included)
'                    ilAcqComm (xxx.xx)
'
'           <output> llNet
'                    llComm
Public Sub gCalcAcqComm(ilAcqComm As Integer, llGross As Long, llNet As Long, llComm As Long)
    Dim slGrossAmt As String
    Dim slNetAmt As String
    Dim slCommAmt As String
    Dim slCommPct As String

    llNet = 0
    llComm = 0
    'convert to string math
    slCommPct = gIntToStrDec(ilAcqComm, 2)          'xxx.xx
    slGrossAmt = gLongToStrDec(llGross, 2)             'xxxxxxx.xx
    
    slCommAmt = gMulStr(slCommPct, slGrossAmt)
    slCommAmt = gRoundStr(slCommAmt, ".01", 0)            'round to $
     
    llComm = gStrDecToLong(slCommAmt, 0)
    llNet = llGross - llComm
    Exit Sub
End Sub

'           'build array of acquisition commission percents for 1 year by month by vehicle
'           Acquisition commission percents are retained by standard month only
'           gBuildAcqComm must be called first as this routine is generated
'           from the tgAcqComm array
'           <input> llStartDate - date to start generating the std start dates
'                   ilMonthOrQtr - calculate for a start quarter or month
'           <output> array of 1 year of standard start dates and its associated varying acquisition comm percent by vehicle
'
Public Sub gBuildACQPctByPeriod(llStartDate As Long, ilMonth As Integer, tlACQPctByPeriod() As ACQPctByPeriod)
    Dim ilLoopOnAcqComm As Integer
    Dim ilUpper As Integer
    Dim ilLoopOnPeriod As Integer
    'Dim llStdStartDates(1 To 13) As Long
    Dim llStdStartDates(0 To 13) As Long    'Index zero ignored
    Dim ilTemp As Integer
    Dim llPacingDate As Long                'pacing not applicable for varying commissions
    Dim ilLastBilledInx As Integer          'last billed inx not applicable for the std dates and varying commissions
    Dim llLastBilled As Long                'last billed date not applicable for the commission pct values
    Dim ilLoopOnInx As Integer
    Dim ilLoInx As Integer
    Dim ilHiInx As Integer
    Dim blAcqOK As Boolean

    'ReDim tlACQPctByPeriod(1 To 1) As ACQPctByPeriod
    'ilUpper = 1
    ReDim tlACQPctByPeriod(0 To 0) As ACQPctByPeriod
    ilUpper = UBound(tlACQPctByPeriod)  '0
    llPacingDate = 0
    'igYear needs to have the year to gather; ilMonth has been passed as parameter
    gSetupBOBDates 2, llStdStartDates(), llLastBilled, ilLastBilledInx, llPacingDate, ilMonth  'build array of std start & end dates

    'intialize the array yearly array of months and associated acquisition commission percents
    For ilLoopOnAcqComm = LBound(tgAcqCommInx) To UBound(tgAcqCommInx) - 1
        For ilLoopOnPeriod = 1 To 12
            tlACQPctByPeriod(ilUpper).iAcqCommPct(ilLoopOnPeriod) = 0       'initialize all the years %
            tlACQPctByPeriod(ilUpper).lStdStartDates(ilLoopOnPeriod) = llStdStartDates(ilLoopOnPeriod)
        Next ilLoopOnPeriod
        tlACQPctByPeriod(ilUpper).iVefCode = tgAcqCommInx(ilLoopOnAcqComm).iVefCode
        'ReDim Preserve tlACQPctByPeriod(1 To ilUpper + 1) As ACQPctByPeriod
        ReDim Preserve tlACQPctByPeriod(0 To ilUpper + 1) As ACQPctByPeriod
        ilUpper = ilUpper + 1
    Next ilLoopOnAcqComm
    
    'Get the monthly percentages by vehicle and store it with its associated month
    For ilLoopOnAcqComm = LBound(tlACQPctByPeriod) To UBound(tlACQPctByPeriod) - 1        'setup the correct comm % for the vehicle in the 12 bucket array of months
        blAcqOK = gGetAcqCommInfoByVehicle(tlACQPctByPeriod(ilLoopOnAcqComm).iVefCode, ilLoInx, ilHiInx)
        If blAcqOK Then
            For ilLoopOnInx = ilLoInx To ilHiInx
                For ilLoopOnPeriod = 1 To 12
                  If tlACQPctByPeriod(ilLoopOnAcqComm).lStdStartDates(ilLoopOnPeriod) >= tgAcqComm(ilLoopOnInx).lEffStartDate And tlACQPctByPeriod(ilLoopOnAcqComm).lStdStartDates(ilLoopOnPeriod) <= tgAcqComm(ilLoopOnInx).lEndDate Then
                      tlACQPctByPeriod(ilLoopOnAcqComm).iAcqCommPct(ilLoopOnPeriod) = tgAcqComm(ilLoopOnInx).iAcqCommPct
                      
                  End If
                Next ilLoopOnPeriod
            Next ilLoopOnInx
        End If
    Next ilLoopOnAcqComm
    
    Exit Sub
End Sub

'               gVaryAcqComm - Take array of total acquisition costs by line and adjust for net amt, if applicable for Standard reporting
'               <input> llAcquisitionCost() - array of acquisition costs by month
'                       tlAcqPctByPeriod() - array of std months start dates and associated comm pcts for acq adjustments
'               <output> - array of net Acquisition values
'
Public Sub gVaryAcqComm(ilVefCode As Integer, llAcquisitionCost() As Long, tlACQPctByPeriod() As ACQPctByPeriod, llAcquisitionNet() As Long, llAcquisitionComm() As Long)
    Dim ilLoopOnPeriod As Integer
    Dim ilLoopOnVef As Integer
    Dim blFound As Boolean

    'initialize the resulting field of net values
    For ilLoopOnPeriod = LBound(llAcquisitionNet) To UBound(llAcquisitionNet) - 1
        llAcquisitionNet(ilLoopOnPeriod) = 0
        llAcquisitionComm(ilLoopOnPeriod) = 0
    Next ilLoopOnPeriod
               
    If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then 'Acquisition Commissionable then
        blFound = False

        For ilLoopOnVef = LBound(tlACQPctByPeriod) To UBound(tlACQPctByPeriod) - 1
            If tlACQPctByPeriod(ilLoopOnVef).iVefCode = ilVefCode Then
                blFound = True
                Exit For
            End If
        Next ilLoopOnVef
    
        If blFound Then
            For ilLoopOnPeriod = LBound(llAcquisitionCost) To UBound(llAcquisitionCost) - 1
                gCalcAcqComm tlACQPctByPeriod(ilLoopOnVef).iAcqCommPct(ilLoopOnPeriod), llAcquisitionCost(ilLoopOnPeriod), llAcquisitionNet(ilLoopOnPeriod), llAcquisitionComm(ilLoopOnPeriod)
            Next ilLoopOnPeriod
        End If
    Else
        If ((Asc(tgSpf.sUsingFeatures2) And BARTER) = BARTER) Then
            For ilLoopOnPeriod = LBound(llAcquisitionCost) To UBound(llAcquisitionCost) - 1
                llAcquisitionNet(ilLoopOnPeriod) = llAcquisitionCost(ilLoopOnPeriod)
            Next ilLoopOnPeriod
        End If
    End If
    Exit Sub
End Sub

'           gSetActFieldsforOutput - arrays created for line projected $, Gross acquisition $, and net acquisition $
'           Generalized routines will calculate and output based on the projected $, just acquisition $, or adjusting projected $
'           for triple-net figures (subtracting out the net acquisition $)
'           Set up the array so the common routine uses the correct arrays
'           <input> ilAdjustAcquisition :  true if t-net report (net actual costs minus merchan $ minus Promo $ minus net Acq cost
'                   ilUseAcquisitionCost - true if report to show the gross or net acqusition amount (vs the spot price amount)
'                   llAcquisitionNet()- net acquisition costs if varying acq commission; otherwise 0
'           <output> llAcquisition() - array of gross acquisition $ that can be subtracted out for tnet reporting
'                     llProject() - array of actual spot cost or gross acquisition costs as prinary amt to show on report (could be netted down)
'
Public Sub gSetAcqFieldsForOutput(ilAdjustAcquisition As Integer, ilUseAcquisitionCost As Integer, slGrossOrNet As String, llProject() As Long, llAcquisition() As Long, llAcquisitionNet() As Long)
    Dim ilLoop As Integer

    'llProject - gross actual spot cost or gross acquisition costs
    'llAcquisition - gross acquisition costs
    'llAcquisitionNet - net acquisition costs if varying acq commissions; otherwise 0
    If ilAdjustAcquisition Then         'subtract out the acq gross or net values?
        'For ilLoop = LBound(llAcquisition) To UBound(llAcquisition) - 1
        For ilLoop = 1 To UBound(llAcquisition) - 1
            If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then    'varying comm, use the net value
                llAcquisition(ilLoop) = llAcquisitionNet(ilLoop)
            End If
        Next ilLoop
    End If
    If ilUseAcquisitionCost Then        'showing acq cost instead of spot price
        'For ilLoop = LBound(llProject) To UBound(llProject) - 1
        For ilLoop = 1 To UBound(llProject) - 1
            If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE And slGrossOrNet = "N" Then
                llProject(ilLoop) = llAcquisitionNet(ilLoop)
            End If
        Next ilLoop
    End If
    Exit Sub
End Sub

'
'                      Calculate Gross & Net $, and Split Cash/Trade $ from a schedule line
'                      mCalcMonthAmt - Loop and calculate the gross and net values for up to 36 months
'
'                       <input> llTempGross - 36 months of projected $ (from contract line)
'                               ilLastBilledInx - index to last month invoiced.
'                               ilMaxPeriods - # of periods to process
'                               ilCorT - 1 = Cash , 2 = Trade processing
'                               slCashAgyComm - agency comm %
'                               tlChf - contract header buffer
'                       <output> llTempGross - altered if split cash/trade calculation
'                               llTempNet - 36months projected net $
'
Public Sub gCalcMonthAmt(llTempGross() As Long, llTempNet() As Long, llTempAcquisition() As Long, ilLastBilledInx As Integer, ilMaxPeriods As Integer, ilCorT As Integer, slCashAgyComm As String, tlChf As CHF)
    Dim ilTemp As Integer
    Dim slAmount As String
    Dim slSharePct As String
    Dim slStr As String
    Dim slCode As String
    Dim slDollar As String
    Dim slNet As String
    Dim slAcquisition As String
    Dim slAcqAmount As String
    Dim slAcqShare As String
    Dim slPctTrade As String

    For ilTemp = ilLastBilledInx To ilMaxPeriods              'loop on # buckets to process.
        slAmount = gLongToStrDec(llTempGross(ilTemp), 2)
        slSharePct = gLongToStrDec(10000, 2)
        slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
        slStr = gRoundStr(slStr, "1", 0)
        'calc the acquisition share for split cash/trade
        slAcqAmount = gLongToStrDec(llTempAcquisition(ilTemp), 2)
        slAcqShare = gMulStr(slSharePct, slAcqAmount)
        slAcqShare = gRoundStr(slAcqShare, "1", 0)
        If ilCorT = 1 Then                 'all cash commissionable
            slPctTrade = gIntToStrDec(tlChf.iPctTrade, 0)

            slCode = gSubStr("100.", slPctTrade)                'get the cash % (100-trade%)
            slDollar = gDivStr(gMulStr(slStr, slCode), "100")              'slsp gross
            slNet = gRoundStr(gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "100.00"), ".01", 0)
            slAcquisition = gDivStr(gMulStr(slAcqShare, slCode), "100")              'Acquisition is always net (same as gross)
        Else
            If ilCorT = 2 Then                'at least cash is commissionable
                slCode = gIntToStrDec(tlChf.iPctTrade, 0)
                slDollar = gDivStr(gMulStr(slStr, slCode), "100")
                slAcquisition = gDivStr(gMulStr(slAcqShare, slCode), "100")
                
                If tlChf.iAgfCode > 0 And tlChf.sAgyCTrade = "Y" Then
                    slNet = gRoundStr(gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "100.00"), "1", 0)
                Else
                    slNet = slDollar    'no commission , net is same as gross
                End If
            End If
        End If
        llTempGross(ilTemp) = Val(slDollar)
        llTempNet(ilTemp) = Val(slNet)
        llTempAcquisition(ilTemp) = Val(slAcquisition)
    Next ilTemp
    Exit Sub
End Sub
