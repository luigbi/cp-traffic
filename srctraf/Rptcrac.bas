Attribute VB_Name = "RPTCRAC"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrac.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
'Public igYear As Integer                'budget year used for filtering
'Public igMonthOrQtr As Integer          'entered month or qtr
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
Dim hmCHF As Integer            'Contract header file handle
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChf As CHF
Dim tlChfAdvtExt() As CHFADVTEXT
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmSlf As Integer            'Salesperson file handle
Dim imSlfRecLen As Integer      'SLF record length
Dim tmSlf As SLF
Dim hmUrf As Integer            'User file handle
Dim imUrfRecLen As Integer      'URF record length
Dim tmUrf As URF
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Type PRICELIST                      'record containing slsp actuals for a year
    iIndex As Integer               'relative year offset (1 = base year, 2 = last year, etc)
    iVefCode  As Integer
    iSlfCode As Integer
    iSofCode As Integer
    lTotalYear As Long
    'lDollars(1 To 12)  As Long      'no pennies
    lDollars(0 To 12)  As Long      'no pennies. Index zero ignored
End Type
Type SLSVEHLIST
    iSofCode As Integer              'sales office  code
    iSlfCode As Integer                 'slsp code
    iVefCode As Integer                 'vehicle coe
End Type
Dim tmBaseYear() As SLSVEHLIST
Dim tlSofList() As SOFLIST    'list of selling office codes and sales sourcecodes
'********************************************************************************************
'
'                   gCrActSlspCmp - Prepass for Actual Comparison Salesperson report
'                   showing all 12 months.  Up to 4 years of comparison data.
'
'                   User selectivity:  Base Year
'                                      # of years to compare (0 indicates show base year only
'                                      Corp or Std Month
'                   This is a report where all contracts are gather regardless of the entry
'                   date and totalled in its respective corporate or standard month.
'                   Records are created one record per slsp and vehicle and year.  Before
'                   records are written to disk, a check is made to insure that a record
'                   exists for the "Base Year", even if it is all zero.  Crystal needs
'                   something to compare to.
'                   Created:  11/7/97 D. Hosaka
'                   4-20-00 New commission structure by subcompany/vehicle
'********************************************************************************************
Sub gCrActSlspCmp()
Dim ilRet As Integer                    '
Dim ilClf As Integer                    'loop for schedule lines
Dim ilHOState As Integer                'retrieve only latest order or revision
Dim slCntrTypes As String               'retrieve remnants, PI, DR, etc
Dim slCntrStatus As String              'retrieve H, O G or N (holds, orders, unsch hlds & orders)
Dim ilCurrentRecd As Integer            'loop for processing last years contracts
Dim llContrCode As Long                 'Internal Contr code to build all lines & flights of a contract
Dim ilFoundOne As Integer               'Found a matching  office built into mem
Dim ilFoundVeh As Integer               'true if selection by vehicle and line matches the selection
Dim ilSlspLoop As Integer
Dim ilTemp As Integer
Dim ilLoop As Integer                   'temp loop variable
Dim il12X As Integer
Dim ilUpper As Integer
Dim slTemp As String                    'temp string for dates
Dim ilCalType As Integer                '1 = std, 2 = corp calendar
'ReDim llTempProject(1 To 12) As Long    '1 years projection
ReDim llTempProject(0 To 12) As Long    '1 years projection. Index zero ignored
'ReDim llProject(1 To 12) As Long       '$ for 12 months
Dim llDate As Long                      'temp date variable
Dim llDate2 As Long
Dim llCalcGross As Long                 'total cntr $ for the grf rcord
Dim slNameCode As String
Dim slCode As String
Dim ilSaveSof As Integer                'office processed from the max of 10 splits possible
Dim ilNoYears As Integer                '# years to compare
Dim ilYearInx As Integer                'index to date array to process for start dates
Dim ilBySlsp As Integer                   'true if selection by slsp
Dim ilByVeh As Integer                  'true if selection by vehicle
Dim ilStartMonth As Integer             'start month of year:  for std always1, otherwise corp start month from corp calendar
'Date used to gather information
'String formats for generalized date conversions routines
'Long formats for testing
'Packed formats to store in GRF record
Dim ilTYStartYr As Integer              'year of this years start date     (1997-1998)
Dim slTYStart As String                 'start date of this year to begin gathering  (string)
Dim llTYStart As Long                   'start date of this year to begin gathering (Long)
Dim slTYEnd As String
'Dim slWeekTYStart As String              'start date of week for this years new business entered this week
'Dim llWeekTYStart As Long                'start date of week for this years new business entered on te user entered week
'ReDim ilWeekTYStart(0 To 1) As Integer     'packed format for GRF record
'Dim llEntryDate As Long                 'date entered from cntr header
'Month Starts to gather projection $ from flights
'ReDim llTYStartDates(1 To 13) As Long        'this year corp or std start dates for next 5 quarters
'ReDim llTempStarts(1 To 13) As Long         'start dates for one of 5 years
ReDim llTempStarts(0 To 13) As Long         'start dates for one of 5 years. Index zero ignored
'ReDim llAllStarts(1 To 5, 1 To 13) As Long         'temp array for start dates for 13 months  for 5 years (base + 4 years back)
ReDim llAllStarts(0 To 5, 0 To 13) As Long         'temp array for start dates for 13 months  for 5 years (base + 4 years back). Index zero ignored
Dim ilMnfSubCo As Integer           '4-20-00
'   end of date variables
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
    hmUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmUrf)
        btrDestroy hmUrf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imUrfRecLen = Len(tmUrf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        btrDestroy hmUrf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmUrf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmUrf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        btrDestroy hmSlf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmUrf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imSlfRecLen = Len(tmSlf)
    ReDim tgClfAC(0 To 0) As CLFLIST
    tgClfAC(0).iStatus = -1 'Not Used
    tgClfAC(0).lRecPos = 0
    tgClfAC(0).iFirstCff = -1
    ReDim tgCffAC(0 To 0) As CFFLIST
    tgCffAC(0).iStatus = -1 'Not Used
    tgCffAC(0).lRecPos = 0
    tgCffAC(0).iNextCff = -1
    'build array of selling office codes and their sales sources.
    ilTemp = 0
    ilRet = btrGetFirst(hmSlf, tmSlf, imSlfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tlSofList(0 To ilTemp) As SOFLIST
        tlSofList(ilTemp).iSofCode = tmSlf.iSofCode         'save selling office code to compare to selectivity
        tlSofList(ilTemp).iMnfSSCode = tmSlf.iCode          'replace Sales source code with slsp code
        ilRet = btrGetNext(hmSlf, tmSlf, imSlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
    Loop
    ilBySlsp = RptSelAc!rbcSelCInclude(0).Value
    ilByVeh = RptSelAc!rbcSelCInclude(1).Value
    'setup year from user input
    ilTYStartYr = Val(RptSelAc!edcSelCFrom.Text)
    ilNoYears = Val(RptSelAc!edcSelCFrom1.Text)
    ilFoundOne = 0
    For ilLoop = 1 To ilNoYears + 1                 '0 indicates show base year only
        If RptSelAc!rbcSelCSelect(0).Value Then      'corp month or qtr? (vs std)
            ilCalType = 2                   'corp flag to store in grf
            ilYearInx = ilTYStartYr - ilLoop + 1
                ilTemp = gGetCorpCalIndex(ilYearInx)
            'If ilTemp > 0 Then
            If ilTemp >= 0 Then
                gGetStartEndYear 1, ilYearInx, slTYStart, slTYEnd
                ilFoundOne = ilFoundOne + 1
                ilStartMonth = tgMCof(ilTemp).iStartMnthNo
            End If
        Else
            ilCalType = 1                   'std flag to store in grf
            ilYearInx = ilTYStartYr - ilLoop + 1
            gGetStartEndYear 2, ilYearInx, slTYStart, slTYEnd
            ilFoundOne = ilFoundOne + 1
            ilStartMonth = 1
        End If
        'Determine startdates for this year for 13 months
        gBuildStartDates slTYStart, ilCalType, 13, llTempStarts()
        'Store all years start dates in array
        For ilTemp = 1 To 13
            llAllStarts(ilLoop, ilTemp) = llTempStarts(ilTemp)
        Next ilTemp
    Next ilLoop
    ilNoYears = ilFoundOne            '# of years to compare may be altred due to lack of corporate dates defined
    'all done obtaining start dates of all years of comparison
    'Build array of the base date incase user is only requesting base date (no comparisons(
    For ilLoop = 1 To 13
        llTempStarts(ilLoop) = llAllStarts(1, ilLoop)
    Next ilLoop
    slTYEnd = Format$(llAllStarts(1, 13) - 1, "m/d/yy")     'last date reqd so the contr gathering has an earliest & latest date
    slTYStart = Format$(llAllStarts(ilNoYears, 1), "m/d/yy")  'earliest date rqd so contr gathering has an earliest & latest date
    llTYStart = gDateValue(slTYStart)
                                                            'tosearch & gather
    'Gather all contracts for previous year and current year whose effective date entered
    'is prior to the effective date that affects either previous year or current year
    slCntrTypes = gBuildCntTypes()
    slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)
    ilRet = gObtainCntrForDate(RptSelAc, slTYStart, slTYEnd, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())

    ReDim tmSlspTot(0 To 0) As PRICELIST            'array of records containing unique ofc/slsp/veh for a given year
    'ReDim tmBaseYear(1 To 1) As SLSVEHLIST          'array of ofc/sls/veh that have been gathered for base year
    ReDim tmBaseYear(0 To 0) As SLSVEHLIST          'array of ofc/sls/veh that have been gathered for base year
    'All contracts have been retrieved for all of this year
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
        ilFoundOne = True
        llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
        'Retrieve the contract, schedule lines and flights
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfAC, tgClfAC(), tgCffAC())
        'Test for selective slsp
        ilFoundOne = True
        If Not (RptSelAc!ckcAll.Value = vbUnchecked) And ilBySlsp Then           'selected slsp picked, see if the one selected belongs to this cnt
            ilFoundOne = False
            For ilSlspLoop = 0 To 9               'max 10 slsp
                If tgChfAC.iSlfCode(ilSlspLoop) > 0 Then
                    'find the associated office from the slsp
                    'For ilLoop = 0 To UBound(tlSofList)
                    '    If tgChfAC.islfCode(ilSlspLoop) = tlSofList(ilLoop).imnfSsCode Then
                    '        ilFoundOne = True
                    '        Exit For
                    '    End If
                    'Next ilLoop
                    'If ilFoundOne Then          'got the office reference, was it selected?
                        ilFoundOne = False
                        For ilTemp = 0 To RptSelAc!lbcSelection(0).ListCount - 1 Step 1
                            If RptSelAc!lbcSelection(0).Selected(ilTemp) Then              'selected advt
                                slNameCode = tgSalesperson(ilTemp).sKey      'pick up slsp code
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                'If Val(slCode) = tlSofList(ilLoop).isofCode Then
                                If Val(slCode) = tgChfAC.iSlfCode(ilSlspLoop) Then
                                    ilFoundOne = True
                                    Exit For
                                End If
                            End If
                        Next ilTemp
                    'End If
                End If
                If ilFoundOne Then          'only looking for at least 1 valid office, then ok to project contract
                    ilSlspLoop = 9          'terminate the loop
                End If
            Next ilSlspLoop
        End If                      'not ckcall

        If ilFoundOne Then           'this cnts office has been selected, go ahead and process
            For ilClf = LBound(tgClfAC) To UBound(tgClfAC) - 1 Step 1
                tmClf = tgClfAC(ilClf).ClfRec
                ilFoundVeh = True
                If ((Not RptSelAc!ckcAll.Value = vbUnchecked)) And (ilByVeh) Then    'if by vehicle, only include those selected
                    ilFoundVeh = False
                    For ilTemp = 0 To RptSelAc!lbcSelection(1).ListCount - 1 Step 1
                        If RptSelAc!lbcSelection(1).Selected(ilTemp) Then              'selected veh
                            slNameCode = tgCSVNameCode(ilTemp).sKey      'pick up veh code
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            If Val(slCode) = tmClf.iVefCode Then
                                ilFoundVeh = True
                                Exit For
                            End If
                        End If
                    Next ilTemp
                End If
                If ilFoundVeh Then
                    'get cnts earliest and latest dates of line to see if it spans the requested period(s)
                    gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), slTemp       '
                    llDate = gDateValue(slTemp)
                    gUnpackDate tmClf.iEndDate(0), tmClf.iEndDate(1), slTemp
                    llDate2 = gDateValue(slTemp)
                    'Determine which year this contract belongs in
                    'ReDim tmLnYrTot(1 To 5) As PRICELIST        'init for this line for 5 max years comparison figures
                    ReDim tmLnYrTot(0 To 5) As PRICELIST        'init for this line for 5 max years comparison figures. Index zero ignored
                    For ilYearInx = 1 To ilNoYears + 1
                        If gSpanDates(llDate, llDate2, llAllStarts(ilYearInx, 1), llAllStarts(ilYearInx, 13)) Then
                            For ilLoop = 1 To 13            'these dates span- move into common date field to process for the year
                                llTempStarts(ilLoop) = llAllStarts(ilYearInx, ilLoop)
                            Next ilLoop
                            If tmClf.sType = "S" Or tmClf.sType = "H" Then
                                gBuildFlights ilClf, llTempStarts(), 1, 13, llTempProject(), 1, tgClfAC(), tgCffAC()
                                'ilTemp = UBound(tmLnYrTot)
                                tmLnYrTot(ilYearInx).iIndex = ilYearInx
                                tmLnYrTot(ilYearInx).iVefCode = tmClf.iVefCode
                                llCalcGross = 0                     'schedule line gross $ for year 'Project the monthly $ from the flights
                                'accumulate the years total to place  into record
                                For il12X = 1 To 12
                                    tmLnYrTot(ilYearInx).lDollars(il12X) = llTempProject(il12X) \ 100    'remove pennies
                                    llCalcGross = llCalcGross + llTempProject(il12X) \ 100
                                    llTempProject(il12X) = 0
                                Next il12X
                                tmLnYrTot(ilYearInx).lTotalYear = llCalcGross
                            End If
                        End If
                    Next ilYearInx                  'all years completed for this line if spanning the corp or std 52 weeks


                    'Test for selective offices and split the gross $ projected
                    ilFoundOne = True
                    'Process for the split offices (max 10)
                    If tgChfAC.lComm(0) = 0 Then              'only 1 slsp must be 100%
                        tgChfAC.lComm(0) = 1000000
                        tgChfAC.iMnfSubCmpy(0) = 0            '4-20-00
                    End If

                    ReDim llSlfSplit(0 To 9) As Long           '4-20-00 slsp rev share %
                    ReDim ilSlfCode(0 To 9) As Integer         '4-20-00
                    ReDim llSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)

                    ilMnfSubCo = gGetSubCmpy(tgChfAC, ilSlfCode(), llSlfSplit(), tmClf.iVefCode, False, llSlfSplitRev())

                    For ilSlspLoop = 0 To 9
                        '4-20-00 If tgChfAC.islfCode(ilSlspLoop) > 0 Then
                        If ilSlfCode(ilSlspLoop) > 0 Then       '4-20-00
                            'find the associated office in memory table from the slsp
                            For ilLoop = 0 To UBound(tlSofList)         'loop thru the slsp codes to find the associated office
                                '4-20-00 If tgChfAC.islfCode(ilSlspLoop) = tlSofList(ilLoop).imnfSsCode Then       'find matching slsp code in memory table
                                If ilSlfCode(ilSlspLoop) = tlSofList(ilLoop).iMnfSSCode Then       'find matching slsp code in memory table
                                    ilSaveSof = tlSofList(ilLoop).iSofCode
                                    If ilBySlsp Then                    'by slsp, check the selection
                                        For ilTemp = 0 To RptSelAc!lbcSelection(0).ListCount - 1 Step 1    'loop to see which office has been selected
                                            If RptSelAc!lbcSelection(0).Selected(ilTemp) Then              'selected office?
                                                slNameCode = tgSalesperson(ilTemp).sKey                      'pick up office code
                                                ilRet = gParseItem(slNameCode, 2, "\", slCode)

                                                '4-20-00 If Val(slCode) = tgChfAC.islfCode(ilSlspLoop) Then
                                                If Val(slCode) = ilSlfCode(ilSlspLoop) Then
                                                    tmGrf.iSofCode = ilSaveSof
                                                    'also, always write records withs splits
                                                    '4-20-00 If tgChfAC.islfCode(ilSlspLoop) > 0 Then            'splits involved, determine split amounts
                                                     If ilSlfCode(ilSlspLoop) > 0 Then
                                                        mSlspSplits ilSlspLoop, ilNoYears, tmLnYrTot(), tmSlspTot(), ilSlfCode(), llSlfSplit()     '4-20-00
                                                    End If
                                                End If                      'val(slcode) = tlsoflist(illoop).isofcode
                                            End If                          'RptSelAc!lbcselection.selected
                                        Next ilTemp
                                    Else                                    'option by vehicle; line has already been excluded
                                        tmGrf.iSofCode = ilSaveSof
                                        'also, always write records withs splits
                                        '4-20-00 If tgChfAC.islfCode(ilSlspLoop) > 0 Then            'splits involved, determine split amounts
                                        If ilSlfCode(ilSlspLoop) > 0 Then            'splits involved, determine split amounts
                                            mSlspSplits ilSlspLoop, ilNoYears, tmLnYrTot(), tmSlspTot(), ilSlfCode(), llSlfSplit()  '4-20-00
                                        End If
                                    End If

                                End If
                            Next ilLoop
                        Else
                            ilFoundOne = False
                            Exit For                    'no more offices, exit the loop
                        End If
                    Next ilSlspLoop                             'next slsp split for same contract/line
                End If                                  'ilFoundVeh
            Next ilClf                                      'loop thru schedule lines
        End If                                              'ilfoundone
    Next ilCurrentRecd                                      'loop for CHF records

    'all contrcts have been processed for all years requested.  Write all records to disk,
    'one per slsp/vehicle/year
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)
    'tmGrf.iGenTime(1) = igNowTime(1)
    tmGrf.lGenTime = lgNowTime
    If ilCalType = 1 Then
        tmGrf.sDateType = "S"
    Else
        tmGrf.sDateType = "C"
    End If
    'tmGrf.iPerGenl(2) = ilTYStartYr             'base year requested (reqd for report heading)
    'tmGrf.iPerGenl(3) = ilStartMonth            'required for report month column headings
    tmGrf.iPerGenl(1) = ilTYStartYr             'base year requested (reqd for report heading)
    tmGrf.iPerGenl(2) = ilStartMonth            'required for report month column headings
    'tmGrf.lDollars(17) = 0                      'init yearly total
    'tmGrf.lDollars(13) = 0                      'init qtr totals
    'tmGrf.lDollars(14) = 0
    'tmGrf.lDollars(15) = 0
    'tmGrf.lDollars(16) = 0
    
    tmGrf.lDollars(16) = 0                      'init yearly total
    tmGrf.lDollars(12) = 0                      'init qtr totals
    tmGrf.lDollars(13) = 0
    tmGrf.lDollars(14) = 0
    tmGrf.lDollars(15) = 0
    For ilLoop = LBound(tmSlspTot) To UBound(tmSlspTot) - 1 Step 1
        'tmGrf.iPerGenl(1) = tmSlspTot(ilLoop).iIndex            'Year index , 1-5 base year = 1, base year minus 1 = 2, base year minus 2 =  3, etc
        tmGrf.iPerGenl(0) = tmSlspTot(ilLoop).iIndex            'Year index , 1-5 base year = 1, base year minus 1 = 2, base year minus 2 =  3, etc
        tmGrf.iSlfCode = tmSlspTot(ilLoop).iSlfCode
        tmGrf.iVefCode = tmSlspTot(ilLoop).iVefCode
        tmGrf.iSofCode = tmSlspTot(ilLoop).iSofCode
        For il12X = 1 To 12
            tmGrf.lDollars(il12X - 1) = tmSlspTot(ilLoop).lDollars(il12X)
            tmGrf.lDollars(16) = tmGrf.lDollars(16) + tmGrf.lDollars(il12X - 1)   'accumulate yearly total
            'Accumulate quarterly totals
            ilTemp = (il12X - 1) \ 3 + 1
            tmGrf.lDollars(12 + ilTemp - 1) = tmGrf.lDollars(12 + ilTemp - 1) + tmGrf.lDollars(il12X - 1)
        Next il12X
        tmGrf.iYear = ilTYStartYr - tmSlspTot(ilLoop).iIndex + 1               'actual year (1999, 1998, etc)
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        For il12X = 1 To 17                         'init $ buckets for next record
            tmGrf.lDollars(il12X - 1) = 0
        Next il12X

        'If tmGrf.iPerGenl(1) <> 1 Then                    'not the base year, see if theres an entry associated
        If tmGrf.iPerGenl(0) <> 1 Then                    'not the base year, see if theres an entry associated
                                                    'with this slp/veh in base year to compare against.  If not,
                                                    'create one with zeros
            ilFoundOne = False
            'For ilSlspLoop = 1 To UBound(tmBaseYear) - 1 Step 1
            For ilSlspLoop = LBound(tmBaseYear) To UBound(tmBaseYear) - 1 Step 1
                If tmBaseYear(ilSlspLoop).iVefCode = tmGrf.iVefCode And tmBaseYear(ilSlspLoop).iSofCode = tmGrf.iSofCode And tmBaseYear(ilSlspLoop).iSlfCode = tmGrf.iSlfCode Then
                    ilFoundOne = True
                    Exit For
                End If
            Next ilSlspLoop
            If Not ilFoundOne Then
                ilUpper = UBound(tmBaseYear)
                tmBaseYear(ilUpper).iVefCode = tmGrf.iVefCode
                tmBaseYear(ilUpper).iSofCode = tmGrf.iSofCode
                tmBaseYear(ilUpper).iSlfCode = tmGrf.iSlfCode
                ReDim Preserve tmBaseYear(LBound(tmBaseYear) To ilUpper + 1)

                'tmGrf.iPerGenl(1) = 1                             'base year index
                tmGrf.iPerGenl(0) = 1                             'base year index
                tmGrf.iYear = ilTYStartYr                         'base year (1999, 2000, etc)
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            End If
        End If
    Next ilLoop
    Erase tlChfAdvtExt, tlSofList, tmSlspTot, tmLnYrTot, tmBaseYear
    'Erase llTempProject, llProject, llTempStarts, llAllStarts
    Erase llTempProject, llTempStarts, llAllStarts
    Erase llSlfSplit, ilSlfCode
    Erase tgClfAC, tgCffAC
    sgCntrForDateStamp = ""
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmUrf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmSlf)
End Sub
'
'
'               mSlspsplits - determine all salesperson splits offices totals
'                               for all years gathered for the line (most cases
'                               should only be 1 or 2 years at most)
'                               Another table (tmBaseYear) built of all office/slf/veh
'                               that exists for the base year.  Required later when
'                               writing records todisk for the other years.  There must
'                               be an entry for the base year for all the other comparison years.
'               <input> ilSlspLoop - index to slsp (1 of 10)
'                       ilNoYears - # years to process (user entered), add 1 for the base year
'                       tlLnYrTot() - array of max 5 records containing each years
'                                       slsp, veh, and yearly $ for 1 sch line
'                       ilSlfCode() - slsp array to process for the matching vehicles sub-compny
'                       llSlfSplit() - slsp share of split
'
'               <output> tlSlsptot() - array of records containing totals for each
'                                   vehicle and slsp for a given year
'
'           4-20-00 implement new structure for slsp share & comm
Sub mSlspSplits(ilSlspLoop As Integer, ilNoYears As Integer, tlLnYrTot() As PRICELIST, tlSlspTot() As PRICELIST, ilSlfCode() As Integer, llSlfSplit() As Long)
Dim ilYearInx As Integer
Dim il12X As Integer
Dim slAmount As String
Dim slSharePct As String
Dim slTemp As String
'ReDim llProject(1 To 12) As Long
ReDim llProject(0 To 12) As Long    'Index zero ignored
Dim ilFndVehSlsp As Integer
Dim ilVehSlsp As Integer
Dim ilLoop As Integer
Dim ilUpper As Integer
    For ilYearInx = 1 To ilNoYears + 1
        If tlLnYrTot(ilYearInx).lTotalYear > 0 Then     'dont do anything if the year is zero
            For il12X = 1 To 12
                slAmount = gLongToStrDec(tlLnYrTot(ilYearInx).lDollars(il12X), 0)            'cents have already been removed
                '4-20-00 slSharePct = gLongToStrDec(tgChfAC.lComm(ilSlspLoop), 4)                    'slsp split share in %
                slSharePct = gLongToStrDec(llSlfSplit(ilSlspLoop), 4)                    '4-20-00 slsp split share in %
                slTemp = gDivStr(gMulStr(slSharePct, slAmount), "100")         'slsp gross portion of possible split
                slTemp = gRoundStr(slTemp, "01.", 0)
                llProject(il12X) = Val(slTemp)    'no cents
            Next il12X
            If ilYearInx = 1 Then           'Base year, build table of all slsp/office/veh for the base year so that
                                            'Before writing all entries to disk, all the comparison years must have
                                            'an entry for the base year to compare against so Crystal will not
                                            'have any problems
                ilFndVehSlsp = False
                For ilLoop = LBound(tmBaseYear) To UBound(tmBaseYear) - 1 Step 1
                    '4-20-00 If tmBaseYear(ilLoop).ivefCode = tmClf.ivefCode And tmBaseYear(ilLoop).islfCode = tgChfAC.islfCode(ilSlspLoop) And tmBaseYear(ilLoop).isofCode = tmGrf.isofCode Then
                    If tmBaseYear(ilLoop).iVefCode = tmClf.iVefCode And tmBaseYear(ilLoop).iSlfCode = ilSlfCode(ilSlspLoop) And tmBaseYear(ilLoop).iSofCode = tmGrf.iSofCode Then
                        ilFndVehSlsp = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFndVehSlsp Then
                    ilUpper = UBound(tmBaseYear)
                    tmBaseYear(ilUpper).iVefCode = tmClf.iVefCode
                    tmBaseYear(ilUpper).iSofCode = tmGrf.iSofCode
                    tmBaseYear(ilUpper).iSlfCode = tgChfAC.iSlfCode(ilSlspLoop)
                    ReDim Preserve tmBaseYear(LBound(tmBaseYear) To UBound(tmBaseYear) + 1)
                End If
            End If
            'Year has been split for this slsp, build it into memory
            'Add to existing entry if already there; otherwise create new entry
            ilFndVehSlsp = False
            For ilVehSlsp = LBound(tlSlspTot) To UBound(tlSlspTot) - 1         'determine if this slsp, veh and year has been built in memory yet
                '4-20-00 If tlSlspTot(ilVehSlsp).ivefCode = tmClf.ivefCode And tlSlspTot(ilVehSlsp).islfCode = tgChfAC.islfCode(ilSlspLoop) And tlSlspTot(ilVehSlsp).iIndex = ilYearInx Then
                If tlSlspTot(ilVehSlsp).iVefCode = tmClf.iVefCode And tlSlspTot(ilVehSlsp).iSlfCode = ilSlfCode(ilSlspLoop) And tlSlspTot(ilVehSlsp).iIndex = ilYearInx Then
                    'Total the $ for this slsp, vehicle and year
                    For il12X = 1 To 12
                        tlSlspTot(ilVehSlsp).lDollars(il12X) = tlSlspTot(ilVehSlsp).lDollars(il12X) + llProject(il12X)   'adjusted  $ split share
                    Next il12X
                    ilFndVehSlsp = True
                    Exit For
                End If
            Next ilVehSlsp
            If Not ilFndVehSlsp Then
                ilUpper = UBound(tlSlspTot)
                tlSlspTot(ilUpper).iVefCode = tmClf.iVefCode
                tlSlspTot(ilUpper).iSlfCode = tgChfAC.iSlfCode(ilSlspLoop)
                tlSlspTot(ilUpper).iSofCode = tmGrf.iSofCode
                tlSlspTot(ilUpper).iIndex = ilYearInx
                For il12X = 1 To 12
                    tlSlspTot(UBound(tlSlspTot)).lDollars(il12X) = llProject(il12X)    'adjusted  $ split share
                Next il12X
                ReDim Preserve tlSlspTot(0 To ilUpper + 1) As PRICELIST
            End If
        End If
    Next ilYearInx
End Sub
