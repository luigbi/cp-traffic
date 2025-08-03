Attribute VB_Name = "RptCrAvgComp"
Option Explicit
Option Compare Text


Dim hmCHF As Integer                'Contract header file handle
Dim imCHFRecLen As Integer          'CHF record length
Dim tmChf As CHF
Dim tmChfSrchKey1 As CHFKEY1
Dim tmChfAdvtExt() As CHFADVTEXT
Dim hmClf As Integer                'Contract line file handle
Dim imClfRecLen As Integer          'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer                'Contract flight file handle
Dim imCffRecLen As Integer          'CFF record length
Dim tmCff As CFF
Dim hmAgf As Integer                'AGency file handle
Dim imAgfRecLen As Integer          'record length
Dim tmAgf As AGF
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim tmVef As VEF
Dim hmVef As Integer                'Vehicles code handle
Dim imGrfRecLen As Integer          'GPF record length
Dim tmCntTypes As CNTTYPES          'image for user requested parameters
Dim lmSingleCntr As Long
Dim imInclAdvtCodes As Integer
Dim imUseAdvtCodes() As Integer
Dim imInclVefCodes As Integer
Dim imUsevefcodes() As Integer
Dim imInclVGCodes As Integer
Dim imUseVGCodes() As Integer
Dim imInclSlspCodes As Integer
Dim imUseSlspCodes() As Integer
Dim imMajorSet As Integer           'vehicle group selected for majort sort, could be NONE (0)
Dim imWhichLine As Integer          'package or airing lines

Type RATEPRICEDATES
    sYearStart As String
    sYearEnd As String
    lYearStart As Long
    lYearEnd As Long
    iYearStart(0 To 1) As Integer
    iYearEnd(0 To 1) As Integer
End Type
Dim tmRatePriceDates() As RATEPRICEDATES

Public Type COMPARESTATS
    sKey As String * 5
    iVefCode As Integer
    dRates(0 To 6) As Double
    dPropPrice(0 To 6) As Double
    dSpots(0 To 6) As Double
    dRates60s(0 To 6) As Double
    dPropPrice60s(0 To 6) As Double
    dSpots60s(0 To 6) As Double
End Type
Public tmCompareStats() As COMPARESTATS

'  Receivables File
'********************************************************************************************
'
'           mCreateAverageCompare - Prepass to create Average 30" Rate / Spot Price Comparison report
'           Create prepass to produce a report to compare acquisition costs to net revenue
'           to arrive at a margin percent to determine if a contract and/or vehicle
'           is profitable or not.  Margin calculation is based on cost (expanded acq)
'           divided by revenue (net $)
'
'
'
'
'********************************************************************************************
Sub mCreateAverageCompare()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilWhichset                                                                            *
'******************************************************************************************

Dim llRet As Long                    '
Dim llRet60s As Long                 '
Dim ilClf As Integer                    'loop for schedule lines
Dim ilHOState As Integer                'retrieve only latest order or revision
Dim slCntrTypes As String               'retrieve remnants, PI, DR, etc
Dim slCntrStatus As String              'retrieve H, O G or N (holds, orders, unsch hlds & orders)
Dim ilCurrentRecd As Integer            'loop for processing last years contracts
Dim blFoundOne As Boolean               'Found a matching  office built into mem
Dim ilValidVehicle As Integer
Dim llTemp As Long
Dim llDate As Long                      'temp date variable
Dim llDate2 As Long
Dim llLineStartDate As Long
Dim llLineEndDate As Long
Dim dlActTotal As Double
Dim dlPropTotal As Double
Dim llSpotsTotal As Long

'Date used to gather information
'String formats for generalized date conversions routines
'Long formats for testing
'Packed formats to store in GRF record
Dim slWeekStart As String               'start date of week for this years new business entered this week
Dim llWeekStart As Long                 'start date of week for this years new business entered on te user entered week
Dim ilWhichRate As Integer
Dim ilWeekOrMonth As Integer
Dim llLineGrossWithoutAcq As Long
Dim ilAgyCommPct As Integer
Dim slCashAgyComm As String
Dim ilIndex As Integer
Dim slYearStart As String
Dim slYearEnd As String
Dim ilCounter As Integer
Dim slChfWeekStart As String
Dim slChfWeekEnd As String
Dim ilRatePrice As Integer               '0=AVG Rate, 1=AVG Spot Price
Dim slCode As String
Dim slKey As String
Dim tmGrf As GRF
Dim slSpotLen As String                 'Combine (all spot lenghts) or Separate (30s / 60s spot lenghts only)
Dim slAvgBy As String                   'Spot Price Comparison: Separate (30s/60s spots) or Combined (all spot lenghts)
Dim slShowUnitPrice As String           'Show Percentage or Unit Price in CR for Avg Rate Comparison

Dim dlDollars_0 As Double
Dim dlDollars_1 As Double
Dim dlDollars_2 As Double
Dim dlDollars_3 As Double
Dim dlDollars_4 As Double
Dim dlDollars_5 As Double
Dim dlDollars_6 As Double
Dim dlDollars_7 As Double
Dim dlDollars_8 As Double
Dim dlDollars_9 As Double
Dim dlDollars_15 As Double
Dim dlDollars_16 As Double
Dim dlDollars_Total As Double
Dim dlDollarsProp_Total As Double
Dim llCount_Total As Long

Dim dlDollarsProp_0 As Double
Dim dlDollarsProp_1 As Double
Dim dlDollarsProp_2 As Double
Dim dlDollarsProp_3 As Double
Dim dlDollarsProp_4 As Double
Dim dlDollarsProp_5 As Double
Dim dlDollarsProp_6 As Double
Dim dlDollarsProp_7 As Double
Dim dlDollarsProp_8 As Double
Dim dlDollarsProp_9 As Double
Dim dlDollarsProp_16 As Double

Dim llCountDollars_0 As Long
Dim llCountDollars_1 As Long
Dim llCountDollars_2 As Long
Dim llCountDollars_3 As Long
Dim llCountDollars_4 As Long
Dim llCountDollars_5 As Long
Dim llCountDollars_6 As Long
Dim llCountDollars_7 As Long
Dim llCountDollars_8 As Long
Dim llCountDollars_9 As Long
Dim llCountDollars_15 As Long
Dim llCountDollars_16 As Long

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    llRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If llRet <> BTRV_ERR_NONE Then
        llRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
 
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    llRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If llRet <> BTRV_ERR_NONE Then
        llRet = btrClose(hmClf)
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    llRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If llRet <> BTRV_ERR_NONE Then
        llRet = btrClose(hmCff)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    llRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If llRet <> BTRV_ERR_NONE Then
        llRet = btrClose(hmCHF)
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)
    
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    llRet = btrOpen(hmAgf, "", sgDBPath & "AGf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If llRet <> BTRV_ERR_NONE Then
        llRet = btrClose(hmAgf)
        btrDestroy hmAgf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)
    
    hmVef = CBtrvTable(ONEHANDLE)
    llRet = btrOpen(hmVef, "", sgDBPath & "VEF.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If llRet <> BTRV_ERR_NONE Then
        llRet = btrClose(hmVef)
        btrDestroy hmVef
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAgfRecLen = Len(tmVef)
    
    ReDim tgClf(0 To 0) As CLFLIST
    tgClf(0).iStatus = -1 'Not Used
    tgClf(0).lRecPos = 0
    tgClf(0).iFirstCff = -1
    ReDim tgCff(0 To 0) As CFFLIST
    tgCff(0).iStatus = -1 'Not Used
    tgCff(0).lRecPos = 0
    tgCff(0).iNextCff = -1
    
    mObtainSelectivityAVG
    
    ReDim tmRatePriceDates(0 To CInt(Trim(RptSelAvgCmp!edcYears.Text)))
    ReDim tmCompareStats(0 To 0) As COMPARESTATS
    ReDim llStartDates(0 To CInt(Trim(RptSelAvgCmp!edcYears.Text) + 1)) As Long     'start dates of each period to gather (only one). Index zero ignored
    
    ReDim dlProject(0 To CInt(Trim(RptSelAvgCmp!edcYears.Text))) As Double          '$ projected for 1 period ("X" weeks). Index zero ignored
    ReDim dlProjectSpots(0 To CInt(Trim(RptSelAvgCmp!edcYears.Text))) As Double     'spot counts for 1 period ("X" weeks). Index zero ignored
    ReDim dlProjectProp(0 To CInt(Trim(RptSelAvgCmp!edcYears.Text))) As Double
    ReDim dlProject60s(0 To CInt(Trim(RptSelAvgCmp!edcYears.Text))) As Double       '$ projected for 60s spots
    ReDim dlProjectSpots60s(0 To CInt(Trim(RptSelAvgCmp!edcYears.Text))) As Double  'spot counts for 60s
    ReDim dlProjectProp60s(0 To CInt(Trim(RptSelAvgCmp!edcYears.Text))) As Double   'project proposoal for 60s spots
    
    ilRatePrice = IIF(RptSelAvgCmp!rbcAvgRatePrice.Item(0).Value = True, 0, 1)      '0=AVG Rate Comparison Report, 1=AVG Spot Price Comparison Report
    
    slShowUnitPrice = "Price"
    If RptSelAvgCmp!ckcShowUnitPrice.Visible = True Then
        slShowUnitPrice = IIF(RptSelAvgCmp!ckcShowUnitPrice.Value = 0, "Percent", "Price")  'For Avg Rate Comparison: show Percentage or Unit Price in report
    End If
    
    'Get Start and end dates of user requested date to find all advertisers airing
    For ilCounter = 0 To (UBound(tmRatePriceDates) - 1)
        slYearEnd = CDate("12/15/" & RptSelAvgCmp!edcStartYear.Text)
        slYearStart = CDate("1/15/" & str$(Year(slYearEnd) - Trim(RptSelAvgCmp!edcYears.Text) + 1))
        
        slYearStart = DateAdd("yyyy", ilCounter, slYearStart)
        tmRatePriceDates(ilCounter).sYearStart = gObtainYearStartDate(0, slYearStart)
        slYearEnd = CDate("12/15/" & Year(slYearStart))
        tmRatePriceDates(ilCounter).sYearEnd = gObtainYearEndDate(0, slYearEnd)
        tmRatePriceDates(ilCounter).lYearStart = gDateValue(DateAdd("yyyy", ilCounter, tmRatePriceDates(ilCounter).sYearStart))
        tmRatePriceDates(ilCounter).lYearEnd = gDateValue(DateAdd("yyyy", ilCounter, tmRatePriceDates(ilCounter).sYearEnd))
        gPackDate tmRatePriceDates(ilCounter).sYearStart, tmRatePriceDates(ilCounter).iYearStart(0), tmRatePriceDates(ilCounter).iYearStart(1)
        gPackDate tmRatePriceDates(ilCounter).sYearEnd, tmRatePriceDates(ilCounter).iYearEnd(0), tmRatePriceDates(ilCounter).iYearEnd(1)
    Next ilCounter
    
    'Start date of each period to accumulate spot counts (only one period applicable)
    For ilCounter = 0 To UBound(tmRatePriceDates)
        If ilCounter < UBound(tmRatePriceDates) Then
            llStartDates(ilCounter + 1) = gDateValue(tmRatePriceDates(ilCounter).sYearStart)
        Else
            llStartDates(ilCounter + 1) = gDateValue(gObtainYearStartDate(0, DateAdd("yyyy", 1, CDate("1/15/" & str$(Year(tmRatePriceDates(ilCounter - 1).sYearEnd))))))
        End If
    Next ilCounter
    
    'populate tmCompareStats with selected vehicles
    'create sKey and vehicle code
    For ilCounter = 0 To RptSelAvgCmp!lbcSelection(0).ListCount - 1 Step 1
        slKey = tgVehicle(ilCounter).sKey
        llRet = gParseItem(slKey, 2, "\", slCode)
        If RptSelAvgCmp!lbcSelection(0).Selected(ilCounter) Then               'selected ?
            tmCompareStats(UBound(tmCompareStats)).sKey = Format(slCode, "00000")
            tmCompareStats(UBound(tmCompareStats)).iVefCode = Val(slCode)
            ReDim Preserve tmCompareStats(0 To UBound(tmCompareStats) + 1)
        End If
    Next ilCounter
    
    'sort array
    If (UBound(tmCompareStats) - 1 > 0) Then
        ArraySortTyp fnAV(tmCompareStats(), 0), UBound(tmCompareStats), 0, LenB(tmCompareStats(0)), 0, LenB(tmCompareStats(0).sKey), 0
    End If

    If lmSingleCntr > 0 Then
        ReDim tmChfAdvtExt(0 To 1) As CHFADVTEXT
        tmChfSrchKey1.lCntrNo = lmSingleCntr
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        llRet = btrGetGreaterOrEqual(hmCHF, tmChf, Len(tmChf), tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If llRet <> BTRV_ERR_NONE Then
           ReDim tmChfAdvtExt(0 To 0) As CHFADVTEXT
        Else
            'setup 1 entry in the active contract array for processing single contract
            tmChfAdvtExt(0).lCntrNo = tmChf.lCntrNo
            tmChfAdvtExt(0).lCode = tmChf.lCode
            tmChfAdvtExt(0).iSlfCode(0) = tmChf.iSlfCode(0)
            tmChfAdvtExt(0).iAdfCode = tmChf.iAdfCode
        End If
    Else
        'Gather all contracts for previous year and current year whose effective date entered
        'is prior to the effective date that affects either previous year or current year
        slCntrTypes = gBuildCntTypesForAll()
        slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order, proposals working, complete, unapproved
        
        If tmCntTypes.iComplete = True Then
            slCntrStatus = slCntrStatus & "C"
        End If
        If tmCntTypes.iIncomplete = True Then
            slCntrStatus = slCntrStatus & "I"
        End If
        If tmCntTypes.iWorking = True Then
            slCntrStatus = slCntrStatus & "W"
        End If
        
        ilHOState = 2                       'H or O or G or N or W or C or I (if G or N or W or C or I exists show it over H or O)
        
        llRet = gObtainCntrForDate(RptSelAvgCmp, tmRatePriceDates(LBound(tmRatePriceDates)).sYearStart, tmRatePriceDates(UBound(tmRatePriceDates) - 1).sYearEnd, _
                slCntrStatus, slCntrTypes, ilHOState, tmChfAdvtExt())
    End If

    'common GRF fields that wont change
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    
    'All contracts have been retrieved; 5 years max
    For ilCurrentRecd = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1

        gUnpackDate tmChfAdvtExt(ilCurrentRecd).iStartDate(0), tmChfAdvtExt(ilCurrentRecd).iStartDate(1), slChfWeekStart
        gUnpackDate tmChfAdvtExt(ilCurrentRecd).iEndDate(0), tmChfAdvtExt(ilCurrentRecd).iEndDate(1), slChfWeekEnd
        
        blFoundOne = mFilterSelectivityAVG(ilCurrentRecd)
        If blFoundOne Then
            'determine if the contracts start & end dates fall within the requested period
            gUnpackDateLong tgChf.iEndDate(0), tgChf.iEndDate(1), llDate2      'hdr end date converted to long
            gUnpackDateLong tgChf.iStartDate(0), tgChf.iStartDate(1), llDate    'hdr start date converted to long
            
            tmGrf.lChfCode = tgChf.lCode
            
            ilAgyCommPct = 0      'direct, no comm
            If tgChf.iAgfCode > 0 Then
                ilIndex = gBinarySearchAgf(tgChf.iAgfCode)
                If ilIndex >= 0 Then
                     ilAgyCommPct = tgCommAgf(ilIndex).iCommPct
                End If
            End If
            slCashAgyComm = gIntToStrDec(ilAgyCommPct, 2)
            If tgChf.iPctTrade = 100 And tgChf.sAgyCTrade = "N" Then
                slCashAgyComm = ".00"
            End If

            For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                tmClf = tgClf(ilClf).ClfRec
                'Project the monthly spots from the flights (standard / hidden)
                If ((imWhichLine = 1) And (tmClf.sType = "S" Or tmClf.sType = "H")) Or ((imWhichLine = 0) And (tmClf.sType = "S" Or tmClf.sType = "O" Or tmClf.sType = "A")) Then
                
                    ilValidVehicle = True
                    ilValidVehicle = gFilterLists(tmClf.iVefCode, imInclVefCodes, imUsevefcodes())
                    
                    ilIndex = gBinarySearchVef(tmClf.iVefCode)
                    If ilIndex > 0 Then
                        If (tgMVef(ilIndex).sType = "R") Then
                            'include REP vehicles
                            If (tmCntTypes.iRep = True) Then ilValidVehicle = True Else ilValidVehicle = False
                        Else            'If (tgMVef(ilIndex).sType <> "R") Then
                            'include non-REP vehicles with AIRTIME spots
                            If (tmCntTypes.iAirTime = True) Then ilValidVehicle = True Else ilValidVehicle = False
                        End If
                    End If
                    
                    'Spot Lenghts: ALL or 30s/60s spots only
                    slSpotLen = IIF(RptSelAvgCmp!rbcSpotLen(0).Value = True, "All", "3060Only")
                    'For AVG Spot Price Report: Combine all spot lenghts or Separate 30s/60s spot lengths
                    slAvgBy = IIF(RptSelAvgCmp!rbcAvgBy(0).Value = True, "Separate", "Combined")
                    
                    ilWhichRate = 0             'use true line rate
                    ilWeekOrMonth = 1           'assume month gathering since it tests date spans; wkly version calculates a week index
                    llLineGrossWithoutAcq = 0
                    If ilValidVehicle Then      'got a valid vehicle group code
                        gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llLineEndDate     'line end date
                        gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llLineStartDate    'Line start date

                        'build llProject array (for $), llProjects array (for # of spots), dlProjectProp (proposal)
                        mBuildFlightSpotRevPropPrice ilClf, llStartDates(), 1, UBound(tmRatePriceDates), dlProject(), dlProjectSpots(), 1, 0, tgClf(), tgCff(), dlProjectProp(), _
                            ilRatePrice, slAvgBy, slSpotLen, dlProject60s(), dlProjectSpots60s(), dlProjectProp60s(), "G"
                            
                        llRet = 0: llRet60s = 0
                        For ilCounter = 0 To UBound(dlProjectSpots)
                            llRet = llRet + dlProjectSpots(ilCounter)
                        Next ilCounter
                        For ilCounter = 0 To UBound(dlProjectSpots60s)
                            llRet60s = llRet60s + dlProjectSpots60s(ilCounter)
                        Next ilCounter

                        If (llRet > 0) Or (llRet60s > 0) Then
                            'loop through the vehicles (spots, prop/actual dollars for 5 years)
                            'do binary search to match on vehicle code
                            ilIndex = mBinarySearchAVG(tmClf.iVefCode, tmCompareStats())
                            If ilIndex >= 0 Then
                                For ilCounter = 0 To UBound(tmRatePriceDates) - 1
                                    If (llRet > 0) Then
                                        'for Avg Spot Price comparison for ALL Spots or Avg 30s Rate Comparison
                                        tmCompareStats(ilIndex).dRates(ilCounter) = tmCompareStats(ilIndex).dRates(ilCounter) + (dlProject(ilCounter + 1))
                                        tmCompareStats(ilIndex).dPropPrice(ilCounter) = tmCompareStats(ilIndex).dPropPrice(ilCounter) + (dlProjectProp(ilCounter + 1))
                                        tmCompareStats(ilIndex).dSpots(ilCounter) = tmCompareStats(ilIndex).dSpots(ilCounter) + dlProjectSpots(ilCounter + 1)
                                    End If
                                    If (llRet60s > 0) Then 'for 30s vs 60s Spots Comparison
                                        tmCompareStats(ilIndex).dRates60s(ilCounter) = tmCompareStats(ilIndex).dRates60s(ilCounter) + (dlProject60s(ilCounter + 1))
                                        tmCompareStats(ilIndex).dPropPrice60s(ilCounter) = tmCompareStats(ilIndex).dPropPrice60s(ilCounter) + (dlProjectProp60s(ilCounter + 1))
                                        tmCompareStats(ilIndex).dSpots60s(ilCounter) = tmCompareStats(ilIndex).dSpots60s(ilCounter) + dlProjectSpots60s(ilCounter + 1)
                                    End If
                                Next ilCounter
                            End If
                            For ilCounter = 0 To UBound(tmRatePriceDates)
                                dlProject(ilCounter) = 0            'init for next schedule line
                                dlProjectSpots(ilCounter) = 0
                                dlProjectProp(ilCounter) = 0
                                dlProject60s(ilCounter) = 0
                                dlProjectSpots60s(ilCounter) = 0
                                dlProjectProp60s(ilCounter) = 0
                            Next ilCounter
                        End If
                    End If
                End If
            Next ilClf                      'loop thru schedule lines
        End If                              'blFoundOne = true
    Next ilCurrentRecd                      'loop for CHF records
    
    '       Grf parameters:
    '       grfGenDAte - generation date (key)
    '       grfGenTime - generation time (key)
    '       grfvefcode - vehicle code
    '       GrfPer1 - Rates 1 for ALL or 30s spots
    '       grfPer2 - Rates 2 for ALL or 30s spots
    '       grfPer3 - Rates 3 for ALL or 30s spots
    '       grfPer4 - Rates 4 for ALL or 30s spots
    '       grfPer5 - Rates 5 for ALL or 30s spots
    '       GrfPer6 - Rates 1 for 60s spots
    '       grfPer7 - Spots 2 for 60s spots
    '       grfPer8 - Spots 3 for 60s spots
    '       grfPer9 - Spots 4 for 60s spots
    '       grfPer10 - Spots 5 for 60s spots
    '       grfPer16 - Total Average Spot Price (ALL or 30s spots)
    '       grfPer17 - Total Average Spot Price (60s spots)
    '       grfPer1Genl - Year 1 for ALL or 30s spots
    '       grfPer2Genl - Year 2 for ALL or 30s spots
    '       grfPer3Genl - Year 3 for ALL or 30s spots
    '       grfPer4Genl - Year 4 for ALL or 30s spots
    '       grfPer5Genl - Year 5 for ALL or 30s spots
    '       grfPer6Genl - Year 1 for 60s spots
    '       grfPer7Genl - Year 2 for 60s spots
    '       grfPer8Genl - Year 3 for 60s spots
    '       grfPer9Genl - Year 4 for 60s spots
    '       grfPer10Genl - Year 5 for 60s spots
    
    dlDollars_0 = 0: dlDollars_1 = 0: dlDollars_2 = 0: dlDollars_3 = 0: dlDollars_4 = 0: dlDollars_5 = 0
    dlDollars_6 = 0: dlDollars_7 = 0: dlDollars_8 = 0: dlDollars_9 = 0: dlDollars_15 = 0: dlDollars_16 = 0:
    
    llCountDollars_0 = 0: llCountDollars_1 = 0: llCountDollars_2 = 0: llCountDollars_3 = 0: llCountDollars_4 = 0: llCountDollars_5 = 0
    llCountDollars_6 = 0: llCountDollars_7 = 0: llCountDollars_8 = 0: llCountDollars_9 = 0: llCountDollars_15 = 0: llCountDollars_16 = 0
    
    For ilIndex = LBound(tmCompareStats) To UBound(tmCompareStats) - 1
    
        tmGrf.iVefCode = tmCompareStats(ilIndex).iVefCode
        tmGrf.iAdfCode = 0
            
        tmGrf.lDollars(0) = 0: tmGrf.lDollars(1) = 0: tmGrf.lDollars(2) = 0: tmGrf.lDollars(3) = 0: tmGrf.lDollars(4) = 0
        tmGrf.lDollars(5) = 0: tmGrf.lDollars(6) = 0: tmGrf.lDollars(7) = 0: tmGrf.lDollars(8) = 0: tmGrf.lDollars(9) = 0

        'ALL Spots Price Comparison or Average Rate Comparison
        If tmCompareStats(ilIndex).dSpots(0) > 0 Then tmGrf.lDollars(0) = ((tmCompareStats(ilIndex).dRates(0) / 100) / tmCompareStats(ilIndex).dSpots(0))
        If tmCompareStats(ilIndex).dSpots(1) > 0 Then tmGrf.lDollars(1) = ((tmCompareStats(ilIndex).dRates(1) / 100) / tmCompareStats(ilIndex).dSpots(1))
        If tmCompareStats(ilIndex).dSpots(2) > 0 Then tmGrf.lDollars(2) = ((tmCompareStats(ilIndex).dRates(2) / 100) / tmCompareStats(ilIndex).dSpots(2))
        If tmCompareStats(ilIndex).dSpots(3) > 0 Then tmGrf.lDollars(3) = ((tmCompareStats(ilIndex).dRates(3) / 100) / tmCompareStats(ilIndex).dSpots(3))
        If tmCompareStats(ilIndex).dSpots(4) > 0 Then tmGrf.lDollars(4) = ((tmCompareStats(ilIndex).dRates(4) / 100) / tmCompareStats(ilIndex).dSpots(4))
        
        If ilRatePrice = 1 Then
            'Avg Spot Price Comparison: Separate 30s v 60s
            If (RptSelAvgCmp!rbcAvgBy(0).Value = True) Then
                'average price calculation: spots * rate / spots for lengths 30s or 60s
                If tmCompareStats(ilIndex).dSpots60s(0) > 0 Then tmGrf.lDollars(5) = ((tmCompareStats(ilIndex).dRates60s(0) / 100) / tmCompareStats(ilIndex).dSpots60s(0))
                If tmCompareStats(ilIndex).dSpots60s(1) > 0 Then tmGrf.lDollars(6) = ((tmCompareStats(ilIndex).dRates60s(1) / 100) / tmCompareStats(ilIndex).dSpots60s(1))
                If tmCompareStats(ilIndex).dSpots60s(2) > 0 Then tmGrf.lDollars(7) = ((tmCompareStats(ilIndex).dRates60s(2) / 100) / tmCompareStats(ilIndex).dSpots60s(2))
                If tmCompareStats(ilIndex).dSpots60s(3) > 0 Then tmGrf.lDollars(8) = ((tmCompareStats(ilIndex).dRates60s(3) / 100) / tmCompareStats(ilIndex).dSpots60s(3))
                If tmCompareStats(ilIndex).dSpots60s(4) > 0 Then tmGrf.lDollars(9) = ((tmCompareStats(ilIndex).dRates60s(4) / 100) / tmCompareStats(ilIndex).dSpots60s(4))
            End If
        Else
            'Average 30s Rate Comparison
            If (slShowUnitPrice = "Percent") Then
                'Show Rate Card Percentage instead of Unit Price
                'if Proposal price = zero, show asterisk in the report (marked by -32767)
                'divide revenue $ by 100 and same for prop price; multipled by 1000 to show xx.x in crystal (change formula formatting in CR)
                tmGrf.lDollars(5) = -32767
                If (tmCompareStats(ilIndex).dPropPrice(0) > 0) Then
                    If tmCompareStats(ilIndex).dSpots(0) > 0 Then
                        If ((tmCompareStats(ilIndex).dRates(0) / 100) / tmCompareStats(ilIndex).dSpots(0)) > 0 Then
                            tmGrf.lDollars(5) = ((((tmCompareStats(ilIndex).dRates(0) / 100) / tmCompareStats(ilIndex).dSpots(0)) - _
                                                ((tmCompareStats(ilIndex).dPropPrice(0) / 100) / tmCompareStats(ilIndex).dSpots(0))) / _
                                                ((tmCompareStats(ilIndex).dRates(0) / 100) / tmCompareStats(ilIndex).dSpots(0))) * 1000
                        End If
                    End If
                End If
                
                tmGrf.lDollars(6) = -32767
                If (tmCompareStats(ilIndex).dPropPrice(1) > 0) Then
                    If tmCompareStats(ilIndex).dSpots(1) > 0 Then
                        If ((tmCompareStats(ilIndex).dRates(1) / 100) / tmCompareStats(ilIndex).dSpots(1)) > 0 Then
                            tmGrf.lDollars(6) = ((((tmCompareStats(ilIndex).dRates(1) / 100) / tmCompareStats(ilIndex).dSpots(1)) - _
                                                ((tmCompareStats(ilIndex).dPropPrice(1) / 100) / tmCompareStats(ilIndex).dSpots(1))) / _
                                                ((tmCompareStats(ilIndex).dRates(1) / 100) / tmCompareStats(ilIndex).dSpots(1))) * 1000
                        End If
                    End If
                End If
                
                tmGrf.lDollars(7) = -32767
                If (tmCompareStats(ilIndex).dPropPrice(2) > 0) Then
                    If tmCompareStats(ilIndex).dSpots(2) > 0 Then
                        If ((tmCompareStats(ilIndex).dRates(2) / 100) / tmCompareStats(ilIndex).dSpots(2)) > 0 Then
                            tmGrf.lDollars(7) = ((((tmCompareStats(ilIndex).dRates(2) / 100) / tmCompareStats(ilIndex).dSpots(2)) - _
                                                ((tmCompareStats(ilIndex).dPropPrice(2) / 100) / tmCompareStats(ilIndex).dSpots(2))) / _
                                                ((tmCompareStats(ilIndex).dRates(2) / 100) / tmCompareStats(ilIndex).dSpots(2))) * 1000
                        End If
                    End If
                End If
                
                tmGrf.lDollars(8) = -32767
                If (tmCompareStats(ilIndex).dPropPrice(3) > 0) Then
                    If tmCompareStats(ilIndex).dSpots(3) > 0 Then
                        If ((tmCompareStats(ilIndex).dRates(3) / 100) / tmCompareStats(ilIndex).dSpots(3)) > 0 Then
                            tmGrf.lDollars(8) = ((((tmCompareStats(ilIndex).dRates(3) / 100) / tmCompareStats(ilIndex).dSpots(3)) - _
                                                ((tmCompareStats(ilIndex).dPropPrice(3) / 100) / tmCompareStats(ilIndex).dSpots(3))) / _
                                                ((tmCompareStats(ilIndex).dRates(3) / 100) / tmCompareStats(ilIndex).dSpots(3))) * 1000
                        End If
                    End If
                End If
                
                tmGrf.lDollars(9) = -32767
                If (tmCompareStats(ilIndex).dPropPrice(4) > 0) Then
                    If tmCompareStats(ilIndex).dSpots(4) > 0 Then
                        If ((tmCompareStats(ilIndex).dRates(4) / 100) / tmCompareStats(ilIndex).dSpots(4)) > 0 Then
                            tmGrf.lDollars(9) = ((((tmCompareStats(ilIndex).dRates(4) / 100) / tmCompareStats(ilIndex).dSpots(4)) - _
                                                ((tmCompareStats(ilIndex).dPropPrice(4) / 100) / tmCompareStats(ilIndex).dSpots(4))) / _
                                                ((tmCompareStats(ilIndex).dRates(4) / 100) / tmCompareStats(ilIndex).dSpots(4))) * 1000
                        End If
                    End If
                End If
            Else
                'Show Rate Card Unit Price instead of percentage
                tmGrf.lDollars(5) = -32767
                If (tmCompareStats(ilIndex).dPropPrice(0) > 0) Then
                    If tmCompareStats(ilIndex).dSpots(0) > 0 Then tmGrf.lDollars(5) = ((tmCompareStats(ilIndex).dPropPrice(0) / 100) / tmCompareStats(ilIndex).dSpots(0))
                End If
                
                tmGrf.lDollars(6) = -32767
                If (tmCompareStats(ilIndex).dPropPrice(1) > 0) Then
                    If tmCompareStats(ilIndex).dSpots(1) > 0 Then tmGrf.lDollars(6) = ((tmCompareStats(ilIndex).dPropPrice(1) / 100) / tmCompareStats(ilIndex).dSpots(1))
                End If
                
                tmGrf.lDollars(7) = -32767
                If (tmCompareStats(ilIndex).dPropPrice(2) > 0) Then
                    If tmCompareStats(ilIndex).dSpots(2) > 0 Then tmGrf.lDollars(7) = ((tmCompareStats(ilIndex).dPropPrice(2) / 100) / tmCompareStats(ilIndex).dSpots(2))
                End If
                
                tmGrf.lDollars(8) = -32767
                If (tmCompareStats(ilIndex).dPropPrice(3) > 0) Then
                    If tmCompareStats(ilIndex).dSpots(3) > 0 Then tmGrf.lDollars(8) = ((tmCompareStats(ilIndex).dPropPrice(3) / 100) / tmCompareStats(ilIndex).dSpots(3))
                End If
                
                tmGrf.lDollars(9) = -32767
                If (tmCompareStats(ilIndex).dPropPrice(4) > 0) Then
                    If tmCompareStats(ilIndex).dSpots(4) > 0 Then tmGrf.lDollars(9) = ((tmCompareStats(ilIndex).dPropPrice(4) / 100) / tmCompareStats(ilIndex).dSpots(4))
                End If
            End If
        End If

        'Avg for columns: 1,3,5,7,9
        llTemp = 0: llLineGrossWithoutAcq = 0: dlActTotal = 0
        For ilCounter = 0 To UBound(tmRatePriceDates) - 1
            dlActTotal = dlActTotal + tmCompareStats(ilIndex).dRates(ilCounter)
            llTemp = llTemp + tmCompareStats(ilIndex).dSpots(ilCounter)
            'start/end year range for display in report
            tmGrf.iPerGenl(ilCounter) = Val(RptSelAvgCmp!edcStartYear.Text) - (Trim(RptSelAvgCmp!edcYears.Text) - (ilCounter + 1))
        Next ilCounter
        If llTemp > 0 Then tmGrf.lDollars(15) = ((dlActTotal / 100) / llTemp) Else tmGrf.lDollars(15) = 0
        
        'Avg for columns: 2,4,6,8,10
        'Spot Price Comparison and Separate 30s/60s Spots OR Average 30s Rate Comparison
        dlActTotal = 0: dlPropTotal = 0: llSpotsTotal = 0
        If (((ilRatePrice = 1) And (slAvgBy = "Separate")) Or (ilRatePrice = 0)) Then
            'we shouldn't be here if we're doing Avg. Spot Price for ALL spots
            llTemp = 0: llLineGrossWithoutAcq = 0
            For ilCounter = 0 To UBound(tmRatePriceDates) - 1
                If (ilRatePrice = 1) Then
                    dlActTotal = dlActTotal + tmCompareStats(ilIndex).dRates60s(ilCounter)
                    llSpotsTotal = llSpotsTotal + tmCompareStats(ilIndex).dSpots60s(ilCounter)
                Else
                    If (tmCompareStats(ilIndex).dPropPrice(ilCounter) > 0) Then
                        dlActTotal = dlActTotal + tmCompareStats(ilIndex).dRates(ilCounter)
                        llSpotsTotal = llSpotsTotal + tmCompareStats(ilIndex).dSpots(ilCounter)
                        dlPropTotal = dlPropTotal + tmCompareStats(ilIndex).dPropPrice(ilCounter)
                    End If
                End If
                'start/end year range for display in report
                tmGrf.iPerGenl(5 + ilCounter) = Val(RptSelAvgCmp!edcStartYear.Text) - (Trim(RptSelAvgCmp!edcYears.Text) - (ilCounter + 1))
            Next ilCounter
            
            tmGrf.lDollars(16) = 0
            If ilRatePrice = 0 Then
                tmGrf.lDollars(16) = -32767
                If (dlPropTotal > 0) Then
                    If (slShowUnitPrice = "Percent") Then
                        If ((dlActTotal > 0) And (llSpotsTotal > 0)) Then
                            tmGrf.lDollars(16) = ((((dlActTotal / 100) / llSpotsTotal) - ((dlPropTotal / 100) / llSpotsTotal)) / ((dlActTotal / 100) / llSpotsTotal) * 1000)
                        End If
                    Else
                        If llSpotsTotal > 0 Then tmGrf.lDollars(16) = ((dlPropTotal / 100) / llSpotsTotal)
                    End If
                End If
            Else
                If llSpotsTotal > 0 Then tmGrf.lDollars(16) = ((dlActTotal / 100) / llSpotsTotal)
            End If
        End If
        
        'column 1
        dlDollars_0 = dlDollars_0 + (tmCompareStats(ilIndex).dRates(0) / 100)
        dlDollarsProp_0 = dlDollarsProp_0 + (tmCompareStats(ilIndex).dPropPrice(0) / 100)
        llCountDollars_0 = llCountDollars_0 + tmCompareStats(ilIndex).dSpots(0)
        'column 3
        dlDollars_1 = dlDollars_1 + (tmCompareStats(ilIndex).dRates(1) / 100)
        dlDollarsProp_1 = dlDollarsProp_1 + (tmCompareStats(ilIndex).dPropPrice(1) / 100)
        llCountDollars_1 = llCountDollars_1 + tmCompareStats(ilIndex).dSpots(1)
        'column 5
        dlDollars_2 = dlDollars_2 + (tmCompareStats(ilIndex).dRates(2) / 100)
        dlDollarsProp_2 = dlDollarsProp_2 + (tmCompareStats(ilIndex).dPropPrice(2) / 100)
        llCountDollars_2 = llCountDollars_2 + tmCompareStats(ilIndex).dSpots(2)
        'column 7
        dlDollars_3 = dlDollars_3 + (tmCompareStats(ilIndex).dRates(3) / 100)
        dlDollarsProp_3 = dlDollarsProp_3 + (tmCompareStats(ilIndex).dPropPrice(3) / 100)
        llCountDollars_3 = llCountDollars_3 + tmCompareStats(ilIndex).dSpots(3)
        'column 9
        dlDollars_4 = dlDollars_4 + (tmCompareStats(ilIndex).dRates(4) / 100)
        dlDollarsProp_4 = dlDollarsProp_4 + (tmCompareStats(ilIndex).dPropPrice(4) / 100)
        llCountDollars_4 = llCountDollars_4 + tmCompareStats(ilIndex).dSpots(4)
        
        'column 2
        dlDollars_5 = dlDollars_5 + (tmCompareStats(ilIndex).dRates60s(0) / 100)
        'dlDollarsProp_5 = dlDollarsProp_5 + tmCompareStats(ilIndex).dPropPrice60s(0)
        llCountDollars_5 = llCountDollars_5 + tmCompareStats(ilIndex).dSpots60s(0)
        'column 4
        dlDollars_6 = dlDollars_6 + (tmCompareStats(ilIndex).dRates60s(1) / 100)
        'dlDollarsProp_6 = dlDollarsProp_6 + tmCompareStats(ilIndex).dPropPrice60s(1)
        llCountDollars_6 = llCountDollars_6 + tmCompareStats(ilIndex).dSpots60s(1)
        'column 6
        dlDollars_7 = dlDollars_7 + (tmCompareStats(ilIndex).dRates60s(2) / 100)
        'dlDollarsProp_7 = dlDollarsProp_7 + tmCompareStats(ilIndex).dPropPrice60s(2)
        llCountDollars_7 = llCountDollars_7 + tmCompareStats(ilIndex).dSpots60s(2)
        'column 8
        dlDollars_8 = dlDollars_8 + (tmCompareStats(ilIndex).dRates60s(3) / 100)
        'dlDollarsProp_8 = dlDollarsProp_8 + tmCompareStats(ilIndex).dPropPrice60s(3)
        llCountDollars_8 = llCountDollars_8 + tmCompareStats(ilIndex).dSpots60s(3)
        'column 10
        dlDollars_9 = dlDollars_9 + (tmCompareStats(ilIndex).dRates60s(4) / 100)
        'dlDollarsProp_9 = dlDollarsProp_9 + tmCompareStats(ilIndex).dPropPrice60s(4)
        llCountDollars_9 = llCountDollars_9 + tmCompareStats(ilIndex).dSpots60s(4)
        
        'Avg for columns: 1,3,5,7,9  (YYYY-YYYY Avg Rate)
        dlDollars_15 = dlDollars_15 + dlDollars_0 + dlDollars_1 + dlDollars_2 + dlDollars_3 + dlDollars_4
        llCountDollars_15 = llCountDollars_15 + llCountDollars_0 + llCountDollars_1 + llCountDollars_2 + llCountDollars_3 + llCountDollars_4
        
        'Avg for columns: 2,4,6,8,10 (YYYY-YYYY Avg Using R/C)
        dlDollars_16 = dlDollars_16 + dlDollars_5 + dlDollars_6 + dlDollars_7 + dlDollars_8 + dlDollars_9
        'dlDollarsProp_16 = dlDollarsProp_5 + dlDollarsProp_6 + dlDollarsProp_7 + dlDollarsProp_8 + dlDollarsProp_9
        llCountDollars_16 = llCountDollars_16 + llCountDollars_5 + llCountDollars_6 + llCountDollars_7 + llCountDollars_8 + llCountDollars_9
        
        'write record to DB
        llRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
    Next ilIndex
    
    'create the totals record
    tmGrf.iVefCode = 0
    tmGrf.iAdfCode = 1
    If llCountDollars_0 > 0 Then tmGrf.lDollars(0) = (dlDollars_0 / llCountDollars_0) Else tmGrf.lDollars(0) = 0
    If llCountDollars_1 > 0 Then tmGrf.lDollars(1) = (dlDollars_1 / llCountDollars_1) Else tmGrf.lDollars(1) = 0
    If llCountDollars_2 > 0 Then tmGrf.lDollars(2) = (dlDollars_2 / llCountDollars_2) Else tmGrf.lDollars(2) = 0
    If llCountDollars_3 > 0 Then tmGrf.lDollars(3) = (dlDollars_3 / llCountDollars_3) Else tmGrf.lDollars(3) = 0
    If llCountDollars_4 > 0 Then tmGrf.lDollars(4) = (dlDollars_4 / llCountDollars_4) Else tmGrf.lDollars(4) = 0
    
    tmGrf.lDollars(5) = 0
    If ilRatePrice = 0 Then
        tmGrf.lDollars(5) = -32767
        If (dlDollarsProp_0 > 0) Then
            If (slShowUnitPrice = "Percent") Then
                'If ((dlDollars_0 > 0) And (llCountDollars_0 > 0)) Then tmGrf.lDollars(5) = ((((dlDollars_0 - dlDollarsProp_0) / dlDollars_0) / llCountDollars_0) * 1000)
                If ((dlDollars_0 > 0) And (llCountDollars_0 > 0)) Then
                    tmGrf.lDollars(5) = ((((dlDollars_0 / llCountDollars_0) - (dlDollarsProp_0 / llCountDollars_0)) / (dlDollars_0 / llCountDollars_0)) * 1000)
                End If
            Else
                If llCountDollars_0 > 0 Then tmGrf.lDollars(5) = (dlDollarsProp_0 / llCountDollars_0)
            End If
        End If
    Else
        If llCountDollars_5 > 0 Then tmGrf.lDollars(5) = ((dlDollars_5) / llCountDollars_5)
    End If
    
    tmGrf.lDollars(6) = 0
    If ilRatePrice = 0 Then
        tmGrf.lDollars(6) = -32767
        If (dlDollarsProp_1 > 0) Then
            If (slShowUnitPrice = "Percent") Then
                If ((dlDollars_1 > 0) And (llCountDollars_1 > 0)) Then
                    tmGrf.lDollars(6) = ((((dlDollars_1 / llCountDollars_1) - (dlDollarsProp_1 / llCountDollars_1)) / (dlDollars_1 / llCountDollars_1)) * 1000)
                End If
            Else
                If llCountDollars_1 > 0 Then tmGrf.lDollars(6) = (dlDollarsProp_1 / llCountDollars_1)
            End If
        End If
    Else
        If llCountDollars_6 > 0 Then tmGrf.lDollars(6) = ((dlDollars_6) / llCountDollars_6)
    End If
    
    tmGrf.lDollars(7) = 0
    If ilRatePrice = 0 Then
        tmGrf.lDollars(7) = -32767
        If (dlDollarsProp_2 > 0) Then
            If (slShowUnitPrice = "Percent") Then
                If ((dlDollars_2 > 0) And (llCountDollars_2 > 0)) Then
                    tmGrf.lDollars(7) = ((((dlDollars_2 / llCountDollars_2) - (dlDollarsProp_2 / llCountDollars_2)) / (dlDollars_2 / llCountDollars_2)) * 1000)
                End If
            Else
                If llCountDollars_2 > 0 Then tmGrf.lDollars(7) = (dlDollarsProp_2 / llCountDollars_2)
            End If
        End If
    Else
        If llCountDollars_7 > 0 Then tmGrf.lDollars(7) = ((dlDollars_7) / llCountDollars_7)
    End If
    
    tmGrf.lDollars(8) = 0
    If ilRatePrice = 0 Then
        tmGrf.lDollars(8) = -32767
        If (dlDollarsProp_3 > 0) Then
            If (slShowUnitPrice = "Percent") Then
                If ((dlDollars_3 > 0) And (llCountDollars_3 > 0)) Then
                    tmGrf.lDollars(8) = ((((dlDollars_3 / llCountDollars_3) - (dlDollarsProp_3 / llCountDollars_3)) / (dlDollars_3 / llCountDollars_3)) * 1000)
                End If
            Else
                If llCountDollars_3 > 0 Then tmGrf.lDollars(8) = (dlDollarsProp_3 / llCountDollars_3)
            End If
        End If
    Else
        If llCountDollars_8 > 0 Then tmGrf.lDollars(8) = ((dlDollars_8) / llCountDollars_8)
    End If
    
    tmGrf.lDollars(9) = 0
    If ilRatePrice = 0 Then
        tmGrf.lDollars(9) = -32767
        If (dlDollarsProp_4 > 0) Then
            If (slShowUnitPrice = "Percent") Then
                If ((dlDollars_4 > 0) And (llCountDollars_4 > 0)) Then
                    tmGrf.lDollars(9) = ((((dlDollars_4 / llCountDollars_4) - (dlDollarsProp_4 / llCountDollars_4)) / (dlDollars_4 / llCountDollars_4)) * 1000)
                End If
            Else
                If llCountDollars_4 > 0 Then tmGrf.lDollars(9) = (dlDollarsProp_4 / llCountDollars_4)
            End If
        End If
    Else
        If llCountDollars_9 > 0 Then tmGrf.lDollars(9) = ((dlDollars_9) / llCountDollars_9)
    End If
    
    tmGrf.lDollars(15) = 0
    If llCountDollars_15 > 0 Then tmGrf.lDollars(15) = (dlDollars_15 / llCountDollars_15)
    
    dlDollars_Total = IIF(dlDollarsProp_0 > 0, dlDollars_0, 0) + IIF(dlDollarsProp_1 > 0, dlDollars_1, 0) + IIF(dlDollarsProp_2 > 0, dlDollars_2, 0) + IIF(dlDollarsProp_3 > 0, dlDollars_3, 0) + IIF(dlDollarsProp_4 > 0, dlDollars_4, 0)
    dlDollarsProp_Total = dlDollarsProp_0 + dlDollarsProp_1 + dlDollarsProp_2 + dlDollarsProp_3 + dlDollarsProp_4
    llCount_Total = IIF(dlDollarsProp_0 > 0, llCountDollars_0, 0) + IIF(dlDollarsProp_1 > 0, llCountDollars_1, 0) + IIF(dlDollarsProp_2 > 0, llCountDollars_2, 0) + IIF(dlDollarsProp_3 > 0, llCountDollars_3, 0) + IIF(dlDollarsProp_4 > 0, llCountDollars_4, 0)
    
    tmGrf.lDollars(16) = 0
    If ilRatePrice = 0 Then
        tmGrf.lDollars(16) = -32767
        If (dlDollarsProp_Total > 0) Then
            If (slShowUnitPrice = "Percent") Then
                If ((dlDollars_Total > 0) And (llCount_Total > 0)) Then
                    tmGrf.lDollars(16) = ((((dlDollars_Total / llCount_Total) - (dlDollarsProp_Total / llCount_Total)) / (dlDollars_Total / llCount_Total)) * 1000)
                End If
            Else
                If llCount_Total > 0 Then tmGrf.lDollars(16) = (dlDollarsProp_Total / llCount_Total)
            End If
        End If
    Else
        If llCountDollars_16 > 0 Then tmGrf.lDollars(16) = (dlDollars_16 / llCountDollars_16)
    End If
    
    'write record to DB
    llRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)

    Erase tmChfAdvtExt, tgClf, tgCff, tmCompareStats
    Erase llStartDates, dlProject, dlProjectSpots, dlProjectProp, dlProject60s, dlProjectSpots60s, dlProjectProp60s
    Erase imUsevefcodes, imUseAdvtCodes, imUseSlspCodes, imUseVGCodes
    sgCntrForDateStamp = ""
    llRet = btrClose(hmCHF)
    llRet = btrClose(hmClf)
    llRet = btrClose(hmCff)
    llRet = btrClose(hmGrf)
    llRet = btrClose(hmAgf)
    Exit Sub
End Sub

'
'               mObtainSelectivityAVG - get the user selected parameters
'
Private Sub mObtainSelectivityAVG()
'Dim ilVGSort As Integer
Dim ilLoop As Integer

    'create a module integer for imWhichLine: package or airing lines
    imWhichLine = IIF(RptSelAvgCmp!rbcUseLines(0).Value = True, 0, 1)
    
    'Selective contract #
    lmSingleCntr = Val(RptSelAvgCmp!edcContract.Text)
    
    tmCntTypes.iHold = gSetCheck(RptSelAvgCmp!ckcAllTypes(0).Value)
    tmCntTypes.iOrder = gSetCheck(RptSelAvgCmp!ckcAllTypes(1).Value)
    tmCntTypes.iStandard = gSetCheck(RptSelAvgCmp!ckcAllTypes(3).Value)
    tmCntTypes.iReserv = gSetCheck(RptSelAvgCmp!ckcAllTypes(4).Value)
    tmCntTypes.iRemnant = gSetCheck(RptSelAvgCmp!ckcAllTypes(5).Value)
    tmCntTypes.iDR = gSetCheck(RptSelAvgCmp!ckcAllTypes(6).Value)
    tmCntTypes.iPI = gSetCheck(RptSelAvgCmp!ckcAllTypes(7).Value)
    tmCntTypes.iPSA = gSetCheck(RptSelAvgCmp!ckcAllTypes(8).Value)
    tmCntTypes.iPromo = gSetCheck(RptSelAvgCmp!ckcAllTypes(9).Value)
    tmCntTypes.iTrade = gSetCheck(RptSelAvgCmp!ckcAllTypes(10).Value)
    tmCntTypes.iAirTime = gSetCheck(RptSelAvgCmp!ckcAllTypes(11).Value)
    tmCntTypes.iRep = gSetCheck(RptSelAvgCmp!ckcAllTypes(12).Value)
    tmCntTypes.iNC = gSetCheck(RptSelAvgCmp!ckcAllTypes(14).Value)
    tmCntTypes.iPolit = gSetCheck(RptSelAvgCmp!ckcAllTypes(15).Value)           'as previously for Feed spots
    tmCntTypes.iNonPolit = gSetCheck(RptSelAvgCmp!ckcAllTypes(16).Value)
    
    ReDim imUsevefcodes(0 To 0) As Integer
    gObtainCodesForMultipleLists 0, tgVehicle(), imInclVefCodes, imUsevefcodes(), RptSelAvgCmp
    
    imMajorSet = 0
End Sub


'
'                   mFilterSelectivityAVG - test user selectivity to determine if valid contract to process
'                   <input> index of active contract array
'                   return - true if passed selectivity
'
Public Function mFilterSelectivityAVG(ilCurrentRecd As Integer) As Boolean
Dim llContrCode As Long
Dim blValidCType As Boolean
Dim blFoundOne As Boolean
Dim ilRet As Integer
Dim ilIsItPolitical As Integer

        blFoundOne = True                              'set default to true incase by vehicle, advt should not be filtered
        ilIsItPolitical = gIsItPolitical(tmChfAdvtExt(ilCurrentRecd).iAdfCode)           'its a political, include this contract?
        'test for inclusion if its political adv and politicals requested, or
        'its not a political adv and politicals
        If (tmCntTypes.iPolit And ilIsItPolitical) Or ((tmCntTypes.iNonPolit) And (Not ilIsItPolitical)) Then           'ok
            blFoundOne = blFoundOne
        Else
            blFoundOne = False
        End If
        If blFoundOne Then
            'Retrieve the contract, schedule lines and flights
            llContrCode = tmChfAdvtExt(ilCurrentRecd).lCode
            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChf, tgClf(), tgCff())

            If Not ilRet Then
                On Error GoTo mFilterSelectivityAVGErr
                gBtrvErrorMsg ilRet, "gCreateMarginAcqErr (mFilterSelectivity: gObtainCntr):" & "Chf.Btr", RptSelAvgCmp
                On Error GoTo 0
            End If
    
            blValidCType = gFilterContractType(tgChf, tmCntTypes, True)         'include proposal type checks
            If blValidCType Then                                        'test for 100% trade inclusion
                'if TRADE option is selected, include ALL trades; if not, exclude if trade = 100%
                If tmCntTypes.iTrade = False And tgChf.iPctTrade = 100 Then
                    blValidCType = False
                End If
            End If
        End If
            
        mFilterSelectivityAVG = True
        If Not blFoundOne Or Not blValidCType Then
            mFilterSelectivityAVG = False
        End If
        Exit Function
mFilterSelectivityAVGErr:
    Resume Next
    
End Function

'*************************************************************
'*                                                           *
'*      Procedure Name:mBinarySearchAVG                      *
'*                                                           *
'*             Created:6/13/93       By:D. LeVine            *
'*            Modified:11/14/2018    By:FYM                  *
'*                                                           *
'*            Comments:Obtain Advt index into tmCompareStats *
'*                                                           *
'*************************************************************
Public Function mBinarySearchAVG(ilCode As Integer, tmCompareStats() As COMPARESTATS) As Integer
    Dim ilMin As Integer
    Dim ilMax As Integer
    Dim ilMiddle As Integer
    ilMin = LBound(tmCompareStats)
    ilMax = UBound(tmCompareStats) - 1
    Do While ilMin <= ilMax
        ilMiddle = (ilMin + ilMax) \ 2
        If ilCode = tmCompareStats(ilMiddle).sKey Then
            'found the match
            mBinarySearchAVG = ilMiddle
            Exit Function
        ElseIf ilCode < tmCompareStats(ilMiddle).sKey Then
            ilMax = ilMiddle - 1
        Else
            'search the right half
            ilMin = ilMiddle + 1
        End If
    Loop
    mBinarySearchAVG = -1
End Function



'
'                   mBuildFlightSpotRevPropPrice - Loop through the flights of the schedule line
'                           and build the projections dollars into dllProject array,
'                           and build projection # of spots into dlProjectSpots array
'                   <input> ilclf = sched line index into tlClfInp; index of the current record from the main loop
'                           llStdStartDates() - array of dates to build $ from flights; max of 6 years
'                           ilFirstProjInx - index of 1st month/week to start projecting; use 1
'                           ilMaxInx - max # of buckets to loop thru (???)
'                           ilWkOrMonth - 1 = Month, 2 = Week; use 1
'                           ilUseWhichRate - 0 = use true line rate, 1 = use acquisition rate, 2 =use acq rate if non0, otherwise use linerate; use 0
'                           slGrossOrNet - G = Gross , N = Net (default to Net).  USed to acquisition costs computation if using Acq commissions; use "G"
'                           ilRatePrice - 0=rate, 1=spot price: adjust if 0, llSpot based on spot length; 30s = 1 unit and less than 30s = 1 unit
'                  <output> dllProject() = array of $ buckets corresponding to array of dates; buffer for contract lines
'                           dlProjectSpots() array of spot count buckets corresponding to array of dates; buffer for flights (schedules)
'                           dlProjectProp() array of proposal prices
'
'                           General routine to build flight $/cpot count into week, month, qtr buckets
'                   Created : 10-31-2018
Public Sub mBuildFlightSpotRevPropPrice(ilClf As Integer, llStdStartDates() As Long, ilFirstProjInx As Integer, ilMaxInx As Integer, dlProject() As Double, _
            dlProjectSpots() As Double, ilWkOrMonth As Integer, ilUseWhichRate As Integer, tlClfInp() As CLFLIST, tlCffInp() As CFFLIST, dlProjectProp() As Double, _
            ilRatePrice As Integer, slAvgBy As String, slSpotLen As String, dlProject60s() As Double, dlProjectSpots60s() As Double, dlProjectProp60s() As Double, Optional slGrossOrNet As String = "N")
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
Dim ilUnits As Integer
Dim iMod As Integer

    llStdStart = llStdStartDates(ilFirstProjInx)
    llStdEnd = llStdStartDates(ilMaxInx + 1)          'look into later;needs to be 6
    ilCff = tlClfInp(ilClf).iFirstCff
    Do While ilCff <> -1
        tlCff = tlCffInp(ilCff).CffRec
    
        'if start date > end date, it's a cancel before start
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
            
            If ilUseWhichRate = 0 Then             'always use linerate (actual price of each spot)
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
                llWhichRate = tlCff.lActPrice
            End If
            
            'Include/Exclude Price Types: T,N,M,B,S,P,R,A --> can be zero dollars
            If ((llWhichRate > 0) Or (llWhichRate = 0 And tmCntTypes.iNC = True)) Then
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
                    
                    'ilRatePrice: 0=Rate Comparison: adjust llSpot based on spot length 30s or less = 1 unit; 1=Spot Price Comparison
                    If ilRatePrice = 0 Then
                        'len=45s --> 2 units
                        'llSpots * ((tlClfInp(ilClf).ClfRec.iLen / 30) + (mod (tlClfInp(ilClf).ClfRec.iLen / 30)))
                        If tlClfInp(ilClf).ClfRec.iLen < 30 Then
                            ilUnits = 1
                        Else
                            ilUnits = (tlClfInp(ilClf).ClfRec.iLen / 30)
                            iMod = (tlClfInp(ilClf).ClfRec.iLen Mod 30)
                            If (iMod > 0 And iMod < 15) Then ilUnits = ilUnits + 1
                        End If
                        ilUnits = llSpots * ilUnits
                    End If
                    
                    If ilWkOrMonth = 1 Then                     'monthly buckets
                        'determine month that this week belongs in, then accumulate the gross and net $
                        'currently, the projections are based on STandard bdcst
                        For ilMonthInx = ilFirstProjInx To ilMaxInx Step 1        'loop thru months to find the match
                            If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                                If (ilRatePrice = 1) Then           'Avg Spot Price Comparison
                                    If slAvgBy = "Combined" Then    'For all spots length
                                        If slSpotLen = "All" Then
                                            'Spot Lengths to include: All
                                            dlProject(ilMonthInx) = dlProject(ilMonthInx) + ((llSpots * llWhichRate))
                                            'dlProjectProp(ilMonthInx) = dlProjectProp(ilMonthInx) + ((llSpots * tlCff.lPropPrice))
                                            dlProjectSpots(ilMonthInx) = dlProjectSpots(ilMonthInx) + llSpots
                                        Else
                                            'Spot Lengths to include: 30s or 60s Only
                                            If ((tlClfInp(ilClf).ClfRec.iLen = 30) Or (tlClfInp(ilClf).ClfRec.iLen = 60)) Then
                                                dlProject(ilMonthInx) = dlProject(ilMonthInx) + ((llSpots * llWhichRate))
                                                'dlProjectProp(ilMonthInx) = dlProjectProp(ilMonthInx) + ((llSpots * tlCff.lPropPrice))
                                                dlProjectSpots(ilMonthInx) = dlProjectSpots(ilMonthInx) + llSpots
                                            End If
                                        End If
                                    Else
                                        'Avg Spot Price Comparison of 30s/60s (Only)
                                        If slSpotLen = "All" Then
                                            'ALL spots: 10,15,30,45,60,90,120 etc.
                                            If (tlClfInp(ilClf).ClfRec.iLen < 60) Then
                                                'anything less than 60s in length goes to 30s bucket
                                                dlProject(ilMonthInx) = dlProject(ilMonthInx) + ((llSpots * llWhichRate))
                                                'dlProjectProp(ilMonthInx) = dlProjectProp(ilMonthInx) + ((llSpots * tlCff.lPropPrice))
                                                dlProjectSpots(ilMonthInx) = dlProjectSpots(ilMonthInx) + llSpots
                                            ElseIf (tlClfInp(ilClf).ClfRec.iLen >= 60) Then
                                                'anything 60s or greater in length goes to 60s bucket
                                                dlProject60s(ilMonthInx) = dlProject60s(ilMonthInx) + ((llSpots * llWhichRate))
                                                'dlProjectProp60s(ilMonthInx) = dlProjectProp60s(ilMonthInx) + (llSpots * tlCff.lPropPrice)
                                                dlProjectSpots60s(ilMonthInx) = dlProjectSpots60s(ilMonthInx) + llSpots
                                            End If
                                        Else
                                            '30s or 60s spots only
                                            If (tlClfInp(ilClf).ClfRec.iLen = 30) Then
                                                'only 30s in length goes to 30s bucket
                                                dlProject(ilMonthInx) = dlProject(ilMonthInx) + ((llSpots * llWhichRate))
                                                dlProjectProp(ilMonthInx) = dlProjectProp(ilMonthInx) + ((llSpots * tlCff.lPropPrice))
                                                dlProjectSpots(ilMonthInx) = dlProjectSpots(ilMonthInx) + llSpots
                                            ElseIf (tlClfInp(ilClf).ClfRec.iLen = 60) Then
                                                'only 60s in length goes to 60s bucket
                                                dlProject60s(ilMonthInx) = dlProject60s(ilMonthInx) + ((llSpots * llWhichRate))
                                                dlProjectProp60s(ilMonthInx) = dlProjectProp60s(ilMonthInx) + (llSpots * tlCff.lPropPrice)
                                                dlProjectSpots60s(ilMonthInx) = dlProjectSpots60s(ilMonthInx) + llSpots
                                            End If
                                        End If
                                    End If
                                    Exit For
                                Else
                                    'Avg. 30s Rate Comparison: ALL Spots
                                    If slSpotLen = "All" Then
                                        dlProject(ilMonthInx) = dlProject(ilMonthInx) + (llSpots * llWhichRate)
                                        ' add the proposal rates on this array (cff.PropPrice)
                                        dlProjectProp(ilMonthInx) = dlProjectProp(ilMonthInx) + (llSpots * tlCff.lPropPrice * 100)
                                        dlProjectSpots(ilMonthInx) = dlProjectSpots(ilMonthInx) + ilUnits
                                    Else
                                        'Avg. 30s Rate Comparison: 30s/60s Spots Only
                                        If ((tlClfInp(ilClf).ClfRec.iLen = 30) Or (tlClfInp(ilClf).ClfRec.iLen = 60)) Then
                                            dlProject(ilMonthInx) = dlProject(ilMonthInx) + (llSpots * llWhichRate)
                                            ' add the proposal rates on this array (cff.PropPrice)
                                            dlProjectProp(ilMonthInx) = dlProjectProp(ilMonthInx) + (llSpots * tlCff.lPropPrice * 100)
                                            dlProjectSpots(ilMonthInx) = dlProjectSpots(ilMonthInx) + ilUnits
                                        End If
                                    End If
                                    Exit For
                                End If
                            End If
                        Next ilMonthInx
                    Else                                    'weekly buckets
                        ilWkInx = (llDate - llStdStartDates(1)) \ 7 + 1
                        ''4-3-07 make sure the data isnt gathered beyond the period requested
                        'If ilWkInx > 0 And llDate >= llStdStartDates(LBound(llStdStartDates)) And llDate < llStdStartDates(ilMaxInx) Then   '1-24-08(UBound(llStdStartDates)) Then
                        If ilWkInx > 0 And llDate >= llStdStartDates(1) And llDate < llStdStartDates(ilMaxInx) Then   '1-24-08(UBound(llStdStartDates)) Then
                            dlProject(ilWkInx) = dlProject(ilWkInx) + (llSpots * llWhichRate)
                            dlProjectSpots(ilWkInx) = dlProjectSpots(ilWkInx) + llSpots
                            'do the same here for weekly
                            dlProjectProp(ilWkInx) = dlProjectProp(ilWkInx) + (llSpots * tlCff.lPropPrice)
                        End If
                    End If
                Next llDate                                     'for llDate = llFltStart To llFltEnd
            End If
        End If                                          '
        ilCff = tlCffInp(ilCff).iNextCff                   'get next flight record from mem
    Loop                                            'while ilcff <> -1
End Sub

