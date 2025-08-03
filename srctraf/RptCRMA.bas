Attribute VB_Name = "RptCrMA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RptCrMA.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
'Public igYear As Integer                'budget year used for filtering
'Public igMonthOrQtr As Integer          'entered month or qtr
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Public Const TA_LEFT = 0
'Public Const TA_RIGHT = 2
'Public Const TA_CENTER = 6
'Public Const TA_TOP = 0
'Public Const TA_BOTTOM = 8
'Public Const TA_BASELINE = 24
Dim hmCHF As Integer            'Contract header file handle
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChf As CHF
Dim tmChfSrchKey1 As CHFKEY1
Dim tmChfAdvtExt() As CHFADVTEXT
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmAgf As Integer            'AGency file handle
Dim imAgfRecLen As Integer      ' record length
Dim tmAgf As AGF
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
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
'  Receivables File
'********************************************************************************************
'
'           gCreateMarginAcquisition - Prepass to create Margin Acquisition report
'           Create prepass to produce a report to compare acquisition costs to net revenue
'           to arrive at a margin percent to determine if a contract and/or vehicle
'           is profitable or not.  Margin calculation is basec on cost (expanded acq)
'           divided by revenue (net $)
'
'
'
'
'********************************************************************************************
Sub gCreateMarginAcquisition()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilWhichset                                                                            *
'******************************************************************************************

Dim ilRet As Integer                    '
Dim ilClf As Integer                    'loop for schedule lines
Dim ilHOState As Integer                'retrieve only latest order or revision
Dim slCntrTypes As String               'retrieve remnants, PI, DR, etc
Dim slCntrStatus As String              'retrieve H, O G or N (holds, orders, unsch hlds & orders)
Dim ilCurrentRecd As Integer            'loop for processing last years contracts
Dim llContrCode As Long                 'Internal Contr code to build all lines & flights of a contract
Dim blFoundOne As Boolean               'Found a matching  office built into mem
Dim ilValidVehicle As Integer
Dim ilTemp As Integer
Dim ilLoop As Integer                   'temp loop variable
'ReDim llProject(1 To 2) As Long        '$ projected for 1 period ("X" weeks)
ReDim llProject(0 To 2) As Long        '$ projected for 1 period ("X" weeks). Index zero ignored
'ReDim llProjectSpots(1 To 2) As Long        'spot counts for 1 period ("X" weeks)
ReDim llProjectSpots(0 To 2) As Long        'spot counts for 1 period ("X" weeks). Index zero ignored
Dim llDate As Long                      'temp date variable
Dim llDate2 As Long
Dim llLineDate As Long
Dim llLineStartDate As Long
Dim llLineEndDate As Long
Dim ilmnfMinorCode As Integer

'Date used to gather information
'String formats for generalized date conversions routines
'Long formats for testing
'Packed formats to store in GRF record
Dim ilNoWeeks As Integer              '# weeks to gather from start date entered
Dim slWeekStart As String              'start date of week for this years new business entered this week
Dim llWeekStart As Long                'start date of week for this years new business entered on te user entered week
ReDim ilWeekStart(0 To 1) As Integer     'packed format for GRF record
ReDim ilWeekEnd(0 To 1) As Integer
Dim llWeekEnd As Long
Dim slWeekEnd As String
Dim ilWhichRate As Integer
Dim ilWeekOrMonth As Integer
Dim llLineGrossWithoutAcq As Long
Dim ilAgyCommPct As Integer
Dim slCashAgyComm As String
Dim ilIndex As Integer
Dim slAmount As String
Dim slNet As String
'Month Starts to gather projection $ from flights
'ReDim llStartDates(1 To 2) As Long        'start dates of each period to gather (only one)
ReDim llStartDates(0 To 2) As Long        'start dates of each period to gather (only one). Index zero ignored

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)
 
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
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
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)
    
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "AGf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        btrDestroy hmAgf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)
    
    ReDim tgClf(0 To 0) As CLFLIST
    tgClf(0).iStatus = -1 'Not Used
    tgClf(0).lRecPos = 0
    tgClf(0).iFirstCff = -1
    ReDim tgCff(0 To 0) As CFFLIST
    tgCff(0).iStatus = -1 'Not Used
    tgCff(0).lRecPos = 0
    tgCff(0).iNextCff = -1
    
    mObtainSelectivityMA
    'Get STart and end dates of user requested date to find all advertisers airing
    slWeekStart = RptSelMA!calStartDate.Text
    llWeekStart = gDateValue(slWeekStart)
 
    slWeekStart = Format$(llWeekStart, "m/d/yy")
    gPackDate slWeekStart, ilWeekStart(0), ilWeekStart(1)    'conversion to store in prepass record
    ilNoWeeks = Val(RptSelMA!edcNoWeeks.Text)
    llWeekEnd = llWeekStart + ((ilNoWeeks - 1) * 7) + 6
    slWeekEnd = Format$(llWeekEnd, "m/d/yy")
    gPackDate slWeekEnd, ilWeekEnd(0), ilWeekEnd(1)  'btrieve format for prepas record

    'Start date of eachperiod to accumulate spot counts for (only one period applicable)
    llStartDates(1) = llWeekStart
    llStartDates(2) = llWeekEnd + 1
    
    'ilRet = gBuildAcqCommInfo(RptSelCt)         'build acq rep commission table, if applicable
    
    If lmSingleCntr > 0 Then
        ReDim tmChfAdvtExt(0 To 1) As CHFADVTEXT
        tmChfSrchKey1.lCntrNo = lmSingleCntr
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, Len(tmChf), tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
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
        
        ilHOState = 3                       'H or O or G or N or W or C or I (if G or N or W or C or I exists show it over H or O)
        
        ilRet = gObtainCntrForDate(RptSelMA, slWeekStart, slWeekEnd, slCntrStatus, slCntrTypes, ilHOState, tmChfAdvtExt())
    End If

    'common GRF fields that wont change
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    tmGrf.iStartDate(0) = ilWeekStart(0)
    tmGrf.iStartDate(1) = ilWeekStart(1)
    tmGrf.iDate(0) = ilWeekEnd(0)
    tmGrf.iDate(1) = ilWeekEnd(1)
    'tmGrf.iPerGenl(1) = imMajorSet      'used for heading in Crystal report

    'All contracts have been retrieved for all of this year
    For ilCurrentRecd = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1

        blFoundOne = mFilterSelectivity(ilCurrentRecd)
        If blFoundOne Then
            'determine if the contracts start & end dates fall within the requested period
            gUnpackDateLong tgChf.iEndDate(0), tgChf.iEndDate(1), llDate2      'hdr end date converted to long
            gUnpackDateLong tgChf.iStartDate(0), tgChf.iStartDate(1), llDate    'hdr start date converted to long
            
            tmGrf.lChfCode = tgChf.lCode
            tmGrf.iSlfCode = tgChf.iSlfCode(0)
            tmGrf.iAdfCode = tgChf.iAdfCode
            
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

            If llDate2 >= llWeekStart And llDate <= llWeekEnd Then          'the contract dates must span the user requested dates
                For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                    tmClf = tgClf(ilClf).ClfRec
                    'Project the monthly spots from the flights
                    If tmClf.sType = "S" Or tmClf.sType = "H" Then
                    
                        'Setup the major sort factor
                        gGetVehGrpSets tmClf.iVefCode, 0, imMajorSet, ilmnfMinorCode, tmGrf.iCode2
                        'check selectivity of vehicle groups
                        blFoundOne = True
                        If (imMajorSet > 0) Then
                            blFoundOne = False
                            If gFilterLists(tmGrf.iCode2, imInclVGCodes, imUseVGCodes()) Then
                                blFoundOne = True
                            End If
                        Else
                            blFoundOne = True
                        End If
                        
                        ilValidVehicle = True
                        ilValidVehicle = gFilterLists(tmClf.iVefCode, imInclVefCodes, imUsevefcodes())
                            
                        ilWhichRate = 0             'use true line rate
                        ilWeekOrMonth = 1           'assume month gathering since it tests date spans; wkly version calculates a week index
                        llLineGrossWithoutAcq = 0
                        If (blFoundOne) And (ilValidVehicle) Then      'got a valid vehicle group code
                            gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llLineEndDate     'line end date
                            gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llLineStartDate    'Line start date
 
                            If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then        'if using acq rep commissions, need to break out lines by the different acq costs if showing detail
                                mBuildFlightSpotsAndRevenue ilClf, llStartDates(), llWeekStart, llWeekEnd, slCashAgyComm
                             
                            Else
                                gBuildFlightSpotsAndRevenue ilClf, llStartDates(), 1, 2, llProject(), llProjectSpots(), ilWeekOrMonth, ilWhichRate, tgClf(), tgCff()
                                If (llProject(1) <> 0 Or tmClf.lAcquisitionCost <> 0) And llProjectSpots(1) > 0 Then        'must have either true spot rate or acq $, along with a spot count
                                     
                                    'tmGrf.lDollars(1) = llProject(1)        'spot count for period
                                    'tmGrf.lDollars(1) = llProject(1)            'cnt gross for selected period
                                    'tmGrf.lDollars(3) = llProjectSpots(1)
                                    
                                    'tmGrf.lDollars(4) = tmClf.lAcquisitionCost
                                    tmGrf.lDollars(0) = llProject(1)        'spot count for period
                                    tmGrf.lDollars(0) = llProject(1)            'cnt gross for selected period
                                    tmGrf.lDollars(2) = llProjectSpots(1)
                                    
                                    tmGrf.lDollars(3) = tmClf.lAcquisitionCost
                                    
                                    slAmount = gLongToStrDec(llProject(1), 2)
                                    slNet = gRoundStr(gMulStr(slAmount, gSubStr("100.00", slCashAgyComm)), ".01", 2)
    
                                    'tmGrf.lDollars(2) = Val(slNet)
                                    'tmGrf.lDollars(5) = tmGrf.lDollars(3) * tmGrf.lDollars(4)   '# spots * acq cost
                                    tmGrf.lDollars(1) = Val(slNet)
                                    tmGrf.lDollars(4) = tmGrf.lDollars(2) * tmGrf.lDollars(3)   '# spots * acq cost
                              
                                    tmGrf.iVefCode = tmClf.iVefCode
                                    tmGrf.sBktType = "N"
                                    If llLineEndDate >= llStartDates(2) Then
                                        tmGrf.sBktType = "Y"
                                    End If
                                    
                                    If llWeekStart >= llLineStartDate Then
                                        'requested STart date is later than the actual line end date, show the requested start date
                                        'tmGrf.iDateGenl(0, 1) = ilWeekStart(0)
                                        'tmGrf.iDateGenl(1, 1) = ilWeekStart(1)
                                        tmGrf.iDateGenl(0, 0) = ilWeekStart(0)
                                        tmGrf.iDateGenl(1, 0) = ilWeekStart(1)
                                    Else                        'requested start date is earlier than the start of the schedule line, show shed line date
                                        'tmGrf.iDateGenl(0, 1) = tmClf.iStartDate(0)
                                        'tmGrf.iDateGenl(1, 1) = tmClf.iStartDate(1)
                                        tmGrf.iDateGenl(0, 0) = tmClf.iStartDate(0)
                                        tmGrf.iDateGenl(1, 0) = tmClf.iStartDate(1)
                                    End If
                                    If llWeekEnd <= llLineEndDate Then
                                        'requested end date is later than the actual line end date, show the requested start date
                                        'tmGrf.iDateGenl(0, 2) = ilWeekEnd(0)
                                        'tmGrf.iDateGenl(1, 2) = ilWeekEnd(1)
                                        tmGrf.iDateGenl(0, 1) = ilWeekEnd(0)
                                        tmGrf.iDateGenl(1, 1) = ilWeekEnd(1)
                                    Else                        'requested end date is later than the end of the schedule line, show shed line date
                                        'tmGrf.iDateGenl(0, 2) = tmClf.iEndDate(0)
                                        'tmGrf.iDateGenl(1, 2) = tmClf.iEndDate(1)
                                        tmGrf.iDateGenl(0, 1) = tmClf.iEndDate(0)
                                        tmGrf.iDateGenl(1, 1) = tmClf.iEndDate(1)
                                    End If
                                    
                                    '       Grf parameters:
                                    '       grfGenDAte - generation date (key)
                                    '       grfGenTime - generation time (key)
                                    '       grfchfCode - Contract Code
                                    '       grfvefcode - vehicle code
                                    '       grfadfcode - advertiser code
                                    '       grfslfcode - salesperson code
                                    '       grfSofCode - vehicle group selected (Participants, sub-company, format, etc)
                                    '       grfbkttype - Y/N - flight continues past the date of the report
                                    '       grfcode2 - vehicle group item
                                    '       grfStartDat - requested start date
                                    '       grfDate - Requested End date
                                    '       GrfPer1 - Cnt gross (requested period)
                                    '       grfPer2 - cnt net (requested period)
                                    '       grfPer3 - # spots (requested period)
                                    '       grfPer4 - Acquisition cost
                                    '       grfPer5 - # Spots * Acq cost (extended acq cost)
                                    '
                                    
                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                End If
                            End If
                        End If
                    End If
                    llProject(1) = 0            'init for next schedule line
                    llProjectSpots(1) = 0
                Next ilClf                                      'loop thru schedule lines
            End If              'lldate2 >= llweekstart and lldate <= llweekend
        End If                  'blFoundOne = true
    Next ilCurrentRecd                                      'loop for CHF records
    Erase tmChfAdvtExt, tgClf, tgCff
    Erase llStartDates, llProject, llProjectSpots
    Erase imUsevefcodes, imUseAdvtCodes, imUseSlspCodes, imUseVGCodes
    sgCntrForDateStamp = ""
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmAgf)
    Exit Sub
'gCreateMarginAcqErr:
'    sgCntrForDateStamp = ""
'    Erase tmChfAdvtExt, tgClf, tgCff
'    ilRet = btrClose(hmCHF)
'    ilRet = btrClose(hmClf)
'    ilRet = btrClose(hmCff)
'    ilRet = btrClose(hmGrf)
'    Exit Sub
End Sub
'
'               mObtainSelectivityMA - get the user selected parameters
'
Private Sub mObtainSelectivityMA()
Dim ilVGSort As Integer
Dim ilLoop As Integer

        'Selective contract #
        lmSingleCntr = Val(RptSelMA!edcContract.Text)
        
        tmCntTypes.iWorking = gSetCheck(RptSelMA!ckcProposals(0).Value)
        tmCntTypes.iComplete = gSetCheck(RptSelMA!ckcProposals(1).Value)
        tmCntTypes.iIncomplete = gSetCheck(RptSelMA!ckcProposals(2).Value)
        
        tmCntTypes.iHold = gSetCheck(RptSelMA!ckcCType(0).Value)
        tmCntTypes.iOrder = gSetCheck(RptSelMA!ckcCType(1).Value)
        tmCntTypes.iStandard = gSetCheck(RptSelMA!ckcCType(3).Value)
        tmCntTypes.iReserv = gSetCheck(RptSelMA!ckcCType(4).Value)
        tmCntTypes.iRemnant = gSetCheck(RptSelMA!ckcCType(5).Value)
        tmCntTypes.iDR = gSetCheck(RptSelMA!ckcCType(6).Value)
        tmCntTypes.iPI = gSetCheck(RptSelMA!ckcCType(7).Value)
        tmCntTypes.iPSA = gSetCheck(RptSelMA!ckcCType(8).Value)
        tmCntTypes.iPromo = gSetCheck(RptSelMA!ckcCType(9).Value)
        tmCntTypes.iTrade = gSetCheck(RptSelMA!ckcCType(10).Value)
        tmCntTypes.iPolit = gSetCheck(RptSelMA!ckcCType(2).Value)           'as previously for Feed spots
        tmCntTypes.iNonPolit = gSetCheck(RptSelMA!ckcCType(11).Value)
        
        ReDim imUseAdvtCodes(0 To 0) As Integer
        ReDim imUsevefcodes(0 To 0) As Integer
        ReDim imUseSlspCodes(0 To 0) As Integer
        gObtainCodesForMultipleLists 2, tgVehicle(), imInclVefCodes, imUsevefcodes(), RptSelMA
        gObtainCodesForMultipleLists 0, tgAdvertiser(), imInclAdvtCodes, imUseAdvtCodes(), RptSelMA
        gObtainCodesForMultipleLists 1, tgSalesperson(), imInclSlspCodes, imUseSlspCodes(), RptSelMA
        
        imMajorSet = 0
        ilVGSort = RptSelMA!cbcSet1.ListIndex
        
        If ilVGSort >= 0 And (RptSelMA!cbcSort1.ListIndex = 2 Or RptSelMA!cbcSort2.ListIndex = 2) Then
            imMajorSet = gFindVehGroupInx(ilVGSort, tgVehicleSets1())
            gObtainCodesForMultipleLists 3, tgSOCode(), imInclVGCodes, imUseVGCodes(), RptSelMA
        Else
            imInclVGCodes = 0
            ReDim imUseVGCodes(0 To 0) As Integer
        End If

        
    

End Sub
'
'                   mFilterSelectivity - test user selectivity to determine if valid contract to process
'                   <input> index of active contract array
'                   return - true if passed selectivity
'
Public Function mFilterSelectivity(ilCurrentRecd As Integer) As Boolean
Dim llContrCode As Long
Dim blValidCType As Boolean
Dim blFoundOne As Boolean
Dim ilRet As Integer
Dim ilIsItPolitical As Integer

        blFoundOne = True                              'set default to true incase by vehicle, advt should not be filtered
        If Not gFilterLists(tmChfAdvtExt(ilCurrentRecd).iAdfCode, imInclAdvtCodes, imUseAdvtCodes()) Then
            blFoundOne = False
        Else
            ilIsItPolitical = gIsItPolitical(tmChfAdvtExt(ilCurrentRecd).iAdfCode)           'its a political, include this contract?
            'test for inclusion if its political adv and politicals requested, or
            'its not a political adv and politicals
            If (tmCntTypes.iPolit And ilIsItPolitical) Or ((tmCntTypes.iNonPolit) And (Not ilIsItPolitical)) Then           'ok
                blFoundOne = blFoundOne
            Else
                blFoundOne = False
            End If
        End If
        If Not gFilterLists(tmChfAdvtExt(ilCurrentRecd).iSlfCode(0), imInclSlspCodes, imUseSlspCodes()) Then
            blFoundOne = False
        End If
        If blFoundOne Then
            'Retrieve the contract, schedule lines and flights
            llContrCode = tmChfAdvtExt(ilCurrentRecd).lCode
            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChf, tgClf(), tgCff())

            If Not ilRet Then
                On Error GoTo mFilterSelectivityErr
                gBtrvErrorMsg ilRet, "gCreateMarginAcqErr (mFilterSelectivity: gObtainCntr):" & "Chf.Btr", RptSelMA
                On Error GoTo 0
            End If
    
            blValidCType = gFilterContractType(tgChf, tmCntTypes, True)         'include proposal type checks
            If blValidCType Then                                        'test for 100% trade inclusion
                'only include trade if 100%
                If tmCntTypes.iTrade = True And tgChf.iPctTrade > 0 And tgChf.iPctTrade < 100 Then
                    blValidCType = False
                End If
            End If
        End If
            
        mFilterSelectivity = True
        If Not blFoundOne Or Not blValidCType Then
            mFilterSelectivity = False
        End If
        
        
        Exit Function
mFilterSelectivityErr:
    Resume Next
    
End Function
'
'                   gBuildFlightSpotsandRevenue - Loop through the flights of the schedule line
'                           and build the projections dollars into llproject ,
'                           and build projection # of spots into llprojectspots
'                   <input> ilclf = sched line index into tlClfInp
'                           llStdStartDates() - array of dates to build $ from flights
'                           ilFirstProjInx - index of 1st month/week to start projecting
'                           ilMaxInx - max # of buckets to loop thru
'                   <output>
'                   General routine to build flight $/cpot count into week, month, qtr buckets
'            Created : 7-12-05
'
Public Sub mBuildFlightSpotsAndRevenue(ilClf As Integer, llStdStartDates() As Long, llWeekStart As Long, llWeekEnd As Long, slCashAgyComm As String)
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
Dim blAcqOK As Boolean
Dim ilAcqLoInx As Integer
Dim ilAcqHiInx As Integer
Dim ilAcqCommPct As Integer
Dim llAcqComm As Long
Dim llAcqNet As Long
Dim llAcquisitionCost As Long
Dim slAmount As String
Dim slNet As String
Dim llLineEndDate As Long
Dim llLineStartDate As Long
Dim llProject As Long
Dim llProjectSpots As Long
Dim ilWeekStart(0 To 1) As Integer
Dim ilWeekEnd(0 To 1) As Integer
Dim slWeekStart As String
Dim slWeekEnd As String
Dim ilRet As Integer


    gUnpackDateLong tmClf.iEndDate(0), tmClf.iEndDate(1), llLineEndDate     'line end date
    gUnpackDateLong tmClf.iStartDate(0), tmClf.iStartDate(1), llLineStartDate    'Line start date

    slWeekStart = Format$(llWeekStart, "m/d/yy")
    gPackDate slWeekStart, ilWeekStart(0), ilWeekStart(1)  'btrieve format for prepas record

    slWeekEnd = Format$(llWeekEnd, "m/d/yy")
    gPackDate slWeekEnd, ilWeekEnd(0), ilWeekEnd(1)  'btrieve format for prepas record

    llStdStart = llStdStartDates(1)
    llStdEnd = llStdStartDates(2)
    
    'blAcqOK = gGetAcqCommInfoByVehicle(tlClfInp(ilClf).ClfRec.iVefCode, ilAcqLoInx, ilAcqHiInx) 'determine the starting and ending indices of acq percents for this lines vehicle
    blAcqOK = gGetAcqCommInfoByVehicle(tgClf(ilClf).ClfRec.iVefCode, ilAcqLoInx, ilAcqHiInx) 'determine the starting and ending indices of acq percents for this lines vehicle
    
    ilCff = tgClf(ilClf).iFirstCff
    'ilCff = tlClfInp(ilClf).iFirstCff
    Do While ilCff <> -1
        'tlCff = tlCffInp(ilCff).CffRec
        tlCff = tgCff(ilCff).CffRec
        llWhichRate = tlCff.lActPrice
       
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

         ilAcqCommPct = gGetEffectiveAcqComm(llFltStart, ilAcqLoInx, ilAcqHiInx)     'if varying commissions for acq costs on Insertion order, get the % to be used to calc.
         'gCalcAcqComm ilAcqCommPct, tlClfInp(ilClf).ClfRec.lAcquisitionCost, llAcqNet, llAcqComm
         gCalcAcqComm ilAcqCommPct, tgClf(ilClf).ClfRec.lAcquisitionCost, llAcqNet, llAcqComm
                
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

            llProjectSpots = 0
            llProject = 0
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
                
                'determine month that this week belongs in, then accumulate the gross and net $
                'currently, the projections are based on STandard bdcst
                If llDate >= llStdStartDates(1) And llDate < llStdStartDates(2) Then
                    llProject = llProject + (llSpots * llWhichRate)
                    llProjectSpots = llProjectSpots + llSpots
                End If
            Next llDate                                     'for llDate = llFltStart To llFltEnd
            
            If (llProject <> 0 Or tmClf.lAcquisitionCost <> 0) And llProjectSpots > 0 Then        'must have either true spot rate or acq $, along with a spot count
                                         
                  'tmGrf.lDollars(1) = llProject        'spot count for period
                  'tmGrf.lDollars(1) = llProject            'cnt gross for selected period
                  'tmGrf.lDollars(3) = llProjectSpots
                  
                  'tmGrf.lDollars(4) = llAcqNet
                  tmGrf.lDollars(0) = llProject        'spot count for period
                  tmGrf.lDollars(0) = llProject            'cnt gross for selected period
                  tmGrf.lDollars(2) = llProjectSpots
                  
                  tmGrf.lDollars(3) = llAcqNet
                  
                  slAmount = gLongToStrDec(llProject, 2)
                  slNet = gRoundStr(gMulStr(slAmount, gSubStr("100.00", slCashAgyComm)), ".01", 2)
    
                  'tmGrf.lDollars(2) = Val(slNet)
                  'tmGrf.lDollars(5) = tmGrf.lDollars(3) * tmGrf.lDollars(4)   '# spots * acq cost
                  tmGrf.lDollars(1) = Val(slNet)
                  tmGrf.lDollars(4) = tmGrf.lDollars(2) * tmGrf.lDollars(3)   '# spots * acq cost
            
                  tmGrf.iVefCode = tmClf.iVefCode
                  tmGrf.sBktType = "N"
                  If llLineEndDate >= llStdStartDates(2) Then
                      tmGrf.sBktType = "Y"
                  End If
                  
                  If llWeekStart >= llLineStartDate Then
                      'requested STart date is later than the actual line end date, show the requested start date
                      'tmGrf.iDateGenl(0, 1) = ilWeekStart(0)
                      'tmGrf.iDateGenl(1, 1) = ilWeekStart(1)
                      tmGrf.iDateGenl(0, 0) = ilWeekStart(0)
                      tmGrf.iDateGenl(1, 0) = ilWeekStart(1)
                  Else                        'requested start date is earlier than the start of the schedule line, show shed line date
                      'tmGrf.iDateGenl(0, 1) = tmClf.iStartDate(0)
                      'tmGrf.iDateGenl(1, 1) = tmClf.iStartDate(1)
                      tmGrf.iDateGenl(0, 0) = tmClf.iStartDate(0)
                      tmGrf.iDateGenl(1, 0) = tmClf.iStartDate(1)
                  End If
                  If llWeekEnd <= llLineEndDate Then
                      'requested end date is later than the actual line end date, show the requested start date
                      'tmGrf.iDateGenl(0, 2) = ilWeekEnd(0)
                      'tmGrf.iDateGenl(1, 2) = ilWeekEnd(1)
                      tmGrf.iDateGenl(0, 1) = ilWeekEnd(0)
                      tmGrf.iDateGenl(1, 1) = ilWeekEnd(1)
                  Else                        'requested end date is later than the end of the schedule line, show shed line date
                      'tmGrf.iDateGenl(0, 2) = tmClf.iEndDate(0)
                      'tmGrf.iDateGenl(1, 2) = tmClf.iEndDate(1)
                      tmGrf.iDateGenl(0, 1) = tmClf.iEndDate(0)
                      tmGrf.iDateGenl(1, 1) = tmClf.iEndDate(1)
                  End If
                  
                  '       Grf parameters:
                  '       grfGenDAte - generation date (key)
                  '       grfGenTime - generation time (key)
                  '       grfchfCode - Contract Code
                  '       grfvefcode - vehicle code
                  '       grfadfcode - advertiser code
                  '       grfslfcode - salesperson code
                  '       grfSofCode - vehicle group selected (Participants, sub-company, format, etc)
                  '       grfbkttype - Y/N - flight continues past the date of the report
                  '       grfcode2 - vehicle group item
                  '       grfStartDat - requested start date
                  '       grfDate - Requested End date
                  '       GrfPer1 - Cnt gross (requested period)
                  '       grfPer2 - cnt net (requested period)
                  '       grfPer3 - # spots (requested period)
                  '       grfPer4 - Acquisition cost
                  '       grfPer5 - # Spots * Acq cost (extended acq cost)
                  '
                  
                  ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            End If
        End If                                          '
        'ilCff = tlCffInp(ilCff).iNextCff                   'get next flight record from mem
        ilCff = tgCff(ilCff).iNextCff
    Loop                                            'while ilcff <> -1
    Exit Sub
End Sub
