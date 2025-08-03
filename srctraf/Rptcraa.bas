Attribute VB_Name = "RPTCRAA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcraa.bas on Wed 6/17/09 @ 12:56 PM
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
Dim tlChfAdvtExt() As CHFADVTEXT
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmUrf As Integer            'User file handle
Dim imUrfRecLen As Integer      'URF record length
Dim tmUrf As URF
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim tmCntTypes As CNTTYPES          'image for user requested parameters

'  Receivables File
'********************************************************************************************
'
'           gCrClientRecap - Prepass to create Advertisers Airing
'           Gather all contracts that are valid airing dates for the user dates entered
'           and show the advertisers/contracts for the period
'
'           D.hosaka   2/20/01
'
'       Grf parameters:
'       grfGenDAte - generation date (key)
'       grfGenTime - generation time (key)
'       grfchfCode - Contract Code
'       grfPerGenl(1) - vehicle group type (1 = participants, 2 = sub-totals 3 =  market
'                                           4 = format 5 = research)
'       grfPerGenl(2) - mnf code (for vehicle group)
'       grfDollars(1) - spot count
'
'
'********************************************************************************************
Sub gCrClientRecap()
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
Dim ilFoundOne As Integer               'Found a matching  office built into mem
Dim ilValidVehicle As Integer
Dim ilTemp As Integer
Dim ilLoop As Integer                   'temp loop variable
'ReDim llProject(1 To 4) As Long        '$ projected for 4 quarters
ReDim llProject(0 To 4) As Long        '$ projected for 4 quarters. Index zero ignored
Dim llDate As Long                      'temp date variable
Dim llDate2 As Long
Dim slNameCode As String
Dim slCode As String
Dim ilmnfMinorCode As Integer
Dim ilMajorSet As Integer
ReDim ilCodes(0 To 0) As Integer     'array of valid advertisrs/vehicles to gather
ReDim ilMnfcodes(0 To 0) As Integer     'array of valid vehicle groups to gather
Dim ilCkcAll As Integer                 'All vehicle groups selcted, force to true if no vehicle group (NONE) selected
Dim ilCkcAllAdvVeh As Integer              'all advt or vehicles selected

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
'Month Starts to gather projection $ from flights
'ReDim llStartDates(1 To 2) As Long        'start dates of each period to gather (only one)
ReDim llStartDates(0 To 2) As Long        'start dates of each period to gather (only one). Index zero ignored
Dim ilSelectionIndex As Integer             'index into lbcSelection array
Dim blInclude As Boolean

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
    ReDim tgClfAA(0 To 0) As CLFLIST
    tgClfAA(0).iStatus = -1 'Not Used
    tgClfAA(0).lRecPos = 0
    tgClfAA(0).iFirstCff = -1
    ReDim tgCffAA(0 To 0) As CFFLIST
    tgCffAA(0).iStatus = -1 'Not Used
    tgCffAA(0).lRecPos = 0
    tgCffAA(0).iNextCff = -1
    ilLoop = RptSelAA!cbcSet.ListIndex
    ilMajorSet = gFindVehGroupInx(ilLoop, tgVehicleSets1())
    If ilLoop = 0 Then              'no vehicle group selected
        ilCkcAll = True '9-13-02vbChecked             'force everything included
    Else
        If (RptSelAA!ckcAll.Value = vbChecked) Then
            ilCkcAll = True '9-13-02 vbChecked
        Else
            ilCkcAll = False    '9-13-02vbUnchecked
        End If
    End If
    
    ilCkcAllAdvVeh = False
    If RptSelAA!rbcSortBy(0).Value Then                 'advt sort
        ilSelectionIndex = 1                                'selection index into array
        If (RptSelAA!ckcAllAdv.Value = vbChecked) Then    'all advertisrs selected?
            ilCkcAllAdvVeh = True
         End If
    Else
        ilSelectionIndex = 2                            'selection index into array
        If (RptSelAA!ckcAllVehicles.Value = vbChecked) Then    'all vehicles selected?
            ilCkcAllAdvVeh = True
        End If
    End If
    
    'Get STart and end dates of user requested date to find all advertisers airing
'    slWeekStart = RptSelAA!edcSelCFrom.Text
    slWeekStart = RptSelAA!CSI_CalFrom.Text '8-14-19
    llWeekStart = gDateValue(slWeekStart)
    'make sure the date is a Monday
    ilLoop = gWeekDayLong(llWeekStart)
    Do While ilLoop <> 0
        llWeekStart = llWeekStart - 1
        ilLoop = gWeekDayLong(llWeekStart)
    Loop
    slWeekStart = Format$(llWeekStart, "m/d/yy")
    gPackDate slWeekStart, ilWeekStart(0), ilWeekStart(1)    'conversion to store in prepass record
    ilNoWeeks = Val(RptSelAA!edcSelCTo.Text)
    llWeekEnd = llWeekStart + ((ilNoWeeks - 1) * 7) + 6
    slWeekEnd = Format$(llWeekEnd, "m/d/yy")
    gPackDate slWeekEnd, ilWeekEnd(0), ilWeekEnd(1)  'btrieve format for prepas record

    'Start date of eachperiod to accumulate spot counts for (only one period applicable)
    llStartDates(1) = llWeekStart
    llStartDates(2) = llWeekEnd + 1
    'Gather all contracts for previous year and current year whose effective date entered
    'is prior to the effective date that affects either previous year or current year
    'slCntrTypes = gBuildCntTypes()
    'slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
    
    slCntrStatus = ""                 'statuses: hold, order, unsch hold, uns order
    If RptSelAA!ckcCType(0).Value = vbChecked Then          'include sched/unsch holds
        slCntrStatus = "HG"             'include orders and uns orders
    End If
    If RptSelAA!ckcCType(1).Value = vbChecked Then          'include sched/unsch order
        slCntrStatus = slCntrStatus & "ON"             'include orders and uns orders
    End If
    slCntrTypes = ""
    If RptSelAA!ckcCType(3).Value = vbChecked Then
        slCntrTypes = "C"
    End If
    If RptSelAA!ckcCType(4).Value = vbChecked Then
        slCntrTypes = slCntrTypes & "V"
    End If
    If RptSelAA!ckcCType(5).Value = vbChecked Then
        slCntrTypes = slCntrTypes & "T"
    End If
    If RptSelAA!ckcCType(6).Value = vbChecked Then
        slCntrTypes = slCntrTypes & "R"
    End If
    If RptSelAA!ckcCType(7).Value = vbChecked Then
        slCntrTypes = slCntrTypes & "Q"
    End If
    If RptSelAA!ckcCType(8).Value = vbChecked Then
        slCntrTypes = slCntrTypes & "S"
    End If
    If RptSelAA!ckcCType(9).Value = vbChecked Then
        slCntrTypes = slCntrTypes & "M"
    End If
    If slCntrTypes = "CVTRQSM" Then          'all types: PI, DR, etc.  except PSA(p) and Promo(m)
        slCntrTypes = ""                     'blank out string for "All"
    End If

    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)

    tmCntTypes.iHold = gSetCheck(RptSelAA!ckcCType(0).Value)
    tmCntTypes.iOrder = gSetCheck(RptSelAA!ckcCType(1).Value)
    tmCntTypes.iStandard = gSetCheck(RptSelAA!ckcCType(3).Value)
    tmCntTypes.iReserv = gSetCheck(RptSelAA!ckcCType(4).Value)
    tmCntTypes.iRemnant = gSetCheck(RptSelAA!ckcCType(5).Value)
    tmCntTypes.iDR = gSetCheck(RptSelAA!ckcCType(6).Value)
    tmCntTypes.iPI = gSetCheck(RptSelAA!ckcCType(7).Value)
    tmCntTypes.iPSA = gSetCheck(RptSelAA!ckcCType(8).Value)
    tmCntTypes.iPromo = gSetCheck(RptSelAA!ckcCType(9).Value)
    tmCntTypes.iTrade = gSetCheck(RptSelAA!ckcCType(10).Value)
    tmCntTypes.iPolit = gSetCheck(RptSelAA!ckcCType(2).Value)           'as previously for Feed spots
    tmCntTypes.iNonPolit = gSetCheck(RptSelAA!ckcCType(11).Value)
    
    'common GRF fields that wont change
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)
    'tmGrf.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    tmGrf.iStartDate(0) = ilWeekStart(0)
    tmGrf.iStartDate(1) = ilWeekStart(1)
    tmGrf.iDate(0) = ilWeekEnd(0)
    tmGrf.iDate(1) = ilWeekEnd(1)
    'tmGrf.iPerGenl(1) = ilMajorSet      'used for heading in Crystal report
    tmGrf.iPerGenl(0) = ilMajorSet      'used for heading in Crystal report

    'If RptSelAA.rbcSortBy(0).Value Then             'advt sort
        If Not (ilCkcAllAdvVeh) Then    'build array of the selected advertisers or vehicles
            For ilTemp = 0 To RptSelAA!lbcSelection(ilSelectionIndex).ListCount - 1 Step 1
                If RptSelAA!lbcSelection(ilSelectionIndex).Selected(ilTemp) Then
                    If RptSelAA!rbcSortBy(0).Value Then         'advt
                        slNameCode = tgAdvertiser(ilTemp).sKey
                    Else
                        slNameCode = tgVehicle(ilTemp).sKey
                    End If
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    ilCodes(UBound(ilCodes)) = Val(slCode)
                    ReDim Preserve ilCodes(0 To UBound(ilCodes) + 1)
                End If
            Next ilTemp
        End If
   ' End If
    If Not (ilCkcAll) Then            'build array of selected vehicle groups
        For ilTemp = 0 To RptSelAA!lbcSelection(0).ListCount - 1 Step 1
            If RptSelAA!lbcSelection(0).Selected(ilTemp) Then
                slNameCode = tgSOCodeAA(ilTemp).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilMnfcodes(UBound(ilMnfcodes)) = Val(slCode)
                ReDim Preserve ilMnfcodes(0 To UBound(ilMnfcodes) + 1)
            End If
        Next ilTemp
    End If

    ilRet = gObtainCntrForDate(RptSelAA, slWeekStart, slWeekEnd, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())

    'All contracts have been retrieved for all of this year
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1

         ilFoundOne = True                              'set default to true incase by vehicle, advt should not be filtered
         If RptSelAA!rbcSortBy(0).Value Then
             If Not ilCkcAllAdvVeh Then                 '9-13-02
                ilFoundOne = False
                For ilLoop = 0 To UBound(ilCodes) - 1
                    If tlChfAdvtExt(ilCurrentRecd).iAdfCode = ilCodes(ilLoop) Then
                        ilFoundOne = True
                        Exit For
                    End If
                Next ilLoop
            Else
                ilFoundOne = True
            End If
        End If
        If ilFoundOne Then
            'Retrieve the contract, schedule lines and flights
            llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfAA, tgClfAA(), tgCffAA())

            If Not ilRet Then
                On Error GoTo gCrClientRecapErr
                gBtrvErrorMsg ilRet, "gCrClientRecapErr (btrExtGetNextExt):" & "Chf.Btr", RptSelAA
                On Error GoTo 0
            End If
            
            'contract obtained, determine if a selected contract type, along with trade and polit/non-polit
            blInclude = mFilterSelectivity()
            If blInclude Then
                'determine if the contracts start & end dates fall within the requested period
                gUnpackDateLong tgChfAA.iEndDate(0), tgChfAA.iEndDate(1), llDate2      'hdr end date converted to long
                gUnpackDateLong tgChfAA.iStartDate(0), tgChfAA.iStartDate(1), llDate    'hdr start date converted to long
                If llDate2 >= llWeekStart And llDate <= llWeekEnd Then
                    For ilClf = LBound(tgClfAA) To UBound(tgClfAA) - 1 Step 1
                        tmClf = tgClfAA(ilClf).ClfRec
                        'Project the monthly spots from the flights
                        If tmClf.sType = "S" Or tmClf.sType = "H" Then
    
                            'Setup the major sort factor
                            'gGetVehGrpSets tmClf.iVefCode, 0, ilMajorSet, ilmnfMinorCode, tmGrf.iPerGenl(2)
                            gGetVehGrpSets tmClf.iVefCode, 0, ilMajorSet, ilmnfMinorCode, tmGrf.iPerGenl(1)
                            'check selectivity of vehicle groups
                            '9-13-02 If Not ilCkcAll = vbChecked Or (ilMajorSet > 0 And Not ilCkcAll = vbChecked) Then
                                                 
                            If Not ilCkcAll Or (ilMajorSet > 0 And Not ilCkcAll) Then     '9-13-02
                                ilFoundOne = False
                                    'ilLoop = gBinarySearchVef(tmClf.iVefCode)
                                    'If ilLoop <> -1 Then
                                    '    If ilMajorSet = 1 Then
                                    '        ilWhichset = tgMVef(ilLoop).iMnfGroup(1)
                                    '    ElseIf ilMajorSet = 2 Then
                                    '        ilWhichset = tgMVef(ilLoop).iMnfVehGp2
                                    '    ElseIf ilMajorSet = 3 Then
                                    '        ilWhichset = tgMVef(ilLoop).iMnfVehGp3Mkt
                                    '    ElseIf ilMajorSet = 4 Then
                                    '        ilWhichset = tgMVef(ilLoop).iMnfVehGp4Fmt
                                    '    ElseIf ilMajorSet = 5 Then
                                    '        ilWhichset = tgMVef(ilLoop).iMnfVehGp5Rsch
                                    '    ElseIf ilMajorSet = 6 Then
                                    '        ilWhichset = tgMVef(ilLoop).iMnfVehGp6Sub
                                    '    End If
    
                                        If Not (ilCkcAll) Then
                                            For ilTemp = 0 To UBound(ilMnfcodes) - 1
                                                ''If ilMnfCodes(ilTemp) = ilWhichset Then
                                                'If ilMnfcodes(ilTemp) = tmGrf.iPerGenl(2) Then
                                                If ilMnfcodes(ilTemp) = tmGrf.iPerGenl(1) Then
                                                    ilFoundOne = True
                                                    Exit For
                                                End If
                                            Next ilTemp
                                            If ilFoundOne Then
                                '                Exit For
                                            End If
                                        Else
                                            ilFoundOne = True
                                        End If
                                    'End If
                            Else
                                ilFoundOne = True
                            End If
                            
                            ilValidVehicle = True
                            If RptSelAA!rbcSortBy(1).Value Then            'sort by vehicle?
                                If Not ilCkcAllAdvVeh Then
                                    ilValidVehicle = False
                                    For ilLoop = 0 To UBound(ilCodes) - 1
                                        If tmClf.iVefCode = ilCodes(ilLoop) Then
                                            ilValidVehicle = True
                                            Exit For
                                        End If
                                    Next ilLoop
                                End If
                            End If
                            
                            If (ilFoundOne) And (ilValidVehicle) Then      'got a valid vehicle group code
                                gBuildFlightSpots ilClf, llStartDates(), 1, 2, llProject(), 1, tgClfAA(), tgCffAA()
    
                                'Setup the major sort factor
                                'gGetVehGrpSets tmClf.iVefCode, 0, ilMajorSet, ilmnfMinorCode, tmGrf.iPerGenl(2)
                                tmGrf.lChfCode = tgChfAA.lCode
                                'tmGrf.lDollars(1) = llProject(1)        'spot count for period
                                tmGrf.lDollars(0) = llProject(1)        'spot count for period
                                tmGrf.iVefCode = tmClf.iVefCode
                                'If tmGrf.lDollars(1) > 0 Then
                                If tmGrf.lDollars(0) > 0 Then
                                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                End If
                            End If
                        End If
                        llProject(1) = 0            'init for next schedule line
                    Next ilClf                                      'loop thru schedule lines
                End If              'lldate2 >= llweekstart and lldate <= llweekend
            End If              'blinclude
        End If                  'ilfoundone = true
    Next ilCurrentRecd                                      'loop for CHF records
    Erase tlChfAdvtExt, tgClfAA, tgCffAA
    Erase llStartDates, llProject
    sgCntrForDateStamp = ""
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmUrf)
    ilRet = btrClose(hmGrf)
    Exit Sub
gCrClientRecapErr:
    sgCntrForDateStamp = ""
    Erase tlChfAdvtExt, tgClfAA, tgCffAA
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmUrf)
    ilRet = btrClose(hmGrf)
    Exit Sub
End Sub
'
'                   mFilterSelectivity - test user selectivity to determine if valid contract to process
'                   <input>  none
'                   return - true if passed selectivity
' tgChfAA buffer assumed to have contract
Public Function mFilterSelectivity() As Boolean
Dim llContrCode As Long
Dim blValidCType As Boolean
Dim blFoundOne As Boolean
Dim ilRet As Integer
Dim ilIsItPolitical As Integer

        blFoundOne = True
        ilIsItPolitical = gIsItPolitical(tgChfAA.iAdfCode)           'its a political, include this contract?
        'test for inclusion if its political adv and politicals requested, or
        'its not a political adv and politicals
        If (tmCntTypes.iPolit And ilIsItPolitical) Or ((tmCntTypes.iNonPolit) And (Not ilIsItPolitical)) Then           'ok
            blFoundOne = blFoundOne
        Else
            blFoundOne = False
        End If

        blValidCType = gFilterContractType(tgChfAA, tmCntTypes, False)         'exclude proposal type checks
        If blValidCType Then                                        'test for 100% trade inclusion
            'only include trade if 100%
            'If (tmCntTypes.iTrade = True And tgChfAA.iPctTrade > 0 And tgChfAA.iPctTrade < 100) Or (tmCntTypes.iTrade = False And tgChfAA.iPctTrade <> 0) Then
            'change to include trades if any part is trade
            If Not (tmCntTypes.iTrade) And tgChfAA.iPctTrade <> 0 Then
                blValidCType = False
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
