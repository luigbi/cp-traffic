Attribute VB_Name = "RPTCRSP"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrsp.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
'Public igYear As Integer                'budget year used for filtering
'Public igMonthOrQtr As Integer          'entered month or qtr
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
'Quarterly Avails

'3-11-13 feature added to use avg 30" spot price to calc inv valuation, along with following 2 adjustment factors
Dim imRCvsAvgPrice As Integer       '0 = Use R/C rates to calc inv val; 1 = use Avg Price to calc inv val
Dim imPctChg As Integer             '+/- percent change to adjust price used to calc inv valuation
Dim imPctSellout As Integer         'est percent sellout against calc inv valuation

Dim tmAvr() As AVR                'AVR record image
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim tmBvf As BVF                  'Budgets by office & vehicle
Dim hmBvf As Integer
Dim imBvfRecLen As Integer        'BVF record length
Dim tmBvfVeh() As BVF               'Budget by vehicle
Dim hmMnf As Integer                'Multiname file handle
Dim imMnfRecLen As Integer          'MNF record length
Dim tmMnf As MNF
Dim tlMMnf() As MNF                    'array of MNF records for specific type
'  Receivables File
Dim tmRvf As RVF            'RVF record image
Dim tlSlsList() As SLSLIST      'Sales Analysis Summary
Dim tmSpotLenRatio As SPOTLENRATIO
'
'
'                   Create Sales vs Plan prepass file
'                   Generate GRF file by vehicle.  Each record  contains the vehicle,
'                   plan $, Business on Books for current years Qtr, OOB w/ holds for
'                   current years qtr
'
'                   Created: 10/29/97 D. Hosaka
'                   4/8/98 Ignore 100% trade contracts
'       2-11-05 Chg array of PHF/RVF array from integer to long (prevent overflow)
'       12-14-06 add parm to gObtainRvfPhf to test on tran date (vs entry date)
Sub gCRSalesPlan()
Dim llDate As Long
Dim slAirOrder As String * 1                'from site pref - bill as air or ordered
Dim ilLoop As Integer
Dim ilTemp As Integer
Dim ilSlsLoop As Integer
Dim ilRet As Integer
ReDim ilEnterDate(0 To 1) As Integer        'Btrieve format for date entered by user
Dim slDate As String
Dim slStr As String
Dim ilYear As Integer
Dim llEnterTo As Long
Dim slEnterTo As String                       'effective date entere:  tested against the contract header entry date
'ReDim llProject(1 To 2) As Long               'projected $, only using 1 bucket, common rtn needs assumes array
ReDim llProject(0 To 2) As Long               'projected $, only using 1 bucket, common rtn needs assumes array. Index zero ignored
'ReDim llTYDates(1 To 2) As Long               'range of qtr dates for contract retrieval (last year)
ReDim llTYDates(0 To 2) As Long               'range of qtr dates for contract retrieval (last year). Index zero ignored
Dim llTYGetFrom As Long                       'range of dates for contract access (last year)
Dim ilBdMnfCode As Integer                      'budget name to get
Dim ilBdYear As Integer                         'budget year to get
Dim slNameCode As String
Dim slYear As String
Dim slCode As String
Dim slCntrTypes As String                       'valid contract types to access
Dim slCntrStatus As String                      'valid status (holds, orders, working, etc) to access
Dim ilHOState As Integer                        'states to use in contract gathering:  use unsch holds/orders and sch holds/orders
Dim ilFound As Integer
Dim ilStartWk As Integer                        'starting week index to gather budget data
Dim ilEndWk As Integer                          'ending week index to gather budgets
Dim ilFirstWk As Integer                        'true if week 0 needs to be added when start wk = 1
Dim ilLastWk As Integer                         'true if week 53 needs to be added when end wk = 52
Dim llContrCode As Long                         'contr code from gObtainCntrforDate
Dim ilCurrentRecd As Integer            'index to contract processed from tlChfAdvtext
Dim ilClf As Integer                    'index for line to be processed from tgClfSP
Dim ilCorpStd As Integer                '1 = corp, 2 = std
Dim ilBvfCalType As Integer             '0=std, 1 = reg, 2 & 3 = julian, 4 = corp for jan thru dec, 5 = corp for fiscal year
Dim ilFoundSls As Integer
Dim tlChfAdvtExt() As CHFADVTEXT        'array of potential valid contracts to process based on header active dates and qtr requested
Dim slStartQtr As String                'start date of qtr for std or corp cal
Dim slEndQtr As String                  'end date of qtr for std or corp cal
Dim slStartYr As String                 'start date of yr for std or corp cal
Dim slEndYr As String                   'end date of yr for std or corp cal
Dim tlCntTypes As CNTTYPES              'structure of contrct spot types to include
Dim ilVehicle As Integer
Dim ilVefFound As Integer
Dim ilMinorSet As Integer
Dim ilMajorSet As Integer
Dim ilmnfMinorCode As Integer           'field used to sort the minor sort with
Dim ilMnfMajorCode As Integer           'field used to sort the major sort with
Dim llTemp1 As Long
Dim llTemp2 As Long
'ReDim ilVefList(1 To 1) As Integer        'list of selected veh codes fromlist box (to avoid parsing with every schline)
ReDim ilVefList(0 To 0) As Integer        'list of selected veh codes fromlist box (to avoid parsing with every schline)
Dim tlTranType As TRANTYPES
'ReDim tlRvf(1 To 1) As RVF
ReDim tlRvf(0 To 0) As RVF
Dim llRvfLoop As Long                   '2-11-05
Dim tlValuationInfo As VALUATIONINFO
Dim ilDone As Integer
ReDim ilAnfCodes(0 To 0) As Integer
Dim slLen(0 To 9) As String
Dim slIndex(0 To 9) As String

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmGrf
        btrDestroy hmCHF
    Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
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
        ilRet = btrClose(hmCHF)
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imBvfRecLen = Len(tmBvf)

    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmBvf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmMnf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    'tgMMnf contains the Vehicle Groups (use Ubound(tgMMnf))
    ilLoop = RptSelSP!cbcSet1.ListIndex
    ilMajorSet = tgVehicleSets1(ilLoop).iCode

    ilLoop = RptSelSP!cbcSet2.ListIndex
    ilMinorSet = tgVehicleSets2(ilLoop).iCode
    '***********   Temporarily setup defaults for sorting use SET 1  until ABC approves
    '
    'ilMajorSet = 1
    'ilMinorSet = 0
    '
    '***********
    
    
    tlValuationInfo.iRCvsAvgPrice = 0           'assume to run by Rate card
    If RptSelSP!rbcValRate(1).Value = True Then 'run by avg 30" spot price
        tlValuationInfo.iRCvsAvgPrice = 1
    End If
    
    tlValuationInfo.sBaseReptFlag = "B"         'get base dayparts only

    tlCntTypes.iHold = True
    tlCntTypes.iOrder = True
    tlCntTypes.iStandard = True
    tlCntTypes.iReserv = True
    tlCntTypes.iRemnant = True
    tlCntTypes.iDR = True
    tlCntTypes.iPI = True
    tlCntTypes.iTrade = False           'chged 7/20/98
    tlCntTypes.iMissed = True
    tlCntTypes.iNC = True
    tlCntTypes.iXtra = True
    tlCntTypes.iPSA = False
    tlCntTypes.iPromo = False
    
     '4-12-18 No selectivity, take the defaults from Site
    If tgSpf.sCIncludeMissDB <> "N" Then
        tlCntTypes.iMissed = True
    Else
        tlCntTypes.iMissed = False
    End If
    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDERESERVATION) <> AVAILINCLUDERESERVATION Then
        tlCntTypes.iReserv = False
    End If
    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEREMNANT) <> AVAILINCLUDEREMNANT Then
        tlCntTypes.iRemnant = False
    End If
    
    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLDEDIRECTRESPONSES) <> AVAILINCLDEDIRECTRESPONSES Then
        tlCntTypes.iDR = False
    End If
    
    If (Asc(tgSaf(0).sFeatures4) And AVAILINCLUDEPERINQUIRY) <> AVAILINCLUDEPERINQUIRY Then
        tlCntTypes.iPI = False
    End If
    
    For ilLoop = 0 To 9
        slLen(ilLoop) = Trim$(RptSelSP!edcLen(ilLoop))
        slIndex(ilLoop) = Trim$(RptSelSP!edcIndex(ilLoop))
    Next ilLoop
    gBuildSpotLenAndIndexTable slLen(), slIndex(), tmSpotLenRatio

    
'    'build the spot lengths lo to hi order, along with its associated index value
'    For ilLoop = 0 To 9
'        tmSpotLenRatio.iLen(ilLoop) = Val(RptSelSP!edcLen(ilLoop))
'        slStr = RptSelSP!edcIndex(ilLoop)
'        gFormatStr slStr, 0, 2, slStr
'        tmSpotLenRatio.iRatio(ilLoop) = gStrDecToInt(slStr, 2)
'    Next ilLoop
'    ilDone = False
'    Do While Not ilDone
'        ilDone = True
'        For ilLoop = 1 To 9
'            If tmSpotLenRatio.iLen(ilLoop - 1) > tmSpotLenRatio.iLen(ilLoop) And tmSpotLenRatio.iLen(ilLoop) > 0 Then
'                'swap the two lengths
'                ilRet = tmSpotLenRatio.iLen(ilLoop - 1)
'                tmSpotLenRatio.iLen(ilLoop - 1) = tmSpotLenRatio.iLen(ilLoop)
'                tmSpotLenRatio.iLen(ilLoop) = ilRet
'
'                'swap the two index values
'                ilRet = tmSpotLenRatio.iRatio(ilLoop - 1)
'                tmSpotLenRatio.iRatio(ilLoop - 1) = tmSpotLenRatio.iRatio(ilLoop)
'                tmSpotLenRatio.iRatio(ilLoop) = ilRet
'
'                ilDone = False
'            End If
'        Next ilLoop
'    Loop
'
'    slStr = ""
'    For ilLoop = 0 To 9
'        If tmSpotLenRatio.iLen(ilLoop) = 0 Then         'done
'            Exit For
'        Else
'            slNameCode = gIntToStrDec(tmSpotLenRatio.iRatio(ilLoop), 2)
'            If Trim$(slStr) <> "" Then
'                slStr = slStr & ","
'            End If
'            slStr = slStr & Str$(tmSpotLenRatio.iLen(ilLoop)) & " @" & Trim$(slNameCode)
'        End If
'    Next ilLoop
    'show in report header
    If Not gSetFormula("UserRatios", "'" & slStr & "'") Then
        Exit Sub
    End If

    '% +/- change of unsold spot price
    slStr = RptSelSP!edcPctChg
    'gFormatStr slStr, FMTNEGATBACK, 0, slStr
    tlValuationInfo.iUnsoldPctAdj = Val(slStr)
    If tlValuationInfo.iUnsoldPctAdj = 0 Then
        tlValuationInfo.iUnsoldPctAdj = 100
    ElseIf tlValuationInfo.iUnsoldPctAdj < 0 Then
        tlValuationInfo.iUnsoldPctAdj = tlValuationInfo.iUnsoldPctAdj + 100
    End If
    
    'est % sellout of unsold avails
    slStr = RptSelSP!edcEstPct
    tlValuationInfo.iEstPctSellout = Val(slStr)
    If tlValuationInfo.iEstPctSellout = 0 Then
        tlValuationInfo.iEstPctSellout = 100
    End If

    slAirOrder = tgSpf.sInvAirOrder     'inv all contracts as aired or ordered
    'get all the dates needed to work with
'    slDate = RptSelSP!edcSelCFrom.Text               'effective date entred
    slDate = RptSelSP!CSI_CalFrom.Text               '12-11-19 change to use csi calendar control, effective date entred
    'obtain the entered dates year based on the std month
    llEnterTo = gDateValue(slDate)                     'gather contracts thru this entered date
    slEnterTo = Format$(llEnterTo, "m/d/yy")           'insure the year is formatted from input
    gPackDateLong llEnterTo, ilEnterDate(0), ilEnterDate(1)    'get btrieve date format for entered to pass to record to show on hdr

    ilYear = Val(RptSelSP!edcSelCTo.Text)           'year requested
    If RptSelSP!rbcSelCSelect(0).Value Then         'corp
        ilRet = gGetCorpCalIndex(ilYear)
        'gUnpackDate tgMCof(ilRet).iStartDate(0, 1), tgMCof(ilRet).iStartDate(1, 1), slTYStart         'convert last bdcst billing date to string
        'gUnpackDate tgMCof(ilRet).iEndDate(0, 12), tgMCof(ilRet).iEndDate(1, 12), slTYEnd
        ilCorpStd = 1
        ilBvfCalType = 5               'get week inx based on fiscal dates
    Else                                'std
        ilCorpStd = 2
        ilBvfCalType = 0             'both projections and budgets will be std
        'slTYStart = "1/15/" & Trim$(Str$(ilYear))
        'slTYStart = gObtainStartStd(slTYStart)              'obtain start and end dates of current std year
        'slTYEnd = "12/15/" & Trim$(Str$(ilYear))
        'slTYEnd = gObtainEndStd(slTYEnd)
    End If
    gGetStartEndQtr ilCorpStd, ilYear, igMonthOrQtr, slStartQtr, slEndQtr
    llTYDates(1) = gDateValue(slStartQtr)
    llTYDates(2) = gDateValue(slEndQtr)
    'gGetStartEndYear ilCorpStd, ilYear, slStartYr, slEndYr
    gGetStartEndYear ilCorpStd, ilYear - 1, slDate, slStr
    slStartYr = slDate
    gGetStartEndYear ilCorpStd, ilYear, slDate, slStr
    slEndYr = slStr

    llTYGetFrom = gDateValue(slStartYr)

    'Determine the Budget name selected
    slNameCode = tgRptSelBudgetCodeSP(igBSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 1, "\", slStr)
    ilRet = gParseItem(slStr, 1, "\", slYear)
    slYear = gSubStr("9999", slYear)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilBdMnfCode = Val(slCode)
    ilBdYear = Val(slYear)
    'ReDim tlSlsList(1 To 1) As SLSLIST          'array of vehicles and their sales
    ReDim tlSlsList(0 To 0) As SLSLIST          'array of vehicles and their sales
    'gather all budget records by vehicle for the requested year, totaling by quarter
    If Not mReadBvfRec(hmBvf, ilBdMnfCode, ilBdYear, tmBvfVeh()) Then
        Exit Sub
    End If

    'use startwk & endwk to gather budgets
    gObtainWkNo ilBvfCalType, slStartQtr, ilStartWk, ilFirstWk          'determine the first week inx to accum (current year)
    gObtainWkNo ilBvfCalType, slEndQtr, ilEndWk, ilLastWk               'determine the last week inx to accum  (current year)
    ilFound = False
    For ilVehicle = 0 To RptSelSP!lbcSelection(2).ListCount - 1 Step 1
        slNameCode = tgCSVNameCode(ilVehicle).sKey 'RptSelSP!lbcCSVNameCode.List(ilVehicle)
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If RptSelSP!lbcSelection(2).Selected(ilVehicle) Then    'only build those vehicles selected
            ilVefList(UBound(ilVefList)) = Val(slCode)
            'ReDim Preserve ilVefList(1 To UBound(ilVefList) + 1) As Integer
            ReDim Preserve ilVefList(LBound(ilVefList) To UBound(ilVefList) + 1) As Integer
        End If
    Next ilVehicle
    For ilLoop = LBound(tmBvfVeh) To UBound(tmBvfVeh) - 1 Step 1
        For ilVehicle = LBound(ilVefList) To UBound(ilVefList) - 1 Step 1
            If ilVefList(ilVehicle) = tmBvfVeh(ilLoop).iVefCode Then   'only build those vehicles selected

                For ilSlsLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1
                    If tmBvfVeh(ilLoop).iVefCode = tlSlsList(ilSlsLoop).iVefCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilSlsLoop
                If Not ilFound Then
                    ilSlsLoop = UBound(tlSlsList)
                    tlSlsList(UBound(tlSlsList)).iVefCode = tmBvfVeh(ilLoop).iVefCode
                    ReDim Preserve tlSlsList(LBound(tlSlsList) To UBound(tlSlsList) + 1)
                End If
                'ilSlsLoop contains index to the correct vehicle
                For ilTemp = ilStartWk To ilEndWk Step 1
                    tlSlsList(ilSlsLoop).lPlan = tlSlsList(ilSlsLoop).lPlan + tmBvfVeh(ilLoop).lGross(ilTemp)
                Next ilTemp
                Exit For
            End If
        Next ilVehicle
    Next ilLoop

    tlTranType.iAdj = True              'look only for adjustments in the History & Rec files
    tlTranType.iInv = False
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = False
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iNTR = False         '9-17-02
   
    ilRet = gObtainPhfRvf(RptSelSP, slStartYr, slEndYr, tlTranType, tlRvf(), 0)

    For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1 Step 1
        tmRvf = tlRvf(llRvfLoop)
        gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slCode
        llDate = gDateValue(slCode)
        ilFound = False
        If llDate >= llTYDates(1) And llDate < llTYDates(2) Then
            If llDate <= llEnterTo Then
                ilFound = True
                gPDNToLong tmRvf.sGross, llProject(1)           'theres only 1qtr to gather
            End If
        End If
        If ilFound Then                             'dates pass, is this a selected vehicle?
            ilFound = False
            For ilVehicle = LBound(ilVefList) To UBound(ilVefList) - 1 Step 1
                If ilVefList(ilVehicle) = tmRvf.iBillVefCode Then
                    ilFound = True                  'found a selected vehicle, proceed with remaining tests
                    Exit For
                End If
            Next ilVehicle
        End If
        If ilFound Then
            'Read the contract
            tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
            tmChfSrchKey1.iCntRevNo = 32000
            tmChfSrchKey1.iPropVer = 32000
            ilRet = btrGetGreaterOrEqual(hmCHF, tgChfSP, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
            'Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo <> tmRvf.lCntrNo Or (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M"))
            Do While (ilRet = BTRV_ERR_NONE) And (tgChfSP.lCntrNo = tmRvf.lCntrNo) And (tgChfSP.sSchStatus <> "F" And tgChfSP.sSchStatus <> "M")
                ilRet = btrGetNext(hmCHF, tgChfSP, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
            If ((ilRet <> BTRV_ERR_NONE) Or (tgChfSP.lCntrNo <> tmRvf.lCntrNo)) Then  'phoney a header from the receivables record so it can be procesed
                For ilLoop = 0 To 9
                    tgChfSP.iSlfCode(ilLoop) = 0
                    tgChfSP.lComm(ilLoop) = 0
                Next ilLoop
                tgChfSP.iPctTrade = 0
                If tmRvf.sCashTrade = "T" Then
                    tgChfSP.iPctTrade = 100           'ignore trades   later
                End If
            End If
            'Accumulate the $ projected into the vehicles buckets
            If llProject(1) <> 0 Then                            'ignore building any data whose lines didnt have $
                llProject(1) = llProject(1) \ 100               'drop pennies
                ilFoundSls = False
                Do While Not ilFoundSls
                    For ilSlsLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1
                        If tlSlsList(ilSlsLoop).iVefCode = tmRvf.iBillVefCode Then
                            ilFoundSls = True                   'current qtrs actuals (dont test for holds in receivables)
                            tlSlsList(ilSlsLoop).lTYAct = tlSlsList(ilSlsLoop).lTYAct + llProject(1)
                            Exit For
                        End If
                    Next ilSlsLoop
                    If Not ilFoundSls Then              'there wasnt a budget for this vehicle to begin with,
                                                        'no entry has been created
                        tlSlsList(UBound(tlSlsList)).iVefCode = tmRvf.iBillVefCode
                        ReDim Preserve tlSlsList(LBound(tlSlsList) To UBound(tlSlsList) + 1)
                    End If
                Loop                    'loop until a vehicle budget has been found
            End If
        End If                          'ilfound
    Next llRvfLoop
    'gather all contracts whose entered date is equal or prior to the requested date (gather from beginning of std year to
    'input date
    slCntrTypes = gBuildCntTypes()
    slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)
    slEndYr = Format$(gDateValue(slEndYr) + 90, "m/d/yy")        'get an extra quarter to make sure all changes included
     'ilRet = gObtainCntrForDate(RptSelSP, slStartYr, slEndQtr, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
    ilRet = gObtainCntrForDate(RptSelSP, slStartYr, slEndYr, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
        'project the $
        llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
        'Retrieve the contract, schedule lines and flights
        llContrCode = gPaceCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llEnterTo, hmCHF, tmChf)
        ilFound = False
        If llContrCode > 0 Then
            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfSP, tgClfSP(), tgCffSP())
            gUnpackDateLong tgChfSP.iOHDDate(0), tgChfSP.iOHDDate(1), llDate
            If llDate <= llEnterTo Then       'entered date must be entered thru effectve date
                ilFound = True
                'determine if the contracts start & end dates fall within the requested period
                gUnpackDateLong tgChfSP.iEndDate(0), tgChfSP.iEndDate(1), llTemp2      'hdr end date converted to long
                gUnpackDateLong tgChfSP.iStartDate(0), tgChfSP.iStartDate(1), llTemp1    'hdr start date converted to long
                If llTemp2 < llTYDates(1) Or llTemp1 >= llTYDates(2) Then
                    ilFound = False
                End If
            End If
        End If



        'ilFound = False
        'gUnPackDateLong tgChfSP.iOHDDate(0), tgChfSP.iOHDDate(1), llDate
        'If llDate <= llEnterTo Then       'entered date must be entered thru effectve date
        '    ilFound = True
        'End If
        If ilFound And tgChfSP.iPctTrade <> 100 Then      'ignore contracts 100% trade, process all other
            For ilClf = LBound(tgClfSP) To UBound(tgClfSP) - 1 Step 1
                llProject(1) = 0                'init bkts to accum qtr $ for this line
                tmClf = tgClfSP(ilClf).ClfRec

                ilVefFound = False
                For ilVehicle = LBound(ilVefList) To UBound(ilVefList) - 1 Step 1
                    If ilVefList(ilVehicle) = tmClf.iVefCode Then
                        ilVefFound = True
                        Exit For
                    End If
                Next ilVehicle
                If ilVefFound Then
                    If tmClf.sType = "H" Or tmClf.sType = "S" Then
                        gBuildFlights ilClf, llTYDates(), 1, 2, llProject(), 1, tgClfSP(), tgCffSP()
                    End If
                    'If slAirOrder = "O" Then                'invoice all contracts as ordered
                    '    If tmClf.sType <> "H" Then          'ignore all hidden lines for ordered billing, should be Pkg or conventional lines
                    '        gBuildFlights ilClf, llTYDates(), 1, 2, llProject(), 1
                    '    End If
                    'Else                                    'inv all contracts as aired
                    '    If tmClf.sType = "H" Then             'but if from pkg and hidden line, ignore hidd
                    '        'if hidden, will project if assoc. package is set to invoice as aired (real)
                    '        For ilTemp = LBound(tgClfSP) To UBound(tgClfSP) - 1    'find the assoc. pkg line for these hidden
                    '        If tmClf.iPkLineNo = tgClfSP(ilTemp).ClfRec.iLine Then
                    '            If tgClfSP(ilTemp).ClfRec.sType = "A" Then        'does the pkg line reflect bill as aired?
                    '                gBuildFlights ilClf, llTYDates(), 1, 2, llProject(), 1 'pkg bills as aired, project the hidden line
                    '            End If
                    '            Exit For
                    '        End If
                    '        Next ilTemp
                    '    Else                            'conventional, VV, or Pkg line
                    '        If tmClf.sType <> "A" Then  'if this package line to be invoiced aired (real times),
                    '                    'it has already been projected above with the hidden line
                    '            gBuildFlights ilClf, llTYDates(), 1, 2, llProject(), 1
                    '        End If
                    '    End If
                    'End If
                    'Accumulate the $ projected into the vehicles buckets
                    If llProject(1) > 0 Then
                        'do not truncate to process same as B & B.  When clients have a lot of pennies in their rates, theres more of a discrepancy with the comparison of reports
                        'If changing to round like B & B, ALL pacing needs to be corrected to round and drop pennies so they all balance
                        'llProject(1) = (llProject(1) + 50) \ 100       'round and drop pennies
                        llProject(1) = llProject(1) \ 100               'drop pennies
                        ilFoundSls = False
                        Do While Not ilFoundSls
                            For ilSlsLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1
                                If tlSlsList(ilSlsLoop).iVefCode = tmClf.iVefCode Then
                                    ilFoundSls = True
                                    If tgChfSP.sStatus = "H" Or tgChfSP.sStatus = "G" Then    'hold or unsch hold
                                        tlSlsList(ilSlsLoop).lTYActHold = tlSlsList(ilSlsLoop).lTYActHold + llProject(1)
                                        Exit For
                                    Else                            'order or unsch order
                                        tlSlsList(ilSlsLoop).lTYAct = tlSlsList(ilSlsLoop).lTYAct + llProject(1)
                                        Exit For
                                    End If
                                End If
                            Next ilSlsLoop
                            If Not ilFoundSls Then              'there wasnt a budget for this vehicle to begin with,
                                            'no entry has been created
                                tlSlsList(UBound(tlSlsList)).iVefCode = tmClf.iVefCode
                                ReDim Preserve tlSlsList(LBound(tlSlsList) To UBound(tlSlsList) + 1)
                            End If
                        Loop                    'loop until a vehicle budget has been found
                    End If                  'llproject > 0
                End If                      'ilveffound
            Next ilClf                      'process nextline
        End If                              'ilfound - llAdjust falls within requested dates
    Next ilCurrentRecd
    Erase tlChfAdvtExt
    'process Avails for 1 quarter
    ReDim tmAvr(0 To 0) As AVR
    '6-30-00  only do 1 quarter
    'gCRQAvails hmChf, 1, "B", slStartQtr, tlCntTypes, RptSelSp!lbcSelection(2), RptSelSp!lbcSelection(1), tmAvr()
    
    'Build array of selected named avails for sports vehicle.
    'if not a sports vehicle, follow normal rules of DP
    ilTemp = 0
    For ilCurrentRecd = 0 To RptSelSP!lbcSelection(3).ListCount - 1
        If RptSelSP!lbcSelection(3).Selected(ilCurrentRecd) Then
            slNameCode = tgNamedAvail(ilCurrentRecd).sKey
            ilRet = gParseItem(slNameCode, 1, "\", slStr)
            ilRet = gParseItem(slStr, 3, "|", slStr)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilAnfCodes(ilTemp) = Val(slCode)
            ilTemp = ilTemp + 1
            ReDim Preserve ilAnfCodes(0 To ilTemp) As Integer
         End If
    Next ilCurrentRecd
    
    gCRQAvails hmCHF, tlValuationInfo, slStartQtr, slEndQtr, tlCntTypes, RptSelSP!lbcSelection(2), RptSelSP!lbcSelection(1), tmAvr(), ilCorpStd, tmSpotLenRatio, ilAnfCodes()
    'loop thru tmAvr to build results for each vehicle
    For ilSlsLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1
        For ilTemp = 0 To UBound(tmAvr) - 1 Step 1
            If tlSlsList(ilSlsLoop).iVefCode = tmAvr(ilTemp).iVefCode Then
                'accumulate the quarters inventory valuation
                'Place the inventory valuation in "lProj" field within tlslsList
                'tlSlsList(ilSlsLoop).lProj = tlSlsList(ilSlsLoop).lProj + tmAvr(ilTemp).lMonth(1) + tmAvr(ilTemp).lMonth(2) + tmAvr(ilTemp).lMonth(3)
                tlSlsList(ilSlsLoop).lProj = tlSlsList(ilSlsLoop).lProj + tmAvr(ilTemp).lMonth(0) + tmAvr(ilTemp).lMonth(1) + tmAvr(ilTemp).lMonth(2)
            End If
        Next ilTemp
    Next ilSlsLoop
    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
    'tmGrf.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    tmGrf.iStartDate(0) = ilEnterDate(0)     'effective date entered
    tmGrf.iStartDate(1) = ilEnterDate(1)
    tmGrf.iCode2 = ilBdMnfCode                          'budget name
    'tmGrf.iPerGenl(3) = ilCorpStd             '1 = corp, 2 = std
    tmGrf.iPerGenl(2) = ilCorpStd             '1 = corp, 2 = std
    For ilLoop = LBound(tlSlsList) To UBound(tlSlsList) - 1 Step 1         'write a record per vehicle as longas something is non zero
        '4-7-04 change the testing to create record from summing to <> 0
        'If tlSlsList(ilLoop).lPlan + tlSlsList(ilLoop).lTYAct + tlSlsList(ilLoop).lProj + tlSlsList(ilLoop).lTYActHold <> 0 Then
         If tlSlsList(ilLoop).lPlan <> 0 Or tlSlsList(ilLoop).lTYAct <> 0 Or tlSlsList(ilLoop).lProj <> 0 Or tlSlsList(ilLoop).lTYActHold <> 0 Then
            tmGrf.iVefCode = tlSlsList(ilLoop).iVefCode
            gGetVehGrpSets tmGrf.iVefCode, ilMinorSet, ilMajorSet, ilmnfMinorCode, ilMnfMajorCode
            'tmGrf.iPerGenl(1) = ilmnfMinorCode
            'tmGrf.iPerGenl(2) = ilMnfMajorCode
            tmGrf.iPerGenl(0) = ilmnfMinorCode
            tmGrf.iPerGenl(1) = ilMnfMajorCode
            'tmGrf.lDollars(1) = tlSlsList(ilLoop).lPlan         'current year, plan $
            'tmGrf.lDollars(2) = tlSlsList(ilLoop).lTYAct        'current year, orders
            'tmGrf.lDollars(3) = tlSlsList(ilLoop).lTYActHold    'current year, holds
            'tmGrf.lDollars(4) = (CSng(tlSlsList(ilLoop).lProj) * tlValuationInfo.iEstPctSellout) / 100
            tmGrf.lDollars(0) = tlSlsList(ilLoop).lPlan         'current year, plan $
            tmGrf.lDollars(1) = tlSlsList(ilLoop).lTYAct        'current year, orders
            tmGrf.lDollars(2) = tlSlsList(ilLoop).lTYActHold    'current year, holds
            tmGrf.lDollars(3) = (CSng(tlSlsList(ilLoop).lProj) * tlValuationInfo.iEstPctSellout) / 100
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
    Next ilLoop
    Erase tlSlsList, tlMMnf, tmAvr, ilVefList, tlRvf
    Erase tlChfAdvtExt, tmBvfVeh
    sgCntrForDateStamp = ""
    ilRet = btrClose(hmBvf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmMnf)
End Sub
