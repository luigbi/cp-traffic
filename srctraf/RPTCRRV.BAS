Attribute VB_Name = "rptcrrv"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RPTCRRV.BAS on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
'Public igYear As Integer                'budget year used for filtering
'Public igMonthOrQtr As Integer          'entered month or qtr
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1    'CHF record image
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
Dim hmSbf As Integer
Dim hmMnf As Integer
Dim tmMnf() As MNF
Dim imMnfRecLen As Integer
Const NOT_SELECTED = 0
'  Receivables File
Dim tmRvf As RVF            'RVF record image
Dim tmRevSetCodes() As Integer 'array of revenue set codes selected
'********************************************************************************************
'
'                   gCrRevSets - Prepass for Revenue Sets report
'                   by Quarter for 1 year
'
'                   User selectivity:  Effective Date
'                                      Start Yr & Qtr
'                                      Corp or Std Month
'                                      One or more revenue sets
'                   This is a pacing report where all contracts are gather
'                   if the contract entered date is equal/prior to the
'                   sunday of the Effective date entred, and whose start/end
'                   dates span the quarter(s) requested.
'                   Records are written to GRF by contract, for 4 quarters
'
'                   Created:  4/30/02 D. Hosaka
'       2-11-05 Chg array of PHF/RVF array from integer to long (prevent overflow)
'********************************************************************************************
Sub gCrRevSets(tlRevSetCode() As SORTCODE)
Dim ilRet As Integer
Dim ilClf As Integer                    'loop for schedule lines
Dim ilHOState As Integer                'retrieve only latest order or revision
Dim slCntrTypes As String               'retrieve remnants, PI, DR, etc
Dim slCntrStatus As String              'retrieve H, O G or N (holds, orders, unsch hlds & orders)
Dim ilCurrentRecd As Integer            'loop for processing last years contracts
Dim llContrCode As Long                 'Internal Contr code to build all lines & flights of a contract
Dim ilFoundOne As Integer               'Found a matching  office built into mem
Dim ilTemp As Integer
Dim ilLoop As Integer                   'temp loop variable
Dim slTemp As String                    'temp string for dates
Dim ilCalType As Integer                '1 = std, 2 = corp calendar
'ReDim llProject(1 To 4) As Long        '$ projected for 4 quarters
ReDim llProject(0 To 4) As Long        '$ projected for 4 quarters. Index zero ignored
Dim llDate As Long                      'temp date variable
Dim llDate2 As Long
Dim slNameCode As String
'Date used to gather information
'String formats for generalized date conversions routines
'Long formats for testing
'Packed formats to store in GRF record
Dim ilStartQtr As Integer             'start qtr to gather data (1-4)
Dim ilTYStartYr As Integer              'year of this years start date     (1997-1998)
Dim slTYStart As String                 'start date of this year to begin gathering  (string)
Dim llTYStart As Long                   'start date of this year to begin gathering (Long)
Dim slTYEnd As String
Dim slWeekTYStart As String              'start date of week for this years new business entered this week
Dim llWeekTYStart As Long                'start date of week for this years new business entered on te user entered week
ReDim ilWeekTYStart(0 To 1) As Integer     'packed format for GRF record
Dim llEntryDate As Long                 'date entered from cntr header
'Month Starts to gather projection $ from flights
'ReDim llTYStartDates(1 To 5) As Long        'this year corp or std start dates for next 5 quarters
ReDim llTYStartDates(0 To 5) As Long        'this year corp or std start dates for next 5 quarters. Index zero ignored
'ReDim llTempStarts(1 To 13) As Long         'temp array for start dates for 13 months
ReDim llTempStarts(0 To 13) As Long         'temp array for start dates for 13 months. Index zero ignored
'   end of date variables
Dim tlTranType As TRANTYPES
'ReDim tlRvf(1 To 1) As RVF
ReDim tlRvf(0 To 0) As RVF
Dim slRevDesc As String
Dim llRvfLoop As Long                   '2-11-05
Dim blIncludeNTR As Boolean             'Dan M Ntr variables
Dim blIncludeHardCost As Boolean
Dim blNTRWithTotal As Boolean
Dim tlNTRInfo() As NTRPacing
Dim ilLowerboundNTR As Integer
Dim ilUpperboundNTR As Integer
Dim ilNTRCounter As Integer
Dim llSingleContract As Long
Dim llDateEntered As Long               'receivables entered date for pacing test
Dim blFailedMatchNtrOrHardCost As Boolean
Dim blFailedBecauseInstallment As Boolean
Dim blExceedFlag As Boolean
Dim slGrossNet As String                'Date:10.28.2019    implement gross/net options

    If Val(RptSelRV!edcSelC3.Text) > 0 And RptSelRV!edcSelC3.Text <> " " Then
         llSingleContract = Val(RptSelRV!edcSelC3.Text)
    Else
        llSingleContract = NOT_SELECTED
    End If
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
         ' Dan M 6-30-08
    hmSbf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmUrf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    hmMnf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmMnf, "", sgDBPath & "mnf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        btrDestroy hmSbf
        btrDestroy hmCHF
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmUrf
        btrDestroy hmGrf
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    ReDim tmMnf(0 To 0) As MNF
    imMnfRecLen = Len(tmMnf(0))
    tlTranType.iAdj = True              'look only for adjustments in the History & Rec files
    tlTranType.iInv = False
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = False
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iNTR = False         '9-17-02
    blExceedFlag = False
    If RptSelRV!ckcSelCInclude(0).Value Or RptSelRV!ckcSelCInclude(1).Value Then    'don't waste time filling array if don't need.
        tlTranType.iNTR = True      '8-13-08 Dan ntr/hard cost adjustments
        'chose ntr, hard, both?
        If RptSelRV!ckcSelCInclude(0).Value = 1 Then
             blIncludeNTR = True
        End If
        If RptSelRV!ckcSelCInclude(1).Value = 1 Then
            blIncludeHardCost = True
        End If

        ilRet = gObtainMnfForType("I", "", tmMnf())
        If ilRet <> True Then
            MsgBox "error retrieving MNF files", vbOKOnly + vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    slGrossNet = "G"                                'Date: 10/28/2019 Implement gross net option; default to gross
    If RptSelRV!rbcSelGrossNet(1).Value Then
        slGrossNet = "N"
    End If
    If Not gSetFormula("GrossNet", "'" & slGrossNet & "'") Then
        MsgBox "Error in RevSetsSelected Formula"
        Exit Sub
    End If

    ReDim tgClfRV(0 To 0) As CLFLIST
    tgClfRV(0).iStatus = -1 'Not Used
    tgClfRV(0).lRecPos = 0
    tgClfRV(0).iFirstCff = -1
    ReDim tgCffRV(0 To 0) As CFFLIST
    tgCffRV(0).iStatus = -1 'Not Used
    tgCffRV(0).lRecPos = 0
    tgCffRV(0).iNextCff = -1
    slRevDesc = ""
    ReDim tmRevSetCodes(0 To 0) As Integer
    For ilLoop = 0 To RptSelRV!lbcSelection.ListCount - 1 Step 1
        If RptSelRV!lbcSelection.Selected(ilLoop) Then
            slTemp = tlRevSetCode(ilLoop).sKey    'lbcMster.List(ilLoop)
            ilRet = gParseItem(slTemp, 2, "\", slNameCode)
            tmRevSetCodes(UBound(tmRevSetCodes)) = Val(slNameCode)
            ilRet = gParseItem(slTemp, 1, "\", slNameCode)
            ilRet = gParseItem(slNameCode, 2, "|", slNameCode)
            If RptSelRV!ckcAll.Value = vbChecked Then
                slRevDesc = "All Revenue Sets"
            Else
                'filter out special characters from name
                slNameCode = gFixQuote(slNameCode)
                If slRevDesc = "" Then
                    slRevDesc = Trim$(slNameCode)
                Else
                    If Len(slRevDesc) + Len(Trim$(slNameCode)) >= 245 Then
                        If Not blExceedFlag Then
                            slRevDesc = slRevDesc & "..."
                            blExceedFlag = True
                        End If
                    Else
                        slRevDesc = slRevDesc & ", " & Trim$(slNameCode)
                    End If
                End If
            End If
            ReDim Preserve tmRevSetCodes(0 To UBound(tmRevSetCodes) + 1) As Integer
        End If
    Next ilLoop
    If Not gSetFormula("RevSetsSelected", "'" & slRevDesc & "'") Then
        MsgBox "Error in RevSetsSelected Formula"
        Exit Sub
    End If
    'Get STart and end dates of current week for for pacing test
    slWeekTYStart = RptSelRV!CSI_CalFrom.Text                       'Date:10/28/2019 using CSI calendar for date entry -->   edcSelCFrom.Text
    llWeekTYStart = gDateValue(slWeekTYStart)
    gPackDate slWeekTYStart, ilWeekTYStart(0), ilWeekTYStart(1)    'conversion to store in prepass record
    'setup year from user input
    ilTYStartYr = Val(RptSelRV!edcSelCTo.Text)
    ilStartQtr = Val(RptSelRV!edcSelCTo1.Text)

    If RptSelRV!rbcSelCSelect(0).Value Then      'corp month or qtr? (vs std)
        ilCalType = 2                   'corp flag to store in grf
        ilLoop = gGetCorpCalIndex(ilTYStartYr)
        gGetStartEndQtr 1, ilTYStartYr, ilStartQtr, slTYStart, slTYEnd
        llTYStart = gDateValue(slTYStart)
    Else
        ilCalType = 1                   'std flag to store in grf
        gGetStartEndQtr 2, ilTYStartYr, ilStartQtr, slTYStart, slTYEnd
        llTYStart = gDateValue(slTYStart)
    End If
    'Determine startdates for this year for 13 months
    gBuildStartDates slTYStart, ilCalType, 13, llTempStarts()
    'Got the 13 monthly start dates, convert to quarter dates
    For ilLoop = 1 To 13 Step 3
        llTYStartDates((ilLoop \ 3) + 1) = llTempStarts(ilLoop)
    Next ilLoop
    slTYEnd = Format$(llTYStartDates(5) - 1, "m/d/yy")      'last date reqd so the contr gathering has an earliest & latest date
                                                            'tosearch & gather


    ilRet = gObtainPhfRvf(RptSelRV, slTYStart, slTYEnd, tlTranType, tlRvf(), 0)
    For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1 Step 1
    tmRvf = tlRvf(llRvfLoop)
    'dan M 7-13-08 added single contract selectivity
        If llSingleContract = NOT_SELECTED Or llSingleContract = tmRvf.lCntrNo Then
            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slTemp
            llDate = gDateValue(slTemp)
            ilFoundOne = False

            gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slTemp
            llDateEntered = gDateValue(slTemp)

            'Dan M 8-13-8 ntr/hard cost adjustments.  Is this record ntr/hard cost and do we want that?
            blFailedMatchNtrOrHardCost = False
            'Dan M 8-13-8 don't allow installment option "I"
            blFailedBecauseInstallment = False
            If tmRvf.sType = "I" Then
                blFailedBecauseInstallment = True
            End If
            If ((blIncludeNTR) Xor (blIncludeHardCost)) And tmRvf.iMnfItem > 0 Then      'one or the other is true, but not both (if both true, don't have to isolate anything)
                ilRet = gIsItHardCost(tmRvf.iMnfItem, tmMnf())
            'if is hard cost but blincludentr  or isn't hard cost but blincludehardcost then it needs to be removed. set failedmatchntrorhardcost true
                If (ilRet And blIncludeNTR) Or ((Not ilRet) And blIncludeHardCost) Then
                    blFailedMatchNtrOrHardCost = True
                End If
            End If
            If Not (blFailedMatchNtrOrHardCost Or blFailedBecauseInstallment) Then  'if both false, continue
                If llDate >= llTYStartDates(1) And llDate < llTYStartDates(5) Then
                   ' If llDate <= llWeekTYStart Then        dan M replaced with below. 8-13-08
                    If llDateEntered <= llWeekTYStart Then
                        For ilTemp = 1 To 4                           'setup general buffer to use This years dates Dan M changed from 5 to 4 6-30-08
                            'llGenlDates(ilTemp) = llTYStarts(ilTemp)
                            If llDate >= llTYStartDates(ilTemp) And llDate < llTYStartDates(ilTemp + 1) Then
                                ilFoundOne = True
                                gPDNToLong tmRvf.sGross, llProject(ilTemp)
                                Exit For
                            End If
                        Next ilTemp
                    End If
                End If
            End If
            If ilFoundOne Then
                'Read the contract
                tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tgChfRV, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
                'Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo <> tmRvf.lCntrNo Or (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M"))
                Do While (ilRet = BTRV_ERR_NONE) And (tgChfRV.lCntrNo = tmRvf.lCntrNo) And (tgChfRV.sSchStatus <> "F" And tgChfRV.sSchStatus <> "M")
                ilRet = btrGetNext(hmCHF, tgChfRV, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If ((ilRet <> BTRV_ERR_NONE) Or (tgChfRV.lCntrNo <> tmRvf.lCntrNo)) Then  'phoney a header from the receivables record so it can be procesed
                    For ilLoop = 0 To 9
                        tgChfRV.iSlfCode(ilLoop) = 0
                        tgChfRV.lComm(ilLoop) = 0
                        tgChfRV.iMnfSubCmpy(ilLoop) = 0       '4-15-00
                    Next ilLoop
                    tgChfRV.iAdfCode = tmRvf.iAdfCode
                    tgChfRV.iSlfCode(0) = tmRvf.iSlfCode
                    tgChfRV.lComm(0) = 1000000
                    tgChfRV.iPctTrade = 0
                    If tmRvf.sCashTrade = "T" Then
                        tgChfRV.iPctTrade = 100           'ignore trades   later
                    End If
                    'Dan M added 8-20-08
                    tgChfRV.sProduct = ""
                    tgChfRV.lCntrNo = tmRvf.lCntrNo
                End If
            End If                          'ilfoundOne
            If ilFoundOne And tgChfRV.iPctTrade <> 100 Then               'valid contract & selling office, and not trade
                mUpdate ilWeekTYStart(), ilCalType, ilTYStartYr, llProject(), tmRvf.iAirVefCode, slGrossNet
            End If
        End If
    Next llRvfLoop


    'Gather all contracts for previous year and current year whose effective date entered
    'is prior to the effective date that affects either previous year or current year
    slCntrTypes = gBuildCntTypes()
    slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)


    'Obtain contracts for the previous year so that when modifications are done to schedules
    'that cause the current vrsions hdr dates to be outside of requested period, the proper
    'revisions are processed.
    If RptSelRV!rbcSelCSelect(0).Value Then      'corp month or qtr? (vs std)
        gGetStartEndYear 1, ilTYStartYr - 1, slTYStart, slTemp
    Else
        gGetStartEndYear 2, ilTYStartYr, slTYStart, slTemp
    End If
    For ilLoop = 1 To 4         '4-15-00, insure the last transaction initialized for the contract processing
        llProject(ilLoop) = 0
    Next ilLoop
    ilRet = gObtainCntrForDate(RptSelRV, slTYStart, slTYEnd, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())

    'All contracts have been retrieved for all of this year
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
        '7-13-08 added single contract selectivity dan M
        If llSingleContract = NOT_SELECTED Or llSingleContract = tlChfAdvtExt(ilCurrentRecd).lCntrNo Then
            ilFoundOne = True
            llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
            'Retrieve the contract, schedule lines and flights
            llContrCode = gPaceCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llWeekTYStart, hmCHF, tmChf)
            If llContrCode > 0 Then
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfRV, tgClfRV(), tgCffRV())

                'determine if the contracts start & end dates fall within the requested period
                gUnpackDateLong tgChfRV.iEndDate(0), tgChfRV.iEndDate(1), llDate2      'hdr end date converted to long
                gUnpackDateLong tgChfRV.iStartDate(0), tgChfRV.iStartDate(1), llDate    'hdr start date converted to long
                If llDate2 < llTYStartDates(1) Or llDate >= llTYStartDates(5) Then
                    ilFoundOne = False
                End If
            Else                            'no entered date of contract falls within effective date
                ilFoundOne = False
            End If

            If ilFoundOne And tgChfRV.iPctTrade <> 100 Then    'ignore 100% trade contracts
                'the date entered must be equal or prior to user entred effective date
                gUnpackDate tgChfRV.iOHDDate(0), tgChfRV.iOHDDate(1), slTemp            'convert date entered
                llEntryDate = gDateValue(slTemp)

                'get cnts earliest and latest dates to see if it spans the requested period
                gUnpackDate tgChfRV.iStartDate(0), tgChfRV.iStartDate(1), slTemp       '
                llDate = gDateValue(slTemp)
                gUnpackDate tgChfRV.iEndDate(0), tgChfRV.iEndDate(1), slTemp
                llDate2 = gDateValue(slTemp)

                'Process all contracts up thru the user entered date
                If llEntryDate <= llWeekTYStart Then                'within the pacing period?
                    For ilClf = LBound(tgClfRV) To UBound(tgClfRV) - 1 Step 1
                        tmClf = tgClfRV(ilClf).ClfRec
                        'Project the monthly $ from the flights
                        If tmClf.sType = "S" Or tmClf.sType = "H" Then
                            gBuildFlights ilClf, llTYStartDates(), 1, 5, llProject(), 1, tgClfRV(), tgCffRV()
                            '4-15-00
                            mUpdate ilWeekTYStart(), ilCalType, ilTYStartYr, llProject(), tmClf.iVefCode, slGrossNet
                        End If
                    Next ilClf                                      'loop thru schedule lines
                End If                                          'cnt entered date <= user entered date
            End If
                ' 6-30-08 Dan M added NTR/hard cost to report
            'Does user want to see HardCost/NTR?  have hard cost? Not pure trade?
            If (blIncludeNTR Or blIncludeHardCost) And (tgChfRV.sNTRDefined = "Y") And (tlChfAdvtExt(ilCurrentRecd).iPctTrade <> 100) And ilFoundOne Then
            'call routine to fill array with choice
                gNtrByContract llContrCode, llDate, llDate2, tlNTRInfo(), tmMnf(), hmSbf, blIncludeNTR, blIncludeHardCost, RptSelRV
                ilLowerboundNTR = LBound(tlNTRInfo)
                ilUpperboundNTR = UBound(tlNTRInfo)
                If ilUpperboundNTR <> ilLowerboundNTR Then
    '            'clear array
                    For ilLoop = 1 To 4
                        llProject(ilLoop) = 0
                    Next ilLoop
                    For ilNTRCounter = ilLowerboundNTR To ilUpperboundNTR - 1 Step 1
                    'flag to see that contract has a value for writing
                        blNTRWithTotal = False
                        For ilTemp = 1 To 4
                             'look at each ntr record's date to see if falls into specific time period.
                            If tlNTRInfo(ilNTRCounter).lSbfDate >= llTYStartDates(ilTemp) And tlNTRInfo(ilNTRCounter).lSbfDate < llTYStartDates(ilTemp + 1) Then
                                'flag so won't write record if all values are 0
                                If tlNTRInfo(ilNTRCounter).lSBFTotal > 0 Then
                                    blNTRWithTotal = True
                                    llProject(ilTemp) = llProject(ilTemp) + tlNTRInfo(ilNTRCounter).lSBFTotal
                                End If
                            End If
                            'send to routine to write to grf
                            If blNTRWithTotal = True Then
                               mUpdate ilWeekTYStart(), ilCalType, ilTYStartYr, llProject(), tlNTRInfo(ilNTRCounter).iVefCode, slGrossNet
                            End If
                        Next ilTemp
                    Next ilNTRCounter
                End If
            End If
        End If  'contract selectivity
    Next ilCurrentRecd                                      'loop for CHF records
    Erase tlChfAdvtExt, tlRvf, tgClfRV, tgCffRV, tlNTRInfo, tmMnf
    Erase llTYStartDates, llTempStarts, llProject
    sgCntrForDateStamp = ""
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmUrf)
    ilRet = btrClose(hmGrf)
End Sub
'
'
'               mUpFScatSplits - Create all the Splits for split offices into
'                   New/Return and Upfront/Scatter revenue sets
'
'              <input> ilWeekTYStart() -  btrieve format: effective date
'                       ilCalType - Corp(2) or Std code (1)
'                       ilTYStartYr - Start of Year requested (1997, 1998, 2000, etc)
'                       llProject() - array of $ generated from contract lines or receivables trans.
'                                       (4 entries, representing each qtr for the year)
'               <output> None
'
'               mUpDate ilWeekTYStart(), ilCalType, ilTYStartYr, llProject(), slGrossNet
'
'               Created:  4/20/98 to include adjustments from PHF/RVF for ABC
'                       4-15-00 Implement changes due to new commission structure
Sub mUpdate(ilWeekTYStart() As Integer, ilCalType As Integer, ilTYStartYr As Integer, llProject() As Long, ilVehicle As Integer, slGrossNet As String)
Dim ilLoop As Integer
Dim llCalcGross As Long
Dim ilRet As Integer
Dim ilInclude As Integer
Dim ilMnf As Integer
Dim llRate As Long

    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)
    'tmGrf.iGenTime(1) = igNowTime(1)
    tmGrf.lGenTime = lgNowTime          '5-1-02
    'tmGrf.lChfCode = tgChfRV.lCode
    tmGrf.lChfCode = tgChfRV.lCntrNo        '8-20-08 replaced above for rvf that don't have contracts Dan M
    tmGrf.iAdfCode = tgChfRV.iAdfCode
    tmGrf.sGenDesc = tgChfRV.sProduct       '8-20-08 added Dan M
    ilInclude = False

    For ilLoop = 0 To 4
        For ilMnf = LBound(tmRevSetCodes) To UBound(tmRevSetCodes) - 1
            If tgChfRV.iMnfRevSet(ilLoop) = tmRevSetCodes(ilMnf) Then
                'tmGrf.iPerGenl(ilLoop + 1) = tgChfRV.iMnfRevSet(ilLoop)
                tmGrf.iPerGenl(ilLoop) = tgChfRV.iMnfRevSet(ilLoop)
                ilInclude = True
                Exit For
            Else
                'tmGrf.iPerGenl(ilLoop + 1) = 0
                tmGrf.iPerGenl(ilLoop) = 0
            End If
        Next ilMnf
    Next ilLoop
    If ilInclude Then
        tmGrf.iCode2 = ilCalType
        'tmGrf.iDateGenl(0, 1) = ilWeekTYStart(0)  'Start date of week for this year
        'tmGrf.iDateGenl(1, 1) = ilWeekTYStart(1)
        tmGrf.iDateGenl(0, 0) = ilWeekTYStart(0)  'Start date of week for this year
        tmGrf.iDateGenl(1, 0) = ilWeekTYStart(1)
        tmGrf.iYear = ilTYStartYr           'year (1997, 1998, etc)  for header
        'Format  quarterly totals  into output buffer

        For ilLoop = 1 To 4
            llRate = 0
            llProject(ilLoop) = llProject(ilLoop) '\ 100 'store the penniless value to avoid redoing it later for splits
            'Date: 10/29/2019  get gross or net amount
            llRate = gGetGrossOrNetFromRate(llProject(ilLoop), slGrossNet, tgChfRV.iAgfCode)
            llCalcGross = llCalcGross + llProject(ilLoop)
            tmGrf.lDollars(ilLoop - 1) = llRate 'llProject(ilLoop)
        Next ilLoop
        If llCalcGross <> 0 Then
            'tmGrf.lDollars(5) = llCalcGross                        'year total
            tmGrf.lDollars(4) = llCalcGross                        'year total

            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)

        End If                                      'llcalcgross <> 0
    End If
    For ilLoop = 1 To 4
        llProject(ilLoop) = 0
    Next ilLoop
    llCalcGross = 0
End Sub
