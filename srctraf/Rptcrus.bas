Attribute VB_Name = "RPTCRUS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrus.bas on Wed 6/17/09 @ 12:56 PM
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
Dim hmSlf As Integer            'Salesperson file handle
Dim imSlfRecLen As Integer      'SLF record length
Dim tmSlf As SLF
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
Dim imIncludeCodes As Integer
Dim imUseCodes() As Integer
Const NOT_SELECTED = 0
'  Receivables File
Dim tmRvf As RVF            'RVF record image
Dim tlSofList() As SOFLIST    'list of selling office codes and sales sourcecodes
'********************************************************************************************
'
'                   gCrUpfScat - Prepass for Upfront/Scatter report
'                   by Quarter for 1 year
'
'                   User selectivity:  Effective Date (backed up to Monday)
'                                      Start Yr & Qtr
'                                      Corp or Std Month
'                   This is a pacing report where all contracts are gather
'                   if the contract entered date is equal/prior to the
'                   sunday of the Effective date entred, and whose start/end
'                   dates span the quarter(s) requested.
'                   Records are written to GRF by contract.
'
'                   Created:  11/3/97 D. Hosaka
'                   4/10/98 Ignore 100% trade contracts
'                   4-15-00 Implement changes due to new commission structure
'       2-11-05 Chg array of PHF/RVF array from integer to long (prevent overflow)
'       12-14-06 add parm to gObtainRvfPhf to test on tran date (vs entry date)
'********************************************************************************************
Sub gCrUpfScat()
Dim ilRet As Integer                    '
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
Dim llRvfLoop As Long                       '2-11-05
Dim blIncludeNTR As Boolean
Dim blIncludeHardCost As Boolean
Dim blNTRWithTotal As Boolean
Dim tlNTRInfo() As NTRPacing
Dim ilLowerboundNTR As Integer
Dim ilUpperboundNTR As Integer
Dim ilNTRCounter As Integer
Dim llSingleContract As Long            '7-11-08 test for option to get single contract
Dim llDateEntered As Long               'receivables entered date for pacing test
Dim blFailedMatchNtrOrHardCost As Boolean   'flag if record doesn't match what user selected
Dim blFailedBecauseInstallment As Boolean   'flag to stop invoice adjustments when installments
If Val(RptSelUS!edcSelC4.Text) > 0 And RptSelUS!edcSelC4.Text <> " " Then
     llSingleContract = Val(RptSelUS!edcSelC4.Text)
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
     ' Dan M 6-23-08
    hmSbf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf
        btrDestroy hmSlf
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
        btrDestroy hmSlf
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
    If RptSelUS!ckcSelCInclude(0).Value Or RptSelUS!ckcSelCInclude(1).Value Then    'don't waste time filling array if don't need.
    'set flags ntr or hard cost or both chosen
        If RptSelUS!ckcSelCInclude(0).Value = 1 Then
             blIncludeNTR = True
             tlTranType.iNTR = True
        End If
        If RptSelUS!ckcSelCInclude(1).Value = 1 Then
            blIncludeHardCost = True
            tlTranType.iNTR = True
        End If
        ilRet = gObtainMnfForType("I", "", tmMnf())
        If ilRet <> True Then
            MsgBox "error retrieving MNF files", vbOKOnly + vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    ReDim tgClfUS(0 To 0) As CLFLIST
    tgClfUS(0).iStatus = -1 'Not Used
    tgClfUS(0).lRecPos = 0
    tgClfUS(0).iFirstCff = -1
    ReDim tgCffUS(0 To 0) As CFFLIST
    tgCffUS(0).iStatus = -1 'Not Used
    tgCffUS(0).lRecPos = 0
    tgCffUS(0).iNextCff = -1
    
    gObtainCodesForMultipleLists 1, tgSalesperson(), imIncludeCodes, imUseCodes(), RptSelUS
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
    'Get STart and end dates of current week for for pacing test
'    slWeekTYStart = RptSelUS!edcSelCFrom.Text
    slWeekTYStart = RptSelUS!CSI_CalFrom.Text           '12-13-19 change touse csi calendar control vs edit box
    llWeekTYStart = gDateValue(slWeekTYStart)
    gPackDate slWeekTYStart, ilWeekTYStart(0), ilWeekTYStart(1)    'conversion to store in prepass record
    'setup year from user input
    ilTYStartYr = Val(RptSelUS!edcSelCTo.Text)
    ilStartQtr = Val(RptSelUS!edcSelCTo1.Text)

    If RptSelUS!rbcSelCSelect(0).Value Then      'corp month or qtr? (vs std)
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

    ilRet = gObtainPhfRvf(RptSelUS, slTYStart, slTYEnd, tlTranType, tlRvf(), 0)
    For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1 Step 1
        tmRvf = tlRvf(llRvfLoop)
        'dan M 7-11-08 added single contract selectivity
        If llSingleContract = NOT_SELECTED Or llSingleContract = tmRvf.lCntrNo Then
            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slTemp
            llDate = gDateValue(slTemp)
            gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slTemp
            llDateEntered = gDateValue(slTemp)

            ilFoundOne = False
            'Dan M 8-11-8 ntr/hard cost adjustments.  Is this record ntr/hard cost and do we want that?
            blFailedMatchNtrOrHardCost = False
            'Dan M 8-12-8 don't allow installment option "I"
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
           ' If Not blFailedMatchNtrOrHardCost Then  'Dan added to remove ntr if only hardcost wanted and viceversa

                If llDate >= llTYStartDates(1) And llDate < llTYStartDates(5) Then
                    'If llDate <= llWeekTYStart Then
                    If llDateEntered <= llWeekTYStart Then          '8-7-08 test effec pacing with date entered; otherwise get too many records
                        For ilTemp = 1 To 4                       'dan M changed from 5 to 4.    'setup general buffer to use This years dates
                            'llGenlDates(ilTemp) = llTYStarts(ilTemp)
                            If llDate >= llTYStartDates(ilTemp) And llDate < llTYStartDates(ilTemp + 1) Then
                                ilFoundOne = True
                                gPDNToLong tmRvf.sGross, llProject(ilTemp)
                                Exit For
                            End If
                        Next ilTemp
                    End If
                End If
            End If 'failed match
            If ilFoundOne Then
                'Read the contract
                tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tgChfUS, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
                'Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo <> tmRvf.lCntrNo Or (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M"))
                Do While (ilRet = BTRV_ERR_NONE) And (tgChfUS.lCntrNo = tmRvf.lCntrNo) And (tgChfUS.sSchStatus <> "F" And tgChfUS.sSchStatus <> "M")
                    ilRet = btrGetNext(hmCHF, tgChfUS, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If ((ilRet <> BTRV_ERR_NONE) Or (tgChfUS.lCntrNo <> tmRvf.lCntrNo)) Then  'phoney a header from the receivables record so it can be procesed
                    For ilLoop = 0 To 9
                        tgChfUS.iSlfCode(ilLoop) = 0
                        tgChfUS.lComm(ilLoop) = 0
                        tgChfUS.iMnfSubCmpy(ilLoop) = 0       '4-15-00
                    Next ilLoop
                    tgChfUS.iAdfCode = tmRvf.iAdfCode
                    tgChfUS.iSlfCode(0) = tmRvf.iSlfCode
                    tgChfUS.lComm(0) = 1000000
                    tgChfUS.iPctTrade = 0
                    If tmRvf.sCashTrade = "T" Then
                        tgChfUS.iPctTrade = 100           'ignore trades   later
                    End If
                    'Dan M added 8-19-08
                    tgChfUS.lCntrNo = tmRvf.lCntrNo
                    tgChfUS.sProduct = ""

                End If
                '4-15-00  remove test for office here, tested in mUpFScatSplits
                'ilFoundOne = mTestSelectedOffice(tlSofList())            'has this office been selected?
            End If                          'ilfoundOne
            If ilFoundOne And tgChfUS.iPctTrade <> 100 Then               'valid contract & selling office, and not trade
                mUpFScatSplits ilWeekTYStart(), ilCalType, ilTYStartYr, tlSofList(), llProject(), tmRvf.iAirVefCode
            End If
        End If  'single contract selectivity
    Next llRvfLoop


    'Gather all contracts for previous year and current year whose effective date entered
    'is prior to the effective date that affects either previous year or current year
    slCntrTypes = gBuildCntTypes()
    slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)


    'Obtain contracts for the previous year so that when modifications are done to schedules
    'that cause the current vrsions hdr dates to be outside of requested period, the proper
    'revisions are processed.
    If RptSelUS!rbcSelCSelect(0).Value Then      'corp month or qtr? (vs std)
        gGetStartEndYear 1, ilTYStartYr - 1, slTYStart, slTemp
    Else
        gGetStartEndYear 2, ilTYStartYr, slTYStart, slTemp
    End If
    For ilLoop = 1 To 4         '4-15-00, insure the last transaction initialized for the contract processing
        llProject(ilLoop) = 0
    Next ilLoop
    ilRet = gObtainCntrForDate(RptSelUS, slTYStart, slTYEnd, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
    'All contracts have been retrieved for all of this year
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
        '7-11-08 added single contract selectivity dan M
        If llSingleContract = NOT_SELECTED Or llSingleContract = tlChfAdvtExt(ilCurrentRecd).lCntrNo Then
            ilFoundOne = True
            llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
            'Retrieve the contract, schedule lines and flights
            llContrCode = gPaceCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llWeekTYStart, hmCHF, tmChf)
            If llContrCode > 0 Then
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfUS, tgClfUS(), tgCffUS())

                '4-11-00 remove test, tested later in mUpFScatSplits - ilFoundOne = mTestSelectedOffice(tlSofList())
                'ilFoundOne = mTestSelectedOffice(tlSofList())
                'determine if the contracts start & end dates fall within the requested period
                gUnpackDateLong tgChfUS.iEndDate(0), tgChfUS.iEndDate(1), llDate2      'hdr end date converted to long
                gUnpackDateLong tgChfUS.iStartDate(0), tgChfUS.iStartDate(1), llDate    'hdr start date converted to long
                If llDate2 < llTYStartDates(1) Or llDate >= llTYStartDates(5) Then
                    ilFoundOne = False
                End If
            Else                            'no entered date of contract falls within effective date
                ilFoundOne = False
            End If

            If ilFoundOne And tgChfUS.iPctTrade <> 100 Then    'ignore 100% trade contracts
                'the date entered must be equal or prior to user entred effective date
                gUnpackDate tgChfUS.iOHDDate(0), tgChfUS.iOHDDate(1), slTemp            'convert date entered
                llEntryDate = gDateValue(slTemp)

                'get cnts earliest and latest dates to see if it spans the requested period
                gUnpackDate tgChfUS.iStartDate(0), tgChfUS.iStartDate(1), slTemp       '
                llDate = gDateValue(slTemp)
                gUnpackDate tgChfUS.iEndDate(0), tgChfUS.iEndDate(1), slTemp
                llDate2 = gDateValue(slTemp)

                'Process all contracts up thru the user entered date
                If llEntryDate <= llWeekTYStart Then                'within the pacing period?
                    For ilClf = LBound(tgClfUS) To UBound(tgClfUS) - 1 Step 1
                        tmClf = tgClfUS(ilClf).ClfRec
                        'Project the monthly $ from the flights
                        If tmClf.sType = "S" Or tmClf.sType = "H" Then
                            gBuildFlights ilClf, llTYStartDates(), 1, 5, llProject(), 1, tgClfUS(), tgCffUS()
                            '4-15-00
                            mUpFScatSplits ilWeekTYStart(), ilCalType, ilTYStartYr, tlSofList(), llProject(), tmClf.iVefCode
                        End If
                    Next ilClf                                      'loop thru schedule lines
                    '4-15-00 mUpFScatSplits ilWeekTYStart(), ilCalType, ilTYStartYr, tlSofList(), llProject()
                End If                                          'cnt entered date <= user entered date
            End If

                    ' Dan M 6-27-08 Add NTR/Hard Cost option
            'Does user want to see HardCost/NTR?  Not pure trade?
         '   If (blIncludeNTR Or blIncludeHardCost) And (llContrCode > 0) And (tmChf.sNTRDefined = "Y") And (tmChf.iPctTrade <> 100) Then
             If (blIncludeNTR Or blIncludeHardCost) And (tlChfAdvtExt(ilCurrentRecd).iPctTrade <> 100) And (tgChfUS.sNTRDefined = "Y") And ilFoundOne Then
                'call routine to fill array with choice

                gNtrByContract llContrCode, llDate, llDate2, tlNTRInfo(), tmMnf(), hmSbf, blIncludeNTR, blIncludeHardCost, RptSelUS
                ilLowerboundNTR = LBound(tlNTRInfo)
                ilUpperboundNTR = UBound(tlNTRInfo)
             'ntr or hard cost found?
                If ilUpperboundNTR <> ilLowerboundNTR Then
                    'clear array
                    For ilLoop = 1 To 4
                        llProject(ilLoop) = 0
                    Next ilLoop
                    'flag to see that contract has a value for writing
                    For ilNTRCounter = ilLowerboundNTR To ilUpperboundNTR - 1 Step 1
                        blNTRWithTotal = False
                            For ilTemp = 1 To 4             'was 1 to 5
                                 'look at each ntr record's date to see if falls into specific time period.
                                If tlNTRInfo(ilNTRCounter).lSbfDate >= llTYStartDates(ilTemp) And tlNTRInfo(ilNTRCounter).lSbfDate < llTYStartDates(ilTemp + 1) Then
                                    'flag so won't write record if all values are 0
                                    If tlNTRInfo(ilNTRCounter).lSBFTotal > 0 Then
                                        blNTRWithTotal = True
                                        llProject(ilTemp) = llProject(ilTemp) + tlNTRInfo(ilNTRCounter).lSBFTotal
                                    End If
                                    'Exit For
                                End If
                                'send to routine to write to grf
                                If blNTRWithTotal = True Then
                                   mUpFScatSplits ilWeekTYStart(), ilCalType, ilTYStartYr, tlSofList(), llProject(), tlNTRInfo(ilNTRCounter).iVefCode
                                End If
                            Next ilTemp
                    Next ilNTRCounter
                End If
            End If
        End If      'single contract
    Next ilCurrentRecd                                      'loop for CHF records



    Erase tlChfAdvtExt, tlSofList, tlRvf, tgClfUS, tgCffUS, tlNTRInfo, tmMnf
    Erase llTYStartDates, llTempStarts, llProject
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
'               mUpFScatSplits - Create all the Splits for split offices into
'                   New/Return and Upfront/Scatter revenue sets
'
'              <input> ilWeekTYStart() -  btrieve format: effective date
'                       ilCalType - Corp(2) or Std code (1)
'                       ilTYStartYr - Start of Year requested (1997, 1998, 2000, etc)
'                       tlSofList() - array of valid sales offices to include
'                       llProject() - array of $ generated from contract lines or receivables trans.
'                                       (4 entries, representing each qtr for the year)
'               <output> None
'
'               mUpFScatSplits ilWeekTYStart(), ilCalType, ilTYStartYr, tlSofList(), llProject()
'
'               Created:  4/20/98 to include adjustments from PHF/RVF for ABC
'                       4-15-00 Implement changes due to new commission structure
Sub mUpFScatSplits(ilWeekTYStart() As Integer, ilCalType As Integer, ilTYStartYr As Integer, tlSofList() As SOFLIST, llProject() As Long, ilVehicle As Integer)
Dim ilLoop As Integer
Dim llCalcGross As Long
Dim ilFoundOne As Integer
Dim ilSlspLoop As Integer
Dim ilSaveSof As Integer
Dim ilTemp As Integer
Dim slNameCode As String
Dim slCode As String
Dim ilRet As Integer
Dim slTemp As String
Dim slAmount As String
Dim slSharePct As String
Dim ilQtrs As Integer
Dim ilPrimarySof As Integer     '4-15-00
Dim ilPrimarySubCo As Integer
'ReDim llTempProject(1 To 4) As Long   '4-15-00
ReDim llTempProject(0 To 4) As Long   '4-15-00. Index zero ignored
Dim ilLastSlsProc As Integer
Dim ilMnfSubCo As Integer
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)
    'tmGrf.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    'tmGrf.lChfCode = tgChfUS.lCode
    tmGrf.lChfCode = tgChfUS.lCntrNo        'Dan M replaced for rvf's that don't have chfcode 8-20-08
    tmGrf.iAdfCode = tgChfUS.iAdfCode
    For ilLoop = 0 To 4
        'tmGrf.iPerGenl(ilLoop + 1) = tgChfUS.iMnfRevSet(ilLoop)
        tmGrf.iPerGenl(ilLoop) = tgChfUS.iMnfRevSet(ilLoop)
    Next ilLoop
    tmGrf.iCode2 = ilCalType
    'tmGrf.iDateGenl(0, 1) = ilWeekTYStart(0)  'Start date of week for this year
    'tmGrf.iDateGenl(1, 1) = ilWeekTYStart(1)
    tmGrf.iDateGenl(0, 0) = ilWeekTYStart(0)  'Start date of week for this year
    tmGrf.iDateGenl(1, 0) = ilWeekTYStart(1)
    tmGrf.iYear = ilTYStartYr           'year (1997, 1998, etc)  for header

    'dan m 8-19-08
    For ilLoop = 0 To 9
        'tmGrf.iPerGenl(ilLoop + 9) = tgChfUS.iSlfCode(ilLoop)
        'tmGrf.lDollars(ilLoop + 9) = tgChfUS.lComm(ilLoop)
        tmGrf.iPerGenl(ilLoop + 8) = tgChfUS.iSlfCode(ilLoop)
        tmGrf.lDollars(ilLoop + 8) = tgChfUS.lComm(ilLoop)
    Next ilLoop
    tmGrf.sGenDesc = tgChfUS.sProduct

    'Format  quarterly totals  into output buffer
    For ilLoop = 1 To 4
        llProject(ilLoop) = llProject(ilLoop) \ 100          'store the penniless value to avoid redoing it later for splits
        llCalcGross = llCalcGross + llProject(ilLoop)
        tmGrf.lDollars(ilLoop - 1) = llProject(ilLoop)
    Next ilLoop
    If llCalcGross <> 0 Then

        'tmGrf.lDollars(5) = llCalcGross                        'year total
        tmGrf.lDollars(4) = llCalcGross                        'year total

        'Test for selective offices
        ilFoundOne = True
        'Process for the split offices (max 10)
        If tgChfUS.lComm(0) = 0 Then              'only 1 slsp must be 100%
            tgChfUS.lComm(0) = 1000000
        End If
        '4-15-00 obtain the office for the primary slsp
        ilPrimarySof = 0
        ilPrimarySubCo = tgChfUS.iMnfSubCmpy(0)
        For ilLoop = 0 To UBound(tlSofList)
            If tgChfUS.iSlfCode(0) = tlSofList(ilLoop).iMnfSSCode Then    'imnfsscode was replaced with slf code during the build of array
                ilPrimarySof = tlSofList(ilLoop).iSofCode
                Exit For
            End If
        Next ilLoop

        'Process without splits only------------
        'find the associated office in memory table from the slsp and see if it should be shown
        'For ilTemp = 0 To UBound(tlSofList)         'loop thru the slsp codes to find the associated office
        '11-9-10 fix office selectivity
        

        For ilTemp = 0 To RptSelUS!lbcSelection(0).ListCount - 1
            'If RptSelUS!lbcSelection.Selected(ilTemp) - 1 Then            'selected advt
            If RptSelUS!lbcSelection(0).Selected(ilTemp) Then            'selected advt
                slNameCode = tgSOCode(ilTemp).sKey      'pick up slsp code
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                If Val(slCode) = ilPrimarySof Then    'test the selcted office code against the slsps office code
                    tmGrf.iSofCode = ilPrimarySof
                    '4-15-00 If ilSlspLoop = 0 Then          'without splits for primary office (for major sort field)
                    If ilPrimarySubCo = tgChfUS.iMnfSubCmpy(ilSlspLoop) Then          'primary, do without splits
                        If gFilterLists(tgChfUS.iSlfCode(ilSlspLoop), imIncludeCodes, imUseCodes()) Then      '7-11-19 test for valid slsp selection
                            'tmGrf.iPerGenl(6) = 0
                            tmGrf.iPerGenl(5) = 0
                            tmGrf.iSlfCode = tgChfUS.iSlfCode(ilSlspLoop)
                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                            Exit For
                        End If
                    End If
                End If                      'val(slcode) = tlsoflist(illoop).isofcode
            End If                          'rptselUS!lbcselection.selected
        Next ilTemp


        For ilQtrs = 1 To 4
            llTempProject(ilQtrs) = 0
        Next ilQtrs
        ilLastSlsProc = 0


        ReDim llSlfSplit(0 To 9) As Long           '4-15-00 slsp rev share %
        ReDim ilSlfCode(0 To 9) As Integer         '4-15-00
        ReDim llSlfSplitRev(0 To 9) As Long         '2-2-04 unused in this report (need for subroutine)

        ilMnfSubCo = gGetSubCmpy(tgChfUS, ilSlfCode(), llSlfSplit(), tmClf.iVefCode, False, llSlfSplitRev())
        For ilSlspLoop = 0 To 9
            '4-12-00 If tgChfUS.islfCode(ilSlspLoop) > 0 And (tgChfUS.iMnfSubCmpy(ilSlspLoop) = 0 Or tgChfUS.iMnfSubCmpy(ilSlspLoop) = ilMnfSubCo) Then  '4-10-00
            If ilSlfCode(ilSlspLoop) > 0 And llSlfSplit(ilSlspLoop) > 0 Then   '4-15-00
                'find the associated office in memory table from the slsp
                For ilLoop = 0 To UBound(tlSofList)         'loop thru the slsp codes to find the associated office
                    '4-15-00 If tgChfUS.islfCode(ilSlspLoop) = tlSofList(ilLoop).imnfSSCode Then       'find matching slsp code in memory table
                    If ilSlfCode(ilSlspLoop) = tlSofList(ilLoop).iMnfSSCode Then       'find matching slsp code in memory table (mnfSSCode has slsp code stored in it)
                        ilSaveSof = tlSofList(ilLoop).iSofCode
                        For ilTemp = 0 To RptSelUS!lbcSelection(0).ListCount - 1 Step 1
                            If RptSelUS!lbcSelection(0).Selected(ilTemp) Then              'selected advt
                                slNameCode = tgSOCode(ilTemp).sKey      'pick up slsp code
                                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                                If Val(slCode) = tlSofList(ilLoop).iSofCode Then    'test the selcted office code against the slsps office code
                                    If gFilterLists(ilSlfCode(ilSlspLoop), imIncludeCodes, imUseCodes()) Then      '7-11-19 test for valid slsp selection

                                        tmGrf.iSofCode = ilSaveSof
                                        '4-15-00 remove
                                        'If ilSlspLoop = 0 Then          'without splits for primary office (for major sort field)
                                        '    tmGrf.iPerGenl(6) = 0
                                        '    tmGrf.islfCode = tgChfUS.islfCode(ilSlspLoop)
                                        '    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                        'End If
                                        'also, always write records withs splits
                                        'tmGrf.iPerGenl(6) = 1       'with splits (for major sort field)
                                        tmGrf.iPerGenl(5) = 1       'with splits (for major sort field)
                                        ilLastSlsProc = ilSlfCode(ilSlspLoop)   '4-15-00
                                        'If tgChfUS.islfCode(ilSlspLoop) > 0 Then            'splits involved, determine split amounts
                                            For ilQtrs = 1 To 4
                                                slAmount = gLongToStrDec(llProject(ilQtrs), 0)              'cents have already been removed
                                                'slSharePct = gLongToStrDec(tgChfUS.lComm(ilSlspLoop), 4)                    'slsp split share in %
                                                slSharePct = gLongToStrDec(llSlfSplit(ilSlspLoop), 4)                    'slsp split share in %
                                                slTemp = gDivStr(gMulStr(slSharePct, slAmount), "100")         'slsp gross portion of possible split
                                                slTemp = gRoundStr(slTemp, "01.", 0)
                                                tmGrf.lDollars(ilQtrs - 1) = Val(slTemp)  'no cents
                                                llTempProject(ilQtrs) = llTempProject(ilQtrs) + tmGrf.lDollars(ilQtrs - 1) '4-15-00
                                            Next ilQtrs
                                            '4-15-00 tmGrf.islfCode = tgChfUS.islfCode(ilSlspLoop)
                                            tmGrf.iSlfCode = ilSlfCode(ilSlspLoop)
                                            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                            ilLoop = UBound(tlSofList)          'force to get out of this loop
                                            Exit For
                                        'End If
                                    End If                      'val(slcode) = tlsoflist(illoop).isofcode
                                End If                  'if gfilterlist
                            End If                          'rptselUS!lbcselection.selected
                        Next ilTemp
                    End If
                Next ilLoop
            'Else               '4-15-00
            '    ilFoundOne = False
            '    Exit For                    'no more offices, exit the loop
            End If
        Next ilSlspLoop
        '4-15-00 If the totals dollars split dont equal the orig amount, create an extra record to avoid loss of balancing
        For ilQtrs = 1 To 4         'handle the loss of rounded dollars due to splits
            tmGrf.lDollars(ilQtrs - 1) = llProject(ilQtrs) - llTempProject(ilQtrs)
        Next ilQtrs
        
        '11-9-10 if not doing all offices, dont try to make up for the total $ not distributed
        If RptSelUS!ckcAll.Value = vbChecked Then
            'update with adjusted extra/under $ only if there was a slsp processed
            'If ilLastSlsProc > 0 And (tmGrf.lDollars(1) <> 0 Or tmGrf.lDollars(2) <> 0 Or tmGrf.lDollars(3) <> 0 Or tmGrf.lDollars(4) <> 0) Then
            If ilLastSlsProc > 0 And (tmGrf.lDollars(0) <> 0 Or tmGrf.lDollars(1) <> 0 Or tmGrf.lDollars(2) <> 0 Or tmGrf.lDollars(3) <> 0) Then
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            End If
        End If

        For ilLoop = 1 To 4
            llProject(ilLoop) = 0
            llTempProject(ilLoop) = 0
        Next ilLoop
        llCalcGross = 0
    End If                                      'llcalcgross <> 0
End Sub
