Attribute VB_Name = "RptCrID"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptcrid.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
Dim tmChfAdvtExt() As CHFADVTEXT
Dim tmContract() As SORTCODE
Dim imMktCode As Integer
Dim tmMktVefCode() As Integer
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer
Dim tmClf As CLF
Dim imClfRecLen As Integer
Dim hmCff As Integer            'Contract line flight file handle
Dim imCffRecLen As Integer        'CFF record length
Dim tmCff As CFF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim tmGrf As GRF

Dim hmSdf As Integer        'Spot recd
Dim tmSdf As SDF    'SDF record image of Spot recd
Dim imSdfRecLen As Integer        'SDF record length

Dim hmSmf As Integer        'MG recd
Dim tmSmf As SMF    'SmF record image of mg recd
Dim imSmfRecLen As Integer        'SmF record length

Dim hmVsf As Integer            'Virtual Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim tmVsfSrchKey As LONGKEY0            'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
Dim imChgMode As Integer    'Change mode status (so change not entered when in change)

Dim imTerminate As Integer  'True = terminating task, False= OK
                            '4=Invoice Date; 5=Combine ID; 6=Ref Inv #; 7=Tax1; 8=Tax2; 9=Referenced Ordered; 10=Ordered Gross; 11=Comm Pct)
Dim imChg As Integer
Dim smStartStd As String    'Starting date for standard billing
Dim smEndStd As String      'Ending date for standard billing
Dim lmStartStd As Long    'Starting date for standard billing
Dim lmEndStd As Long      'Ending date for standard billing
Dim imMarketIndex As Integer
Dim tmChfInfo() As CHFINFO
Dim tmChfOutInfo() As CHFINFO
Dim imDiscrepOnly As Integer

Dim tmSdfExtSort() As SDFEXTSORT
Dim tmSdfExt() As SDFEXT



Type CHFINFO            'array of all vehicles ordered or aired for month requested
    lChfCode As Long
    lCntNo As Long
    iVefCode As Integer
    iSpotsOrd As Long
    lAmtOrd As Long
    iSpotsAired As Long
    lAmtAired As Long
    iBonus As Long
End Type





'*******************************************************
'*                                                     *
'*      Procedure Name:mProcFlight                     *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Sub mProcFlightID(ilCff As Integer, slSFlightDate As String, slEFlightDate As String, ilPass As Integer, slPctTrade As String, slTotalNoPerWk As String, slTotalRate As String)
'
'   Where
'       ilCff(I)- Flight record index
'       slSFlightDate(I)- Flight Start date
'       slEFlightDate(I)- Flight End Date
'       slTotalNoPerWk(O)- Running Total number of spots per week
'       slTotalRate(I/O)- Ordered Total $'s
'

    Dim llDate As Long
    Dim ilDay As Integer
    Dim llSDate As Long
    Dim slRate As String
    'Get flight rate
    Select Case tgCffID(ilCff).CffRec.sPriceType
        Case "T"    'True
            slRate = gLongToStrDec(tgCffID(ilCff).CffRec.lActPrice, 2)
            If (ilPass = 0) And (Val(slPctTrade) <> 0) Then
                slRate = gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100")
            ElseIf (ilPass = 1) And (Val(slPctTrade) <> 100) Then
                slRate = gSubStr(RTrim$(slRate), gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100"))
            End If
        Case "N"    'No Charge
            slRate = "N/C"
        Case "M"    'MG Line
            slRate = "MG"
        Case "B"    'Bonus
            slRate = "Bonus"
        Case "S"    'Spinoff
            slRate = "Spinoff"
        Case "P"    'Package
            slRate = gLongToStrDec(tgCffID(ilCff).CffRec.lActPrice, 2)
            If (ilPass = 0) And (Val(slPctTrade) <> 0) Then
                slRate = gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100")
            ElseIf (ilPass = 1) And (Val(slPctTrade) <> 100) Then
                slRate = gSubStr(RTrim$(slRate), gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100"))
            End If
        Case "R"    'Recapturable
            slRate = "Recapturable"
        Case "A"    'ADU
            slRate = "ADU"
    End Select
    If (tgCffID(ilCff).CffRec.sDyWk <> "D") Then    'Weekly
        'take out calendar testing
        llDate = gDateValue(slSFlightDate)
        Do While llDate <= gDateValue(slEFlightDate)
            If (llDate >= lmStartStd) And (llDate <= lmEndStd) Then
                slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffID(ilCff).CffRec.iSpotsWk)))
                If InStr(RTrim$(slRate), ".") > 0 Then
                    slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffID(ilCff).CffRec.iSpotsWk))))
                End If
            End If
            If llDate > lmEndStd Then
                Exit Do
            End If
            llDate = gDateValue(gObtainNextMonday(Format$(llDate + 1, "m/d/yy")))
        Loop
    Else    'Daily
        'take out calendar testing
        If gDateValue(slSFlightDate) >= lmStartStd Then
            llSDate = gDateValue(slSFlightDate)
        Else
            llSDate = lmStartStd
        End If
        For llDate = llSDate To gDateValue(slEFlightDate) Step 1
            If (llDate >= lmStartStd) And (llDate <= lmEndStd) Then
                ilDay = gWeekDayLong(llDate)
                slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffID(ilCff).CffRec.iDay(ilDay))))
                If InStr(RTrim$(slRate), ".") > 0 Then
                    slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffID(ilCff).CffRec.iDay(ilDay)))))
                End If
            End If
            If llDate >= lmEndStd Then
                Exit For
            End If
        Next llDate
    End If
End Sub

'*************************************************************
'*                                                           *
'*      Procedure Name:gCrInvImpWS                           *
'*      Compares Contract ordered against Spot File aired    *
'*           Created:2/18            By:D. Hosaka            *
'
'                                                            *
'                                                            *
'*************************************************************
Sub gIDSummary()

    Dim ilRet As Integer    'Return Status
    Dim ilLoop As Integer
    Dim slStr1 As String
    Dim ilYear As Integer
    Dim ilCurrentMonth As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim llSingleCntr As Long
    Dim slTempStart As String
    Dim llChfCode As Long
    Dim ilChf As Integer
    Dim ilValidCntr As Integer


    Screen.MousePointer = vbHourglass

    sgCntrForDateStamp = ""

    imMarketIndex = -1
    imTerminate = False
    imChg = False
    imChgMode = False
    tmGrf.sGenDesc = UCase$(RptSelID!edcSelCFrom.Text)
    tmGrf.sGenDesc = Trim$(tmGrf.sGenDesc) & " " & Trim$(RptSelID!edcSelCFrom1.Text)
    If imTerminate Then
        Exit Sub
    End If

    slTempStart = RptSelID!edcContract  'single contract # requested
    llSingleCntr = Val(slTempStart)

    imDiscrepOnly = False
    If RptSelID!CkcDiscrepOnly.Value = vbChecked Then
        imDiscrepOnly = True
    End If

    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gIDSummaryErr
    gBtrvErrorMsg ilRet, "gIDSummary (btrOpen: Chf.Btr)", RptSelID
    On Error GoTo 0
    imCHFRecLen = Len(tmChf) 'btrRecordLength(hmChf)    'Get Chf size
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gIDSummaryErr
    gBtrvErrorMsg ilRet, "gIDSummary (btrOpen: Clf.Btr)", RptSelID
    On Error GoTo 0
    imClfRecLen = Len(tmClf) 'btrRecordLength(hmClf)    'Get Clf size
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gIDSummaryErr
    gBtrvErrorMsg ilRet, "gIDSummary (btrOpen: Cff.Btr)", RptSelID
    On Error GoTo 0
    imCffRecLen = Len(tmCff) 'btrRecordLength(hmCff)    'Get Cff size

    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gIDSummaryErr
    gBtrvErrorMsg ilRet, "gIDSummary (btrOpen: Vsf.Btr)", RptSelID
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)

    hmGrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gIDSummaryErr
    gBtrvErrorMsg ilRet, "gIDSummary (btrOpen: Grf.Btr)", RptSelID
    On Error GoTo 0
    imGrfRecLen = Len(tmGrf)

    hmSdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSdf, "", sgDBPath & "SDf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gIDSummaryErr
    gBtrvErrorMsg ilRet, "gIDSummary (btrOpen: Sdf.Btr)", RptSelID
    On Error GoTo 0
    imSdfRecLen = Len(tmSdf)

    hmSmf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gIDSummaryErr
    gBtrvErrorMsg ilRet, "gIDSummary (btrOpen: Smf.Btr)", RptSelID
    On Error GoTo 0
    imSmfRecLen = Len(tmSmf)

    'obtain vehicle and advertiser lists
    ilRet = gObtainVef()
    ilRet = gObtainAdvt()

    slStr1 = RptSelID!edcSelCFrom.Text             'month in text form (jan..dec)
    ilYear = Val(RptSelID!edcSelCFrom1.Text)
    gGetMonthNoFromString UCase$(slStr1), ilCurrentMonth      'getmonth #
    slStr1 = Trim$(str$(ilCurrentMonth)) & "/15/" & Trim$(RptSelID!edcSelCFrom1.Text)   'format xx/15/xxxx

    smStartStd = gObtainStartStd(slStr1)               'obtain std start date for month
    lmStartStd = gDateValue(smStartStd)
    smEndStd = gObtainEndStd(slStr1)                 'obtain std end date for month
    lmEndStd = gDateValue(smEndStd)

    'From the markets selected, build array of valid vehicle names belonging in those markets
    ReDim tmMktVefCode(0 To 0) As Integer   'init veh list for the mkts selected
    For imMarketIndex = 0 To RptSelID!lbcSelection(0).ListCount - 1 'loop on markets
        If (RptSelID!lbcSelection(0).Selected(imMarketIndex)) Then      'market selected?
            slNameCode = tgMktCode(imMarketIndex).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imMktCode = Val(slCode)
            'maintain a list of the valid vehicles for the markets selected
            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If tgMVef(ilLoop).iMnfVehGp3Mkt = imMktCode Then
                    tmMktVefCode(UBound(tmMktVefCode)) = tgMVef(ilLoop).iCode

                    ReDim Preserve tmMktVefCode(0 To UBound(tmMktVefCode) + 1) As Integer
                    'Exit For
                End If
            Next ilLoop
        End If
    Next imMarketIndex
   ' Find the contracts to process - either single or all active contracts for a 3-month period
   '1 in the past, current, & 1 in future for any makegoods in the current month that doesnt have
   'an active contract in the current month

    mGetCntrID llSingleCntr          'gather contracts for selected markets (for 3 months in order to get contracts
                                    'that have expired with makegoods.  Sort the array by contract code


    'Build images and gather ordered spots & amounts from the contract

    For ilChf = 0 To UBound(tmContract) - 1 Step 1
        ReDim tmChfInfo(0 To 0) As CHFINFO       'active contracts
        ReDim tmChfOutInfo(0 To 0) As CHFINFO    'vehicles not found in the active contract list
        slNameCode = tmContract(ilChf).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        llChfCode = Val(slCode)
        ilValidCntr = mObtainCntrInfoID(llChfCode)      'gather ordered info for current contract
        'gather all the spots for this vehicle for the requested month
        If ilValidCntr Then
            ilRet = gObtainCntrSpot(-1, False, llChfCode, -1, "S", smStartStd, smEndStd, tmSdfExtSort(), tmSdfExt(), 0, False)

            mBuildSDFCounts     'spots have been gathered and placed in array, loop thru spots and
                                'get rates and accum spots & $ aired
            mWriteID         'write prepass records for crystal
        End If

    Next ilChf


    Erase tmMktVefCode, tmContract, tmChfAdvtExt
    Erase tmChfInfo, tmChfOutInfo
    Erase tmSdfExtSort, tmSdfExt


    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSmf)

    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmVsf
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmSdf
    btrDestroy hmSmf

    Exit Sub
gIDSummaryErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub
'
'
'           Dump all the arrays into GRF record to print from Crystal
'
'   tmGrf.iGenDate(0-1) = Generation Date
'   tmGrf.iGenTime(0-1) = Generation Time
'   tmGrf.iSofCode = Market code (mnf)
'   tmGf.ivefCode = vehicle code (vef)
'   tmGrf.lChfCode - Contract code (chf)
'   tmGrf.iPerGenl(1) - Ordered spot count
'   tmGrf.iPerGenl(3) - spot count aired & billed (paid spots)
'   tmGrf.iPerGenl(4) - remote spot count billed (bonus spots)
'   tmGrf.lDollars(1) -  Ordered gross $
'   tmGrf.lDollars(3) -  Aired & billed gross $
'
'
Sub mWriteID()
Dim ilRet As Integer
Dim ilLoop As Integer
    'format remainder of record
    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
    'tmGrf.iGenTime(1) = igNowTime(1)
    tmGrf.lGenTime = lgNowTime
    tmGrf.iSofCode = imMktCode              'Market code

    For ilLoop = LBound(tmChfInfo) To UBound(tmChfInfo) - 1
        tmGrf.iVefCode = tmChfInfo(ilLoop).iVefCode      'vehicle code
        tmGrf.lChfCode = tmChfInfo(ilLoop).lChfCode      'contract code
        'tmGrf.iPerGenl(1) = tmChfInfo(ilLoop).iSpotsOrd ' ordered spots
        tmGrf.iPerGenl(0) = tmChfInfo(ilLoop).iSpotsOrd ' ordered spots
        'tmGrf.lDollars(1) = tmChfInfo(ilLoop).lAmtOrd          '  gross
        tmGrf.lDollars(0) = tmChfInfo(ilLoop).lAmtOrd          '  gross
        'tmGrf.iPerGenl(3) = tmChfInfo(ilLoop).iSpotsAired  'Aired  (billed) spots
        tmGrf.iPerGenl(2) = tmChfInfo(ilLoop).iSpotsAired  'Aired  (billed) spots
        'tmGrf.iPerGenl(4) = tmChfInfo(ilLoop).iBonus 'aired (billed) bonus spots
        tmGrf.iPerGenl(3) = tmChfInfo(ilLoop).iBonus 'aired (billed) bonus spots
        'tmGrf.lDollars(3) = tmChfInfo(ilLoop).lAmtAired 'Aired (billed)  gross
        tmGrf.lDollars(2) = tmChfInfo(ilLoop).lAmtAired 'Aired (billed)  gross
        ''If tmGrf.iPerGenl(1) + tmGrf.lDollars(1) + tmGrf.iPerGenl(3) + tmGrf.iPerGenl(4) + tmGrf.lDollars(3) <> 0 Then  'something exists for this contract,
        'If tmGrf.iPerGenl(1) + tmGrf.lDollars(0) + tmGrf.iPerGenl(3) + tmGrf.iPerGenl(4) + tmGrf.lDollars(2) <> 0 Then  'something exists for this contract,
        If tmGrf.iPerGenl(0) + tmGrf.lDollars(0) + tmGrf.iPerGenl(2) + tmGrf.iPerGenl(3) + tmGrf.lDollars(2) <> 0 Then  'something exists for this contract,
            'determine if its in balance and whether discreps only requested
            ''If (Not imDiscrepOnly) Or ((tmGrf.iPerGenl(1) <> tmGrf.iPerGenl(3) + tmGrf.iPerGenl(4)) Or (tmGrf.lDollars(1) <> tmGrf.lDollars(3)) And imDiscrepOnly) Then
            'If (Not imDiscrepOnly) Or ((tmGrf.iPerGenl(1) <> tmGrf.iPerGenl(3) + tmGrf.iPerGenl(4)) Or (tmGrf.lDollars(0) <> tmGrf.lDollars(2)) And imDiscrepOnly) Then
            If (Not imDiscrepOnly) Or ((tmGrf.iPerGenl(0) <> tmGrf.iPerGenl(2) + tmGrf.iPerGenl(3)) Or (tmGrf.lDollars(0) <> tmGrf.lDollars(2)) And imDiscrepOnly) Then
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            End If
        End If
    Next ilLoop


    For ilLoop = LBound(tmChfOutInfo) To UBound(tmChfOutInfo) - 1
        tmGrf.iVefCode = tmChfOutInfo(ilLoop).iVefCode      'vehicle code
        tmGrf.lChfCode = tmChfOutInfo(ilLoop).lChfCode      'contract code
        'tmGrf.iPerGenl(1) = tmChfOutInfo(ilLoop).iSpotsOrd ' ordered spots
        tmGrf.iPerGenl(0) = tmChfOutInfo(ilLoop).iSpotsOrd ' ordered spots
        'tmGrf.lDollars(1) = tmChfOutInfo(ilLoop).lAmtOrd          '  gross
        tmGrf.lDollars(0) = tmChfOutInfo(ilLoop).lAmtOrd          '  gross
        'tmGrf.iPerGenl(3) = tmChfOutInfo(ilLoop).iSpotsAired  'Aired  (billed) spots
        tmGrf.iPerGenl(2) = tmChfOutInfo(ilLoop).iSpotsAired  'Aired  (billed) spots
        'tmGrf.iPerGenl(4) = tmChfOutInfo(ilLoop).iBonus 'aired (billed) bonus spots
        tmGrf.iPerGenl(3) = tmChfOutInfo(ilLoop).iBonus 'aired (billed) bonus spots
        'tmGrf.lDollars(3) = tmChfOutInfo(ilLoop).lAmtAired 'Aired (billed)  gross
        tmGrf.lDollars(2) = tmChfOutInfo(ilLoop).lAmtAired 'Aired (billed)  gross
        ''If tmGrf.iPerGenl(1) + tmGrf.lDollars(1) + tmGrf.iPerGenl(3) + tmGrf.iPerGenl(4) + tmGrf.lDollars(3) <> 0 Then  'something exists for this contract,
        'If tmGrf.iPerGenl(1) + tmGrf.lDollars(0) + tmGrf.iPerGenl(3) + tmGrf.iPerGenl(4) + tmGrf.lDollars(2) <> 0 Then  'something exists for this contract,
        If tmGrf.iPerGenl(0) + tmGrf.lDollars(0) + tmGrf.iPerGenl(2) + tmGrf.iPerGenl(3) + tmGrf.lDollars(2) <> 0 Then  'something exists for this contract,
            'determine if its in balance and whether discreps only requested
            ''If (Not imDiscrepOnly) Or ((tmGrf.iPerGenl(1) <> tmGrf.iPerGenl(3) + tmGrf.iPerGenl(4)) Or (tmGrf.lDollars(1) <> tmGrf.lDollars(3)) And imDiscrepOnly) Then
            'If (Not imDiscrepOnly) Or ((tmGrf.iPerGenl(1) <> tmGrf.iPerGenl(3) + tmGrf.iPerGenl(4)) Or (tmGrf.lDollars(0) <> tmGrf.lDollars(2)) And imDiscrepOnly) Then
            If (Not imDiscrepOnly) Or ((tmGrf.iPerGenl(0) <> tmGrf.iPerGenl(2) + tmGrf.iPerGenl(3)) Or (tmGrf.lDollars(0) <> tmGrf.lDollars(2)) And imDiscrepOnly) Then
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            End If
        End If
    Next ilLoop

End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetCntr                        *
'*                                                     *
'*             Created:9/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Get Contracts for market       *
'*                                                     *
'*******************************************************
Sub mGetCntrID(llSingleCntr As Long)
    Dim slStatus As String
    Dim slCntrType As String
    Dim ilHOType As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVeh As Integer
    Dim ilVefCode As Integer
    Dim ilFound As Integer
    Dim ilChf As Integer
    Dim slKey As String
    Dim slStr As String
    Dim slTemp As String
    Dim slStartStd As String
    Dim slEndStd As String


    slTemp = Format$(lmStartStd - 1, "m/d/yy")       'go back one month
    slStartStd = gObtainStartStd(slTemp)               'obtain std start date for month
    slTemp = Format$(lmEndStd + 1, "m/d/yy")
    slEndStd = gObtainEndStd(slTemp)

    slStatus = "HO"
    slCntrType = ""
    ilHOType = 2        'GN supercedes HO if exists
    ReDim tmContract(0 To 0) As SORTCODE
    If llSingleCntr > 0 Then          'get only contract # requested
        'ReDim tmChfAdvtExt(1 To 2) As CHFADVTEXT   'fake out the array so its common code
        ReDim tmChfAdvtExt(0 To 1) As CHFADVTEXT   'fake out the array so its common code
        tmChfSrchKey1.lCntrNo = llSingleCntr
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If tmChf.lCntrNo <> llSingleCntr Then
            Exit Sub
        Else                           'fake out an entry in tmchfadvtext  for the single contract
            'tmChfAdvtExt(1).lCntrNo = llSingleCntr
            'tmChfAdvtExt(1).lCode = tmChf.lCode
            'tmChfAdvtExt(1).iStartDate(0) = tmChf.iStartDate(0)
            'tmChfAdvtExt(1).iStartDate(1) = tmChf.iStartDate(1)
            'tmChfAdvtExt(1).iEndDate(0) = tmChf.iEndDate(0)
            'tmChfAdvtExt(1).iEndDate(1) = tmChf.iEndDate(1)

            'slStr = Trim$(str$(tmChfAdvtExt(1).lCode))
            'Do While Len(slStr) < 10
            '    slStr = "0" & slStr
            'Loop

            'tmContract(UBound(tmContract)).sKey = slStr & "\" & Trim$(str$(tmChfAdvtExt(1).lCode))
            tmChfAdvtExt(0).lCntrNo = llSingleCntr
            tmChfAdvtExt(0).lCode = tmChf.lCode
            tmChfAdvtExt(0).iStartDate(0) = tmChf.iStartDate(0)
            tmChfAdvtExt(0).iStartDate(1) = tmChf.iStartDate(1)
            tmChfAdvtExt(0).iEndDate(0) = tmChf.iEndDate(0)
            tmChfAdvtExt(0).iEndDate(1) = tmChf.iEndDate(1)

            slStr = Trim$(str$(tmChfAdvtExt(0).lCode))
            Do While Len(slStr) < 10
                slStr = "0" & slStr
            Loop

            tmContract(UBound(tmContract)).sKey = slStr & "\" & Trim$(str$(tmChfAdvtExt(0).lCode))
            ReDim Preserve tmContract(0 To UBound(tmContract) + 1) As SORTCODE
        End If
    Else
        ilRet = gObtainCntrForDate(RptSelID, slStartStd, slEndStd, slStatus, slCntrType, ilHOType, tmChfAdvtExt())

        For ilChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
            For ilVeh = 0 To UBound(tmMktVefCode) - 1 Step 1
                ilVefCode = tmMktVefCode(ilVeh)
                ilFound = False
                If tmChfAdvtExt(ilChf).lVefCode > 0 Then
                    If tmChfAdvtExt(ilChf).lVefCode = ilVefCode Then
                        ilFound = True
                    End If
                ElseIf tmChfAdvtExt(ilChf).lVefCode < 0 Then
                    tmVsfSrchKey.lCode = -tmChfAdvtExt(ilChf).lVefCode
                    ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                        If tmVsf.iFSCode(ilLoop) > 0 Then
                            If tmVsf.iFSCode(ilLoop) = ilVefCode Then
                                ilFound = True
                                Exit For
                            End If
                        End If
                    Next ilLoop
               End If
                If ilFound Then
                    'Contract Code
                    slKey = Trim$(str$(tmChfAdvtExt(ilChf).lCode))
                    Do While Len(slKey) < 10
                        slKey = "0" & slKey
                    Loop

                    tmContract(UBound(tmContract)).sKey = slKey & "\" & Trim$(str$(tmChfAdvtExt(ilChf).lCode))
                    ReDim Preserve tmContract(0 To UBound(tmContract) + 1) As SORTCODE
                    Exit For
                End If
            Next ilVeh
        Next ilChf
    End If
    If UBound(tmContract) > 0 Then
        'ArraySortTyp fnAV(tmContract(), 0), UBound(tmContract), 0, LenB(tmContract(0)), 0, LenB(tmContract(0).sKey), 0
        ArraySortTyp fnAV(tmContract(), 0), UBound(tmContract), 0, LenB(tmContract(0)), 0, LenB(tmContract(0).sKey), 0

    End If

End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mProcFlight                     *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Function mObtainCntrInfoID(llChfCode As Long) As Integer
    Dim ilRet As Integer
    Dim ilPass As Integer
    Dim ilClf As Integer
    Dim slSFlightDate As String
    Dim slEFlightDate As String
    Dim ilIncludeFlight As Integer
    Dim slTotalNoPerWk As String
    Dim slTotalRate As String
    Dim slPctTrade As String
    Dim ilCff As Integer
    Dim ilLoopCnt As Integer
    Dim ilFoundEntry As Integer
    Dim llActPrice As Long
    Dim ilSearchVeh As Integer

    mObtainCntrInfoID = False
    ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llChfCode, False, tgChfID, tgClfID(), tgCffID())
    If ilRet Then

        'see if anypart of the contract is within the markts selected. Contract
        'could be multi-part with one vehicle a cluster, the others non-cluster.
        'if so, show it
        ilFoundEntry = False
        For ilClf = LBound(tgClfID) To UBound(tgClfID) - 1 Step 1
            For ilSearchVeh = LBound(tmMktVefCode) To UBound(tmMktVefCode) - 1
                If tgClfID(ilClf).ClfRec.iVefCode = tmMktVefCode(ilSearchVeh) Then
                    mObtainCntrInfoID = True
                    ilFoundEntry = True
                    Exit For
                End If
            Next ilSearchVeh
        Next ilClf
        If Not ilFoundEntry Then
            Exit Function
        End If

        For ilClf = LBound(tgClfID) To UBound(tgClfID) - 1 Step 1
            tmClf = tgClfID(ilClf).ClfRec
            If (tmClf.sType = "S") Or (tmClf.sType = "H") Then

                ilFoundEntry = False
                For ilLoopCnt = LBound(tmChfInfo) To UBound(tmChfInfo) - 1
                    If tmChfInfo(ilLoopCnt).lChfCode <= tmClf.lChfCode Then
                        If tmClf.lChfCode = tmChfInfo(ilLoopCnt).lChfCode And tmClf.iVefCode = tmChfInfo(ilLoopCnt).iVefCode Then
                            ilFoundEntry = True
                            Exit For
                        End If
                    Else
                        ilFoundEntry = False
                        Exit For
                    End If
                Next ilLoopCnt
                If Not ilFoundEntry Then    'create new entry, then update values gathered
                    ilLoopCnt = UBound(tmChfInfo)
                    tmChfInfo(ilLoopCnt).lChfCode = tmClf.lChfCode
                    tmChfInfo(ilLoopCnt).iVefCode = tmClf.iVefCode
                    ReDim Preserve tmChfInfo(0 To ilLoopCnt + 1) As CHFINFO
                End If
            'If (tmClf.sType = "S") Or (tmClf.sType = "H") Then
                ilCff = tgClfID(ilClf).iFirstCff
                Do While ilCff <> -1
                    gUnpackDate tgCffID(ilCff).CffRec.iStartDate(0), tgCffID(ilCff).CffRec.iStartDate(1), slSFlightDate
                    gUnpackDate tgCffID(ilCff).CffRec.iEndDate(0), tgCffID(ilCff).CffRec.iEndDate(1), slEFlightDate
                    ilIncludeFlight = True
                    If (gDateValue(slSFlightDate) > lmEndStd) Or (gDateValue(slEFlightDate) < lmStartStd) Then
                        ilIncludeFlight = False
                    End If
                    'Test if CBS
                    If gDateValue(slEFlightDate) < gDateValue(slSFlightDate) Then
                        ilIncludeFlight = False
                    End If
                    If ilIncludeFlight Then
                        slTotalNoPerWk = "0"
                        slTotalRate = "0"
                        'get the $ and # spots ordered for this week
                        mProcFlightID ilCff, slSFlightDate, slEFlightDate, ilPass, slPctTrade, slTotalNoPerWk, slTotalRate
                        If (InStr(slTotalRate, ".") <> 0) Then        'found spot cost
                            llActPrice = gStrDecToLong(slTotalRate, 2)
                        Else
                            llActPrice = 0
                        End If
                        tmChfInfo(ilLoopCnt).iSpotsOrd = tmChfInfo(ilLoopCnt).iSpotsOrd + Val(slTotalNoPerWk)
                        tmChfInfo(ilLoopCnt).lAmtOrd = tmChfInfo(ilLoopCnt).lAmtOrd + llActPrice
                    End If
                    ilCff = tgCffID(ilCff).iNextCff
                Loop
            End If
        Next ilClf
    End If
End Function
'
'
'       Determine $ and # spots scheduled and build in contract memory array
'       with the ordered $ (by vehicle).  If an entry isn't found, create a new one
'       in a different array
'
'

Public Sub mBuildSDFCounts()
'Dim ilLoopSpots As Integer
Dim llLoopSpots As Long
Dim ilClf As Integer
'Dim ilSdfIndex As Integer
Dim llSdfIndex As Long
Dim ilVehicle As Integer
Dim ilVefFound As Integer
Dim ilUseOutInfo As Integer
Dim ilCff As Integer
Dim slSFlightDate As String
Dim slEFlightDate As String
Dim ilIncludeFlight As Integer
Dim llActPrice As Long
Dim slStr As String
Dim llDateTest As Long

    'For ilLoopSpots = LBound(tmSdfExtSort) To UBound(tmSdfExtSort) - 1
    For llLoopSpots = LBound(tmSdfExtSort) To UBound(tmSdfExtSort) - 1
        'ilSdfIndex = tmSdfExtSort(ilLoopSpots).iSdfExtIndex
        llSdfIndex = tmSdfExtSort(llLoopSpots).lSdfExtIndex
        'If tmSdfExt(ilSdfIndex).sSchStatus = "S" Or tmSdfExt(ilSdfIndex).sSchStatus = "G" Or tmSdfExt(ilSdfIndex).sSchStatus = "O" Then
        If tmSdfExt(llSdfIndex).sSchStatus = "S" Or tmSdfExt(llSdfIndex).sSchStatus = "G" Or tmSdfExt(llSdfIndex).sSchStatus = "O" Then
            'find the matching line to process this spot, plus either find or create an entry for
            'this vehicle to hold the ordered info

            For ilClf = LBound(tgClfID) To UBound(tgClfID) - 1
                tmClf = tgClfID(ilClf).ClfRec
                If tmClf.iLine = tmSdfExt(llSdfIndex).iLineNo Then
                    'Now find the matching vehicle entry for this contract
                    'ilCntIndex is the first vehicle belonging to this contract
                    ilUseOutInfo = False
                    ilVefFound = False
                    For ilVehicle = LBound(tmChfInfo) To UBound(tmChfInfo) - 1
                        If tmChfInfo(ilVehicle).iVefCode = tmSdfExt(llSdfIndex).iVefCode Then
                            ilVefFound = True
                            Exit For
                        End If
                    Next ilVehicle
                    If Not ilVefFound Then      'create an entry to place this one
                        'see if this entry is in the contract list of vehicles not ordered for the contract
                        ilUseOutInfo = True
                        ilVefFound = False
                        For ilVehicle = LBound(tmChfOutInfo) To UBound(tmChfOutInfo) - 1
                            If tmChfOutInfo(ilVehicle).lChfCode = tmClf.lChfCode And tmChfOutInfo(ilVehicle).iVefCode = tmSdfExt(llSdfIndex).iVefCode Then
                                ilVefFound = True
                                Exit For
                            End If
                        Next ilVehicle
                        If Not ilVefFound Then
                            ilVehicle = UBound(tmChfOutInfo)
                            tmChfOutInfo(ilVehicle).lChfCode = tmClf.lChfCode
                            tmChfOutInfo(ilVehicle).iVefCode = tmClf.iVefCode
                            ReDim Preserve tmChfOutInfo(0 To ilVehicle + 1) As CHFINFO
                        End If
                    End If
                    Exit For
                End If
            Next ilClf

            'Need to determine which flight belongs with this flight to get spot rate
            'If the spot is a Fill/Bonus, no rates apply; add up all the bonus aired
            If tmSdfExt(llSdfIndex).sSpotType = "X" Then
                If ilUseOutInfo Then
                    tmChfOutInfo(ilVehicle).iBonus = tmChfOutInfo(ilVehicle).iBonus + 1
                Else
                    tmChfInfo(ilVehicle).iBonus = tmChfInfo(ilVehicle).iBonus + 1
                End If
            Else

                ilCff = tgClfID(ilClf).iFirstCff
                Do While ilCff <> -1
                    ilIncludeFlight = True
                    gUnpackDate tgCffID(ilCff).CffRec.iStartDate(0), tgCffID(ilCff).CffRec.iStartDate(1), slSFlightDate
                    gUnpackDate tgCffID(ilCff).CffRec.iEndDate(0), tgCffID(ilCff).CffRec.iEndDate(1), slEFlightDate

                    If tmSdfExt(llSdfIndex).sSchStatus = "G" Or tmSdfExt(llSdfIndex).sSchStatus = "O" Then
                        llDateTest = tmSdfExt(llSdfIndex).lMdDate
                    Else
                        gUnpackDate tmSdfExt(llSdfIndex).iDate(0), tmSdfExt(llSdfIndex).iDate(1), slStr
                        llDateTest = gDateValue(slStr)
                    End If
                    If (gDateValue(slSFlightDate) > llDateTest) Or (gDateValue(slEFlightDate) < llDateTest) Then
                        ilIncludeFlight = False
                    End If
                    'Test if CBS
                    If gDateValue(slEFlightDate) < gDateValue(slSFlightDate) Then
                        ilIncludeFlight = False

                    End If
                    If ilIncludeFlight Then
                        'obtain spot rate for scheduled spots only
                        If tmSdfExt(llSdfIndex).sSchStatus = "S" Or tmSdfExt(llSdfIndex).sSchStatus = "G" Or tmSdfExt(llSdfIndex).sSchStatus = "O" Then
                           llActPrice = tgCffID(ilCff).CffRec.lActPrice

                           'For this report, the spot types are not necessary; only need the rate unless its a Fill spot, in which
                           'case we its zero
                           ' Select Case tgCffID(ilCff).CffRec.sPriceType
                           '      Case "T"    'True
                           '          slTotalRate = gLongToStrDec(tgCffID(ilCff).CffRec.lActPrice, 2)    'tlCff.lActPrice, 2)
                           '      Case "N"    'No Charge
                           '          slTotalRate = "N/C"
                           '      Case "M"    'MG Line
                           '          slTotalRate = "MG"
                           '      Case "B"    'Bonus
                           '          slTotalRate = "Bonus"
                           '      Case "S"    'Spinoff
                           '          slTotalRate = "Spinoff"
                           '      Case "P"    'Package
                           '          slTotalRate = ".00"
                           '      Case "R"    'Recapturable
                           '          slTotalRate = "Recapturable"
                           '     Case "A"    'ADU
                           '          slTotalRate = "ADU"
                            'End Select

                            'accum the spot  & $ totals
                            'If (InStr(slTotalRate, ".") <> 0) Then        'found spot cost
                            '    llActPrice = gStrDecToLong(slTotalRate, 2)
                            'Else
                            '    llActPrice = 0
                            'End If


                            If ilUseOutInfo Then        'these buckets are for those spots that are not on the contract and
                                'moved as outsides or makegoods
                                tmChfOutInfo(ilVehicle).iSpotsAired = tmChfOutInfo(ilVehicle).iSpotsAired + 1
                                tmChfOutInfo(ilVehicle).lAmtAired = tmChfOutInfo(ilVehicle).lAmtAired + llActPrice
                            Else
                                tmChfInfo(ilVehicle).iSpotsAired = tmChfInfo(ilVehicle).iSpotsAired + 1
                                tmChfInfo(ilVehicle).lAmtAired = tmChfInfo(ilVehicle).lAmtAired + llActPrice
                            End If
                        End If
                        'if the spot was accounted for, force exit and process next spot
                        ilCff = -1
                    Else
                        ilCff = tgCffID(ilCff).iNextCff
                    End If
                Loop
            End If
        End If          'bypass missed spots
    Next llLoopSpots
End Sub
