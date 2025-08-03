Attribute VB_Name = "RptCrRI"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RPTCRRI.BAS on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmSbfSrchKey1                                                                         *
'******************************************************************************************

Option Explicit
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
Dim tmChfAdvtExt() As CHFADVTEXT
Dim tmContract() As SORTCODE
Dim imMktCode As Integer
Dim tmMktVefCode() As Integer
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
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
Dim hmSbf As Integer        'Special billing
Dim tmSbf As SBF
Dim tmSbfSrchKey2 As SBFKEY2    'SBF key record image
Dim imSbfRecLen As Integer        'SBF record length

Dim hmRwf As Integer        'Weekly rep posting
Dim tmRwf As RWF
Dim imRwfRecLen As Integer        'RWF record length
Dim tmRwfSrchKey As LONGKEY0      'RWF record image

Dim hmVbf As Integer        'Barter agreement
Dim tmVbf As VBF
Dim imVBfRecLen As Integer        'VBF record length
Dim tmVbfSrchKey1 As VBFKEY1
Dim tmVbfSrchKey0 As LONGKEY0


Dim hmVsf As Integer            'Virtual Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim tmVsfSrchKey As LONGKEY0            'VSF record image
Dim imVsfRecLen As Integer        'VSF record length

Dim hmVaf As Integer            'Vehicle Accts file handle
Dim tmVaf() As VAF                'VAF record image
Dim tmVafSrchKey As LONGKEY0            'VAF record image
Dim imVafRecLen As Integer        'VAF record length


Dim hmAdf As Integer
Dim tmAdf As ADF
Dim imAdfRecLen As Integer
Dim tmAdfSrchKey As INTKEY0

Dim hmRvf As Integer
Dim tmRvf As RVF
Dim imRvfRecLen As Integer
Dim tmRvfSrchKey4 As RVFKEY4
Dim hmPhf As Integer

Dim imChgMode As Integer    'Change mode status (so change not entered when in change)

Dim imTerminate As Integer  'True = terminating task, False= OK
Dim smSave() As String      '11-12-02: Values saved (1=C or T; 2=PctTrade; 3=Ordered No Spots; 4=Ordered Gross; 5=Aired No Spots;
                            '6=Aired Gross; 7= Bonus No Spots; 8=Billed, 9=Calendar Carry; 10=Missed carry; 11 Received Date;
                            '12=Posted Date; 13=MGs; 14=Missed; 15=Initial Aired Spots;16=Next billing no spots-for calendar posting)
Dim imSave() As Integer     'Values saved (1=AdfCode; 2=VefCode; 3=lbcVehicle Index; 4=Valid vehicle, 5 = checked)
Dim lmSave() As Long        'Values saved(1=ChfCode; 2=SbfCode, 3= spot price)
Dim smInfo() As String * 12      'Import Info for Right Mouse (1=Source[I=Import;F=Sbf;C=Contract;T=Total;S=Insert]; 2=Export Date; 3=Import Date;
                            '4=Invoice Date; 5=Combine ID; 6=Ref Inv #; 7=Tax1; 8=Tax2; 9=Referenced Ordered; 10=Ordered Gross; 11=Comm Pct)
Dim imChg As Integer
Dim smStartStd As String    'Starting date for standard billing
Dim smEndStd As String      'Ending date for standard billing
Dim smStartCal As String    'Starting date for standard billing
Dim smEndCal As String      'Ending date for standard billing
Dim lmStartStd As Long    'Starting date for standard billing
Dim lmEndStd As Long      'Ending date for standard billing
Dim lmStartCal As Long
Dim lmEndCal As Long
Dim imMarketIndex As Integer
Dim imTaxDefined As Integer
Dim tmAffByMonth() As AFFBYMONTH
Dim tmSelectiveList() As Integer
Dim tmREPVehicleList() As Integer
Dim tmAffOrdered() As AFFORDERED
Dim ilmnfMinorCode As Integer
Dim ilMnfMajorCode As Integer

Dim imYear As Integer               'Barter report requested year
Dim imCurrentMonth As Integer       'barter report requested month

Type YEARLY_TRACKING
    iVefCode As Integer
    lVbfcode As Long
    lUsedItems As Long
End Type

Type BARTERSTATS
    iVefCode As Integer     'vehicle code
    lCntrNo As Long         'contract #
    lChfCode As Long        'contract code
    lWeek As Long           'start week
    iLen As Integer         'spot length
    iOrderedSpots As Integer    'ordered spot count
    iAiredSpots As Integer      'aired spot count
    lMinAired As Long           'minutes aired in seconds
    iFree As Integer            '# free spots
    iPayable As Integer         '# payable spots
    lBarterValue As Long        '$ value for the spot
    iPostDate(0 To 1) As Integer    'posted date
    sBarterPaid As String * 1             'station paid flag (Y/N)
    iAdfCode As Integer         'advertiser code
    sProduct As String * 35
    sCashTrade As String * 1    'C = cash, T = 100 or partial trade
    iLenForSortMethodX As Integer   'spot length for method X only, and if period is month or year.  Payment is by spot length
    lVbfcode As Long
    iGetSbfStartDate(0 To 1) As Integer        'start date gathering sbf records
    iGetSbfEndDate(0 To 1) As Integer        'end date gathering sbf records
    bIgnoreForXport As Boolean          'true if this set of statistics should be ignored to update the export.  Barter method X
                                        '(buy x and get Y free for monthly and yearly period intervals), does not have a contract
                                        'associated with the payables due station.  Therefore, pick one contract in for each spot
                                        'length within the contract and export that one; ignore all others for same spot length becuase
                                        'the same payable values will be in all records for the matching spot length.
    iTotalAiredByLen As Integer         '10-18-19 total spots aired by spot length for calc of rev
    lRevShareByLen As Long              '10-19-18 % of revenue by cnt
    sMethod As String * 1               '10-19-18 barter posting method from vbf
    sPeriodForMethodX As String * 1     '10-19-18 for barter method X, which period interval are $ determined; required for export portion
    lInvNo As Long                      '10-19-18
End Type

Type BARTERTYPESORT
    sKey As String * 80 'vehicle name (max 40)| date (5) | contract # descending (9)| spot length (3)
    tBarterStats As BARTERSTATS
End Type

Dim tmBarterTypeSort() As BARTERTYPESORT        'array of SBF to process and their stats, sorted by vehicle name, cntr # (descending), week & spot length
Dim tmYearly_Tracking() As YEARLY_TRACKING      'total minutes or units for the yearly barters to keep track of the balance remaining

Const LBONE = 1



'
'
'               Rep Invoice Worksheet report subroutine
'               mObtainCodes - get all codes to process or exclude
'               When selecting advt--make testing
'               of selection more efficient.  If more than half of
'               the entries are selected, create an array with entries
'               to exclude.  If less than half of entries are selected,
'               create an array with entries to include.
'               <input> ilListIndex - list box to test
'                       lbcListbox - array containing sort codes
'               <output> ilIncludeCodes - true if test to include the codes in array
'                                          false if test to exclude the codes in array
'                        ilUseCodes - array of advt, agy or vehicles codes to include/exclude
Sub mObtainCodes(Form As Form, ilListIndex As Integer, lbcListBox() As SORTCODE, ilIncludeCodes, ilUseCodes() As Integer)
Dim ilHowManyDefined As Integer
Dim ilHowMany As Integer
Dim slNameCode As String
Dim ilLoop As Integer
Dim slCode As String
Dim ilRet As Integer
    ilHowManyDefined = Form!lbcSelection(ilListIndex).ListCount
    'ilHowMany = RptSel!lbcSelection(ilListIndex).SelectCount
    ilHowMany = Form!lbcSelection(ilListIndex).SelCount
    If ilHowMany > ilHowManyDefined / 2 Then    'more than half selected
        ilIncludeCodes = False
    Else
        ilIncludeCodes = True
    End If
    For ilLoop = 0 To Form!lbcSelection(ilListIndex).ListCount - 1 Step 1
        slNameCode = lbcListBox(ilLoop).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        If Form!lbcSelection(ilListIndex).Selected(ilLoop) And ilIncludeCodes Then               'selected ?
            ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
            ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
        Else        'exclude these
            If (Not Form!lbcSelection(ilListIndex).Selected(ilLoop)) And (Not ilIncludeCodes) Then
                ilUseCodes(UBound(ilUseCodes)) = Val(slCode)
                ReDim Preserve ilUseCodes(LBound(ilUseCodes) To UBound(ilUseCodes) + 1)
            End If
        End If
    Next ilLoop
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mRecomputeTotals                *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Recompute all totals           *
'*                                                     *
'*******************************************************
Sub mRecomputeTotals()
    Dim slONoSpots As String
    Dim slOGross As String
    Dim slANoSpots As String
    Dim slAGross As String
    Dim slMissedNoSpots As String
    Dim slMGNoSpots As String
    Dim slBonusNoSpots As String
    Dim slCalCarry As String
    Dim slPrior As String
    Dim slCalPrev As String
    Dim ilRow As Integer
    Dim ilStartRow As Integer
    Dim ilEndRow As Integer
    Dim slStr As String
    Dim ilAdj As Integer
    Dim ilComputeMG As Integer
    Dim ilAdf As Integer

    ilStartRow = LBONE  'LBound(smSave, 2)
    ilEndRow = ilStartRow + 1
    If ilEndRow >= UBound(smSave, 2) Then
        Exit Sub
    End If
    If igPostType = 4 Then
        ilAdj = 1
    Else
        ilAdj = 0
    End If
    Do
        If (InStr(1, smInfo(1, ilEndRow), "T", 1) > 0) Then
            slONoSpots = "0"
            slOGross = "0.00"
            slANoSpots = "0"
            slAGross = "0.00"
            slMissedNoSpots = "0"
            slMGNoSpots = "0"
            slBonusNoSpots = "0"
            slCalCarry = "0"
            slCalPrev = "0"
            slPrior = "0"
            For ilRow = ilStartRow To ilEndRow - 1 Step 1
                'Compute # Missed
                slStr = ""
                If (smSave(5, ilRow) <> "") Then
                    slStr = smSave(5, ilRow)
                Else
                    slStr = "0"
                End If
                If (tgSpf.sPostCalAff = "C") And (smSave(9, ilRow) <> "") Then
                    slStr = gAddStr(slStr, smSave(9, ilRow))
                End If
                If (slStr <> "") Then
                    slStr = gSubStr(slStr, smSave(3, ilRow))
                    If Val(slStr) >= 0 Then
                        slStr = ""
                    End If
                End If
                smSave(14, ilRow) = slStr
                'Compute # MGs
                ilComputeMG = True
                'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                '    If tgCommAdf(ilAdf).iCode = imSave(1, ilRow) Then
                    ilAdf = gBinarySearchAdf(imSave(1, ilRow))
                    If ilAdf <> -1 Then
                        If tgCommAdf(ilAdf).sAllowRepMG = "N" Then
                            ilComputeMG = False
                        End If
                '        Exit For
                    End If
                'Next ilAdf

                slStr = ""
                If (smSave(5, ilRow) <> "") Then
                    slStr = smSave(5, ilRow)
                End If
                If (tgSpf.sPostCalAff = "C") And (smSave(9, ilRow) <> "") Then
                    slStr = gAddStr(slStr, smSave(9, ilRow))
                End If
                If ilComputeMG Then
                    If (slStr <> "") Then
                        slStr = gSubStr(slStr, smSave(3, ilRow))
                        If Val(slStr) > 0 Then
                            If Val(smSave(10, ilRow)) < 0 Then
                                If Val(slStr) > Abs(Val(smSave(10, ilRow))) Then
                                    slStr = str$(Abs(Val(smSave(10, ilRow))))
                                End If
                            Else
                                slStr = ""
                            End If
                        Else
                            slStr = ""
                        End If
                    End If
                Else
                    slStr = ""
                End If
                smSave(13, ilRow) = slStr
                'Compute Bonus
                slStr = ""
                If (smSave(5, ilRow) <> "") Then
                    slStr = smSave(5, ilRow)
                End If
                If (tgSpf.sPostCalAff = "C") And (smSave(9, ilRow) <> "") Then
                    slStr = gAddStr(slStr, smSave(9, ilRow))
                End If
                If (slStr <> "") Then
                    slStr = gSubStr(slStr, smSave(3, ilRow))
                    If Val(slStr) > 0 Then
                        If Val(smSave(10, ilRow)) <= 0 Then
                            If ilComputeMG Then
                                If Val(slStr) > Abs(Val(smSave(10, ilRow))) Then
                                    slStr = str$(Val(slStr) - Abs(Val(smSave(10, ilRow))))
                                Else
                                    slStr = ""
                                End If
                            Else
                                slStr = slStr
                            End If
                        Else
                            slStr = ""
                        End If
                    Else
                        slStr = ""
                    End If
                End If
                smSave(7, ilRow) = slStr
                'Compute Aired Gross
                slStr = gAddStr(smSave(3, ilRow), smSave(14, ilRow))
                slStr = gAddStr(slStr, smSave(13, ilRow))
                smSave(6, ilRow) = gLongToStrDec(Val(slStr) * lmSave(3, ilRow), 2)
                'Compute running values
                slONoSpots = gAddStr(slONoSpots, smSave(3, ilRow))
                slOGross = gAddStr(slOGross, smSave(4, ilRow))
                slANoSpots = gAddStr(slANoSpots, smSave(5, ilRow))
                slMissedNoSpots = gAddStr(slMissedNoSpots, smSave(14, ilRow))
                slMGNoSpots = gAddStr(slMGNoSpots, smSave(13, ilRow))
                slBonusNoSpots = gAddStr(slBonusNoSpots, smSave(7, ilRow))
                slAGross = gAddStr(slAGross, smSave(6, ilRow))
                slCalCarry = gAddStr(slCalCarry, smSave(16, ilRow))
                slCalPrev = gAddStr(slCalPrev, smSave(9, ilRow))
                slPrior = gAddStr(slPrior, smSave(10, ilRow))
            Next ilRow
            smSave(3, ilEndRow) = slONoSpots
            smSave(4, ilEndRow) = slOGross
            smSave(5, ilEndRow) = slANoSpots
            smSave(6, ilEndRow) = slAGross
            smSave(14, ilEndRow) = slMissedNoSpots
            smSave(13, ilEndRow) = slMGNoSpots
            smSave(7, ilEndRow) = slBonusNoSpots
            smSave(16, ilEndRow) = slCalCarry
            smSave(9, ilEndRow) = slCalPrev
            smSave(10, ilEndRow) = slPrior

            ilStartRow = ilEndRow + 1
            ilEndRow = ilStartRow + 1
        Else
            ilEndRow = ilEndRow + 1
        End If
    Loop While ilEndRow <= UBound(smSave, 2) - ilAdj

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
Sub mProcFlightRI(ilCff As Integer, slSFlightDate As String, slEFlightDate As String, ilPass As Integer, slPctTrade As String, slTotalNoPerWk As String, slTotalRate As String)
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
    Select Case tgCffRI(ilCff).CffRec.sPriceType
        Case "T"    'True
            slRate = gLongToStrDec(tgCffRI(ilCff).CffRec.lActPrice, 2)
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
            slRate = gLongToStrDec(tgCffRI(ilCff).CffRec.lActPrice, 2)
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
    If (tgCffRI(ilCff).CffRec.sDyWk <> "D") Then    'Weekly
            'take out calendar testing
            llDate = gDateValue(slSFlightDate)
            Do While llDate <= gDateValue(slEFlightDate)
                If (llDate >= lmStartStd) And (llDate <= lmEndStd) Then
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffRI(ilCff).CffRec.iSpotsWk)))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffRI(ilCff).CffRec.iSpotsWk))))
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
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffRI(ilCff).CffRec.iDay(ilDay))))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffRI(ilCff).CffRec.iDay(ilDay)))))
                    End If
                End If
                If llDate >= lmEndStd Then
                    Exit For
                End If
            Next llDate
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mGetCarryRI                     *
'*                                                     *
'*             Created:11/13/02       By:D. Hosaka     *
'*             copy of routine from Post Rep           *
'*                                                     *
'*            Comments: Get spots carried from previous*
'*                      month                          *
'*                                                     *
'*******************************************************
Sub mGetCarryRI()
    Dim ilRet As Integer    'Return status
    Dim llDate As Long
    Dim llStartStd As Long
    Dim llEndStd As Long
    Dim slDate As String
    Dim ilRow As Integer
    Dim ilPass As Integer
    Dim ilFound As Integer

    If (imMarketIndex >= 0) Then
        slDate = gDecOneDay(smStartStd)
        slDate = gObtainStartStd(slDate)
        llStartStd = gDateValue(slDate)
        slDate = gObtainEndStd(slDate)
        llEndStd = gDateValue(slDate)
        tmSbfSrchKey2.sTranType = "T"
        'tmSbfSrchKey2.iDate(0) = 0
        'tmSbfSrchKey2.iDate(1) = 0
        gPackDateLong llStartStd, tmSbfSrchKey2.iDate(0), tmSbfSrchKey2.iDate(1)
        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.sTranType = "T")
            gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
            If llDate > llEndStd Then
                Exit Do
            End If
            If (tmSbf.iMissCarryOver < 0) Or ((tmSbf.iCalCarryOver > 0) And (tgSpf.sPostCalAff = "C") And (igPostType <= 4)) Then    '8/12/03 might need to test for <= 2 or take test out
                ilFound = False
                For ilPass = 0 To 1 Step 1
                    'For ilRow = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
                    For ilRow = LBONE To UBound(smSave, 2) - 1 Step 1
                    If (InStr(1, smInfo(1, ilRow), "T", 1) <= 0) Then
                        'If (InStr(1, smShow(imTOTALINDEX, ilRow), "Total:", 1) <= 0) Then
                            If (tmSbf.lChfCode = lmSave(1, ilRow)) And (tmSbf.iAirVefCode = imSave(2, ilRow)) And (tmSbf.sCashTrade = smSave(1, ilRow)) And ((tmSbf.lSpotPrice = lmSave(3, ilRow)) Or (ilPass = 1)) Then
                                If (tmSbf.iMissCarryOver < 0) Then
                                    If Trim$(smSave(10, ilRow)) = "" Then
                                        smSave(10, ilRow) = Trim$(str$(tmSbf.iMissCarryOver))
                                    Else
                                        smSave(10, ilRow) = Trim$(str$(Val(smSave(10, ilRow)) + tmSbf.iMissCarryOver))
                                    End If
                                End If
                                If ((tmSbf.iCalCarryOver > 0) And (tgSpf.sPostCalAff = "C") And (igPostType <= 4)) Then
                                    If Trim$(smSave(9, ilRow)) = "" Then
                                        smSave(9, ilRow) = Trim$(str$(tmSbf.iCalCarryOver))
                                    Else
                                        smSave(9, ilRow) = Trim$(str$(Val(smSave(9, ilRow)) + tmSbf.iCalCarryOver))
                                    End If
                                End If
                                ilFound = True
                                Exit For
                            End If
                        End If
                    Next ilRow
                    If ilFound Then
                        Exit For
                    End If
                Next ilPass
            End If
            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
        'For ilRow = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
        '    If (InStr(1, smShow(imTOTALINDEX, ilRow), "Total:", 1) <= 0) Then
       '         If imAPREVNOSPOTSINDEX > 0 Then
       '             gSetShow pbcPostRep, smSave(9, ilRow), tmCtrls(imAPREVNOSPOTSINDEX)
       '             smShow(imAPREVNOSPOTSINDEX, ilRow) = tmCtrls(imAPREVNOSPOTSINDEX).sShow
       '         End If
       '         If imPRIORNOSPOTSINDEX > 0 Then
       '             gSetShow pbcPostRep, smSave(10, ilRow), tmCtrls(imPRIORNOSPOTSINDEX)
       '             smShow(imPRIORNOSPOTSINDEX, ilRow) = tmCtrls(imPRIORNOSPOTSINDEX).sShow
       '         End If
       '     End If
       ' Next ilRow
    End If
    Exit Sub

    On Error GoTo 0
    Exit Sub
End Sub
'********************************************************************
'*                                                                  *
'*      Procedure Name:gCheck_affidavits                            *
'*                                                                  *
'*           Created:9-29-02           By:D. Hosaka                 *
'                                                                   *
'       '9-29-02 1)Generate report of affiliates missing            *
'       returned affidavits                                         *
'       2) generate report of unbilliable invoices                  *
'                                                                   *
'       1-13-03 Ignore contracts whose advertiser is an "Externally *
'               billed contract
'       6-18-04 Implement major sort by vehicle group
'       1-19-05 Fix subscript out of range on Unbillable report- only used for
'               Delinquent affidavits
'********************************************************************
Sub gCheck_Affs(Form As Form)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilYear                        ilCurrentMonth                                          *
'******************************************************************************************


    Dim ilRet As Integer    'Return Status
    Dim ilLoop As Integer
    Dim slStr1 As String
    Dim slCode As String
    Dim slNameCode As String
    Dim llUpper As Integer
    Dim ilChf As Integer
    Dim slStatus As String
    Dim slCntrType As String
    Dim ilHOType As Integer
    Dim slStr As String
    Dim ilListIndex As Integer
    Dim slName As String
    Dim slMonthHdr As String * 36
    Dim ilMonth As Integer
    Dim slMonth As String
    Dim ilInclRecd As Integer
    Dim ilMajorSet As Integer               '6-18-04
    Dim ilMinorSet As Integer
    ReDim tmSelectiveList(0 To 0) As Integer
    ReDim tmAffOrdered(0 To 0) As AFFORDERED
    ReDim tmAffByMonth(0 To 0) As AFFBYMONTH
    ReDim tmREPVehicleList(0 To 0) As Integer

    slMonthHdr = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"
    ilListIndex = Form!lbcRptType.ListIndex
    ilInclRecd = False
    If Form!ckcInclRecd.Value = vbChecked Then      'Include Posting received?
        ilInclRecd = True
    End If
    Screen.MousePointer = vbHourglass

    sgCntrForDateStamp = ""
    '12/17/06-Change to tax by agency or vehicle
    'If (tgSpf.iBTax(0) = 0) Or (tgSpf.iBTax(1) = 0) Then
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
        imTaxDefined = True
    Else
        imTaxDefined = False
    End If
    imMarketIndex = -1
    imTerminate = False
    imChg = False
    imChgMode = False

    If ilListIndex = INV_DELINQUENT Then        '1-19-05 delinquent can sort by vehicle group; otherwise N/A
        'tgMMnf contains the Vehicle Groups (use Ubound(tgMMnf))
        ilLoop = Form!cbcSet1.ListIndex
        ilMajorSet = tgVehicleSets1(ilLoop).iCode
    Else
        ilMajorSet = 0
        ilMinorSet = 0
    End If

    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    tmGrf.lGenTime = lgNowTime

    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCheck_AffsErr
    gBtrvErrorMsg ilRet, "gCheck_Affs (btrOpen: Chf.Btr)", Form
    On Error GoTo 0
    imCHFRecLen = Len(tmChf) 'btrRecordLength(hmChf)    'Get Chf size

    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCheck_AffsErr
    gBtrvErrorMsg ilRet, "gCheck_Affs (btrOpen: Clf.Btr)", Form
    On Error GoTo 0
    imClfRecLen = Len(tmClf) 'btrRecordLength(hmClf)    'Get Clf size

    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCheck_AffsErr
    gBtrvErrorMsg ilRet, "gCheck_Affs (btrOpen: Cff.Btr)", Form
    On Error GoTo 0
    imCffRecLen = Len(tmCff) 'btrRecordLength(hmCff)    'Get Cff size

    hmSbf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCheck_AffsErr
    gBtrvErrorMsg ilRet, "gCheck_Affs (btrOpen: Sbf.Btr)", Form
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf) 'btrRecordLength(hmSbf)    'Get Sbf size


    hmGrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCheck_AffsErr
    gBtrvErrorMsg ilRet, "gCheck_Affs (btrOpen: Grf.Btr)", Form
    On Error GoTo 0
    imGrfRecLen = Len(tmGrf)

    hmAdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gCheck_AffsErr
    gBtrvErrorMsg ilRet, "gCheck_Affs (btrOpen: Adf.Btr)", Form
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)

    'obtain vehicle and advertiser lists
    ilRet = gObtainVef()
    ilRet = gObtainAdvt()

    slStr1 = Form!edcSelCFrom.Text             'month in text form (jan..dec)
    imYear = Val(Form!edcSelCFrom1.Text)
    gGetMonthNoFromString UCase$(slStr1), imCurrentMonth      'getmonth #

    If imCurrentMonth = 0 Then                                 'input isn't text month name, try month #
        imCurrentMonth = Val(slStr1)
    End If

    slStr1 = Trim$(str$(imCurrentMonth)) & "/15/" & Trim$(Form!edcSelCFrom1.Text)   'format xx/15/xxxx

   'send the report month headings
    ilMonth = imCurrentMonth
    For ilLoop = 1 To 12
        ilMonth = (imCurrentMonth + ilLoop) - 1
        If ilMonth > 12 Then
            ilMonth = ilMonth - 12
        End If
        slStr = Mid$(slMonthHdr, (ilMonth - 1) * 3 + 1, 3)
        If ilLoop = 1 Then          'save for heading to pass to crystal
            slMonth = slStr
        End If
    Next ilLoop
    tmGrf.sGenDesc = UCase$(slMonth)
    tmGrf.sGenDesc = Trim$(tmGrf.sGenDesc) & " " & Trim$(Form!edcSelCFrom1.Text)
    If imTerminate Then
        Exit Sub
    End If

    '4-16-02 setup common date parameters
    smStartStd = gObtainStartStd(slStr1)               'obtain std start date for month
    lmStartStd = gDateValue(smStartStd)
    smEndStd = gObtainEndStd(slStr1)                 'obtain std end date for month
    lmEndStd = gDateValue(smEndStd)

    slStatus = "HO"
    slCntrType = ""
    ilHOType = 1
    ilRet = gObtainCntrForDate(Form, smStartStd, smEndStd, slStatus, slCntrType, ilHOType, tmChfAdvtExt())


    'build list of selective vehicles or advt
    llUpper = 0
    If ilListIndex = INV_DELINQUENT Then
         For ilLoop = 0 To Form!lbcSelection(1).ListCount - 1 Step 1
            If (Form!lbcSelection(1).Selected(ilLoop)) Then
                slNameCode = tgCSVNameCode(ilLoop).sKey
                ilRet = gParseItem(slNameCode, 1, "\", slName)

                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                tmSelectiveList(llUpper) = Val(slCode)
                ReDim Preserve tmSelectiveList(0 To llUpper + 1) As Integer
                llUpper = llUpper + 1
            End If
        Next ilLoop
    Else                    'Unbillable report
        For ilLoop = 0 To Form!lbcSelection(1).ListCount - 1
            If Form!lbcSelection(1).Selected(ilLoop) Then
                slStr = tgRptSelAdvertiserCode(ilLoop).sKey
                ilRet = gParseItem(slStr, 2, "\", slCode)
                tmSelectiveList(llUpper) = Val(slCode)
                ReDim Preserve tmSelectiveList(0 To llUpper + 1) As Integer
                llUpper = llUpper + 1
            End If
        Next ilLoop
        llUpper = 0
        For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1
            If tgMVef(ilLoop).sType = "R" Then      'rep vehicle
                tmREPVehicleList(llUpper) = tgMVef(ilLoop).iCode
                ReDim Preserve tmREPVehicleList(0 To llUpper + 1) As Integer
                llUpper = llUpper + 1
            End If
        Next ilLoop
    End If

    mBuildChf_Affs Form  'go thru array of active contracts and build their ordered $ and spot table
    mBuildSbf_Affs Form  'gather all billing received from affiliates from SBF

    For ilLoop = LBound(tmAffByMonth) To UBound(tmAffByMonth) - 1
        If tmAffByMonth(ilLoop).iShowFlag = 0 Or (tmAffByMonth(ilLoop).iShowFlag = 2 And ilInclRecd) Then           'No SBF found , flag wasnt set.  Or, SBF found but only received, not posted
                                                        'show two cases as unbillable
            tmGrf.sBktType = ""
            If tmAffByMonth(ilLoop).iShowFlag = 2 Then
                tmGrf.sBktType = "R"                    'show as received only since it still can't be invoiced
            End If
            tmGrf.iVefCode = tmAffByMonth(ilLoop).iVefCode
            gGetVehGrpSets tmGrf.iVefCode, ilMinorSet, ilMajorSet, ilmnfMinorCode, ilMnfMajorCode
            tmGrf.iRdfCode = ilMnfMajorCode          '6-18-04
            tmGrf.iCode2 = ilMajorSet               'sent which vehicle group selected for report heading
            tmGrf.lChfCode = tmAffByMonth(ilLoop).lChfCode
            tmGrf.iAdfCode = tmAffByMonth(ilLoop).iAdfCode
            'tmGrf.lDollars(1) = tmAffByMonth(ilLoop).lOrderedAmt    'vehicle ordered $
            'tmGrf.lDollars(2) = tmAffByMonth(ilLoop).lAiredAmt      'vehicle aired $
            'tmGrf.lDollars(3) = tmAffByMonth(ilLoop).lOrderedSpots  'vehicle ordered # spots
            'tmGrf.lDollars(4) = tmAffByMonth(ilLoop).lAiredSpots    'vehicle aired # spots
            tmGrf.lDollars(0) = tmAffByMonth(ilLoop).lOrderedAmt    'vehicle ordered $
            tmGrf.lDollars(1) = tmAffByMonth(ilLoop).lAiredAmt      'vehicle aired $
            tmGrf.lDollars(2) = tmAffByMonth(ilLoop).lOrderedSpots  'vehicle ordered # spots
            tmGrf.lDollars(3) = tmAffByMonth(ilLoop).lAiredSpots    'vehicle aired # spots
            For ilChf = 0 To UBound(tmAffOrdered) - 1
                If tmAffOrdered(ilChf).lChfCode = tmGrf.lChfCode Then
                    'tmGrf.lDollars(5) = tmAffOrdered(ilChf).lCntGross      'contract ordred gross
                    'tmGrf.lDollars(6) = tmAffOrdered(ilChf).lCntSpots
                    tmGrf.lDollars(4) = tmAffOrdered(ilChf).lCntGross      'contract ordred gross
                    tmGrf.lDollars(5) = tmAffOrdered(ilChf).lCntSpots
                End If
            Next ilChf
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
    Next ilLoop


    Erase tmAffByMonth, tmAffOrdered, tmSelectiveList, tmREPVehicleList

    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSbf)

    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmSbf
    Exit Sub
gCheck_AffsErr:
    On Error GoTo 0
    imTerminate = True
    Exit Sub
End Sub

'**************************************************************************
'*                                                                        *
'*      Procedure Name:gCrInvImpWS                                        *
'*                                                                        *
'*           Created:2/18            By:D. Hosaka                         *
'
'                                                                         *
'       1-13-03 Ignore contracts whose advertiser is an "Externally       *
'               billed contract
'       6-29-04 assume posting by advt (previous was 4 = cluster),        *
'           so that affiliates not posted will show as missed not aired   *
'                                                                         *
'**************************************************************************
Sub gRIWorkSheet(Form As Form)

    Dim ilRet As Integer    'Return Status
    Dim ilLoop As Integer
    Dim ilCol As Integer
    Dim slStr1 As String
    Dim ilYear As Integer
    Dim ilCurrentMonth As Integer
    Dim slCode As String
    Dim slNameCode As String
    Dim ilIncludeCodes As Integer               'true = include codes stored in ilusecode array,
                                            'false = exclude codes store din ilusecode array
    'ReDim ilUseCodes(1 To 1) As Integer       'valid advt (for now), agency or vehicles codes to process--
    ReDim ilUseCodes(0 To 0) As Integer       'valid advt (for now), agency or vehicles codes to process--
                                            'or advt, agy or vehicles codes not to process
    Dim ilListIndex As Integer
    Dim slMonthDesc As String * 36      'JanFeb...Dec

    Screen.MousePointer = vbHourglass
    ilListIndex = Form!lbcRptType.ListIndex
    slMonthDesc = "JanFebMarAprMayJunJulAugSepOctNovDec"

    sgCntrForDateStamp = ""
    '12/17/06-Change to tax by agency or vehicle
    'If (tgSpf.iBTax(0) = 0) Or (tgSpf.iBTax(1) = 0) Then
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
        imTaxDefined = True
    Else
        imTaxDefined = False
    End If
    imMarketIndex = -1
    imTerminate = False
    imChg = False
    imChgMode = False
    If imTerminate Then
        Exit Sub
    End If
    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRIWorkSheetErr
    gBtrvErrorMsg ilRet, "gRIWorkSheetErr (btrOpen: Chf.Btr)", Form
    On Error GoTo 0

    imCHFRecLen = Len(tmChf) 'btrRecordLength(hmChf)    'Get Chf size
    hmClf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRIWorkSheetErr
    gBtrvErrorMsg ilRet, "gRIWorkSheetErr (btrOpen: Clf.Btr)", Form
    On Error GoTo 0

    imClfRecLen = Len(tmClf) 'btrRecordLength(hmClf)    'Get Clf size
    hmCff = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRIWorkSheetErr
    gBtrvErrorMsg ilRet, "gRIWorkSheetErr (btrOpen: Cff.Btr)", Form
    On Error GoTo 0

    imCffRecLen = Len(tmCff) 'btrRecordLength(hmCff)    'Get Cff size
    hmSbf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRIWorkSheetErr
    gBtrvErrorMsg ilRet, "gRIWorkSheetErr (btrOpen: Sbf.Btr)", Form
    On Error GoTo 0

    imSbfRecLen = Len(tmSbf) 'btrRecordLength(hmSbf)    'Get Sbf size
    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRIWorkSheetErr
    gBtrvErrorMsg ilRet, "gCrWorkSheet (btrOpen: Vsf.Btr)", Form
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)

    hmGrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRIWorkSheetErr
    gBtrvErrorMsg ilRet, "gCrWorkSheet (btrOpen: Grf.Btr)", Form
    On Error GoTo 0
    imGrfRecLen = Len(tmGrf)

    hmAdf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gRIWorkSheetErr
    gBtrvErrorMsg ilRet, "gCrWorkSheet (btrOpen: Adf.Btr)", Form
    On Error GoTo 0
    imAdfRecLen = Len(tmAdf)

    'obtain vehicle and advertiser lists
    ilRet = gObtainVef()
    ilRet = gObtainAdvt()
    ilRet = gObtainAgency()

    slStr1 = Form!edcSelCFrom.Text             'month in text form (jan..dec)
    ilYear = Val(Form!edcSelCFrom1.Text)

    gGetMonthNoFromString slStr1, ilCurrentMonth         'getmonth #
    If ilCurrentMonth = 0 Then                                 'input isn't text month name, try month #
        ilCurrentMonth = Val(slStr1)
    End If
    slStr1 = Trim$(str(ilCurrentMonth)) & "/15/" & Trim$(Form!edcSelCFrom1.Text)     'form mm/dd/yy
    tmGrf.sGenDesc = Mid$(slMonthDesc, (ilCurrentMonth - 1) * 3 + 1, 3) 'UCase$(Form!edcSelCFrom.Text)
    tmGrf.sGenDesc = Trim$(tmGrf.sGenDesc) & " " & Trim$(Form!edcSelCFrom1.Text)


    '4-16-02 setup common date parameters
    smStartStd = gObtainStartStd(slStr1)               'obtain std start date for month
    lmStartStd = gDateValue(smStartStd)
    smEndStd = gObtainEndStd(slStr1)                 'obtain std end date for month
    lmEndStd = gDateValue(smEndStd)
    smEndCal = gObtainEndCal(smEndStd)
    smStartCal = gObtainStartCal(smEndCal)
    lmStartCal = gDateValue(smStartCal)
    lmEndCal = gDateValue(smEndCal)

    If ilListIndex = INV_REMOTE_INVWS And Form!rbcWSSort(1).Value Then          'filtering of advt for Rep Inv worksheet by advt
        mObtainCodes Form, 1, tgRptSelAdvertiserCode(), ilIncludeCodes, ilUseCodes()
    End If
    '3-3-03 if Rep Inv worksheet (by ADVt sort), all the markets have been selected ON
    For imMarketIndex = 0 To Form!lbcSelection(0).ListCount - 1 'loop on markets
        If (Form!lbcSelection(0).Selected(imMarketIndex)) Then      'market selected?
            ReDim tmMktVefCode(0 To 0) As Integer   'init veh list for this mkt
            'ReDim smSave(1 To 16, 1 To 1) As String
            'ReDim imSave(1 To 5, 1 To 1) As Integer
            'ReDim lmSave(1 To 3, 1 To 1) As Long
            'ReDim smInfo(1 To 11, 1 To 1) As String * 12
            ReDim smSave(0 To 16, 0 To 1) As String 'Index zero ignored
            ReDim imSave(0 To 5, 0 To 1) As Integer 'Index zero ignored
            ReDim lmSave(0 To 3, 0 To 1) As Long 'Index zero ignored
            ReDim smInfo(0 To 11, 0 To 1) As String * 12 'Index zero ignored
            For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                smInfo(ilCol, 1) = ""
            Next ilCol
            slNameCode = tgMktCode(imMarketIndex).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imMktCode = Val(slCode)
            For ilLoop = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                If tgMVef(ilLoop).iMnfVehGp3Mkt = imMktCode And tgMVef(ilLoop).sType = "R" Then  'find matching market and REP type vehicle
                    tmMktVefCode(UBound(tmMktVefCode)) = tgMVef(ilLoop).iCode
                    ReDim Preserve tmMktVefCode(0 To UBound(tmMktVefCode) + 1) As Integer
                End If
            Next ilLoop

            igPostType = 2          '6-29-04 assume posting by advt (previous was 4 = cluster), so that
                                    'affiliates not posted will show as missed not aired  igPostType = 4          'assume Cluster posting
            mGetCntrRI Form           'gather all the contracts for the billing period
            ilRet = mReadSbfRecRI(False)
            mGetCarryRI
            mRecomputeTotals
            mCrRemoteRecs Form, ilIncludeCodes, ilUseCodes()
        End If
    Next imMarketIndex

    Erase smSave, lmSave, smInfo, imSave
    Erase tmMktVefCode, tmContract, tmChfAdvtExt
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmAdf)

    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmSbf
    btrDestroy hmVsf
    btrDestroy hmAdf
    Exit Sub
gRIWorkSheetErr:
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
'   tmGrf.sBktType - C = Cash, T = trade
'   tmGf.ivefCode = vehicle code (vef)
'   tmGrf.iadfCode = advertiser code (adf)
'   tmGrf.lChfCode - Contract code (chf)
'   tmGrf.iPerGenl(1) - Central Ordered spot count
'   tmGrf.iPerGenl(2) = (unused 11-15-02) Remote Ordered spot count (only total for contract, no vehicle spot counts ordered from remote)
'   tmGrf.iPerGenl(3) - remote spot count aired & billed (paid spots)
'   tmGrf.iPerGenl(4) - remote spot count billed (bonus spots)
'   tmGrf.iPerGenl(11) - mg spots            '11-15-02
'   tmGrf.iPerGenl(6) - missed spots                '11-15-02
'   tmGrf.iPerGenl(7) - bonus spots             '11-15-02
'   tmGrf.iPerGenl(8) - prior missed            '11-15-02
'   tmGrf.iPerGenl(9) - Cal Previous            '11-15-02
'   tmGrf.iPerGenl(10) - cal next               '11'15-02
'   tmGrf.lDollars(1) - Central Ordered gross $
'   tmGrf.lDollars(2) = (Unused 11-15-02) Remote ordered spot count
'   tmGrf.lDollars(3) - remote Aired & billed gross $
'   tmGrf.lDollars(4) = Ordered spot price          '11-15-02
'   tmGrf.iCode2 = Remote agy comm % (xx.xx)
'   tmGrf.iPerGenl(5) - vehicle from remote invoice isnt from same market in the central side
'   tmGrf.lLong - Remote Invoice #
'
'
Sub mCrRemoteRecs(Form As Form, ilIncludeCodes As Integer, ilUseCodes() As Integer)
'           ilIncludeCodes - used for Rep Invoice Worksheet by Advt only.  True to test array of advt codes to include,
'                            false to exclude the array of advt codes
'           iluseCodes() array of advt codes to include or exclude for Rep Invoice Worksheet by Advt only
Dim ilRet As Integer
Dim ilLoop As Integer
Dim ilFound As Integer
Dim ilListIndex As Integer
Dim ilTemp As Integer

    ilListIndex = Form!lbcRptType.ListIndex
    'format remainder of record
    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for removal of records
    'tmGrf.iGenTime(1) = igNowTime(1)
    tmGrf.lGenTime = lgNowTime
    tmGrf.iSofCode = imMktCode              'Market code

    'For ilLoop = LBound(smSave, 1) To UBound(smSave, 2) - 2
    For ilLoop = LBONE To UBound(smSave, 2) - 2
        If Trim$(smInfo(1, ilLoop)) <> "T" Then           'ignore the total record or blank
            tmGrf.iAdfCode = imSave(1, ilLoop)      'adv code
            ilFound = False
            If ilListIndex = INV_REMOTE_INVWS And Form!rbcWSSort(1).Value Then  'if Rep Invoice worksheet by advt, test if selective advt
                If ilIncludeCodes Then
                    For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                        If ilUseCodes(ilTemp) = tmGrf.iAdfCode Then
                            ilFound = True
                            Exit For
                        End If
                    Next ilTemp
                Else
                    ilFound = True        '8/23/99 when more than half selected, selection fixed
                    For ilTemp = LBound(ilUseCodes) To UBound(ilUseCodes) - 1 Step 1
                        If ilUseCodes(ilTemp) = tmGrf.iAdfCode Then
                            ilFound = False
                            Exit For
                        End If
                    Next ilTemp
                End If
            Else
                ilFound = True          'all options will write all records
            End If

            If ilFound Then
                tmGrf.sBktType = smSave(1, ilLoop)      'C or T flag
                tmGrf.iVefCode = imSave(2, ilLoop)      'vehicle code
                tmGrf.lChfCode = lmSave(1, ilLoop)      'contract code
                'tmGrf.iPerGenl(1) = gStrDecToInt(smSave(3, ilLoop), 0)  'central ordered spots
                tmGrf.iPerGenl(0) = gStrDecToInt(smSave(3, ilLoop), 0)  'central ordered spots
                'tmGrf.lDollars(1) = gStrDecToLong(smSave(4, ilLoop), 2) 'central ordered gross
                tmGrf.lDollars(0) = gStrDecToLong(smSave(4, ilLoop), 2) 'central ordered gross
                ''tmGrf.lDollars(2) = gStrDecToLong(smInfo(10, ilLoop), 2) 'Remote ordered gross
                'tmGrf.lDollars(3) = gStrDecToLong(smSave(6, ilLoop), 2) 'Aired (billed)  gross
                tmGrf.lDollars(2) = gStrDecToLong(smSave(6, ilLoop), 2) 'Aired (billed)  gross
                'tmGrf.lDollars(4) = lmSave(3, ilLoop)       '11-15-02 ordered spot price
                tmGrf.lDollars(3) = lmSave(3, ilLoop)       '11-15-02 ordered spot price

                ''tmGrf.iPerGenl(2) = gStrDecToInt(smInfo(9, ilLoop), 0)  'Remote ordered spots
                'tmGrf.iPerGenl(3) = gStrDecToInt(smSave(5, ilLoop), 0)  'Aired  (billed) spots
                'tmGrf.iPerGenl(4) = gStrDecToInt(smSave(7, ilLoop), 0)  'aired (billed) bonus spots
                'tmGrf.iPerGenl(11) = gStrDecToInt(smSave(13, ilLoop), 0) 'mg spots   '11-15-02
                'tmGrf.iPerGenl(6) = gStrDecToInt(smSave(14, ilLoop), 0) 'missed spots   '11-15-02
                'tmGrf.iPerGenl(7) = gStrDecToInt(smSave(7, ilLoop), 0) 'bonus spots   '11-15-02
                'tmGrf.iPerGenl(8) = gStrDecToInt(smSave(10, ilLoop), 0) 'missed carry   '11-15-02
                'tmGrf.iPerGenl(9) = gStrDecToInt(smSave(9, ilLoop), 0) 'prev billng for cal    '11-15-02
                'tmGrf.iPerGenl(10) = gStrDecToInt(smSave(16, ilLoop), 0) 'next billng for cal    '11-15-02
                
                ''tmGrf.iPerGenl(2) = gStrDecToInt(smInfo(9, ilLoop), 0)  'Remote ordered spots
                tmGrf.iPerGenl(2) = gStrDecToInt(smSave(5, ilLoop), 0)  'Aired  (billed) spots
                tmGrf.iPerGenl(3) = gStrDecToInt(smSave(7, ilLoop), 0)  'aired (billed) bonus spots
                tmGrf.iPerGenl(10) = gStrDecToInt(smSave(13, ilLoop), 0) 'mg spots   '11-15-02
                tmGrf.iPerGenl(5) = gStrDecToInt(smSave(14, ilLoop), 0) 'missed spots   '11-15-02
                tmGrf.iPerGenl(6) = gStrDecToInt(smSave(7, ilLoop), 0) 'bonus spots   '11-15-02
                tmGrf.iPerGenl(7) = gStrDecToInt(smSave(10, ilLoop), 0) 'missed carry   '11-15-02
                tmGrf.iPerGenl(8) = gStrDecToInt(smSave(9, ilLoop), 0) 'prev billng for cal    '11-15-02
                tmGrf.iPerGenl(9) = gStrDecToInt(smSave(16, ilLoop), 0) 'next billng for cal    '11-15-02
                tmGrf.iCode2 = gStrDecToInt(smInfo(11, ilLoop), 2)      'remote agy comm % xx.xx
                tmGrf.lLong = gStrDecToLong(smSave(6, ilLoop), 0)       'remote invoice #
                'tmGrf.iPerGenl(5) = imSave(4, ilLoop)       'true if bad vehicle, else false
                tmGrf.iPerGenl(4) = imSave(4, ilLoop)       'true if bad vehicle, else false

                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            End If
        End If
    Next ilLoop

End Sub
'***************************************************************
'*                                                             *
'*      Procedure Name:mGetCntr                                *
'*                                                             *
'*             Created:9/02/93       By:D. LeVine              *
'*            Modified:              By:                       *
'*                                                             *
'*            Comments: Get Contracts for market               *
'*                                                             *
'*            6-29-04 retrieve the continuation records        *
'*              in VSF when # unique vehicles exceed 50        *
'*              in a contract                                  *
'***************************************************************
Sub mGetCntrRI(Form As Form)
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStatus As String
    Dim slCntrType As String
    Dim ilHOType As Integer
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilVeh As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim llChfCode As Long
    Dim ilVefCode As Integer
    Dim ilFound As Integer
    Dim ilChf As Integer
    Dim slKey As String
    Dim slStr As String
    Dim ilAdf As Integer
    Dim ilAgf As Integer



    If (lmStartStd > 0) And (lmStartCal > 0) Then
        If lmStartStd < lmStartCal Then
            slStartDate = Format$(lmStartStd, "m/d/yy")
        Else
            slStartDate = Format$(lmStartCal, "m/d/yy")
        End If
    ElseIf lmStartStd > 0 Then
        slStartDate = Format$(lmStartStd, "m/d/yy")
    Else
        slStartDate = Format$(lmStartCal, "m/d/yy")
    End If
    If (lmEndStd > 0) And (lmEndCal > 0) Then
        If lmEndStd > lmEndCal Then
            slEndDate = Format$(lmEndStd, "m/d/yy")
        Else
            slEndDate = Format$(lmEndCal, "m/d/yy")
        End If
    ElseIf lmEndStd > 0 Then
        slEndDate = Format$(lmEndStd, "m/d/yy")
    Else
        slEndDate = Format$(lmEndCal, "m/d/yy")
    End If
    slStatus = "HO"
    slCntrType = ""
    ilHOType = 1
    ilRet = gObtainCntrForDate(Form, slStartDate, slEndDate, slStatus, slCntrType, ilHOType, tmChfAdvtExt())

    ReDim tmContract(0 To 0) As SORTCODE
    For ilChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
        '1-13-03 Determine if this is an advertiser that is billed externally
        tmAdfSrchKey.iCode = tmChfAdvtExt(ilChf).iAdfCode
        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE And tmAdf.sRepInvGen <> "E" Then   'only process this contract if valid adf read and its not set to externally bill
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
                    '6-29-04 retrieve the continuation records in VSF when # unique vehicles exceed 50 in a contract
                    Do While ilRet = BTRV_ERR_NONE
                        For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                            If tmVsf.iFSCode(ilLoop) > 0 Then
                                If tmVsf.iFSCode(ilLoop) = ilVefCode Then
                                    ilFound = True
                                    Exit For
                                End If
                            End If
                        Next ilLoop
                        If ilFound Then
                            Exit Do
                        Else
                            If tmVsf.lLkVsfCode <= 0 Then
                                Exit Do
                            End If
                            tmVsfSrchKey.lCode = tmVsf.lLkVsfCode
                            ilRet = btrGetEqual(hmVsf, tmVsf, imVsfRecLen, tmVsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        End If
                    Loop


                    'For ilLoop = LBound(tmVsf.iFSCode) To UBound(tmVsf.iFSCode) Step 1
                     '   If tmVsf.iFSCode(ilLoop) > 0 Then
                     '       If tmVsf.iFSCode(ilLoop) = ilVefCode Then
                     '           ilFound = True
                     '           Exit For
                     '       End If
                     '   End If
                    'Next ilLoop
                End If
                '6/24/10: Bypass EDI contracts
                If ilFound Then
                    If tmChfAdvtExt(ilChf).iAgfCode > 0 Then
                        ilAgf = gBinarySearchAgf(tmChfAdvtExt(ilChf).iAgfCode)
                        If ilAgf <> -1 Then
                            If tgCommAgf(ilAgf).iArfInvCode > 0 Then
                                ilFound = False
                            End If
                        End If
                    Else
                        ilAdf = gBinarySearchAdf(tmChfAdvtExt(ilChf).iAdfCode)
                        If ilAdf <> -1 Then
                            If tgCommAdf(ilAdf).iArfInvCode > 0 Then
                                ilFound = False
                            End If
                        End If
                    End If
                End If
                If ilFound Then
                    slStr = "Missing"
                    'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                    '    If tgCommAdf(ilAdf).iCode = tmChfAdvtExt(ilChf).iAdfCode Then
                        ilAdf = gBinarySearchAdf(tmChfAdvtExt(ilChf).iAdfCode)
                        If ilAdf <> -1 Then
                            ''slStr = Trim$(tgCommAdf(ilAdf).sName)
                            'If (tgCommAdf(ilAdf).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdf).sAddrID) <> "") Then
                            '    slStr = Trim$(tgCommAdf(ilAdf).sName) & ", " & Trim$(tgCommAdf(ilAdf).sAddrID)
                            'Else
                                slStr = Trim$(tgCommAdf(ilAdf).sName)
                            'End If
                    '        Exit For
                        End If
                    'Next ilAdf
                    Do While Len(slStr) < 30
                        slStr = slStr & " "
                    Loop
                    slKey = slStr
                    'Contract number
                    slStr = Trim$(str$(tmChfAdvtExt(ilChf).lCntrNo))
                    Do While Len(slStr) < 8
                        slStr = "0" & slStr
                    Loop
                    slKey = slKey & slStr
                    tmContract(UBound(tmContract)).sKey = slKey & "\" & Trim$(str$(tmChfAdvtExt(ilChf).lCode))
                    ReDim Preserve tmContract(0 To UBound(tmContract) + 1) As SORTCODE
                    Exit For
                End If
            Next ilVeh
        End If
    Next ilChf
    If UBound(tmContract) > 0 Then
        'ArraySortTyp fnAV(tmContract(), 0), UBound(tmContract), 0, LenB(tmContract(0)), 0, LenB(tmContract(0).sKey), 0
        ArraySortTyp fnAV(tmContract(), 0), UBound(tmContract), 0, LenB(tmContract(0)), 0, LenB(tmContract(0).sKey), 0

    End If

    'Build images
    For ilChf = 0 To UBound(tmContract) - 1 Step 1
        slNameCode = tmContract(ilChf).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        llChfCode = Val(slCode)
        mObtainCntrInfoRI llChfCode
    Next ilChf


End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mMerge                          *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Merge SBF record into Save     *
'*                      images                         *
'*                                                     *
'*******************************************************
Function mMergeRI() As Integer
'
'
    Dim slStr As String
    Dim ilRet As Integer    'Return status
    Dim ilLoop As Integer
    Dim ilRow As Integer
    Dim ilCol As Integer
    Dim ilIndex As Integer
    Dim ilNewRow As Integer
    Dim ilPass As Integer
    Dim ilSPass As Integer
    Dim ilEPass As Integer
    Dim ilFound As Integer
    Dim ilAdf As Integer
    Dim ilVef As Integer

    ilFound = False
    'Pass zero- match on price
    'Pass 1- bypass price test
    For ilPass = 0 To 1 Step 1
        'For ilRow = LBound(smSave, 2) To UBound(smSave, 2) - 1 Step 1
        For ilRow = LBONE To UBound(smSave, 2) - 1 Step 1
            If (InStr(1, smInfo(1, ilRow), "T", 1) <= 0) Then
            'If (InStr(1, smShow(imTOTALINDEX, ilRow), "Total:", 1) <= 0) Then
                If (tmSbf.lChfCode = lmSave(1, ilRow)) And (tmSbf.iAirVefCode = imSave(2, ilRow)) And (tmSbf.sCashTrade = smSave(1, ilRow)) And ((tmSbf.lSpotPrice = lmSave(3, ilRow)) Or (ilPass = 1)) Then
                    ilFound = True
                    lmSave(2, ilRow) = tmSbf.lCode
                    If Trim$(smInfo(1, ilRow)) = "C" Then
                        If tmSbf.sPostStatus = "P" Then
                            smSave(5, ilRow) = Trim$(str$(tmSbf.iAirNoSpots))
                            smSave(16, ilRow) = Trim$(str$(tmSbf.iCalCarryOver))
                        Else
                            smSave(5, ilRow) = ""
                            smSave(16, ilRow) = ""
                        End If

                        smSave(6, ilRow) = ""   'gLongToStrDec(tmSbf.lGross, 2)
                        smSave(7, ilRow) = ""   'Trim$(Str$(tmSbf.iBonusNoSpots))
                    Else
                        If tmSbf.sPostStatus = "P" Then
                            smSave(5, ilRow) = gAddStr(smSave(5, ilRow), Trim$(str$(tmSbf.iAirNoSpots)))
                            smSave(16, ilRow) = gAddStr(smSave(16, ilRow), Trim$(str$(tmSbf.iCalCarryOver)))
                        End If
                        smSave(6, ilRow) = ""   'gAddStr(smSave(6, ilRow), gLongToStrDec(tmSbf.lGross, 2))
                        smSave(7, ilRow) = ""   'gAddStr(smSave(7, ilRow), Trim$(Str$(tmSbf.iBonusNoSpots)))
                    End If

                    smSave(15, ilRow) = smSave(5, ilRow)

                    smSave(8, ilRow) = tmSbf.sBilled

                    If (tmSbf.sPostStatus = "R") Or (tmSbf.sPostStatus = "P") Then
                        imSave(5, ilRow) = True
                    Else
                        imSave(5, ilRow) = False
                    End If
                    gUnpackDate tmSbf.iRecDate(0), tmSbf.iRecDate(1), smSave(11, ilRow)
                    gUnpackDate tmSbf.iPostDate(0), tmSbf.iPostDate(1), smSave(12, ilRow)


                    smInfo(1, ilRow) = "F"      'source
                    gUnpackDate tmSbf.iExportDate(0), tmSbf.iExportDate(1), smInfo(2, ilRow)
                    gUnpackDate tmSbf.iImportDate(0), tmSbf.iImportDate(1), smInfo(3, ilRow)
                    gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), smInfo(4, ilRow)
                    smInfo(5, ilRow) = Trim$(str$(tmSbf.iCombineID))
                    smInfo(6, ilRow) = Trim$(str$(tmSbf.lRefInvNo))
                    '12/17/06-Change to tax by agency or vehicle
                    'smInfo(7, ilRow) = gLongToStrDec(tmSbf.lTax1, 2)
                    'smInfo(8, ilRow) = gLongToStrDec(tmSbf.lTax2, 2)
                    smInfo(9, ilRow) = Trim$(str$(tmSbf.iNoItems))
                    smInfo(10, ilRow) = gLongToStrDec(tmSbf.lOGross, 2)
                    smInfo(11, ilRow) = gIntToStrDec(tmSbf.iCommPct, 2)
                    Exit For
                End If
            End If
        Next ilRow
        If ilFound Then
            Exit For
        End If
    Next ilPass
    If Not ilFound Then
        'Add in above total record for contract
        ilSPass = 0
        If (igPostType <> 2) Then
            ilEPass = 2
        Else
            'If by advertiser, don't add sbf unless unless contract matched
            'If slSource = "F" Then
                ilEPass = 1
            'Else
            '    ilEPass = 2
            'End If
        End If
        ilFound = False
        For ilPass = ilSPass To ilEPass Step 1
            'For ilRow = LBound(smSave, 2) To UBound(smSave, 2) - 2 Step 1       '-2 to adjust for Grand Total entry
            For ilRow = LBONE To UBound(smSave, 2) - 2 Step 1       '-2 to adjust for Grand Total entry
                If ((tmSbf.lChfCode = lmSave(1, ilRow)) And (tmSbf.sCashTrade = smSave(1, ilRow)) And (ilPass = 0)) Or ((tmSbf.lChfCode = lmSave(1, ilRow)) And (ilPass = 1)) Or ((ilRow = UBound(smSave, 2) - 2) And (ilPass = 2)) Then
                    ilFound = True
                    'Continue search until total record
                    For ilLoop = ilRow + 1 To UBound(smSave, 2) - 1 Step 1
                        If (InStr(1, smInfo(1, ilLoop), "T", 1) > 0) Or (lmSave(1, ilRow) <> lmSave(1, ilLoop)) Then
                        'If (InStr(1, smShow(imTOTALINDEX, ilLoop), "Total:", 1) > 0) Or (lmSave(1, ilRow) <> lmSave(1, ilLoop)) Then
                            'If ilPass = 0 Then
                                ilNewRow = ilLoop
                            'Else
                           '     ilNewRow = ilLoop + 1
                            '    If InStr(1, smShow(imTOTALINDEX, ilLoop), "Grand Total:", 1) > 0 Then
                            '        ilNewRow = ilNewRow - 1
                            '    End If
                            'End If
                            'Move all records from and including ilLoop dowm one
                            For ilIndex = UBound(smSave, 2) To ilNewRow Step -1
                                For ilCol = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                                    smSave(ilCol, ilIndex) = smSave(ilCol, ilIndex - 1)
                                Next ilCol
                                For ilCol = LBound(imSave, 1) To UBound(imSave, 1) Step 1
                                    imSave(ilCol, ilIndex) = imSave(ilCol, ilIndex - 1)
                                Next ilCol
                                For ilCol = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
                                    lmSave(ilCol, ilIndex) = lmSave(ilCol, ilIndex - 1)
                                Next ilCol
                                'For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                                '    smShow(ilCol, ilIndex) = smShow(ilCol, ilIndex - 1)
                                'Next ilCol
                                For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                                    smInfo(ilCol, ilIndex) = smInfo(ilCol, ilIndex - 1)
                                Next ilCol
                            Next ilIndex
                            'Add row
                            If (ilPass = 0) Or (igPostType = 1) Or (igPostType = 3) Then

                                'ReDim Preserve smSave(1 To 16, 1 To UBound(smSave, 2) + 1) As String
                                'ReDim Preserve imSave(1 To 5, 1 To UBound(imSave, 2) + 1) As Integer
                                'ReDim Preserve lmSave(1 To 3, 1 To UBound(lmSave, 2) + 1) As Long
                                ''ReDim Preserve smShow(1 To 10, 1 To UBound(smShow, 2) + 1) As String * 40
                                'ReDim Preserve smInfo(1 To 11, 1 To UBound(smInfo, 2) + 1) As String * 12
                            
                                ReDim Preserve smSave(0 To 16, 0 To UBound(smSave, 2) + 1) As String
                                ReDim Preserve imSave(0 To 5, 0 To UBound(imSave, 2) + 1) As Integer
                                ReDim Preserve lmSave(0 To 3, 0 To UBound(lmSave, 2) + 1) As Long
                                ReDim Preserve smInfo(0 To 11, 0 To UBound(smInfo, 2) + 1) As String * 12
                            Else
                                'ReDim Preserve smSave(1 To 16, 1 To UBound(smSave, 2) + 2) As String
                                'ReDim Preserve imSave(1 To 5, 1 To UBound(imSave, 2) + 2) As Integer
                                'ReDim Preserve lmSave(1 To 3, 1 To UBound(lmSave, 2) + 2) As Long
                                ''ReDim Preserve smShow(1 To 10, 1 To UBound(smShow, 2) + 2) As String * 40
                                'ReDim Preserve smInfo(1 To 11, 1 To UBound(smInfo, 2) + 2) As String * 12
                            
                                ReDim Preserve smSave(0 To 16, 0 To UBound(smSave, 2) + 2) As String
                                ReDim Preserve imSave(0 To 5, 0 To UBound(imSave, 2) + 2) As Integer
                                ReDim Preserve lmSave(0 To 3, 0 To UBound(lmSave, 2) + 2) As Long
                                ReDim Preserve smInfo(0 To 11, 0 To UBound(smInfo, 2) + 2) As String * 12
                            End If
                            If (ilPass = 0) Or (ilPass = 1) Then
                                For ilCol = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                                    smSave(ilCol, ilNewRow) = smSave(ilCol, ilNewRow - 1)
                                Next ilCol
                                For ilCol = LBound(imSave, 1) To UBound(imSave, 1) Step 1
                                    imSave(ilCol, ilNewRow) = imSave(ilCol, ilNewRow - 1)
                                Next ilCol
                                For ilCol = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
                                    lmSave(ilCol, ilNewRow) = lmSave(ilCol, ilNewRow - 1)
                                Next ilCol
                                'For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                                '    smShow(ilCol, ilNewRow) = smShow(ilCol, ilNewRow - 1)
                                'Next ilCol
                                For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                                    smInfo(ilCol, ilNewRow) = smInfo(ilCol, ilNewRow - 1)
                                Next ilCol
                            Else
                                For ilCol = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                                    smSave(ilCol, ilNewRow) = ""
                                Next ilCol
                                For ilCol = LBound(imSave, 1) To UBound(imSave, 1) Step 1
                                    imSave(ilCol, ilNewRow) = 0
                                Next ilCol
                                For ilCol = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
                                    lmSave(ilCol, ilNewRow) = 0
                                Next ilCol
                                'For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                                '    smShow(ilCol, ilNewRow) = ""
                                'Next ilCol
                                For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                                    smInfo(ilCol, ilNewRow) = ""
                                Next ilCol
                            End If
                            'Set values into new at ilLoop
                            smSave(1, ilNewRow) = tmSbf.sCashTrade
                            If ilPass = 0 Then
                                smSave(2, ilNewRow) = smSave(2, ilLoop - 1)
                            End If
                            smSave(3, ilNewRow) = ""   'Trim$(Str$(tmSbf.iNoItems))
                            smSave(4, ilNewRow) = ""
                            smSave(5, ilNewRow) = Trim$(str$(tmSbf.iAirNoSpots))
                            smSave(16, ilNewRow) = Trim$(str$(tmSbf.iCalCarryOver))

                            smSave(15, ilNewRow) = smSave(5, ilNewRow)

                            smSave(6, ilNewRow) = gLongToStrDec(tmSbf.lGross, 2)
                            smSave(7, ilNewRow) = ""    'Trim$(Str$(tmSbf.iBonusNoSpots))
                            smSave(8, ilNewRow) = tmSbf.sBilled
                            tmChfSrchKey.lCode = tmSbf.lChfCode
                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If ilRet = BTRV_ERR_NONE Then
                                imSave(1, ilNewRow) = tmChf.iAdfCode
                                slStr = Trim$(str$(tmChf.lCntrNo))
                                'gSetShow pbcPostRep, slStr, tmCtrls(imCONTRACTINDEX)
                                'smShow(imCONTRACTINDEX, ilNewRow) = tmCtrls(imCONTRACTINDEX).sShow
                                If ilPass <> 0 Then
                                    smSave(2, ilNewRow) = gIntToStrDec(tmChf.iPctTrade, 0)
                                End If
                            Else
                                imSave(1, ilNewRow) = -1
                                slStr = "Missing:" & Trim$(str$(tmSbf.lChfCode))
                                'gSetShow pbcPostRep, slStr, tmCtrls(imCONTRACTINDEX)
                                'smShow(imCONTRACTINDEX, ilNewRow) = tmCtrls(imCONTRACTINDEX).sShow
                                If ilPass <> 0 Then
                                    smSave(2, ilNewRow) = "0"
                                End If
                            End If
                            imSave(2, ilNewRow) = tmSbf.iAirVefCode
                            imSave(3, ilNewRow) = 0
                            lmSave(1, ilNewRow) = tmSbf.lChfCode
                            lmSave(2, ilNewRow) = tmSbf.lCode

                            lmSave(3, ilNewRow) = tmSbf.lSpotPrice
                            If (tmSbf.sPostStatus = "R") Or (tmSbf.sPostStatus = "P") Then
                                imSave(5, ilRow) = True
                            Else
                                imSave(5, ilRow) = False
                            End If
                            gUnpackDate tmSbf.iRecDate(0), tmSbf.iRecDate(1), smSave(11, ilRow)
                            gUnpackDate tmSbf.iPostDate(0), tmSbf.iPostDate(1), smSave(12, ilRow)

                            'slStr = smSave(1, ilNewRow)
                            'gSetShow pbcPostRep, slStr, tmCtrls(imCashTradeIndex)
                            'smShow(imCashTradeIndex, ilNewRow) = tmCtrls(imCashTradeIndex).sShow
                            'Advertiser
                            slStr = "Missing"
                            'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                            '    If tgCommAdf(ilAdf).iCode = imSave(1, ilNewRow) Then
                                ilAdf = gBinarySearchAdf(imSave(1, ilNewRow))
                                If ilAdf <> -1 Then
                                    ''slStr = Trim$(tgCommAdf(ilAdf).sName)
                                    'If (tgCommAdf(ilAdf).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdf).sAddrID) <> "") Then
                                    '    slStr = Trim$(tgCommAdf(ilAdf).sName) & ", " & Trim$(tgCommAdf(ilAdf).sAddrID)
                                    'Else
                                        slStr = Trim$(tgCommAdf(ilAdf).sName)
                                    'End If
                            '        Exit For
                                End If
                            'Next ilAdf
                            'If imAdvtIndex > 0 Then
                            '    gSetShow pbcPostRep, slStr, tmCtrls(imAdvtIndex)
                            '    smShow(imAdvtIndex, ilNewRow) = tmCtrls(imAdvtIndex).sShow
                            'End If
                            'Vehicle
                            slStr = "Missing"
                            'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                            '    If tgMVef(ilVef).iCode = imSave(2, ilNewRow) Then
                                ilVef = gBinarySearchVef(imSave(2, ilNewRow))
                                If ilVef <> -1 Then
                                    slStr = Trim$(tgMVef(ilVef).sName)
                            '        Exit For
                                End If
                            'Next ilVef

                            imSave(4, ilNewRow) = False
                            For ilVef = 0 To UBound(tmMktVefCode) - 1 Step 1
                                If imSave(2, ilNewRow) = tmMktVefCode(ilVef) Then
                                    imSave(4, ilNewRow) = True
                                    Exit For
                                End If
                            Next ilVef
                            If (igPostType = 4) Or (igPostType = 2) Then
                                If ilPass <> 0 Then

                                    For ilCol = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                                        smSave(ilCol, ilNewRow + 1) = ""
                                    Next ilCol
                                    For ilCol = LBound(imSave, 1) To UBound(imSave, 1) Step 1
                                        imSave(ilCol, ilNewRow + 1) = 0
                                    Next ilCol
                                    For ilCol = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
                                        lmSave(ilCol, ilNewRow + 1) = 0
                                    Next ilCol
                                    'For ilCol = LBound(smShow, 1) To UBound(smShow, 1) Step 1
                                    '    smShow(ilCol, ilNewRow + 1) = ""
                                   ' Next ilCol
                                    For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                                        smInfo(ilCol, ilNewRow + 1) = ""
                                    Next ilCol
                                    smSave(1, ilNewRow + 1) = ""
                                    smSave(2, ilNewRow + 1) = ""
                                    smSave(3, ilNewRow + 1) = ""
                                    smSave(4, ilNewRow + 1) = ""
                                    smSave(5, ilNewRow + 1) = ""
                                    smSave(15, ilNewRow + 1) = ""
                                    smSave(6, ilNewRow + 1) = ""
                                    smSave(7, ilNewRow + 1) = ""
                                    smSave(8, ilNewRow + 1) = ""
                                    smSave(9, ilNewRow + 1) = ""
                                    smSave(10, ilNewRow + 1) = ""
                                    smSave(16, ilNewRow + 1) = ""

                                    imSave(2, ilNewRow + 1) = 0
                                    imSave(3, ilNewRow + 1) = -1
                                    imSave(4, ilNewRow + 1) = True
                                    If imSave(1, ilNewRow) = -1 Then
                                        slStr = "# Missing"
                                    Else
                                        slStr = "Total: " & Trim$(str$(tmChf.lCntrNo))
                                    End If

                                    smInfo(1, ilNewRow + 1) = "T"
                                End If
                            End If
                            'If tmSbf.lCode = 0 Then
                            '    smInfo(1, ilNewRow) = "I"
                            'Else
                            '    smInfo(1, ilNewRow) = "F"
                            'End If
                            smInfo(1, ilNewRow) = "F"
                            gUnpackDate tmSbf.iExportDate(0), tmSbf.iExportDate(1), smInfo(2, ilNewRow)
                            gUnpackDate tmSbf.iImportDate(0), tmSbf.iImportDate(1), smInfo(3, ilNewRow)
                            gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), smInfo(4, ilNewRow)
                            smInfo(5, ilNewRow) = Trim$(str$(tmSbf.iCombineID))
                            smInfo(6, ilNewRow) = Trim$(str$(tmSbf.lRefInvNo))
                            '12/17/06-Change to tax by agency or vehicle
                            'smInfo(7, ilNewRow) = gLongToStrDec(tmSbf.lTax1, 2)
                            'smInfo(8, ilNewRow) = gLongToStrDec(tmSbf.lTax2, 2)
                            smInfo(9, ilNewRow) = Trim$(str$(tmSbf.iNoItems))
                            smInfo(10, ilNewRow) = gLongToStrDec(tmSbf.lOGross, 2)
                            smInfo(11, ilNewRow) = gIntToStrDec(tmSbf.iCommPct, 2)
                            Exit For
                        End If
                    Next ilLoop
                End If
                If ilFound Then
                    Exit For
                End If
            Next ilRow
            If ilFound Then
                Exit For
            End If
        Next ilPass
    End If
    mMergeRI = True
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainCntrInfoRI               *
'*                                                     *
'*      Obtain the contract and all line/flight info
'       Build into Save arrays                         *
'*                                                     *
'*******************************************************
Sub mObtainCntrInfoRI(llChfCode As Long)
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilRowNo As Integer
    Dim ilPass As Integer
    Dim ilAdf As Integer
    Dim ilClf As Integer
    Dim ilVef As Integer
    Dim ilInsertLine As Integer
    Dim slSFlightDate As String
    Dim slEFlightDate As String
    Dim ilIncludeFlight As Integer
    Dim slTotalNoPerWk As String
    Dim slTotalRate As String
    Dim slCTotalNoPerWk As String
    Dim slCTotalRate As String
    Dim slPctTrade As String
    Dim ilAddTo As Integer
    Dim ilLoop As Integer
    Dim ilCff As Integer
    Dim ilStartRowNo As Integer
    Dim ilCol As Integer
    Dim ilNoPasses As Integer
    Dim slSpotRate As String

    If igPostType = 4 Then
        ilNoPasses = 0  '1
    Else
        ilNoPasses = 0
    End If
    ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llChfCode, False, tgChfRI, tgClfRI(), tgCffRI())
    If ilRet Then
        For ilPass = 0 To ilNoPasses Step 1
            ilStartRowNo = UBound(smSave, 2)
            ilInsertLine = True
            ilRowNo = UBound(smSave, 2)
            'Contract Number
            lmSave(1, ilRowNo) = tgChfRI.lCode
            lmSave(2, ilRowNo) = 0
            slStr = Trim$(str$(tgChfRI.lCntrNo))
            'Cash/Trade flag
            slPctTrade = gIntToStrDec(tgChfRI.iPctTrade, 0)
            If ilNoPasses = 1 Then
                If (ilPass = 0) And (tgChfRI.iPctTrade <> 100) Then
                    slStr = "C"
                ElseIf (ilPass = 1) And (tgChfRI.iPctTrade <> 0) Then
                    slStr = "T"
                Else
                    ilInsertLine = False
                End If
            Else
                If tgChfRI.iPctTrade = 0 Then
                    slStr = "C"
                ElseIf tgChfRI.iPctTrade = 100 Then
                    slStr = "T"
                Else
                    slStr = "B"
                End If
            End If
            If ilInsertLine Then
                smSave(1, ilRowNo) = slStr
                smSave(2, ilRowNo) = slPctTrade
                'Advertiser
                slStr = "Missing"
                imSave(1, ilRowNo) = tgChfRI.iAdfCode
                'For ilAdf = LBound(tgCommAdf) To UBound(tgCommAdf) - 1 Step 1
                '    If tgCommAdf(ilAdf).iCode = tgChfRI.iAdfCode Then
                    ilAdf = gBinarySearchAdf(tgChfRI.iAdfCode)
                    If ilAdf <> -1 Then
                        ''slStr = Trim$(tgCommAdf(ilAdf).sName)
                        'If (tgCommAdf(ilAdf).sBillAgyDir = "D") And (Trim$(tgCommAdf(ilAdf).sAddrID) <> "") Then
                        '    slStr = Trim$(tgCommAdf(ilAdf).sName) & ", " & Trim$(tgCommAdf(ilAdf).sAddrID)
                        'Else
                            slStr = Trim$(tgCommAdf(ilAdf).sName)
                        'End If
                '        Exit For
                    End If
                'Next ilAdf

                For ilClf = LBound(tgClfRI) To UBound(tgClfRI) - 1 Step 1
                    ilRowNo = UBound(smSave, 2)
                    ilInsertLine = False
                    For ilVef = LBound(tmMktVefCode) To UBound(tmMktVefCode) - 1 Step 1
                        If tgClfRI(ilClf).ClfRec.iVefCode = tmMktVefCode(ilVef) Then
                            ilInsertLine = True
                            Exit For
                        End If
                    Next ilVef
                    If (ilInsertLine) And ((tgClfRI(ilClf).ClfRec.sType = "S") Or (tgClfRI(ilClf).ClfRec.sType = "H")) Then
                        'Vehicle
                        slStr = "Missing"
                        imSave(2, ilRowNo) = tgClfRI(ilClf).ClfRec.iVefCode
                        'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        '    If tgMVef(ilVef).iCode = tgClfRI(ilClf).ClfRec.iVefCode Then
                            ilVef = gBinarySearchVef(tgClfRI(ilClf).ClfRec.iVefCode)
                            If ilVef <> -1 Then
                                slStr = Trim$(tgMVef(ilVef).sName)
                        '        Exit For
                            End If
                        'Next ilVef
                        imSave(4, ilRowNo) = False
                        For ilVef = 0 To UBound(tmMktVefCode) - 1 Step 1
                            If imSave(2, ilRowNo) = tmMktVefCode(ilVef) Then
                                imSave(4, ilRowNo) = True
                                Exit For
                            End If
                        Next ilVef
                        slTotalNoPerWk = "0"
                        slTotalRate = "0"
                        ''Test if vehicle match, later test if spot price match
                        'ilAddTo = False
                        'For ilLoop = ilStartRowNo To ilRowNo - 1 Step 1
                        '    If imSave(2, ilLoop) = imSave(2, ilRowNo) Then
                        '        ilAddTo = True
                        '        ilRowNo = ilLoop
                        '        slTotalNoPerWk = smSave(3, ilRowNo)
                        '        slTotalRate = smSave(4, ilRowNo)
                        '        Exit For
                        '    End If
                        'Next ilLoop
                        ilCff = tgClfRI(ilClf).iFirstCff
                        Do While ilCff <> -1
                            ilRowNo = UBound(smSave, 2)
                            slTotalNoPerWk = "0"
                            slTotalRate = "0"
                            'Test if vehicle match, later test if spot price match
                            ilAddTo = False
                            For ilLoop = ilStartRowNo To ilRowNo - 1 Step 1
                                'If imSave(2, ilLoop) = imSave(2, ilRowNo) Then
                                If imSave(2, ilLoop) = tgClfRI(ilClf).ClfRec.iVefCode Then
                                    If (lmSave(3, ilLoop) = tgCffRI(ilCff).CffRec.lActPrice) Then
                                        ilAddTo = True
                                        ilRowNo = ilLoop
                                        slTotalNoPerWk = smSave(3, ilRowNo)
                                        slTotalRate = smSave(4, ilRowNo)
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                            gUnpackDate tgCffRI(ilCff).CffRec.iStartDate(0), tgCffRI(ilCff).CffRec.iStartDate(1), slSFlightDate
                            gUnpackDate tgCffRI(ilCff).CffRec.iEndDate(0), tgCffRI(ilCff).CffRec.iEndDate(1), slEFlightDate
                            ilIncludeFlight = True
                            If tgChfRI.sBillCycle = "C" Then
                                If (gDateValue(slSFlightDate) > lmEndCal) Or (gDateValue(slEFlightDate) < lmStartCal) Then
                                    ilIncludeFlight = False
                                End If
                            Else
                                If (gDateValue(slSFlightDate) > lmEndStd) Or (gDateValue(slEFlightDate) < lmStartStd) Then
                                    ilIncludeFlight = False
                                End If
                            End If
                            'Test if CBS
                            If gDateValue(slEFlightDate) < gDateValue(slSFlightDate) Then
                                ilIncludeFlight = False
                            End If
                            If ilIncludeFlight Then
                                mProcFlight ilCff, slSFlightDate, slEFlightDate, ilPass, slPctTrade, slSpotRate, slTotalNoPerWk, slTotalRate
                                If slTotalNoPerWk <> "0" Then
                                    smSave(3, ilRowNo) = slTotalNoPerWk
                                    smSave(4, ilRowNo) = slTotalRate
                                    If igPostType = 4 Then
                                        smSave(5, ilRowNo) = slTotalNoPerWk
                                        smSave(6, ilRowNo) = slTotalRate
                                        smSave(16, ilRowNo) = ""
                                    Else
                                        smSave(5, ilRowNo) = ""
                                        smSave(6, ilRowNo) = ""
                                        smSave(16, ilRowNo) = ""
                                    End If
                                    smSave(7, ilRowNo) = ""
                                    smSave(8, ilRowNo) = ""
                                    smSave(9, ilRowNo) = ""
                                    smSave(10, ilRowNo) = ""
                                    lmSave(3, ilRowNo) = tgCffRI(ilCff).CffRec.lActPrice
                                    smSave(11, ilRowNo) = ""    'Received Date
                                    smSave(12, ilRowNo) = ""    'Posted Date
                                    smSave(15, ilRowNo) = ""    'Original aired spots
                                    If Not ilAddTo Then
                                        For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                                            smInfo(ilCol, ilRowNo) = ""
                                        Next ilCol
                                        smInfo(1, ilRowNo) = "C"
                                        'ReDim Preserve smSave(1 To 16, 1 To ilRowNo + 1) As String
                                        'ReDim Preserve imSave(1 To 5, 1 To ilRowNo + 1) As Integer
                                        'ReDim Preserve lmSave(1 To 3, 1 To ilRowNo + 1) As Long
                                        'ReDim Preserve smInfo(1 To 11, 1 To ilRowNo + 1) As String * 12
                                        ReDim Preserve smSave(0 To 16, 0 To ilRowNo + 1) As String
                                        ReDim Preserve imSave(0 To 5, 0 To ilRowNo + 1) As Integer
                                        ReDim Preserve lmSave(0 To 3, 0 To ilRowNo + 1) As Long
                                        ReDim Preserve smInfo(0 To 11, 0 To ilRowNo + 1) As String * 12
                                        'imSave(1, ilRowNo + 1) = imSave(1, ilRowNo)
                                        'imSave(2, ilRowNo + 1) = imSave(2, ilRowNo)
                                        'lmSave(1, ilRowNo + 1) = lmSave(1, ilRowNo)
                                        'lmSave(2, ilRowNo + 1) = lmSave(2, ilRowNo)
                                        'lmSave(3, ilRowNo + 1) = lmSave(3, ilRowNo)
                                        'smSave(1, ilRowNo + 1) = smSave(1, ilRowNo)
                                        'smSave(2, ilRowNo + 1) = smSave(2, ilRowNo)
                                        For ilCol = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                                            smSave(ilCol, ilRowNo + 1) = smSave(ilCol, ilRowNo)
                                        Next ilCol
                                        For ilCol = LBound(imSave, 1) To UBound(imSave, 1) Step 1
                                            imSave(ilCol, ilRowNo + 1) = imSave(ilCol, ilRowNo)
                                        Next ilCol
                                        For ilCol = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
                                            lmSave(ilCol, ilRowNo + 1) = lmSave(ilCol, ilRowNo)
                                        Next ilCol
                                        For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                                            smInfo(ilCol, ilRowNo + 1) = smInfo(ilCol, ilRowNo)
                                        Next ilCol
                                        ilRowNo = ilRowNo + 1
                                    End If
                                End If
                            End If
                            ilCff = tgCffRI(ilCff).iNextCff
                        Loop

                    End If
                Next ilClf
                If ilStartRowNo < UBound(smSave, 2) Then
                    'Add Total line
                    slCTotalNoPerWk = "0"
                    slCTotalRate = "0"
                    For ilLoop = ilStartRowNo To UBound(smSave, 2) - 1 Step 1
                        slCTotalNoPerWk = gAddStr(slCTotalNoPerWk, smSave(3, ilLoop))
                        If InStr(RTrim$(smSave(4, ilLoop)), ".") > 0 Then
                            slCTotalRate = gAddStr(slCTotalRate, smSave(4, ilLoop))
                        End If
                    Next ilLoop
                    ilRowNo = UBound(smSave, 2)
                    For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                        smInfo(ilCol, ilRowNo) = ""
                    Next ilCol
                    smSave(1, ilRowNo) = ""
                    smSave(2, ilRowNo) = ""
                    smSave(3, ilRowNo) = slCTotalNoPerWk
                    smSave(4, ilRowNo) = slCTotalRate
                    If igPostType = 4 Then
                        smSave(5, ilRowNo) = slCTotalNoPerWk
                        smSave(6, ilRowNo) = slCTotalRate
                        smSave(16, ilRowNo) = ""
                    Else
                        smSave(5, ilRowNo) = ""
                        smSave(6, ilRowNo) = ""
                        smSave(16, ilRowNo) = ""
                    End If
                    smSave(7, ilRowNo) = ""
                    smSave(8, ilRowNo) = ""
                    smSave(9, ilRowNo) = ""
                    smSave(10, ilRowNo) = ""
                    imSave(2, ilRowNo) = 0
                    imSave(3, ilRowNo) = -1
                    If (igPostType = 4) Or (igPostType = 2) Then
                        slStr = "Total: " & Trim$(str$(tgChfRI.lCntrNo))
                        smInfo(1, ilRowNo) = "T"
                        'ReDim Preserve smSave(1 To 16, 1 To ilRowNo + 1) As String
                        'ReDim Preserve imSave(1 To 5, 1 To ilRowNo + 1) As Integer
                        'ReDim Preserve lmSave(1 To 3, 1 To ilRowNo + 1) As Long
                        'ReDim Preserve smInfo(1 To 11, 1 To ilRowNo + 1) As String * 12
                        ReDim Preserve smSave(0 To 16, 0 To ilRowNo + 1) As String
                        ReDim Preserve imSave(0 To 5, 0 To ilRowNo + 1) As Integer
                        ReDim Preserve lmSave(0 To 3, 0 To ilRowNo + 1) As Long
                        ReDim Preserve smInfo(0 To 11, 0 To ilRowNo + 1) As String * 12
                        For ilCol = LBound(smSave, 1) To UBound(smSave, 1) Step 1
                            smSave(ilCol, ilRowNo + 1) = ""
                        Next ilCol
                        For ilCol = LBound(imSave, 1) To UBound(imSave, 1) Step 1
                            imSave(ilCol, ilRowNo + 1) = 0
                        Next ilCol
                        For ilCol = LBound(lmSave, 1) To UBound(lmSave, 1) Step 1
                            lmSave(ilCol, ilRowNo + 1) = 0
                        Next ilCol
                        For ilCol = LBound(smInfo, 1) To UBound(smInfo, 1) Step 1
                            smInfo(ilCol, ilRowNo + 1) = ""
                        Next ilCol
                    End If
                End If
            End If
        Next ilPass
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mProcFlightRI                   *
'*                                                     *
'*             Created:8/02/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*                                                     *
'*******************************************************
Sub mProcFlight(ilCff As Integer, slSFlightDate As String, slEFlightDate As String, ilPass As Integer, slPctTrade As String, slSpotRate As String, slTotalNoPerWk As String, slTotalRate As String)
'   Where
'       ilCff(I)- Flight record index
'       slSFlightDate(I)- Flight Start date
'       slEFlightDate(I)- Flight End Date
'       slSpotRate(O)- Spot Rate
'       slTotalNoPerWk(O)- Running Total number of spots per week
'       slTotalRate(I/O)- Ordered Total $'s
'

    Dim llDate As Long
    Dim ilDay As Integer
    Dim llSDate As Long
    Dim slRate As String

    'Get flight rate
    Select Case tgCffRI(ilCff).CffRec.sPriceType
        Case "T"    'True
            slRate = gLongToStrDec(tgCffRI(ilCff).CffRec.lActPrice, 2)
            'Remove separate records for trade and cash
            'If (ilPass = 0) And (Val(slPctTrade) <> 0) Then
            '    slRate = gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100")
            'ElseIf (ilPass = 1) And (Val(slPctTrade) <> 100) Then
            '    slRate = gSubStr(RTrim$(slRate), gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100"))
            'End If
        Case "N"    'No Charge
            slRate = "N/C"
        Case "M"    'MG Line
            slRate = "MG"
        Case "B"    'Bonus
            slRate = "Bonus"
        Case "S"    'Spinoff
            slRate = "Spinoff"
        Case "P"    'Package
            slRate = gLongToStrDec(tgCffRI(ilCff).CffRec.lActPrice, 2)
            'Remove separate records for trade and cash
            'If (ilPass = 0) And (Val(slPctTrade) <> 0) Then
            '    slRate = gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100")
            'ElseIf (ilPass = 1) And (Val(slPctTrade) <> 100) Then
            '    slRate = gSubStr(RTrim$(slRate), gDivStr(gMulStr(RTrim$(slRate), gSubStr("100", slPctTrade)), "100"))
            'End If
        Case "R"    'Recapturable
            slRate = "Recapturable"
        Case "A"    'ADU
            slRate = "ADU"
    End Select
    'Ignore rates for all type except True and Package
    If (tgCffRI(ilCff).CffRec.sPriceType <> "T") And (tgCffRI(ilCff).CffRec.sPriceType <> "P") Then
        slRate = "0.00"
    End If
    slSpotRate = slRate
    If (tgCffRI(ilCff).CffRec.sDyWk <> "D") Then    'Weekly
        If tgChfRI.sBillCycle = "C" Then
            llDate = gDateValue(slSFlightDate)
            Do While llDate <= gDateValue(slEFlightDate)
                If llDate < lmStartCal Then
                    If llDate + 6 >= lmStartCal Then
                        slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffRI(ilCff).CffRec.iXSpotsWk)))
                        If InStr(RTrim$(slRate), ".") > 0 Then
                            slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffRI(ilCff).CffRec.iXSpotsWk))))
                        End If
                    End If
                ElseIf (llDate <= lmEndCal) Then
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffRI(ilCff).CffRec.iSpotsWk)))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffRI(ilCff).CffRec.iSpotsWk))))
                    End If
                Else
                    Exit Do
                End If
                llDate = gDateValue(gObtainNextMonday(Format$(llDate + 1, "m/d/yy")))
            Loop
        Else
            llDate = gDateValue(slSFlightDate)
            Do While llDate <= gDateValue(slEFlightDate)
                If (llDate >= lmStartStd) And (llDate <= lmEndStd) Then
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffRI(ilCff).CffRec.iSpotsWk)))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffRI(ilCff).CffRec.iSpotsWk))))
                    End If
                End If
                If llDate > lmEndStd Then
                    Exit Do
                End If
                llDate = gDateValue(gObtainNextMonday(Format$(llDate + 1, "m/d/yy")))
            Loop
        End If
    Else    'Daily
        If tgChfRI.sBillCycle = "C" Then
            If gDateValue(slSFlightDate) >= lmStartCal Then
                llSDate = gDateValue(slSFlightDate)
            Else
                llSDate = lmStartCal
            End If
            For llDate = llSDate To gDateValue(slEFlightDate) Step 1
                If (llDate >= lmStartCal) And (llDate <= lmEndCal) Then
                    ilDay = gWeekDayLong(llDate)
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffRI(ilCff).CffRec.iDay(ilDay))))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffRI(ilCff).CffRec.iDay(ilDay)))))
                    End If
                End If
                If llDate >= lmEndCal Then
                    Exit For
                End If
            Next llDate
        Else
            If gDateValue(slSFlightDate) >= lmStartStd Then
                llSDate = gDateValue(slSFlightDate)
            Else
                llSDate = lmStartStd
            End If
            For llDate = llSDate To gDateValue(slEFlightDate) Step 1
                If (llDate >= lmStartStd) And (llDate <= lmEndStd) Then
                    ilDay = gWeekDayLong(llDate)
                    slTotalNoPerWk = gAddStr(slTotalNoPerWk, Trim$(str$(tgCffRI(ilCff).CffRec.iDay(ilDay))))
                    If InStr(RTrim$(slRate), ".") > 0 Then
                        slTotalRate = gAddStr(slTotalRate, gMulStr(RTrim$(slRate), Trim$(str$(tgCffRI(ilCff).CffRec.iDay(ilDay)))))
                    End If
                End If
                If llDate >= lmEndStd Then
                    Exit For
                End If
            Next llDate
        End If
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadSbfRec                     *
'*                                                     *
'*             Created:9/04/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read item bill records         *
'*                                                     *
'*******************************************************
Function mReadSbfRecRI(ilTestForSbf As Integer) As Integer
'
'   iRet = mReadSbfRec
'   Where:
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilVeh As Integer
    Dim llDate As Long
    If (imMarketIndex >= 0) Then            '-1 if market doesnt exist
        imSbfRecLen = Len(tmSbf)
        tmSbfSrchKey2.sTranType = "T"
        'tmSbfSrchKey2.iDate(0) = 0
        'tmSbfSrchKey2.iDate(1) = 0
        gPackDate smStartStd, tmSbfSrchKey2.iDate(0), tmSbfSrchKey2.iDate(1)
        ilRet = btrGetGreaterOrEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmSbf.sTranType = "T")
            '6/24/10: Bypass Lines posted by Date/Time (tmSbf.iLineNo > 0)
            If tmSbf.iLineNo <= 0 Then
                gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
                If llDate > lmEndStd Then
                    Exit Do
                End If
                For ilVeh = 0 To UBound(tmMktVefCode) - 1 Step 1
                    If tmSbf.iAirVefCode = tmMktVefCode(ilVeh) Then
                        If ilTestForSbf Then
                            mReadSbfRecRI = True
                            Exit Function
                        End If
                        imChg = False
                        ilRet = mMergeRI()
                        Exit For
                    End If
                Next ilVeh
            End If
            ilRet = btrGetNext(hmSbf, tmSbf, imSbfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    If ilTestForSbf Then
        mReadSbfRecRI = False
    Else
        mReadSbfRecRI = True
    End If

    Exit Function

    On Error GoTo 0
    mReadSbfRecRI = False
    Exit Function
End Function

'
'       Cycle through all the active contracts and build arrays of
'       ordered $ and spots for the month
'
'       <input> ilListIndex = Index of report generating (INV_DELINQUENT or INV_UNBILLABLE)
Public Sub mBuildChf_Affs(Form As Form)
Dim ilRet As Integer    'Return Status
Dim ilLoop As Integer
Dim ilClf As Integer
Dim ilCff As Integer
Dim ilFound As Integer
Dim llUpper As Integer
Dim ilChf As Integer
Dim llChfCode As Long
Dim ilPass As Integer
Dim llKey As Long
Dim slTotalRate As String
Dim slTotalNoPerWk As String
Dim slPctTrade As String
Dim slSFlightDate As String
Dim slEFlightDate As String

Dim slStr As String
Dim ilIncludeFlight As Integer
Dim ilListIndex As Integer
Dim ilFoundAdv As Integer
Dim ilFoundVef As Integer
Dim llCntGross As Long
Dim llCntSpots As Long

    ilListIndex = Form!lbcRptType.ListIndex
    For ilChf = LBound(tmChfAdvtExt) To UBound(tmChfAdvtExt) - 1 Step 1
        'Filter advertiser selectivity for Unbillable report
        ilFoundAdv = False

        '1-13-03 Determine if this is an advertiser that is billed externally
        tmAdfSrchKey.iCode = tmChfAdvtExt(ilChf).iAdfCode
        ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE And tmAdf.sRepInvGen <> "E" Then   'only process this contract if valid adf read and its not set to externally bill
            If ilListIndex = INV_UNBILLABLE Then
                'For Unbillable report, see if selective advt requested
                If Form!ckcAll.Value = vbChecked Then
                    ilFoundAdv = True
                Else
                    For ilLoop = 0 To UBound(tmSelectiveList) - 1
                        If tmSelectiveList(ilLoop) = tmChfAdvtExt(ilChf).iAdfCode Then
                            ilFoundAdv = True
                            Exit For
                        End If
                    Next ilLoop
                End If
            Else
                ilFoundAdv = True
            End If
        End If
        If ilFoundAdv Then  'if selected go ahead and process the contract
            llChfCode = tmChfAdvtExt(ilChf).lCode
            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llChfCode, False, tgChfRI, tgClfRI(), tgCffRI())
            If ilRet Then
                ilPass = 0          'no need to split cash/trade portions, only 1 pass required.  this is
                                    'only to use common code for the flight subroutines
                slPctTrade = "0"     'no need to split cash/trade portions, assume all cash for sake of
                                    'calculating spots & $
                llCntGross = 0
                llCntSpots = 0
                For ilClf = LBound(tgClfRI) To UBound(tgClfRI) - 1 Step 1
                    tmClf = tgClfRI(ilClf).ClfRec

                    'Filter vehicle selectivity for Deqlinquent Affidavit report
                    ilFoundVef = False


                    If ilListIndex = INV_DELINQUENT Then
                        '1-14-03 dont test for "all vehicles". need tomake sure ony rep vehicles are included and
                        'the selective list to test contains only REP vehicles

                        'For Delinquent Affidavit report, see if selective vehicles requested
                        'If Form!ckcAll.Value = vbChecked Then
                        '    ilFoundVef = True
                        'Else
                            For ilLoop = LBound(tmSelectiveList) To UBound(tmSelectiveList) - 1
                                If tmSelectiveList(ilLoop) = tmClf.iVefCode Then
                                    ilFoundVef = True
                                    Exit For
                                End If
                            Next ilLoop
                        'End If
                    Else                'INV_UNBILLABLE, need to only process REP vehicles here, too
                        For ilLoop = LBound(tmREPVehicleList) To UBound(tmREPVehicleList) - 1
                            If tmREPVehicleList(ilLoop) = tmClf.iVefCode Then
                                ilFoundVef = True
                                Exit For
                            End If
                        Next ilLoop
                    End If
                    If ilFoundVef Then

                        If (tmClf.sType = "S") Or (tmClf.sType = "H") Then
                            llKey = 0
                            'If ilListIndex = INV_UNBILLABLE Then
                                llKey = tmClf.lChfCode
                            'End If
                            'Vehicle
                            slStr = "Missing"

                            slTotalNoPerWk = "0"
                            slTotalRate = "0"

                            ilCff = tgClfRI(ilClf).iFirstCff
                            Do While ilCff <> -1
                                gUnpackDate tgCffRI(ilCff).CffRec.iStartDate(0), tgCffRI(ilCff).CffRec.iStartDate(1), slSFlightDate
                                gUnpackDate tgCffRI(ilCff).CffRec.iEndDate(0), tgCffRI(ilCff).CffRec.iEndDate(1), slEFlightDate
                                ilIncludeFlight = True

                                If (gDateValue(slSFlightDate) > lmEndStd) Or (gDateValue(slEFlightDate) < lmStartStd) Then
                                    ilIncludeFlight = False
                                End If

                                'Test if CBS
                                If gDateValue(slEFlightDate) < gDateValue(slSFlightDate) Then
                                    ilIncludeFlight = False
                                End If
                                If ilIncludeFlight Then
                                    mProcFlightRI ilCff, slSFlightDate, slEFlightDate, ilPass, slPctTrade, slTotalNoPerWk, slTotalRate
                                End If
                                ilCff = tgCffRI(ilCff).iNextCff
                            Loop
                            'slTotalNoPerWk = total spots ordered for month
                            'slTotalRate = total $ ordered for month
                            If Val(slTotalNoPerWk) > 0 Then 'only create an entry if theres at least 1 spot, even if its zero $
                                ilFound = False
                                For ilLoop = LBound(tmAffByMonth) To UBound(tmAffByMonth) - 1
                                    If tmClf.iVefCode = tmAffByMonth(ilLoop).iVefCode And llKey = tmAffByMonth(ilLoop).lChfCode Then
                                        ilFound = True
                                        'accumulate the $ and spots
                                        tmAffByMonth(ilLoop).lOrderedSpots = tmAffByMonth(ilLoop).lOrderedSpots + Val(slTotalNoPerWk)
                                        tmAffByMonth(ilLoop).lOrderedAmt = tmAffByMonth(llUpper).lOrderedAmt + gStrDecToLong(slTotalRate, 2)
                                        Exit For
                                    End If
                                Next ilLoop
                                If Not ilFound Then         'create new entry
                                    llUpper = UBound(tmAffByMonth)
                                    tmAffByMonth(llUpper).iVefCode = tmClf.iVefCode
                                    tmAffByMonth(llUpper).lChfCode = llKey          'contr code or 0
                                    tmAffByMonth(llUpper).iAdfCode = tgChfRI.iAdfCode
                                    tmAffByMonth(llUpper).lOrderedSpots = Val(slTotalNoPerWk)
                                    tmAffByMonth(llUpper).lOrderedAmt = gStrDecToLong(slTotalRate, 2)
                                    ReDim Preserve tmAffByMonth(0 To llUpper + 1)
                                End If
                                'Accumulte the total $ & spots for each line
                                llCntGross = llCntGross + gStrDecToLong(slTotalRate, 2)
                                llCntSpots = llCntSpots + Val(slTotalNoPerWk)
                            End If
                        End If
                    End If          'ilFoundVef = true
                Next ilClf
                'Lines completed, now save the total ordered for this contract which needs to be
                'send along with each GRF record created for the Crystal report
                llUpper = UBound(tmAffOrdered)
                tmAffOrdered(llUpper).lChfCode = tgChfRI.lCode
                tmAffOrdered(llUpper).lCntGross = llCntGross
                tmAffOrdered(llUpper).lCntSpots = llCntSpots
                ReDim Preserve tmAffOrdered(0 To llUpper + 1) As AFFORDERED
            End If
        End If              'ilfoundadv = true
    Next ilChf
    Erase tmChfAdvtExt
End Sub
'
'
'       mBuildSbf_Affs - Read SBF and find all records that fall within the
'           standard broadcast month that is to be invoiced.  Build the aired
'           $ and spot counts into tmAffByMonth.  tmAFfByMonth was created based
'           on the active contracts for the month, gathering ordered information
'           If record found from sBF that has not been created from the contracts,
'           ignore it since the report is only concerned about ordered $ unbilled
'
'           8-30-02
Public Sub mBuildSbf_Affs(Form As Form)
Dim ilRet As Integer
Dim ilLoop As Integer
Dim ilLoopSBF As Integer
Dim llKey As Long
Dim ilListIndex As Integer
Dim ilFound As Integer
Dim llUpper As Integer
Dim tlSBFTypes As SBFTypes
ReDim tlSbf(0 To 0) As SBF

    ilListIndex = Form!lbcRptType.ListIndex
    'get all contracts between for SBF Import types only
    tlSBFTypes.iNTR = False
    tlSBFTypes.iInstallment = False
    tlSBFTypes.iImport = True
    ilRet = gObtainSBF(Form, hmSbf, 0, smStartStd, smEndStd, tlSBFTypes, tlSbf(), 0)    '11-28-06 add last parm to indicate which key to use

    For ilLoopSBF = LBound(tlSbf) To UBound(tlSbf)
        ilFound = False
        llKey = 0       'if by Delinquent Affidavits, dont need the contract breakout
        'If ilListIndex = INV_UNBILLABLE Then
            llKey = tlSbf(ilLoopSBF).lChfCode
        'End If
        For ilLoop = LBound(tmAffByMonth) To UBound(tmAffByMonth) - 1
            If tlSbf(ilLoopSBF).iAirVefCode = tmAffByMonth(ilLoop).iVefCode And llKey = tmAffByMonth(ilLoop).lChfCode Then
                ilFound = True
                'accumulate the $ and spots
                tmAffByMonth(ilLoop).lAiredSpots = tmAffByMonth(ilLoop).lAiredSpots + tlSbf(ilLoopSBF).iAirNoSpots + tlSbf(ilLoopSBF).iBonusNoSpots
                tmAffByMonth(ilLoop).lAiredAmt = tmAffByMonth(llUpper).lAiredAmt + tlSbf(ilLoopSBF).lGross
                tmAffByMonth(ilLoop).iShowFlag = 1     'sbf received, dont show on report
                If tlSbf(ilLoopSBF).sPostStatus = "R" Then
                    tmAffByMonth(ilLoop).iShowFlag = 2          'only received, not posted yet
                End If
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then         'Found a record for a vehicle that was not ordered for the
            'month.  Spot was moved to a different airing vehicle.  Ignore--only concerned with
            'ordered  comparisons
            'llUpper = UBound(tmAffByMonth)
            'tmAffByMonth(llUpper).iVefCode = tlSbf(ilLoopSBF).iAirVefCode
            'tmAffByMonth(llUpper).lChfCode = llKey      'contract code or 0
            'tmAffByMonth(llUpper).lAiredSpots = tlSbf(ilLoopSBF).iAirNoSpots + tlSbf(ilLoopSBF).iBonusNoSpots
            'tmAffByMonth(llUpper).lAiredAmt = tlSbf(ilLoopSBF).lGross
            'tmAffByMonth(llUpper).iShowFlag = 1     'sbf received, dont show on report
            'ReDim Preserve tmAffByMonth(0 To llUpper + 1)
        End If
    Next ilLoopSBF

    Erase tlSbf
End Sub
'
'           Generate the BArter Payments by week for a standard broadcast
'           month for all REP vehicles.  Include anything newly posted rcords
'           that fall within the requested standard broadcast month.
'           Use the Barter agreements by date assigned (VBF) to each vehicle (station).
'           11-30-06
'
'            gGenBarterpayments
'
Public Function gGenBarterPayment(Form As Form, hlExportFile As Integer) As Integer
'           Form - form coming from
'           hlExportFile - handle to csv file if Export Barter
'                          otherwise -1
'   8-24-10  Interface Barter Payment export as a CSV file with this report.
'           if there is an Exportfile handle, its an export; otherwise a report
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llStartWeek                   llTemp                        ilVefLoop                 *
'*  tlSBFTypes                    llStdStartDate                llStdEndDate              *
'*  slVehicleName                 ilYear                        ilCurrentMonth            *
'*  ilCalcLoop                                                                            *
'******************************************************************************************

Dim slKey As String
Dim slStr As String
Dim llStartStdBarter As Long        'last date barter paid
Dim slStartStdBarter As String
Dim llEndStdBarter As Long        'user entered posted thru date (defaulted to today)
Dim slEndStdBarter As String
Dim slYear As String
Dim ilRet As Integer
Dim llSBFLoop As Long
Dim llUpper As Long
Dim llLower As Long
Dim slMonthHdr As String * 36
Dim llWeekToBill As Long
Dim ilWeekInx As Integer
Dim ilIncludeCodes As Integer       'true = include the vehicle code, false = exclude the vehicle code
Dim ilValid As Integer
Dim llLoIndex As Long
Dim llHiIndex As Long
Dim ilVefCode As Integer
Dim ilInclPreviousPaid As Integer
Dim ilUpdatePaid As Integer
ReDim ilUseCodes(0 To 0) As Integer
ReDim tlSbf(0 To 0) As SBF
ReDim tlSbfTemp(0 To 0) As SBF
ReDim tlVbf(0 To 0) As VBF
Dim tlSBFType As SBFTypes
Dim tlBarterStats As BARTERSTATS
Dim slStdStartYear As String        'Start date of std bdcst year
Dim llSTdStartYear As Long
Dim ilLoopOnVef As Integer
Dim slTempStart As String
Dim slTempEnd As String
Dim llTempStart As Long
Dim llTempEnd As Long
Dim ilVbfInx As Integer
Dim slNameCode As String
Dim slCode As String
Dim ilWkIncrement As Integer
Dim ilCurrentVbfInx As Integer
Dim ilLoopOnVbf As Integer
Dim ilLoop As Integer
Dim blFound As Boolean
Dim ilGetSbfStartDate(0 To 1) As Integer        'start date gathering sbf records
Dim ilGetSbfEndDate(0 To 1) As Integer        'end date gathering sbf records
Dim ilVefList() As Integer
Dim ilLoopOnFile As Integer                 '10-19-18 loop on rvf , then phf to find assoc inv #
Dim hlWhichFile As Integer
Dim tlRvfSrchKey4 As RVFKEY4
Dim blFoundTrans As Boolean


'GRF field definitions
'
'grfchfCode        Contract code
'grfSTartDate      Week of
'grfYear           Year requested (user input)
'grfCode2          Month requested (user input)
'grfvefcode        vehicle code (station)
'grfDate           Posted date (date the user posted the advt or station)
'                   SBF is searched for all posted dates falling in the standard month requested
'grfPer1Genl       spot length (for detail)
'grfPer2Genl       Ordered # spots
'grfPer3Genl       Aired spots in units
'grfPer4Genl       # free spots (calculated)
'grfPer5Genl       # payable spots (calculated)
'grfPer6Genl       Type:  1 = header record for barter specifications
'                         11 = detail
'grfPer7Genl       spot length for sorting (0 for all methods except pay x for y spots and its per period is month or year)
'grfPer9Genl       Spot Length #1  (from barter table)
'grfPer10Genl      spot Length #2
'grfPer11Genl -
'  grfPer18Genl     spot lengths #3 - #10
'grfPer1           Aired Minutes stored in seconds
'grfPer2           cost for detail record
'grfper3           VBF pointer
'grfPer9 -
'  grfPer18        Actual acquistion cost by spot length (corresponds to grfPer9Genl - grfPer18Genl)
'grfBktType         Barter Paid (Y/N)
'grfDateGenl(1)     start date of year gathering
'grfDateGenl(2)     end date of year gathering


    gGenBarterPayment = -1           'assume all ok
    mOpenBarterFiles Form
    If imTerminate = True Then          'open errors?
        gGenBarterPayment = 0           'file open error
        Exit Function
    End If
    ReDim lgPrintedCnts(0 To 0) As Long     'init the array of SBF records to update if setting the PAID flag
    slMonthHdr = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"

    ReDim ilVefList(0 To 0) As Integer
    If Form.Name = "ExpGPBarter" Then            'Export barter payment
        hmVaf = CBtrvTable(ONEHANDLE) 'CBtrvObj
        ilRet = btrOpen(hmVaf, "", sgDBPath & "Vaf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo mInitErr
        gBtrvErrorMsg ilRet, "gGenbarterPayment (btrOpen VAF)", Form
        On Error GoTo 0
        gPopVaf hmVaf, tmVaf()
        'build list of vehicle codes to include or exclude for faster processing
        gObtainCodes Form!lbcSelection, tgCSVNameCode(), ilIncludeCodes, ilUseCodes(), Form
        'determine the start and end dates of the month requested based on month/year
        slStr = Form!edcSelCFrom.Text             'month in text form (jan..dec)
        imYear = Val(Form!edcSelCFrom1.Text)
        slYear = Trim$(Form!edcSelCFrom1.Text)
        ilInclPreviousPaid = False          'exclude previous barters paid
        ilUpdatePaid = False                'do not udpate the barter paid
        For ilLoopOnVef = 0 To Form!lbcSelection.ListCount - 1
            If Form!lbcSelection.Selected(ilLoopOnVef) Then
                slNameCode = tgCSVNameCode(ilLoopOnVef).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)          'vehicle code
                ilVefList(UBound(ilVefList)) = Val(slCode)
                ReDim Preserve ilVefList(0 To UBound(ilVefList) + 1) As Integer
            End If
        Next ilLoopOnVef

    Else                                'barter payment report
        'build list of vehicle codes to include or exclude for faster processing
        gObtainCodesForMultipleLists 1, tgCSVNameCode(), ilIncludeCodes, ilUseCodes(), Form
        'determine the start and end dates of the month requested based on month/year
        slStr = Form!edcSelCFrom.Text             'month in text form (jan..dec)
        imYear = Val(Form!edcSelCFrom1.Text)
        slYear = Trim$(Form!edcSelCFrom1.Text)
        ilInclPreviousPaid = False          'exclude previous barters paid
        If Form!ckcInclPreviousPaid.Value = vbChecked Then
            ilInclPreviousPaid = True       'include previous barters paid
        End If
        ilUpdatePaid = False                'do not udpate the barter paid
        If Form!ckcInclRecd.Value = vbChecked Then
            ilUpdatePaid = True             'update barter paid
        End If
        For ilLoopOnVef = 0 To Form!lbcSelection(1).ListCount - 1
            If Form!lbcSelection(1).Selected(ilLoopOnVef) Then
                slNameCode = tgCSVNameCode(ilLoopOnVef).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)          'vehicle code
                ilVefList(UBound(ilVefList)) = Val(slCode)
                ReDim Preserve ilVefList(0 To UBound(ilVefList) + 1) As Integer
            End If
        Next ilLoopOnVef
    End If

    gGetMonthNoFromString slStr, imCurrentMonth         'getmonth #
    If imCurrentMonth = 0 Then                                 'input isn't text month name, try month #
        imCurrentMonth = Val(slStr)
    End If
    slStr = Trim$(str(imCurrentMonth)) & "/15/" & slYear     'form mm/dd/yy
    slStartStdBarter = gObtainStartStd(slStr)               'obtain std start date for month
    llStartStdBarter = gDateValue(slStartStdBarter)
    slEndStdBarter = gObtainEndStd(slStr)                 'obtain std end date for month
    llEndStdBarter = gDateValue(slEndStdBarter)
    
    slStr = "1/15/" & slYear     'form mm/dd/yy
    slStdStartYear = gObtainStartStd(slStr)               'obtain start date for std bdcst year
    llSTdStartYear = gDateValue(slStdStartYear)

    ReDim tmBarterTypeSort(0 To 0) As BARTERTYPESORT
    ReDim tmYearly_Tracking(0 To 0) As YEARLY_TRACKING        'if any yearly barters for Exceeding Minutes or Units, keep track and update vbf
    llUpper = 0
    'retrieve all rep billing only for transactions posted between the requested start/end dates of the std bdcst month requested
    tlSBFType.iNTR = False           'no NTRs
    tlSBFType.iInstallment = False   'no installments
    tlSBFType.iImport = True         'only rep posting

'    For ilLoopOnVef = 0 To Form!lbcSelection(1).ListCount - 1 Step 1
     For ilLoopOnVef = 0 To UBound(ilVefList) - 1
'        If Form!lbcSelection(1).Selected(ilLoopOnVef) Then
'            slNameCode = tgCSVNameCode(ilLoopOnVef).sKey
'            ilRet = gParseItem(slNameCode, 2, "\", slCode)          'vehicle code
'            ilVefCode = Val(slCode)
            ilVefCode = ilVefList(ilLoopOnVef)

            'check vehicle to detrmine its barter interval, need to change dates
            'if year interval, use std bdct start year thru end of the month requested
            
            'determine the barter agreement. Find the one whose start date is prior/equal to the bdcst date requested
            gPackDateLong llStartStdBarter, tmVbfSrchKey1.iStartDate(0), tmVbfSrchKey1.iStartDate(1)
            tmVbfSrchKey1.iVefCode = ilVefCode
            ilRet = btrGetLessOrEqual(hmVbf, tmVbf, imVBfRecLen, tmVbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE And tmVbf.iVefCode = ilVefCode) Then
                'convert dates to test if this is the proper agreement
                'determine per period intervals to use
                If ((tmVbf.sPerPeriod = "W" Or tmVbf.sPerPeriod = " ") And (Form!ckcMethod(0).Value = vbChecked)) Or (tmVbf.sPerPeriod = "Y" And Form!ckcMethod(2).Value = vbChecked) Or (tmVbf.sPerPeriod = "M" And Form!ckcMethod(1).Value = vbChecked) Then
                    'ok to include exports if weekly (monthly are totalled by spot length and do not have contract totals, so they must be ignored)
                    'If ((Form.Name = "ExpGPBarter") And (tmVbf.sPerPeriod <> "M")) Or (Form.Name = "RptSelRI") Then
                        gUnpackDateLong tmVbf.iStartDate(0), tmVbf.iStartDate(1), llTempStart
                        gUnpackDateLong tmVbf.iEndDate(0), tmVbf.iEndDate(1), llTempEnd
                        If llStartStdBarter >= llTempStart Then        'latest agreement falls within the requested period
                            'determine interval of barter (week/month uses std month, year uses std bdcst year)
                            If tmVbf.sPerPeriod = "Y" Then          'year interval
                                If llTempStart < llSTdStartYear Then        'agreement start date is prior to the start of year to process, use the start of the current year
                                    slTempStart = slStdStartYear
                                    slTempEnd = slEndStdBarter
                                Else                                        'agreement starts after the start of the std year, use the agreement start date
                                    slTempStart = Format(llTempStart, "m/d/yy")
                                    slTempEnd = slEndStdBarter
                                End If
                            Else            'week or month barter
                                slTempStart = slStartStdBarter
                                slTempEnd = slEndStdBarter
                            End If
                            tlVbf(UBound(tlVbf)) = tmVbf
                            ReDim Preserve tlVbf(0 To UBound(tlVbf) + 1) As VBF
                            'need new routine to gather by vehicle, tran type & postdate
                            ilRet = gObtainSBF(Form, hmSbf, 0, slTempStart, slTempEnd, tlSBFType, tlSbfTemp(), 4, ilVefCode)   'obtain sbf by vehicle, tran type & post date
                            llUpper = UBound(tlSbf)
                            For llSBFLoop = LBound(tlSbfTemp) To UBound(tlSbfTemp) - 1            'append the vehicles posting for the month or year
                                tlSbf(llUpper) = tlSbfTemp(llSBFLoop)
                                ReDim Preserve tlSbf(LBound(tlSbf) To llUpper + 1) As SBF
                                llUpper = llUpper + 1
                            Next llSBFLoop
                        End If
                   ' End If      'ignore export for monthly or year
                End If          'tmvbf.sperPeriod test
            End If
'        End If
    Next ilLoopOnVef
    
    llUpper = 0
    gPackDate slTempStart, ilGetSbfStartDate(0), ilGetSbfStartDate(1)
    gPackDate slTempEnd, ilGetSbfEndDate(0), ilGetSbfEndDate(1)
    For llSBFLoop = LBound(tlSbf) To UBound(tlSbf) - 1
        tmSbf = tlSbf(llSBFLoop)
        ilValid = gTestIncludeExclude(tmSbf.iBillVefCode, ilIncludeCodes, ilUseCodes())
        'If ilValid And tmSbf.sBarterPaid <> "Y" Then        'valid vehicle and the station hasnt been paid yet
        'Posting included if:
        '     1)  Updating and the posting has not been paid,
        '     2)  Not updating and Incl Previous Paid = Yes, or
        '     3)  Not updating and Incl Previous paid = No and Posted BarterPaid flag  = No
        'If (ilValid) And ((Form!ckcInclRecd.Value = vbChecked And tmSbf.sBarterPaid <> "Y") Or (Form!ckcInclRecd.Value = vbUnchecked And Form!ckcInclPreviousPaid.Value = vbChecked) Or (Form!ckcInclRecd.Value = vbUnchecked And Form!ckcInclPreviousPaid.Value = vbUnchecked And tmSbf.sBarterPaid = "N")) Then
        If (ilValid) And (((ilUpdatePaid = True) And (tmSbf.sBarterPaid <> "Y")) Or ((ilUpdatePaid = False) And (ilInclPreviousPaid = True)) Or ((ilUpdatePaid = False) And (ilInclPreviousPaid = False) And (tmSbf.sBarterPaid = "N"))) Then
            tmChfSrchKey.lCode = tmSbf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
            If ilRet <> BTRV_ERR_NONE Then
                tmChf.lCntrNo = 999999999       'none found
            End If
            
            ilCurrentVbfInx = -1
            For ilLoopOnVbf = LBound(tlVbf) To UBound(tlVbf) - 1
                If tlVbf(ilLoopOnVbf).iVefCode = tmSbf.iBillVefCode Then
                    ilCurrentVbfInx = ilLoopOnVbf
                    Exit For
                End If
            Next ilLoopOnVbf
            
            If ilCurrentVbfInx >= 0 Then
                tmVbf = tlVbf(ilCurrentVbfInx)
                
                hlWhichFile = hmRvf
' 10/30/18 no need to access RVF/PHF
'                For ilLoopOnFile = 1 To 2
'                    blFoundTrans = False
'                    'find matching receivables to get the inv #
'                    tlRvfSrchKey4.iTranDate(0) = tmSbf.iDate(0)
'                    tlRvfSrchKey4.iTranDate(1) = tmSbf.iDate(1)
'                    tlRvfSrchKey4.lCntrNo = tmChf.lCntrNo
'                    ilRet = btrGetGreaterOrEqual(hlWhichFile, tmRvf, imRvfRecLen, tlRvfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)
'                    'keep reading until a matching contr # found for the billing period, and an "IN" transaction
'                    Do While (ilRet = BTRV_ERR_NONE) And (tmRvf.iTranDate(0) = tmSbf.iDate(0) And tmRvf.iTranDate(1) = tmSbf.iDate(1) And tmRvf.lCntrNo = tmChf.lCntrNo)
'                        If tmRvf.sTranType = "IN" And tmRvf.iAirVefCode = tmSbf.iAirVefCode Then
'                            blFoundTrans = True
'                            Exit Do
'                        Else
'                            ilRet = btrGetNext(hmRvf, tmRvf, imRvfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
'                        End If
'                    Loop
'                    If blFoundTrans Then
'                        ilLoopOnFile = 2
'                    Else
'                        If ilLoopOnFile = 1 Then                          'if 1, then just finished receivables, go do history
'                           hlWhichFile = hmPhf
'                        End If
'                    End If
'                Next ilLoopOnFile
               
'                If blFoundTrans Then
'                    tlBarterStats.lInvNo = tmRvf.lInvNo
'                Else
'                    tlBarterStats.lInvNo = 0
'                End If

                tlBarterStats.lInvNo = tmSbf.lRefInvNo              '10-30-18 this is the station inv # entered when posting the rep spots
                tlBarterStats.lVbfcode = tmVbf.lCode
                tlBarterStats.sMethod = tmVbf.sMethod       '10-19-18
                tlBarterStats.sPeriodForMethodX = tmVbf.sPerPeriod   '10-19-18 used for method X , period interval Month, for export option
                tlBarterStats.lCntrNo = tmChf.lCntrNo       'contract was accessed in mBuildBarterKey
                tlBarterStats.iLen = tmSbf.iSpotLen
                
                tlBarterStats.lChfCode = tmSbf.lChfCode
                tlBarterStats.iAdfCode = tmChf.iAdfCode
                tlBarterStats.sProduct = Trim$(tmChf.sProduct)
                tlBarterStats.sCashTrade = "C"                  'assume cash
                If tmChf.iPctTrade > 0 Then
                    tlBarterStats.sCashTrade = "T"
                End If
                tlBarterStats.iVefCode = tmSbf.iBillVefCode
                tlBarterStats.sBarterPaid = tmSbf.sBarterPaid
                tlBarterStats.iPostDate(0) = tmSbf.iPostDate(0)
                tlBarterStats.iPostDate(1) = tmSbf.iPostDate(1)
                tlBarterStats.iGetSbfStartDate(0) = ilGetSbfStartDate(0)    'start and end dates gathered to show on report for year agreements
                tlBarterStats.iGetSbfStartDate(1) = ilGetSbfStartDate(1)
                tlBarterStats.iGetSbfEndDate(0) = ilGetSbfEndDate(0)
                tlBarterStats.iGetSbfEndDate(1) = ilGetSbfEndDate(1)
                tlBarterStats.bIgnoreForXport = False               'see description in type statement above
                
                'get the start of the standard bdcst month that has been posted
                gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llWeekToBill
                slStr = Format(llWeekToBill, "m/d/yy")
                slStr = gObtainStartStd(slStr)
                llWeekToBill = gDateValue(slStr)
                ilWkIncrement = 7
                tlBarterStats.iLenForSortMethodX = 0
                If tmVbf.sPerPeriod = "M" Or tmVbf.sPerPeriod = "Y" Then
                    ilWkIncrement = 0          'do not separate the weeks
                    'override the start date so all entries will have the same date.  Do not want to split out dates like the weekly version does
                    llWeekToBill = llStartStdBarter
                    If tmVbf.sMethod = "X" Then
                        tlBarterStats.iLenForSortMethodX = tmSbf.iSpotLen    'payment for month or year by spot length for method X
                    End If
                End If
                tlBarterStats.iAiredSpots = 0
                tlBarterStats.lMinAired = 0
                tlBarterStats.iOrderedSpots = 0
                tlBarterStats.iFree = 0
                tlBarterStats.iPayable = 0
                'see if any weekly posting exists.  If so, build a sort record for each week with its aired/ordered info
                tmRwfSrchKey.lCode = tmSbf.lCode
                ilRet = btrGetEqual(hmRwf, tmRwf, imRwfRecLen, tmRwfSrchKey, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                If ilRet = BTRV_ERR_NONE And tmRwf.lSbfCode = tmSbf.lCode Then     'only process this contract if valid adf read and its not set to externally bill
                    For ilWeekInx = 1 To 5
                        'dont create an entry if there are no aired spots for the week
                        If (tmRwf.iWkNoSpots(ilWeekInx - 1) + tmRwf.iWkNoCarried(ilWeekInx - 1)) <> 0 Then 'only build an entry if theres data for the week
                            tlBarterStats.lWeek = llWeekToBill            'billing week that has been posted
                            tlBarterStats.iAiredSpots = tmRwf.iWkNoSpots(ilWeekInx - 1) + tmRwf.iWkNoCarried(ilWeekInx - 1) 'no spots aired plus makegood in same week
                            tlBarterStats.lMinAired = (tlBarterStats.iLen * (tmRwf.iWkNoSpots(ilWeekInx - 1) + tmRwf.iWkNoCarried(ilWeekInx - 1)))   'calc minutes aired (store in seconds)
                            tlBarterStats.iOrderedSpots = tmRwf.iWkOrderSpotNo(ilWeekInx - 1)
                            slKey = mBuildBarterKey(llWeekToBill)
                            If ilWkIncrement > 0 Then           'increment by 7 if by week; otherwise keep all dates together so as not to sort separately in report
                                tmBarterTypeSort(llUpper).sKey = slKey
                                tmBarterTypeSort(llUpper).tBarterStats = tlBarterStats
                                llUpper = llUpper + 1
                                ReDim Preserve tmBarterTypeSort(0 To llUpper) As BARTERTYPESORT
                            Else
                                blFound = False
                                For ilLoop = LBound(tmBarterTypeSort) To UBound(tmBarterTypeSort) - 1
                                    If Trim$(slKey) = Trim$(tmBarterTypeSort(ilLoop).sKey) Then
                                        blFound = True
                                        Exit For
                                    End If
                                Next ilLoop
                                If blFound Then
                                    'increment info into already created entry within ilLoop
                                    tmBarterTypeSort(ilLoop).tBarterStats.iAiredSpots = tmBarterTypeSort(ilLoop).tBarterStats.iAiredSpots + tlBarterStats.iAiredSpots
                                    tmBarterTypeSort(ilLoop).tBarterStats.lMinAired = tmBarterTypeSort(ilLoop).tBarterStats.lMinAired + tlBarterStats.lMinAired
                                    tmBarterTypeSort(ilLoop).tBarterStats.iOrderedSpots = tmBarterTypeSort(ilLoop).tBarterStats.iOrderedSpots + tlBarterStats.iOrderedSpots
                                Else
                                    tmBarterTypeSort(llUpper).sKey = slKey
                                    tmBarterTypeSort(llUpper).tBarterStats = tlBarterStats
                                    llUpper = llUpper + 1
                                    ReDim Preserve tmBarterTypeSort(0 To llUpper) As BARTERTYPESORT
                                End If
                            End If
                        End If
                        llWeekToBill = llWeekToBill + ilWkIncrement     'increment by 1 week for week per period only for sorting.  if other than week per period, need to keep together without sorting apart
                    Next ilWeekInx

                Else                            'no weekly posting done
                    tlSbf(llSBFLoop).lCode = -tlSbf(llSBFLoop).lCode          'ignore the negated codes if updating paid status
                End If
            End If              ' ilCurrentVbfInx >= 0
        Else                    'not a valid posted record or has already been paid,
                                'set the record # so it doesnt get updated
            tlSbf(llSBFLoop).lCode = -tlSbf(llSBFLoop).lCode          'ignore the negated codes if updating paid status
        End If                  'if ilvalid
    Next llSBFLoop

    If UBound(tmBarterTypeSort) - 1 > 0 Then
        ArraySortTyp fnAV(tmBarterTypeSort(), 0), UBound(tmBarterTypeSort), 0, LenB(tmBarterTypeSort(0)), 0, LenB(tmBarterTypeSort(0).sKey), 0
    End If

    'Array has been created and sorted by vehicle.  Loop thru one vehicle at a time and calculate their barter payments.
    'Get the barter agreement associated with each vehicle based on the std month requested.
    llLoIndex = -1
    llHiIndex = -1
    llLower = 0

    'determine the span of indexes to process for 1 vehicle
    For llSBFLoop = llLower To llUpper
        If llLoIndex >= 0 Then       'not first time thru
             If tmBarterTypeSort(llSBFLoop).tBarterStats.iVefCode = tmBarterTypeSort(llLoIndex).tBarterStats.iVefCode Then
                llHiIndex = llSBFLoop
                ilVefCode = tmBarterTypeSort(llHiIndex).tBarterStats.iVefCode
            Else            'different vehicles
                'process the vehicle for the span of indices obtained.  Calculate the barter payments due
            
                mCalcBarterDue ilVefCode, llStartStdBarter, llLoIndex, llHiIndex, tlVbf()
                If Form.Name = "RptSelRI" Then
                    mWriteBarterDue llLoIndex, llHiIndex
                Else
                    mWriteBarterExport llLoIndex, llHiIndex, hlExportFile
                End If
                'go get the next set of vehicles to process
                llLoIndex = llHiIndex + 1
                llHiIndex = llLoIndex
                'llLower = llLoIndex
                ilVefCode = tmBarterTypeSort(llHiIndex).tBarterStats.iVefCode
            End If
        Else
            llLoIndex = llSBFLoop
            llHiIndex = llSBFLoop
            ilVefCode = tmBarterTypeSort(llHiIndex).tBarterStats.iVefCode
       End If
    Next llSBFLoop

    'if  updating the barter paid, go thru the SBF array create a global array for those
    'SBF records to update.  Update after the report has been generated
    'llUpper = 1
    llUpper = 0
    'If Form!ckcInclRecd.Value = vbChecked Then          'update the paid flags
    If ilUpdatePaid = True Then          'update the paid flags
        For llSBFLoop = LBound(tlSbf) To UBound(tlSbf) - 1
            tmSbf = tlSbf(llSBFLoop)
            If tmSbf.lCode > 0 Then
                lgPrintedCnts(llUpper) = tmSbf.lCode
                llUpper = llUpper + 1
                ReDim Preserve lgPrintedCnts(0 To llUpper) As Long
            End If
        Next llSBFLoop
    End If

    Erase tmBarterTypeSort, tlSbf, ilUseCodes
    Erase tlSbfTemp, ilVefList
    Erase tlVbf, tmYearly_Tracking
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmRwf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmVbf)
    ilRet = btrClose(hmRvf)
    ilRet = btrClose(hmPhf)
    ilRet = btrClose(hmVaf)

    btrDestroy hmVsf
    btrDestroy hmRwf
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmSbf
    btrDestroy hmVbf
    btrDestroy hmRvf
    btrDestroy hmPhf
    btrDestroy hmVaf
    Exit Function
    
mInitErr:
    On Error GoTo 0
    imTerminate = True
    gGenBarterPayment = 0
    Exit Function
End Function
'
'           mBuildBarterKey - build key for sorting the rep posting
'           sort major to minor:  vehicle,  week of, contract # (descending), spot length
'           <input> none
'           <return> key string
Public Function mBuildBarterKey(llWeekToBill As Long) As String
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilDay                         ilRet                                                   *
'******************************************************************************************

Dim ilVefLoop As Integer
Dim slVehicleName As String
Dim slStr As String
Dim slKey As String
Dim llTemp As Long

        ilVefLoop = gBinarySearchVef(tmSbf.iBillVefCode)
        If ilVefLoop = -1 Then
            slVehicleName = ""
        Else
            slVehicleName = tgMVef(ilVefLoop).sName
        End If
        'format key for the sort:  vehicle name (max 40 char; contract # descending (9), date as number (5), spot length (3)
        slStr = Trim$(slVehicleName)           'vehicle name
        Do While Len(slStr) < 40
            slStr = slStr & " "
        Loop
        slKey = slKey & slStr & "|"

        'start week
        slStr = str$(llWeekToBill)
        Do While Len(slStr) < 5
            slStr = "0" & slStr
        Loop
        slKey = slKey & slStr & "|"

        If ((tmVbf.sMethod = "X") And (tmVbf.sPerPeriod = "M" Or tmVbf.sPerPeriod = "Y")) Then      'pay x for every Y by month or year, sort spot length first
            
            slStr = Trim$(str$(tmSbf.iSpotLen))            'spot length
            Do While Len(slStr) < 3
                slStr = "0" & slStr
            Loop
            slKey = slKey & slStr & "|"

            'get the contract
            'llTemp = 999999999 - tmChf.lCntrNo
            '1-21-10 Client wants to give newest contract the barter $, not the oldest
            llTemp = tmChf.lCntrNo
            slStr = Trim$(str$(llTemp))          'contract #
            Do While Len(slStr) < 9
                slStr = "0" & slStr
            Loop
    
        Else
            'get the contract
            'llTemp = 999999999 - tmChf.lCntrNo
            '1-21-10 Client wants to give newest contract the barter $, not the oldest
            llTemp = tmChf.lCntrNo
            slStr = Trim$(str$(llTemp))          'contract #
            Do While Len(slStr) < 9
                slStr = "0" & slStr
            Loop
            slKey = slKey & slStr & "|"
    
            slStr = Trim$(str$(tmSbf.iSpotLen))            'spot length
            Do While Len(slStr) < 3
                slStr = "0" & slStr
            Loop
        End If
        
        slKey = slKey & slStr
        mBuildBarterKey = slKey
End Function
'
'           mCalcBarterDue - calculate the payment due for the station based on
'           its barter agreement
'           <input> ilLoIndex
'                   ilHiIndex
Public Sub mCalcBarterDue(ilVefCode As Integer, llStdStart As Long, llLoIndex As Long, llHiIndex As Long, tlVbf() As VBF)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilStdStart                                                                            *
'******************************************************************************************

Dim llTempStart As Long
Dim llTempEnd As Long
Dim ilRet As Integer
Dim llBarterLoop As Long
Dim ilAiredUnits As Integer
Dim llAiredSec As Long
Dim ilThreshHoldMet As Integer
Dim ilfirstTime As Integer
Dim llPrevWeek As Long
Dim ilGroups As Integer
Dim ilTemp As Integer
Dim llTemp As Long
Dim ilTempXFree As Integer      '7-8-15
Dim ilTempYSold As Integer
Dim ilCurrentVbfInx As Integer
Dim ilLoopOnVbf As Integer
'Dim ilSpotCountByLen(1 To 10) As Integer
Dim ilSpotCountByLen(0 To 10) As Integer    'Index zero ignored
Dim blCalcMethodX As Boolean
Dim ilPayable As Integer
Dim ilFree As Integer
Dim llBarterValue As Long
Dim llUsedItems As Long     'count of units or minutes used for bart methods Pay when Minutes or units exceeded
Dim blFoundTrack As Boolean
Dim blFirstForThisSpotLength As Boolean 'for barter method X only period interval month and year because they are totalled
                            'for all contracts.  Need to put the total on any contract with matching spot length for export
Dim ilTotalAiredByLen As Integer    '10-19-18 count of total spots aired by spot length
Dim ilPct As Integer
Dim llRevShare As Long
Dim tlStats As BARTERSTATS
Dim slStr As String
Dim slDollar As String
Dim slPct As String
Dim slAirCount As String
Dim llAccumRev As Long

        ilCurrentVbfInx = -1
        For ilLoopOnVbf = LBound(tlVbf) To UBound(tlVbf) - 1
            If tlVbf(ilLoopOnVbf).iVefCode = ilVefCode Then
                ilCurrentVbfInx = ilLoopOnVbf
                Exit For
            End If
        Next ilLoopOnVbf
        If ilCurrentVbfInx < 0 Then
            Exit Sub
        End If
        
        tmVbf = tlVbf(ilCurrentVbfInx)
'        'determine the barter agreement. Find the one whose start date is prior/equal to the bdcst date requested
'        gPackDateLong llStdStart, tmVbfSrchKey1.iStartDate(0), tmVbfSrchKey1.iStartDate(1)
'        tmVbfSrchKey1.iVefCode = ilVefCode
'        ilRet = btrGetLessOrEqual(hmVbf, tmVbf, imVBfRecLen, tmVbfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
'        If (ilRet = BTRV_ERR_NONE And tmVbf.iVefCode = ilVefCode) Then
            'convert dates to test if this is the proper agreement
            gUnpackDateLong tmVbf.iStartDate(0), tmVbf.iStartDate(1), llTempStart
            gUnpackDateLong tmVbf.iEndDate(0), tmVbf.iEndDate(1), llTempEnd
            If llStdStart < llTempStart Then        'the agreement starts after the std month requested
                Exit Sub
            End If
'        Else
'            tmVbf.sMethod = ""
'            For ilTemp = 1 To 10
'                tmVbf.iSpotLen(ilTemp) = 0
'                tmVbf.lActAcqCost(ilTemp) = 0
'            Next ilTemp
'            Exit Sub                'dont process this vehicle
'        End If

        'accumulate by spot length if by method "X" (pay X for every Y spot)
        
        ilfirstTime = True
        For ilTemp = 1 To 10
            ilSpotCountByLen(ilTemp) = 0
        Next ilTemp
        
        blCalcMethodX = False
        llUsedItems = 0
        For llBarterLoop = llLoIndex To llHiIndex
            tlStats = tmBarterTypeSort(llBarterLoop).tBarterStats      'place into single record for easier reading
            'Computations start over (buckets initialized) for each new week
            If ilfirstTime Then
                llPrevWeek = tlStats.lWeek
                ilfirstTime = False
                ilAiredUnits = 0
                llAiredSec = 0
                ilThreshHoldMet = False
            End If
            If llPrevWeek <> tlStats.lWeek Then
                llPrevWeek = tlStats.lWeek
                ilAiredUnits = 0
                llAiredSec = 0
                ilThreshHoldMet = False
            End If
            tlStats.lBarterValue = mGetBarterValueForLen(tlStats.iLen)      'get the actual $ that should be paid to station
            If tmVbf.sMethod = "A" Then     'all cash.  Everything is payable
                tlStats.iPayable = tlStats.iAiredSpots                          'total spots that are payable to station
                tlStats.iFree = 0
                tlStats.lMinAired = tlStats.iAiredSpots * tlStats.iLen           'total minutes sold stored in seconds
                tmBarterTypeSort(llBarterLoop).tBarterStats = tlStats      'place the updated image backinto array

            ElseIf tmVbf.sMethod = "M" Then   'minutes exceeded.  i.e. Payable after 15M exceeded
                llUsedItems = llUsedItems + (tlStats.iAiredSpots * tlStats.iLen)        'accumulate for the entire period of time
              
                llAiredSec = llAiredSec + (tlStats.iAiredSpots * tlStats.iLen)
                If ilThreshHoldMet Then
                    tlStats.iPayable = tlStats.iAiredSpots
                    tlStats.iFree = 0
                    tlStats.lMinAired = tlStats.iAiredSpots * tlStats.iLen
                    tmBarterTypeSort(llBarterLoop).tBarterStats = tlStats      'place the updated image backinto array
                Else                                                'threshold complete met, everything is payable
                    If llAiredSec > (CLng(tmVbf.iThreshold) * 60) Then      'exact # of seconds make up the threshold; remaining are all paid
                        ilThreshHoldMet = True

                        llTemp = llAiredSec - (CLng(tmVbf.iThreshold) * 60)
                        ilTemp = llTemp \ tlStats.iLen
                        tlStats.iPayable = (llAiredSec - (CLng(tmVbf.iThreshold) * 60)) / tlStats.iLen     'threshold in minutes, convert to seconds
                        If llTemp > ilTemp * tlStats.iLen Then
                            tlStats.iPayable = tlStats.iPayable + 1
                        End If
                        tlStats.iFree = tlStats.iAiredSpots - tlStats.iPayable
                        tlStats.lMinAired = tlStats.iAiredSpots * tlStats.iLen
                        tmBarterTypeSort(llBarterLoop).tBarterStats = tlStats  'place the updated image backinto array
                    Else                                            'threshold hasnt been met yet, everything is still free
                        tlStats.iPayable = 0
                        tlStats.iFree = tlStats.iAiredSpots
                        tlStats.lMinAired = tlStats.iAiredSpots * tlStats.iLen
                        tmBarterTypeSort(llBarterLoop).tBarterStats = tlStats
                    End If
                End If

            ElseIf tmVbf.sMethod = "U" Then     'units exceeded.  ie. Payable after 15 units exceeded
                llUsedItems = llUsedItems + tlStats.iAiredSpots     'accum # spots aired for the entire period of time (not just by week)
                ilAiredUnits = ilAiredUnits + tlStats.iAiredSpots
                If ilThreshHoldMet Then
                    tlStats.iPayable = tlStats.iAiredSpots
                    tlStats.iFree = 0
                    tlStats.lMinAired = tlStats.iAiredSpots * tlStats.iLen
                    tmBarterTypeSort(llBarterLoop).tBarterStats = tlStats      'place the updated image backinto array
                Else                                                'threshold complete met, everything is payable
                    If ilAiredUnits > tmVbf.iThreshold Then         'remaining are all paid
                        ilThreshHoldMet = True
                        tlStats.iPayable = ilAiredUnits - tmVbf.iThreshold
                        tlStats.iFree = tlStats.iAiredSpots - tlStats.iPayable
                        tlStats.lMinAired = tlStats.iAiredSpots * tlStats.iLen
                        tmBarterTypeSort(llBarterLoop).tBarterStats = tlStats  'place the updated image backinto array
                    Else                                            'threshold hasnt been met yet, everything is still free
                        tlStats.iPayable = 0
                        tlStats.iFree = tlStats.iAiredSpots
                        tlStats.lMinAired = tlStats.iAiredSpots * tlStats.iLen
                        tmBarterTypeSort(llBarterLoop).tBarterStats = tlStats
                    End If
                End If

            ElseIf tmVbf.sMethod = "X" Then                                'X free for Y paid. ie: 1 free spot for every 3 paid.
                If Trim$(tmVbf.sPerPeriod) = "" Or tmVbf.sPerPeriod = "W" Then
                    'start with 1st spot paid, then give "X" free ones; repeat until spots used up
                    ilAiredUnits = tlStats.iAiredSpots          'total spots aired
                    ilTempYSold = tmVbf.iYSold
                    Do While ilAiredUnits > 0      'done when no more spots remaining, or # Paid exceeded
                        ilTempYSold = tmVbf.iYSold              'paid
                        Do While ilTempYSold > 0 And ilAiredUnits > 0
                            tlStats.iPayable = tlStats.iPayable + 1
                            ilAiredUnits = ilAiredUnits - 1             'decrease spots aired
                            ilTempYSold = ilTempYSold - 1
                        Loop
                         ilTempXFree = tmVbf.iXFree
                         Do While ilTempXFree > 0 And ilAiredUnits > 0
                             tlStats.iFree = tlStats.iFree + 1
                             ilAiredUnits = ilAiredUnits - 1
                             ilTempXFree = ilTempXFree - 1
                         Loop
                    Loop
                    tlStats.lMinAired = tlStats.iAiredSpots * tlStats.iLen
                    tmBarterTypeSort(llBarterLoop).tBarterStats = tlStats
                Else                'bymonth or year,need to total by spot length to determine paid spots
                    'add up spot counts by length
                    blCalcMethodX = True
                    For ilTemp = 0 To 9             '10-5-18 zero base
'                    For ilTemp = 1 To 10

                        If tmVbf.iSpotLen(ilTemp) = tlStats.iLen Then
                            ilSpotCountByLen(ilTemp + 1) = ilSpotCountByLen(ilTemp + 1) + tlStats.iAiredSpots
                            Exit For
                        End If
                    Next ilTemp
                End If
                
                
'                'calculation is done on a contract by contract & spot length basis because of different rates across the
'                'different spot lengths.  Any spots remaining is considered payable
'                ilfirstTime = True
'
'                '# groups = # units aired / (free + payable)   i.e. 9 spots aired / (1 free + 3 payable)
'                '                                        9/4 = 2 groups
'                'Free (X): 1 * 2 (groups) = 2 free
'                'Payable(Y): 3 * 2 (groups) = 6 paid
'                '# aired - (total free + total payable) = remaining spots payable
'
'                '7-8-15 when spots aired is less than the # of free spots, need to alter the algorithm so all the spots are not payable
'                ilTempXFree = tmVbf.iXFree
'                ilTempYSold = tmVbf.iYSold
'
'                If ilTempXFree >= tlStats.iAiredSpots Then
'                    ilTempXFree = tlStats.iAiredSpots - ilTempYSold
'                    If ilTempXFree < 0 Then
'                        ilTempXFree = 0
'                    End If
'                End If
'                'If (tmVbf.iXFree + tmVbf.iYSold) > 0 Then
'                If (ilTempXFree + ilTempYSold) > 0 Then
'                    ilGroups = tlStats.iAiredSpots \ (ilTempXFree + ilTempYSold)      'obtain # of groups
'                Else
'                    ilGroups = 0
'                End If
'                tlStats.iFree = ilTempXFree * ilGroups                 'get the # of free spots
'                tlStats.iPayable = ilTempYSold * ilGroups              'get the number of payable spots
'                ilTemp = tlStats.iFree + tlStats.iPayable
'                tlStats.iPayable = tlStats.iPayable + (tlStats.iAiredSpots - ilTemp)
'                tlStats.lMinAired = tlStats.iAiredSpots * tlStats.iLen
'                tmBarterTypeSort(llBarterLoop).tBarterStats = tlStats
            Else                        'no method defined

            End If
        Next llBarterLoop
        
        If tmVbf.sPerPeriod = "Y" Then         'period interval is by year
            If tmVbf.sMethod = "M" Or tmVbf.sMethod = "U" Then      'barter method is by minutes or units exceeded
                'track minutes or units used so far
                blFoundTrack = False
                For ilTemp = 0 To UBound(tmYearly_Tracking) - 1
                     If tmYearly_Tracking(ilTemp).lVbfcode = tmVbf.lCode Then
                        blFoundTrack = True
                        Exit For
                     End If
                Next ilTemp
                If Not blFoundTrack Then
                    ilTemp = UBound(tmYearly_Tracking)
                    tmYearly_Tracking(ilTemp).lVbfcode = tmVbf.lCode
                    tmYearly_Tracking(ilTemp).iVefCode = tmVbf.iVefCode
                    tmYearly_Tracking(ilTemp).lUsedItems = llUsedItems
                    ReDim Preserve tmYearly_Tracking(0 To ilTemp + 1) As YEARLY_TRACKING
                End If
            End If
        End If

        If blCalcMethodX Then               'calc for every x get y by month or year
            'loop by the spot length counts

            For ilTemp = 1 To 10
                If ilSpotCountByLen(ilTemp) > 0 Then
                    llBarterValue = mGetBarterValueForLen(tmVbf.iSpotLen(ilTemp - 1))         '10-5-18 tmvbf.ispotlen is 0 base (adjust it); get the actual $ that should be paid to station
                    ilFree = 0
                    ilPayable = 0

                     'start with 1st spot paid, then give "X" free ones; repeat until spots used up
                    ilAiredUnits = ilSpotCountByLen(ilTemp)         'total spots aired
                    ilTotalAiredByLen = ilSpotCountByLen(ilTemp)    '10-19-18 save to place into cnt record to give % of revenue
                    ilTempYSold = tmVbf.iYSold
                    
                    Do While ilAiredUnits > 0      'done when no more spots remaining, or # Paid exceeded
                        ilTempYSold = tmVbf.iYSold              'paid
                        Do While ilTempYSold > 0 And ilAiredUnits > 0
                            'tlStats.iPayable = tlStats.iPayable + 1
                            ilPayable = ilPayable + 1
                            ilAiredUnits = ilAiredUnits - 1             'decrease spots aired
                            ilTempYSold = ilTempYSold - 1
                        Loop
                         ilTempXFree = tmVbf.iXFree
                         Do While ilTempXFree > 0 And ilAiredUnits > 0
                             'tlStats.iFree = tlStats.iFree + 1
                             ilFree = ilFree + 1
                             ilAiredUnits = ilAiredUnits - 1
                             ilTempXFree = ilTempXFree - 1
                         Loop
                    Loop
                    'look for entry by spot length in barter array.  place the free/payable in each record
                    'the report will not tally up these counts, but instead show the values on the total line
                    blFirstForThisSpotLength = True
                    llAccumRev = 0
                    For llBarterLoop = llLoIndex To llHiIndex
                        
                        '10-5-18 tmvbf.ispotlen is 0 base , adjust it
                        If tmVbf.iSpotLen(ilTemp - 1) = tmBarterTypeSort(llBarterLoop).tBarterStats.iLen And ilVefCode = tmBarterTypeSort(llBarterLoop).tBarterStats.iVefCode Then
'                            If blFirstForThisSpotLength Then
                            '10-19-18 for method X by month, now shows ratio of $ per contract, no longer ignore all cnts of same length in week
'                                blFirstForThisSpotLength = False            'set only 1 contracts within matching spot length
'                            Else
'                                tmBarterTypeSort(llBarterLoop).tBarterStats.bIgnoreForXport = True  'ignore duplicate spot lengths for the Barter method X on month & yearly intervals
'                                                                            'export it only once per spot length
'                            End If
                            tmBarterTypeSort(llBarterLoop).tBarterStats.iFree = ilFree
                            tmBarterTypeSort(llBarterLoop).tBarterStats.iPayable = ilPayable
                            tmBarterTypeSort(llBarterLoop).tBarterStats.lBarterValue = llBarterValue
                            '10-18-18 User wants to see the ratio of spots applied to each cnts by spot length
                            tmBarterTypeSort(llBarterLoop).tBarterStats.iTotalAiredByLen = ilTotalAiredByLen   'total spots aired by the spot length\
                            slStr = gLongToStrDec((CLng(tmBarterTypeSort(llBarterLoop).tBarterStats.iAiredSpots) * 10000), 2)  'convert # spots to string dec
                            slAirCount = gIntToStrDec(ilTotalAiredByLen, 0)     'convert total aired by length to string dec
                            slPct = gDivStr(slStr, slAirCount)
                            slDollar = gMulStr(gLongToStrDec((llBarterValue * ilPayable), 2), slPct)
                            llRevShare = (gStrDecToLong(slDollar, 2)) \ 100
                            If llAccumRev + llRevShare > llBarterValue * ilPayable Then     'do not distribute more than the max allowed, give the last one found that is processed, the remainder not to exceed
                                llRevShare = (llBarterValue * ilPayable) - llAccumRev
                                llAccumRev = llBarterValue * ilPayable      'max the accumulation revenue
                            Else
                                llAccumRev = llAccumRev + llRevShare
                                '2-13-19 make up the pennies to first in list
                                If ((llBarterValue * ilPayable) - llAccumRev > 0) And (llBarterLoop = llHiIndex) Then
                                    llRevShare = llRevShare + (llBarterValue * ilPayable) - llAccumRev
                                End If
                                
                            End If
                            'ilPct = tmBarterTypeSort(llBarterLoop).tBarterStats.iAiredSpots * 100 / ilTotalAiredByLen
                            'llRevShare = ((llBarterValue * ilPayable) * ilPct) / 100
                            tmBarterTypeSort(llBarterLoop).tBarterStats.lRevShareByLen = llRevShare
                        End If
                    Next llBarterLoop
                End If
            Next ilTemp

        End If
        
        
        Exit Sub
End Sub
'
'           mGetBarterValueForLen - grab the actual acquisition $ due the station
'           for the spot length
'           <input> spot length to match
'           <return> value of the spot length to pay each spot
Public Function mGetBarterValueForLen(ilLen As Integer) As Long
Dim ilLoop As Integer
    mGetBarterValueForLen = 0
'    For ilLoop = 1 To 10
    For ilLoop = 0 To 9         '10-5-18 zero base
        If ilLen = tmVbf.iSpotLen(ilLoop) Then
            mGetBarterValueForLen = tmVbf.lActAcqCost(ilLoop)
            Exit For
        End If
    Next ilLoop
End Function
'
'           mWriteBarterdue - create prepass record for printing
'           <input>  llLoIndex - lo index into array to process for 1 vehicle
'                    llHiindex - hi index into array to process for 1 vehicle
Public Sub mWriteBarterDue(llLoIndex As Long, llHiIndex As Long)
Dim llLoopBarter As Long
Dim ilfirstTime As Integer
Dim ilLen As Integer
Dim ilRet As Integer
Dim ilLenUsed As Integer
Dim tlStats As BARTERSTATS
Dim slPerPeriod As String
Dim ilTemp As Integer
Dim llTemp As Long
Dim llPrevBalDate As Long

        tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
        tmGrf.iGenDate(1) = igNowDate(1)
        tmGrf.lGenTime = lgNowTime
        tmGrf.iCode2 = imCurrentMonth           'month requested
        tmGrf.iYear = imYear                    'year requested
        ilfirstTime = True

        
        For llLoopBarter = llLoIndex To llHiIndex
            tlStats = tmBarterTypeSort(llLoopBarter).tBarterStats
            If ilfirstTime Then
                ilfirstTime = False
                'create the header record that contains the info for the vehicle agreement
                tmGrf.iVefCode = tlStats.iVefCode
                'clear all fields that detail uses and header does not
                'tmGrf.iPerGenl(1) = 0
                'tmGrf.iPerGenl(2) = 0
                'tmGrf.iPerGenl(3) = 0
                'tmGrf.iPerGenl(4) = 0
                'tmGrf.iPerGenl(5) = 0
                'tmGrf.iPerGenl(7) = 0       'spot length field for sorting when method is pay x for y spots and per period is month or year; otherwise 0
                tmGrf.iPerGenl(0) = 0
                tmGrf.iPerGenl(1) = 0
                tmGrf.iPerGenl(2) = 0
                tmGrf.iPerGenl(3) = 0
                tmGrf.iPerGenl(4) = 0
                tmGrf.iPerGenl(6) = 0       'spot length field for sorting when method is pay x for y spots and per period is month or year; otherwise 0
                'tmGrf.lDollars(1) = 0
                'tmGrf.lDollars(2) = 0
                tmGrf.lDollars(0) = 0
                tmGrf.lDollars(1) = 0
                tmGrf.lDollars(2) = 0
                tmGrf.lDollars(3) = 0
                tmGrf.sBktType = ""             'flag previously paid
                'tmGrf.iDateGenl(0, 1) = tlStats.iGetSbfStartDate(0)         'start/end dates to show for the year gathered on yearly agreements
                'tmGrf.iDateGenl(1, 1) = tlStats.iGetSbfStartDate(1)
                'tmGrf.iDateGenl(0, 2) = tlStats.iGetSbfEndDate(0)
                'tmGrf.iDateGenl(1, 2) = tlStats.iGetSbfEndDate(1)
                tmGrf.iDateGenl(0, 0) = tlStats.iGetSbfStartDate(0)         'start/end dates to show for the year gathered on yearly agreements
                tmGrf.iDateGenl(1, 0) = tlStats.iGetSbfStartDate(1)
                tmGrf.iDateGenl(0, 1) = tlStats.iGetSbfEndDate(0)
                tmGrf.iDateGenl(1, 1) = tlStats.iGetSbfEndDate(1)
                
                For ilLen = 1 To 10             'spot length and acq $ for agreement, show in header
                    'tmGrf.iPerGenl(8 + ilLen) = 0
                    'tmGrf.lDollars(8 + ilLen) = 0
                    tmGrf.iPerGenl(7 + ilLen) = 0
                    tmGrf.lDollars(7 + ilLen) = 0
                Next ilLen
                ilLenUsed = 1
'                For ilLen = 1 To 10
                For ilLen = 0 To 9             '10-5-18 zero base
                    If tmVbf.lActAcqCost(ilLen) <> 0 Then      'if no actual acq. cost, no need to show on the report
                        'tmGrf.iPerGenl(8 + ilLenUsed) = tmVbf.iSpotLen(ilLen)
                        'tmGrf.lDollars(8 + ilLenUsed) = tmVbf.lActAcqCost(ilLen)
                        tmGrf.iPerGenl(7 + ilLenUsed) = tmVbf.iSpotLen(ilLen)
                        tmGrf.lDollars(7 + ilLenUsed) = tmVbf.lActAcqCost(ilLen)
                        ilLenUsed = ilLenUsed + 1
                    End If
                Next ilLen

                'build description for barter method
                
                If tmVbf.sMethod = "A" Then
                    tmGrf.sGenDesc = "All Cash"
                ElseIf tmVbf.sMethod = "M" Then        'minutes exceeding
                    tmGrf.sGenDesc = "Paid spots exceeding " & Trim$(str$(tmVbf.iThreshold)) & " minutes per "
                   
                ElseIf tmVbf.sMethod = "U" Then         'units exceeding
                    tmGrf.sGenDesc = "Paid spots exceeding " & Trim$(str$(tmVbf.iThreshold)) & " spots per "
                    For ilTemp = 0 To UBound(tmYearly_Tracking) - 1
                        If tmVbf.lCode = tmYearly_Tracking(ilTemp).lVbfcode Then        'update the minutes used
                            
                            Exit For
                        End If
                    Next ilTemp
                ElseIf tmVbf.sMethod = "X" Then             'X free for Y
                    tmGrf.sGenDesc = Trim$(str$(tmVbf.iXFree)) & " free spots for every " & Trim$(str$(tmVbf.iYSold)) & " paid per "
                Else
                    tmGrf.sGenDesc = "None"
                End If
                
                If tmVbf.sPerPeriod = "Y" Then          'if yearly method for either Minutes or Units exceeded, update the balance of minutes/units used so far.
                                                        'update only if the end of the requested period is later than the period already updated
                    For ilTemp = 0 To UBound(tmYearly_Tracking) - 1
                        If tmVbf.lCode = tmYearly_Tracking(ilTemp).lVbfcode Then        'update the minutes used
                            tmVbfSrchKey0.lCode = tmYearly_Tracking(ilTemp).lVbfcode
                            imVBfRecLen = Len(tmVbf)
                            ilRet = btrGetEqual(hmVbf, tmVbf, imVBfRecLen, tmVbfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)
                            If ilRet = BTRV_ERR_NONE Then
                                'gUnpackDateLong tmGrf.iDateGenl(0, 2), tmGrf.iDateGenl(1, 2), llTemp       'convert the end of the period
                                gUnpackDateLong tmGrf.iDateGenl(0, 1), tmGrf.iDateGenl(1, 1), llTemp       'convert the end of the period
                                 
                                If tmVbf.iBalanceDate(0) = 0 Or tmVbf.iBalanceDate(1) = 0 Then
                                    'tmVbf.iBalanceDate(0) = tmGrf.iDateGenl(0, 2)
                                    'tmVbf.iBalanceDate(1) = tmGrf.iDateGenl(1, 2)
                                    tmVbf.iBalanceDate(0) = tmGrf.iDateGenl(0, 1)
                                    tmVbf.iBalanceDate(1) = tmGrf.iDateGenl(1, 1)
                                Else
                                    gUnpackDateLong tmVbf.iBalanceDate(0), tmVbf.iBalanceDate(1), llPrevBalDate
                                End If
                                If llTemp >= llPrevBalDate Then         'later date requested, update the balance remaining
                                    'tmVbf.iBalanceDate(0) = tmGrf.iDateGenl(0, 2)
                                    'tmVbf.iBalanceDate(1) = tmGrf.iDateGenl(1, 2)
                                    tmVbf.iBalanceDate(0) = tmGrf.iDateGenl(0, 1)
                                    tmVbf.iBalanceDate(1) = tmGrf.iDateGenl(1, 1)
                                    If tmVbf.sMethod = "M" Then
                                       tmVbf.lBalance = (CLng(tmVbf.iThreshold) * 60) - (tmYearly_Tracking(ilTemp).lUsedItems)        'used time is in seconds
                                    Else
                                        tmVbf.lBalance = tmVbf.iThreshold - tmYearly_Tracking(ilTemp).lUsedItems
                                    End If
                                    ilRet = btrUpdate(hmVbf, tmVbf, imVBfRecLen)
                                    Exit For
                                End If
                            End If
                        End If
                    Next ilTemp
                End If
                
                'tmGrf.iPerGenl(6) = 1               'header record
                tmGrf.iPerGenl(5) = 1               'header record
                'tmGrf.lDollars(3) = tlStats.lVbfcode
                tmGrf.lDollars(2) = tlStats.lVbfcode
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            End If
            tmGrf.lChfCode = tlStats.lChfCode
            tmGrf.iVefCode = tlStats.iVefCode
            'tmGrf.iPerGenl(1) = tlStats.iLen
            'tmGrf.iPerGenl(2) = tlStats.iOrderedSpots
            'tmGrf.iPerGenl(3) = tlStats.iAiredSpots
            'tmGrf.iPerGenl(4) = tlStats.iFree
            'tmGrf.iPerGenl(5) = tlStats.iPayable
            'tmGrf.iPerGenl(6) = 11                 'detail (vs header)
            'tmGrf.iPerGenl(7) = tlStats.iLenForSortMethodX
            tmGrf.iPerGenl(0) = tlStats.iLen
            tmGrf.iPerGenl(1) = tlStats.iOrderedSpots
            tmGrf.iPerGenl(2) = tlStats.iAiredSpots
            tmGrf.iPerGenl(3) = tlStats.iFree
            tmGrf.iPerGenl(4) = tlStats.iPayable
            tmGrf.iPerGenl(5) = 11                 'detail (vs header)
            tmGrf.iPerGenl(6) = tlStats.iLenForSortMethodX
            'tmGrf.lDollars(1) = tlStats.lMinAired
            'tmGrf.lDollars(2) = tlStats.lBarterValue
            'tmGrf.lDollars(3) = tlStats.lVbfcode
            tmGrf.lDollars(0) = tlStats.lMinAired
            tmGrf.lDollars(1) = tlStats.lBarterValue
            tmGrf.lDollars(2) = tlStats.lVbfcode
            tmGrf.lDollars(3) = tlStats.lRevShareByLen          '10-19-18
            tmGrf.lDollars(4) = tlStats.lInvNo                  '10-19-18
            tmGrf.sBktType = tlStats.sBarterPaid
            tmGrf.iDate(0) = tlStats.iPostDate(0)
            tmGrf.iDate(1) = tlStats.iPostDate(1)
            gPackDateLong tlStats.lWeek, tmGrf.iStartDate(0), tmGrf.iStartDate(1)
            'tmGrf.iDateGenl(0, 1) = tlStats.iGetSbfStartDate(0)         'start/end dates to show for the year gathered on yearly agreements
            'tmGrf.iDateGenl(1, 1) = tlStats.iGetSbfStartDate(1)
            'tmGrf.iDateGenl(0, 2) = tlStats.iGetSbfEndDate(0)
            'tmGrf.iDateGenl(1, 2) = tlStats.iGetSbfEndDate(1)
            tmGrf.iDateGenl(0, 0) = tlStats.iGetSbfStartDate(0)         'start/end dates to show for the year gathered on yearly agreements
            tmGrf.iDateGenl(1, 0) = tlStats.iGetSbfStartDate(1)
            tmGrf.iDateGenl(0, 1) = tlStats.iGetSbfEndDate(0)
            tmGrf.iDateGenl(1, 1) = tlStats.iGetSbfEndDate(1)

            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)

        Next llLoopBarter
    Exit Sub
End Sub

Public Sub mOpenBarterFiles(Form As Form)
Dim ilRet As Integer

    imTerminate = False

    hmCHF = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gOpenBarterFiles
    gBtrvErrorMsg ilRet, "gOpenBarterFiles (btrOpen: Chf.Btr)", Form
    On Error GoTo 0
    imCHFRecLen = Len(tmChf) 'btrRecordLength(hmChf)    'Get Chf size

    hmSbf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gOpenBarterFiles
    gBtrvErrorMsg ilRet, "gOpenBarterFiles (btrOpen: Sbf.Btr)", Form
    On Error GoTo 0
    imSbfRecLen = Len(tmSbf) 'btrRecordLength(hmSbf)    'Get Sbf size

    hmVsf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gOpenBarterFiles
    gBtrvErrorMsg ilRet, "gOpenBarterFiles (btrOpen: Vsf.Btr)", Form
    On Error GoTo 0
    imVsfRecLen = Len(tmVsf)

    hmGrf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gOpenBarterFiles
    gBtrvErrorMsg ilRet, "gOpenBarterFiles (btrOpen: Grf.Btr)", Form
    On Error GoTo 0
    imGrfRecLen = Len(tmGrf)

    hmRwf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmRwf, "", sgDBPath & "Rwf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gOpenBarterFiles
    gBtrvErrorMsg ilRet, "gOpenBarterFiles (btrOpen: Rwf.Btr)", Form
    On Error GoTo 0
    imRwfRecLen = Len(tmRwf)

    'hmVbf = CBtrvTable(ONEHANDLE)
    hmVbf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmVbf, "", sgDBPath & "Vbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gOpenBarterFiles
    gBtrvErrorMsg ilRet, "gOpenBarterFiles (btrOpen: Vbf.Btr)", Form
    On Error GoTo 0
    imVBfRecLen = Len(tmVbf)
    
    '10-19-18 need receivables/history for inv #
    hmRvf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gOpenBarterFiles
    gBtrvErrorMsg ilRet, "gOpenBarterFiles (btrOpen: Rvf.Btr)", Form
    On Error GoTo 0
    imRvfRecLen = Len(tmRvf)

    hmPhf = CBtrvTable(TWOHANDLES)
    ilRet = btrOpen(hmPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gOpenBarterFiles
    gBtrvErrorMsg ilRet, "gOpenBarterFiles (btrOpen: Phf.Btr)", Form
    On Error GoTo 0
    
    Exit Sub
gOpenBarterFiles:
    On Error GoTo 0
    imTerminate = True
    ilRet = btrClose(hmVsf)
    ilRet = btrClose(hmRwf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmVbf)

    btrDestroy hmVsf
    btrDestroy hmRwf
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmSbf
    btrDestroy hmVbf
    Exit Sub


End Sub
'           mWriteBarterExport - create a record for the barter paid export
'           <input>  llLoIndex - lo index into array to process for 1 vehicle
'                    llHiindex - hi index into array to process for 1 vehicle
Public Sub mWriteBarterExport(llLoIndex As Long, llHiIndex As Long, hlExportFile As Integer)
Dim llLoopBarter As Long
Dim ilRet As Integer
Dim ilAdfIndex As Integer
Dim ilVafIndex As Integer
Dim slField1 As String              'Vendor ID from VAF
Dim slField2 As String              'Brnch code from VAF
Dim slField3 As String              'Start Week
Dim slField4 As String              'Contract #
Dim slField5 As String              'Advertiser/Product
Dim slField6 As String              'due station
Dim slField7 As String               '10-19-18 Assoc. Inv #
Dim slStr As String
Dim ilErr As Integer
Dim tlStats As BARTERSTATS
Dim ilVefInx As Integer

        For llLoopBarter = llLoIndex To llHiIndex
            tlStats = tmBarterTypeSort(llLoopBarter).tBarterStats
            If Not tlStats.bIgnoreForXport Then     'entries that are method type X with monthly and yearly will be excluded, except for one of those
                                                    'for each spot length within the barter vehicle
                ilAdfIndex = gBinarySearchAdf(tlStats.iAdfCode)
                ilVafIndex = gBinarySearchVaf(tlStats.iVefCode, tmVaf())
                slField1 = ""
                slField2 = ""
                If ilVafIndex <> -1 Then
                    slField1 = tmVaf(ilVafIndex).sVendorID                  'Vendor ID
                    If tlStats.sCashTrade = "C" Then
                        slField2 = tmVaf(ilVafIndex).sBranchCodeCash
                    Else
                        slField2 = tmVaf(ilVafIndex).sBranchCodeTrade
                    End If
                End If
                slField3 = Format$(tlStats.lWeek, "m/d/yy")
                slField4 = Trim$(str(tlStats.lCntrNo))       'contract #
                slField5 = ""                   'advt/product
                If ilAdfIndex <> -1 Then
                    slField5 = Trim$(tgCommAdf(ilAdfIndex).sName)
                End If
                If tlStats.sProduct <> "" Then
                    slField5 = slField5 & "/" & Trim$(tlStats.sProduct)
                End If
                If tlStats.sMethod = "X" And tlStats.sPeriodForMethodX = "M" Then                    '10-19-18 this method needs to see ratio of $ per contract when interval by month
                    slField6 = gLongToStrDec(tlStats.lRevShareByLen, 2)
                Else
                    slField6 = gLongToStrDec((tlStats.iPayable * tlStats.lBarterValue), 2)
                End If
                slField7 = str$(tlStats.lInvNo)                                                 '10-19-18  Inv # from receivables matching the sbf contract # and date
    
                slStr = """" & Trim$(slField1) & """" & ","             'vendor id
                slStr = slStr & """" & Trim$(slField2) & """" & ","     'branch code
                'slStr = slStr & """" & Trim$(slField3) & """" & ","     'date
                slStr = slStr & Trim$(slField3) & ","      'date
                'slStr = slStr & """" & Trim$(slField4) & """" & ","     'contract
                slStr = slStr & Trim$(slField4) & ","      'contract
                slStr = slStr & """" & Trim$(slField5) & """" & ","     'advt/prod
                'slStr = slStr & """" & Trim$(slField6) & """"           'due station
                slStr = slStr & Trim$(slField6)             'due station
                
                slStr = slStr & "," & Trim$(slField7)               'Inv #
                
'                'debug only
'                ilVefInx = gBinarySearchVef(tlStats.iVefCode)
'                slStr = slStr & "," & Trim$(tgMVef(ilVefInx).sName)
                
                On Error GoTo cmcExportErr
                Print #hlExportFile, slStr        'write header description
                On Error GoTo 0
            End If
        Next llLoopBarter
        Exit Sub
cmcExportErr:
        ilErr = err.Number
        Resume Next
    
End Sub
