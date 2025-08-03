Attribute VB_Name = "RPTCRPC"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrpc.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
'Public igYear As Integer                'budget year used for filtering
'Public igMonthOrQtr As Integer          'entered month or qtr
'Public igNowDate(0 To 1) As Integer
'Public igNowTime(0 To 1) As Integer
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer            'Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
'Quarterly Avails
Dim hmAvr As Integer            'Quarterly Avails file handle
Dim tmAvr() As AVR                'AVR record image
Dim imAvrRecLen As Integer        'AVR record length
'Dim tmBooked() As BOOKSPOTS
'Dim lmSAvailsDates(1 To 13) As Long   'Start Dates of avail week
Dim lmSAvailsDates(0 To 13) As Long   'Start Dates of avail week. Index zero ignored
'Dim lmEAvailsDates(1 To 13) As Long   'End dates of avail week
Dim lmEAvailsDates(0 To 13) As Long   'End dates of avail week. Index zero ignored
Dim hmCHF As Integer            'Contract header file handle
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmSdf As Integer            'Spot detail file handle
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdf As SDF
Dim hmSsf As Integer
Dim hmSmf As Integer            'MG file handle
Dim tmSmf As SMF                'SMF record image
Dim imSmfRecLen As Integer        'SMF record length
'Log Calendar
Dim hmLcf As Integer            'Log Calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim imLcfRecLen As Integer        'LCF record length
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim tmRcf As RCF
Dim hmFsf As Integer            'Feed file handle
Dim imFsfRecLen As Integer        'FSF record length
Dim tmFsf As FSF
Dim hmAnf As Integer            'Named avail file handle
Dim imAnfRecLen As Integer        'AnF record length
Dim tmAnf As ANF
'

'*********************************************************************
'*
'*      Procedure Name:gCRAvailsProposal
'*
'*             Created:12/29/97      By:D. Hosaka
'*            Modified:              By:
'*
'*            Comments: Generate avails & spot Data for
'*                      any report requiring Inventory,
'*                      avails, sold.
'*            Find availabilty by vehicle for a given
'*            rate card.
'*            Copy of gCreateAvails
'*
'*      3/11/98 Look at "Base DP" only, (not Report DP)
'*      4/12/98 Remove duplication of spots from vehicle
'*              These spots appeared to be moved across vehicles
'*      7-29-04 Option to include/exclude contract/feed spots
'**********************************************************************
Sub gCRAvailsProposal()
'
    Dim slCntrTypes As String
    Dim slCntrStatus As String
    Dim ilHOState As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilVehicle As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim slStartDate As String
    Dim ilQNo As Integer
    Dim slQSDate As String
    Dim ilFirstQ As Integer
    Dim ilRec As Integer
    Dim ilVpfIndex As Integer
    Dim ilUpper As Integer
    Dim ilDateOk As Integer
    Dim llRif As Long
    Dim ilRdf As Integer
    Dim ilRcf As Integer
    Dim ilFound As Integer
    Dim ilIndex As Integer
    Dim llStart As Long
    Dim llEnd As Long
    Dim slStartQtr As String                    'start date of qtr based on user date entered
    Dim slEndQtr As String
    Dim llEffDate As Long
    Dim ilNoQtrs As Integer                     'user # qtrs requested
    Dim tlCntTypes As CNTTYPES
    Dim ilSaveSort As Integer                   'DP or RIF field:  sort code
    Dim slSaveReport As String                  'DP or RIF field:  Save on report
    Dim tlAvr As AVR                            'Avails generated from SDF
    Dim ilStdQtr As Integer
    Dim llStartOfReport As Long                 'earliest date to begin processing avails & proposals
    Dim llEndOfReport As Long                   'latest date of end processing avails & proposals
    Dim ilCurrentRecd As Integer
    Dim llContrCode As Long
    Dim ilClf As Integer
    Dim ilMinorSet As Integer                   'sort index selected (minor sort) - unused
    Dim ilMajorSet As Integer                   'sort index selected (major sort)
    Dim ilmnfMinorCode As Integer
    Dim ilMnfMajorCode As Integer               'Multi name sort code for major sort selection
    'ReDim llProject(1 To 53) As Long              '53 weeks of spot counts  from flights
    ReDim llProject(0 To 53) As Long              '53 weeks of spot counts  from flights. Index zero ignored
    'ReDim llStartWeeks(1 To 53) As Long           'max 53 start dates of weeks
    ReDim llStartWeeks(0 To 53) As Long           'max 53 start dates of weeks. Index zero ignored
    ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
    ReDim tlAvrContr(0 To 0) As AVR
    Dim tlAvailInfo As AVAILCOUNT
    Dim llToday As Long                     'todays date
    Dim ilProp As Integer
    Dim slUserRequest As String * 1  '8-13-10 by 30/60 (B) or Units (U)


    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmAvr = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAvr, "", sgDBPath & "Avr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAvr)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAvr
        btrDestroy hmCHF
    Exit Sub
    End If
    imAvrRecLen = Len(tlAvr)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAvr)
        ilRet = btrClose(hmCHF)
        btrDestroy hmClf
        btrDestroy hmAvr
        btrDestroy hmCHF
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAvr)
        ilRet = btrClose(hmCHF)
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmAvr
        btrDestroy hmCHF
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAvr)
        ilRet = btrClose(hmCHF)
        btrDestroy hmMnf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmAvr
        btrDestroy hmCHF
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        btrDestroy hmVef
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmSsf
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmLcf
        btrDestroy hmSsf
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    imLcfRecLen = Len(tmLcf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmLcf)
        ilRet = btrClose(hmSsf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmSmf
        btrDestroy hmLcf
        btrDestroy hmSsf
        btrDestroy hmSdf
        btrDestroy hmVef
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)

    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAvr)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmAvr
        btrDestroy hmCHF
        Exit Sub
    End If
    imVsfRecLen = Len(tmVsf)

    hmFsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAvr)
        ilRet = btrClose(hmCHF)
        btrDestroy hmFsf
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmAvr
        btrDestroy hmCHF
        Exit Sub
    End If
    imFsfRecLen = Len(tmFsf)

    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAvr)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAnf
        btrDestroy hmFsf
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmAvr
        btrDestroy hmCHF
        Exit Sub
    End If
    imAnfRecLen = Len(tmAnf)

    slDate = Format$(gNow(), "m/d/yy")   'Date$
    llToday = gDateValue(slDate) + 1  'Not include todays date

    tlCntTypes.iHold = gSetCheck(RptSelPC!ckcSelC1(0).Value)            'contract spots
    tlCntTypes.iOrder = gSetCheck(RptSelPC!ckcSelC1(1).Value)           'contract spots
    tlCntTypes.iNetwork = gSetCheck(RptSelPC!ckcSelC1(13).Value)        'feed spot
    tlCntTypes.iStandard = gSetCheck(RptSelPC!ckcSelC1(2).Value)
    tlCntTypes.iReserv = gSetCheck(RptSelPC!ckcSelC1(3).Value)
    tlCntTypes.iRemnant = gSetCheck(RptSelPC!ckcSelC1(4).Value)
    tlCntTypes.iDR = gSetCheck(RptSelPC!ckcSelC1(5).Value)
    tlCntTypes.iPI = gSetCheck(RptSelPC!ckcSelC1(6).Value)
    tlCntTypes.iTrade = gSetCheck(RptSelPC!ckcSelC1(9).Value)
    tlCntTypes.iMissed = gSetCheck(RptSelPC!ckcSelC1(10).Value)
    tlCntTypes.iNC = gSetCheck(RptSelPC!ckcSelC1(11).Value)
    tlCntTypes.iXtra = gSetCheck(RptSelPC!ckcSelC1(12).Value)
    tlCntTypes.iPSA = gSetCheck(RptSelPC!ckcSelC1(7).Value)
    tlCntTypes.iPromo = gSetCheck(RptSelPC!ckcSelC1(8).Value)
    'following added for Qtrly booked spots
    tlCntTypes.iOrphan = False          'do not show separate line for orphan spots, combine with
                                        'daypart the orphan spot falls within (not spans)
    tlCntTypes.iDayOption = 0           'Daypart option (vs days within dp or DP within days)
    tlCntTypes.iBuildKey = False         'for qtrly booked, build additl tables by line-- otherwise dont build special key
    tlCntTypes.iShowReservLine = False    'qtrly booked does not show reserved as separate line
    If RptSelPC!rbcSelC6(1).Value Then      'show reserve line separately
        tlCntTypes.iShowReservLine = True
    End If

    '4-18-02 SEtup proposaltypes to include
    tlCntTypes.iWorking = gSetCheck(RptSelPC!ckcPropType(0).Value)
    tlCntTypes.iComplete = gSetCheck(RptSelPC!ckcPropType(1).Value)
    tlCntTypes.iIncomplete = gSetCheck(RptSelPC!ckcPropType(2).Value)

    ilStdQtr = False                    'assume to use the start date entered
    If RptSelPC!rbcSelC4(0).Value Then  'use start of quarter
        ilStdQtr = True
    End If
    
    '8-13-10 determine whether output should be shown by units, or 30/60
    'changed subrtn to not exceed the number of units for a unit output
    If RptSelPC!rbcSelC2(0).Value = True Then  'units
        slUserRequest = "U"
    Else
        slUserRequest = "B"
    End If

    If RptSelPC!rbcSelC5(1).Value Then  'summary (bottom line for avails only_
        tlCntTypes.iDetail = False
        tlCntTypes.sAvailType = "A"     'assume Avails version for Pressure report
    Else                                'detail (show separate lines for holds, reserves, sold)
        tlCntTypes.iDetail = True
        tlCntTypes.sAvailType = "S"         'assume Avails detail version for Pressure report
    End If
    ilRet = gObtainRcfRifRdf()          'get the rate cards and assoc dayparts
    ilLoop = RptSelPC!cbcSet1.ListIndex 'get the type of sorting for output
    ilMajorSet = gFindVehGroupInx(ilLoop, tgVehicleSets1())
    'get all the dates needed to work with
    slDate = RptSelPC!edcSelCFrom.Text               'effective date entred
    slStartDate = slDate                               'save orig date entered
    llEffDate = gDateValue(slStartDate)
    If ilStdQtr Then
        slDate = gObtainYearStartDate(0, slDate)    'get start of std year so it can advance to the current requested qtr
        llDate = gDateValue(slDate)
        Do While (llEffDate < llDate) Or (llEffDate > llDate + (13 * 7) - 1)  'advance to current requested qtr
            llDate = llDate + (13 * 7)
        Loop
    Else                                            'use the date entered
        llDate = gDateValue(slDate)
    End If
    slQSDate = Format$(llDate, "m/d/yy")

    tmVef.iCode = 0
    ilNoQtrs = Val(RptSelPC!edcSelCFrom1.Text)
    'Save the start of the quarter or requested date, use for Proposals gathering of contracts
    llStartOfReport = llDate
    llEndOfReport = llDate + ((13 * 7) * ilNoQtrs) - 1



    'gather all contracts whose entered date is equal or prior to the requested date (gather from beginning of std year to
    'input date
    slCntrTypes = gBuildCntTypes()
    slCntrStatus = "HOGN"            'get status sch hold & order , unsch hold & order
    'determine the proposal statuses
    If tlCntTypes.iComplete = True Then
        slCntrStatus = slCntrStatus & "C"
    End If
    If tlCntTypes.iIncomplete = True Then
        slCntrStatus = slCntrStatus & "I"
    End If
    If tlCntTypes.iWorking = True Then
        slCntrStatus = slCntrStatus & "W"
    End If
    ilHOState = 4                      'get everything for the date span that is not deleted
    slStartQtr = Format$(llStartOfReport - 90, "m/d/yy")
    slEndQtr = Format$(llEndOfReport + 90, "m/d/yy")

    ilRet = gObtainCntrForDate(RptSelPC, slStartQtr, slEndQtr, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
    'build array of contracts that are HO with another that is either C, I, N, or G.  There are the contracts that
    'need to be ignored while gathering the spots  to prevent duplicating contracts processed due to modifications
    ReDim llIgnoreCodes(0 To 0) As Long
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
        If tlChfAdvtExt(ilCurrentRecd).sStatus = "H" Or tlChfAdvtExt(ilCurrentRecd).sStatus = "O" Then
            For ilIndex = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
                If ilIndex <> ilCurrentRecd Then        'ignore itself

                    If tlChfAdvtExt(ilCurrentRecd).lCntrNo = tlChfAdvtExt(ilIndex).lCntrNo Then
                        'Also ignore any proposals that were not selected
                        For ilProp = LBound(tgProcessProp) To UBound(tgProcessProp) - 1
                            If tgProcessProp(ilProp).lCntrNo = tlChfAdvtExt(ilCurrentRecd).lCntrNo Then
                                llIgnoreCodes(UBound(llIgnoreCodes)) = tlChfAdvtExt(ilCurrentRecd).lCode   'save the schedule order/hold code so
                                'that when processing spots it will be ignored
                                ReDim Preserve llIgnoreCodes(0 To UBound(llIgnoreCodes) + 1) As Long
                                Exit For
                            End If
                        Next ilProp
                    End If

                End If
            Next ilIndex
        End If
    Next ilCurrentRecd

    For ilQNo = 1 To ilNoQtrs Step 1
        llDate = gDateValue(slQSDate)
        For ilLoop = 1 To 13 Step 1
            If ilStdQtr Then
                If (llEffDate >= llDate) And (llEffDate <= llDate + 6) Then
                    ilFirstQ = ilLoop
                End If
            Else
                ilFirstQ = 1
            End If
            lmSAvailsDates(ilLoop) = llDate
            lmEAvailsDates(ilLoop) = llDate + 6
            llDate = llDate + 7
        Next ilLoop
        For ilVehicle = 0 To RptSelPC!lbcSelection(0).ListCount - 1 Step 1
            If (RptSelPC!lbcSelection(0).Selected(ilVehicle)) Then
                slNameCode = tgCSVNameCode(ilVehicle).sKey 'RptSelSP!lbcCSVNameCode.List(ilVehicle)
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slName, 3, "|", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                ilVpfIndex = -1
                'For ilLoop = 0 To UBound(tgVpf) Step 1
                '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
                    ilLoop = gBinarySearchVpf(ilVefCode)
                    If ilLoop <> -1 Then
                        ilVpfIndex = ilLoop
                '        Exit For
                    End If
                'Next ilLoop
                If ilVpfIndex >= 0 Then
                    For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                        tmRcf = tgMRcf(ilRcf)
                        ilDateOk = False
                        For ilLoop = 0 To RptSelPC!lbcSelection(1).ListCount - 1 Step 1
                            slNameCode = tgRateCardCode(ilLoop).sKey
                            ilRet = gParseItem(slNameCode, 3, "\", slCode)
                            If Val(slCode) = tgMRcf(ilRcf).iCode Then
                                If (RptSelPC!lbcSelection(1).Selected(ilLoop)) Then
                                    ilDateOk = True
                                End If
                                Exit For
                            End If
                        Next ilLoop

                        If ilDateOk Then
                            ReDim tmAvr(0 To 0) As AVR
                            ReDim tmAvRdf(0 To 0) As RDF
                            ReDim tmRifRate(0 To 0) As RIF

                            ilUpper = 0
                            'Setup the DP with rates to process
                            For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                                If tgMRif(llRif).iRcfCode = tgMRcf(ilRcf).iCode Then
                                    'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                    ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                                    If ilRdf <> -1 Then
                                        'determine if rate card items have sort codes, etc; otherwise use DP
                                        If tgMRif(llRif).sRpt <> "Y" And tgMRif(llRif).sRpt <> "N" Then
                                            slSaveReport = tgMRdf(ilRdf).sReport
                                        Else
                                            slSaveReport = tgMRif(llRif).sRpt
                                        End If
                                        If tgMRif(llRif).iSort = 0 Then
                                            ilSaveSort = tgMRdf(ilRdf).iSortCode
                                        Else
                                            ilSaveSort = tgMRif(llRif).iSort
                                        End If
                                        If tgMRdf(ilRdf).iCode = tgMRif(llRif).iRdfCode And slSaveReport = "Y" And tgMRdf(ilRdf).sState <> "D" And tgMRif(llRif).iVefCode = ilVefCode Then
                                            ilFound = False
                                            For ilLoop = LBound(tmAvRdf) To ilUpper - 1 Step 1
                                                If tmAvRdf(ilLoop).iCode = tgMRdf(ilRdf).iCode Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next ilLoop
                                            If Not ilFound Then
                                                tmAvRdf(ilUpper) = tgMRdf(ilRdf)
                                                tmRifRate(ilUpper) = tgMRif(llRif)
                                                tmRifRate(ilUpper).iSort = ilSaveSort
                                                ilUpper = ilUpper + 1
                                                ReDim Preserve tmAvRdf(0 To ilUpper) As RDF
                                                ReDim Preserve tmRifRate(0 To ilUpper) As RIF
                                            End If
                                    '        Exit For
                                        End If
                                    End If
                                    'Next ilRdf
                                End If
                            Next llRif
                            llStart = lmSAvailsDates(ilFirstQ)
                            llEnd = lmEAvailsDates(13)
                            'gGetSpotCounts ilVefcode, ilVpfIndex, ilFirstQ, llStart, llEnd, lmSAvailsDates(), lmEAvailsDates(), tmAvRdf(), tmRifRate(), tlCntTypes, llIgnoreCodes()
                            tlAvailInfo.iVefCode = ilVefCode           'setup variables sent via structure
                            tlAvailInfo.iVpfIndex = ilVpfIndex
                            tlAvailInfo.iFirstBkt = ilFirstQ
                            tlAvailInfo.lSDate = llStart
                            tlAvailInfo.lEDate = llEnd
                            tlAvailInfo.hLcf = hmLcf
                            tlAvailInfo.hSdf = hmSdf
                            tlAvailInfo.hSsf = hmSsf
                            tlAvailInfo.hSmf = hmSmf
                            tlAvailInfo.hChf = hmCHF
                            tlAvailInfo.hClf = hmClf
                            tlAvailInfo.hCff = hmCff
                            tlAvailInfo.hVef = hmVef
                            tlAvailInfo.hVsf = hmVsf
                            tlAvailInfo.hFsf = hmFsf
                            tlAvailInfo.hAnf = hmAnf
                            gGetSpotCounts tlAvailInfo, lmSAvailsDates(), lmEAvailsDates(), tmAvRdf(), tmRifRate(), tlCntTypes, llIgnoreCodes(), tmAvr(), slUserRequest
                            For ilIndex = 0 To UBound(tmAvr) - 1 Step 1
                                    'build the AVR for each daypart/vehicle so that the proposals can be incorporated into these totals
                                    tlAvrContr(UBound(tlAvrContr)) = tmAvr(ilIndex)
                                    ReDim Preserve tlAvrContr(0 To UBound(tlAvrContr) + 1) As AVR
                            Next ilIndex
                        End If
                    Next ilRcf                      'next rate card (should only be 1)
                End If                              'vpfindex > 0
            End If                                  'vehicle selected
        Next ilVehicle                              'For ilvehicle = 0 To RptSelPC!lbcSelection(0).ListCount - 1
        llDate = gDateValue(slQSDate) + 13 * 7
        slQSDate = Format$(llDate, "m/d/yy")
        llEffDate = llDate
    Next ilQNo                                      'next quarter
    Erase tmAvr
    'gather all contracts whose entered date is equal or prior to the requested date (gather from beginning of std year to
    'input date
    'slCntrTypes = gBuildCntTypes()
    'slCntrStatus = "CI"               'Complete & incomplete props
    'ilHOState = 3                       'get latest orders & revisions   (plus revised orders turned proposals WCI)
    'slStartQtr = Format$(llStartOfReport, "m/d/yy")
    'slEndQtr = Format$(llEndOfReport, "m/d/yy")
    llStart = llStartOfReport
    ilLoop = 1
    Do While llStart < llEndOfReport
        llStartWeeks(ilLoop) = llStart
        llStart = llStart + 7
        ilLoop = ilLoop + 1
    Loop
    llStartWeeks(ilLoop) = llStart              'get the last start date

    'ilret = gObtainCntrForDate(RptSelPC, slStartQtr, slEndQtr, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())


    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
        'project the $
        llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
        'Retrieve the contract, schedule lines and flights
        'llContrCode = gPaceCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llEnterTo, hmChf, tmChf)
        ilFound = False
        If llContrCode > 0 And ((tlChfAdvtExt(ilCurrentRecd).sStatus = "C" And tlCntTypes.iComplete = True) Or (tlChfAdvtExt(ilCurrentRecd).sStatus = "I" And tlCntTypes.iIncomplete = True) Or (tlChfAdvtExt(ilCurrentRecd).sStatus = "W" And tlCntTypes.iWorking = True) Or tlChfAdvtExt(ilCurrentRecd).sStatus = "G" Or tlChfAdvtExt(ilCurrentRecd).sStatus = "N") Then
            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfPC, tgClfPC(), tgCffPC())
            gUnpackDateLong tgChfPC.iOHDDate(0), tgChfPC.iOHDDate(1), llDate
            ilFound = True
            'determine if the contracts start & end dates fall within the requested period
            gUnpackDateLong tgChfPC.iEndDate(0), tgChfPC.iEndDate(1), llEnd      'hdr end date converted to long
            gUnpackDateLong tgChfPC.iStartDate(0), tgChfPC.iStartDate(1), llStart    'hdr start date converted to long

            'ignore if end date of contract prior to start of requested report,
            'ignore if start date of contract is later than end date of report,
             'unscheduled hold or order, dont do this test--it needs to still be scheduled
            If tgChfPC.sStatus = "G" Or tgChfPC.sStatus = "N" Then  'unsched hold or order, need to count it regardless
                If llEnd < llStartOfReport Or llStart >= llEndOfReport Then  'determine if prop is within bounds
                    ilFound = False
                End If
            Else
                'see if the propsal has been selected by user
                ilFound = False
                For ilProp = LBound(tgProcessProp) To UBound(tgProcessProp) - 1
                    If tgProcessProp(ilProp).lCode = tgChfPC.lCode Then
                        ilFound = True
                        Exit For
                    End If
                Next ilProp

                If llEnd < llStartOfReport Or llStart >= llEndOfReport Then  'determine if prop is within bounds
                    ilFound = False
                End If
            End If

            If tgChfPC.sType = "T" And Not tlCntTypes.iRemnant Then
                ilFound = False
            End If
            If tgChfPC.sType = "Q" And Not tlCntTypes.iPI Then
                ilFound = False
            End If
            If tgChfPC.iPctTrade = 100 And Not tlCntTypes.iTrade Then
                ilFound = False
            End If
            If tgChfPC.sType = "M" And Not tlCntTypes.iPromo Then
                ilFound = False
            End If
            If tgChfPC.sType = "S" And Not tlCntTypes.iPSA Then
                ilFound = False
            End If
            '3-16-10 wrong code was tested for standard (tested S not C)
            If tgChfPC.sType = "C" And Not tlCntTypes.iStandard Then      'include Standard types?
                ilFound = False
            End If
            If tgChfPC.sType = "V" And Not tlCntTypes.iReserv Then      'include reservations ?
                ilFound = False
            End If
            If tgChfPC.sType = "R" And Not tlCntTypes.iDR Then      'include DR?
                ilFound = False
            End If
        End If

        If ilFound Then
            For ilClf = LBound(tgClfPC) To UBound(tgClfPC) - 1 Step 1
                tmClf = tgClfPC(ilClf).ClfRec

                If tmClf.sType = "H" Or tmClf.sType = "S" Then
                    'test for selective vehicles
                    For ilVehicle = 0 To RptSelPC!lbcSelection(0).ListCount - 1 Step 1
                        slNameCode = tgCSVNameCode(ilVehicle).sKey
                        ilRet = gParseItem(slNameCode, 1, "\", slName)
                        ilRet = gParseItem(slName, 3, "|", slName)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        ilVefCode = Val(slCode)
                        ilVpfIndex = -1
                        If ilVefCode = tmClf.iVefCode And RptSelPC!lbcSelection(0).Selected(ilVehicle) Then
                            'ilVpfIndex = -1
                            'For ilLoop = 0 To UBound(tgVpf) Step 1
                            '    If ilVefCode = tgVpf(ilLoop).iVefKCode Then
                                ilLoop = gBinarySearchVpf(ilVefCode)
                                If ilLoop <> -1 Then
                                    ilVpfIndex = ilLoop
                            '        Exit For
                                End If
                            'Next ilLoop
                            If ilVpfIndex >= 0 Then
                                tlAvailInfo.iVefCode = ilVefCode           'setup variables sent via structure
                                tlAvailInfo.iVpfIndex = ilVpfIndex
                                For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                                    ilFound = False
                                    tmRcf = tgMRcf(ilRcf)
                                    ilDateOk = False
                                    For ilLoop = 0 To RptSelPC!lbcSelection(1).ListCount - 1 Step 1
                                        slNameCode = tgRateCardCode(ilLoop).sKey
                                        ilRet = gParseItem(slNameCode, 3, "\", slCode)
                                        If Val(slCode) = tgMRcf(ilRcf).iCode Then
                                            If (RptSelPC!lbcSelection(1).Selected(ilLoop)) Then
                                                ilDateOk = True
                                            End If
                                            Exit For
                                        End If
                                    Next ilLoop

                                    If ilDateOk Then
                                        'ReDim tmAvr(0 To 0) As AVR
                                        ReDim tmAvRdf(0 To 0) As RDF
                                        ReDim tmRifRate(0 To 0) As RIF

                                        ilUpper = 0
                                        'Setup the DP with rates to process
                                        For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                                            If tgMRif(llRif).iRcfCode = tgMRcf(ilRcf).iCode Then
                                                'For ilRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                                                ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                                                If ilRdf <> -1 Then
                                                    'determine if rate card items have sort codes, etc; otherwise use DP
                                                    If tgMRif(llRif).sRpt <> "Y" And tgMRif(llRif).sRpt <> "N" Then
                                                        slSaveReport = tgMRdf(ilRdf).sReport
                                                    Else
                                                        slSaveReport = tgMRif(llRif).sRpt
                                                    End If
                                                    If tgMRif(llRif).iSort = 0 Then
                                                        ilSaveSort = tgMRdf(ilRdf).iSortCode
                                                    Else
                                                        ilSaveSort = tgMRif(llRif).iSort
                                                    End If
                                                    'Must be matching DP against Items, printed DP to be shown, active DP, DP vehicle matches line, & DP found matches line DP
                                                    If tgMRdf(ilRdf).iCode = tgMRif(llRif).iRdfCode And slSaveReport = "Y" And tgMRdf(ilRdf).sState <> "D" And tgMRif(llRif).iVefCode = ilVefCode And tmClf.iRdfCode = tgMRdf(ilRdf).iCode Then
                                                        ilFound = False
                                                        For ilLoop = LBound(tmAvRdf) To ilUpper - 1 Step 1
                                                            If tmAvRdf(ilLoop).iCode = tgMRdf(ilRdf).iCode Then
                                                                ilFound = True
                                                                Exit For
                                                            End If
                                                        Next ilLoop
                                                        If Not ilFound Then
                                                            tmAvRdf(ilUpper) = tgMRdf(ilRdf)
                                                            tmRifRate(ilUpper) = tgMRif(llRif)
                                                            tmRifRate(ilUpper).iSort = ilSaveSort
                                                            ilUpper = ilUpper + 1
                                                            ReDim Preserve tmAvRdf(0 To ilUpper) As RDF
                                                            ReDim Preserve tmRifRate(0 To ilUpper) As RIF
                                                            ilFound = True
                                                        End If
                                                '        Exit For
                                                    End If
                                                End If
                                                'Next ilRdf
                                            End If
                                            'Use only the matching daypart that the line was sold to
                                            If ilFound Then
                                                Exit For
                                            End If
                                        Next llRif
                                    End If
                                    If ilFound Then     'found a matching r/c card and daypart.  get the spot counts for htis line
                                        For ilLoop = 1 To 53
                                            llProject(ilLoop) = 0                'init bkts to accum qtr $ for this line
                                        Next ilLoop
                                        gBuildFlightSpots ilClf, llStartWeeks(), 1, ilNoQtrs * 13 + 1, llProject(), 2, tgClfPC(), tgCffPC()       'project the # spots for this lines dp definition
                                        tlAvailInfo.lSDate = llStartOfReport
                                        mGetProposalCounts tlAvailInfo, tmAvRdf(), tmRifRate(), tlCntTypes, ilNoQtrs, llProject(), tlAvrContr()
                                        Exit For
                                    End If
                                Next ilRcf          'only continue with next r/c if not a matching card with user requested one
                                Exit For
                            End If
                        End If                      'tmclf.ivefcode <> selected vehicle
                        If ilVpfIndex >= 0 Then
                            Exit For                'exit the vehicle loop, line was processed
                        End If
                    Next ilVehicle
                End If
            Next ilClf                      'process nextline
        End If                              'ilfound - llAdjust falls within requested dates
    Next ilCurrentRecd

    'gGetVehGrpSets ilVefcode, ilMinorSet, ilMajorSet, ilmnfMinorCode, ilmnfMajorCode
    'Output records
    For ilRec = 0 To UBound(tlAvrContr) - 1 Step 1
        gGetVehGrpSets tlAvrContr(ilRec).iVefCode, ilMinorSet, ilMajorSet, ilmnfMinorCode, ilMnfMajorCode
        tlAvrContr(ilRec).imnfMajorCode = ilMnfMajorCode
        If tlAvrContr(ilRec).iRdfSortCode = -1 Then      'this is the orphaned missed group,
            tlAvrContr(ilRec).iRdfSortCode = 0           'Crystal will test for absence of code to detect
        End If

        ilRet = btrInsert(hmAvr, tlAvrContr(ilRec), imAvrRecLen, INDEXKEY0)
    Next ilRec
    sgCntrForDateStamp = ""
    Erase tlChfAdvtExt
    Erase tmAvRdf, tmRifRate, tmAvr, llProject, llStartWeeks, tlAvrContr

    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmAvr)
    ilRet = btrClose(hmLcf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmAvr)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmMnf)
    btrDestroy hmSdf
    btrDestroy hmVef
    btrDestroy hmSsf
    btrDestroy hmLcf
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmAvr
    btrDestroy hmCHF
    btrDestroy hmMnf
    Exit Sub
End Sub
'
'
'           gGetProposalCounts - Spot counts have been gathered for
'           the flights of the line. Place spot counts into the
'           buckets for the avails by week.
'
'           5/24/99  (extracted from gGetSpot Counts)
'           8/13/99 Hold & Reserve counts are not tested.  All Proposals
'                   are included in Proposal bucket regardless of type
'
'Sub gGetProposalCounts (ilVpfIndex As Integer, ilFirstQ As Integer, llStartDate As Long, tlAvRdf() As RDF, tlRifRate() As RIF, tlCntTypes As CNTTYPES, ilNoQTrs As Integer, llProject() As Long, tlAvr() As AVR)
Sub mGetProposalCounts(tlAvailInfo As AVAILCOUNT, tlAvRdf() As RDF, tlRifRate() As RIF, tlCntTypes As CNTTYPES, ilNoQtrs As Integer, llProject() As Long, tlAvr() As AVR)
Dim ilNo30 As Integer
Dim ilNo60 As Integer
Dim ilLen As Integer
Dim slBucketType As String
Dim ilRecIndex As Integer            'week array index
Dim ilBucketIndex As Integer
Dim ilBucketIndexMinusOne As Integer
Dim ilAdjAdd As Integer
Dim ilAdjSub As Integer
Dim llLoopDate As Long
Dim ilOrphanMax As Integer
Dim ilOrphanFound As Integer
Dim ilOrphanMissedLoop As Integer
Dim ilAvailOk As Integer
Dim ilFound As Integer
Dim ilDay As Integer
Dim ilSaveDay As Integer
Dim ilRec As Integer
Dim ilRdf As Integer
Dim llStartQtr As Long
Dim llEndQtr As Long
Dim ilQtrLoop As Integer
Dim ilLoop As Integer
Dim ilVpfIndex As Integer
Dim ilFirstQ As Integer
Dim llStartDate As Long
ReDim ilStartQtr(0 To 1) As Integer
Dim ilTrueWeekInx As Integer              '1-24-08 true week index from 52 week bucket
'ReDim ilRdfCodes(0 To 1) As Integer    '8/4/99
'ReDim tmAvr(0 To 0) As AVR

    ilVpfIndex = tlAvailInfo.iVpfIndex
    ilFirstQ = tlAvailInfo.iFirstBkt
    llStartDate = tlAvailInfo.lSDate
    ilOrphanMax = 1                                     'do not show orphans on separate line (ignore)
    If tlCntTypes.iOrphan Then                          'place orphans on separate line?, (dont fall within any shown daypart)
        ilOrphanMax = 2
    End If

    slBucketType = tlCntTypes.sAvailType                'Avails, Sellout, Inventory flag
    If tgVpf(ilVpfIndex).sSSellOut = "B" Then           'if units & seconds - add 2 to 30 sec unit and take away 1 fro 60
        ilAdjAdd = 2
        ilAdjSub = 1
    ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then       'if units only - take 1 away from 60 count and add 1 to 30 count
        ilAdjAdd = 1
        ilAdjSub = 1
    End If
    For ilQtrLoop = 1 To ilNoQtrs                            'repeat for possible 4 quarters requested
        llStartQtr = llStartDate + ((ilQtrLoop - 1) * 91)            'start of quarter to be processed
        llEndQtr = llStartQtr + 90
        For llLoopDate = llStartQtr To llEndQtr Step 7
            ilBucketIndex = (llLoopDate - llStartQtr) / 7 + 1        'week to process
            ilBucketIndexMinusOne = ilBucketIndex - 1
            ilTrueWeekInx = (llLoopDate - llStartDate) / 7 + 1
            ilDay = gWeekDayLong(llLoopDate)
            'convert to btrieve to store in AVR record
            gPackDateLong llStartQtr, ilStartQtr(0), ilStartQtr(1)
            For ilOrphanMissedLoop = 1 To ilOrphanMax
                ilOrphanFound = False
                ilAvailOk = False
                For ilRdf = LBound(tlAvRdf) To UBound(tlAvRdf) - 1 Step 1
                    ilAvailOk = False
                    If tlAvRdf(ilRdf).iCode = tmClf.iRdfCode Then
                        ilAvailOk = True
                        If ilAvailOk Or ilOrphanMissedLoop = 2 Then        'if pass 2, OK to use this dp if not showing orphans on same line
                            'if pass 1 and showing orphans on separate line, the dayparts must match the sched line ordered
                            If ilOrphanMissedLoop = 1 And tlCntTypes.iOrphan Then
                                If tmClf.iRdfCode <> tlAvRdf(ilRdf).iCode Then
                                    ilAvailOk = False
                                End If
                            End If
                            If ilAvailOk Then
                                ilOrphanFound = True
                                'Determine if Avr created
                                ilFound = False
                                ilSaveDay = ilDay
                                If tlCntTypes.iDayOption = 0 Then              'daypart option, place all values in same record
                                                                                    'to get better availability
                                    ilDay = 0                                       'force all data in same day of week
                                End If

                                '8/4/99 remove code to test for orphans--wont have any for proposals; anything not
                                'matching aDP will be ignored
                                'If ilOrphanMissedLoop = 2 Then
                                '    For ilRec = 0 To UBound(tlAvr) - 1 Step 1
                                '    If (ilRdfCodes(ilRec) = -1) And (tlAvr(ilRec).iFirstBucket = ilFirstQ) And (tlAvr(ilRec).iDay = ilDay) Then
                                '        ilFound = True
                                '        ilRecIndex = ilRec
                                '        Exit For
                                '    End If
                                '    Next ilRec
                                'Else

                                '5-29-14 wrong index tested resulting in many bad results for proposal totals
                                For ilRec = 0 To UBound(tlAvr) - 1 Step 1
                                    If (tlAvr(ilRec).iRdfSortCode = tlAvRdf(ilRdf).iCode) And (tlAvr(ilRec).iFirstBucket = ilFirstQ) And (tlAvr(ilRec).iDay = ilDay) And tlAvr(ilRec).iVefCode = tmClf.iVefCode And tlAvr(ilRec).iQStartDate(0) = ilStartQtr(0) And tlAvr(ilRec).iQStartDate(1) = ilStartQtr(1) Then
                                        ilFound = True
                                        ilRecIndex = ilRec
                                        Exit For
                                    End If
                                Next ilRec
                                'End If

                                If Not ilFound Then
                                    ilRecIndex = UBound(tlAvr)
                                    tlAvr(ilRecIndex).iGenDate(0) = igNowDate(0)
                                    tlAvr(ilRecIndex).iGenDate(1) = igNowDate(1)
                                    'tlAvr(ilRecIndex).iGenTime(0) = igNowTime(0)
                                    'tlAvr(ilRecIndex).iGenTime(1) = igNowTime(1)
                                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                    tlAvr(ilRecIndex).lGenTime = lgNowTime
                                    tlAvr(ilRecIndex).iDay = ilDay
                                    tlAvr(ilRecIndex).iQStartDate(0) = ilStartQtr(0)
                                    tlAvr(ilRecIndex).iQStartDate(1) = ilStartQtr(1)
                                    tlAvr(ilRecIndex).iFirstBucket = ilFirstQ
                                    tlAvr(ilRecIndex).sBucketType = slBucketType
                                    tlAvr(ilRecIndex).iDPStartTime(0) = tlAvRdf(ilRdf).iStartTime(0, 6) '7)
                                    tlAvr(ilRecIndex).iDPStartTime(1) = tlAvRdf(ilRdf).iStartTime(1, 6) '7)
                                    tlAvr(ilRecIndex).iDPEndTime(0) = tlAvRdf(ilRdf).iEndTime(0, 6) '7)
                                    tlAvr(ilRecIndex).iDPEndTime(1) = tlAvRdf(ilRdf).iEndTime(1, 6) '7)
                                    'tlAvr(ilRecIndex).sDPDays = slDays
                                    tlAvr(ilRecIndex).sNot30Or60 = "N"

                                    tlAvr(ilRecIndex).iVefCode = tmClf.iVefCode
                                    'tlAvr(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iCode           'DP code

                                    tlAvr(ilRecIndex).iRdfCode = tlRifRate(ilRdf).iSort   'DP Sort code  from RIF
                                    tlAvr(ilRecIndex).iRdfSortCode = tlAvRdf(ilRdf).iCode   'DP code to retrieve DP name description
                                    'ilRdfCodes(ilRecIndex) = tlAvRdf(ilRdf).icode
                                    tlAvr(ilRecIndex).sInOut = tlAvRdf(ilRdf).sInOut
                                    tlAvr(ilRecIndex).ianfCode = tlAvRdf(ilRdf).ianfCode

                                    '8/4/99 dont need orphan testing here-- they are ignored for proposals
                                    'If ilOrphanMissedLoop = 2 Then
                                    '    'override some of the codes if its in the orphan pass (where no shown DP equals the DP of the missed spot)
                                    '    ilRdfCodes(ilRecIndex) = -1         'phoney daypart for orphaned missed spots 'tmAvRdf(ilRdf).icode
                                    '    tlAvr(ilRecIndex).irdfCode = 32000     'sort it last tmRifSorts(ilRdf).isort   'DP Sort code  from RIF
                                    '    tlAvr(ilRecIndex).iRdfSortCode = -1 'tmAvRdf(ilRdf).icode   'DP code to retrieve DP name description
                                    'End If

                                    ReDim Preserve tlAvr(0 To ilRecIndex + 1) As AVR
                                    'ReDim Preserve ilRdfCodes(0 To ilRecIndex + 1)
                                End If
                                tlAvr(ilRecIndex).lRate(ilBucketIndexMinusOne) = tlRifRate(ilRdf).lRate(ilBucketIndex)
                                ilDay = ilSaveDay
                                ilNo30 = 0
                                ilNo60 = 0
                                ilLen = tmClf.iLen
                                If tgVpf(ilVpfIndex).sSSellOut = "B" Then           'use units & seconds
                                'Convert inventory to number of 30's and 60's
                                     If ilLen >= 60 Then
                                        ilNo60 = llProject(ilTrueWeekInx)   '1-24-08(ilBucketIndex)   '4-26-02 take out multiplying * 2
                                    Else             'any length under 60 gets counted as 1-30" unit

                                        ilNo30 = llProject(ilTrueWeekInx)   '1-24-08 (ilBucketIndex)      '30 second units
                                    End If
                                    If (slBucketType = "S") Or (slBucketType = "P") Then    'sellout or %sellout, accum sold
                                        If tlCntTypes.iDetail Then                        'qtrly detail report (has detail for sch lines)
                                            '2-14-03
                                            'tlAvr(ilRecIndex).i30Count(ilBucketIndex) = tlAvr(ilRecIndex).i30Count(ilBucketIndex) + ilNo30
                                            'tlAvr(ilRecIndex).i60Count(ilBucketIndex) = tlAvr(ilRecIndex).i60Count(ilBucketIndex) + ilNo60

                                            If tgChfPC.iCntRevNo > 0 And tgChfPC.sStatus = "W" Then 'ignore working revs as a real proposal
                                            Else
                                                tlAvr(ilRecIndex).i30Prop(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Prop(ilBucketIndexMinusOne) + ilNo30
                                                tlAvr(ilRecIndex).i60Prop(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Prop(ilBucketIndexMinusOne) + ilNo60
                                            End If

                                        Else                    'not qtrly detail
                                            tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                            tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                        End If
                                    Else                    'summary version
                                        tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo60
                                        tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) - ilNo30
                                    End If
                                    'adjust the available buckets (used for qtrly detail  report only)
                                    tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                    tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                    'Count 30 or 60 and set flag if neither
                                    If ilLen = 60 Then
                                        ilNo60 = llProject(ilTrueWeekInx)   '1-24-08 (ilBucketIndex)
                                    ElseIf ilLen <= 30 Then
                                        ilNo30 = llProject(ilTrueWeekInx)    '1-24-08 (ilBucketIndex)
                                    Else
                                        tlAvr(ilRecIndex).sNot30Or60 = "Y"
                                        If ilLen <= 30 Then
                                            ilNo30 = llProject(ilTrueWeekInx)   '1-24-08(ilBucketIndex)
                                        Else
                                            ilNo60 = llProject(ilTrueWeekInx)  '1-24-08(ilBucketIndex)
                                        End If
                                    End If
                                    If (ilNo60 <> 0) Or (ilNo30 <> 0) Then
                                        If (slBucketType = "S") Or (slBucketType = "P") Then    'Sellout or Percent option (vs Avails)
                                            If tlCntTypes.iDetail Then                        'qtrly detail spots option
                                                tlAvr(ilRecIndex).i30Prop(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Prop(ilBucketIndexMinusOne) + ilNo30
                                                tlAvr(ilRecIndex).i60Prop(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Prop(ilBucketIndexMinusOne) + ilNo60
                                            Else
                                                tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                            End If
                                        Else
                                            If ilNo60 > 0 Then                     'spot found a 60?
                                                tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo60
                                                'tlAvr(ilRecIndex).i60Prop(ilBucketIndex) = tlAvr(ilRecIndex).i60Prop(ilBucketIndex) - ilNo60
                                            Else
                                                If tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) > 0 Then
                                                    tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) - ilNo30
                                                Else
                                                    If tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) > 0 Then
                                                        tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo30
                                                    Else                        'oversold units
                                                        tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) - ilNo30
                                                    End If
                                                End If
                                                'If tlAvr(ilRecIndex).i30Prop(ilBucketIndex) > 0 Then
                                                '    tlAvr(ilRecIndex).i30Prop(ilBucketIndex) = tlAvr(ilRecIndex).i30Prop(ilBucketIndex) - ilNo30
                                                'Else
                                                '    If tlAvr(ilRecIndex).i60Prop(ilBucketIndex) > 0 Then
                                                '        tlAvr(ilRecIndex).i60Prop(ilBucketIndex) = tlAvr(ilRecIndex).i60Prop(ilBucketIndex) - ilNo30
                                                '    Else                        'oversold units
                                                '        tlAvr(ilRecIndex).i30Prop(ilBucketIndex) = tlAvr(ilRecIndex).i30Prop(ilBucketIndex) - ilNo30
                                                '    End If
                                                'End If
                                            End If
                                        End If
                                    End If
                                    'adjust the available buckets (used for qtrly detail report only)
                                    tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                    tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                    'Count 30 or 60 and set flag if neither
                                    If ilLen = 60 Then
                                        ilNo60 = llProject(ilTrueWeekInx)   '1-24-08(ilBucketIndex)
                                    ElseIf ilLen = 30 Then
                                        ilNo30 = llProject(ilTrueWeekInx)   '1-24-08(ilBucketIndex)
                                    Else
                                        tlAvr(ilRecIndex).sNot30Or60 = "Y"
                                    End If
                                    If (slBucketType = "S") Or (slBucketType = "P") Then        'if Sellout or % sellout, accum the seconds sold
                                        'Qtrly detail has been forced to "Sellout" for internal testing
                                        If tlCntTypes.iDetail Then                    'qtrly detail booked has more options

                                            tlAvr(ilRecIndex).i30Prop(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Prop(ilBucketIndexMinusOne) + ilNo30
                                            tlAvr(ilRecIndex).i60Prop(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Prop(ilBucketIndexMinusOne) + ilNo60
                                        Else
                                            tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                            tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                        End If
                                    Else                                                    'holds & reserve n/a for othr qtrly summary options
                                        tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo60
                                        tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) - ilNo30
                                    End If
                                    'adjust the available bucket (used for qrtrly detail report only)
                                    tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                    tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                End If
                                If ilAvailOk Then
                                    Exit For                'force exit on this missed if found a matching daypart
                                End If
                            End If                          'ilAvailOK
                            If ilOrphanMissedLoop = 2 Then
                                Exit For
                            End If
                        End If
                    End If                      'tmclf.icode = tlAvrdf
                Next ilRdf
                If ilOrphanFound Then
                    Exit For
                End If
            Next ilOrphanMissedLoop
        Next llLoopDate
    Next ilQtrLoop
    'Adjust counts
    If (slBucketType = "A" And tgVpf(ilVpfIndex).sSSellOut = "B") Then
        For ilRec = 0 To UBound(tlAvr) - 1 Step 1
            'For ilLoop = 1 To 13 Step 1
            For ilLoop = LBound(tlAvr(ilRec).i30Count) To UBound(tlAvr(ilRec).i30Count) Step 1
                If tlAvr(ilRec).i30Count(ilLoop) < 0 Then
                    Do While (tlAvr(ilRec).i60Count(ilLoop) > 0) And (tlAvr(ilRec).i30Count(ilLoop) < 0)
                        tlAvr(ilRec).i60Count(ilLoop) = tlAvr(ilRec).i60Count(ilLoop) - ilAdjSub    '1
                        tlAvr(ilRec).i30Count(ilLoop) = tlAvr(ilRec).i30Count(ilLoop) + ilAdjAdd    '2
                    Loop
                ElseIf (tlAvr(ilRec).i60Count(ilLoop) < 0) Then
                End If
            Next ilLoop
            'For ilLoop = 1 To 13 Step 1
            For ilLoop = LBound(tlAvr(ilRec).i30Count) To UBound(tlAvr(ilRec).i30Count) Step 1
                If tlAvr(ilRec).i30Prop(ilLoop) < 0 Then
                    Do While (tlAvr(ilRec).i60Prop(ilLoop) > 0) And (tlAvr(ilRec).i30Prop(ilLoop) < 0)
                        tlAvr(ilRec).i60Prop(ilLoop) = tlAvr(ilRec).i60Prop(ilLoop) - ilAdjSub    '1
                        tlAvr(ilRec).i30Prop(ilLoop) = tlAvr(ilRec).i30Prop(ilLoop) + ilAdjAdd    '2
                    Loop
                ElseIf (tlAvr(ilRec).i60Prop(ilLoop) < 0) Then
                End If
            Next ilLoop
        Next ilRec
    End If
    'Adjust counts for qtrly detail availbilty
    If (tgVpf(ilVpfIndex).sSSellOut = "B") Then
        For ilRec = 0 To UBound(tlAvr) - 1 Step 1
            'For ilLoop = 1 To 13 Step 1
            For ilLoop = LBound(tlAvr(ilRec).i30Avail) To UBound(tlAvr(ilRec).i30Avail) Step 1
                If tlAvr(ilRec).i30Avail(ilLoop) < 0 Then
                    Do While (tlAvr(ilRec).i60Avail(ilLoop) > 0) And (tlAvr(ilRec).i30Avail(ilLoop) < 0)
                        tlAvr(ilRec).i60Avail(ilLoop) = tlAvr(ilRec).i60Avail(ilLoop) - 1
                        tlAvr(ilRec).i30Avail(ilLoop) = tlAvr(ilRec).i30Avail(ilLoop) + 2
                    Loop
                ElseIf (tlAvr(ilRec).i60Avail(ilLoop) < 0) Then
                End If
            Next ilLoop
        Next ilRec
    End If
End Sub
