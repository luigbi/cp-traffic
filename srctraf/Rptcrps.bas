Attribute VB_Name = "RPTCRPS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrps.bas on Wed 6/17/09 @ 12:56 PM
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
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim tlMMnf() As MNF                    'array of MNF records for specific type
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim tlGrf() As GRF
Dim tmBvf As BVF                  'Budgets by office & vehicle
Dim hmBvf As Integer
Dim tmBvfSrchKey As BVFKEY0       'Gen date and time
Dim imBvfRecLen As Integer        'BVF record length
Dim tmBvfPlan() As BVF             'Budget plan
Dim tmBvfFC() As BVF               'Budget  forecast
Dim tmPjf As PJF                  'Slsp Projections
Dim hmPjf As Integer
Dim imPjfRecLen As Integer        'PJF record length
Dim hmSbf As Integer
Dim tmMnfNtr() As MNF
Const NOT_SELECTED = 0

Const LBONE = 1

'  Receivables File
Dim tmRvf As RVF            'RVF record image
'*******************************************************
'*                                                     *
'*      Procedure Name:mReadBvfRec                     *
'*                                                     *
'*             Created:6/28/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Read a record                  *
'*            (Taken from RCImpact.bas and modified    *
'*             to build budgets by selling office)     *
'*******************************************************
Private Function mReadBvfRec(hlBvf As Integer, ilMnfCode As Integer, ilYear As Integer, tlBvfSof() As BVF) As Integer
'
'   iRet = mReadBvfRec (hlBvf As Integer, iMnfCode as integer, ilYear as integer)
'   Where:
'       ilMnfCode(I)-Budget Name Code
'       ilYears(I)-Year to retrieve
'       tlBvfSof(O) - array of Budget records by selling office
'       iRet (O)- True if record read,
'                 False if not read
'
    Dim ilRet As Integer    'Return status
    Dim ilUpper As Integer
    Dim ilLoop As Integer
    Dim ilIndex As Integer
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    Dim ilFound As Integer
    Dim tlIntTypeBuff As POPINTEGERTYPE   'Type field record

    'ReDim tlBvfSof(1 To 1) As BVF
    ReDim tlBvfSof(0 To 0) As BVF

    ilUpper = UBound(tlBvfSof)
    btrExtClear hlBvf   'Clear any previous extend operation
    'ilExtLen = Len(tlBvfSof(1))  'Extract operation record size
    ilExtLen = Len(tlBvfSof(0))  'Extract operation record size
    imBvfRecLen = Len(tmBvf)
    tmBvfSrchKey.iYear = ilYear
    tmBvfSrchKey.iSeqNo = 1
    tmBvfSrchKey.iMnfBudget = ilMnfCode
    ilRet = btrGetGreaterOrEqual(hlBvf, tmBvf, imBvfRecLen, tmBvfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    'ilRet = btrGetFirst(hlBvf, tgBvfRec(1).tBvf, imBvfRecLen, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    If ilRet <> BTRV_ERR_END_OF_FILE Then
        llNoRec = gExtNoRec(ilExtLen) 'btrRecords(hlAdf) 'Obtain number of records
        Call btrExtSetBounds(hlBvf, llNoRec, -1, "UC", "BVF", "") '"EG") 'Set extract limits (all records)
        ilOffSet = gFieldOffset("Bvf", "BvfMnfBudget")
        tlIntTypeBuff.iType = ilMnfCode
        ilRet = btrExtAddLogicConst(hlBvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
        'On Error GoTo mReadBvfRecErr
        'gBtrvErrorMsg ilRet, "mReadBvfRec (btrExtAddLogicConst):" & "Bvf.Btr", RCImpact
        'On Error GoTo 0
        ilOffSet = gFieldOffset("Bvf", "BvfYear")
        tlIntTypeBuff.iType = ilYear
        ilRet = btrExtAddLogicConst(hlBvf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlIntTypeBuff, 2)
        ilRet = btrExtAddField(hlBvf, 0, ilExtLen) 'Extract the whole record
        ilRet = btrExtGetNext(hlBvf, tmBvf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            ilExtLen = Len(tmBvf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlBvf, tmBvf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                ilFound = False
                'For ilLoop = 1 To ilUpper - 1 Step 1
                For ilLoop = 0 To ilUpper - 1 Step 1
                    If tlBvfSof(ilLoop).iSofCode = tmBvf.iSofCode Then
                        For ilIndex = LBound(tmBvf.lGross) To UBound(tmBvf.lGross) Step 1
                            tlBvfSof(ilLoop).lGross(ilIndex) = tlBvfSof(ilLoop).lGross(ilIndex) + tmBvf.lGross(ilIndex)
                        Next ilIndex
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If Not ilFound Then
                    tlBvfSof(ilUpper) = tmBvf
                    ilUpper = ilUpper + 1
                    'ReDim Preserve tlBvfSof(1 To ilUpper) As BVF
                    ReDim Preserve tlBvfSof(0 To ilUpper) As BVF
                End If

                ilRet = btrExtGetNext(hlBvf, tmBvf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlBvf, tmBvf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    mReadBvfRec = True
    Exit Function

    On Error GoTo 0
    mReadBvfRec = False
    Exit Function
End Function
'
'
'                   Create Salesperson Projection Scenario prepass file
'                   Generate GRF file by vehicle.  Each record  contains the vehicle,
'                   plan $, Forecast $, Business on Books (OOB) for current years Qtr,
'                   slsp projection $ unadjusted (pjf),Most likely projection Qtr $ (adj)
'                   pessimistic projection Qtr $ (adjusted), and Optimistic Qtr $ (adj)
'
'                   Rollover date is determined by the user input date.  All slsp
'                   are checked trying to find a date equal or greater (and within
'                   one week of the entered date) to the entered date.  The closest
'                   date found from all slsp against the user entered date is used
'                   to find all matching rollover records.
'
'                   <input>  llCurrStart - start of quarter
'                            llCurrEnd - end of quarter
'
'                   4/10/98 Exclude 100% trade contracts
'       2-11-05 Chg array of PHF/RVF array from integer to long (prevent overflow)
'       12-14-06 add parm to gObtainRvfPhf to test on tran date (vs entry date)

Sub gCrScenario()
Dim slMnfStamp As String
Dim slAirOrder As String * 1                'from site pref - bill as air or ordered
Dim ilPotnInx As Integer                    'index to ilLikePct (which % to use)
Dim slPotn As String * 3                    'Order of potential codes
Dim ilLoop As Integer
Dim ilTemp As Integer
Dim ilSlsLoop As Integer
Dim ilRet As Integer
ReDim ilRODate(0 To 1) As Integer           'Effective Date to match retrieval of Projection record
Dim llClosestDate As Long                   'Closest date to rollover date entered
Dim slDate As String
Dim slStr As String
Dim ilMonth As Integer
Dim ilYear As Integer
Dim llEnterTo As Long                         'date used to test contracts (any date equal/prior to date entered is used)
'ReDim llProject(1 To 3) As Long               'projected $, for 3 months (qtr)
ReDim llProject(0 To 3) As Long               'projected $, for 3 months (qtr). Index zero ignored
Dim llTYGetFrom As Long                       'range of dates for contract access
Dim llTYGetTo As Long
ReDim ilEnterDate(0 To 1) As Integer            'enterd date, used to show effective date on report
'ReDim llTYDates(1 To 4) As Long
ReDim llTYDates(0 To 4) As Long             'Index zero ignored
Dim slStartDate As String                       'llLYGetFrom or llTYGetFrom converted to string
Dim slEndDate As String                         'llLYGetTo or LLTYGetTo converted to string
Dim ilBdMnfPlan As Integer                      'budget name Planto get
Dim ilBdMnfFC As Integer                         'budget name Foecast
Dim ilBdPlanYear As Integer                       'budget year plan to get
Dim ilBdFCYear As Integer                         'budget year forecast to get
Dim slNameCode As String
Dim slYear As String
Dim slCode As String
Dim slCntrTypes As String                       'valid contract types to access
Dim slCntrStatus As String                      'valid status (holds, orders, working, etc) to access
Dim ilHOState As Integer
Dim ilFound As Integer
Dim ilStartWk As Integer                        'starting week index to gather budget data
Dim ilEndWk As Integer                          'ending week index to gather budgets
Dim ilFirstWk As Integer                        'true if week 0 needs to be added when start wk = 1
Dim ilLastWk As Integer                         'true if week 53 needs to be added when end wk = 52
Dim llContrCode As Long                         'contr code from gObtainCntrforDate
Dim ilCurrentRecd As Integer
Dim ilClf As Integer
Dim llAdjust As Long                            'Adjusted gross using the potential codes most likely %
Dim ilCorpStd As Integer                        '1 = corp, 2 = std
Dim ilSofCode As Integer
Dim llTemp As Long
Dim ilBvfCalType As Integer                 '(search type for budget file) 0=std, 1 = reg, 2 & 3 = julian, 4= corp jan thru dec, 5 = corp fiscal
Dim ilPjfCalType As Integer                 'same as above, search type for projection file
Dim slTYStartYr As String                     'start date of year requested
Dim slTYEndYr As String                       'end date of year requested
Dim llProjYearEnd As Long                       'projection recds standard end date of year
Dim llProjYearStart As Long                     'projection recd stdard start date of year
Dim llWeek As Long
Dim tlTranType As TRANTYPES
'ReDim tlRvf(1 To 1) As RVF
ReDim tlRvf(0 To 0) As RVF
Dim llRvfLoop As Long                       '2-11-05
Dim blIncludeNTR As Boolean                 '7-15-08 Dan M allow ntr/hard cost in report
Dim blIncludeHardCost As Boolean
Dim blNTRWithTotal As Boolean
Dim tlNTRInfo() As NTRPacing
Dim ilLowerboundNTR As Integer
Dim ilUpperboundNTR As Integer
Dim ilNTRCounter As Integer
Dim llSingleContract As Long                '7-15-08 single contract selectivity
Dim llDateEntered As Long               'receivables entered date for pacing test
Dim blFailedMatchNtrOrHardCost As Boolean
Dim blFailedBecauseInstallment As Boolean
Dim slTemp As String

If Val(RptSelPS!edcSelC2.Text) > 0 And RptSelPS!edcSelC2.Text <> " " Then
     llSingleContract = Val(RptSelPS!edcSelC2.Text)
Else
    llSingleContract = NOT_SELECTED
End If

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

    hmPjf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmPjf, "", sgDBPath & "Pjf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmBvf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmPjf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imPjfRecLen = Len(tmPjf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmBvf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmMnf
        btrDestroy hmPjf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmPjf)
        ilRet = btrClose(hmBvf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSlf
        btrDestroy hmMnf
        btrDestroy hmPjf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imSlfRecLen = Len(tmSlf)
     ' Dan M 7-14-08
    hmSbf = CBtrvTable(ONEHANDLE)
    ilRet = btrOpen(hmSbf, "", sgDBPath & "sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSbf
        btrDestroy hmSlf
        btrDestroy hmMnf
        btrDestroy hmPjf
        btrDestroy hmBvf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
        'Setup start and end dates of contracts to obtain (this is testing the contract headers start & end date range
    tlTranType.iAdj = True              'look only for adjustments in the History & Rec files
    tlTranType.iInv = False
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = False
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iNTR = False         '9-17-02

    If RptSelPS!ckcSelCInclude(0).Value Or RptSelPS!ckcSelCInclude(1).Value Then    'don't waste time filling array if don't need.
        tlTranType.iNTR = True
        ReDim tmMnfNtr(0 To 0) As MNF
   'set flags ntr or hard cost or both chosen
        If RptSelPS!ckcSelCInclude(0).Value = 1 Then
             blIncludeNTR = True
        End If
        If RptSelPS!ckcSelCInclude(1).Value = 1 Then
            blIncludeHardCost = True
        End If
        ilRet = gObtainMnfForType("I", "", tmMnfNtr())
        If ilRet <> True Then
            MsgBox "error retrieving MNF files", vbOKOnly + vbCritical
            Exit Sub
        End If
    End If

    slAirOrder = tgSpf.sInvAirOrder     'inv all contracts as aired or ordered
    ilCorpStd = 2                       'force standard
    If RptSelPS!rbcSelCSelect(0).Value Then     'corporate selected
        ilCorpStd = 1
        ilPjfCalType = 4                'get projetions based on std
        ilBvfCalType = 5                'get budgets based on corp
    Else
        ilCorpStd = 2                       'force standard
        ilPjfCalType = 0                   'retrieve both projections & budgets
        ilBvfCalType = 0                   'based on std months
    End If
    'ReDim tlMMnf(1 To 1) As MNF
    ReDim tlMMnf(0 To 0) As MNF
    'get all the Potential codes from MNF
    ilRet = gObtainMnfForType("P", slMnfStamp, tlMMnf())
    'ReDim tlPotn(1 To 1) As ADJUSTLIST          'arry of potential codes and their percentages
    ReDim tlPotn(0 To 1) As ADJUSTLIST          'arry of potential codes and their percentages. Index zero ignored
    slPotn = ""
    'For ilLoop = 1 To UBound(tlMMnf) - 1 Step 1
    For ilLoop = LBound(tlMMnf) To UBound(tlMMnf) - 1 Step 1
        'Bypass any potential codes that isnt an A, B or C
        If Trim$(tlMMnf(ilLoop).sName) = "A" Or Trim$(tlMMnf(ilLoop).sName) = "B" Or Trim$(tlMMnf(ilLoop).sName) = "C" Then
            For ilSlsLoop = LBONE To UBound(tlPotn) - 1 Step 1
                If tlMMnf(ilLoop).iCode = tlPotn(ilSlsLoop).iVefCode Then    'see if this potn code has been created in mem yet
                    ilFound = True
                    Exit For
                End If
            Next ilSlsLoop
            If Not ilFound Then
                ilSlsLoop = UBound(tlPotn)
                tlPotn(ilSlsLoop).iVefCode = tlMMnf(ilLoop).iCode
                tlPotn(ilSlsLoop).lProject(1) = Val(tlMMnf(ilLoop).sUnitType)            'most likely percentage
                gPDNToLong tlMMnf(ilLoop).sRPU, tlPotn(ilSlsLoop).lProject(2)          'optimistc percentage
                tlPotn(ilSlsLoop).lProject(2) = tlPotn(ilSlsLoop).lProject(2) \ 100
                gPDNToLong tlMMnf(ilLoop).sSSComm, tlPotn(ilSlsLoop).lProject(3)           'pessimistic percentage
                tlPotn(ilSlsLoop).lProject(3) = tlPotn(ilSlsLoop).lProject(3) \ 10000
                'ReDim Preserve tlPotn(1 To UBound(tlPotn) + 1)
                ReDim Preserve tlPotn(0 To UBound(tlPotn) + 1)
                slPotn = Trim$(slPotn) & Trim$(tlMMnf(ilLoop).sName)
            End If
        End If
    Next ilLoop
    'Build  slsp in memory to avoid rereading for office
    'ReDim tlSlf(1 To 1) As SLF
    ReDim tlSlf(0 To 0) As SLF
    ilRet = gObtainSlf(RptSelPS, hmSlf, tlSlf())
    'get all the dates needed to work with
    slDate = RptSelPS!edcSelCFrom.Text               'effective date entred
    'obtain the entered dates year based on the std month
    llTYGetTo = gDateValue(slDate)                     'gather contracts thru this date
    gPackDateLong llTYGetTo, ilEnterDate(0), ilEnterDate(1)    'get btrieve date format for entered to pass to record to show on hdr
    'setup Projection rollover date
    'gPackDate slDate, ilRODate(0), ilRODate(1)
    gGetRollOverDate RptSelPS, 2, slDate, llClosestDate   'send the lbcselection index to search, plust rollover date
    gPackDateLong llClosestDate, ilRODate(0), ilRODate(1)

    slStr = gObtainEndStd(Format$(llTYGetTo, "m/d/yy"))
    gObtainMonthYear 0, slDate, ilMonth, ilYear           'get year  of effective date (to figure out the beginning of std year)
    slStr = "1/15/" & Trim$(str$(ilYear))                 'Jan of std year effective dat entered
    If ilCorpStd = 1 Then
        ilYear = Val(RptSelPS!edcSelCTo.Text)           'year requested
        ilRet = gGetCorpCalIndex(ilYear)
        'gUnpackDate tgMCof(ilRet).iStartDate(0, 1), tgMCof(ilRet).iStartDate(1, 1), slStr
        gUnpackDate tgMCof(ilRet).iStartDate(0, 0), tgMCof(ilRet).iStartDate(1, 0), slStr
        llTYGetFrom = gDateValue(gObtainStartCorp(slStr, True))  'gather contracts from this date thru effective entered date
        'Determine this years quarter span
        ilLoop = (igMonthOrQtr - 1) * 3 + 1             'determine starting month based on qtr entred
        gUnpackDate tgMCof(ilRet).iStartDate(0, ilLoop - 1), tgMCof(ilRet).iStartDate(1, ilLoop - 1), slStr     'convert last bdcst billing date to string
        'slStr = Trim$(Str$(ilLoop)) & "/15/" & Trim$(RptSelPS!edcSelCTo.Text)
        For ilLoop = 1 To 4 Step 1
            llTYDates(ilLoop) = gDateValue(gObtainStartCorp(slStr, True))
            slStr = gObtainEndCorp(slStr, False)
            llAdjust = gDateValue(slStr) + 1          'get to next month
            slStr = Format$(llAdjust, "m/d/yy")
        Next ilLoop
        gGetStartEndYear 1, ilYear - 1, slDate, slStr
        slTYStartYr = slDate
        gGetStartEndYear 1, ilYear, slDate, slStr
        slTYEndYr = slStr
    Else
        llTYGetFrom = gDateValue(gObtainStartStd(slStr))  'gather contracts from this date thru effective entered date
        ilYear = Val(RptSelPS!edcSelCTo.Text)           'year requested
        'Determine this years quarter span
        ilLoop = (igMonthOrQtr - 1) * 3 + 1             'determine starting month based on qtr entred
        slStr = Trim$(str$(ilLoop)) & "/15/" & Trim$(RptSelPS!edcSelCTo.Text)
        For ilLoop = 1 To 4 Step 1
            llTYDates(ilLoop) = gDateValue(gObtainStartStd(slStr))
            slStr = gObtainEndStd(slStr)
            llAdjust = gDateValue(slStr) + 1          'get to next month
            slStr = Format$(llAdjust, "m/d/yy")
        Next ilLoop
        gGetStartEndYear 2, ilYear - 1, slDate, slStr
        slTYStartYr = slDate
        gGetStartEndYear 2, ilYear, slDate, slStr
        slTYEndYr = slStr
    End If

    'Determine the Budget name Plan selected
    slNameCode = tgRptSelBudgetCodePS(igBSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 1, "\", slStr)
    ilRet = gParseItem(slStr, 1, "\", slYear)
    slYear = gSubStr("9999", slYear)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilBdMnfPlan = Val(slCode)
    ilBdPlanYear = Val(slYear)

    'Determine the Budget name Forecast selected
    slNameCode = tgRptSelBudgetCodePS(igBFCSelectedIndex).sKey
    ilRet = gParseItem(slNameCode, 1, "\", slStr)
    ilRet = gParseItem(slStr, 1, "\", slYear)
    slYear = gSubStr("9999", slYear)
    ilRet = gParseItem(slNameCode, 2, "\", slCode)
    ilBdMnfFC = Val(slCode)
    ilBdFCYear = Val(slYear)
    'ReDim tlGrf(1 To 1) As GRF          'array of vehicles and their sales
    ReDim tlGrf(0 To 0) As GRF          'array of vehicles and their sales
    'gather all budget records by vehicle for the requested years plan, totaling by quarter
    If Not mReadBvfRec(hmBvf, ilBdMnfPlan, ilBdPlanYear, tmBvfPlan()) Then
        Exit Sub
    End If
    'gather all budget records by vehicle for the requested years forecast, totaling by quarter
    If Not mReadBvfRec(hmBvf, ilBdMnfFC, ilBdFCYear, tmBvfFC()) Then
        Exit Sub
    End If

    'Record Type Definition GRF is used (originally form Sales Analysis Summary report, some variable names may be misleading)
    'Gather budget $ for Plan selected
    For ilMonth = 1 To 3 Step 1
        ilFound = False
        slStartDate = Format$(llTYDates(ilMonth), "m/d/yy")
        slEndDate = Format$(llTYDates(ilMonth + 1) - 1, "m/d/yy")
        'use startwk & endwk to gather budgets
        gObtainWkNo ilBvfCalType, slStartDate, ilStartWk, ilFirstWk          'determine the first week inx to accum (current year)
        gObtainWkNo ilBvfCalType, slEndDate, ilEndWk, ilLastWk               'determine the last week inx to accum  (current year)
        If ilCorpStd = 2 And ilFirstWk = 1 Then                   'if std and week 1 is start, always add week 0
            ilFirstWk = True
        End If
        For ilLoop = LBound(tmBvfPlan) To UBound(tmBvfPlan) - 1 Step 1
            For ilSlsLoop = LBound(tlGrf) To UBound(tlGrf) - 1 Step 1
                'If tmBvfPlan(ilLoop).iSofCode = tlGrf(ilSlsLoop).iSofCode And tlGrf(ilSlsLoop).iPerGenl(2) = ilMonth Then
                If tmBvfPlan(ilLoop).iSofCode = tlGrf(ilSlsLoop).iSofCode And tlGrf(ilSlsLoop).iPerGenl(1) = ilMonth Then
                    ilFound = True
                    Exit For
                End If
            Next ilSlsLoop
            If Not ilFound Then
                tlGrf(UBound(tlGrf)).iSofCode = tmBvfPlan(ilLoop).iSofCode
                'tlGrf(UBound(tlGrf)).iPerGenl(2) = ilMonth
                tlGrf(UBound(tlGrf)).iPerGenl(1) = ilMonth
                'ReDim Preserve tlGrf(1 To UBound(tlGrf) + 1)
                ReDim Preserve tlGrf(0 To UBound(tlGrf) + 1)
            End If
            'ilSlsLoop contains index to the correct vehicle
            For ilTemp = ilStartWk To ilEndWk Step 1
                'tlGrf(ilSlsLoop).lDollars(1) = tlGrf(ilSlsLoop).lDollars(1) + tmBvfPlan(ilLoop).lGross(ilTemp)
                tlGrf(ilSlsLoop).lDollars(0) = tlGrf(ilSlsLoop).lDollars(0) + tmBvfPlan(ilLoop).lGross(ilTemp)
            Next ilTemp
            If ilFirstWk Then       'adjust for the partial weeks at the beginning or end of the year
                                    'due to corp or calendar months
                'tlGrf(ilSlsLoop).lDollars(1) = tlGrf(ilSlsLoop).lDollars(1) + tmBvfPlan(ilLoop).lGross(0)
                tlGrf(ilSlsLoop).lDollars(0) = tlGrf(ilSlsLoop).lDollars(0) + tmBvfPlan(ilLoop).lGross(0)
            End If
            If ilLastWk Then
                'tlGrf(ilSlsLoop).lDollars(1) = tlGrf(ilSlsLoop).lDollars(1) + tmBvfPlan(ilLoop).lGross(53)
                tlGrf(ilSlsLoop).lDollars(0) = tlGrf(ilSlsLoop).lDollars(0) + tmBvfPlan(ilLoop).lGross(53)
            End If
        Next ilLoop
    Next ilMonth
    'Gather budget $ for Forecast selected
    For ilMonth = 1 To 3 Step 1
        ilFound = False
        slStartDate = Format$(llTYDates(ilMonth), "m/d/yy")
        slEndDate = Format$(llTYDates(ilMonth + 1) - 1, "m/d/yy")
        'use startwk & endwk to gather budgets
        gObtainWkNo ilBvfCalType, slStartDate, ilStartWk, ilFirstWk          'determine the first week inx to accum (current year)
        gObtainWkNo ilBvfCalType, slEndDate, ilEndWk, ilLastWk               'determine the last week inx to accum  (current year)
        If ilCorpStd = 2 And ilFirstWk = 1 Then                   'if std and week 1 is start, always add week 0
            ilFirstWk = True
        End If
        For ilLoop = LBound(tmBvfFC) To UBound(tmBvfFC) - 1 Step 1
            For ilSlsLoop = LBound(tlGrf) To UBound(tlGrf) - 1 Step 1
                'If tmBvfFC(ilLoop).iSofCode = tlGrf(ilSlsLoop).iSofCode And tlGrf(ilSlsLoop).iPerGenl(2) = ilMonth Then
                If tmBvfFC(ilLoop).iSofCode = tlGrf(ilSlsLoop).iSofCode And tlGrf(ilSlsLoop).iPerGenl(1) = ilMonth Then
                    ilFound = True
                    Exit For
                End If
            Next ilSlsLoop
            If Not ilFound Then
                tlGrf(UBound(tlGrf)).iSofCode = tmBvfFC(ilLoop).iSofCode
                'tlGrf(UBound(tlGrf)).iPerGenl(2) = ilMonth
                tlGrf(UBound(tlGrf)).iPerGenl(1) = ilMonth
                'ReDim Preserve tlGrf(1 To UBound(tlGrf) + 1)
                ReDim Preserve tlGrf(0 To UBound(tlGrf) + 1)
            End If
            'ilSlsLoop contains index to the correct vehicle
            For ilTemp = ilStartWk To ilEndWk Step 1
                'tlGrf(ilSlsLoop).lDollars(2) = tlGrf(ilSlsLoop).lDollars(2) + tmBvfFC(ilLoop).lGross(ilTemp)
                tlGrf(ilSlsLoop).lDollars(1) = tlGrf(ilSlsLoop).lDollars(1) + tmBvfFC(ilLoop).lGross(ilTemp)
            Next ilTemp
            If ilFirstWk Then       'adjust for the partial weeks at the beginning or end of the year
                                    'due to corp or calendar months
                'tlGrf(ilSlsLoop).lDollars(2) = tlGrf(ilSlsLoop).lDollars(2) + tmBvfPlan(ilLoop).lGross(0)
                tlGrf(ilSlsLoop).lDollars(1) = tlGrf(ilSlsLoop).lDollars(1) + tmBvfPlan(ilLoop).lGross(0)
            End If
            If ilLastWk Then
                'tlGrf(ilSlsLoop).lDollars(2) = tlGrf(ilSlsLoop).lDollars(2) + tmBvfFC(ilLoop).lGross(53)
                tlGrf(ilSlsLoop).lDollars(1) = tlGrf(ilSlsLoop).lDollars(1) + tmBvfFC(ilLoop).lGross(53)
            End If
        Next ilLoop
    Next ilMonth


    'use startwk & endwk to gather  slsp projections
    'gObtainWkNo ilPjfCalType, slStartDate, ilStartWk, ilFirstWk          'determine the first week inx to accum (current year)
    'gObtainWkNo ilPjfCalType, slEndDate, ilEndWk, ilLastWk               'determine the last week inx to accum  (current year)
    'gather all Slsp projection records for the matching rollover date (exclude current records)
    ReDim tmTPjf(0 To 0) As PJF
    ilRet = gObtainPjf(RptSelPS, hmPjf, ilRODate(), tmTPjf())                 'Read all applicable Projection records into memory
    'Build slsp projection $ just gathered into vehicle buckets
    For ilLoop = LBound(tmTPjf) To UBound(tmTPjf) Step 1
        'ilPotnInx = 0
        ilPotnInx = -1
        For ilSlsLoop = LBONE To UBound(tlPotn) - 1 Step 1
            If tmTPjf(ilLoop).iMnfBus = tlPotn(ilSlsLoop).iVefCode Then     'match potential codes
                ilPotnInx = ilSlsLoop
                Exit For
            End If
        Next ilSlsLoop

        'Determine start & end dates of the standard year from the proj recd
        slStr = "12/15/" & Trim$(str$(tmTPjf(ilLoop).iYear))
        llProjYearEnd = gDateValue(gObtainEndStd(slStr))
        slStr = "1/15/" & Trim$(str$(tmTPjf(ilLoop).iYear))
        llProjYearStart = gDateValue(gObtainStartStd(slStr))
        'the quarter must be within the Projection year
        'If ilPotnInx > 0 And llTYDates(4) > llProjYearStart And llTYDates(1) <= llProjYearEnd Then           'potential code exists
        If ilPotnInx >= 0 And llTYDates(4) > llProjYearStart And llTYDates(1) <= llProjYearEnd Then           'potential code exists
            For ilMonth = 1 To 3            'loop for # months in  qtr
                llAdjust = 0
                For ilTemp = 1 To 53
                    llWeek = (ilTemp - 1) * 7 + llProjYearStart
                    If llWeek >= llTYDates(ilMonth) And llWeek < llTYDates(ilMonth + 1) Then
                        For ilSlsLoop = LBound(tlGrf) To UBound(tlGrf) - 1 Step 1
                            'If tlGrf(ilSlsLoop).iSofCode = tmTPjf(ilLoop).iSofCode And tlGrf(ilSlsLoop).iPerGenl(2) = ilMonth Then
                            If tlGrf(ilSlsLoop).iSofCode = tmTPjf(ilLoop).iSofCode And tlGrf(ilSlsLoop).iPerGenl(1) = ilMonth Then
                                llAdjust = llAdjust + tmTPjf(ilLoop).lGross(ilTemp)
                                Exit For
                            End If
                        Next ilSlsLoop                  'accumulate next projection into the GRF recd
                    Else
                        If llWeek >= llTYDates(ilMonth + 1) Then
                            Exit For
                        End If
                    End If
                Next ilTemp
                llTemp = (llAdjust * tlPotn(ilPotnInx).lProject(1)) \ 100  'adjust the gross based on the potential codes most likely %
                'tlGrf(ilSlsLoop).lDollars(5) = tlGrf(ilSlsLoop).lDollars(5) + llTemp
                tlGrf(ilSlsLoop).lDollars(4) = tlGrf(ilSlsLoop).lDollars(4) + llTemp
                llTemp = (llAdjust * tlPotn(ilPotnInx).lProject(2)) \ 100  'adjust the gross based on the potential codes optimistic %
                'tlGrf(ilSlsLoop).lDollars(6) = tlGrf(ilSlsLoop).lDollars(6) + llTemp
                tlGrf(ilSlsLoop).lDollars(5) = tlGrf(ilSlsLoop).lDollars(5) + llTemp
                llTemp = (llAdjust * tlPotn(ilPotnInx).lProject(3)) \ 100  'adjust the gross based on the potential codes pessimistic %
                'tlGrf(ilSlsLoop).lDollars(7) = tlGrf(ilSlsLoop).lDollars(7) + llTemp
                tlGrf(ilSlsLoop).lDollars(6) = tlGrf(ilSlsLoop).lDollars(6) + llTemp
                'Accum unadjusted projection $
                'tlGrf(ilSlsLoop).lDollars(4) = tlGrf(ilSlsLoop).lDollars(4) + llAdjust
                tlGrf(ilSlsLoop).lDollars(3) = tlGrf(ilSlsLoop).lDollars(3) + llAdjust
            Next ilMonth
        End If
    Next ilLoop
    'gather all contracts whose entered date is equal or prior to the requested date (gather from beginning of std year to
    'input date
    slCntrTypes = gBuildCntTypes()
    slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)

    'Process requested year/quarter  to obtain contracts (use the date entered by the user to get contracts)
    llEnterTo = llTYGetTo
    slStartDate = Format$(llTYDates(1), "m/d/yy")
    slEndDate = Format$(llTYDates(4) - 1, "m/d/yy")

    ilRet = gObtainPhfRvf(RptSelPS, slStartDate, slEndDate, tlTranType, tlRvf(), 0)
    For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1 Step 1
        tmRvf = tlRvf(llRvfLoop)
        'dan M 7-15-08 added single contract selectivity
        If llSingleContract = NOT_SELECTED Or llSingleContract = tmRvf.lCntrNo Then
            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStr
            llTemp = gDateValue(slStr)

            ilFound = False
            gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slTemp
            llDateEntered = gDateValue(slTemp)
            'Dan M 8-14-8 ntr/hard cost adjustments.  Is this record ntr/hard cost and do we want that?
            blFailedMatchNtrOrHardCost = False
            'Dan M 8-14-8 don't allow installment option "I"
            blFailedBecauseInstallment = False
            If tmRvf.sType = "I" Then
                blFailedBecauseInstallment = True
            End If
            If ((blIncludeNTR) Xor (blIncludeHardCost)) And tmRvf.iMnfItem > 0 Then      'one or the other is true, but not both (if both true, don't have to isolate anything)
                ilRet = gIsItHardCost(tmRvf.iMnfItem, tmMnfNtr())
            'if is hard cost but blincludentr  or isn't hard cost but blincludehardcost then it needs to be removed. set failedmatchntrorhardcost true
                If (ilRet And blIncludeNTR) Or ((Not ilRet) And blIncludeHardCost) Then
                    blFailedMatchNtrOrHardCost = True
                End If
            End If
            If Not (blFailedMatchNtrOrHardCost Or blFailedBecauseInstallment) Then  'if both false, continue
                If llTemp >= llTYDates(1) And llTemp < llTYDates(4) Then
                    If llDateEntered <= llEnterTo Then  'replaced below dan M
                    'If llTemp <= llEnterTo Then
                        For ilTemp = 1 To 4                           'setup general buffer to use This years dates
                            'llGenlDates(ilTemp) = llTYStarts(ilTemp)
                            If llTemp >= llTYDates(ilTemp) And llTemp < llTYDates(ilTemp + 1) Then
                                llProject(1) = 0                'init bkts to accum qtr $ for this line
                                llProject(2) = 0
                                llProject(3) = 0
                                ilFound = True
                                gPDNToLong tmRvf.sGross, llProject(ilTemp)
                                Exit For
                            End If
                        Next ilTemp
                    End If
                End If
            End If
            If ilFound Then
                'Read the contract
                tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tgChfPS, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
                'Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo <> tmRvf.lCntrNo Or (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M"))
                Do While (ilRet = BTRV_ERR_NONE) And (tgChfPS.lCntrNo = tmRvf.lCntrNo) And (tgChfPS.sSchStatus <> "F" And tgChfPS.sSchStatus <> "M")
                    ilRet = btrGetNext(hmCHF, tgChfPS, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
                If ((ilRet <> BTRV_ERR_NONE) Or (tgChfPS.lCntrNo <> tmRvf.lCntrNo)) Then  'phoney a header from the receivables record so it can be procesed
                    For ilLoop = 0 To 9
                        tgChfPS.iSlfCode(ilLoop) = 0
                        tgChfPS.lComm(ilLoop) = 0
                    Next ilLoop
                    tgChfPS.iAdfCode = tmRvf.iAdfCode
                    tgChfPS.iSlfCode(0) = tmRvf.iSlfCode
                    tgChfPS.lComm(0) = 1000000
                    tgChfPS.iPctTrade = 0
                    If tmRvf.sCashTrade = "T" Then
                        tgChfPS.iPctTrade = 100           'ignore trades   later
                    End If
                End If

            End If                          'ilfoundOne
            If ilFound And tgChfPS.iPctTrade <> 100 Then              'valid contract & selling office, and not trade
                ilSofCode = 0
                For ilLoop = LBound(tlSlf) To UBound(tlSlf) - 1 Step 1
                    If tlSlf(ilLoop).iCode = tgChfPS.iSlfCode(0) Then
                        ilSofCode = tlSlf(ilLoop).iSofCode
                        Exit For
                    End If
                Next ilLoop
                For ilTemp = 1 To 3         'loop for # months in qtr
                    If llProject(ilTemp) <> 0 Then
                        llProject(ilTemp) = llProject(ilTemp) \ 100               'drop pennies
                        For ilSlsLoop = LBound(tlGrf) To UBound(tlGrf) - 1 Step 1
                            'match the contract selling office against the ones built in memory
                            'If tlGrf(ilSlsLoop).iSofCode = ilSofCode And tlGrf(ilSlsLoop).iPerGenl(2) = ilTemp Then
                            If tlGrf(ilSlsLoop).iSofCode = ilSofCode And tlGrf(ilSlsLoop).iPerGenl(1) = ilTemp Then
                                'Actuals OOB
                                'tlGrf(ilSlsLoop).lDollars(3) = tlGrf(ilSlsLoop).lDollars(3) + llProject(ilTemp)
                                tlGrf(ilSlsLoop).lDollars(2) = tlGrf(ilSlsLoop).lDollars(2) + llProject(ilTemp)
                                Exit For
                            End If
                        Next ilSlsLoop
                    End If
                Next ilTemp
            End If
        End If          'single contract selectivity
    Next llRvfLoop
    'ilRet = gObtainCntrForDate(RptSelPs, slStartDate, slEndDate, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
    'Two years of data will be read so that contracts that previous revisions can be processed when
    'schedules are dramatically changed in dates

    slTYEndYr = Format$(gDateValue(slTYEndYr) + 90, "m/d/yy")        'get an extra quarter to make sure all changes included
    ilRet = gObtainCntrForDate(RptSelPS, slTYStartYr, slTYEndYr, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
            '7-15-08 added single contract selectivity dan M
        If llSingleContract = NOT_SELECTED Or llSingleContract = tlChfAdvtExt(ilCurrentRecd).lCntrNo Then
            'project the $
            llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
            'Retrieve the contract, schedule lines and flights
            llContrCode = gPaceCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llEnterTo, hmCHF, tmChf)
            If llContrCode > 0 Then
                ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfPS, tgClfPS(), tgCffPS())
                ilFound = True
                'determine if the contracts start & end dates fall within the requested period
                gUnpackDateLong tgChfPS.iEndDate(0), tgChfPS.iEndDate(1), llAdjust      'hdr end date converted to long
                gUnpackDateLong tgChfPS.iStartDate(0), tgChfPS.iStartDate(1), llTemp    'hdr start date converted to long
                If llAdjust < llTYDates(1) Or llTemp >= llTYDates(4) Then
                    ilFound = False
                End If
            Else
                ilFound = False
            End If



            If ilFound And tgChfPS.iPctTrade <> 100 Then  'exclude 100% trades
                ilSofCode = 0
                For ilLoop = LBound(tlSlf) To UBound(tlSlf) - 1 Step 1
                    If tlSlf(ilLoop).iCode = tgChfPS.iSlfCode(0) Then
                        ilSofCode = tlSlf(ilLoop).iSofCode
                        Exit For
                    End If
                Next ilLoop
                gUnpackDateLong tgChfPS.iOHDDate(0), tgChfPS.iOHDDate(1), llAdjust
                If llAdjust <= llEnterTo Then       'entered date must be entered thru effectve date

                    For ilClf = LBound(tgClfPS) To UBound(tgClfPS) - 1 Step 1
                        llProject(1) = 0                'init bkts to accum qtr $ for this line
                        llProject(2) = 0
                        llProject(3) = 0
                        tmClf = tgClfPS(ilClf).ClfRec
                        If tmClf.sType = "H" Or tmClf.sType = "S" Then
                            gBuildFlights ilClf, llTYDates(), 1, 4, llProject(), 1, tgClfPS(), tgCffPS()
                        End If
                        'If slAirOrder = "O" Then                'invoice all contracts as ordered
                        '    If tmClf.sType <> "H" Then          'ignore all hidden lines for ordered billing, should be Pkg or conventional lines
                        '        gBuildFlights ilClf, llTYDates(), 1, 4, llProject(), 1
                        '    End If
                        'Else                                    'inv all contracts as aired
                        '    If tmClf.sType = "H" Then             'but if from pkg and hidden line, ignore hidd
                        '        'if hidden, will project if assoc. package is set to invoice as aired (real)
                        '        For ilTemp = LBound(tgClfPS) To UBound(tgClfPS) - 1    'find the assoc. pkg line for these hidden
                        '            If tmClf.iPkLineNo = tgClfPS(ilTemp).ClfRec.iLine Then
                        '                If tgClfPS(ilTemp).ClfRec.sType = "A" Then        'does the pkg line reflect bill as aired?
                        '                    gBuildFlights ilClf, llTYDates(), 1, 4, llProject(), 1 'pkg bills as aired, project the hidden line
                        '                End If
                        '                Exit For
                        '            End If
                        '        Next ilTemp
                        '    Else                            'conventional, VV, or Pkg line
                        '        If tmClf.sType <> "A" Then  'if this package line to be invoiced aired (real times),
                        '                                    'it has already been projected above with the hidden line
                        '            gBuildFlights ilClf, llTYDates(), 1, 4, llProject(), 1
                        '        End If
                        '    End If
                        'End If
                        'Accumulate the $ projected into the vehicles buckets
                        'If llProject(1) + llProject(2) + llProject(3) > 0 Then
                        For ilTemp = 1 To 3         'loop for # months in qtr
                            If llProject(ilTemp) <> 0 Then
                                llProject(ilTemp) = llProject(ilTemp) \ 100               'drop pennies
                                For ilSlsLoop = LBound(tlGrf) To UBound(tlGrf) - 1 Step 1
                                    'match the contract selling office against the ones built in memory
                                    'If tlGrf(ilSlsLoop).iSofCode = ilSofCode And tlGrf(ilSlsLoop).iPerGenl(2) = ilTemp Then
                                    If tlGrf(ilSlsLoop).iSofCode = ilSofCode And tlGrf(ilSlsLoop).iPerGenl(1) = ilTemp Then
                                        'Actuals OOB
                                        'tlGrf(ilSlsLoop).lDollars(3) = tlGrf(ilSlsLoop).lDollars(3) + llProject(ilTemp)
                                        tlGrf(ilSlsLoop).lDollars(2) = tlGrf(ilSlsLoop).lDollars(2) + llProject(ilTemp)
                                        Exit For
                                    End If
                                Next ilSlsLoop
                            End If
                        Next ilTemp
                        'End If                      'llproject > 0
                    Next ilClf                      'process nextline
                End If                              'llAdjust <= llEnterTo
            End If                                  'ilfound
                   ' Dan M 7-15-08 Add NTR/Hard Cost option
            'Does user want to see HardCost/NTR?  Not pure trade?
            If llAdjust <= llEnterTo And ilFound Then       ' dan m added ilfound 8-15-08 to be on safe side.
                 If (blIncludeNTR Or blIncludeHardCost) And (tlChfAdvtExt(ilCurrentRecd).iPctTrade <> 100) And (tgChfPS.sNTRDefined = "Y") Then
                    'call routine to fill array with choice
                    gNtrByContract llContrCode, llTYDates(1), llTYDates(4), tlNTRInfo(), tmMnfNtr(), hmSbf, blIncludeNTR, blIncludeHardCost, RptSelPS
                    ilLowerboundNTR = LBound(tlNTRInfo)
                    ilUpperboundNTR = UBound(tlNTRInfo)
                 'ntr or hard cost found?
                    If ilUpperboundNTR <> ilLowerboundNTR Then
                        'clear array
                        'flag to see that contract has a value for writing
                        For ilNTRCounter = ilLowerboundNTR To ilUpperboundNTR - 1 Step 1
                            For ilLoop = 1 To 3
                                llProject(ilLoop) = 0
                            Next ilLoop
                            blNTRWithTotal = False
                                For ilTemp = 1 To 3
                                     'look at each ntr record's date to see if falls into specific time period.
                                    If tlNTRInfo(ilNTRCounter).lSbfDate >= llTYDates(ilTemp) And tlNTRInfo(ilNTRCounter).lSbfDate < llTYDates(ilTemp + 1) Then
                                        'flag so won't write record if all values are 0
                                        blNTRWithTotal = True
                                        llProject(ilTemp) = llProject(ilTemp) + tlNTRInfo(ilNTRCounter).lSBFTotal
                                        Exit For
                                    End If
                                Next ilTemp
                                If blNTRWithTotal = True Then
                                    For ilTemp = 1 To 3         'loop for # months in qtr
                                        If llProject(ilTemp) <> 0 Then
                                            llProject(ilTemp) = llProject(ilTemp) \ 100               'drop pennies
                                            For ilSlsLoop = LBound(tlGrf) To UBound(tlGrf) - 1 Step 1
                                                'match the contract selling office against the ones built in memory
                                                'If tlGrf(ilSlsLoop).iSofCode = ilSofCode And tlGrf(ilSlsLoop).iPerGenl(2) = ilTemp Then
                                                If tlGrf(ilSlsLoop).iSofCode = ilSofCode And tlGrf(ilSlsLoop).iPerGenl(1) = ilTemp Then
                                                    'Actuals OOB
                                                    'tlGrf(ilSlsLoop).lDollars(3) = tlGrf(ilSlsLoop).lDollars(3) + llProject(ilTemp)
                                                    tlGrf(ilSlsLoop).lDollars(2) = tlGrf(ilSlsLoop).lDollars(2) + llProject(ilTemp)
                                                    Exit For
                                                End If
                                            Next ilSlsLoop
                                        End If
                                    Next ilTemp
                                End If
                        Next ilNTRCounter
                    End If
                End If
            End If      'lladjust/entrydate


        End If                                  'single contract
    Next ilCurrentRecd


    For ilLoop = LBound(tlGrf) To UBound(tlGrf) - 1 Step 1         'write a record per vehicle
        tlGrf(ilLoop).iGenDate(0) = igNowDate(0)        'todays date used for removal of records
        tlGrf(ilLoop).iGenDate(1) = igNowDate(1)
        'tlGrf(ilLoop).iGenTime(0) = igNowTime(0)        'todays time used for removal of records
        'tlGrf(ilLoop).iGenTime(1) = igNowTime(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tlGrf(ilLoop).lGenTime = lgNowTime
        tlGrf(ilLoop).iStartDate(0) = ilEnterDate(0)             'effective date entered
        tlGrf(ilLoop).iStartDate(1) = ilEnterDate(1)
        tlGrf(ilLoop).iCode2 = ilBdMnfPlan                         'budget name
        tlGrf(ilLoop).iYear = ilBdMnfFC                 'budget forecast name
        'tlGrf(ilLoop).iPerGenl(1) = ilCorpStd             '1 = corp, 2 = std
        tlGrf(ilLoop).iPerGenl(0) = ilCorpStd             '1 = corp, 2 = std
        ''tlGrf.iPerGenl(2) = the month index 1-3, need to put in the actual month # based on the start quarter
        'tlGrf(ilLoop).iPerGenl(2) = (igMonthOrQtr - 1) * 3 + tlGrf(ilLoop).iPerGenl(2)
        tlGrf(ilLoop).iPerGenl(1) = (igMonthOrQtr - 1) * 3 + tlGrf(ilLoop).iPerGenl(1)
        For ilPotnInx = 1 To 3
            If ilPotnInx = 1 Then
                ilRet = InStr(slPotn, "A")
            ElseIf ilPotnInx = 2 Then
                ilRet = InStr(slPotn, "B")
            ElseIf ilPotnInx = 3 Then
                ilRet = InStr(slPotn, "C")
            End If
            If ilRet > 0 Then
                ilRet = ilRet - 1
                'tlGrf(ilLoop).iPerGenl((ilPotnInx - 1) * 3 + 3) = tlPotn(ilRet).lProject(1)     'Most Likely
                'tlGrf(ilLoop).iPerGenl((ilPotnInx - 1) * 3 + 4) = tlPotn(ilRet).lProject(2)   'Optimistic
                'tlGrf(ilLoop).iPerGenl((ilPotnInx - 1) * 3 + 5) = tlPotn(ilRet).lProject(3)   'Pessimistic
                tlGrf(ilLoop).iPerGenl((ilPotnInx - 1) * 3 + 2) = tlPotn(ilRet).lProject(1)     'Most Likely
                tlGrf(ilLoop).iPerGenl((ilPotnInx - 1) * 3 + 3) = tlPotn(ilRet).lProject(2)   'Optimistic
                tlGrf(ilLoop).iPerGenl((ilPotnInx - 1) * 3 + 4) = tlPotn(ilRet).lProject(3)   'Pessimistic
            End If
        Next ilPotnInx
        ilRet = btrInsert(hmGrf, tlGrf(ilLoop), imGrfRecLen, INDEXKEY0)
    Next ilLoop
    Erase tlGrf, tlMMnf, tmBvfPlan, tmBvfFC, tlPotn, tmMnfNtr
    Erase tmTPjf, tlChfAdvtExt, tgClfPS, tgCffPS
    sgCntrForDateStamp = ""             'init time stamp if re-entering
    ilRet = btrClose(hmBvf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmPjf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmSbf)
End Sub
