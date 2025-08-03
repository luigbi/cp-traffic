Attribute VB_Name = "RPTCRAV"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrav.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text
Const NOT_SELECTED = 0
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

'  Receivables File
Dim tmRvf As RVF            'RVF record image
'********************************************************************************************
'
'                   gCrAdvVariance - Prepass for Advertiser Variance report
'                   by Week or by Month.
'
'                   User selectivity:  Effective Date (backed up to Monday)
'                                      Start Yr & Qtr
'                                      Week or Month (if week- ask # Qtrs)
'                                                    (if month, always 4 qtrs)
'                                      Corp or Std Month
'                                      All or selective advertisr
'                   This is a pacing report where all contracts are gather
'                   if the contract entered date is equal/prior to the
'                   sunday of the Effective date entred, and whose start/end
'                   dates span the quarter(s) requested.
'                   Records are written to GRF by contract.
'
'                   Created:  9/9/97 D. Hosaka
'                   4/10/98 Exclude 100% trade contracts
'                   3/5/99 Selective advt causing subscript out of range
'       2-11-05 Chg array of PHF/RVF array from integer to long (prevent overflow)
'       12-08-06 Last Years pacing date is gathering from incorrect date
'       12-14-06 add parm to gObtainRvfPhf to test on tran date (vs entry date)
'********************************************************************************************
Sub gCrAdvVariance()
Dim ilRet As Integer                    '
Dim ilClf As Integer                    'loop for schedule lines
Dim ilHOState As Integer                'retrieve only latest order or revision
Dim slCntrTypes As String               'retrieve remnants, PI, DR, etc
Dim slCntrStatus As String              'retrieve H, O G or N (holds, orders, unsch hlds & orders)
Dim ilCurrentRecd As Integer            'loop for processing last years contracts
Dim llContrCode As Long                 'Internal Contr code to build all lines & flights of a contract
Dim ilFoundOne As Integer               'Found a matching  office built into mem
Dim ilFoundAdvt As Integer
Dim ilTemp As Integer
Dim ilLoop As Integer                   'temp loop variable
Dim slTemp As String                    'temp string for dates
Dim ilPeriod As Integer                 '13 = periods to gather or 53 periods for weekly
Dim ilWkOrMonth As Integer              '1 = Month, 2 = weekly option
Dim ilCalType As Integer                '1 = std, 2 = corp calendar
'ReDim llProject(1 To 53) As Long        '$ projected for 12 months or 53 weeks
ReDim llProject(0 To 53) As Long        '$ projected for 12 months or 53 weeks. Index zero ignored
Dim llDate As Long                      'temp date variable
Dim llDate2 As Long
Dim ilTY As Integer                         'true if the contract processing this year, else false
Dim ilQtr As Integer                    '# qtrs requested
Dim slNameCode As String
Dim slCode As String
Dim slYear As String
Dim slMonth As String
Dim slDay As String
'Date used to gather information
'String formats for generalized date conversions routines
'Long formats for testing
'Packed formats to store in GRF record
Dim ilStartMonth As Integer             'start month to gather data (10=Oct)
Dim ilEndMonth As Integer               'end month to gather data (9=sept)
Dim ilLYStartYr As Integer              ' year of last years start date    (1996-1997)
Dim ilLYEndYr As Integer                ' year of last years end date
Dim ilTYStartYr As Integer              'year of this years start date     (1997-1998)
Dim ilTYEndYr As Integer                'year of this years end date
Dim slTYJan As String                   'true beginning of "calendar year" so that an offset can be obtained to retrieve'the same time for last year
Dim slLYJan As String
Dim slLYStart As String                 'start date of last year to begin gathering (could be std or corp)
Dim slLYEnd As String                   'end date of last year to begin gathering (could be std or corp)
Dim llLYStart As Long                   'start date of last year to begin gathering (long)
Dim llLYEnd As Long                     'end date of last year to begin gathering (long
Dim slTYStart As String                 'start date of this year to begin gathering  (string)
Dim slTYEnd As String                   'end date of this year to begin gathering
Dim llTYStart As Long                   'start date of this year to begin gathering (Long)
Dim llTYEnd As Long                     'end date of this year to begin gthering (long)
Dim llWeekLYStart As Long           'start date of week for same time last year (based on week index)
Dim llWeekLYEnd As Long             'end date of week for same time last year (based on week index)
ReDim ilWeekLYStart(0 To 1) As Integer    'packed format for GRF record
Dim slWeekTYStart As String              'start date of week for this years new business entered this week
Dim llWeekTYStart As Long                'start date of week for this years new business entered on te user entered week
ReDim ilWeekTYStart(0 To 1) As Integer     'packed format for GRF record
Dim llEntryDate As Long                 'date entered from cntr header
Dim slTYEffStart As String          'start date (std or corp) or the week of pacing
Dim slTYEffEnd As String
Dim slLYEffStart As String
Dim slLYEffEnd As String
Dim llTYEffStart As Long
Dim llTYEffEnd As Long
Dim ilYearInx As Integer
Dim ilLoopOnYear As Integer 'loop twice for contracts, 1 for year requested, 1 for previous year
Dim tlNTRInfo() As NTRPacing
Dim blIncludeNTR As Boolean
Dim blIncludeHardCost As Boolean
Dim ilNTRCounter As Integer
Dim blNTRWithTotal As Boolean
Dim ilLowerboundNTR As Integer
Dim ilUpperboundNTR As Integer
Dim llSingleContract As Long            '7-7-08 test for option to get single contract
Dim llDateEntered As Long               'receivables entered date for pacing test
Dim blFailedMatchNtrOrHardCost As Boolean
Dim blFailedBecauseInstallment As Boolean
Dim slGrossNet As String * 1
Dim llGrossNet As Long

'ReDim llGenlDates(1 To 53) As Long          'start dates of each month/week.  This will contain either this years 13 dates or last years 13 dates,
ReDim llGenlDates(0 To 53) As Long          'start dates of each month/week.  This will contain either this years 13 dates or last years 13 dates,
                                            'depending on the contracts dates. Index zero ignored
'Month Starts to gather projection $ from flights
'ReDim llTYStarts(1 To 54) As Long           'this year corp or std start dates or weeks
ReDim llTYStarts(0 To 54) As Long           'this year corp or std start dates or weeks. Index zero ignored
'ReDim llLYStarts(1 To 54) As Long           'last year corp or std start dates or weeks
ReDim llLYStarts(0 To 54) As Long           'last year corp or std start dates or weeks. Index zero ignored
'   end of date variables
Dim tlTranType As TRANTYPES                 'valid trans types to use
'ReDim tlRvf(1 To 1) As RVF
ReDim tlRvf(0 To 0) As RVF
Dim llRvfLoop As Long                       '2-11-05
Dim llAmt As Long                           '3-28-19

    blIncludeNTR = False
    blIncludeHardCost = False
    'dan M 7-07-08 added single contract selectivity
    If Val(RptSelAv!edcContract.Text) > 0 And RptSelAv!edcContract.Text <> " " Then
        llSingleContract = Val(RptSelAv!edcContract.Text)
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
     ' Dan M 6-23-08
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
    
    If RptSelAv!rbcGrossNet(0).Value = True Then    '3-28-19 gross
        slGrossNet = "G"
    Else
        slGrossNet = "N"
    End If

    tlTranType.iAdj = True              'look only for adjustments in the History & Rec files
    tlTranType.iInv = False
    tlTranType.iWriteOff = False
    tlTranType.iPymt = False
    tlTranType.iCash = True
    tlTranType.iTrade = False
    tlTranType.iMerch = False
    tlTranType.iPromo = False
    tlTranType.iNTR = False         '9-17-02

    If RptSelAv!ckcSelCInclude(0).Value Or RptSelAv!ckcSelCInclude(1).Value Then
        tlTranType.iNTR = True
        'set flags that want ntr /hard cost
        If RptSelAv!ckcSelCInclude(0).Value = 1 Then
            blIncludeNTR = True
        End If
        If RptSelAv!ckcSelCInclude(1).Value = 1 Then
            blIncludeHardCost = True
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
        ilRet = mFillMnfArray
        If ilRet <> True Then
            MsgBox "error retrieving MNF files", vbOKOnly + vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    ReDim tgClfAV(0 To 0) As CLFLIST
    tgClfAV(0).iStatus = -1 'Not Used
    tgClfAV(0).lRecPos = 0
    tgClfAV(0).iFirstCff = -1
    ReDim tgCffAV(0 To 0) As CFFLIST
    tgCffAV(0).iStatus = -1 'Not Used
    tgCffAV(0).lRecPos = 0
    tgCffAV(0).iNextCff = -1
    'Get STart and end dates of current week for for pacing test
'    slWeekTYStart = RptSelAv!edcSelCFrom.Text
    slWeekTYStart = RptSelAv!CSI_CalFrom.Text       '9-3-19 use csi cal control vs edit box

    llWeekTYStart = gDateValue(slWeekTYStart)
    gPackDate slWeekTYStart, ilWeekTYStart(0), ilWeekTYStart(1)    'conversion to store in prepass record
    'llWeekTYEnd = llWeekTYStart + 6
    'slTemp = Format$(llWeekTYEnd, "m/d/yy")
    slMonth = RptSelAv!edcSelCFrom1.Text
    ilQtr = Val(slMonth)                            'save # qtrs requested
    'setup year from user input
    'gObtainYearMonthDayStr slTemp, True, slYear, slMonth, slDay
    ilTYStartYr = Val(RptSelAv!edcSelCTo.Text)
    ilStartMonth = Val(RptSelAv!edcSelCTo1.Text)
    ilStartMonth = (ilStartMonth - 1) * 3 + 1       'for std, actual starting month, for corp its the index into corp cal month

    If RptSelAv!rbcSelCSelect(0).Value Then      'corp month or qtr? (vs std)
        'ilCalType = 2                   'corp
        ilLoop = gGetCorpCalIndex(ilTYStartYr)
        'gUnpackDateLong tgMCof(ilLoop).iStartDate(0, 1), tgMCof(ilLoop).iStartDate(1, 1), llDate
        gUnpackDateLong tgMCof(ilLoop).iStartDate(0, 0), tgMCof(ilLoop).iStartDate(1, 0), llDate
        slTYJan = Format$(llDate, "m/d/yy")
        gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilStartMonth - 1), tgMCof(ilLoop).iStartDate(1, ilStartMonth - 1), llDate
        slTYStart = Format$(llDate, "m/d/yy")
        'assume starting from quarter 1 of corporateyear
        'gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 12), tgMCof(ilLoop).iEndDate(1, 12), llDate
        gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), llDate
        slTYEnd = Format$(llDate, "m/d/yy")
        'if not starting at beginning of fiscal month, need to get the next year for the end corp months
        ilEndMonth = ilStartMonth
        ilTYEndYr = ilTYStartYr
        If ilStartMonth <> 1 Then
            ilLoop = gGetCorpCalIndex(ilTYStartYr + 1)
            ilEndMonth = ilStartMonth - 1
            gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilEndMonth - 1), tgMCof(ilLoop).iEndDate(1, ilEndMonth - 1), llDate
            slTYEnd = Format$(llDate, "m/d/yy")
            ilTYEndYr = tgMCof(ilLoop).iYear
        End If

        ilYearInx = gGetCorpInxByDate(llWeekTYStart, llTYEffStart, llTYEffEnd)
        slTYEffStart = Format$(llTYEffStart, "m/d/yy")
        slTYEffEnd = Format$(llTYEffEnd, "m/d/yy")

        llTYEffStart = gDateValue(slTYEffStart)
        ilYearInx = gGetCorpCalIndex(tgMCof(ilYearInx).iYear - 1)   'get previous years effective date start date of yr
        'gUnpackDate tgMCof(ilYearInx).iStartDate(0, 1), tgMCof(ilYearInx).iStartDate(1, 1), slLYEffStart         'convert last bdcst billing date to string
        'gUnpackDate tgMCof(ilYearInx).iEndDate(0, 12), tgMCof(ilYearInx).iEndDate(1, 12), slLYEffEnd
        gUnpackDate tgMCof(ilYearInx).iStartDate(0, 0), tgMCof(ilYearInx).iStartDate(1, 0), slLYEffStart         'convert last bdcst billing date to string
        gUnpackDate tgMCof(ilYearInx).iEndDate(0, 11), tgMCof(ilYearInx).iEndDate(1, 11), slLYEffEnd
    Else
        ilCalType = 1                   'std
        'Determine earliest and latest dates of current year and last year
        slTYStart = Trim$(str$(ilStartMonth)) & "/15/" & Trim$(str$(ilTYStartYr))
        slTYJan = "01/15/" & Trim$(str$(ilTYStartYr))
        ilEndMonth = ilStartMonth
        ilTYEndYr = ilTYStartYr

        ilLYStartYr = ilTYStartYr - 1
        ilLYEndYr = ilLYStartYr
        For ilLoop = 1 To 11            'determine the month of the last month to process
            If ilEndMonth = 12 Then
                ilEndMonth = 0
                ilTYEndYr = ilTYStartYr + 1
                ilLYEndYr = ilLYEndYr + 1
            End If
            ilEndMonth = ilEndMonth + 1
        Next ilLoop

        slTYEnd = Trim$(str$(ilEndMonth)) & "/15/" & Trim$(str$(ilTYEndYr))
        slTYStart = gObtainStartStd(slTYStart)              'obtain start and end dates of current std year
        slTYEnd = gObtainEndStd(slTYEnd)
        slTYJan = gObtainStartStd(slTYJan)

        'determine start/end dates  of the effective date
        gObtainYearMonthDayStr slWeekTYStart, True, slYear, slMonth, slDay
        slTYEffStart = "01/15/" & Trim$(slYear)
        slTYEffEnd = "12/15/" & Trim$(slYear)
        slTYEffStart = gObtainStartStd(slTYEffStart)
        slTYEffEnd = gObtainEndStd(slTYEffEnd)

    End If
    'Last Year
    If RptSelAv!rbcSelCSelect(0).Value Then                   'corp month (vs std)
        ilLYStartYr = ilTYStartYr - 1
        ilLoop = gGetCorpCalIndex(ilLYStartYr)
        'gUnpackDateLong tgMCof(ilLoop).iStartDate(0, 1), tgMCof(ilLoop).iStartDate(1, 1), llDate
        gUnpackDateLong tgMCof(ilLoop).iStartDate(0, 0), tgMCof(ilLoop).iStartDate(1, 0), llDate
        slLYJan = Format$(llDate, "m/d/yy")
        'gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilStartMonth), tgMCof(ilLoop).iStartDate(1, ilStartMonth), llDate
        gUnpackDateLong tgMCof(ilLoop).iStartDate(0, ilStartMonth - 1), tgMCof(ilLoop).iStartDate(1, ilStartMonth - 1), llDate
        slLYStart = Format$(llDate, "m/d/yy")
        'assume running from Qtr1 to Qtr4 and not a wraparound year
        'gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 12), tgMCof(ilLoop).iEndDate(1, 12), llDate
        gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iEndDate(1, 11), llDate
        slLYEnd = Format$(llDate, "m/d/yy")
        'if not starting at beginning of fiscal month, need to get the next year for the end corp months
        ilEndMonth = ilStartMonth
        ilLYEndYr = ilLYStartYr
        If ilStartMonth <> 1 Then
            ilLoop = gGetCorpCalIndex(ilLYStartYr + 1)
            ilEndMonth = ilStartMonth - 1
            'gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilEndMonth), tgMCof(ilLoop).iEndDate(1, ilEndMonth), llDate
            gUnpackDateLong tgMCof(ilLoop).iEndDate(0, ilEndMonth - 1), tgMCof(ilLoop).iEndDate(1, ilEndMonth - 1), llDate
            slLYEnd = Format$(llDate, "m/d/yy")
            ilLYEndYr = tgMCof(ilLoop).iYear
        End If
    Else
        ilCalType = 1                   'std                'obtain start and end dates of previous std year
        slLYStart = Trim$(str$(ilStartMonth)) & "/15/" & Trim$(str$(ilLYStartYr))
        slLYEnd = Trim$(str$(ilEndMonth)) & "/15/" & Trim$(str$(ilLYEndYr))
        slLYJan = "01/15/" & Trim$(str$(ilLYStartYr))
        slLYStart = gObtainStartStd(slLYStart)
        slLYEnd = gObtainEndStd(slLYEnd)
        slLYJan = gObtainStartStd(slLYJan)

        'determine start/end dates of last years  effective date
        slLYEffStart = "01/15/" & Trim$(str(Val(slYear - 1)))
        slLYEffEnd = "12/15/" & Trim$(str(Val(slYear - 1)))
        slLYEffStart = gObtainStartStd(slLYEffStart)
        slLYEffEnd = gObtainEndStd(slLYEffEnd)

    End If
    llTYStart = gDateValue(slTYStart)               'This years start date needed  in long format  to test spans
    llTYEnd = gDateValue(slTYEnd)                   'This years end date needed in long format to test spans
    llLYStart = gDateValue(slLYStart)               'Last years start date needed in long format to test spans
    llLYEnd = gDateValue(slLYEnd)                   'Last years end date needed in long format to test spans


    'get start and end dates (same time last year) for orders on books same time last year
    'based on week index of todays date
    'llWeekLYStart = gDateValue(slLYJan) + (llWeekTYStart - gDateValue(slTYJan))
    llWeekLYStart = gDateValue(slLYEffStart) + (llWeekTYStart - gDateValue(slTYEffStart))
    slTemp = Format(llWeekLYStart, "m/d/yy")
    llWeekLYEnd = llWeekLYStart + 6
    gPackDate slTemp, ilWeekLYStart(0), ilWeekLYStart(1)
    'Determine startdates for for last year and this year   months
    'Std or corp?
    If RptSelAv!rbcSelCInclude(0).Value Then            'weekly option
        llTYStarts(1) = llTYStart
        llLYStarts(1) = llLYStart
        For ilLoop = 1 To 53
            llTYStarts(ilLoop + 1) = llTYStarts(ilLoop) + 7
            llLYStarts(ilLoop + 1) = llLYStarts(ilLoop) + 7
        Next ilLoop
    Else                                                'monthly (qtrly) version
        'lldate and iltemp are filler, not required for return
        'gSetupBOBDates ilTYStartYr, ilStartMonth, ilLoop, llTYStarts(), lldate, ilTemp    'build array of start & end dates
        'gSetupBOBDates ilLYStartYr, ilStartMonth, ilLoop, llLYStarts(), lldate, ilTemp    'build array of start & end dates
        For ilTemp = 1 To 2
            If ilTemp = 1 Then                  'last year
                slCode = slLYStart
            Else
                slCode = slTYStart
            End If
            If RptSelAv!rbcSelCSelect(0).Value Then                   'corp month (vs std)
                For ilLoop = 1 To 13 Step 1
                    slCode = gObtainStartCorp(slCode, True)
                    llTYStarts(ilLoop) = gDateValue(slCode)
                    slCode = gObtainEndCorp(slCode, True)
                    llDate = gDateValue(slCode) + 1                      'increment for next month
                    slCode = Format$(llDate, "m/d/yy")
                Next ilLoop
            Else
                For ilLoop = 1 To 13 Step 1
                    slCode = gObtainStartStd(slCode)
                    llTYStarts(ilLoop) = gDateValue(slCode)
                    slCode = gObtainEndStd(slCode)
                    llDate = gDateValue(slCode) + 1                      'increment for next month
                    slCode = Format$(llDate, "m/d/yy")
                Next ilLoop
            End If
            If ilTemp = 1 Then      'calc last years date, move them to LY date array
                For ilLoop = 1 To 13
                    llLYStarts(ilLoop) = llTYStarts(ilLoop)
                Next ilLoop
            End If
        Next ilTemp
    End If

    ilRet = gObtainPhfRvf(RptSelAv, slLYStart, slTYEnd, tlTranType, tlRvf(), 0)
    ilPeriod = 12                       'assume quarterly report
    ilWkOrMonth = 1
    If RptSelAv!rbcSelCInclude(0).Value Then
        ilPeriod = 53
        ilWkOrMonth = 2
    End If
    For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1 Step 1
        tmRvf = tlRvf(llRvfLoop)
        'dan M 7-07-08 added single contract selectivity
       If llSingleContract = NOT_SELECTED Or llSingleContract = tmRvf.lCntrNo Then
            gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slTemp
            llDate = gDateValue(slTemp)
            'Dan M 8-8-8  adding ntr/hard cost to report means adding it to rvf
            gUnpackDate tmRvf.iDateEntrd(0), tmRvf.iDateEntrd(1), slTemp
            llDateEntered = gDateValue(slTemp)
            ilTY = False
            ilFoundOne = False
            'Dan M 8-8-8 ntr/hard cost adjustments.  Is this record ntr/hard cost and do we want that?
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
                If llDate >= llTYStart And llDate <= llTYEnd Then
                        ilTY = True
                       ' If llDate <= llWeekTYStart Then              'replaced dan M 8-8-8 with below. Make sure adjustments made are less then effective date
                        If llDateEntered <= llWeekTYStart Then
                            For ilTemp = 1 To ilPeriod                            'setup general buffer to use This years dates
                                'If llDate >= llTYStarts(ilTemp) And llDate < llTYStarts(ilTemp + 1) Then
                                If llDate >= llTYStarts(ilTemp) And llDate < llTYStarts(ilTemp + 1) Then
                                    ilFoundOne = True
                                    If slGrossNet = "G" Then                '3-28-19 add gross or net option
                                        gPDNToLong tmRvf.sGross, llGrossNet
                                    Else
                                        gPDNToLong tmRvf.sNet, llGrossNet
                                    End If
                                    'gPDNToLong tmRvf.sGross, llProject(ilTemp)
                                    llProject(ilTemp) = llGrossNet
                                    Exit For
                                End If
                            Next ilTemp
                        End If
                'if trans date not within current year, assume last year
                Else
                    'If llDate <= llWeekLYStart Then        'dan M 8-8-8 changed as above, but probably not necessary
                    If llDateEntered <= llWeekLYStart Then              '
                    For ilTemp = 1 To ilPeriod                        'setup general buffer to use last years dates
                        If llDate >= llLYStarts(ilTemp) And llDate < llLYStarts(ilTemp + 1) Then
                            ilFoundOne = True
                            If slGrossNet = "G" Then                    '3-28-19 add gross or net option
                                gPDNToLong tmRvf.sGross, llGrossNet
                            Else
                                gPDNToLong tmRvf.sNet, llGrossNet
                            End If
                            'gPDNToLong tmRvf.sGross, llProject(ilTemp)
                            llProject(ilTemp) = llGrossNet
                            Exit For
                        End If
                    Next ilTemp
                    End If
                End If
            End If      'failedMatch
            If ilFoundOne Then
                'Read the contract
                tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
                tmChfSrchKey1.iCntRevNo = 32000
                tmChfSrchKey1.iPropVer = 32000
                ilRet = btrGetGreaterOrEqual(hmCHF, tgChfAV, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd
                Do While (ilRet = BTRV_ERR_NONE) And (tgChfAV.lCntrNo = tmRvf.lCntrNo) And (tgChfAV.sSchStatus <> "F" And tgChfAV.sSchStatus <> "M")
                    ilRet = btrGetNext(hmCHF, tgChfAV, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop

                If ((ilRet <> BTRV_ERR_NONE) Or (tgChfAV.lCntrNo <> tmRvf.lCntrNo)) Then  'phoney a header from the receivables record so it can be procesed
                    For ilLoop = 0 To 9
                        tgChfAV.iSlfCode(ilLoop) = 0
                        tgChfAV.lComm(ilLoop) = 0
                    Next ilLoop
                    tgChfAV.iSlfCode(0) = tmRvf.iSlfCode
                    tgChfAV.lComm(0) = 1000000
                    tgChfAV.iPctTrade = 0
                    If tmRvf.sCashTrade = "T" Then
                        tgChfAV.iPctTrade = 100           'ignore trades   later
                    End If
                    tgChfAV.iAdfCode = tmRvf.iAdfCode
                End If
                ilFoundOne = True
                'Test for selective advertisers
                If Not gSetCheck(RptSelAv!ckcAll.Value) Then
                    ilFoundOne = False
                    For ilTemp = 0 To RptSelAv!lbcSelection.ListCount - 1 Step 1
                        If RptSelAv!lbcSelection.Selected(ilTemp) Then              'selected advt
                            slNameCode = tgAdvertiser(ilTemp).sKey 'Traffic!lbcAdvertiser.List(ilTemp)         'pick up slsp code
                            ilRet = gParseItem(slNameCode, 2, "\", slCode)
                            'If Val(slCode) = tlChfadvtExt(ilCurrentRecd).iadfCode Then
                            If Val(slCode) = tgChfAV.iAdfCode Then
                                ilFoundOne = True
                                Exit For
                            End If
                        End If
                    Next ilTemp
                End If                      'not ckcall
                If ilFoundOne Then
                    If ilTY Then
                        mWriteAV ilCalType, ilWkOrMonth, ilWeekTYStart(), ilWeekLYStart(), ilTYStartYr, llProject(), llTYStarts(1), ilQtr
                    Else
                        'If ilFoundOne Then
                        mWriteAV ilCalType, ilWkOrMonth, ilWeekTYStart(), ilWeekLYStart(), ilLYStartYr, llProject(), llLYStarts(1), ilQtr
                        'End If
                    End If
                Else
                    For ilLoop = 1 To 53
                        llProject(ilLoop) = 0
                    Next ilLoop
                End If
            End If                          'ilfoundOne
        End If                              ' singlecontract
    Next llRvfLoop

    'Gather all contracts for previous year and current year whose effective date entered
    'is prior to the effective date that affects either previous year or current year
    slCntrTypes = gBuildCntTypes()
    slCntrStatus = "HOGN"               'Holds, orders, unsch hold, unsch order.
    ilHOState = 2                       'get latest orders & revisions   (may include G & N if later, plus revised orders turned proposals WCI)
    'Alter last years start date to pick up the contracts that may have been modified
    'after the pacing date, removing the period that should be included.
    slLYStart = Format$(gDateValue(slLYStart) - 90, "m/d/yy")   'go back 1 qtr for last years active dates becuase
    '                                       the latest version might have been altered, which cancelled the period
    '                                       the period to be reported

    slTYEnd = Format$(gDateValue(slTYEnd) + 90, "m/d/yy")        'get an extra quarter to make sure all changes included
    ilRet = gObtainCntrForDate(RptSelAv, slLYStart, slTYEnd, slCntrStatus, slCntrTypes, ilHOState, tlChfAdvtExt())

    'All contracts have been retrieved for all of this year plus all of last year
    'If running by week, # of qtrs produced is entered by the user.  Alter the
    'end date of previous and current years end date if less than 4 qtrs requested
    If RptSelAv!rbcSelCInclude(0).Value Then        'week option
        ilLoop = ilQtr * 13 + 1              '#months to gather
        llTYEnd = llTYStarts(ilLoop) - 1
        'convert the true this years end date for weekly option to string
        slTYEnd = Format$(llTYEnd, "m/d/yy")
        llLYEnd = llLYStarts(ilLoop) - 1
        'convert the true last years end date for weekly option to string
        slLYEnd = Format$(llLYEnd, "m/d/yy")
        ilPeriod = 53                        'weekly version, do 52 weeks buckets
        ilWkOrMonth = 2                     'weekly flag
    Else
        ilPeriod = 13                       'monthly version, do 12 monthly buckets
        ilWkOrMonth = 1                     'monthly flag
    End If
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
            'dan M 7-07-08 added single contract selectivity
        If llSingleContract = NOT_SELECTED Or (llSingleContract <> NOT_SELECTED And llSingleContract = tlChfAdvtExt(ilCurrentRecd).lCntrNo) Then
            ilFoundAdvt = True
            If Not gSetCheck(RptSelAv!ckcAll.Value) Then
                ilFoundAdvt = False
                For ilTemp = 0 To RptSelAv!lbcSelection.ListCount - 1 Step 1
                    If RptSelAv!lbcSelection.Selected(ilTemp) Then              'selected advt
                        slNameCode = tgAdvertiser(ilTemp).sKey 'Traffic!lbcAdvertiser.List(ilTemp)         'pick up slsp code
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = tlChfAdvtExt(ilCurrentRecd).iAdfCode Then
                            ilFoundAdvt = True
                            Exit For
                        End If
                    End If
                Next ilTemp
            End If                      'not ckcall
            ilFoundOne = ilFoundAdvt
            For ilLoopOnYear = 1 To 2
                If (ilFoundOne) Then
                    'get cnts earliest and latest dates to see if it spans the requested period
                    'gUnPackDate tlChfAdvtExt(ilCurrentRecd).iStartDate(0), tlChfAdvtExt(ilCurrentRecd).iStartDate(1), slTemp
                    'llDate = gDateValue(slTemp)
                    'gUnPackDate tlChfAdvtExt(ilCurrentRecd).iEndDate(0), tlChfAdvtExt(ilCurrentRecd).iEndDate(1), slTemp
                    'llDate2 = gDateValue(slTemp)
                    'Test this year
                    'llContrCode = 0
                    ''If ilLoopOnYear = 1 Then
                    '    If llDate <= llTYEnd And llDate2 >= llTYStart Then          'Does contract dates span requested period for this year?
                     '       llContrCode = gPaceCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llWeekTYStart, hmChf, tmChf)
                    '    End If
                    'Else                'previous year?
                        'Else                            'test if cnt with previous year
                            'If llDate >= llLYStart And llDate <= llLYEnd Then
                            'is hdr start date <= last years end date and hdr end date >= last years start?
                    '        If llDate <= llLYEnd And llDate2 >= llLYStart Then          'Does contract dates span requested period for this year?
                    '            llContrCode = gPaceCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llWeekLYStart, hmChf, tmChf)
                    '        End If
                        'End If
                    'End If




                    llContrCode = 0
                    ilTY = False
                    If ilLoopOnYear = 1 Then                   'current
                        llContrCode = gPaceCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llWeekTYStart, hmCHF, tmChf)
                        gUnpackDate tmChf.iStartDate(0), tmChf.iStartDate(1), slTemp
                        llDate = gDateValue(slTemp)
                        gUnpackDate tmChf.iEndDate(0), tmChf.iEndDate(1), slTemp
                        llDate2 = gDateValue(slTemp)
                        If llContrCode > 0 And llDate <= llTYEnd And llDate2 >= llTYStart Then
                            ilTY = True
                        'Else
                        '    llContrCode = 0             'force not found for this year
                        End If
                        'dan M 6-23-08 grab NTR records
                        If (blIncludeNTR Or blIncludeHardCost) And (ilTY) And (tlChfAdvtExt(ilCurrentRecd).iPctTrade <> 100) Then
                            mNTRByContract llContrCode, llDate, llDate2, tlNTRInfo(), blIncludeNTR, blIncludeHardCost
                            ilLowerboundNTR = LBound(tlNTRInfo)
                            ilUpperboundNTR = UBound(tlNTRInfo)
                            'find Ntr?
                            If ilUpperboundNTR <> ilLowerboundNTR Then
                            'clear array
                                For ilLoop = 1 To 53
                                    llProject(ilLoop) = 0
                                Next ilLoop
                                blNTRWithTotal = False
                                For ilNTRCounter = ilLowerboundNTR To ilUpperboundNTR - 1 Step 1
                                    If tlNTRInfo(ilNTRCounter).bGarbage = False Then 'garbage if ntr and only want hard cost, and vice-versa
                                        For ilTemp = 1 To ilPeriod
                                             'look at each ntr record's date to see if falls into specific time period.
                                            If tlNTRInfo(ilNTRCounter).lSbfDate >= llTYStarts(ilTemp) And tlNTRInfo(ilNTRCounter).lSbfDate < llTYStarts(ilTemp + 1) Then
                                                'flag so won't write record if all values are 0
                                                If tlNTRInfo(ilNTRCounter).lSBFTotal > 0 Then
                                                    blNTRWithTotal = True
                                                    '3-28-19 determine the net amt for an NTR item if Net selected
                                                    llGrossNet = gGetGrossOrNetFromRate(tlNTRInfo(ilNTRCounter).lSBFTotal, slGrossNet, tlChfAdvtExt(ilCurrentRecd).iAgfCode, True, tlNTRInfo(ilNTRCounter).slNTRAgyCommFlag)

                                                    'llProject(ilTemp) = llProject(ilTemp) + tlNTRInfo(ilNTRCounter).lSBFTotal
                                                    llProject(ilTemp) = llProject(ilTemp) + llGrossNet                 '3-28-19
                                                End If
                                            Exit For
                                            End If
                                        Next ilTemp
                                    End If
                                Next ilNTRCounter
                                If blNTRWithTotal = True Then
                                    'mWrite looks at the tgChfAV, so must set it.Dan M 8-18-08 no longer looks at lcode, because rvf files may not have contract. Now need adf and slf
                                   ' tgChfAV.lCode = llContrCode
                                   tgChfAV.iAdfCode = tlChfAdvtExt(ilCurrentRecd).iAdfCode
                                   tgChfAV.iSlfCode(0) = tlChfAdvtExt(ilCurrentRecd).iSlfCode(0)
                                   '3-28-19
                                   
                                    mWriteAV ilCalType, ilWkOrMonth, ilWeekTYStart(), ilWeekLYStart(), ilTYStartYr, llProject(), llGenlDates(1), ilQtr
                                End If
                            End If
                        End If      'ty, include harcost or ntr, not trade
                   Else                                    'past
                        'does cnt dates span last year ?
                        llContrCode = gPaceCntr(tlChfAdvtExt(ilCurrentRecd).lCntrNo, llWeekLYStart, hmCHF, tmChf)
                        gUnpackDate tmChf.iStartDate(0), tmChf.iStartDate(1), slTemp
                        llDate = gDateValue(slTemp)
                        gUnpackDate tmChf.iEndDate(0), tmChf.iEndDate(1), slTemp
                        llDate2 = gDateValue(slTemp)

                        If llContrCode > 0 And llDate <= llLYEnd And llDate2 >= llLYStart Then
                            ilTY = False            'last year found
                        Else
                            ilTY = True
                          '  llContrCode = 0         'force not found for last year.
                        End If
                         'dan M 6-23-08 grab NTR records
                        If (blIncludeNTR Or blIncludeHardCost) And (Not (ilTY)) And (tlChfAdvtExt(ilCurrentRecd).iPctTrade <> 100) Then
                            mNTRByContract llContrCode, llDate, llDate2, tlNTRInfo(), blIncludeNTR, blIncludeHardCost
                            ilLowerboundNTR = LBound(tlNTRInfo)
                            ilUpperboundNTR = UBound(tlNTRInfo)
                            If ilLowerboundNTR <> ilUpperboundNTR Then
                                'clear array
                                For ilLoop = 1 To 53
                                    llProject(ilLoop) = 0
                                Next ilLoop
                                blNTRWithTotal = False
                                For ilNTRCounter = ilLowerboundNTR To ilUpperboundNTR - 1 Step 1
                                    If tlNTRInfo(ilNTRCounter).bGarbage = False Then 'garbage if ntr and only want hard cost, and vice-versa
                                        For ilTemp = 1 To ilPeriod
                                            'look at each ntr record's date to see if falls into specific time period.
                                                If tlNTRInfo(ilNTRCounter).lSbfDate >= llLYStarts(ilTemp) And tlNTRInfo(ilNTRCounter).lSbfDate < llLYStarts(ilTemp + 1) Then
                                                'flag so won't write record if all values are 0
                                                    If tlNTRInfo(ilNTRCounter).lSBFTotal > 0 Then
                                                        blNTRWithTotal = True
                                                        '3-28-19 need to check each ntr for agency commissionable
                                                        '3-28-19 determine the net amt for an NTR item if Net selected
                                                        llGrossNet = gGetGrossOrNetFromRate(tlNTRInfo(ilNTRCounter).lSBFTotal, slGrossNet, tlChfAdvtExt(ilCurrentRecd).iAgfCode, True, tlNTRInfo(ilNTRCounter).slNTRAgyCommFlag)
                                                      
                                                        'llProject(ilTemp) = llProject(ilTemp) + tlNTRInfo(ilNTRCounter).lSBFTotal
                                                        llProject(ilTemp) = llProject(ilTemp) + llGrossNet
                                                    End If
                                                    Exit For
                                                End If
                                        Next ilTemp
                                    End If
                                Next ilNTRCounter
                                If blNTRWithTotal = True Then
                                    'tgChfAV.lCode = llContrCode
                                    tgChfAV.iAdfCode = tlChfAdvtExt(ilCurrentRecd).iAdfCode
                                   tgChfAV.iSlfCode(0) = tlChfAdvtExt(ilCurrentRecd).iSlfCode(0)
                                    mWriteAV ilCalType, ilWkOrMonth, ilWeekTYStart(), ilWeekLYStart(), ilLYStartYr, llProject(), llGenlDates(1), ilQtr
                                End If
                            End If
                        End If

                    End If
                    'Retrieve the contract, schedule lines and flights
                    'llContrCode = gPaceCntr(tlChfAdvtext(ilCurrentRecd).lCntrNo, llWeekTYStart, hmChf, tmChf)
                    If llContrCode > 0 Then
                        'Retrieve the contract, schedule lines and flights
                        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfAV, tgClfAV(), tgCffAV())
                        'test current or last year date range
                        gUnpackDate tgChfAV.iOHDDate(0), tgChfAV.iOHDDate(1), slTemp            'convert date entered
                        llEntryDate = gDateValue(slTemp)

                        'get cnts earliest and latest dates to see if it spans the requested period
                        gUnpackDate tgChfAV.iStartDate(0), tgChfAV.iStartDate(1), slTemp       '
                        llDate = gDateValue(slTemp)
                        gUnpackDate tgChfAV.iEndDate(0), tgChfAV.iEndDate(1), slTemp
                        llDate2 = gDateValue(slTemp)
                        ilFoundOne = False
                        'Determine what years is being processed and if its within pacing period

                        If ilLoopOnYear = 1 Then
                            'is hdr start date <= This years end date and hdr end date >= this years start?
                            If llDate <= llTYEnd And llDate2 >= llTYStart Then          'Does contract dates span requested period for this year?
                                If llEntryDate <= llWeekTYStart Then
                                    ilFoundOne = True
                                    For ilTemp = 1 To 53                            'setup general buffer to use This years dates
                                        llGenlDates(ilTemp) = llTYStarts(ilTemp)
                                        llProject(ilTemp) = 0
                                    Next ilTemp
                                    ilTY = True
                                End If
                            Else
                                ilFoundOne = False
                            End If
                        Else
                            'Else                                                    'test if cnt with previous year
                                'is hdr start date <= last years end date and hdr end date >= last years start?
                                If llDate <= llLYEnd And llDate2 >= llLYStart Then          'Does contract dates span requested period for this year?
                                    If llEntryDate <= llWeekLYStart Then
                                        ilFoundOne = True
                                        For ilTemp = 1 To 53                        'setup general buffer to use last years dates
                                            llGenlDates(ilTemp) = llLYStarts(ilTemp)
                                            llProject(ilTemp) = 0
                                        Next ilTemp
                                        ilTY = False
                                    End If
                                Else
                                    ilFoundOne = False
                                End If
                            'End If
                        End If
                    Else
                        ilFoundOne = False
                    End If                                              'ilfoundOne
                End If

                'Find all contracts for this year (start of the year thru the user requested week) or last years business from the start of last year
                'thru the same week last year
                If ilFoundOne And tgChfAV.iPctTrade <> 100 Then
                    For ilClf = LBound(tgClfAV) To UBound(tgClfAV) - 1 Step 1
                        tmClf = tgClfAV(ilClf).ClfRec
                        If tmClf.sType = "S" Or tmClf.sType = "H" Then
                            'Project the monthly $ from the flights
                            gBuildFlights ilClf, llGenlDates(), 1, ilPeriod, llProject(), ilWkOrMonth, tgClfAV(), tgCffAV()
                        End If
                    Next ilClf                                      'loop thru schedule lines
                'End If                                              'ilfound
                    '3-28-19 determine the net amt for an NTR item if Net selected
                    For ilTemp = LBound(llProject) To UBound(llProject)
                        llGrossNet = gGetGrossOrNetFromRate(llProject(ilTemp), slGrossNet, tlChfAdvtExt(ilCurrentRecd).iAgfCode, False)
                        llProject(ilTemp) = llGrossNet              'alter if its net selected
                    Next ilTemp

                    If ilTY Then                            'this yr
                        mWriteAV ilCalType, ilWkOrMonth, ilWeekTYStart(), ilWeekLYStart(), ilTYStartYr, llProject(), llGenlDates(1), ilQtr
                    Else                                    'last yr
                        mWriteAV ilCalType, ilWkOrMonth, ilWeekTYStart(), ilWeekLYStart(), ilLYStartYr, llProject(), llGenlDates(1), ilQtr
                    End If
                End If                              'if found
                ilFoundOne = ilFoundAdvt            'if current year just procesed, do previous year--but restore flag
                                                    'if contract advt was found
            Next ilLoopOnYear
        End If                          ' select single contract
    Next ilCurrentRecd                                      'loop for CHF records
    sgCntrForDateStamp = ""             'init time stamp for re-entrant without returning to report list
    Erase tlChfAdvtExt, tlRvf
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmUrf)
    ilRet = btrClose(hmGrf)
End Sub
'
'
'
'           mWriteAV - write the Advertisr Variance record to disk (Grf)
'               from Contract or Receivables file
'           <input> ilCalType - 1 = std, 2 = corp
'                   ilWkOrMonth - 1 = Month, 2 = week
'                   ilWeekTYStart - btrieve date, start date of week for this year
'                   ilWeekLYStart - btrieve date, start date of week for last year
'                   ilStartYr - this year or last years Year (1998, 1999, 2000, etc)
'                   llproject - array of dollars gathered for the yr or month (from receivables or cnt)
'                   llGenlDates(1) - string containing date of first month of data
'                   ilQTr - # qtrs requested if week option
'           Created:  4/21/98
'
'           mWriteAV ilCalType, ilWkOrMonth, ilWeekTYStart(), ilWeekLYStart(), ilStartYr, llProject(),llGenlDates(1),ilQtr
'
'           5/6/99 Keep values stored in whole dollars, dont round to 1000;  do it in Crystal to maintain accuracy
Sub mWriteAV(ilCalType As Integer, ilWkOrMonth As Integer, ilWeekTYStart() As Integer, ilWeekLYStart() As Integer, ilStartYr As Integer, llProject() As Long, llGenlDates As Long, ilQtr As Integer)
Dim ilLoop As Integer
Dim llCalcGross As Long
Dim llQtrGross As Long
Dim ilTemp As Integer
Dim ilRet As Integer
Dim slMonth As String
    tmGrf.iGenDate(0) = igNowDate(0)
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)
    'tmGrf.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    'tmGrf.lChfCode = tgChfAV.lCode         'dan M removed 8-19-08  rvf may return element without a contract.
    tmGrf.iCode2 = ilCalType
    tmGrf.iAdfCode = tgChfAV.iAdfCode       'Dan M added 8-19-08
    tmGrf.iSlfCode = tgChfAV.iSlfCode(0)     'Dan M added 8-19-08
    'tmGrf.iDateGenl(0, 1) = ilWeekTYStart(0)  'Start date of week for this year
    'tmGrf.iDateGenl(1, 1) = ilWeekTYStart(1)
    'tmGrf.iDateGenl(0, 2) = ilWeekLYStart(0)  'Start date of week for this year
    'tmGrf.iDateGenl(1, 2) = ilWeekLYStart(1)
    tmGrf.iDateGenl(0, 0) = ilWeekTYStart(0)  'Start date of week for this year
    tmGrf.iDateGenl(1, 0) = ilWeekTYStart(1)
    tmGrf.iDateGenl(0, 1) = ilWeekLYStart(0)  'Start date of week for this year
    tmGrf.iDateGenl(1, 1) = ilWeekLYStart(1)
    'If ilTY Then                            'this year
    '    tmGrf.iYear = ilTYStartYr           'year (1997, 1998, etc)
    'Else
    '    tmGrf.iYear = ilLYStartYr           'last year
    'End If                                  'year (1996, 1997, etc)
    tmGrf.iYear = ilStartYr                  'last year or this year (1998, 1999, 2000,etc)
    If ilWkOrMonth = 1 Then                   'month
        slMonth = Format$(llGenlDates, "m/d/yy")
        gPackDate slMonth, tmGrf.iStartDate(0), tmGrf.iStartDate(1)
        tmGrf.sBktType = "M"
        For ilLoop = 1 To 12
            tmGrf.lDollars(ilLoop - 1) = llProject(ilLoop) \ 100      'drop pennies
            '5/6/99 round to 1000s in crystal to maintain accuracy
            'tmGrf.lDollars(ilLoop) = (tmGrf.lDollars(ilLoop) + 500) \ 1000  'round to thousands
            llCalcGross = llCalcGross + tmGrf.lDollars(ilLoop - 1)
            llQtrGross = llQtrGross + tmGrf.lDollars(ilLoop - 1)
            If ilLoop Mod 3 = 0 Then
                tmGrf.lDollars(13 + (ilLoop / 3) - 1) = llQtrGross
                llQtrGross = 0
            End If
        Next ilLoop
        If llCalcGross <> 0 Then
            'tmGrf.lDollars(13) = llCalcGross                        'year total
            tmGrf.lDollars(12) = llCalcGross                        'year total
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        End If
    Else
        tmGrf.sBktType = "W"                    'week
        'loop for as many quarters are requested
        For ilTemp = 1 To ilQtr
            For ilLoop = 1 To 13
                tmGrf.lDollars(ilLoop - 1) = llProject((ilTemp - 1) * 13 + ilLoop) \ 100    'drop pennies
                '5/6/99 round to 1000s in Crystal to maintain accuracy
                'tmGrf.lDollars(ilLoop) = (tmGrf.lDollars(ilLoop) + 500) \ 1000              'round to thousands
                llCalcGross = llCalcGross + tmGrf.lDollars(ilLoop - 1)
            Next ilLoop
            'slMonth = Format$(llGenlDates((ilTemp - 1) * 13 + 1), "m/d/yy")
            'gPackDate slMonth, tmGrf.iStartDate(0), tmGrf.iStartDate(1)
            'tmGrf.iPerGenl(1) = ilTemp          'qtr flag for sorting
            tmGrf.iPerGenl(0) = ilTemp          'qtr flag for sorting
            If llCalcGross <> 0 Then
                'tmGrf.lDollars(14) = llCalcGross            'qtr total
                tmGrf.lDollars(13) = llCalcGross            'qtr total
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            End If
            llCalcGross = 0             'init the qtr total
        Next ilTemp
    End If
    For ilLoop = 1 To 53
        llProject(ilLoop) = 0
    Next ilLoop
    llCalcGross = 0
End Sub

'********************* mNTRByContract*****************
' Dan M. 6-23-08  NTR/HardCost to pacing reports
'   llCurrentChfCode (I)
'   llStartDate(I)
'   llEndDate(I)
'   tlNTRInfo (O)
'   blIncludeNTR (I)
'   blIncludeHardCost(I)
'   tmMNF() (I)
Private Sub mNTRByContract(llCurrentChfCode As Long, llStartDate As Long, llEndDate As Long, tlNTRInfo() As NTRPacing, blIncludeNTR As Boolean, blIncludeHardCost As Boolean) '(llCurrentChfCode As Long, llStartDate As long, llEndDate As Long, tlNTRInfo() As NTRPacing, blIncludeNTR As Boolean, blIncludeHardCost As Boolean)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llLineAmount                                                                          *
'******************************************************************************************

Dim ilRet As Integer
Dim tlSbf() As SBF
Dim tlSBFTypes As SBFTypes
Dim slEndDate As String
Dim slStartDate As String
Dim ilLowEndArray As Integer
Dim ilHighEndArray As Integer
Dim ilSbfCounter As Integer
Dim ilMnfItem As Integer
Const HardCost = -1

tlSBFTypes.iNTR = True
tlSBFTypes.iImport = False
tlSBFTypes.iInstallment = False
slStartDate = Format(llStartDate, "m/d/yy")
slEndDate = Format(llEndDate, "m/d/yy")
ilRet = gObtainSBF(RptSelAv, hmSbf, llCurrentChfCode, slStartDate, slEndDate, tlSBFTypes, tlSbf(), 0)
ilLowEndArray = LBound(tlSbf)
ilHighEndArray = UBound(tlSbf)
ReDim tlNTRInfo(ilLowEndArray To ilHighEndArray) As NTRPacing
For ilSbfCounter = ilLowEndArray To ilHighEndArray - 1 Step 1
    If blIncludeHardCost And blIncludeNTR Then
        gUnpackDateLong tlSbf(ilSbfCounter).iDate(0), tlSbf(ilSbfCounter).iDate(1), tlNTRInfo(ilSbfCounter).lSbfDate
        tlNTRInfo(ilSbfCounter).lSBFTotal = tlSbf(ilSbfCounter).lGross * tlSbf(ilSbfCounter).iNoItems
        tlNTRInfo(ilSbfCounter).bGarbage = False
        tlNTRInfo(ilSbfCounter).slNTRAgyCommFlag = tlSbf(ilSbfCounter).sAgyComm             '3-28-19
    Else
        ilMnfItem = tlSbf(ilSbfCounter).iMnfItem
        ilRet = gIsItHardCost(ilMnfItem, tmMnf())
        If ((blIncludeNTR) And (ilRet <> HardCost)) Or ((blIncludeHardCost) And (ilRet = HardCost)) Then
            gUnpackDateLong tlSbf(ilSbfCounter).iDate(0), tlSbf(ilSbfCounter).iDate(1), tlNTRInfo(ilSbfCounter).lSbfDate
            tlNTRInfo(ilSbfCounter).lSBFTotal = tlSbf(ilSbfCounter).lGross * tlSbf(ilSbfCounter).iNoItems
            tlNTRInfo(ilSbfCounter).bGarbage = False
            tlNTRInfo(ilSbfCounter).slNTRAgyCommFlag = tlSbf(ilSbfCounter).sAgyComm             '3-28-19
        Else
            tlNTRInfo(ilSbfCounter).bGarbage = True
            tlNTRInfo(ilSbfCounter).slNTRAgyCommFlag = tlSbf(ilSbfCounter).sAgyComm             '3-28-19
        End If
    End If
Next ilSbfCounter
End Sub

Private Function mFillMnfArray() As Boolean
Dim ilRet As Integer
Dim ilOffSet As Integer
Dim ilUpperBound As Integer
Dim ilExtLen As Integer
Dim llNoRec As Long
Dim llRecPos As Long

mFillMnfArray = False

ilOffSet = 0
ReDim tmMnf(0 To 0) As MNF
ilUpperBound = UBound(tmMnf)
ilExtLen = Len(tmMnf(ilUpperBound))
llNoRec = gExtNoRec(ilExtLen)

btrExtClear hmMnf
ilRet = btrGetFirst(hmMnf, tmMnf(ilUpperBound), imMnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
If ilRet = BTRV_ERR_END_OF_FILE Then
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    Exit Function
End If
Call btrExtSetBounds(hmMnf, llNoRec, -1, "UC", "mnf", "")
ilRet = btrExtAddField(hmMnf, ilOffSet, imMnfRecLen)
If ilRet <> BTRV_ERR_NONE Then
    ilRet = btrClose(hmMnf)
    btrDestroy hmMnf
    Exit Function
End If
ilRet = btrExtGetNext(hmMnf, tmMnf(ilUpperBound), ilExtLen, llRecPos)
If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
    If (ilRet <> BTRV_ERR_NONE) And (ilRet <> BTRV_ERR_REJECT_COUNT) Then
        ilRet = btrClose(hmMnf)
        btrDestroy hmMnf
        Exit Function
    End If
    ilUpperBound = UBound(tmMnf)                          'precaution
    ilExtLen = Len(tmMnf(ilUpperBound))
    Do While ilRet = BTRV_ERR_REJECT_COUNT
        ilRet = btrExtGetNext(hmMnf, tmMnf(ilUpperBound), ilExtLen, llRecPos)
    Loop
    Do While ilRet = BTRV_ERR_NONE
        ilUpperBound = ilUpperBound + 1
        ReDim Preserve tmMnf(0 To ilUpperBound) As MNF
        ilRet = btrExtGetNext(hmMnf, tmMnf(ilUpperBound), ilExtLen, llRecPos)
        Do While ilRet = BTRV_ERR_REJECT_COUNT
            ilRet = btrExtGetNext(hmMnf, tmMnf(ilUpperBound), ilExtLen, llRecPos)
        Loop
    Loop
End If
mFillMnfArray = True
End Function
