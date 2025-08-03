Attribute VB_Name = "RptcrPR"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptcrpr.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

Option Explicit
Option Compare Text


'contract header
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1    'CHF record image
Dim imCHFRecLen As Integer      'CHF record length
Dim tmChf As CHF

'  Vehicle File

'  General Report File
Dim hmCbf As Integer            'prepass file handle
Dim tmCbf As CBF
Dim imCbfRecLen As Integer      'CBF record length

Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF

Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF

Dim hmDnf As Integer            'Book Name file handle
Dim imDnfRecLen As Integer      '
Dim tmDnf As DNF
Dim tmDnfSrchKey As INTKEY0     'DNF key image

Dim hmDrf As Integer            'Demo Research data
Dim imDrfRecLen As Integer
Dim tmDrf As DRF
Dim tmDrfSrchKey As LONGKEY0

Dim hmDpf As Integer            'Demo Plus Research data
Dim imDpfRecLen As Integer
Dim tmDpf As DPF

Dim hmRdf As Integer            'Dayparts file handle
Dim imRdfRecLen As Integer      'RD record length
Dim tmRdfSrchKey As INTKEY0     'RDF key image
Dim tmRdf As RDF

Dim hmDef As Integer            'Demo Plus Research data
Dim hmRaf As Integer            'Split Network Region definition

Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF

Dim lmSingleCntr As Long              'selective contr # for Billed & booked
Dim imWkSpots() As Integer    '2 years weekly spot counts, one line at a time
Dim lmWkRate() As Long        '2 years weekly rates, one line at a time
Dim lmWkPop() As Long         '2 years weekly population, one line at a time
Dim lmWkAud() As Long         '2 years weekly audience,one line at a time
Dim imAudFromSource() As Integer    '2 years weekly codes :  where the aud was obtained from (0= sold DP, 1=extra DP,
                                    '2 = time, 3 = vehicle, 4 = best fit sold DP, 5 = best fit extra dp
Dim smAudFromDesc() As String * 10 '2 years weekly DP or DRF descriptions (10 char each week)
Dim lmStartQtrs() As Long   'start dates of each quarter
Dim lmEndQtrs() As Long     'end dates of each quarter
'
'
'           mChfDatesToLong - get contract start and end dates and
'           convert to long for math
'
'           <input/Output> lldate - chf start date as long
'                           llDate2 - chf end date as long
'                           slStartDate - chf start date as a string
'                           slEndDate - chf end date as a string
'   3-19-13 remove and use gChfDatesToLong
'Sub mChfDatesToLong(llDate As Long, llDate2 As Long, slStartDate As String, slEndDate As String)
'    gUnpackDate tgChfCT.iStartDate(0), tgChfCT.iStartDate(1), slStartDate
'    If slStartDate = "" Then
'    llDate = 0
'    Else
'    llDate = gDateValue(slStartDate)
'    End If
'    gUnpackDate tgChfCT.iEndDate(0), tgChfCT.iEndDate(1), slEndDate
'    If slEndDate = "" Then
'    llDate2 = 0
'    Else
'    llDate2 = gDateValue(slEndDate)
'    End If
'End Sub
'
'
'           mCreateResearchRecap - this report is a dump of the schedule
'           lines research information.  Gathers the spot counts, rates,
'           population, and audience by week.  Due to Demo estimates,
'           different populations and audience can vary, resulting in
'           difficult proof of research numbers on the contract screen/printout.
'           A detail version of this report dumps all weeks of the order,
'           printed by the standard quarter, showing the spot counts, rate,
'           population and audience by week.
'           Each schedule line is summarized with:  Sch line, vehicle, start/end
'           dates of the line, daypart (followed by * if it is overriden),
'           book name (and book date), line population and audience.  If the
'           population or audience varies across the weeks, the overall schedule
'           line population and/or audience will show "Varies".
'           6-17-04
Public Sub gCreateResearchRecap()
Dim llContrCode As Long
Dim ilRet As Integer
Dim slStartDate As String           'cntr start date
Dim slEndDate As String             'cntr end date
Dim llStartDate As Long             'cntr start date
Dim llEndDate As Long               'cntr end date
Dim slStartQtr As String
Dim slEndQtr As String
Dim ilQtr As Integer
Dim ilLoop As Integer
Dim ilClf As Integer                'loop for sch lines
Dim llStartTemp As Long
Dim llEndTemp As Long
Dim slMonth As String
Dim slDay As String
Dim slYear As String
Dim ilCff As Integer
Dim ilDemoLoop As Integer       'loop on demo
'ReDim ilMnfDemo(1 To 4) As Integer  'up to 4 demos to process for 1 contract
ReDim ilMnfDemo(0 To 4) As Integer  'up to 4 demos to process for 1 contract. Index zero ignored
Dim ilMnfQualitative As Integer     'qualitative group if selected
Dim slCode As String
Dim slNameCode As String
Dim ilWk As Integer
Dim ilLoopOnWk As Integer
Dim slStr As String
Dim ilWkInArray As Integer
Dim llTotalSpots As Long        'total spot count for the quarter
Dim ilOVDays As Integer         'true if line has days override
Dim ilOVTimes As Integer        'true if line has time override


    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        GoTo CloseAndExit
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    ReDim tgClfCT(0 To 0) As CLFLIST
    tgClfCT(0).iStatus = -1 'Not Used
    tgClfCT(0).lRecPos = 0
    tgClfCT(0).iFirstCff = -1
    ReDim tgCffCT(0 To 0) As CFFLIST
    tgCffCT(0).iStatus = -1 'Not Used
    tgCffCT(0).lRecPos = 0
    tgCffCT(0).iNextCff = -1
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        GoTo CloseAndExit
    End If
    imClfRecLen = Len(tgClfCT(0).ClfRec)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        GoTo CloseAndExit
    End If
    imCffRecLen = Len(tgCffCT(0).CffRec)

    hmDnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        GoTo CloseAndExit
    End If
    imDnfRecLen = Len(tmDnf)

    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        GoTo CloseAndExit
    End If
    imMnfRecLen = Len(tmMnf)

    hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        GoTo CloseAndExit
    End If
    imCbfRecLen = Len(tmCbf)

    hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        GoTo CloseAndExit
    End If
    imDrfRecLen = Len(tmDrf)

    hmDpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()      '7-23-01
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        GoTo CloseAndExit
    End If
    imDpfRecLen = Len(tmDpf)

    hmDef = CBtrvTable(ONEHANDLE) 'CBtrvObj()      '7-23-01
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        GoTo CloseAndExit
    End If

    hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()      '7-23-01
    ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        GoTo CloseAndExit
    End If

    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()      '7-23-01
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        GoTo CloseAndExit
    End If
    imRdfRecLen = Len(tmRdf)


    'retrieve the contract header so the rest of the contract can be retrieved
    lmSingleCntr = 0
    lmSingleCntr = Val(RptSelPr!edcContract.Text)
    tmChfSrchKey1.lCntrNo = lmSingleCntr
    tmChfSrchKey1.iCntRevNo = 32000
    tmChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Or tmChf.lCntrNo <> lmSingleCntr Then         'exit if err or does not exist
        Exit Sub
    End If
    
    If RptSelPr!ckcUseDefault.Value = vbChecked Then
        If Not gSetFormula("UseDefaultBook", "'Y'") Then
            MsgBox "Formula UseDefault does not exist - Call Counterpoint"
            Exit Sub
        End If
    Else
        If Not gSetFormula("UseDefaultBook", "'N'") Then
            MsgBox "Formula UseDefault does not exist - Call Counterpoint"
            Exit Sub
        End If
    End If

    If RptSelPr!cbcSet1.ListIndex > 0 Then
        slNameCode = tgSocEcoCode(RptSelPr!cbcSet1.ListIndex - 1).sKey
        ilRet = gParseItem(slNameCode, 2, "\", slCode)
        ilMnfQualitative = Val(slCode)
    End If

    'get the schedule lines and flights for the current version
    llContrCode = tmChf.lCode
    ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT())

    'mChfDatesToLong llStartDate, llEndDate, slStartDate, slEndDate   'convert contract header start & end dates to long
    '3-19-13 use common routine and remove mChfDAtestoLong
    gChfDatesToLong tgChfCT.iStartDate(), tgChfCT.iEndDate(), llStartDate, llEndDate, slStartDate, slEndDate   'convert contract header start & end dates to long

    'ReDim lmStartQtrs(1 To 9) As Long   'start dates of each quarter.  Create an extra qtr to obtain the end of qtr 8
    ReDim lmStartQtrs(0 To 9) As Long   'start dates of each quarter.  Create an extra qtr to obtain the end of qtr 8. Index zero ignored
    'ReDim lmEndQtrs(1 To 8) As Long     'end dates of each quarter
    ReDim lmEndQtrs(0 To 8) As Long     'end dates of each quarter. Index zero ignored
    'determine current qtr start date and build array of 2 years quarter start/end dates.  The detail
    'information will be output by quarters
    llStartTemp = gDateValue(gObtainStartStd(slStartDate))    'get the std start date of the contract
    llEndTemp = gDateValue(gObtainEndStd(slEndDate))       'get the std end date of the contract
    slStartQtr = Format$(llStartTemp, "m/d/yy")
    gObtainYearMonthDayStr slStartQtr, True, slYear, slMonth, slDay
    Do While (slMonth <> "01") And (slMonth <> "04") And (slMonth <> "07") And (slMonth <> "10")
        slMonth = str$((Val(slMonth) - 1))
        slDay = "15"
        slStartQtr = slMonth & "/" & slDay & "/" & slYear
        gObtainYearMonthDayStr slStartQtr, True, slYear, slMonth, slDay
        slStartQtr = gObtainStartStd(slStartQtr)        'get std bdcst end date
    Loop
    llStartTemp = gDateValue(slStartQtr)
    'Setup the start dates of the 8 quarters
    For ilLoop = 1 To 9
        lmStartQtrs(ilLoop) = llStartTemp
        For ilQtr = 1 To 3 Step 1
            slStartQtr = Format$(llStartTemp, "m/d/yy")
            slEndQtr = gObtainEndStd(slStartQtr)
            slStartQtr = gIncOneDay(slEndQtr)
            llStartTemp = gDateValue(slStartQtr)
        Next ilQtr
    Next ilLoop
    'set the quarter end dates
    For ilLoop = 1 To 8
        lmEndQtrs(ilLoop) = lmStartQtrs(ilLoop + 1) - 1
    Next ilLoop

    'setup demos to process (up to 4)

    ilMnfDemo(1) = tgChfCT.iMnfDemo(0)
    If RptSelPr!rbcDemo(1).Value = True Then      'all
        ilMnfDemo(2) = tgChfCT.iMnfDemo(1)
        ilMnfDemo(3) = tgChfCT.iMnfDemo(2)
        ilMnfDemo(4) = tgChfCT.iMnfDemo(3)
    End If

    For ilDemoLoop = 1 To 4                     'loop on max 4 demos
        If ilMnfDemo(ilDemoLoop) > 0 Then                  'process the demo

            'loop thru schedule lines and gather the weekly information
            For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                tmClf = tgClfCT(ilClf).ClfRec
    
                tmCbf.sMixTypes = ""            'flag to indicate default book used
                If RptSelPr!ckcUseDefault.Value = vbChecked Then
                    
                    '3-21-19 see if the demo exist, if not, use default book if different than the current one stored in line
                    'if same default book, 0 the book reference
                    tmDnfSrchKey.iCode = tmClf.iDnfCode
    
                    ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        'book doesnt exist, find the default book
                        ilRet = gBinarySearchVef(tmClf.iVefCode)
                        If ilRet <> -1 Then
                            If tmClf.iDnfCode <> tgMVef(ilRet).iDnfCode Then
                                tgClfCT(ilClf).ClfRec.iDnfCode = tgMVef(ilRet).iDnfCode        'negate the number to know to set flag for report
                                tmCbf.sMixTypes = "D"
                            Else                'same invalid reference
                                tgClfCT(ilClf).ClfRec.iDnfCode = 0
                            End If
                        End If
                    Else            'book exists
                        ilRet = ilRet
                    End If
                End If
                
                tmClf = tgClfCT(ilClf).ClfRec

                'Arrays of weekly spots, rates, population and audience is shown on the detail version

                'ReDim imWkSpots(1 To 160) As Integer     '2+ years weekly spot counts
                'ReDim lmWkRate(1 To 160) As Long         '2+ years weekly rates
                'ReDim lmWkPop(1 To 160) As Long          '2+ years weekly population
                'ReDim lmWkAud(1 To 160) As Long          '2+ years weekly audience
                'ReDim imAudFromSource(1 To 160) As Integer  '2+ years weekly code to determine the source of the aud (DP, extra, time, etc)
                'ReDim smAudFromDesc(1 To 160) As String * 10    '2+ years weekly aud source desc from DRF or RDF
                'Index zero ignored in the arrays below
                ReDim imWkSpots(0 To 160) As Integer     '2+ years weekly spot counts
                ReDim lmWkRate(0 To 160) As Long         '2+ years weekly rates
                ReDim lmWkPop(0 To 160) As Long          '2+ years weekly population
                ReDim lmWkAud(0 To 160) As Long          '2+ years weekly audience
                ReDim imAudFromSource(0 To 160) As Integer  '2+ years weekly code to determine the source of the aud (DP, extra, time, etc)
                ReDim smAudFromDesc(0 To 160) As String * 10    '2+ years weekly aud source desc from DRF or RDF
                'For ilLoop = 1 To 160                     'init the pop & audience to know if a week has zero pop or not week not active
                For ilLoop = 0 To 160                     'init the pop & audience to know if a week has zero pop or not week not active
                    lmWkPop(ilLoop) = -1
                    lmWkAud(ilLoop) = -1
                    imAudFromSource(ilLoop) = -1
                    smAudFromDesc(ilLoop) = "          "
                Next ilLoop

                ilOVDays = False                        'assume no days overridden until flights processed
                ilOVTimes = False                       'assume no times overridden
                 'get the daypart record
                tmRdfSrchKey.iCode = tmClf.iRdfCode
                ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching r/c recd
                If ilRet <> BTRV_ERR_NONE Then
                    tmCbf.sDysTms = "Missing DP"
                Else
                    tmCbf.sDysTms = Trim$(tmRdf.sName)
                End If

                'check if times overridden.  Schline has the overridden times
                If ((tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And ((tmClf.iEndTime(0) <> 1) Or (tmClf.iEndTime(1) <> 0)) Then
                    ilOVTimes = True
                End If

                'loop on flights
                ilCff = tgClfCT(ilClf).iFirstCff
                Do While ilCff <> -1
                    tmCff = tgCffCT(ilCff).CffRec

                    mProcessFlight ilMnfDemo(ilDemoLoop), ilMnfQualitative, ilOVDays                                    '
                    ilCff = tgCffCT(ilCff).iNextCff             'get next flight record from mem
                Loop                                            'while ilcff <> -1

                'All rates, spots, pop and audience have been gathered for each week
                gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                tmCbf.lGenTime = lgNowTime
                tmCbf.iGenDate(0) = igNowDate(0)
                tmCbf.iGenDate(1) = igNowDate(1)
                tmCbf.lChfCode = tgChfCT.lCode                'contract internal code
                tmCbf.iOurShare = ilMnfDemo(ilDemoLoop)        'demo code
                tmCbf.iMnfGroup = ilMnfQualitative              'qualitative group code (socio economic)
                'tmCbf.lPop & tmCbf.lAvgAud are set to either varying pop/aud (-1) or the actual pop/audience


                tmCbf.lLineNo = tmClf.iLine             'line #
                tmCbf.iDnfCode = tmClf.iDnfCode         'book name
                tmCbf.iStartDate(0) = tmClf.iStartDate(0)   'sch line start date
                tmCbf.iStartDate(1) = tmClf.iStartDate(1)
                tmCbf.iEndDate(0) = tmClf.iEndDate(0)       'sch line end date
                tmCbf.iEndDate(1) = tmClf.iEndDate(1)
                tmCbf.iVefCode = tmClf.iVefCode             'vehicle code
                tmCbf.sType = tmClf.sType                   'vehicle type (for package notation)
                tmCbf.lCurrMod = tmClf.iPkLineNo            'package line reference
                tmCbf.lPop = -1               'init for varying pop
                tmCbf.lAvgAud = -1             ' init for varying audience
                If ilOVDays Or ilOVTimes Then   'either days or times were overridden
                    tmCbf.sDysTms = Trim(tmCbf.sDysTms) & "*"   'override exists, flag the daypart
                End If

                If tmClf.iDnfCode > 0 Then
                    For ilLoop = 1 To 160
                        'dont count packages for varying pop & aud
                        If tmClf.sType <> "E" And tmClf.sType <> "O" And tmClf.sType <> "A" Then
                            If tmCbf.lPop < 0 Then
                                If lmWkPop(ilLoop) >= 0 Then             'active week, otherwise ignore
                                    tmCbf.lPop = lmWkPop(ilLoop)
                                End If
                            Else
                                If ((tmCbf.lPop <> lmWkPop(ilLoop)) And (lmWkPop(ilLoop) <> -1)) Then '3-20-19 And (lmWkPop(ilLoop) <> 0 And lmWkPop(ilLoop) <> -1) Then
                                    tmCbf.lPop = 0                          'varying populations
                                End If
                            End If
                            If tmCbf.lAvgAud < 0 Then
                                If lmWkAud(ilLoop) >= 0 Then           'active week, otherwise ignore
                                    tmCbf.lAvgAud = lmWkAud(ilLoop)
                                End If
                            Else
                                If ((tmCbf.lAvgAud <> lmWkAud(ilLoop)) And (lmWkAud(ilLoop) <> -1)) Then    '3-20-19 And (lmWkAud(ilLoop) <> 0 And lmWkAud(ilLoop) <> -1) Then
                                    tmCbf.lAvgAud = 0                       'varying audience
                                End If
                            End If
                        End If
                    Next ilLoop
                Else
                    tmCbf.lPop = 0
                    tmCbf.lAvgAud = 0
                End If

                ilWkInArray = 0                         'index to week 1 - 160, running index as each qtr is processed
                If RptSelPr!rbcSel(0).Value = True Then              'detail
                    For ilQtr = 1 To 8
                        'Calculate # of weeks in this quarter
                        ilWk = (lmStartQtrs(ilQtr + 1) - lmStartQtrs(ilQtr)) / 7
                        gPackDateLong lmStartQtrs(ilQtr), tmCbf.iStartQtr(0), tmCbf.iStartQtr(1)
                        llTotalSpots = 0            'keep count of spots in qtr, dont write out if zero
                        tmCbf.iTotalWks = ilWk   'set the # of weeks in this quarter

                        For ilLoopOnWk = 1 To ilWk
                            llTotalSpots = llTotalSpots + imWkSpots(ilWkInArray + ilLoopOnWk)
                            If ilLoopOnWk = 14 Then         '14 weeks in this qtr, not enough buckets in each of the arrays so put
                                                                'the 14th value in different fields
'                                tmCbf.lMonth(1) = lmWkRate(ilWkInArray + ilLoopOnWk)     'Weekly Rates: move for one of 160 buckets in buckets 1-13
'                                tmCbf.lMonth(2) = lmWkPop(ilWkInArray + ilLoopOnWk)     'Weekly Populations: move for one of 160 buckets in buckets 1-13
'                                tmCbf.lMonth(3) = lmWkAud(ilWkInArray + ilLoopOnWk)     'Weekly Audience: move for one of 160 buckets in buckets 1-13
'                                tmCbf.lMonth(4) = imWkSpots(ilWkInArray + ilLoopOnWk)     'Weekly Spots: move for one of 160 buckets in buckets 1-13
                                tmCbf.lQGRP = lmWkRate(ilWkInArray + ilLoopOnWk)     'Weekly Rates: move for one of 160 buckets in buckets 1-13
                                tmCbf.lQCPP = lmWkPop(ilWkInArray + ilLoopOnWk)     'Weekly Populations: move for one of 160 buckets in buckets 1-13
                                tmCbf.lQCPM = lmWkAud(ilWkInArray + ilLoopOnWk)     'Weekly Audience: move for one of 160 buckets in buckets 1-13
                                tmCbf.lQGrimp = imWkSpots(ilWkInArray + ilLoopOnWk)     'Weekly Spots: move for one of 160 buckets in buckets 1-13
                                tmCbf.lVQGRP = imAudFromSource(ilWkInArray + ilLoopOnWk)    'source of where aud came from (dp, extra, time, etc)
                                'tmCbf.lVQCPP = lmAudFromCode(ilWkInArray + ilLoopOnWk)       'internal code of RDF or DRF, could be 0
                            Else
                                tmCbf.lValue(ilLoopOnWk - 1) = lmWkRate(ilWkInArray + ilLoopOnWk)   'Weekly Rates: move for one of 160 buckets in buckets 1-13
                                tmCbf.lWkVehGrp(ilLoopOnWk - 1) = lmWkPop(ilWkInArray + ilLoopOnWk)   'Weekly Populations: move for one of 160 buckets in buckets 1-13
                                tmCbf.lWkCntGrp(ilLoopOnWk - 1) = lmWkAud(ilWkInArray + ilLoopOnWk)  'Weekly Audience: move for one of 160 buckets in buckets 1-13
                                tmCbf.lWeek(ilLoopOnWk - 1) = imWkSpots(ilWkInArray + ilLoopOnWk)  'Weekly Spots: move for one of 160 buckets in buckets 1-13
                                tmCbf.lMonthUnits(ilLoopOnWk - 1) = imAudFromSource(ilWkInArray + ilLoopOnWk)  'source of where aud came from (dp, extra, time, etc)
                                'tmCbf.lMonth(ilLoopOnWk) = lmAudFromCode(ilWkInArray + ilLoopOnWk)       'internal code of RDF or DRF, could be 0
                            End If
                            'the Aud source description must be handled with each week since there is no
                            'continuous array that the string can be placed into
                            
                            
                            'tmcbf.sSortField1 = Wk 1 (1-10), Wk 2 (11-20)
                            If ilLoopOnWk = 1 Then
                                tmCbf.sSortField1 = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            ElseIf ilLoopOnWk = 2 Then
                                Mid(tmCbf.sSortField1, 11, 10) = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            
                            'tmcbf.sSortField2 = Wk 3 (1-10), Wk 4 (11-20)
                            ElseIf ilLoopOnWk = 3 Then
                                tmCbf.sSortField2 = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            ElseIf ilLoopOnWk = 4 Then
                                Mid(tmCbf.sSortField2, 11, 10) = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            
                            'tmCbf.sSurvey = Wk5 (1-10), wk6 (11-20), wk7 (21-30)
                            ElseIf ilLoopOnWk = 5 Then
                                tmCbf.sSurvey = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            ElseIf ilLoopOnWk = 6 Then
                                Mid(tmCbf.sSurvey, 11, 10) = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            ElseIf ilLoopOnWk = 7 Then
                                Mid(tmCbf.sSurvey, 21, 10) = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            
                            'tmCbf.sProduct = wk 8 (1-10), wk9 (11-20), wk10 (21-30)
                            ElseIf ilLoopOnWk = 8 Then
                                tmCbf.sProduct = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            ElseIf ilLoopOnWk = 9 Then
                                Mid(tmCbf.sProduct, 11, 10) = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            ElseIf ilLoopOnWk = 10 Then
                                Mid(tmCbf.sProduct, 21, 10) = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            
                            'tmCbf.sPrefDT = wk 11 (1-10), wk12 (11-20), wk13 (21-30), wk14 (31-40)
                            ElseIf ilLoopOnWk = 11 Then
                                tmCbf.sPrefDT = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            ElseIf ilLoopOnWk = 12 Then
                                Mid(tmCbf.sPrefDT, 11, 10) = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            ElseIf ilLoopOnWk = 13 Then
                                Mid(tmCbf.sPrefDT, 21, 10) = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            ElseIf ilLoopOnWk = 14 Then
                                Mid(tmCbf.sPrefDT, 31, 10) = smAudFromDesc(ilWkInArray + ilLoopOnWk)
                            End If
                                
                        Next ilLoopOnWk
                        If llTotalSpots <> 0 Then
                            ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                        End If
                        ilWkInArray = ilWkInArray + ilWk
                    Next ilQtr
                Else
                    ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                End If

            Next ilClf                  'process next sched line
        End If                          'ilMnfDemo(ilDemoLoop) > 0
    Next ilDemoLoop                     'next demo

    Erase imWkSpots
    Erase lmWkRate
    Erase lmWkPop
    Erase lmWkAud
    Erase imAudFromSource
    Erase smAudFromDesc
    Erase lmStartQtrs, lmEndQtrs

CloseAndExit:
    'close all files
     btrDestroy hmRaf
     btrDestroy hmDef
     btrDestroy hmRdf
     btrDestroy hmDpf
     btrDestroy hmDrf
     btrDestroy hmCbf
     btrDestroy hmMnf
     btrDestroy hmDnf
     btrDestroy hmCff
     btrDestroy hmClf
     btrDestroy hmCHF
     'ilRet = btrClose(hmRdf)
     'ilRet = btrClose(hmDpf)
     'ilRet = btrClose(hmDrf)
     'ilRet = btrClose(hmCbf)
     'ilRet = btrClose(hmMnf)
     'ilRet = btrClose(hmDnf)
     'ilRet = btrClose(hmCff)
     'ilRet = btrClose(hmClf)
     'ilRet = btrClose(hmChf)
     Exit Sub
End Sub

'
'       mProcessFlight - process the flight week by week to obtain the array by week of
'        spot count, $, audience and population
'
'       <input>
'               ilMnfDemo - demo mnf code
'               ilMnfQualitative - socio economic mnf code if requested
'               ilOVDays - set to true if days found to be overridden
'
Public Sub mProcessFlight(ilMnfDemo As Integer, ilMnfQualitative As Integer, ilOVDays As Integer)
Dim ilLoop As Integer
Dim llStartTemp As Long
Dim llEndTemp As Long
Dim slStr As String
Dim llSpots As Long
Dim llDate As Long
Dim llDateLoop As Long
Dim ilWkInx As Integer
Dim ilDay As Integer
Dim ilRet As Integer
Dim llOvStartTime As Long
Dim llOvEndTime As Long
Dim llAvgAud As Long
Dim llPopEst As Long
Dim llPop As Long
Dim ilDemoAvgAudDays(0 To 6) As Integer
Dim ilAudFromSource As Integer
Dim llAudFromCode As Long
Dim ilDays(0 To 6) As Integer
Dim slDays(0 To 6) As String * 1
Dim slDPDays As String
Dim slStartTime As String
Dim slDesc As String * 10
Dim slStrippedDays As String * 10
Dim slStrippedTime As String * 10
Dim ilPos As Integer
Dim ilLen As Integer

    For ilLoop = 0 To 6                 'init all days to not airing, setup for research results later
        ilDemoAvgAudDays(ilLoop) = False        'initalize valid days of schedule line
    Next ilLoop
    gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr     'start date of flight
    llStartTemp = gDateValue(slStr)

    gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr         'end date of flight
    llEndTemp = gDateValue(slStr)

    'cannot be a Cancel Before Start
    If (llEndTemp >= llStartTemp) Then
        'backup start date to Monday
        ilLoop = gWeekDayLong(llStartTemp)
        Do While ilLoop <> 0
            llStartTemp = llStartTemp - 1
            ilLoop = gWeekDayLong(llStartTemp)      'retain illoop for start day of week if daily buy
        Loop

        'loop on each week of the flight
        For llDate = llStartTemp To llEndTemp Step 7
            'Loop on the number of weeks in this flight
            'calc week into of this flight to accum the spot count
            If tmCff.sDyWk = "W" Then            'weekly
                llSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                For ilDay = 0 To 6 Step 1
                    If (llDate + ilDay >= llStartTemp) And (llDate + ilDay <= llEndTemp) Then
                        If tmCff.iDay(ilDay) > 0 Or tmCff.sXDay(ilDay) = "1" Then
                            ilDemoAvgAudDays(ilDay) = True          'set the valid days for avg aud routine
                        End If
                    End If
                Next ilDay
            Else                                        'daily
                If llDate + 6 < llEndTemp Then           'we have a whole week
                    llSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)
                    For ilDay = 0 To 6 Step 1
                        If tmCff.iDay(ilDay) > 0 Then
                            ilDemoAvgAudDays(ilDay) = True
                        End If
                    Next ilDay
                Else
                    For llDateLoop = llStartTemp To llEndTemp
                        ilDay = gWeekDayLong(llDateLoop)
                        llSpots = llSpots + tmCff.iDay(ilDay)
                        If tmCff.iDay(ilDay) > 0 Then
                            ilDemoAvgAudDays(ilDay) = True
                            llSpots = llSpots + tmCff.iDay(ilDay)
                        End If
                    Next llDateLoop
                End If
            End If

            ilWkInx = (llDate - lmStartQtrs(1)) \ 7 + 1
            If ilWkInx > 0 Then
                imWkSpots(ilWkInx) = llSpots                  'spots in week
                '7-22-14 option to show either actual spot rate or proposal (for debugging purposes on various reports)
                If RptSelPr!rbcWhichRate(0).Value = True Then
                    lmWkRate(ilWkInx) = tmCff.lActPrice         'spot price
                Else
                    lmWkRate(ilWkInx) = tmCff.lPropPrice         'Proposal price
                End If
            
                If tmClf.iDnfCode > 0 Then
                    ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, tmClf.iDnfCode, ilMnfQualitative, ilMnfDemo, llPop)
                Else
                    llPop = 0
                End If
                'test for day overrides
                For ilLoop = 0 To 6 Step 1
                    'If tmRdf.sWkDays(7, ilLoop + 1) = "Y" Then             'this day is a valid day for DP
                    If tmRdf.sWkDays(6, ilLoop) = "Y" Then             'this day is a valid day for DP
                        If ilDemoAvgAudDays(ilLoop) = False Then
                            ilOVDays = True
                            Exit For
                        End If
                    End If
                Next ilLoop

                If tmClf.sType <> "E" And tmClf.sType <> "A" And tmClf.sType <> "O" Then        'populations dont exist for package lines
                    lmWkPop(ilWkInx) = llPop

                    If tmClf.iStartTime(0) = 1 And tmClf.iStartTime(1) = 0 Then
                        llOvStartTime = 0
                        llOvEndTime = 0
                    Else
                        'override times exist
                        gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llOvStartTime
                        gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llOvEndTime
                    End If
                    ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmClf.iDnfCode, tmClf.iVefCode, ilMnfQualitative, ilMnfDemo, llDate, llDate, tmClf.iRdfCode, llOvStartTime, llOvEndTime, ilDemoAvgAudDays(), tmClf.sType, tmClf.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                    lmWkAud(ilWkInx) = llAvgAud
                    imAudFromSource(ilWkInx) = ilAudFromSource      'source 0 = sold dp, 1 = extra, 2 = time, 3 = veh, 4 = bestfit sold d, 5 = bestfit extra dp
                    If tgSpf.sDemoEstAllowed = "Y" Then     'estimtes allowed
                        lmWkPop(ilWkInx) = llPopEst
                    End If
                    'get the dp name or the days/times from DRF
                    If ilAudFromSource = 0 Or ilAudFromSource = 4 Then
                        'get the daypart record
                        tmRdfSrchKey.iCode = llAudFromCode
                        ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching r/c recd
                        If ilRet <> BTRV_ERR_NONE Then
                            smAudFromDesc(ilWkInx) = "?" & Trim$(str(llAudFromCode))
                        Else
                            'strip out the " "
                            ilPos = 1
                            slStrippedDays = ""
                            For ilLoop = 1 To 10
                                If Mid(tmRdf.sName, ilLoop, 1) <> " " Then
                                    Mid(slStrippedDays, ilPos, 1) = Mid(tmRdf.sName, ilLoop, 1)
                                    ilPos = ilPos + 1
                                End If
                            Next ilLoop
                            smAudFromDesc(ilWkInx) = Trim$(slStrippedDays)
                        End If
                    'get the drf demo research description
                    ElseIf ilAudFromSource = 1 Or ilAudFromSource = 5 Then      'source is extra dp or best fit extra dp
                        tmDrfSrchKey.lCode = llAudFromCode
                        ilRet = btrGetEqual(hmDrf, tmDrf, imDrfRecLen, tmDrfSrchKey, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching r/c recd
                        If ilRet <> BTRV_ERR_NONE Then
                            smAudFromDesc(ilWkInx) = "?" & Trim$(str(llAudFromCode))
                        Else
                            'determine days and format
                            For ilLoop = 0 To 6
                                If tmDrf.sDay(ilLoop) = "Y" Then
                                    ilDays(ilLoop) = 1
                                Else
                                    ilDays(ilLoop) = 0
                                End If
                            Next ilLoop
                            slDPDays = gDayNames(ilDays(), slDays(), 1, slStr)
                            'strip out the "-"
                            ilLen = Len(slDPDays)
                            ilPos = 1
                            slStrippedDays = ""
                            For ilLoop = 1 To ilLen
                                If Mid(slDPDays, ilLoop, 1) <> "-" Then
                                    Mid(slStrippedDays, ilPos, 1) = Mid(slDPDays, ilLoop, 1)
                                    ilPos = ilPos + 1
                                End If
                            Next ilLoop

                            'determine start time
                            gUnpackTime tmDrf.iStartTime(0), tmDrf.iStartTime(1), "A", "2", slStartTime
                             'strip out the "M" from AM or PM
                          
                            ilPos = 1
                            slStrippedTime = ""
                            For ilLoop = 1 To 10
                                If Mid(slStartTime, ilLoop, 1) <> "M" Then
                                    Mid(slStrippedTime, ilPos, 1) = Mid(slStartTime, ilLoop, 1)
                                    ilPos = ilPos + 1
                                End If
                            Next ilLoop
                            slDesc = Trim$(slStrippedDays) & Trim$(slStrippedTime)
                            
                            'determine end time
                            gUnpackTime tmDrf.iEndTime(0), tmDrf.iEndTime(1), "A", "2", slStartTime
                             'strip out the "M" from AM or PM

                            ilPos = 1
                            slStrippedTime = ""
                            For ilLoop = 1 To 10
                                If Mid(slStartTime, ilLoop, 1) <> "M" Then
                                    Mid(slStrippedTime, ilPos, 1) = Mid(slStartTime, ilLoop, 1)
                                    ilPos = ilPos + 1
                                End If
                            Next ilLoop

                            smAudFromDesc(ilWkInx) = Trim$(slDesc) & "-" & Trim$(slStrippedTime)
                        End If
                    Else
                        smAudFromDesc(ilWkInx) = "          "
                    End If
                End If

            End If                                      'ilwkinx > 0
        Next llDate                                     'for llDate = llFltStart To llFltEnd
    End If
End Sub
