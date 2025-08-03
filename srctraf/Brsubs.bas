Attribute VB_Name = "BRSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Brsubs.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmGsfSrchKey1                 imRvfRecLen                   tmSbfSrchKey              *
'*  tmSbfSrchKey2                 imSbfRecLen                   imRafRecLen               *
'*  imSefRecLen                   imTxrRecLen                   imShfRecLen               *
'*  imMktRecLen                   imVlfRecLen                   imAttRecLen               *
'*                                                                                        *
'*                                                                                        *
'* Private Procedures (Removed)                                                           *
'*  mGetUniquePkgVehPop                                                                   *
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
'
' Description:
'   Generate the Pre-pass file (cbf.btr) required to produce a
'   printed contract from Traffic Snapshot
Option Explicit
Option Compare Text
'7/30/19: Contract difference from order/proposal screen
Public lgCurrChfCode As Long
Public lgPrevChfCode As Long
'Global igPdStartDate(0 To 1) As Integer
'Global sgPdType As String * 1
'Global igNowDate(0 To 1) As Integer
'Global igNowTime(0 To 1) As Integer
'Global igYear As Integer                'budget year used for filtering
'Global igMonthOrQtr As Integer          'entered month or qtr
'Arrays for BR Generation
'Global lgPrintedCnts() As Long             'table to maintain the contr pointers
'Global lgNowtime As Long                '10-30-01                                                'when contracts are finished printing, update print flag
                                                'from this table
'Dim lmMatchingCnts() As Long            'stored contr numbers to process
'Dim imSpotsByWeek() As Integer
Dim imMaxQtrs As Integer                '4-12-19 Default to 8 quarters to process, but change for each contract to speed up processing
Const MAXWEEKSFOR2YRS = 105
Dim imDnfCodes() As Integer
''Dim imAirWks(1 To 104) As Integer     'flag of weeks with spots airing, built after all lines
'Dim imAirWks(1 To MAXWEEKSFOR2YRS) As Integer       '11-14-11 account for 14 week qtr
Dim imAirWks(0 To MAXWEEKSFOR2YRS) As Integer       '11-14-11 account for 14 week qtr. Index zero ignored
                                      'from the contract is gathered, then build unique airing weeks
Dim imProcessFlag() As Integer        'flag to indicate what to do with the line: what processing mods (differences)
                                      'all of history is built, but should not be shown.  Also, all current lines
                                      'must be examined for the overall flight weeks.  0=ignore line altogether,
                                      '1=previous revision (use for differences), 2=current line but not on
                                      'same revision as header, don't show, 3=curent line (no previous revisions),
                                      'show on BR, 4 = current line same as header revision (show on BR)
'Dim tlChfAdvtExt() As ChfAdvtExt
Dim tlMMnf() As MNF                    'array of MNF records for specific type
'The following arrays are built by the schedule line for as many weeks as there are in the order
'Dim imWeeksPerQtr(1 To 8) As Integer           '11-11-11
Dim imWeeksPerQtr(0 To 8) As Integer           '11-11-11. Index zero ignored

Dim lmWklySpots() As Long          'array of spots by week for one sched line, reqd for research results
Dim lmWklyRates() As Long             'array of rates by week for one sch line, reqd for research results
Dim lmAvgAud() As Long                'array of avg aud by week for one sch line, reqd for research
Dim lmPopEst() As Long                'array of pop by week
Dim lmPop() As Long                   'array of population by sch line (1 entry per line)
Dim lmPopPkg() As Long                '11-24-04 retain separate population for the package vehicles because user has same
                                      'hidden vehicle as pkg vehicle reference
Dim lmPopPkgByLine() As Long          '11-11-05 population by pkg line (i.e. whenmore than 1 pkg vehicle exists for 2 packages )
Dim tmLRch() As RESEARCHINFO         '10-30-01 chged from ResearchList to ResearchInfo for V5.0 : research data for current revision of order by unique sch line, rate & daypart
Dim tmVCost() As Long
Dim tmVRtg() As Integer
Dim tmVGrimp() As Long
Dim tmVGRP() As Long
Dim tmVehQtrList() As VEHQTRLIST        'list of vehicles and their quarterly cpp, cpms, grimps and grps (max 2 years , 8 qtrs)
Dim tmPkVCost() As Long
Dim tmPkVRtg() As Integer
Dim tmPkVGrimp() As Long
Dim tmPkVGRP() As Long
Dim tmPkLnCost() As Long
Dim tmPkLnGrimp() As Long
Dim tmPkLnGRP() As Long
Dim tmWkCntGrps() As WEEKLYGRPS         '2-18-00
Dim tmWkCntVGrps() As WEEKLYGRPS        '2-18-00
Dim tmWkHiddenVGrps() As WEEKLYGRPS     '2-20-00
Dim tmWkPkgVGrps() As WEEKLYGRPS        '2-20-00

'Dim tmPkLnQtrList() As VEHQTRLIST        'list of pkg lines and their quarterly cpp, cpms, grimps and grps (max 2 years , 8 qtrs)
'end of BR Generation Arrays
Dim tmLnr() As LNR              'arry of unique line & rates for BR
Dim hmCHF As Integer            'Contract header file handle
Dim hmTChf As Integer           'secondary contr header handle, so get next is not destroyed
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim tmPChf As CHF                  'CHF - previous version header (for Differences BR)
Dim hmClf As Integer            'Contract line file handle
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim tmPclf() As CLFLIST             'CLF previous version lines (for differences BR)
Dim tmCClf() As CLFLIST             'CLF current version lines (for differences BR)
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim tmCCff() As CFFLIST         'CFF current version lines (for differences BR)
Dim tmPCff() As CFFLIST         'CFF previous version lines (for differences BR)

'10-20-05   sports files
Dim tmCgf As CGF
Dim imCgfRecLen As Integer
Dim hmCgf As Integer
Dim tmCgfSrchKey1 As CGFKEY1
Dim tmGsf As GSF
Dim imGsfRecLen As Integer
Dim hmGsf As Integer

Dim hmAdf As Integer            'Advertisr file handle
Dim imAdfRecLen As Integer      'ADF record length
Dim tmAdfSrchKey As INTKEY0     'ADF key image
Dim tmAdf As ADF
Dim hmAgf As Integer            'Agency file handle
Dim imAgfRecLen As Integer      'AGF record length
Dim tmAgfSrchKey As INTKEY0     'AGF key image
Dim tmAgf As AGF
Dim hmSof As Integer            'Office file handle
Dim imSofRecLen As Integer      'SOF record length
Dim tmSof As SOF
Dim hmSlf As Integer            'Salesperson file handle
Dim imSlfRecLen As Integer      'SLF record length
Dim tmSlfSrchKey As INTKEY0     'SLF key image
Dim tmSlf As SLF
Dim hmUrf As Integer            'User file handle
Dim imUrfRecLen As Integer      'URF record length
Dim tmUrf As URF
Dim hmRdf As Integer            'Dayparts file handle
Dim imRdfRecLen As Integer      'RD record length
Dim tmRdfSrchKey As INTKEY0     'RDF key image
Dim tmRdf As RDF
#If programmatic = 1 Then
Dim tmRcf As RCF
#End If
Dim tmMRif() As RIF
Dim tmMRDF() As RDF

Dim tmAnf As ANF
Dim tmAnfTable() As ANF
Dim hmAnf As Integer
Dim imAnfRecLen As Integer

'Dim tmAvRdf() As RDF            'array of dayparts
'Dim tmRifSorts() As RIF         'array of Rate Card items to obtain sort fields & whether to use Base or show on Rept
Dim hmDnf As Integer            'Demo file handle
Dim imDnfRecLen As Integer      'DNF record length
Dim tmDnfSrchKey As INTKEY0     'DNF key image
Dim tmDnf As DNF
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnfSrchKey As INTKEY0
Dim tmMnf As MNF
'Dim tmMnfList() As MNFLIST        'array of mnf codes for Missed reasons and billing rules
Dim hmCxf As Integer            'Comment file handle
Dim imCxfRecLen As Integer      'CXF record length
Dim tmCxfSrchKey As LONGKEY0    'CXF key record image
Dim tmCxf As CXF
Dim hmCbf As Integer            'Contract BR file handle
Dim imCbfRecLen As Integer      'CBF record length
Dim tmCbf As CBF
Dim tmZeroCbf As CBF
Dim tmPkgCbf As CBF             'TTP 8410 for holding package header & replace with hidden lines

'4-23-13 File to hold the string of book names for a pkg on the contract summary (with research info), as well as the split network list shown on the detail page
'Split Network data matches on chfcode and time gen
'Package book names will match on chfcode, date/time gen, and seq = -1
Dim hmTxr As Integer            'Text string file handle
Dim tmTxr As TXR
Dim imTxrRecLen As Integer

'  Rating Book File
Dim hmDrf As Integer        'Rating book file handle
Dim tmDrf As DRF            'DRF record image
Dim imDrfRecLen As Integer  'DRF record length
'  Demo Plus Book File       '7-24-01
Dim hmDpf As Integer        'Demo Plus book file handle
Dim hmDef As Integer
Dim imDiffExceeds104Wks As Integer      '3-9-06 flag to indicate over 104 weeks on differences only version
Dim imHiddenOverride As Integer         '3-10-06 hidden overrides ignored for research

'  Receivables File
Dim hmRvf As Integer        'receivables file handle
Dim tmRvf As RVF            'RVF record image

Dim hmPhf As Integer        'receivables file handle

Dim hmSbf As Integer        'SBF file handle
Dim tmSbf As SBF            'SBF record image

'6-19-07 files for contract print to show list of split networks
Dim hmRaf As Integer        'RAF file handle
Dim hmSef As Integer        'SEF file handle
'Dim hmTxr As Integer        'TXR file handle
Dim hmShf As Integer        'SHTT file handle
Dim hmMkt As Integer        'MKT file handle
Dim hmVLF As Integer        'VLF file handle
Dim hmAtt As Integer        'ATT file handle

'12-11-20  Podcast cpm file
Dim hmPcf As Integer            'cpm podcast handle
Dim tmPcf() As PCF
Dim imPcfRecLen As Integer    'PCF record length
Dim hmThf As Integer            'cpm Pod Target header handle
Dim tmThf As PCF              'Thf record image
Dim imThRecLen As Integer    'Thf record length
Dim hmTif As Integer            'cpm pod Targe items handle
Dim tmTif As PCF              'Tif record image
Dim imTifRecLen As Integer    'Tif record length

Dim lmMonthlyTNet() As Long        't-net monthly amts for max 3 years of contract
Dim lmContractSpots As Long                 'total contract spot count

Dim tmCPM_IDs() As CPM_BR         '12-16-20 podcast line IDs data for BR
Dim tmCPMSummary() As CPM_BR        '12-16-20 podcast vehicle summary data for BR

Dim tmTranTypes As TRANTYPES
Type BRSELECTIONS
    lActiveStart As Long       'Active Start Date -for all mode, include cnts whose latest version are active
                                'for user requested dates
    lActiveEnd As Long         'Active End Date
    lEnterStart As Long        'Entered Start Date - for all mode, include cnts whose latest version entered are spanning
                                'the user requested dates
    lEnterEnd As Long          'Entered End Date
    iDetail As Integer          'true if include detail
    iSummary As Integer         'true if include summary
    lDiffChfCode As Long        'if difference,  contr code to compare against
    iDiffOnly As Integer        'true if difference only
    iThisCntMod As Integer      'for printables, if show mods as differences- Current cnt to be printed as difference
    iPrintables As Integer     'True if Printables only
    iShowMods As Integer       'For Printables only - True if to show Mods as differences
    iPropOrOrder As Integer    '0 = Proposal, 1 = Order
    iAllDemos As Integer       'True = All demos
    iShowRates As Integer      'True if include rates
    iShowResearch As Integer   'True if include research
    iShowProof As Integer      'True if Show Proof (include hidden lines)
    iCorpOrStd As Integer      '0 = Corp, 1 = std
    iWhichSort As Integer      '0 = Advt, 1 = slsp, 2 = agy
    lPrevSpots As Long      '1-14-02 chg from integer to long: For diff only, totals spots on previous version
    lPrevGross As Long         'For diff only, total gross on previous version
    iCurrTotWks As Integer     'total weeks (start to end) (printed on BR, ie 35 airing wks over 70 totals weeks)
    iCurrAirWks As Integer     'total airing weeks  (printed on BR)
    iGenDate(0 To 1) As Integer    'CBF generation key
    iGenTime(0 To 1) As Integer    'CBF generation key
    sSnapshot As String * 1         'flag to indicate Snapshot:   blank = from reports, S = from snapshot
    iSocEcoMnfCode As Integer       '10-29-03 social economic mnf code
    iShowSplits As Integer      '2-13-04 true if show comm splits
    iShowNTRBillSummary As Integer  '2-2-10 option to merge ntr bill summary with air time
    iShowNetAmtOnProps As Integer   '2-3-10 show agy comm and net amt on proposals
    iShowProdProt As Integer        '8-25-15 show product protection categories (competitives)
    iShowAct1 As Integer        'TTP 10382 - Contract report: Option To not show Act1 codes on PDF
End Type

Dim tmAuditInfo() As AUDITINFO

Public Type AUDITINFO
    iVefCode As Integer
    iType As Integer            '0 = air time, 1 = ntr
    iMnfCode As Integer         'ntr type
    lGross As Long              'gross $
    lAgyComm As Long            'agency comm
    lNet As Long                'net amount
    lMerch As Long              'merch $
    lPromo As Long              'promotions $
    lAcquisition As Long        'acquisition $
    lTNet As Long               'triple net
    lRateCard As Long           'rate card $
    lSpots As Long
    'lMonthly(1 To 36) As Long   '3 years monthly $. Not used
End Type

Type PKGVEHICLEFORSURVEYLIST            '12-19-13  this is only for the summary version to show book names (from hidden) for package vehicle
    iPkgVefCode As Integer              'pkg vehicle code (one entry per unique vehicle code)
    iHiddenDnfCode As Integer           'demo book code (one entry per unique vehicle code & book)
End Type

'10/27/14: 1 or 2 place rating
Dim sm1or2PlaceRating As String
Dim bmShowAudIfPodcast As Boolean           '4-18-18 option to show Aud % for Podcast vehicles on BR

Dim tmPodcast_Info() As PODCAST_INFO

Const LBONE = 1

'
'
'       Create the record to show a CBS line
'       <input> ilListIndex - report type
'               tlSofList() array of slsp offices
'       2-24-06
Public Sub mShowCBS(ilListIndex As Integer, tlSofList() As SOFLIST, slDailyExists As String)
    Dim ilTemp As Integer
    Dim llLineRef As Long

    tmCbf.lRate = 0
    tmCbf.sPriceType = ""       '2-23-01
    tmCbf.sDysTms = "Cancel Before Start"
    'force extra options to unused so they wont show on report
    tmCbf.iOBBLen = 0
    tmCbf.iCBBLen = 0
    tmCbf.s1stPosition = "N"
    tmCbf.sSoloAvail = "N"
    tmCbf.sPrefDT = " "
    tmCbf.lLineNo = CLng((tmClf.iLine) * CLng(1000)) + tmClf.iCntRevNo
    tmCbf.iVefCode = tmClf.iVefCode
    tmCbf.lRafCode = tmClf.lRafCode '8-30-06
    llLineRef = tmClf.iLine
    If tmClf.sType = "H" Then
        llLineRef = tmClf.iPkLineNo
    End If
    mSetResortField tmClf.sType, llLineRef               '12-22-20

    'init all the fields that are shown on report
    tmCbf.lCPP = 0
    tmCbf.lCPM = 0
    tmCbf.lGrImp = 0
    tmCbf.iAvgRate = 0
    tmCbf.lAvgAud = 0
    tmCbf.lGRP = 0
    If ilListIndex = CNT_INSERTION Then 'Insertion orders needs to separate daily/weekly by vehicle not contract
        For ilTemp = LBound(tlSofList) To UBound(tlSofList) - 1
            If tmClf.iVefCode = tlSofList(ilTemp).iMnfSSCode Then
                If tlSofList(ilTemp).iSofCode = True Then   'at least one daily exists for this vehicle
                    tmCbf.sDailyExists = "Y"            '12-9-03 make sure the Insertion Order shows the CBS on dailies
                    tmCbf.sDailyWkly = "2"
                Else
                    tmCbf.sDailyWkly = "0"          'all weekly
                    tmCbf.sDailyExists = "N"
                End If
                Exit For
            End If
        Next ilTemp
    Else
        'tmCbf.sDailyWkly: 0 = no dailies, 1 = daily line, 2 = wkly line but there are dailies on this order
        If slDailyExists = "Y" Then       '5-21-03
            tmCbf.sDailyWkly = "2"
        Else
            tmCbf.sDailyWkly = "0"
        End If
    End If
    Exit Sub
End Sub

'           mGetPopPkgByLine - set up that population because there
'           could be more than 1 package line for the same vehicle, with varying books
'           referenced.  Build array of package populations by the package lines (ilPkgLineList)
'           <input> ilPkgLineList - array of all package lines
'                   llPopByLine - array of population by line
Private Sub mGetPopPkgByLine(ilPkgLineList() As Integer, llPopByLine() As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                                                                                *
'******************************************************************************************
    Dim ilLoopOnPkgLine As Integer
    Dim ilPkg As Integer
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim slStr As String
    '10-23-17 adjust ilPkglinelist for 0 based
    For ilLoopOnPkgLine = LBound(ilPkgLineList) To UBound(ilPkgLineList) - 1 Step 1
        'If tmClf.iVefCode = ilPkgLineList(ilLoopOnPkgLine) Then
            'ilLoopOnPkgLine contains index to store package population
            lmPopPkgByLine(ilLoopOnPkgLine) = -1    '11-24-04 pop for the pkg because of using same hidden vehicle reference as the pkg vehicle
            For ilPkg = 0 To UBound(tgClf) - 1 Step 1
                gUnpackDate tgClf(ilPkg).ClfRec.iStartDate(0), tgClf(ilPkg).ClfRec.iStartDate(1), slStr
                llFltStart = gDateValue(slStr)
                gUnpackDate tgClf(ilPkg).ClfRec.iEndDate(0), tgClf(ilPkg).ClfRec.iEndDate(1), slStr
                llFltEnd = gDateValue(slStr)
                If ilPkgLineList(ilLoopOnPkgLine) = tgClf(ilPkg).ClfRec.iPkLineNo And llFltEnd >= llFltStart Then    'find the assoc pkg vehicle name
                    'For ilLoop = 1 To UBound(ilVehList)
                    '    If tgClf(ilPkg).ClfRec.iVefCode = ilVehList(ilLoop) Then
                    '        Exit For
                    '    End If
                    'Next ilLoop
                    If lmPopPkgByLine(ilLoopOnPkgLine) = -1 And llPopByLine(ilPkg + 1) <> 0 Then '11-24-04 pop for the pkg because of using same hidden vehicle reference as the pkg vehicle
                        lmPopPkgByLine(ilLoopOnPkgLine) = llPopByLine(ilPkg + 1)
                    Else
                        If (lmPopPkgByLine(ilLoopOnPkgLine) <> 0) And (lmPopPkgByLine(ilLoopOnPkgLine) <> llPopByLine(ilPkg + 1)) And (llPopByLine(ilPkg + 1) <> 0) Then  'test to see if this pop is different that the prev one.
                            lmPopPkgByLine(ilLoopOnPkgLine) = 0                                           'if different pops, calculate the contract  summary different
                        Else
                            'if current line has population, but there was already a different across
                            'lines in pop, dont save new one
                            If llPopByLine(ilPkg + 1) <> 0 And (lmPopPkgByLine(ilLoopOnPkgLine) <> 0 And lmPopPkgByLine(ilLoopOnPkgLine) <> -1) Then '2/1/99
                                lmPopPkgByLine(ilLoopOnPkgLine) = llPopByLine(ilPkg + 1)
                            End If
                        End If
                    End If
                End If
            Next ilPkg
        'End If
        Next ilLoopOnPkgLine
    Exit Sub
End Sub

'           mCreateBRSportsComments - go thru each schedule line and determine
'           if its a sports vehicle, and user wants to show the vehicle, game #
'           teams, date & time of event in comments section of the printed contrct.
'           These 'type 5' records will be used in a subreport called from the
'           summary versions of the printed contract.
'           10-20-05
'
Public Sub mCreateBRSportsComments()
    Dim ilClf As Integer
    Dim ilRet As Integer
    For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
        tmClf = tgClf(ilClf).ClfRec
            If tmClf.sGameLayout = "Y" Then     'show the games defined for this line
                mSetupBrHdr 0
                '8-3-12 do need sort order for the DP, sort the game info by line
                tmCbf.iRdfDPSort = tmClf.iRdfCode   'daypart
                '3-6-07 determine how to sort the order:  using R/C Items sort code, DP sort code, or Sch Line #
                'tmCbf.iRdfDPSort = gFindDPSort(tmMRif(), tmMRDF(), tmClf.iRdfcode, tmClf.iVefCode)
                'If tmCbf.iRdfDPSort < 0 Then
                '    tmCbf.iRdfDPSort = tmClf.iLine
                'End If
                tmCbf.iExtra2Byte = 5           'Sports comments flag
                tmCbf.lChfCode = tmClf.lChfCode 'contrct code
                tmCbf.iVefCode = tmClf.iVefCode
                'get the game defined for this line
                tmCgfSrchKey1.lClfCode = tmClf.lCode
                ilRet = btrGetEqual(hmCgf, tmCgf, imCgfRecLen, tmCgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                Do While (tmClf.lCode = tmCgf.lClfCode And ilRet = BTRV_ERR_NONE)
                    'tmCbf.lLineNo = tmClf.iLine
                    '12-26-12 change saving line code instead of line # for matching in subreport for game info
                    tmCbf.lLineNo = tmCgf.lCode

                    tmCbf.iTotalWks = tmCgf.iGameNo
                    'loop and get all games for this
                    'gen date and time set previously and should still be in buffer
                    ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                    ilRet = btrGetNext(hmCgf, tmCgf, imCgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                Loop
            End If
    Next ilClf

End Sub

'           m104PlusWks - test if contract or previous version
'           due to differences are over 104 weeks.  if so,
'           cannot process the contract.
'           <input> lldate - start date of contract
'                   lldate2 - end date of contract
'           <return> true if OK to process, else false
'
'       1-18-06 Print NTR summary even tho contract exceeds 104 weeks
Function m104PlusWks(llDate As Long, llDate2 As Long, tlBR As BRSELECTIONS, tlRvf() As RVF) As Integer
    Dim ilRet As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim ilTotalsWks As Integer
    ReDim tlSbf(0 To 0) As SBF
    ReDim ilVehiclesDone(0 To 0) As Integer
    'ReDim llStdStartDates(1 To 13) As Long
    ReDim llStdStartDates(0 To 13) As Long  'Index zero ignored
    Dim blPodExists As Boolean

    m104PlusWks = True
    slStartDate = Format$(llDate, "m/d/yy")
    slStartDate = gObtainEndStd(slStartDate)       'get std bdcst end date
    gObtainYearMonthDayStr slStartDate, True, slYear, slMonth, slDay
    Do While (slMonth <> "01") And (slMonth <> "04") And (slMonth <> "07") And (slMonth <> "10")
        slMonth = str$((Val(slMonth) - 1))
        slDay = "15"
        slStartDate = slMonth & "/" & slDay & "/" & slYear
        gObtainYearMonthDayStr slStartDate, True, slYear, slMonth, slDay
        slStartDate = gObtainEndStd(slStartDate)        'get std bdcst end date
    Loop
    'slStartDate now contains the end date of a quarter
    slStartDate = gObtainStartStd(slStartDate)      'start date of the quarter for this cnt
    slEndDate = Format$(llDate2, "m/d/yy")

    'recalculate the totals wks because the start date was pushed back to the start of the qtr
    ilTotalsWks = (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 7 + 1

    'If ((llDate2 - llDate + 1) / 7 > 104) Then
    If ilTotalsWks > 104 Then
        tmCbf = tmZeroCbf
        'tmCbf.iGenTime(0) = igNowTime(0)
        'tmCbf.iGenTime(1) = igNowTime(1)
        'gUnpackTimeLong tlBR.iGenTime(0), tlBR.iGenTime(1), False, lgNowTime
        tmCbf.lGenTime = lgNowTime
        tmCbf.iGenDate(0) = tlBR.iGenDate(0)    'igNowDate(0)
        tmCbf.iGenDate(1) = tlBR.iGenDate(1)   'igNowDate(1)
        tmCbf.lChfCode = tgChf.lCode                'contract internal code
        tmCbf.sSurvey = "Reduce contract to 104 weeks"
        mSetupBrHdr tlBR.iWhichSort                             'format remaining header fields to output
        tmCbf.sDailyExists = "N"
        'tmCbf.sDailyWkly: 0 = no dailies, 1 = daily line, 2 = wkly line but there are dailies on this order
        tmCbf.sDailyWkly = "0"
        ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
        igBR_SchLinesExist = True            'force contrct to print
        m104PlusWks = False

        tmCbf.sSurvey = ""              'dont show the exceed week for NTR version
        mProcessBrNTR tlBR.iWhichSort, tlBR.iShowNTRBillSummary, tlRvf(), tlSbf(), ilVehiclesDone(), llStdStartDates()
        blPodExists = gBuildCPMIDs(hmPcf, tgChf, llStdStartDates(), tmCPM_IDs(), tmCPMSummary())
        If blPodExists Then
            mProcessBR_CPM tlBR.iWhichSort, tlBR.iShowProof, tmCPM_IDs(), tmCPMSummary()      'create the detail IDs, summary by vehicle Research, and vehicle billing vehicle
        End If
    End If
End Function

'
'           mGetHiddenWkVGrs - Calculate the grps for each week individually
'           for hidden line totals by vehicle.  For packages only.
'           <input> llResearchPop - population
'                   ilVehlist() array of vehicle codes
'                   ilPkgLineList() array of package line #s
'           <output> 104 weekly grps for all lines
'           1-23-03
'
'
Sub mGetHiddenWkVGrps(llInResearchPop As Long, ilVehList() As Integer, ilPkgLineList() As Integer)
    'Dim ilQtr As Integer
    Dim ilWeekly As Integer
    Dim ilWeeklyMinusOne As Integer
    'Dim ilRchQtr As Integer
    Dim llLnSpots As Long
    Dim llCPP As Long
    Dim llCPM As Long
    Dim llAvgAud As Long
    'Dim llTotalCost As Long
    Dim dlTotalCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim llGrimps As Long
    Dim ilVehicle As Integer
    Dim ilUpperGrp As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim ilPkgVehLoop As Integer
    Dim llResearchPop As Long           '6-29-04
    'Dim ilVaryingQtrs As Integer
    Dim llQtr As Long                   '4-10-19    replace ilQtr due to subscript out of range
    Dim llRchQtr As Long                '4-10-19    replace ilRchQtr
    Dim llVaryingQtrs As Long           '4-10-19    replace ilVaryingQtrs

    '2-11-00 Build weekly grps for all unique vehicles within a package.  Each line entry is 8 quarters (max 2 years).  Entries are unique
    'for  vehicle/daypart/#spots/line rate
    '1-24-03 Loop on all Package lines; for each package line gather each unique vehicles weekly grps for 8 qtrs (13 wks at a time)
    'ReDim tmWkHiddenVGrps(1 To 1) As WEEKLYGRPS
    'ilUpperGrp = 1
    ReDim tmWkHiddenVGrps(0 To 0) As WEEKLYGRPS
    ilUpperGrp = 0
    For ilPkgVehLoop = LBound(ilPkgLineList) To UBound(ilPkgLineList) - 1 Step 1
        For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1
            llVaryingQtrs = 0
            For llQtr = 1 To imMaxQtrs Step 1              'cycle thru 8 quarters
                'For ilWeekly = 1 To 13              'each qtr has 13 weeks stored
                 For ilWeekly = 1 To imWeeksPerQtr(llQtr)       '12/13/14 week qtrs, reset for each contract
                    ilWeeklyMinusOne = ilWeekly - 1
                    'ReDim tmVRtg(1 To 1) As Integer
                    'ReDim tmVCost(1 To 1) As Long
                    'ReDim tmVGRP(1 To 1) As Long
                    'ReDim tmVGrimp(1 To 1) As Long
                    ReDim tmVRtg(0 To 0) As Integer
                    ReDim tmVCost(0 To 0) As Long
                    ReDim tmVGRP(0 To 0) As Long
                    ReDim tmVGrimp(0 To 0) As Long

                    llLnSpots = 0
                    If tgSpf.sDemoEstAllowed = "Y" Then
                        llResearchPop = -1
                    Else
                        llResearchPop = llInResearchPop
                    End If
                    For llRchQtr = llQtr To UBound(tmLRch) - 1 Step 8   'process as many lines that are built
                        If (tmLRch(llRchQtr).sType = "H") And (tmLRch(llRchQtr).iVefCode = ilVehList(ilVehicle)) And (tmLRch(llRchQtr).iPkLineNo = ilPkgLineList(ilPkgVehLoop)) Then
                            llLnSpots = llLnSpots + tmLRch(llRchQtr).lSpots(ilWeeklyMinusOne)
                            If (tmLRch(llRchQtr).lSpots(ilWeeklyMinusOne) <> 0) Then              'only process if spots exist in the qtr
                                'tmVRtg(UBound(tmVRtg)) = tmLRch(ilRchQtr).lAvgAud(ilWeekly)
                                tmVRtg(UBound(tmVRtg)) = tmLRch(llRchQtr).iWklyRating(ilWeeklyMinusOne) '2-17-00
                                tmVCost(UBound(tmVCost)) = tmLRch(llRchQtr).lRates(ilWeeklyMinusOne)
                                tmVGRP(UBound(tmVGRP)) = tmLRch(llRchQtr).lWklyGRP(ilWeeklyMinusOne)
                                tmVGrimp(UBound(tmVGrimp)) = tmLRch(llRchQtr).lWklyGrimp(ilWeeklyMinusOne)
                                If tgSpf.sDemoEstAllowed = "Y" Then
                                    If llResearchPop = -1 Then
                                        llResearchPop = tmLRch(llRchQtr).lPopEst(ilWeeklyMinusOne)
                                    Else
                                        If llResearchPop <> tmLRch(llRchQtr).lPopEst(ilWeeklyMinusOne) And llResearchPop <> 0 Then
                                            llResearchPop = 0
                                        End If
                                    End If
                                End If
                                'ReDim Preserve tmVRtg(1 To UBound(tmVRtg) + 1) As Integer
                                'ReDim Preserve tmVCost(1 To UBound(tmVCost) + 1) As Long
                                'ReDim Preserve tmVGRP(1 To UBound(tmVGRP) + 1) As Long
                                'ReDim Preserve tmVGrimp(1 To UBound(tmVGrimp) + 1) As Long
                                ReDim Preserve tmVRtg(0 To UBound(tmVRtg) + 1) As Integer
                                ReDim Preserve tmVCost(0 To UBound(tmVCost) + 1) As Long
                                ReDim Preserve tmVGRP(0 To UBound(tmVGRP) + 1) As Long
                                ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) + 1) As Long
                            End If
                        End If
                    Next llRchQtr
                    'If UBound(tmVRtg) > 1 Then
                    If UBound(tmVRtg) > 0 Then
                        'dimensions must be exact sizes
                        'ReDim Preserve tmVRtg(1 To UBound(tmVRtg) - 1) As Integer
                        'ReDim Preserve tmVCost(1 To UBound(tmVCost) - 1) As Long
                        'ReDim Preserve tmVGRP(1 To UBound(tmVGRP) - 1) As Long
                        'ReDim Preserve tmVGrimp(1 To UBound(tmVGrimp) - 1) As Long
                        ReDim Preserve tmVRtg(0 To UBound(tmVRtg) - 1) As Integer
                        ReDim Preserve tmVCost(0 To UBound(tmVCost) - 1) As Long
                        ReDim Preserve tmVGRP(0 To UBound(tmVGRP) - 1) As Long
                        ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) - 1) As Long
                        'ilRchQtr = (ilQtr - 1) * 13 + ilWeekly
                        '11-14-11 detrmine the next starting index based on the # weeks per qtr
                        llRchQtr = llVaryingQtrs + ilWeekly
                        'only care about the weeks grps
                        ilFound = False
                        For ilLoop = LBound(tmWkHiddenVGrps) To UBound(tmWkHiddenVGrps) - 1
                            If tmWkHiddenVGrps(ilLoop).iVefCode = ilVehList(ilVehicle) And tmWkHiddenVGrps(ilLoop).iPkLineNo = ilPkgLineList(ilPkgVehLoop) Then
                                ilFound = True
                                Exit For
                            End If
                        Next ilLoop
                        If ilFound Then
                            'gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, llTotalCost, tmCbf.iAvgRate, llGrimps, tmWkHiddenVGrps(ilLoop).lGrps(llRchQtr), llCPP, llCPM, llAvgAud
                            gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, dlTotalCost, tmCbf.iAvgRate, llGrimps, tmWkHiddenVGrps(ilLoop).lGrps(llRchQtr), llCPP, llCPM, llAvgAud 'TTP 10439 - Rerate 21,000,000
                        Else
                            tmWkHiddenVGrps(ilUpperGrp).iVefCode = ilVehList(ilVehicle)
                            tmWkHiddenVGrps(ilUpperGrp).iPkLineNo = ilPkgLineList(ilPkgVehLoop)

                            'gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, llTotalCost, tmCbf.iAvgRate, llGrimps, tmWkHiddenVGrps(ilUpperGrp).lGrps(llRchQtr), llCPP, llCPM, llAvgAud
                            gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, dlTotalCost, tmCbf.iAvgRate, llGrimps, tmWkHiddenVGrps(ilUpperGrp).lGrps(llRchQtr), llCPP, llCPM, llAvgAud 'TTP 10439 - Rerate 21,000,000
                            ilUpperGrp = UBound(tmWkHiddenVGrps) + 1
                            'ReDim Preserve tmWkHiddenVGrps(1 To ilUpperGrp) As WEEKLYGRPS
                            ReDim Preserve tmWkHiddenVGrps(0 To ilUpperGrp) As WEEKLYGRPS
                        End If
                    End If
                Next ilWeekly
                llVaryingQtrs = llVaryingQtrs + imWeeksPerQtr(llQtr)     '11-14-11
            Next llQtr
        Next ilVehicle
    Next ilPkgVehLoop
    Erase tmVRtg
    Erase tmVCost
    Erase tmVGRP
    Erase tmVGrimp
End Sub

'
'           mProcessBRNTR - process NTR transactions for this order to show on Printed Contract Summary
'           this printed summary comes from the Snapshot button on the Orders/Proposals screen.
'           The records are stored in global array, tgIBSbf.  As each entry is read, it is created in CBF
'           <input> ilWhichSort - parameters to setup sort field for Crystal
'                   ilShowNTRBillSummary - true to merge NTR billing summary with air time billing summary
'                    tlRVF() - array of RVF records found for tax computations
'                    tlInstallSBF() - array of Instllment $ from sBF
'                    ilVehiclesDone() - array of vehicles processed already (monthly installments gathered for vehicle
'                    llSTdStartDates() - array of 12 months to determine what month to place $ into
Public Sub mProcessBrNTR(ilWhichSort As Integer, ilShowNTRBillSummary As Integer, tlRvf() As RVF, tlInstallSBF() As SBF, ilVehiclesDone() As Integer, llStdStartDates() As Long)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llProjectedFlights                                                                    *
'******************************************************************************************
    'TTP 10855 - fix potential NTR Overflow errors
    'Dim ilSbf As Integer
    Dim llSbf As Long
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim llTax1Pct As Long
    Dim llTax2Pct As Long
    Dim slGrossNet As String
    Dim llProjectedTax1 As Long
    Dim llProjectedTax2 As Long
    Dim ilAgyComm As Integer
    Dim llNTRAmt As Long
    Dim slLastBillDate As String
    Dim llLastBillDate As Long
    Dim llSbfBillDate As Long
    Dim llTax1CalcAmt As Long
    Dim llTax2CalcAmt As Long
    'Dim llInstallBilling(1 To 13) As Long
    Dim llInstallBilling(0 To 13) As Long       'Index zero ignored
    Dim ilMonthLoop As Integer
    Dim ilFound As Integer
    Dim ilSelectedLoop As Integer
    Dim ilNTRBillLoop As Integer
    Dim ilFoundNTRBill As Integer
    Dim llNTRNet As Long
    Dim llNTRComm As Long
    Dim ilVefInxForCallLetters As Integer
    ReDim tlntrbillsummary(0 To 0) As NTRBILLSUMMARY

    mSetupBrHdr ilWhichSort                                     'build the advt, agy, slsp and sort fields specs that ned to be built

    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slLastBillDate
    If slLastBillDate = "" Then
        slLastBillDate = "1/1/1970"
    End If
    llLastBillDate = gDateValue(slLastBillDate)

    'accumulate taxes for IN airtime only thru the end of the last bill date, NTR will be processed later
    llProjectedTax1 = 0
    llProjectedTax2 = 0
    llTax1CalcAmt = 0
    llTax2CalcAmt = 0
'        For ilLoop = LBound(tlRvf) To UBound(tlRvf) - 1
'            gUnpackDateLong tlRvf(ilLoop).iTranDate(0), tlRvf(ilLoop).iTranDate(1), llSbfBillDate
'            If tlRvf(ilLoop).iMnfItem > 0 And llSbfBillDate <= llLastBillDate Then          'look for NTR items from PHF/RVF onlly
'                llTax1CalcAmt = llTax1CalcAmt + tlRvf(ilLoop).lTax1
'                llTax2CalcAmt = llTax2CalcAmt + tlRvf(ilLoop).lTax2
'            End If
'        Next ilLoop
    If tgChf.iAgfCode > 0 Then        'get agy comm in case taxes need to be obtained
        ilAgyComm = tmAgf.iComm
    Else
        ilAgyComm = 0                       'direct
    End If

    'create a CBF record for each SBF found
    For llSbf = LBound(tgIBSbf) To UBound(tgIBSbf) - 1
        If tgIBSbf(llSbf).iStatus >= 0 Then
            tmCbf.iVefCode = tgIBSbf(llSbf).SbfRec.iBillVefCode       'vehicle code for sorting
            tmCbf.iDtFrstBkt(0) = tgIBSbf(llSbf).SbfRec.iDate(0)     'Transaction date
            tmCbf.iDtFrstBkt(1) = tgIBSbf(llSbf).SbfRec.iDate(1)
            tmCbf.sDysTms = tgIBSbf(llSbf).SbfRec.sDescr             'NTR description
            tmCbf.iPctDist = tgIBSbf(llSbf).SbfRec.iMnfItem          'NTR Item Type
            tmCbf.lRate = tgIBSbf(llSbf).SbfRec.lGross               'Rate
            tmCbf.iProdPct = tgIBSbf(llSbf).SbfRec.iNoItems          '# items per unit
            tmCbf.sPriceType = tgIBSbf(llSbf).SbfRec.sAgyComm
            tmCbf.iExtra2Byte = 4                           'NTR detail flag
            
            gUnpackDateLong tgIBSbf(llSbf).SbfRec.iDate(0), tgIBSbf(llSbf).SbfRec.iDate(1), llSbfBillDate
            llNTRAmt = tgIBSbf(llSbf).SbfRec.lGross * tgIBSbf(llSbf).SbfRec.iNoItems
            
            llTax1CalcAmt = 0           '12-15-11
            llTax2CalcAmt = 0

            If llSbfBillDate > llLastBillDate Then      '
                'gGetAirTimeTaxRates tgChf.iAdfCode, tgChf.iAgfCode, tmCbf.iVefCode, llTax1Pct, llTax2Pct
                gGetNTRTaxRates tgIBSbf(llSbf).SbfRec.iTrfCode, llTax1Pct, llTax2Pct, slGrossNet
                If tgIBSbf(llSbf).SbfRec.iTrfCode = 0 Then       'is this ntr item taxable?
                    llTax1Pct = 0
                    llTax2Pct = 0
                End If
'                    llNTRAmt = tgIBSbf(llSbf).SbfRec.lGross * tgIBSbf(llSbf).SbfRec.iNoItems

                gFutureTaxesForNTR llNTRAmt, llProjectedTax1, llProjectedTax2, ilAgyComm, tgChf.iPctTrade, llTax1Pct, llTax2Pct, slGrossNet
            Else
                For ilLoop = LBound(tlRvf) To UBound(tlRvf) - 1
                    If tlRvf(ilLoop).iMnfItem > 0 And tlRvf(ilLoop).iBillVefCode = tgIBSbf(llSbf).SbfRec.iBillVefCode And tlRvf(ilLoop).lSbfCode = tgIBSbf(llSbf).SbfRec.lCode Then
'                            llTax1CalcAmt = llTax1CalcAmt + tlRvf(ilLoop).lTax1
'                            llTax2CalcAmt = llTax2CalcAmt + tlRvf(ilLoop).lTax2
                        llProjectedTax1 = tlRvf(ilLoop).lTax1
                        llProjectedTax2 = tlRvf(ilLoop).lTax2
                    End If
                Next ilLoop

'                    llNTRAmt = tgIBSbf(llSbf).SbfRec.lGross * tgIBSbf(llSbf).SbfRec.iNoItems
            End If


            '12/17/06-Change to tax by agency or vehicle
            'tmCbf.sLineType = tgIBSbf(llSbf).SbfRec.sSlsTax          'tax applicable from site (y/n)
            llTax1CalcAmt = llTax1CalcAmt + llProjectedTax1
            llTax2CalcAmt = llTax2CalcAmt + llProjectedTax2
            tmCbf.lTax1 = llTax1CalcAmt
            tmCbf.lTax2 = llTax2CalcAmt
            ilVefInxForCallLetters = gBinarySearchVef(tmCbf.iVefCode)
            If ilVefInxForCallLetters <> -1 Then        '10-23-06 if new package vehicle entered for proposal, the index wont exist.  The vehicle
                gFindStationMkt ilVefInxForCallLetters, tmCbf            '2-21-18 show market name?
            End If
            ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            
            '12-14-11 receivables tax $ in these 2 fields; put in detail NTR records only once.  crystal will accumulate all the detail records to get total taxes
            llTax1CalcAmt = 0
            llTax2CalcAmt = 0
            
            'Create array of NTR vehicles for the billing summary to be merged with a subreport with air time billing summary
            ilFoundNTRBill = False
            For ilNTRBillLoop = LBound(tlntrbillsummary) To UBound(tlntrbillsummary) - 1
                If tlntrbillsummary(ilNTRBillLoop).iVefCode = tmCbf.iVefCode Then
                    ilFoundNTRBill = True
                    Exit For
                End If
            Next ilNTRBillLoop
            If Not ilFoundNTRBill Then
                ilNTRBillLoop = UBound(tlntrbillsummary)
                ReDim Preserve tlntrbillsummary(0 To UBound(tlntrbillsummary) + 1) As NTRBILLSUMMARY
            End If
            tlntrbillsummary(ilNTRBillLoop).iVefCode = tmCbf.iVefCode
            tlntrbillsummary(ilNTRBillLoop).lTax1 = tlntrbillsummary(ilNTRBillLoop).lTax1 + llProjectedTax1
            tlntrbillsummary(ilNTRBillLoop).lTax2 = tlntrbillsummary(ilNTRBillLoop).lTax2 + llProjectedTax2
            'calc agy comm
            If tmCbf.sPriceType = "Y" Then          'agy commissionable
                llNTRComm = (llNTRAmt * CDbl(ilAgyComm)) / 10000        'round for proper decimals due to multiplication of agy comm
            Else
                llNTRComm = 0       'no commission
            End If
            tlntrbillsummary(ilNTRBillLoop).lAgyComm = tlntrbillsummary(ilNTRBillLoop).lAgyComm + llNTRComm
            tlntrbillsummary(ilNTRBillLoop).lNet = tlntrbillsummary(ilNTRBillLoop).lNet + (llNTRAmt - llNTRComm)
            If llSbfBillDate >= llStdStartDates(13) Then    'over 12 months
                tlntrbillsummary(ilNTRBillLoop).lMonth(13) = tlntrbillsummary(ilNTRBillLoop).lMonth(13) + llNTRAmt
            Else                    'determine what month the billing falls into
                For ilMonthLoop = 1 To 12
                    If llSbfBillDate >= llStdStartDates(ilMonthLoop) And llSbfBillDate < llStdStartDates(ilMonthLoop + 1) Then
                        tlntrbillsummary(ilNTRBillLoop).lMonth(ilMonthLoop) = tlntrbillsummary(ilNTRBillLoop).lMonth(ilMonthLoop) + llNTRAmt
                        Exit For
                    End If
                Next ilMonthLoop
            End If
            
            
            If tgChf.sInstallDefined = "Y" Then   'if installment contract, show install $ on summary page
                gBuildInstallMonths tgIBSbf(llSbf).SbfRec.iBillVefCode, ilVehiclesDone(), llStdStartDates(), llInstallBilling(), tlInstallSBF()
                For ilMonthLoop = 1 To 13
                    tmCbf.lMonth(ilMonthLoop - 1) = llInstallBilling(ilMonthLoop)
                Next ilMonthLoop
                
                'find the vehicle this belongs to
                If ilNTRBillLoop <= UBound(tlntrbillsummary) Then     'make sure not out of subscript range
                    For ilMonthLoop = 1 To 13
                        tlntrbillsummary(ilNTRBillLoop).lInstallment(ilMonthLoop) = tlntrbillsummary(ilNTRBillLoop).lInstallment(ilMonthLoop) + llInstallBilling(ilMonthLoop)
                    Next ilMonthLoop
                End If
'                    tmCbf.iExtra2Byte = 6               'for installment summaries (12month summary)
'                    tmCbf.iVefCode = tgIBSbf(llSbf).SbfRec.iBillVefCode
'                    ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                'init monthly $ and units buckets for next flight
                For ilMonthLoop = 1 To 13 Step 1
                    tmCbf.lMonth(ilMonthLoop - 1) = 0
                    tmCbf.lMonthUnits(ilMonthLoop - 1) = 0
                    tmCbf.lWeek(ilMonthLoop - 1) = 0
                    llInstallBilling(ilMonthLoop) = 0
                Next ilMonthLoop
            End If

        End If
    Next llSbf
    
     If tgChf.sInstallDefined = "Y" Then       'if contract is installment billing, do not show the NTR as separate billing
        For ilNTRBillLoop = LBound(tlntrbillsummary) To UBound(tlntrbillsummary) - 1
            tmCbf.iVefCode = tlntrbillsummary(ilNTRBillLoop).iVefCode
            
            For ilMonthLoop = 1 To 13
                tmCbf.lMonth(ilMonthLoop - 1) = tlntrbillsummary(ilNTRBillLoop).lInstallment(ilMonthLoop)
            Next ilMonthLoop
            
            tmCbf.lTax1 = tlntrbillsummary(ilNTRBillLoop).lTax1
            tmCbf.lTax2 = tlntrbillsummary(ilNTRBillLoop).lTax2
            
            'tmCbf.lValue(1) = tlntrbillsummary(ilNTRBillLoop).lNet
            'tmCbf.lValue(2) = tlntrbillsummary(ilNTRBillLoop).lAgyComm
            tmCbf.lValue(0) = tlntrbillsummary(ilNTRBillLoop).lNet
            tmCbf.lValue(1) = tlntrbillsummary(ilNTRBillLoop).lAgyComm
            tmCbf.iExtra2Byte = 6               'ntr billing summary for subreport
            ilVefInxForCallLetters = gBinarySearchVef(tmCbf.iVefCode)
            If ilVefInxForCallLetters <> -1 Then        '10-23-06 if new package vehicle entered for proposal, the index wont exist.  The vehicle
                gFindStationMkt ilVefInxForCallLetters, tmCbf            '2-21-18 show market name?
            End If
            ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
        Next ilNTRBillLoop
    Else
        'check to see if NTR should be merged with Air time as a separate section
        If ilShowNTRBillSummary Then        'show the NTR billing summary?
            For ilNTRBillLoop = LBound(tlntrbillsummary) To UBound(tlntrbillsummary) - 1
                tmCbf.iVefCode = tlntrbillsummary(ilNTRBillLoop).iVefCode
                
                For ilMonthLoop = 1 To 13
                    If tgChf.sInstallDefined = "Y" Then
                        tmCbf.lMonth(ilMonthLoop - 1) = tlntrbillsummary(ilNTRBillLoop).lInstallment(ilMonthLoop)
                    Else
                        tmCbf.lMonth(ilMonthLoop - 1) = tlntrbillsummary(ilNTRBillLoop).lMonth(ilMonthLoop)
                    End If
                Next ilMonthLoop
                
                tmCbf.lTax1 = tlntrbillsummary(ilNTRBillLoop).lTax1
                tmCbf.lTax2 = tlntrbillsummary(ilNTRBillLoop).lTax2
                
                'tmCbf.lValue(1) = tlntrbillsummary(ilNTRBillLoop).lNet
                'tmCbf.lValue(2) = tlntrbillsummary(ilNTRBillLoop).lAgyComm
                tmCbf.lValue(0) = tlntrbillsummary(ilNTRBillLoop).lNet
                tmCbf.lValue(1) = tlntrbillsummary(ilNTRBillLoop).lAgyComm
                tmCbf.iExtra2Byte = 8               'ntr billing summary for subreport
                ilVefInxForCallLetters = gBinarySearchVef(tmCbf.iVefCode)
                If ilVefInxForCallLetters <> -1 Then        '10-23-06 if new package vehicle entered for proposal, the index wont exist.  The vehicle
                    gFindStationMkt ilVefInxForCallLetters, tmCbf            '2-21-18 show market name?
                End If
                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            Next ilNTRBillLoop
        End If
    End If
    
    Erase tlntrbillsummary
    Exit Sub
End Sub

'
'
'                   mGenDiff - Compare the current contract version against
'                       the previous versions and build the difference
'                       in "TG" array (tgchf, tgclf, tgcff)
'
'                   <input> ilTask - REPORTSJOB = from rptsel (not Traffic)
'                           lmCurrCode = current versions contract code
'                           lmPrevCode = previous versions contract code
'                   <output> llPrevSpots - total spot count of current version
'                            llPrevGross - total gross $ of current version
'                            ilCurrTotWks - total weeks from start date to end date (ignoring
'                                           if spots aired in every week
'                            ilCurrAirWks - total weeks having spots to air
'                   Created:  4/25/98 (copied & modified from Contract.bas (cbcDifference)
'
'           2-23-06 ignore hidden lines when calculating previous gross & spots
'
'
Sub gGenDiff(ilTask As Integer, hmCHF As Integer, hmClf As Integer, hmCff As Integer, lmCurrCode As Long, lmPrevcode As Long, llPrevSpots As Long, llPrevGross As Long, ilCurrTotWks As Integer, ilCurrAirWks As Integer, Optional slSnapShot As String = "")
    Dim ilRet As Integer
    Dim ilFound As Integer
    Dim ilClf As Integer
    Dim ilDiffClf As Integer
    Dim ilCff As Integer
    Dim ilDiffCff As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llLnStartDate As Long
    Dim llLnEndDate As Long
    Dim llMoStartDate As Long
    Dim ilFirstCff As Integer
    Dim ilSpots As Integer
    Dim ilDiffSpots As Integer
    Dim ilDay As Integer
    Dim llDate As Long
    Dim ilCffFound As Integer
    Dim ilDiffCffFound As Integer
    Dim ilPrevCff As Integer
    Dim ilCffIndex As Integer
    Dim ilAddWk As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim ilClfIndex As Integer
    Dim ilLoop As Integer
    Dim llPrice As Long
    Dim ilPass As Integer
    Dim ilSpotsChgd As Integer
    Dim ilPriceChgd As Integer
    'ReDim ilAllWks(1 To MAXWEEKSFOR2YRS) As Integer     'map of 2 years weeks, set on indicating spot airing in that week
    ReDim ilAllWks(0 To MAXWEEKSFOR2YRS) As Integer     'map of 2 years weeks, set on indicating spot airing in that week. Index zero ignored
    '7/30/19: Contract difference from order/proposal screen
    If ilTask = REPORTSJOB Or slSnapShot = "D" Then
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, lmCurrCode, False, tgChf, tmCClf(), tmCCff())
        If Not ilRet Then
           Exit Sub
        End If
    Else


        'Need to retain tgChf since it may be in an altered state from the Proposal/change mode
        'Move all the lines & flights to Current line and flight arrays and avoid above read
        ReDim tmCClf(0 To 0) As CLFLIST
        For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
            tmCClf(ilClf) = tgClf(ilClf)
            ReDim Preserve tmCClf(0 To UBound(tmCClf) + 1) As CLFLIST
        Next ilClf
        ReDim tmCCff(0 To 0) As CFFLIST
        For ilCff = LBound(tgCff) To UBound(tgCff) - 1 Step 1
            tmCCff(ilCff) = tgCff(ilCff)
            ReDim Preserve tmCCff(0 To UBound(tmCCff) + 1) As CFFLIST
        Next ilCff
    End If

    If lmPrevcode = 0 Then                  'no prvious, show everything as Added
        ReDim tmPclf(0 To 0) As CLFLIST
        tmPclf(0).iStatus = -1 'Not Used
        tmPclf(0).lRecPos = 0
        tmPclf(0).iFirstCff = -1
        ReDim tmPCff(0 To 0) As CFFLIST
        tmPCff(0).iStatus = -1 'Not Used
        tmPCff(0).lRecPos = 0
        tmPCff(0).iNextCff = -1
    Else
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, lmPrevcode, False, tmPChf, tmPclf(), tmPCff())
        If Not ilRet Then
            Exit Sub
        End If
    End If
    ReDim tgClf(0 To 0) As CLFLIST
    tgClf(0).iStatus = -1 'Not Used
    tgClf(0).lRecPos = 0
    tgClf(0).iFirstCff = -1
    ReDim tgCff(0 To 0) As CFFLIST
    tgCff(0).iStatus = -1 'Not Used
    tgCff(0).lRecPos = 0
    tgCff(0).iNextCff = -1



    'Compute Total spots and cost for current version
    llPrevSpots = 0
    llPrevGross = 0
    For ilClf = LBound(tmPclf) To UBound(tmPclf) - 1 Step 1
        ilCff = tmPclf(ilClf).iFirstCff
        Do While ilCff <> -1
        If (tmPCff(ilCff).iStatus = 0) Or (tmPCff(ilCff).iStatus = 1) Then
            llStartDate = tmPCff(ilCff).lStartDate
            llEndDate = tmPCff(ilCff).lEndDate
            llMoStartDate = llStartDate
            Do While gWeekDayLong(llMoStartDate) <> 0
                llMoStartDate = llMoStartDate - 1
            Loop
            For llDate = llMoStartDate To llEndDate Step 7
                If tmPCff(ilCff).CffRec.sDyWk = "D" Then
                    ilSpots = 0         '2-11-03 init for totals spots / week
                    For ilDay = 0 To 6 Step 1
                        If (llDate + ilDay >= llStartDate) And (llDate + ilDay <= llEndDate) Then
                            ilSpots = ilSpots + tmPCff(ilCff).CffRec.iDay(ilDay)
                        End If
                    Next ilDay
                Else
                    ilSpots = tmPCff(ilCff).CffRec.iSpotsWk + tmPCff(ilCff).CffRec.iXSpotsWk
                End If
                Select Case tmPCff(ilCff).CffRec.sPriceType
                    Case "T"
                    llPrice = tmPCff(ilCff).CffRec.lActPrice
                    Case Else
                    llPrice = 0
                End Select
                If tmPclf(ilClf).ClfRec.sType <> "H" Then           '2-23-06
                    llPrevGross = llPrevGross + (llPrice * ilSpots)
                    llPrevSpots = llPrevSpots + ilSpots
                End If
            Next llDate
        End If
        ilCff = tmPCff(ilCff).iNextCff
        Loop
    Next ilClf

    'Calculate the total airing weeks and total weeks (contract headr
    'start/end) from the current version for the output since differences
    'contract wont be correct
    ilCurrTotWks = 0
    ilCurrAirWks = 0
    'get start date of current header
    gUnpackDateLong tgChf.iStartDate(0), tgChf.iStartDate(1), llLnStartDate
    gUnpackDateLong tgChf.iEndDate(0), tgChf.iEndDate(1), llLnEndDate
    ilCurrTotWks = (llLnEndDate - llLnStartDate) \ 7 + 1
    For ilClf = LBound(tmCClf) To UBound(tmCClf) - 1 Step 1
        ilCff = tmCClf(ilClf).iFirstCff
        Do While ilCff <> -1
        If (tmCCff(ilCff).iStatus = 0) Or (tmCCff(ilCff).iStatus = 1) Then
            llStartDate = tmCCff(ilCff).lStartDate
            llEndDate = tmCCff(ilCff).lEndDate
            llMoStartDate = llStartDate
            Do While gWeekDayLong(llMoStartDate) <> 0
                llMoStartDate = llMoStartDate - 1
            Loop
            For llDate = llMoStartDate To llEndDate Step 7
                If tmCCff(ilCff).CffRec.sDyWk = "D" Then
                    For ilDay = 0 To 6 Step 1
                    If (llDate + ilDay >= llStartDate) And (llDate + ilDay <= llEndDate) Then
                        ilSpots = ilSpots + tmCCff(ilCff).CffRec.iDay(ilDay)
                    End If
                    Next ilDay
                Else
                    ilSpots = tmCCff(ilCff).CffRec.iSpotsWk + tmCCff(ilCff).CffRec.iXSpotsWk
                End If
                If ilSpots <> 0 Then
                    ilLoop = (llDate - llLnStartDate) \ 7 + 1
                    If ilLoop > 0 Then
                        ilAllWks(ilLoop) = 1
                    End If
                End If
                ilSpots = 0             'init for next week
            Next llDate
        End If
        ilCff = tmCCff(ilCff).iNextCff
        Loop
    Next ilClf
    For ilLoop = 1 To MAXWEEKSFOR2YRS                   'add up all the weeks that have spots across the lines examined
        If ilAllWks(ilLoop) <> 0 Then
            ilCurrAirWks = ilCurrAirWks + 1
        End If
    Next ilLoop

    'Match up lines and determine differences.  Loop thru current versions.  for each line,
    'find matching line in the previous version
    For ilClf = LBound(tmCClf) To UBound(tmCClf) - 1 Step 1
        'Find matching line
        ilFirstCff = -1
        'Loop thru previous version, find a matching line with the current
        For ilDiffClf = LBound(tmPclf) To UBound(tmPclf) - 1 Step 1
            If tmCClf(ilClf).ClfRec.iLine = tmPclf(ilDiffClf).ClfRec.iLine Then
                gUnpackDateLong tmCClf(ilClf).ClfRec.iStartDate(0), tmCClf(ilClf).ClfRec.iStartDate(1), llStartDate
                gUnpackDateLong tmCClf(ilClf).ClfRec.iEndDate(0), tmCClf(ilClf).ClfRec.iEndDate(1), llEndDate
                gUnpackDateLong tmPclf(ilDiffClf).ClfRec.iStartDate(0), tmPclf(ilDiffClf).ClfRec.iStartDate(1), llLnStartDate
                gUnpackDateLong tmPclf(ilDiffClf).ClfRec.iEndDate(0), tmPclf(ilDiffClf).ClfRec.iEndDate(1), llLnEndDate
                If llEndDate >= llStartDate Then                        'current end date >= Prev start date?
                    If llLnEndDate >= llLnStartDate Then
                        If llLnStartDate < llStartDate Then             'prev start date < that current start date?
                            llStartDate = llLnStartDate                 'get earliest start date from previous
                        End If
                        If llLnEndDate > llEndDate Then                 'prev end date > that current start?
                            llEndDate = llLnEndDate                     'get latest end date from previous
                        End If
                    End If
                Else                                                    'current is a CBS, use the previous start & end dates
                    llStartDate = llLnStartDate
                    llEndDate = llLnEndDate
                End If
                ilCffFound = -1
                ilDiffCffFound = -1
                ilSpotsChgd = False
                ilPriceChgd = False
                For ilPass = 1 To 2                                      '1st pass:  look for matching lines with both spot count & price changed
                                                                        '2nd pass:  Ignore if both found, adjust for this case in removal & addition of week
                                                                        '2nd pass:  create the other differences
                    If llEndDate >= llStartDate Then                        'test for cancel before start
                        'Build airing week
                        llMoStartDate = llStartDate
                        Do While gWeekDayLong(llMoStartDate) <> 0
                            llMoStartDate = llMoStartDate - 1
                        Loop
                        For llDate = llMoStartDate To llEndDate Step 7        'loop from the start to the end, builld one week at a time
                            ilCffFound = -1
                            ilCff = tmCClf(ilClf).iFirstCff
                            ilSpots = 0
                            Do While ilCff <> -1
                                If (tmCCff(ilCff).iStatus = 0) Or (tmCCff(ilCff).iStatus = 1) Then
                                    gUnpackDateLong tmCCff(ilCff).CffRec.iStartDate(0), tmCCff(ilCff).CffRec.iStartDate(1), llLnStartDate    'Week Start date
                                    gUnpackDateLong tmCCff(ilCff).CffRec.iEndDate(0), tmCCff(ilCff).CffRec.iEndDate(1), llLnEndDate    'Week Start date
                                    If (llDate >= llLnStartDate) And (llDate <= llLnEndDate) Then
                                        ilCffFound = ilCff
                                        If tmCCff(ilCff).CffRec.sDyWk = "D" Then
                                            For ilDay = 0 To 6 Step 1
                                                ilSpots = ilSpots + tmCCff(ilCff).CffRec.iDay(ilDay)
                                            Next ilDay
                                        Else
                                            ilSpots = tmCCff(ilCff).CffRec.iSpotsWk + tmCCff(ilCff).CffRec.iXSpotsWk
                                        End If
                                        Exit Do
                                    End If
                                End If
                                ilCff = tmCCff(ilCff).iNextCff
                            Loop
                            ilDiffCffFound = -1
                            ilDiffCff = tmPclf(ilDiffClf).iFirstCff
                            ilDiffSpots = 0
                            Do While ilDiffCff <> -1
                                If (tmPCff(ilDiffCff).iStatus = 0) Or (tmPCff(ilDiffCff).iStatus = 1) Then
                                    gUnpackDateLong tmPCff(ilDiffCff).CffRec.iStartDate(0), tmPCff(ilDiffCff).CffRec.iStartDate(1), llLnStartDate    'Week Start date
                                    gUnpackDateLong tmPCff(ilDiffCff).CffRec.iEndDate(0), tmPCff(ilDiffCff).CffRec.iEndDate(1), llLnEndDate    'Week Start date
                                    If (llDate >= llLnStartDate) And (llDate <= llLnEndDate) Then
                                        ilDiffCffFound = ilDiffCff
                                        If tmPCff(ilDiffCff).CffRec.sDyWk = "D" Then
                                            For ilDay = 0 To 6 Step 1
                                                ilDiffSpots = ilDiffSpots + tmPCff(ilDiffCff).CffRec.iDay(ilDay)
                                            Next ilDay
                                        Else
                                            ilDiffSpots = tmPCff(ilDiffCff).CffRec.iSpotsWk + tmPCff(ilDiffCff).CffRec.iXSpotsWk
                                        End If
                                        Exit Do
                                    End If
                                End If
                                ilDiffCff = tmPCff(ilDiffCff).iNextCff
                            Loop
                            ilAddWk = 0
                            If (ilCffFound >= 0) And (ilDiffCffFound >= 0) Then
                                If ilSpots <> ilDiffSpots Then
                                    ilAddWk = 1
                                    ilSpotsChgd = True
                                End If
                                If tmCCff(ilCffFound).CffRec.sPriceType <> tmPCff(ilDiffCffFound).CffRec.sPriceType Then
                                    ilAddWk = 1
                                    ilPriceChgd = True
                                End If
                                If tmCCff(ilCffFound).CffRec.lActPrice <> tmPCff(ilDiffCffFound).CffRec.lActPrice Then
                                    ilAddWk = 1
                                    ilPriceChgd = True
                                End If

                            ElseIf (ilCffFound >= 0) Then
                                ilAddWk = 2                         'current version new airing week
                            ElseIf (ilDiffCffFound >= 0) Then
                                ilAddWk = 3                         'current version removed airing week
                            End If
                            'The first pass is to only set the schedule lines to negative if something needs to be done
                            'in removal or addition of week logic (below)
                            If ((ilSpotsChgd) And (ilPriceChgd) And (ilPass = 1)) Or ((Not ilSpotsChgd) And (ilPriceChgd) And (ilPass = 1)) Then   'both rate & spot count changed, negate the
                                                                                    'line #s to indicate to process these as removal
                                                                                    'and addition of week
                                tmCClf(ilClf).ClfRec.iLine = -tmCClf(ilClf).ClfRec.iLine
                                tmPclf(ilDiffClf).ClfRec.iLine = -tmPclf(ilDiffClf).ClfRec.iLine
                                Exit For
                            End If
                            If ilPass = 2 Then
                                If ilAddWk > 0 Then
                                    ilCffIndex = UBound(tgCff)
                                    tgCff(ilCffIndex).iNextCff = -1
                                    If ilFirstCff = -1 Then
                                        ilClfIndex = UBound(tgClf)
                                        tgClf(ilClfIndex) = tmCClf(ilClf)
                                        tgClf(ilClfIndex).ClfRec.iStartDate(0) = 0
                                        tgClf(ilClfIndex).ClfRec.iStartDate(1) = 0
                                        tgClf(ilClfIndex).iFirstCff = ilCffIndex
                                        ReDim Preserve tgClf(0 To UBound(tgClf) + 1) As CLFLIST
                                        tgClf(UBound(tgClf)).iFirstCff = -1
                                        tgClf(UBound(tgClf)).iStatus = -1
                                        ilFirstCff = ilCffIndex
                                    Else
                                        tgCff(ilPrevCff).iNextCff = ilCffIndex
                                    End If
                                    ilPrevCff = ilCffIndex
                                    ReDim Preserve tgCff(0 To UBound(tgCff) + 1) As CFFLIST
                                    tgCff(UBound(tgCff)).iStatus = -1 'Not Used
                                    tgCff(UBound(tgCff)).lRecPos = 0
                                    tgCff(UBound(tgCff)).iNextCff = -1
                                    tgCff(ilCffIndex).iStatus = 0   'New to not used
                                    tgCff(ilCffIndex).CffRec.lChfCode = tgChf.lCode
                                    tgCff(ilCffIndex).CffRec.iClfLine = tmCClf(ilClf).ClfRec.iLine
                                    tgCff(ilCffIndex).CffRec.iCntRevNo = tmCClf(ilClf).ClfRec.iCntRevNo
                                    tgCff(ilCffIndex).CffRec.iPropVer = tmCClf(ilClf).ClfRec.iPropVer
                                    If llDate = llStartDate Then
                                        slStartDate = Format$(llDate, "m/d/yy")
                                    Else
                                        slStartDate = Format$(llDate, "m/d/yy")
                                        slStartDate = gObtainPrevMonday(slStartDate)
                                    End If
                                    If tgClf(ilClfIndex).ClfRec.iStartDate(0) = 0 And tgClf(ilClfIndex).ClfRec.iStartDate(1) = 0 Then
                                        gPackDate slStartDate, tgClf(ilClfIndex).ClfRec.iStartDate(0), tgClf(ilClfIndex).ClfRec.iStartDate(1)
                                    End If
                                    slEndDate = Format$(llDate, "m/d/yy")
                                    slEndDate = gObtainNextSunday(slEndDate)
                                    If gDateValue(slEndDate) > llEndDate Then
                                        slEndDate = Format$(llEndDate, "m/d/yy")
                                    End If
                                    gPackDate slEndDate, tgClf(ilClfIndex).ClfRec.iEndDate(0), tgClf(ilClfIndex).ClfRec.iEndDate(1)
                                    gPackDate slStartDate, tgCff(ilCffIndex).CffRec.iStartDate(0), tgCff(ilCffIndex).CffRec.iStartDate(1)
                                    gPackDate slEndDate, tgCff(ilCffIndex).CffRec.iEndDate(0), tgCff(ilCffIndex).CffRec.iEndDate(1)
                                    tgCff(ilCffIndex).lStartDate = gDateValue(slStartDate)
                                    tgCff(ilCffIndex).lEndDate = gDateValue(slEndDate)
                                    If ilAddWk = 1 Then
                                        If tmCCff(ilCffFound).CffRec.sDyWk = tmPCff(ilDiffCffFound).CffRec.sDyWk Then
                                            If tmCCff(ilCffFound).CffRec.sDyWk = "D" Then
                                                tgCff(ilCffIndex).CffRec.sDyWk = "D"
                                                For ilDay = 0 To 6 Step 1
                                                    tgCff(ilCffIndex).CffRec.iDay(ilDay) = tmCCff(ilCffFound).CffRec.iDay(ilDay) - tmPCff(ilDiffCffFound).CffRec.iDay(ilDay)
                                                Next ilDay
                                            Else
                                                tgCff(ilCffIndex).CffRec.sDyWk = "W"
                                                tgCff(ilCffIndex).CffRec.iSpotsWk = tmCCff(ilCffFound).CffRec.iSpotsWk - tmPCff(ilDiffCffFound).CffRec.iSpotsWk
                                                For ilDay = 0 To 6 Step 1
                                                    tgCff(ilCffIndex).CffRec.iDay(ilDay) = tmCCff(ilCffFound).CffRec.iDay(ilDay)
                                                Next ilDay
                                            End If
                                        Else
                                            If tmCCff(ilCffFound).CffRec.sDyWk = "D" Then
                                                'Convert to weekly as only showing difference in spot count
                                                tgCff(ilCffIndex).CffRec.sDyWk = "W"
                                                tgCff(ilCffIndex).CffRec.iSpotsWk = ilSpots - ilDiffSpots
                                                For ilDay = 0 To 6 Step 1
                                                    If tmCCff(ilCffFound).CffRec.iDay(ilDay) > 0 Then
                                                        tgCff(ilCffIndex).CffRec.iDay(ilDay) = 1
                                                    Else
                                                        tgCff(ilCffIndex).CffRec.iDay(ilDay) = 0
                                                    End If
                                                Next ilDay
                                            Else
                                                tgCff(ilCffIndex).CffRec.sDyWk = "W"
                                                tgCff(ilCffIndex).CffRec.iSpotsWk = ilSpots - ilDiffSpots
                                                For ilDay = 0 To 6 Step 1
                                                    tgCff(ilCffIndex).CffRec.iDay(ilDay) = tmCCff(ilCffFound).CffRec.iDay(ilDay)
                                                Next ilDay
                                            End If
                                        End If
                                    ElseIf ilAddWk = 2 Then
                                        If tmCCff(ilCffFound).CffRec.sDyWk = "D" Then
                                            tgCff(ilCffIndex).CffRec.sDyWk = "D"
                                            For ilDay = 0 To 6 Step 1
                                                tgCff(ilCffIndex).CffRec.iDay(ilDay) = tmCCff(ilCffFound).CffRec.iDay(ilDay)
                                            Next ilDay
                                        Else
                                            tgCff(ilCffIndex).CffRec.sDyWk = "W"
                                            tgCff(ilCffIndex).CffRec.iSpotsWk = tmCCff(ilCffFound).CffRec.iSpotsWk
                                            For ilDay = 0 To 6 Step 1
                                                tgCff(ilCffIndex).CffRec.iDay(ilDay) = tmCCff(ilCffFound).CffRec.iDay(ilDay)
                                            Next ilDay
                                        End If
                                    ElseIf ilAddWk = 3 Then
                                        If tmPCff(ilDiffCffFound).CffRec.sDyWk = "D" Then
                                            tgCff(ilCffIndex).CffRec.sDyWk = "D"
                                            For ilDay = 0 To 6 Step 1
                                                tgCff(ilCffIndex).CffRec.iDay(ilDay) = -tmPCff(ilDiffCffFound).CffRec.iDay(ilDay)
                                            Next ilDay
                                        Else
                                            tgCff(ilCffIndex).CffRec.sDyWk = "W"
                                            tgCff(ilCffIndex).CffRec.iSpotsWk = -tmPCff(ilDiffCffFound).CffRec.iSpotsWk
                                            For ilDay = 0 To 6 Step 1
                                                tgCff(ilCffIndex).CffRec.iDay(ilDay) = tmPCff(ilDiffCffFound).CffRec.iDay(ilDay)
                                            Next ilDay
                                        End If
                                    End If
                                    tgCff(ilCffIndex).CffRec.sDelete = "N"
                                    If ilAddWk = 1 Then
                                        If tmCCff(ilCffFound).CffRec.sPriceType <> tmPCff(ilDiffCffFound).CffRec.sPriceType Then
                                            If tmCCff(ilCffFound).CffRec.sPriceType = "T" Then
                                                tgCff(ilCffIndex).CffRec.sPriceType = tmCCff(ilCffFound).CffRec.sPriceType
                                                tgCff(ilCffIndex).CffRec.lActPrice = tmCCff(ilCffFound).CffRec.lActPrice
                                            ElseIf tmPCff(ilDiffCffFound).CffRec.sPriceType = "T" Then
                                                tgCff(ilCffIndex).CffRec.sPriceType = tmPCff(ilDiffCffFound).CffRec.sPriceType
                                                tgCff(ilCffIndex).CffRec.lActPrice = -tmPCff(ilDiffCffFound).CffRec.lActPrice
                                            End If
                                        Else
                                            tgCff(ilCffIndex).CffRec.sPriceType = tmCCff(ilCffFound).CffRec.sPriceType
                                            If ilSpotsChgd Then
                                                tgCff(ilCffIndex).CffRec.lActPrice = tmCCff(ilCffFound).CffRec.lActPrice
                                            Else
                                                tgCff(ilCffIndex).CffRec.lActPrice = tmCCff(ilCffFound).CffRec.lActPrice - tmPCff(ilDiffCffFound).CffRec.lActPrice
                                            End If
                                        End If
                                    ElseIf ilAddWk = 2 Then                 'adding a week in current
                                            tgCff(ilCffIndex).CffRec.lActPrice = tmCCff(ilCffFound).CffRec.lActPrice
                                            tgCff(ilCffIndex).CffRec.sPriceType = tmCCff(ilCffFound).CffRec.sPriceType
                                    ElseIf ilAddWk = 3 Then             'removing a week , show decrease in $
                                            tgCff(ilCffIndex).CffRec.lActPrice = tmPCff(ilDiffCffFound).CffRec.lActPrice
                                            tgCff(ilCffIndex).CffRec.sPriceType = tmPCff(ilDiffCffFound).CffRec.sPriceType

                                    End If

                                End If              'iladdwk > 0
                            End If              'if ilpass = 2
                        Next llDate             '
                        If (ilPass = 1 And ilSpotsChgd And ilPriceChgd) Or ((ilPass = 1) And (Not ilSpotsChgd) And (ilPriceChgd)) Then
                            Exit For
                        End If
                        'Exit For                'llDate = llStartDate To llEndDate Step 7
                    End If                      'llenddate >= llstartdate
                Next ilPass                 'ilpass = 1 to 2
                Exit For
            End If                          'tmCClf(ilClf).ClfRec.iLine = tmPClf(ilDiffClf).ClfRec.iLine
        Next ilDiffClf                      'ilDiffClf = LBound(tmPClf) To UBound(tmPClf)
    Next ilClf                              'ilClf = LBound(tmCClf) To UBound(tmCClf

    'Find a current line that  has been added (not in previous)
    For ilClf = LBound(tmCClf) To UBound(tmCClf) - 1 Step 1
        'Find matching line
        ilFirstCff = -1
        ilFound = False
        If tmCClf(ilClf).ClfRec.iLine > 0 Then          'both rate and spot count not changed
            For ilDiffClf = LBound(tmPclf) To UBound(tmPclf) - 1 Step 1
                If tmCClf(ilClf).ClfRec.iLine = tmPclf(ilDiffClf).ClfRec.iLine Then
                    ilFound = True
                    Exit For
                End If
            Next ilDiffClf
        Else                                'both rate & spot count changed, do a addition and removal for differences
            tmCClf(ilClf).ClfRec.iLine = -tmCClf(ilClf).ClfRec.iLine
        End If
        If Not ilFound Then
            'Add line
            ilClfIndex = UBound(tgClf)
            tgClf(ilClfIndex) = tmCClf(ilClf)
            ReDim Preserve tgClf(0 To UBound(tgClf) + 1) As CLFLIST
            'ilLineNo = 0
            'For ilLoop = LBound(tgClf) To UBound(tgClf) - 1 Step 1
            '    If tgClf(ilLoop).ClfRec.iLine > ilLineNo Then
            '        ilLineNo = tgClf(ilLoop).ClfRec.iLine
            '    End If
            'Next ilLoop
            'tgClf(ilClfIndex).ClfRec.iLine = ilLineNo + 1
            tgClf(ilClfIndex).iFirstCff = -1
            tgClf(UBound(tgClf)).iFirstCff = -1
            tgClf(UBound(tgClf)).iStatus = -1
            ilCff = tmCClf(ilClf).iFirstCff
            Do While ilCff <> -1
                If (tmCCff(ilCff).iStatus = 0) Or (tmCCff(ilCff).iStatus = 1) Then
                    tgCff(UBound(tgCff)) = tmCCff(ilCff)
                    If tgClf(ilClfIndex).iFirstCff = -1 Then
                        tgClf(ilClfIndex).iFirstCff = UBound(tgCff)
                    Else
                        tgCff(UBound(tgCff) - 1).iNextCff = UBound(tgCff)
                    End If
                    tgCff(UBound(tgCff)).iNextCff = -1
                    tgCff(UBound(tgCff)).lRecPos = 0
                    tgCff(UBound(tgCff)).iStatus = 0
                    ReDim Preserve tgCff(0 To UBound(tgCff) + 1) As CFFLIST
                    tgCff(UBound(tgCff)).iStatus = -1 'Not Used
                    tgCff(UBound(tgCff)).iNextCff = -1
                    tgCff(UBound(tgCff)).lRecPos = 0
                End If
                ilCff = tmCCff(ilCff).iNextCff
            Loop
        End If
    Next ilClf
    'Find a line that has been removed (deleted) in the current version (found in previous
    'but not in current)
    For ilDiffClf = LBound(tmPclf) To UBound(tmPclf) - 1 Step 1
        'Find matching line
        ilFirstCff = -1
        ilFound = False
        If tmPclf(ilDiffClf).ClfRec.iLine > 0 Then          'both spot count and rate not changed
            For ilClf = LBound(tmCClf) To UBound(tmCClf) - 1 Step 1
                If tmCClf(ilClf).ClfRec.iLine = tmPclf(ilDiffClf).ClfRec.iLine Then
                    ilFound = True
                    Exit For
                End If
            Next ilClf
        Else                            'both spot count & rate changed, do removal
            tmPclf(ilDiffClf).ClfRec.iLine = -tmPclf(ilDiffClf).ClfRec.iLine
        End If
        If Not ilFound Then
            ilClfIndex = UBound(tgClf)
            tgClf(ilClfIndex) = tmPclf(ilDiffClf)
            ReDim Preserve tgClf(0 To UBound(tgClf) + 1) As CLFLIST
            'ilLineNo = 0
            'For ilLoop = LBound(tgClf) To UBound(tgClf) - 1 Step 1
            '    If tgClf(ilLoop).ClfRec.iLine > ilLineNo Then
            '        ilLineNo = tgClf(ilLoop).ClfRec.iLine
            '    End If
            'Next ilLoop
            'tgClf(ilClfIndex).ClfRec.iLine = ilLineNo + 1
            tgClf(ilClfIndex).iFirstCff = -1
            tgClf(UBound(tgClf)).iFirstCff = -1
            tgClf(UBound(tgClf)).iStatus = -1
            ilDiffCff = tmPclf(ilDiffClf).iFirstCff
            Do While ilDiffCff <> -1
                If (tmPCff(ilDiffCff).iStatus = 0) Or (tmPCff(ilDiffCff).iStatus = 1) Then
                    tgCff(UBound(tgCff)) = tmPCff(ilDiffCff)
                    If tgClf(ilClfIndex).iFirstCff = -1 Then
                        tgClf(ilClfIndex).iFirstCff = UBound(tgCff)
                    Else
                        tgCff(UBound(tgCff) - 1).iNextCff = UBound(tgCff)
                    End If
                    tgCff(UBound(tgCff)).iNextCff = -1
                    tgCff(UBound(tgCff)).lRecPos = 0
                    tgCff(UBound(tgCff)).iStatus = 0
                    If tmPCff(ilDiffCff).CffRec.sDyWk = "D" Then        '4-21-15 wrong index, chg fom ildiffclf to ildiffcff
                        tgCff(UBound(tgCff)).CffRec.sDyWk = "D"
                        For ilDay = 0 To 6 Step 1
                            tgCff(UBound(tgCff)).CffRec.iDay(ilDay) = -tmPCff(ilDiffCff).CffRec.iDay(ilDay)
                        Next ilDay
                    Else
                        tgCff(UBound(tgCff)).CffRec.sDyWk = "W"
                        tgCff(UBound(tgCff)).CffRec.iSpotsWk = -tmPCff(ilDiffCff).CffRec.iSpotsWk
                    End If
                    'tgCff(UBound(tgCff)).CffRec.lActPrice = -tmPCff(ilDiffCff).CffRec.lActPrice
                    tgCff(UBound(tgCff)).CffRec.lActPrice = tmPCff(ilDiffCff).CffRec.lActPrice
                    ReDim Preserve tgCff(0 To UBound(tgCff) + 1) As CFFLIST
                    tgCff(UBound(tgCff)).iStatus = -1 'Not Used
                    tgCff(UBound(tgCff)).iNextCff = -1
                    tgCff(UBound(tgCff)).lRecPos = 0
                End If
                ilDiffCff = tmPCff(ilDiffCff).iNextCff
            Loop
        End If
    Next ilDiffClf

    'Determine earliest and latest dates of previous and current to put in diff header
    gUnpackDateLong tgChf.iStartDate(0), tgChf.iStartDate(1), llStartDate
    gUnpackDateLong tgChf.iEndDate(0), tgChf.iEndDate(1), llEndDate
    gUnpackDateLong tmPChf.iStartDate(0), tmPChf.iStartDate(1), llLnStartDate
    gUnpackDateLong tmPChf.iEndDate(0), tmPChf.iEndDate(1), llLnEndDate
    If llLnStartDate < llStartDate Then
        gPackDateLong llLnStartDate, tgChf.iStartDate(0), tgChf.iStartDate(1)
    End If
    If llLnEndDate > llEndDate Then
        gPackDateLong llLnEndDate, tgChf.iEndDate(0), tgChf.iEndDate(1)
    End If
    Erase tmPclf, tmPCff
    Erase tmCClf, tmCCff
End Sub

Function gProcessBR(tlBR As BRSELECTIONS, ilProcessFlag As Integer, ilTask As Integer) As Integer
'           tlBR - answers from user options
'           ilProcessFlag : 0 = open all files, then process cnt
'                           1 = process cnt (looping on multiple cnts
'                           2 = close files and exit
'
'           ilTask - if REPORTSJOB (12), init all arrays
'                    otherwise dont initialize the contract "tgclf" & "tgCff" arrays
'           <return>  0 = file opens OK
'                     -1 = close files, error in open
'       10-29-03 Add ability to show social economic research
'       12-16-03 set flag if any lines exist.  Used for exporting: if ntr only, blank detail
'                and summary lines wont be generated
'       6-15-04 Ignore CBS lines whose original start date is prior to the contract start/end dates
'               When in prior periods, cases problems with researchtotals
'           11-24-04 CPP and sometimes weekly grps invalid when a package name is the same
'                       vehicle as the hidden lines.  Also, when excluding hidden lines with
'                       with the same case above, and a conventional line also exists, the
'                       research information on the summary resulted in all zeroes.
'           2-24-06 If the last line of a quarter is a CBS, the bottom line research #s were
'                   picked up from another quarter, or could possibly be zero.
'
'       CBF Record types :  cbfExtra2Byte = 0:  detail
'                                           2 = vehicle summary
'                                           3 = contract tots
'                                           4 = NTR
'                                           5 = sports (games)
'                                           6 = for Installment contracts only for those vehicles that
'                                               were not on the air time schedule (i.e. NTRs).  These
'                                               vehicles need to show on the Installment Summary version
'                                           -1 = Key record for Insertion Order
'                                           8 = NTR billing summary (non-installment)
'                                           9 = CPM detail IDs
'                                           10 = CPM vehicle summary (for Research page)
'                                           11 = CPM billing summary by vehicle
'*************************************************************************************
'
'       Broadcast Contract - Order or Proposal version
'       Snapshot from Proposals/Order screen to print the
'       current version viewing of a differences BR against
'       the most current version.
'
'
'       D Hosaka: 7/16/98
'            2/4/99  Fix Research totals for Contract.  Previously, each
'               quarters research values (by line) were used for the final
'               contract totals.  Change to gather entire line, regardless
'               of duration of line (not by quarters).
'           11/29/99 Fix research (avg rtg & grps) when multiple lines
'               for same vehicle use different books
'           2/4/00  Add option (from site pref) to show weeks based on start of
'                   std qtr even if less than 13 weeks vs based on start date if
'                   less than 13 weeks
'           2-10-00 Calculate grps for individual weeks by vehicle and contract
'                   all the weekly totals must be stored in the CBF record;
'                   it cannot be calculated in crystal
'           2-18-00 Fix Package vehicle summary when more than 1 line of the same pkg
'           7-24-01 change aud call routines to allow for Dpf (demo plus table)
'           5-28-04  Add Satellite Research estimates
'**************************************************************************************
    Dim ilStdQtrLess13 As Integer             '2-4-00 Show std qtr if less than 13 weeks
    Dim ilShowStdQtr As Integer             '12-19-20 0 = std, 1 = cal, 2 = corp: was true if std quarters, else false for corp
    Dim ilCorpStdYear As Integer            'corp or std year that contracts belongs in
    Dim ilStartMonth As Integer             'starting month of corp or std year that contract belongs in
    Dim ilRet As Integer
                                            'the same contr # when more than 1 header is active (due to
                                            'planner changes
    Dim ilDiffOnly As Integer               'from user input (Differences only - ckcSelC5(0))
    'ReDim llStdStartDates(1 To 37) As Long  'max 3 yr contract:  only 12 months  used for summary monthly totals;
    ReDim llStdStartDates(0 To 37) As Long  'max 3 yr contract:  only 12 months  used for summary monthly totals; Index zero ignored
                                            '3 years dates to calc taxes ; start date of each std month, starting with start date or current
                                            'order, then each months start date for max 13.  Calc for each contr processing
    Dim ilCurrTotalMonths As Integer        'total months of order processing (to be stored in cbf)
    ReDim ilCurrStartQtr(0 To 1) As Integer   'start date of first qtr processing (to be stored in cbf)
    Dim ilFoundCnt As Integer               'valid contract to process
    Dim ilDemo As Integer                   'loop variable for each of the 4 demos to process
    Dim ilFirstDemo As Integer              'loop variable for the demo name to show (if using all demos categories)
    Dim ilLastDemo As Integer               'loop for demo name to show if using all demo categories
    Dim ilLoop As Integer                   'temp loop variable
    Dim ilLoop3 As Integer                  'temp loop variable
    Dim ilHiddenAndConv As Integer          '11-24-04
    Dim ilRchQtr As Integer
    Dim ilTemp As Integer
    Dim ilSavePkVeh As Integer
    Dim ilMonthLoop As Integer
    Dim ilFoundVeh As Integer
    Dim slStartDate As String               'Contract start date
    Dim slEndDate As String                 'contract end date
    Dim ilTotalWks As Integer               'Total weeks of contract start to end
    Dim slMonth As String
    Dim slDay As String
    Dim slYear As String
    Dim ilClf As Integer                    'loop for lines
    Dim ilCff As Integer                    'loop for flights
    Dim ilLineRate As Integer               'count of # of line items due to unique rates
    Dim ilLineSpots As Integer              'count by line & rate of # spots per week
    Dim ilLineSpotsInx As Integer           'index to entry of unique line in table
    Dim slStr As String                     'temp string for conversions
    Dim ilWeekInx As Integer
    Dim ilQtr As Integer
    Dim llDate As Long                      'temp serial date
    Dim llTaxDate As Long
    Dim llDate2 As Long
    Dim llChfStart As Long                  'start date from contr header
    Dim llChfEnd As Long                    'end date from contr header
    Dim ilDay As Integer
    Dim llAccumSpotCount As Long
    Dim llFltStart As Long                  'serial flight start date
    Dim llFltEnd As Long                    'serial flight end date
    Dim ilFoundLine As Integer              'flag that unique line & rate built in memory
    Dim ilManyFlts As Integer               'true if more than 1 flt for a line, otherwise false.
                                            'This is so Crytal knows to show the package research totals when theres only 1 flt, and
                                            'not to put it on the total line only
    Dim ilSpots As Integer                  'total spots for the flight week
    Dim slValidDays As String * 40           '11-19-02 represents days of the week airing string (without blanks & commas)
    Dim slTempDays As String                 'represents days of the week airing string (with blanks & commas)
    Dim ilUpperDnf As Integer               'number of different demo research codes form contr
    Dim slNameCode As String
    Dim slCode As String
    Dim slDemos As String                   'string of 4 demos defined for contr
    Dim slSurvey As String                  'string of research books from lines
    ReDim ilInputDays(0 To 6) As Integer      'valid days of the week for gGetAvgAud
    Dim ilfirstTime As Integer
    Dim llOvStartTime As Long               'override start time for Get Avg Aud
    Dim llOvEndTime As Long                 'overrride end time for gGetAvgAud
    Dim llPop As Long                       'population from gDemoPOP
    Dim llResearchPop As Long               'pop to use for Research Totals (if different across lines, send 0 and pop will be calc based on grimps)
                                            'otherwise send the pop from the book
    Dim llOverallPopEst As Long             '6-1-04 contract pop estimates for all lines
    Dim llAvgAud As Long
    Dim llLnSpots As Long                   'total spots to send to Research routines
    Dim ilUpperClf As Integer
    Dim ilRealTotalWks As Integer           'total number of weeks based on contr start/end dates
    Dim ilCmpny As Integer                  'Company competitor "Us"
    Dim slTimeStamp As String               'time stamp to MNF file
    Dim dlTemp As Double                      'temporary long variable'TTP 10439 - Rerate 21,000,000
    Dim ilGot1ToPrint As Integer             'true if at least 1 line to print
    Dim llCntGrimps As Long                    'contract total grimps calculated before each line written todisk so that
                                               'the % distribution can be obtained
    Dim llPrevCntGrimps As Long                 'contr total grimps calculated o previous order before each line written to disk so that
                                                'the % distr can be obtained
    Dim llCurrMod As Long                      'total contrct $ for current mod
    Dim llCurrModSpots As Long              'total contract spots for current mod
    Dim llPrevSpots As Long                  'total contracts spots for previous version (differences option)
    Dim llPrevGross As Long                     'total contracts gross $ for previous version (differences option)
    Dim ilVehicle As Integer                  'loop field for ilVehList
    Dim ilPkg As Integer                      'loop field for ilPkgVehList & ilPkgLineList
    Dim ilPkgClf As Integer
    Dim ilPkgOrSpot As Integer
    Dim ilCurrTotalWks As Integer             'for differences - total weeks of current order (start date to end date, obtained from gGenDiff)
    Dim ilCurrAirWks As Integer             'for differences - total weeks airing spots from the current order - obtained from gGenDiff
    ReDim ilPkgVehList(0 To 0) As Integer        'list of unique packages (vehicles)  ie - more than one package line of the same name
    ReDim ilPkgLineList(0 To 0) As Integer    'list of line #s that are packages
    ReDim ilVehList(0 To 0) As Integer        'list of unique vehicles
    ReDim llWklySpots(0 To MAXWEEKSFOR2YRS - 1) As Long    'arrays set up for each line to pass wkly research data to subrtn
    ReDim llWklyRates(0 To MAXWEEKSFOR2YRS - 1) As Long       'arrays set up for each line to pass wkly research data to subrtn
    ReDim llWklyAvgAud(0 To MAXWEEKSFOR2YRS - 1) As Long          'arrays set up for each line to pass wkly research data to subrtn
    ReDim ilWklyRtg(0 To MAXWEEKSFOR2YRS - 1) As Integer      'arrays set up for each line to return contracts research data
    ReDim llWklyGrimp(0 To MAXWEEKSFOR2YRS - 1) As Long       'arrays set up for each line to return contracts research data
    ReDim llWklyGRP(0 To MAXWEEKSFOR2YRS - 1) As Long         'arrays set up for each line to return contracts research data
    ReDim llWklyPopEst(0 To MAXWEEKSFOR2YRS - 1) As Long      'array set up for each line
    ReDim llPopByLine(0 To 1) As Long           '11-23-99. Index zero ignored
    ReDim tmWkHiddenVGrps(0 To 0) As WEEKLYGRPS     '2-20-00
    ReDim tmWkPkgVGrps(0 To 0) As WEEKLYGRPS        '2-20-00
    Dim ilDemoAvgAudDays(0 To 6) As Integer     '11-19-02, rquired for valid days sent to audienc rtn
    Dim slDailyExists As String * 1      '5-20-03 "Y" if at least 1 daily line exist for the contract
    Dim slSpotCount As String        '5-21-03
    Dim ilCBS As Integer             '5-24-03
    Dim slSocEco As String
    Dim ilListIndex As Integer
    Dim llPopEst As Long
    Dim ilVefIndex As Integer
    Dim ilVefInxForCallLetters As Integer       '2-23-18
    Dim ilValue As Integer
    ReDim tlRvf(0 To 0) As RVF
    Dim llTax1Pct As Long           'tax 1 pct from tax table
    Dim llTax2Pct As Long           'tax 2 pct from tax table
    Dim slGrossNet As String
    Dim llTax1CalcAmt As Long
    Dim llTax2CalcAmt As Long
    Dim ilAgyComm As Integer
    Dim slLastBillDate As String
    Dim ilLastBilledInx As Integer
    ReDim llProjectedTax1(0 To 37) As Long      'Index zero ignored
    ReDim llProjectedTax2(0 To 37) As Long      'Index zero ignored
    ReDim llProjectedFlights(0 To 37) As Long   'Index zero ignored
    Dim llLastBillDate As Long
    Dim tlInstallSBFType As SBFTypes
    ReDim ilVehiclesDone(0 To 1) As Integer       'array of installment vehicles processed
    ReDim llInstallBilling(0 To 13) As Long     'Index zero ignored
    ReDim tlSbf(0 To 0) As SBF
    Dim blDefaultToQtr As Boolean
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim ilVaryingQtrs As Integer
    Dim tlPkgVehicleForSurveyList() As PKGVEHICLEFORSURVEYLIST  '12-19-13  List of unique pakcage vehicles and book references.  i.e. 1 package vehicle may be defined for multiple lines.,
                                                                'the same hidden vehicle name name be across those 2 pkg lines and each have different book references
                                                                'i.e. Line 1 pkg: Acme pkg, hidden line Billy Crystal using book spring 2015; Line 2 pkg:  Acme, hidden line Billy Crystal using book Fall 2015
    Dim ilVpfIndex As Integer               '3-9-18
    Dim blPodExists As Boolean
    Dim blNonPodExists As Boolean
    Dim llLoopOnQtr As Long             '4-10-19 replace ilQTr subscript out of range issue
    Dim llRchQtr As Long              '4-10-19 replace ilRchQtr subscript out of range issue
    Dim llTempLRch As Long              '4-10-19 replace ilTemp subscript out of range issue
    Dim slTempGrimp As String
    Dim ilClfInx As Integer 'TTO 8410
    gProcessBR = 0
    If ilProcessFlag = 0 Then                       'open files once, if flag = 1 just process the contract
        ilRet = mOpenBRFiles()                      'open all BR files
        If ilRet <> BTRV_ERR_NONE Then              'any error, exit and quit
            Screen.MousePointer = vbDefault
            gProcessBR = -1                         'return and close all files
            Exit Function
        End If
    ElseIf ilProcessFlag = 1 Then
        ilShowStdQtr = tlBR.iCorpOrStd      '12-19-20 0 = std, 1 = cal, 2 = corp , was true if std
        'These fields are for differences only option
        llPrevSpots = tlBR.lPrevSpots       'retrieve the values from Traffic for Diff only calcs
        llPrevGross = tlBR.lPrevGross
        ilCurrTotalWks = tlBR.iCurrTotWks   'retrieve the values from Traffic for Air wks vs Contractual weeks for header info
        ilCurrAirWks = tlBR.iCurrAirWks

        'if contract is greater than 2 years, we need to ignore it for now.  Tables are exceeded
        gUnpackDate tgChf.iStartDate(0), tgChf.iStartDate(1), slStartDate
        If slStartDate = "" Then
            llDate = 0
        Else
            llDate = gDateValue(slStartDate)
        End If
        gUnpackDate tgChf.iEndDate(0), tgChf.iEndDate(1), slEndDate
        If slEndDate = "" Then
            llDate2 = 0
        Else
            llDate2 = gDateValue(slEndDate)
        End If

        If (tgChf.sDelete = "Y") Or ((llDate2 - llDate + 1) / 7 > 104) Then
            ilFoundCnt = False
        End If

        ilRet = m104PlusWks(llDate, llDate2, tlBR, tlRvf())            '7-27-05 see if over 104 weeks
        If ilRet = False Then                           'over 104 weeks
            igBR_SSLinesExist = True            '12-16-03 force output
            ilRet = btrClose(hmSof)
            ilRet = btrClose(hmUrf)
            ilRet = btrClose(hmMnf)
            ilRet = btrClose(hmSlf)
            ilRet = btrClose(hmRdf)
            ilRet = btrClose(hmDnf)
            ilRet = btrClose(hmAgf)
            ilRet = btrClose(hmAdf)
            ilRet = btrClose(hmCff)
            ilRet = btrClose(hmClf)
            ilRet = btrClose(hmCHF)
            ilRet = btrClose(hmTChf)
            ilRet = btrClose(hmCbf)
            ilRet = btrClose(hmCxf)
            ilRet = btrClose(hmDrf)
            ilRet = btrClose(hmDpf)
            ilRet = btrClose(hmAnf)
            btrDestroy hmSof
            btrDestroy hmUrf
            btrDestroy hmMnf
            btrDestroy hmSlf
            btrDestroy hmRdf
            btrDestroy hmDnf
            btrDestroy hmAgf
            btrDestroy hmAdf
            btrDestroy hmCff
            btrDestroy hmClf
            btrDestroy hmCHF
            btrDestroy hmTChf
            btrDestroy hmCbf
            btrDestroy hmCxf
            btrDestroy hmDrf
            btrDestroy hmDpf
            btrDestroy hmAnf
            ilValue = Asc(tgSpf.sSportInfo)
            If (ilValue And USINGSPORTS) = USINGSPORTS Then 'Using Sports, close game files
                ilRet = btrClose(hmCgf)
                ilRet = btrClose(hmGsf)
                btrDestroy hmCgf
                btrDestroy hmGsf
            End If
            Erase llPopByLine           '11-23-99
            Erase lmWklySpots, lmWklyRates, lmAvgAud, lmPopEst
            Erase llStdStartDates
            Erase ilPkgVehList, ilPkgLineList, ilVehList
            If ilTask = REPORTSJOB Then
                Erase tgClf, tgCff
            End If
            Exit Function


        End If
        '12-16-03
        igBR_SSLinesExist = False         'if selective contract, dont show if no sch lines; if more than 1 contract,
                                        'show when theres at least 1 sch line to a contract

        '7-23-01 setup global variable for Demo Plus file (to see if any exists)
        lgDpfNoRecs = btrRecords(hmDpf)
        If lgDpfNoRecs = 0 Then
            lgDpfNoRecs = -1
        End If

        ilStdQtrLess13 = False
        If tgSpf.sSBrStdQt = "Y" Then               'always show std qtr regardless if less than 13 week order
            ilStdQtrLess13 = True
        End If

        imHiddenOverride = Asc(tgSpf.sUsingFeatures)                '3-10-06 determine if usinig hidden overrides on package hidden lines

        bmShowAudIfPodcast = True                                   '4-18-18 default to show aud % for Podcast vehicles
#If programmatic <> 1 Then
        If BrSnap!ckcProdcastInfo.Value = vbUnchecked Then
            bmShowAudIfPodcast = False
        End If
#Else
        bmShowAudIfPodcast = False
#End If
'Get all company competitor records - specifically get "Us"
        ilRet = gObtainMnfForType("O", slTimeStamp, tlMMnf())
        For ilLoop = LBound(tlMMnf) To UBound(tlMMnf) Step 1
            If tlMMnf(ilLoop).iGroupNo = 1 Then
                ilCmpny = tlMMnf(ilLoop).iCode
            End If
        Next ilLoop
        
        gPopAnf hmAnf, tmAnfTable()              'populate the Named Avails in array for DP Overrides printing rules
         'billed receivables, future will be calculated
        gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slLastBillDate
        If slLastBillDate = "" Then
            slLastBillDate = "1/1/1970"
        End If
        llLastBillDate = gDateValue(slLastBillDate)

        tmTranTypes.iAdj = False
        tmTranTypes.iAirTime = True
        tmTranTypes.iCash = True
        tmTranTypes.iInv = True
        tmTranTypes.iMerch = False
        tmTranTypes.iNTR = True
        tmTranTypes.iPromo = True
        tmTranTypes.iPymt = False
        tmTranTypes.iTrade = False
        tmTranTypes.iWriteOff = False

        'if Installment applicable, setup types of records to retrieve from SBF
        tlInstallSBFType.iNTR = False
        tlInstallSBFType.iInstallment = True
        tlInstallSBFType.iImport = False

    '           Valid contract read into memory.  Cycle thru sch lines,
    '           build array of weeks research, spots, $ and create
    '           quarterly records for output.
    '

            'retrieve R/C, items, and DP info for the rate card that contract uses
#If programmatic <> 1 Then
            gRCRead Contract, tgChf.iRcfCode, tmRcf, tmMRif(), tmMRDF()
#Else
            gRCRead ProgrammaticBuy, tgChf.iRcfCode, tmRcf, tmMRif(), tmMRDF()
#End If

            ilCurrStartQtr(0) = 0                       'init to place start date of starting qtr of order
            ilCurrStartQtr(1) = 0
            ilCurrTotalMonths = 0
            'Set up earliest/latest dates of contr, set to std dates.  Set array of starting bdcst months for summary page
            'mFindMaxDates slStartDate, slEndDate, llChfStart, llChfEnd, ilCurrStartQtr(), ilCurrTotalMonths, llStdStartDates()

            For ilDemo = 0 To 3 Step 1                  'user input Entered and Active dates
            If tgChf.iMnfDemo(ilDemo) > 0 Or ilDemo = 0 Then
            tmCbf = tmZeroCbf

            ReDim Preserve ilVehiclesDone(0 To 0) As Integer    'installmentvehicles gathered for monthly summary
            'ReDim Preserve llInstallBilling(1 To 13) As Long    '12 months max of installment billing, anything over is accum in one bucket
            ReDim Preserve llInstallBilling(0 To 13) As Long    '12 months max of installment billing, anything over is accum in one bucket. Index zero ignored
            
            tmCbf.iVehSort = ilDemo             '8-21-15 for sorting of multiple demos, to keep in same order as entered
            tmCbf.sHiddenOverride = "N"         '3-10-06 assume not using hidden override feature from Site
            llCurrMod = 0                       'init total $ this version
            llCurrModSpots = 0                  'init total spot count this version
            imDiffExceeds104Wks = False         '3-9-06 this is for difference only version
            'move to BR driver
            'ReDim Preserve lgPrintedCnts(1 To ilUpperPrint) As Long
            'lgPrintedCnts(ilUpperPrint) = tgChf.lCode
            'ilUpperPrint = ilUpperPrint + 1

'11-14-11   account for 14 week qtr
            ReDim lmWklySpots(0 To MAXWEEKSFOR2YRS, 0 To 0) As Long           'array of spots by week for one sched line, reqd for research results. Index zero ignored
            ReDim lmWklyRates(0 To MAXWEEKSFOR2YRS, 0 To 0) As Long              'array of rates by week for one sch line, reqd for research results. Index zero ignored
            ReDim lmAvgAud(0 To MAXWEEKSFOR2YRS, 0 To 0) As Long                 'array of avg aud by week for one sch line, reqd for research. Index zero ignored
            ReDim lmPopEst(0 To MAXWEEKSFOR2YRS, 0 To 0) As Long        'Index zero ignored
            
            ReDim lmPop(0 To 0) As Long                              'population per line
            ReDim lmPopPkg(0 To 0) As Long                          '11-24-04 pop for the pkg because of using same hidden vehicle reference as the pkg vehicle
            ReDim tmLRch(0 To 1) As RESEARCHINFO    '10-30-01. Index zero ignored
            'ReDim tmPrevLRch(1 To 1) As RESEARCHLIST
'11-14-11 account for 14 week qtr
            ReDim lmSpotsByWk(0 To MAXWEEKSFOR2YRS, 0 To 1) As Long  'array of 105 weeks (2 yrs) containing spot counts for unique spot rate. Index zero ignored
                                                            'each one of these 104 week arrays correspond to one tmLnr entry
            ReDim lmratesbywk(0 To MAXWEEKSFOR2YRS, 0 To 1) As Long  'array of 105 weeks (2 yrs) containing spot rates per week. Index zero ignored
                                                            'each one of these 104 week arrays correspond to one tmLnr entry
                                                            'each one of these 104 week arrays correspond to one tmLnr entry
            ReDim tmLnr(0 To 1) As LNR                        'array of unique spot rates for a line.  One line can vary in
                                                            'rates from week to week.  By building all lines in memory at once,
                                                            'we can get the actual # of airing weeks.  For Differences option,
                                                            'previous versions schedule lines are also built; for non-differences
                                                            'option, only the current lines are built
                                                            'i.e.  Line #       Rate
                                                            '        1           25000
                                                            '        1           27500
                                                            '        2           25000
                                                            '        3           40000
                                                            'Index zero ignored
            ReDim imDnfCodes(0 To 0) As Integer               'array of Demo Research names used from contr
            
            ReDim imProcessFlag(0 To UBound(tgClf)) As Integer          'flags to indicate how to process line (show current vs mods, differences)
            ilLineRate = 1
            ilLineSpots = 1
            ilUpperClf = 0                                  'no lines in research table
            
            '3-31-21 initialize the variables to combine air time and adserver Impressions and cost to calculate CPM on the Research Summary page
            sgAirTimeGrimp = ""
            lgAirTimeGross = 0

            ilUpperDnf = UBound(imDnfCodes) '1
            ilLoop = 0
            llResearchPop = -1                       'pop from book if all same books across lines, else
                                                    'its zero
            gUnpackDate tgChf.iStartDate(0), tgChf.iStartDate(1), slStartDate
            If slStartDate = "" Then
                llDate = 0
            Else
                llDate = gDateValue(slStartDate)
            End If
            gUnpackDate tgChf.iEndDate(0), tgChf.iEndDate(1), slEndDate
            If slEndDate = "" Then
                llDate2 = 0
            Else
                llDate2 = gDateValue(slEndDate)
            End If
                'get the Installment records, if applicable
            ReDim tlSbf(0 To 0) As SBF
#If programmatic <> 1 Then
            ilRet = gObtainSBF(Contract, hmSbf, tgChf.lCode, slStartDate, slEndDate, tlInstallSBFType, tlSbf(), 0)
#Else
            ilRet = gObtainSBF(ProgrammaticBuy, hmSbf, tgChf.lCode, slStartDate, slEndDate, tlInstallSBFType, tlSbf(), 0)
#End If

            'Set up earliest/latest dates of contr, set to std dates.  Set array of starting bdcst months for summary page
            blDefaultToQtr = True
            gFindMaxDates slStartDate, slEndDate, llChfStart, llChfEnd, ilCurrStartQtr(), ilCurrTotalMonths, llStdStartDates(), ilShowStdQtr, ilCorpStdYear, ilStartMonth, blDefaultToQtr

            'determine first month in the future (after last billing date)
            llTaxDate = gDateValue(slLastBillDate)
            If llTaxDate < llStdStartDates(1) Then        'everything is in the future
                ilLastBilledInx = 1
             ElseIf llTaxDate >= llStdStartDates(37) Then       'everything in past
                ilLastBilledInx = 1
            Else
                'contract spans the last bill date
                ilLastBilledInx = 1
                For ilLoop = 1 To 36 Step 1
                    If llTaxDate > llStdStartDates(ilLoop) And llTaxDate < llStdStartDates(ilLoop + 1) Then
                        ilLastBilledInx = ilLoop
                        Exit For
                    End If
                Next ilLoop
            End If
            
            '11-14-11 Some quarters may be 14 week qtrs; default all to 13 week qtrs and set the quarter than is non-std
            For ilLoop = 1 To 8
                imWeeksPerQtr(ilLoop) = 13
            Next ilLoop

            'build table of how to process the lines for the CBF file and output shown on BR
            'for differences, gather the current lines and most recent mod
            'for Full BR, gather only current lines
            'array corresponds to the lines built in tgclf:
            '0 = ignore altogether, old mod.
            '1 = previous mod, process for BR
            '2 = don't show on BR but use to calc flights because its a current line from an older mod,
            '    and need to figure out # of active weeks vs actual weeks
            '3 = current, same revision # as header, need to show and process
            '4 = hidden line shown on proof. Show it, but don't use in any calculations of flights
            If UBound(tgClf) > UBound(llPopByLine) Then                 '12-20-99
                'ReDim llPopByLine(1 To UBound(tgClf)) As Long               '11-23-99
                ReDim llPopByLine(0 To UBound(tgClf)) As Long               '11-23-99. Index zero ignored
            End If
            slDailyExists = "N"           '5-20-30 assume no dailys on this order
            ReDim tlSofList(0 To 0) As SOFLIST  'make use of existing table to store the vehicle and flag if there is at least one day for the vehicle.
                                                'this is for Insertion Orders only because one vehicle may be only weekly and another daily; but
                                                'dont want to show dotted vertical lines for the weekly only vehicles

            ReDim tmPodcast_Info(0 To 0) As PODCAST_INFO         '3-9-18
            mSetupBrHdr tlBR.iWhichSort         'build the advt, agy, slsp and sort fields specs that ned to be built
                                                'only once per contract

            If tgChf.iAgfCode > 0 Then        'get agy comm in case taxes need to be obtained
                ilAgyComm = tmAgf.iComm
            Else
                ilAgyComm = 0                       'direct
            End If
            
            tmCbf.sRschColHdr = ""               'init to show column headers for avg rtg, grp and cpp
            tmCbf.sMixTypes = ""                 'init to show field values for avg rtg, grp & cpp
            tmCbf.iExtra2Byte = -1                  '1-28-10 key record required to link NTR and air time for billing summary
            ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)

            For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                tmClf = tgClf(ilClf).ClfRec

                gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), slStartDate
                gUnpackDate tmClf.iEndDate(0), tmClf.iEndDate(1), slEndDate

                'test for cancel before start which is always start date one day later that end date
                'This is code to fix up schedule lines that had bad start & end dates in them.  They didnt coincide with the flights or header
                If ((gDateValue(slStartDate) - gDateValue(slEndDate)) <> 1) Then      'not cancel before start
                    If gDateValue(slStartDate) < llChfStart Or gDateValue(slStartDate) > llChfEnd Then
                        slStartDate = Format$(llChfStart, "m/d/yy")
                    End If
                    If gDateValue(slEndDate) < llChfStart Or gDateValue(slEndDate) > llChfEnd Then
                        slEndDate = Format$(llChfEnd, "m/d/yy")
                    End If
                End If
                If slStartDate = "" Then
                    llFltStart = 0
                Else
                    llFltStart = gDateValue(slStartDate)
                End If
                If slEndDate = "" Then
                    llFltEnd = 0
                Else
                    llFltEnd = gDateValue(slEndDate)
                End If
                If tmClf.sType = "H" Then         'Hidden?, if so-ignore if not proof option
                    imProcessFlag(ilClf) = 3
                Else
                    imProcessFlag(ilClf) = 3
                End If
            Next ilClf
            'Adjust dates to Monday and to start of the quarter.  Just obtained the earliest and latest
            'dtes of the contract.  If differences, considered all past lines for earliest and latest dates
            '
            'backup start date of contract to Monday, then decide how many quarters to process for this contract.
            'Always show data from a start quarter unless the total # of weeks is less than 13.
            '
            'llDate = gDateValue(slStartDate)
            'llDate contains earliest start date of sch lines

            ilLoop = gWeekDayLong(llDate)
            Do While ilLoop <> 0
                llDate = llDate - 1
                ilLoop = gWeekDayLong(llDate)
            Loop
            slStartDate = Format$(llDate, "m/d/yy")
            slEndDate = Format$(llDate2, "m/d/yy")
            'gUnpackDate tgChf.iEndDate(0), tgChf.iEndDate(1), slEndDate
            ilRealTotalWks = (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 7 + 1
            If ilRealTotalWks < 0 Then          'this condition only for Cancel Before start contracts
                                                'where there isn't a start & end date in header
                ilRealTotalWks = 1
            End If
            ilTotalWks = ilRealTotalWks         'this will be the total number of weeks including the start of the qtr
            If (ilTotalWks > 13) Or (ilStdQtrLess13) Then            'only show complete quarters on BR if the schedule lasts longers
                                                'than 13 weeks, or site is set up to always show on start of std qtrs
                slStartDate = gObtainEndStd(slStartDate)        'get std bdcst end date
                gObtainYearMonthDayStr slStartDate, True, slYear, slMonth, slDay
                Do While (slMonth <> "01") And (slMonth <> "04") And (slMonth <> "07") And (slMonth <> "10")
                    slMonth = str$((Val(slMonth) - 1))
                    slDay = "15"
                    slStartDate = slMonth & "/" & slDay & "/" & slYear
                    gObtainYearMonthDayStr slStartDate, True, slYear, slMonth, slDay
                    slStartDate = gObtainEndStd(slStartDate)        'get std bdcst end date
                Loop
                'slStartDate now contains the end date of a quarter
                slStartDate = gObtainStartStd(slStartDate)      'start date of the quarter for this cnt
                'recalculate the totals wks because the start date was pushed back to the start of the qtr
                ilTotalWks = (gDateValue(slEndDate) - gDateValue(slStartDate)) \ 7 + 1
            
                slTempDays = slStartDate
                llDate = gDateValue(slTempDays)
                
                '11-14-11 determine if any quarter has 14 weeks to set the array of weeks per qtr for 8 quarters
                For ilLoop = 1 To 8
                    slTempDays = Format$((llDate + 75), "m/d/yy")
                    slTempDays = gObtainEndStd(slTempDays)
                    llDate2 = gDateValue(slTempDays)
                    If (llDate2 - llDate) = 83 Then          '12 week qtr
                        imWeeksPerQtr(ilLoop) = 12
                    ElseIf (llDate2 - llDate) > 90 Then            '14 week qtr
                        imWeeksPerQtr(ilLoop) = 14
                    End If
                    llDate = llDate2 + 1        'start of next qtr
                    slTempDays = Format$(llDate, "m/d/yy")
            
                 Next ilLoop
            End If                                              'if ilTotalsWks > 13
            'slStartDate contains the start of quarter or week that is to being on BR
            '4-12-19 2 years (8 qtrs is the max to process for any contract).  Process only the number of quarters on a contract to speed up processing.
            imMaxQtrs = 8
            'llTemp = gDateValue(slStartDate)
            dlTemp = gDateValue(slStartDate) 'TTP 10439 - Rerate 21,000,000
            For ilLoop = 1 To 8
                'llTemp = (7 * imWeeksPerQtr(ilLoop)) + llTemp
                dlTemp = (7 * imWeeksPerQtr(ilLoop)) + dlTemp 'TTP 10439 - Rerate 21,000,000
                'If llTemp > gDateValue(slEndDate) Then
                If dlTemp > gDateValue(slEndDate) Then 'TTP 10439 - Rerate 21,000,000
                    imMaxQtrs = ilLoop
                    Exit For
                End If
            Next ilLoop
            If imMaxQtrs > 8 Then
                imMaxQtrs = 8
            End If
            ReDim Preserve ilVehList(0 To 0) As Integer
            ReDim Preserve ilPkgVehList(0 To 0) As Integer
            ReDim Preserve ilPkgLineList(0 To 0) As Integer
            ReDim lmPopPkgByLine(0 To 1) As Long                '11-11-05. Index zero ignored

            If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then        '12-22-11 test if ntr only taxes
                'if taxes apply, find all IN receivables prior to last billing date.  Add up tax1 and tax2 amounts
                'When going thru each line, calc the taxes after last billing date
#If programmatic <> 1 Then
                ilRet = gObtainPhfRvfbyCntr(Contract, tgChf.lCntrNo, "1/1/1970", slLastBillDate, tmTranTypes, tlRvf())
#Else
                ilRet = gObtainPhfRvfbyCntr(ProgrammaticBuy, tgChf.lCntrNo, "1/1/1970", slLastBillDate, tmTranTypes, tlRvf())
#End If
                llTax1CalcAmt = 0
                llTax2CalcAmt = 0
                'accumulate taxes for IN airtime only thru the end of the last bill date, NTR will be processed later
                For ilLoop = LBound(tlRvf) To UBound(tlRvf) - 1
                    gUnpackDateLong tlRvf(ilLoop).iTranDate(0), tlRvf(ilLoop).iTranDate(1), llDate
                    If tlRvf(ilLoop).iMnfItem = 0 And llDate <= llLastBillDate Then
                        llTax1CalcAmt = llTax1CalcAmt + tlRvf(ilLoop).lTax1
                        llTax2CalcAmt = llTax2CalcAmt + tlRvf(ilLoop).lTax2
                    End If
                Next ilLoop
                For ilLoop = LBound(llProjectedTax1) To UBound(llProjectedTax2) - 1
                    llProjectedTax1(ilLoop) = 0
                    llProjectedTax2(ilLoop) = 0
                Next ilLoop
            End If

            'cycle thru all the lines and build arrays that have same line, vehicle, daypart, length, & rate.  Each of these unique entries
            'shows on a separate line on the printed contract.
            ReDim tlPkgVehicleForSurveyList(0 To 0) As PKGVEHICLEFORSURVEYLIST          '12-19-13
  
            For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                ReDim Preserve lmWklySpots(0 To MAXWEEKSFOR2YRS, 0 To ilUpperClf) As Long          'array of spots by week for one sched line, reqd for research results. Index zero ignored
                ReDim Preserve lmWklyRates(0 To MAXWEEKSFOR2YRS, 0 To ilUpperClf) As Long             'array of rates by week for one sch line, reqd for research results. Index zero ignored
                ReDim Preserve lmAvgAud(0 To MAXWEEKSFOR2YRS, 0 To ilUpperClf) As Long                'array of avg aud by week for one sch line, reqd for research. Index zero ignored
                ReDim Preserve lmPopEst(0 To MAXWEEKSFOR2YRS, 0 To ilUpperClf) As Long          'Index zero ignored
                tmClf = tgClf(ilClf).ClfRec
                
                ilVefIndex = gBinarySearchVef(tmClf.iVefCode)
                If ilVefIndex <> -1 Then        '10-23-06 if new package vehicle entered for proposal, the index wont exist.  The vehicle
                    '3-9-18 determine if podcast vehicles to show rating info or not
                    'determine if combination of podcast cast vs no podcast vehicles for output to show avgrtg, grp or ccp
                    ilVpfIndex = gBinarySearchVpf(tmClf.iVefCode)
                    If ilVpfIndex <> -1 Then
                        'Build array from sch lines that contain the type of line (podcast, non podcast, package , non package) to
                        'help determine if the line should show grp,cpp or avg rtg on the contract
                        tmPodcast_Info(UBound(tmPodcast_Info)).iLine = tmClf.iLine
                        If tgVpf(ilVpfIndex).sGMedium = "P" Then
                            If bmShowAudIfPodcast Then              '4-18-18 show Aud % for Podcast vehicles?
                                'override the feature and show all Aud % for this podcast vehicle
                                If tmClf.sType = "H" Then
                                    tmPodcast_Info(UBound(tmPodcast_Info)).sType = "L"      'force to other non-pkg hidden line
                                Else
                                    If tmClf.sType = "O" Or tmClf.sType = "A" Then          'pkg line?
                                        tmPodcast_Info(UBound(tmPodcast_Info)).sType = "K"      'force to other non-pkg hidden line
                                    Else
                                        tmPodcast_Info(UBound(tmPodcast_Info)).sType = "L"      'force to other non-pkg hidden line
                                    End If
                                End If
                            Else                                    '4-18-18 do not show aud % for Podcast vehicle
                                'podcast:  is it a hidden line of package or not
                                If tmClf.sType = "H" Then
                                    'type of line:  P = Podcast, K = package line, H = Podcast Hidden Line, L = Other, not podcast Hidden Line, O = other, not poscast (conventional, selling)
                                    tmPodcast_Info(UBound(tmPodcast_Info)).sType = "H"
                                Else
                                    If tmClf.sType = "O" Or tmClf.sType = "A" Then      'package line?
                                        tmPodcast_Info(UBound(tmPodcast_Info)).sType = "K"
                                    Else
                                        tmPodcast_Info(UBound(tmPodcast_Info)).sType = "P"      'podcast vehicle, not within a package
                                    End If
                                End If
                            End If
                        Else            'not podcast
                            If tmClf.sType = "H" Then
                                tmPodcast_Info(UBound(tmPodcast_Info)).sType = "L"      'not podcast, hidden line
                            Else
                                If tmClf.sType = "O" Or tmClf.sType = "A" Then          'package line?
                                    tmPodcast_Info(UBound(tmPodcast_Info)).sType = "K"
                                Else                                                    'standard line, not a podcast
                                    tmPodcast_Info(UBound(tmPodcast_Info)).sType = "O"
                                End If
                            End If
                        End If
                        tmPodcast_Info(UBound(tmPodcast_Info)).iPkgRefLine = tmClf.iPkLineNo
                        tmPodcast_Info(UBound(tmPodcast_Info)).bShowResearch = True             'assume to show research fields
                        tmPodcast_Info(UBound(tmPodcast_Info)).iVefCode = tmClf.iVefCode
                        ReDim Preserve tmPodcast_Info(0 To UBound(tmPodcast_Info) + 1) As PODCAST_INFO
                    End If
                End If

                If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Then
                    'determine the tax amt for future dates (after last billing date)
                    'a contract can be max 3 years
                    gGetAirTimeTaxRates tgChf.iAdfCode, tgChf.iAgfCode, tmClf.iVefCode, llTax1Pct, llTax2Pct, slGrossNet
                    gBuildFlights ilClf, llStdStartDates(), ilLastBilledInx, 36, llProjectedFlights(), 1, tgClf(), tgCff()
                    gFutureTaxes ilLastBilledInx, llProjectedFlights(), llProjectedTax1(), llProjectedTax2(), ilAgyComm, tgChf.iPctTrade, llTax1Pct, llTax2Pct, slGrossNet
                End If

                'Build table of unique vehicles and whether they have at least one daily schedule line (for Insertion Orders) to show dotted vertical lines
                'use existing variable from another report
                ilFoundLine = False
                For ilTemp = LBound(tlSofList) To UBound(tlSofList) - 1
                    If tlSofList(ilTemp).iMnfSSCode = tmClf.iVefCode Then
                        ilFoundLine = True
                        Exit For
                    End If
                Next ilTemp
                If Not ilFoundLine Then
                    tlSofList(UBound(tlSofList)).iMnfSSCode = tmClf.iVefCode         'setup unqiue vehicle and assume not dailies yet
                    tlSofList(UBound(tlSofList)).iSofCode = False       'assume no daily exists yet for this vehicle
                    ReDim Preserve tlSofList(UBound(tlSofList) + 1) As SOFLIST
                End If

                gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), slStr
                llFltStart = gDateValue(slStr)
                gUnpackDate tmClf.iEndDate(0), tmClf.iEndDate(1), slStr
                llFltEnd = gDateValue(slStr)
                If llFltStart >= gDateValue(slStartDate) Then       '6-15-04 ignore the CBS lines that are outside the contract dates
                                                                    'including the CBS line outside the contract dates causes problems in research numbers
                    If llFltEnd >= llFltStart Then
                        ilCBS = False
                        'Build array of the unique package vehicle names
                        '11-11-05 build array of the package vehicle pops (lmpoppkgbyLine) associated with the package line line (ilPkgLineList)
                        If tmClf.sType = "O" Or tmClf.sType = "A" Or tmClf.sType = "E" Then      'pkg line
                            'build array of the Package line #s & vehicle
                            ilFoundLine = False
                            For ilPkg = LBound(ilPkgLineList) To UBound(ilPkgLineList) - 1 Step 1
                                If tmClf.iLine = ilPkgLineList(ilPkg) Then
                                    ilFoundLine = True
                                    Exit For
                                End If
                            Next ilPkg
                            If Not ilFoundLine Then
                                ilPkgLineList(UBound(ilPkgLineList)) = tmClf.iLine
                                ReDim Preserve ilPkgLineList(0 To UBound(ilPkgLineList) + 1)
                                ReDim Preserve lmPopPkgByLine(0 To UBound(lmPopPkgByLine) + 1)      '11-11-05 package pop (not by unqiue packages vehicle names, but each individual pkg). Index zero ignored
                                ilPkgVehList(UBound(ilPkgVehList)) = tmClf.iVefCode
                                ReDim Preserve ilPkgVehList(0 To UBound(ilPkgVehList) + 1)
                            End If
                        Else                        '12-19-13 not package; build association with package vehicle if hidden
                            If tmClf.sType = "H" Then
                                If tmClf.iDnfCode > 0 Then
                                    'got a hidden line, find the associated package line to get the vehicle code
                                    ilSavePkVeh = -1
                                    For ilPkg = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                                        If tmClf.iPkLineNo = tgClf(ilPkg).ClfRec.iLine Then   'found the associated pkg line
                                            ilSavePkVeh = tgClf(ilPkg).ClfRec.iVefCode        'save the associated pkg vehicle code
                                            Exit For
                                        End If
                                    Next ilPkg
                                    
                                    ilFoundLine = False
                                    For ilPkg = 0 To UBound(tlPkgVehicleForSurveyList) - 1 Step 1
                                        If ilSavePkVeh = tlPkgVehicleForSurveyList(ilPkg).iPkgVefCode And tmClf.iDnfCode = tlPkgVehicleForSurveyList(ilPkg).iHiddenDnfCode Then
                                            ilFoundLine = True
                                            Exit For
                                        End If
                                    Next ilPkg
                                    If (Not ilFoundLine) And (ilSavePkVeh > 0) Then
                                        tlPkgVehicleForSurveyList(UBound(tlPkgVehicleForSurveyList)).iPkgVefCode = ilSavePkVeh
                                        tlPkgVehicleForSurveyList(UBound(tlPkgVehicleForSurveyList)).iHiddenDnfCode = tmClf.iDnfCode
                                        ReDim Preserve tlPkgVehicleForSurveyList(0 To UBound(tlPkgVehicleForSurveyList) + 1) As PKGVEHICLEFORSURVEYLIST
                                    End If
                                End If
                            End If
                        End If

                        ilFoundLine = False
                        For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1 Step 1
                            If tmClf.iVefCode = ilVehList(ilVehicle) Then
                                ilFoundLine = True
                                Exit For
                            End If
                        Next ilVehicle
                        If Not ilFoundLine Then
                            ilVehList(UBound(ilVehList)) = tmClf.iVefCode
                            ilVehicle = UBound(ilVehList)
                            ReDim Preserve ilVehList(0 To UBound(ilVehList) + 1)
                            ReDim Preserve lmPop(0 To UBound(lmPop) + 1)
                            ReDim Preserve lmPopPkg(0 To UBound(lmPopPkg) + 1)  '11-24-04 pop for the pkg because of using same hidden vehicle reference as the pkg vehicle
                        End If

                        'Build population table by vehicle
                        ilRet = gGetDemoPop(hmDrf, hmMnf, hmDpf, tmClf.iDnfCode, tlBR.iSocEcoMnfCode, tgChf.iMnfDemo(ilDemo), llPop)
                        If tmClf.iDnfCode = 0 Then                      '3-12-19 if no book assigned to line, ignore the population
                            llPop = 0
                        End If
                        If tmClf.sType = "H" Or tmClf.sType = "S" Then      '11-24-04 exclude packages
                            '5-28-04
                            If tgSpf.sDemoEstAllowed <> "Y" Then
                                llPopByLine(ilClf + 1) = llPop    'save the pop by line (each linemay have different survey books)
                                If lmPop(ilVehicle) = 0 Then            'same vehicle, dont wipe out the pop if already obtained, current one could be zero
                                    lmPop(ilVehicle) = llPop            'associate the population with the vehicle
                                End If
                                If llResearchPop = -1 And llPop <> 0 Then           'first time , but when a population exists
                                    llResearchPop = llPop
                                Else
                                    If (llResearchPop <> 0) And (llResearchPop <> llPop) And (llPop <> 0) Then      'test to see if this pop is different that the prev one.
                                        If tmClf.iDnfCode > 0 Then                      '3-12-19 ignore this line pop if no book defined
                                            llResearchPop = 0                                           'if different pops, calculate the contract  summary different
                                        End If
                                        If llPop <> lmPop(ilVehicle) Then  '11-30-99
                                            lmPop(ilVehicle) = -1          '11-30-99
                                        End If                             '11-30-99

                                    Else
                                        If llPop <> 0 And (llResearchPop <> 0 And llResearchPop <> -1) Then '2/1/99
                                            llResearchPop = llPop
                                        Else        '5-8-02
                                            If lmPop(ilVehicle) <> llPop And lmPop(ilVehicle) <> -1 Then
                                                lmPop(ilVehicle) = -1
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        ilFoundLine = False
                        For ilLoop = LBound(imDnfCodes) To ilUpperDnf - 1 Step 1
                            If imDnfCodes(ilLoop) = tmClf.iDnfCode Then
                                ilFoundLine = True
                            End If
                        Next ilLoop
                        If Not (ilFoundLine) Then
                            ReDim Preserve imDnfCodes(0 To ilUpperDnf) As Integer
                            imDnfCodes(ilUpperDnf) = tmClf.iDnfCode
                            ilUpperDnf = ilUpperDnf + 1
                        End If

                    Else
                        ilCBS = True
                    End If
                    ilManyFlts = False
                    ilCff = tgClf(ilClf).iFirstCff
                    Do While ilCff <> -1
                        If tgCff(ilCff).iStatus = 0 Or tgCff(ilCff).iStatus = 1 Then
                            tmCff = tgCff(ilCff).CffRec
                            If ilVefIndex <> -1 Then        '10-23-06 if new package vehicle entered for proposal, the index wont exist.  The vehicle
                                'name wont show on the printout since the vehicle hasnt been written to disk yet
                                If tgMVef(ilVefIndex).sType = "G" Then          'sports vehicle
                                    tmCff.sDyWk = "W"
                                End If
                                
                            End If
                            ilfirstTime = True                  'set to calc avg aud one time only for this flight
                            If tmCff.sDyWk = "W" Or ilCBS Then           '6-20-03
                                slTempDays = gDayNames(tmCff.iDay(), tmCff.sXDay(), 2, slStr)            'slstr not needed when returned
                                slStr = ""
                                For ilLoop = 1 To Len(slTempDays) Step 1
                                    slYear = Mid$(slTempDays, ilLoop, 1)
                                    If slYear <> " " And slYear <> "," Then
                                        slStr = Trim$(slStr) & Trim$(slYear)
                                    End If
                                Next ilLoop
                            Else                '11-19-02
                                'Setup # spots/day
                                If tmCff.sDyWk = "D" Then            '5-20-03 daily
                                    slDailyExists = "Y"
                                    For ilTemp = LBound(tlSofList) To UBound(tlSofList) - 1
                                        If tlSofList(ilTemp).iMnfSSCode = tmClf.iVefCode Then
                                            tlSofList(ilTemp).iSofCode = True       'set flag to indicate at least 1 daily for this vehicle
                                            Exit For
                                        End If
                                    Next ilTemp
                                End If
                                slStr = ""
                                For ilLoop = 0 To 6
                                    slSpotCount = Trim$(str$(tmCff.iDay(ilLoop)))
                                    Do While Len(slSpotCount) < 4
                                        slSpotCount = " " & slSpotCount
                                    Loop
                                    slStr = slStr & " " & slSpotCount
                                Next ilLoop
                            End If
                            slValidDays = ""
                            slValidDays = slStr
                            For ilLoop = 0 To 6                 'init all days to not airing, setup for research results later
                                ilInputDays(ilLoop) = False
                                ilDemoAvgAudDays(ilLoop) = False        'initalize to 0
                            Next ilLoop
                            ilFoundLine = False
                            For ilLineSpotsInx = 1 To ilLineRate - 1 Step 1
                                If (tmLnr(ilLineSpotsInx).iLineInx = ilClf) Then            'same lines
                                    If ((tmCff.lActPrice = tmLnr(ilLineSpotsInx).lRate) And (slValidDays = tmLnr(ilLineSpotsInx).sValidDays)) Then
                                        ilFoundLine = True
                                        Exit For
                                    Else
                                        ilManyFlts = True
                                    End If
                                End If
                            Next ilLineSpotsInx
                            'maintain ilLineSpotsInx (if found) so that we can get back to the data in that entry
                            If Not ilFoundLine Then                         'create the 2 year buffer for spot counts
                                ReDim Preserve tmLnr(0 To ilLineRate) As LNR    'Index zero ignored
                                ReDim Preserve lmSpotsByWk(0 To MAXWEEKSFOR2YRS, 0 To ilLineSpots) As Long  'Index zero ignored
                                ReDim Preserve lmratesbywk(0 To MAXWEEKSFOR2YRS, 0 To ilLineSpots) As Long  'Index zero ignored
                                tmLnr(ilLineRate).iLineInx = ilClf
                                tmLnr(ilLineRate).lRate = tmCff.lActPrice
                                tmLnr(ilLineRate).sPriceType = tmCff.sPriceType     '2-23-01
                                tmLnr(ilLineRate).sValidDays = slValidDays
                                tmLnr(ilLineRate).iManyFlts = ilManyFlts            'true if more than 1 flt this sched line
                                ilLineSpotsInx = ilLineSpots            'need to get to this entry to accum spots later
                                ilLineRate = ilLineRate + 1
                                ilLineSpots = ilLineSpots + 1
                            End If

                            gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
                            llFltStart = gDateValue(slStr)
                            'backup start date to Monday
                            ilLoop = gWeekDayLong(llFltStart)
                            Do While ilLoop <> 0
                                llFltStart = llFltStart - 1
                                ilLoop = gWeekDayLong(llFltStart)
                            Loop
                            gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
                            llFltEnd = gDateValue(slStr)
                            '
                            'Loop thru the flight by week and build the number of spots for each week
                            For llDate2 = llFltStart To llFltEnd Step 7
                                If tmCff.sDyWk = "W" Then            'weekly
                                    ilSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                                    For ilDay = 0 To 6 Step 1
                                    If (llDate2 + ilDay >= llFltStart) And (llDate2 + ilDay <= llFltEnd) Then
                                        If tmCff.iDay(ilDay) > 0 Or tmCff.sXDay(ilDay) = "1" Then
                                            ilInputDays(ilDay) = True
                                            ilDemoAvgAudDays(ilDay) = True      '11-19-02 for weekly, each day is indicated by true false as a valid airing day of week
                                        End If
                                    End If
                                    Next ilDay
                                Else                                        'daily
                                    If ilLoop + 6 < llFltEnd Then           'we have a whole week
                                        ilSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)    'got entire week
                                        For ilDay = 0 To 6 Step 1
                                            If tmCff.iDay(ilDay) > 0 Then
                                                ilInputDays(ilDay) = tmCff.iDay(ilDay)          '11-19-02 chged from true/false to # spots/day
                                                ilDemoAvgAudDays(ilDay) = True      '11-19-02 for daily, each day is indicated by # spots per day as a valid airing day
                                             End If
                                        Next ilDay
                                    Else                                    'do partial week
                                        For llDate = llDate2 To llFltEnd Step 1
                                            ilDay = gWeekDayLong(llDate)
                                            ilSpots = ilSpots + tmCff.iDay(ilDay)
                                            If tmCff.iDay(ilDay) > 0 Then
                                                ilInputDays(ilDay) = tmCff.iDay(ilDay)          '11-19-02 chged from true/false to # spots/day
                                                ilDemoAvgAudDays(ilDay) = True      '11-19-02 for daily, each day is indicated by # spots per day as a valid airing day
                                            End If
                                        Next llDate
                                    End If
                                End If
                                'Days of week from flight - saved to compare against daypart days for overrides
                                tmLnr(ilLineSpotsInx).sDailyWkly = tmCff.sDyWk  '5-20-03 daily vs weekly flag
                                For ilLoop = 0 To 6 Step 1
                                    tmLnr(ilLineSpotsInx).iCffDays(ilLoop) = ilInputDays(ilLoop)
                                Next ilLoop

                                ilWeekInx = (llDate2 - gDateValue(slStartDate)) / 7 + 1
                                If ilWeekInx > 0 And ilWeekInx < 105 Then           ' < 1, its a CBS .  3-9-06 cant exceed 2 yrs or subscript error
                                    lmSpotsByWk(ilWeekInx, ilLineSpotsInx) = lmSpotsByWk(ilWeekInx, ilLineSpotsInx) + ilSpots
                                    lmratesbywk(ilWeekInx, ilLineSpotsInx) = tmCff.lActPrice
                                    lmWklySpots(ilWeekInx, ilUpperClf) = lmWklySpots(ilWeekInx, ilUpperClf) + ilSpots
                                    lmWklyRates(ilWeekInx, ilUpperClf) = tmCff.lActPrice

                                    'Accum previous modifications $ and spot count
                                    If imProcessFlag(ilClf) = 2 Or imProcessFlag(ilClf) = 3 And (tmClf.sType <> "E" And tmClf.sType <> "A" And tmClf.sType <> "O") Then     '10-7-08 only get the pop estimates for conventional& hidden lines ;'2=current line, but not on same rev, 3 = current line, same rev#,   '2=current line, but not on same rev, 3 = current line, same rev#
                                        llCurrModSpots = llCurrModSpots + ilSpots
                                        If tmClf.sType = "E" Then
                                            If ilSpots <> 0 Then        'insure that this is an airing week, dont just add in package amt
                                                llCurrMod = llCurrMod + tmCff.lActPrice
                                            End If
                                        Else
                                            llCurrMod = llCurrMod + (ilSpots * tmCff.lActPrice)
                                        End If
                                    End If

                                    If ilfirstTime Then
                                        If tgSpf.sDemoEstAllowed <> "Y" Then
                                            ilfirstTime = False
                                        End If
                                        If tmClf.iStartTime(0) = 1 And tmClf.iStartTime(1) = 0 Then
                                            llOvStartTime = 0
                                            llOvEndTime = 0
                                        Else
                                            'override times exist

                                            gUnpackTimeLong tmClf.iStartTime(0), tmClf.iStartTime(1), False, llOvStartTime
                                            gUnpackTimeLong tmClf.iEndTime(0), tmClf.iEndTime(1), True, llOvEndTime
                                        End If
                                        '11-19-02 Daily and weekly need the valid airing day, not the spots per day if daily (ilDemoAvgAudDays)
                                         ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, tmClf.iDnfCode, tmClf.iVefCode, tlBR.iSocEcoMnfCode, tgChf.iMnfDemo(ilDemo), llDate2, llDate2, tmClf.iRdfCode, llOvStartTime, llOvEndTime, ilDemoAvgAudDays(), tmClf.sType, tmClf.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                    End If
                                    lmAvgAud(ilWeekInx, ilUpperClf) = llAvgAud
                                    lmPopEst(ilWeekInx, ilUpperClf) = llPopEst
                                ElseIf ilWeekInx > 104 And (tlBR.iThisCntMod Or ilDiffOnly) Then
                                    imDiffExceeds104Wks = True      '3-9-06 set flag to show message in header
                                End If
                            Next llDate2
                        End If                                      'tgCff(ilCff).istatus = 0 or 1
                        ilCff = tgCff(ilCff).iNextCff               'get next flight record from mem
                    Loop                                            'while ilcff <> -1
                    'Set flag to send to Crystal for:  Multiple flights this line, or single flight this line.
                    'If multiple flights, the Research totals will go on separate total line.  If single flight,
                    'the research cpp/cpm will go on same line as spot data.
                    For ilLineSpotsInx = 1 To ilLineRate - 1 Step 1
                        If (tmLnr(ilLineSpotsInx).iLineInx = ilClf) Then            'same lines
                            tmLnr(ilLineSpotsInx).iManyFlts = ilManyFlts
                        End If
                    Next ilLineSpotsInx
                End If                                          'llFltStart >= gDateValue(slStartDate)
                ilUpperClf = ilUpperClf + 1
            Next ilClf                                          'for ilclf = lbound(tmclf) to ubound(tmclf)
        
            For ilClf = 0 To UBound(tmPodcast_Info) - 1
                'type of line:  P = Podcast, K = package line, H = Podcast Hidden Line, L = Other, not podcast Hidden Line, O = other not in pkg, not poscast (conventional, selling)
                If tmPodcast_Info(ilClf).sType = "P" Then
                    tmPodcast_Info(ilClf).bShowResearch = False
                ElseIf tmPodcast_Info(ilClf).sType = "O" Then
                    tmPodcast_Info(ilClf).bShowResearch = True
                ElseIf tmPodcast_Info(ilClf).sType = "H" Then       'hidden podcast line in pkg
                    tmPodcast_Info(ilClf).bShowResearch = False
                ElseIf tmPodcast_Info(ilClf).sType = "L" Then       'other type of vehicle in package, not a podcast vehicle
                    tmPodcast_Info(ilClf).bShowResearch = True
                Else
                    If tmPodcast_Info(ilClf).sType = "K" Then       'package, loop thru all the lines to find the references to see if all podcast, mixture
                        blPodExists = False
                        blNonPodExists = False
                        For ilTemp = 0 To UBound(tmPodcast_Info) - 1
                            If ilClf <> ilTemp Then                 'ignore itself
                                If tmPodcast_Info(ilClf).iLine = tmPodcast_Info(ilTemp).iPkgRefLine Then    'test for associated hidden lines to the package
                                    If tmPodcast_Info(ilTemp).sType = "H" Then           'hidden lineis podcast
                                        blPodExists = True
                                    End If
                                    If tmPodcast_Info(ilTemp).sType = "L" Then       'hidden line is not  a podcase=t
                                        blNonPodExists = True
                                    End If
                                End If
                            End If
                        Next ilTemp
                        'set how the package line should show for research info
                        If blPodExists And blNonPodExists Then
                            tmPodcast_Info(ilClf).bShowResearch = True          'mixture of podcast and non podcast, show research for package
                        ElseIf blPodExists And Not (blNonPodExists) Then        'all podcast hidden lines
                            tmPodcast_Info(ilClf).bShowResearch = False
                        Else
                            tmPodcast_Info(ilClf).bShowResearch = True          'no podcast, show research
                        End If
                        'now set how the hidden lines should show for research info
                        For ilTemp = 0 To UBound(tmPodcast_Info) - 1
                            If tmPodcast_Info(ilClf).iLine = tmPodcast_Info(ilTemp).iPkgRefLine Then        'matching hidden line
                                If tmPodcast_Info(ilTemp).sType = "H" Then       'podcast vehicle in the pkg
                                    tmPodcast_Info(ilTemp).bShowResearch = False
                                ElseIf tmPodcast_Info(ilTemp).sType = "L" Then
                                    tmPodcast_Info(ilTemp).bShowResearch = True
                                End If
                            End If
                        Next ilTemp
                    End If
                End If
            Next ilClf
            
            blPodExists = False
            blNonPodExists = False
            'determine if any podcast hidden lines within package, whether research header titles should be shown or not.
            'if all podcast within package, do not show research.  If mixed podcast and non-podcast, show them all so the research info
            'type of line:  P = Podcast, K = package line, H = Podcast Hidden Line, L = Other, not podcast Hidden Line, O = other, not poscast (conventional, selling)
            For ilClf = 0 To UBound(tmPodcast_Info) - 1
                If tmPodcast_Info(ilClf).sType = "P" Or tmPodcast_Info(ilClf).sType = "H" Then      'std line that is podcast, or hidden line that is podcast
                    blPodExists = True
                ElseIf tmPodcast_Info(ilClf).sType <> "K" Then          'ignore package lines to determine if showing column header
                    blNonPodExists = True
                End If
            Next ilClf
            If blPodExists And blNonPodExists Then
                tmCbf.sRschColHdr = ""          'mixture of podcast and non podcast, blnak indicates to show research hdr
            ElseIf blPodExists And Not (blNonPodExists) Then        'all podcast hidden lines, do not show research hdr
                tmCbf.sRschColHdr = "H"                                  'hide research hdr column
            Else
                tmCbf.sRschColHdr = ""           'no podcast, show research (blank indicates to show research hdr)
            End If
            
            'create CBF record from contract just gathered
            '11-23-99
            If tgSpf.sDemoEstAllowed <> "Y" Then
                For ilLoop = LBound(lmPop) To UBound(lmPop) - 1
                    If lmPop(ilLoop) < 0 Then       'neg was to indicate there was different
                                    'books/pop across vehicles.  Use 0 pop for TotalResearch
                        lmPop(ilLoop) = 0
                    End If
                Next ilLoop
            End If

            tmCbf.lGenTime = lgNowTime              '10-30-01
            tmCbf.iGenDate(0) = tlBR.iGenDate(0)
            tmCbf.iGenDate(1) = tlBR.iGenDate(1)
            tmCbf.lChfCode = tgChf.lCode                'contract internal code
            tmCbf.lCxfComment = 0               'comment code from Site for either BR or Insertion Orders

            If ilListIndex = CNT_INSERTION Then
                If tgSpf.lCxfInsertComment > 2 Then         '2-12-03 comments to show on all Insrtion order
                    tmCbf.lCxfComment = tgSpf.lCxfInsertComment
                End If
            Else
                tmCbf.sDailyExists = slDailyExists      '5-21-03  at least one day line exists flag (Y/N)
                If tgSpf.lCxfContrComment > 2 Then         '2-12-03 comments to show on all printed contrcts
                    tmCbf.lCxfComment = tgSpf.lCxfContrComment
                End If
            End If

            If ilTotalWks > 0 Then
                tmCbf.iTotalWks = ilRealTotalWks            'overall contract span in weeks
            End If
            If tlBR.iThisCntMod Then
                tmCbf.iTotalWks = ilCurrTotalWks        'differences only, use the Totl weeks determined from gGenDiff rtn from
                                                        'the current revision
            End If

            tmCbf.iExtra2Byte = 0                   'flag to include for Detail version only, ignore other types
            tmCbf.lExtra4Byte = 0                   'if corp summary totals, set flag to place legend on summart page
            tmCbf.iAffMktCode = -1                  '2-21-18 init mkt code for station market code reference (-1 = dont show mkt name, ignore; 0 = use vehicle group market name from vehicle; non zero = use affiliate mkt name
            tmCbf.lPop = 0                          '2-27-18 This is NOT population, field used to point to CLF to retrieve the Act1 Line Up code
            'Year and month is concatenated in the long field - Crystal uses to print out quarter headings on summary page
            'get start month of the contract start date
            slStr = Format$(llStdStartDates(1) + 15, "m/d/yy")
            gObtainYearMonthDayStr slStr, True, slYear, slMonth, slDay
            'xxxx = year, + x = start qtr + x = corp or std flag  (0 =std, 1= corp)
            ilTemp = Val(slMonth)
            If ilTemp < ilStartMonth Then
                ilTemp = ilTemp + 12
            End If

            'ilTemp-ilstartMonth/3+1 gets the starting qtr for the corp or std year based on the first qtr to be shown
            tmCbf.lExtra4Byte = (CLng(ilCorpStdYear) * 100) + (((ilTemp - ilStartMonth) / 3 + 1) * 10)
            If ilShowStdQtr = 2 Then            '12-19-20  for BR legend output
                tmCbf.lExtra4Byte = tmCbf.lExtra4Byte + 1
            End If
            tmCbf.iDnfCode = tgChf.iMnfDemo(ilDemo)   '6-2-16 need demo code, not the index      'ilDemo  demo version xx of possible 4

            'Setup % of budget for our station
            For ilLoop = 0 To 6 Step 1
                If ilCmpny = tgChf.iMnfCmpy(ilLoop) Then
                    tmCbf.iOurShare = tgChf.iCmpyPct(ilLoop)
                End If
            Next ilLoop
            'setup comment fields only if to be shown on contract.  Update record with
            'the comment pointer if it should be shown, else leave it 0
            '
            mGetComment tmCbf.lOtherComment, tgChf.lCxfCode, tlBR.iPropOrOrder
            mGetComment tmCbf.lCancComment, tgChf.lCxfCanc, tlBR.iPropOrOrder
            If tgChf.iExtRevNo > 0 Then                     'print cance, merch & promo clause on props & orders, not on revisions
                tmCbf.lMerchComment = 0
                tmCbf.lPromoComment = 0
                mGetComment tmCbf.lChgRComment, tgChf.lCxfChgR, tlBR.iPropOrOrder 'print mod reasons on all changes
            Else
                mGetComment tmCbf.lMerchComment, tgChf.lCxfMerch, tlBR.iPropOrOrder
                mGetComment tmCbf.lPromoComment, tgChf.lCxfProm, tlBR.iPropOrOrder
                tmCbf.lChgRComment = 0
            End If

            'put in the calculated taxes (if applicable) into the prepass record
            tmCbf.lTax1 = llTax1CalcAmt
            tmCbf.lTax2 = llTax2CalcAmt
            For ilLoop = 1 To 36
                tmCbf.lTax1 = tmCbf.lTax1 + llProjectedTax1(ilLoop)
                tmCbf.lTax2 = tmCbf.lTax2 + llProjectedTax2(ilLoop)
            Next ilLoop

            'Obtain all the books used for this contract, string out the demoresearch names
            'if there are multiple books, precede the description with an *.  Crystal report
            'will test for this flag and specify book name as "See Summary".  Then on the summary
            'page, the book name will be extracted from the line data and shown next to the line summary.
            '
            slSurvey = ""
            tmCbf.sSurvey = ""
            If tlBR.iShowResearch Then              'show Research data in header (survey book)
                For ilLoop = LBound(imDnfCodes) To ilUpperDnf - 1 Step 1
                    If imDnfCodes(ilLoop) > 0 Then
                        tmDnfSrchKey.iCode = imDnfCodes(ilLoop)
                        ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                        If ilRet <> BTRV_ERR_NONE Then
                            tmDnf.sBookName = " "
                        End If
                        If slSurvey = "" Then
                            slSurvey = RTrim$(tmDnf.sBookName)
                        Else
                            slSurvey = slSurvey & ", " & RTrim$(tmDnf.sBookName)
                        End If
                    End If
                Next ilLoop
                If ilUpperDnf > 1 Then
                    tmCbf.sSurvey = "*"     ' & slSurvey      'flag to denote multiple book names
                Else
                    tmCbf.sSurvey = slSurvey
                End If
            End If

            slSocEco = ""           '10-28-03 - show social economic category if selected
#If programmatic <> 1 Then
            If tlBR.iSocEcoMnfCode > 0 Then
                slNameCode = tgSocEcoCode(BrSnap!cbcSet1.ListIndex - 1).sKey
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilListIndex = Val(slCode)
                For ilLoop = LBound(tgSocEcoMnf) To UBound(tgSocEcoMnf)
                    If ilListIndex = tgSocEcoMnf(ilLoop).iCode Then
                        slSocEco = Trim$(tgSocEcoMnf(ilLoop).sName)
                        Exit For
                    End If
                Next ilLoop
            End If
#End If
            'Obtain all the demos used for this contract, string out the  names
            slDemos = " "
            ilFirstDemo = 0                 'assume gathering all 4 demos
            ilLastDemo = 3
            If tlBR.iAllDemos Then              'Use all demos (proposal option only)
                'Only show the demo it's gathering for rather than stringing out all of them (which is
                'what will be shown when only the primary is used)
                ilFirstDemo = ilDemo
                ilLastDemo = ilDemo
            End If
            
            If tgSaf(0).sHideDemoOnBR = "Y" And tgChf.sHideDemo = "Y" Then
                slDemos = "Impressions"
            Else
                For ilLoop = ilFirstDemo To ilLastDemo Step 1
                    If tgChf.iMnfDemo(ilLoop) > 0 Then
                        tmMnfSrchKey.iCode = tgChf.iMnfDemo(ilLoop)         'demo category
                        ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching demo name
                        If ilRet <> BTRV_ERR_NONE Then
                            tmMnf.sName = " "                        'error
                        End If
                        If slDemos = " " Then
                            slDemos = RTrim$(tmMnf.sName)
                        Else
                            slDemos = slDemos & ", " & RTrim$(tmMnf.sName)
                        End If
                        'show CPP or CPM data
                        If tgChf.lTarget(ilLoop) > 0 And tlBR.iShowProof Then        'show cpp or cpm on proof only
                        'If tgChf.lTarget(ilLoop) > 0 And RptSelCt!ckcSelC6(2).Value Then        'show cpp or cpm on proof only
                            If tgChf.sCppCpm = "P" Then         'CPP
                                slStr = Format$(tgChf.lTarget(ilLoop) / 100, "###")
                                slDemos = slDemos & " CPP@" & Trim$(slStr)
                            ElseIf tgChf.sCppCpm = "M" Then     'CPM
                                slStr = Format$(tgChf.lTarget(ilLoop) / 100, "###.00")
                                slDemos = slDemos & " CPM@" & Trim$(slStr)
                            End If
                        End If
                    End If
                Next ilLoop
            End If
            If slSocEco = "" Then       '10-29-03
                tmCbf.sDemos = slDemos
            Else
                tmCbf.sDemos = slSocEco & " " & slDemos
            End If
            
            If tlBR.iThisCntMod Then                            'differences option for this cnt (either printables & show all mods as diff,
                                                            'or requesting diff only on selective cnt)
                tmCbf.lCurrModSpots = llPrevSpots            'total spot count for the previous version - used for differences option
                tmCbf.lCurrMod = llPrevGross                 'total $ for the previous version - used for differences option
            Else
                tmCbf.lCurrModSpots = llCurrModSpots            'total spot count for the current version - used for differences option
                tmCbf.lCurrMod = llCurrMod                      'total $ for the current version - used for differences option
            End If
            tmCbf.iTotalMonths = ilCurrTotalMonths          'total airing months of this order
            tmCbf.iStartQtr(0) = ilCurrStartQtr(0)              'start month of current qtr
            tmCbf.iStartQtr(1) = ilCurrStartQtr(1)
            
            'Calculate total number of actual airing weeks
            'first go thru and find unique airing weeks, then go accum those
            'unique airing weeks
            'i.e.  The contract may have a 52 week start & end date span, but only air
            'every other week.  The report will show 26/52 (26 airing weeks across 52 week span)
            Erase imAirWks                                      'initialize true airing weeks for this cnt

            For ilLoop = 1 To ilLineRate - 1 Step 1             'loop for # of lines gathered, there may be more than
                                                                'one entry for a line due to different rates in the flights for a line
                'first decide which lines gathered shouldbe included to calc the actual number of weeks airing.  If modification option,
                'all history lines have been gathered
                ilLoop3 = tmLnr(ilLoop).iLineInx
                If imProcessFlag(ilLoop3) = 2 Or imProcessFlag(ilLoop3) = 3 Then      'flags 2 = current line, dont show line but process for flights,
                                                                    'flag 3 = current line, show on Br and process for flight
                    For ilDay = 1 To MAXWEEKSFOR2YRS Step 1                     'loop for the line entry for 2 yrs (104 wks)
                        If lmSpotsByWk(ilDay, ilLoop) <> 0 Then
                            imAirWks(ilDay) = 1
                        End If
                    Next ilDay
                '10-20-03
                ElseIf imProcessFlag(ilLoop3) = 1 Then          'from previous revsion, gather $ and spot count
                    For ilDay = 1 To MAXWEEKSFOR2YRS Step 1                     'loop for the line entry for 2 yrs (104 wks)
                        If lmSpotsByWk(ilDay, ilLoop) <> 0 Then          'valid airing week

                        End If
                    Next ilDay
                End If
            Next ilLoop                                         'go to next line entry
            tmCbf.iAirWks = 0
            For ilLoop = 1 To MAXWEEKSFOR2YRS
                If imAirWks(ilLoop) > 0 Then
                    tmCbf.iAirWks = tmCbf.iAirWks + 1
                End If
            Next ilLoop
            If tlBR.iThisCntMod Then
            'If ilThisCntMod Then
                tmCbf.iAirWks = ilCurrAirWks        'differences only, use the airing weeks determined from gGenDiff rtn from
                                                    'the current revision
            End If
            'Cycle through all the line records built (some lines may have multiples records due to
            'different flights for the same line.  Calculate total GRIMPs to place in each record so
            'that the % Distribution can be calculated in Crystal.
            'Need to do this in a single prepass because the overall Grimps need to be stored in each qtr's recd
            'Get Research data for individual lines by unique rates
            llCntGrimps = 0
            For ilLoop = 1 To ilLineSpots - 1
                If imProcessFlag(tmLnr(ilLoop).iLineInx) = 1 Or imProcessFlag(tmLnr(ilLoop).iLineInx) = 3 Or imProcessFlag(tmLnr(ilLoop).iLineInx) = 4 Then  'show current and prev lines depending
                    ilLoop3 = tmLnr(ilLoop).iLineInx            'line index to process
                    tmClf = tgClf(ilLoop3).ClfRec
                    tmLnr(ilLoop).lLRchInx = UBound(tmLRch)     'index to start of this lines 8 quarters - When differences
                                                                'report is run, lines are built in LNR and not shown on BR,
                                                                'so this is index to where the Research data exists for this line
                    ilSavePkVeh = 0                             'init in case its not a hidden line
                    If tmClf.sType = "H" Then                   'hidden line, find assoc. pkg vehicle
                        For ilPkg = 0 To UBound(tgClf) - 1 Step 1
                            If tmClf.iPkLineNo = tgClf(ilPkg).ClfRec.iLine Then    'find the assoc pkg vehicle name
                                ilSavePkVeh = tgClf(ilPkg).ClfRec.iVefCode
                                Exit For
                            End If
                        Next ilPkg
                    ElseIf tmClf.sType = "E" Or tmClf.sType = "O" Or tmClf.sType = "A" Then     'package
                        'update the package vehicles population from all the hidden lines
                        'first, find the package vehicles index to store the pop in array
                        For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1 Step 1
                            If tmClf.iVefCode = ilVehList(ilVehicle) Then
                                ilfirstTime = True              '3-1-11 need to get the package vehicles population.  if hidden lines have varying population, use 0; if all same use that pop
                                                                'ignore the CBS lines population
                                'ilvehicle contains index to store package population
                                '5-28-10 do not set the vehicle pop until a matching package vehicle is found so
                                'that incase of CBS, it doesnt look like theres varying populations (by using 0 in the calculations)
                                'lmPopPkg(ilVehicle) = -1    '11-24-04 pop for the pkg because of using same hidden vehicle reference as the pkg vehicle
                                For ilPkg = 0 To UBound(tgClf) - 1 Step 1
                                    gUnpackDate tgClf(ilPkg).ClfRec.iStartDate(0), tgClf(ilPkg).ClfRec.iStartDate(1), slStr
                                    llFltStart = gDateValue(slStr)
                                    gUnpackDate tgClf(ilPkg).ClfRec.iEndDate(0), tgClf(ilPkg).ClfRec.iEndDate(1), slStr
                                    llFltEnd = gDateValue(slStr)
                                    'find all the hidden lines associated with the package to see if their populations vary across them
                                    If tmClf.iLine = tgClf(ilPkg).ClfRec.iPkLineNo And llFltEnd >= llFltStart Then    'find the assoc pkg vehicle name
                                        '5-28-10 moved from above to avoid using wrong population of 0 when CBS involved
                                        If ilfirstTime Then
                                            lmPopPkg(ilVehicle) = -1    '11-24-04 pop for the pkg because of using same hidden vehicle reference as the pkg vehicle
                                            ilfirstTime = False
                                        End If
                                        For ilQtr = LBound(ilVehList) To UBound(ilVehList)
                                            If tgClf(ilPkg).ClfRec.iVefCode = ilVehList(ilQtr) Then         'get the index to hidden vehicle to to index into the line population table (lmpop)
                                                Exit For
                                            End If
                                        Next ilQtr
                                        '3-16-12 Always on the first time, use the first population it finds
                                        If lmPopPkg(ilVehicle) = -1 Then    'And lmPop(ilQtr) <> 0 Then          'first time
                                            If lmPop(ilQtr) > 0 Then                        '3-13-19
                                                lmPopPkg(ilVehicle) = lmPop(ilQtr)
                                            End If
                                        Else
                                            If (lmPopPkg(ilVehicle) <> 0) And (lmPopPkg(ilVehicle) <> lmPop(ilQtr)) And (lmPop(ilQtr) <> 0) Then      'test to see if this pop is different that the prev one.
                                                 lmPopPkg(ilVehicle) = 0                                           'if different pops, calculate the contract  summary different
                                            Else
                                            'if current line has population, but there was already a different across
                                            'lines in pop, dont save new one
                                                If lmPop(ilQtr) <> 0 And (lmPopPkg(ilVehicle) <> 0 And lmPopPkg(ilVehicle) <> -1) Then   '2/1/99
                                                    lmPopPkg(ilVehicle) = lmPop(ilQtr)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next ilPkg

                                Exit For
                            End If
                        Next ilVehicle
                    End If
                    
                    'find the vehicle and associated population
                    llPop = 0
                    For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1 Step 1
                        If tmClf.iVefCode = ilVehList(ilVehicle) Then
                            'llPop = lmPop(ilVehicle)
                            llPop = llPopByLine(ilLoop3 + 1)      '11-23-99
                            Exit For
                        End If
                    Next ilVehicle
                    
                    '4-10-19 replace with Long due to subscript out of range
                    'replace ilQtr with llLoopOnQtr
                    'replace ilTemp with llTempLRch
                    'replace ilRchQtr with llRchQtr
                    llRchQtr = 0                            '11-14-11
                    
                    For llLoopOnQtr = 1 To 8 Step 1         'retain 8 quarters to fall thru and get
                        llTempLRch = UBound(tmLRch)
                        tmLRch(llTempLRch).lQSpots = 0
                        For ilSpots = 1 To imWeeksPerQtr(llLoopOnQtr)         '11-14-11
                            tmLRch(llTempLRch).lSpots(ilSpots - 1) = lmSpotsByWk(llRchQtr + ilSpots, ilLoop)
                            tmLRch(llTempLRch).lQSpots = tmLRch(llTempLRch).lQSpots + tmLRch(llTempLRch).lSpots(ilSpots - 1)
                            tmLRch(llTempLRch).lRates(ilSpots - 1) = lmWklyRates(llRchQtr + ilSpots, ilLoop3)
                            tmLRch(llTempLRch).lAvgAud(ilSpots - 1) = lmAvgAud(llRchQtr + ilSpots, ilLoop3)
                            tmLRch(llTempLRch).lPopEst(ilSpots - 1) = lmPopEst(llRchQtr + ilSpots, ilLoop3)
                        Next ilSpots
                        mBuildLRch tmLRch(), llTempLRch, ilSavePkVeh, llCntGrimps, llPop            '4-10-19 remove ilLoopOnQtr parameter
                        llRchQtr = llRchQtr + imWeeksPerQtr(llLoopOnQtr)
                    Next llLoopOnQtr
                End If
            Next ilLoop

            '6-1-04 Need to get average pop estimates for each line for all the quarterly calc if Demo Estimates allowed
            llOverallPopEst = -1
            If tgSpf.sDemoEstAllowed = "Y" Then
                For ilLoop = 1 To ilLineSpots - 1
                    If imProcessFlag(tmLnr(ilLoop).iLineInx) = 1 Or imProcessFlag(tmLnr(ilLoop).iLineInx) = 3 Or imProcessFlag(tmLnr(ilLoop).iLineInx) = 4 Then  'show current and prev lines depending
                        ilLoop3 = tmLnr(ilLoop).iLineInx            'line index to process
                                                                    'on the option selected.  Mods show current & prev, Full BR shows curent only
                                                                    'Move spots, $ and aud to single dimension array for gAvgAudToLnResearch routine
                        tmClf = tgClf(ilLoop3).ClfRec
                        If tmClf.sType <> "E" And tmClf.sType <> "A" And tmClf.sType <> "O" Then    'only get the pop estimates for conventional& hidden lines

                            For ilSpots = 1 To MAXWEEKSFOR2YRS
                                llWklySpots(ilSpots - 1) = lmSpotsByWk(ilSpots, ilLoop)
                                llWklyRates(ilSpots - 1) = lmWklyRates(ilSpots, ilLoop3)
                                llWklyAvgAud(ilSpots - 1) = lmAvgAud(ilSpots, ilLoop3)
                                llWklyPopEst(ilSpots - 1) = lmPopEst(ilSpots, ilLoop3)
                            Next ilSpots
                            'gAvgAudToLnResearch sm1or2PlaceRating, True, llPop, llWklyPopEst(), llWklySpots(), llWklyRates(), llWklyAvgAud(), llTemp, tmCbf.lAvgAud, ilWklyRtg(), tmCbf.iAvgRate, llWklyGrimp(), tmCbf.lGrImp, llWklyGRP(), tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, llPopEst
                            gAvgAudToLnResearch sm1or2PlaceRating, True, llPop, llWklyPopEst(), llWklySpots(), llWklyRates(), llWklyAvgAud(), dlTemp, tmCbf.lAvgAud, ilWklyRtg(), tmCbf.iAvgRate, llWklyGrimp(), tmCbf.lGrImp, llWklyGRP(), tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, llPopEst 'TTP 10439 - Rerate 21,000,000

                            If llOverallPopEst = -1 Then      '6-1-04
                                llOverallPopEst = llPopEst
                            Else
                                If llOverallPopEst <> llPopEst And llOverallPopEst <> 0 Then
                                    llOverallPopEst = 0
                                End If
                            End If
                        End If                      'tmclf.stype
                    End If                          'imProcessFlag(tmLnr(ilLoop).iLineInx) = 1
                Next ilLoop
            End If

            mGetPopPkgByLine ilPkgLineList(), llPopByLine()      '11-11-05 update the pkg vehicles by line # (reqd if more than 1 line for the same package vehicle)

            'For Packages Only........
            'Create array of grps, weekly rating, grimps and rates to pass to routine to generate qtrly totals for individual
            'line packages.  (combination of all hidden lines for 1 package line)
            mBRPkgQtrTotals ilPkgLineList(), ilVehList(), ilPkgVehList()
            'Create array of grps, wkly rating, grimps and rates to pass to routine to generate qtrly totals for all vehicles (detail version)
            'Store the qtr totals in all detail records - it's the only way to get the accurate figures
            ReDim tmQGRP(0 To 0) As Long
            ReDim tmQCPP(0 To 0) As Long
            ReDim tmQCPM(0 To 0) As Long
            ReDim tmQGrimp(0 To 0) As Long
            llPop = -1                          '6-1-04
            For llLoopOnQtr = 1 To imMaxQtrs Step 1
                ReDim tmVRtg(0 To 0) As Integer
                ReDim tmVCost(0 To 0) As Long
                ReDim tmVGRP(0 To 0) As Long
                ReDim tmVGrimp(0 To 0) As Long
                '3/13/04- moved llLnSpots = 0 here
                llLnSpots = 0
                If tgSpf.sDemoEstAllowed = "Y" Then     '6-1-04
                    llResearchPop = -1
                End If
                For llRchQtr = llLoopOnQtr To UBound(tmLRch) - 1 Step 8
                    'bypass the Package lines, the quarterly totals will be obtained from the individual hidden lines because
                    'each package line can contain the same data repeated in each flight.
                    'Exclude pkg to avoid duplicate results (the hidden lines are calc. separately)
                    '12/17/15: initialized above, line was not removed
                    'llLnSpots = 0           '11-14-11
                    If (tmLRch(llRchQtr).sType <> "A" And tmLRch(llRchQtr).sType <> "O" And tmLRch(llRchQtr).sType <> "E") Then
                        If (tmLRch(llRchQtr).lQSpots <> 0) Then 'And (tmLRch(llRchQtr).lTotalGrimps <> 0) Then            'only process if spots exist in the qtr
                            llLnSpots = llLnSpots + tmLRch(llRchQtr).lQSpots
                            If llResearchPop = -1 Then      '6-1-04
                                llResearchPop = tmLRch(llRchQtr).lSatelliteEst
                            Else
                                '9-27-04 ignore pop if its 0, which means no book was designated or not found
                                If (llResearchPop <> tmLRch(llRchQtr).lSatelliteEst And tmLRch(llRchQtr).lSatelliteEst <> 0) And llResearchPop <> 0 Then
                                    llResearchPop = 0
                                End If
                            End If
                            tmVRtg(UBound(tmVRtg)) = tmLRch(llRchQtr).iTotalAvgRating
                            tmVCost(UBound(tmVCost)) = tmLRch(llRchQtr).dTotalCost 'TTP 10439 - Rerate 21,000,000
                            tmVGRP(UBound(tmVGRP)) = tmLRch(llRchQtr).lTotalGRP
                            tmVGrimp(UBound(tmVGrimp)) = tmLRch(llRchQtr).lTotalGrimps
                            ReDim Preserve tmVRtg(0 To UBound(tmVRtg) + 1) As Integer
                            ReDim Preserve tmVCost(0 To UBound(tmVCost) + 1) As Long
                            ReDim Preserve tmVGRP(0 To UBound(tmVGRP) + 1) As Long
                            ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) + 1) As Long
                        End If
                    End If
                Next llRchQtr
                If UBound(tmVRtg) > 0 Then
                    'dimensions must be exact sizes
                    ReDim Preserve tmVRtg(0 To UBound(tmVRtg) - 1) As Integer
                    ReDim Preserve tmVCost(0 To UBound(tmVCost) - 1) As Long
                    ReDim Preserve tmVGRP(0 To UBound(tmVGRP) - 1) As Long
                    ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) - 1) As Long
                    gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, dlTemp, tmCbf.iAvgRate, tmQGrimp(UBound(tmQGrimp)), tmQGRP(UBound(tmQGRP)), tmQCPP(UBound(tmQCPP)), tmQCPM(UBound(tmQCPM)), tmCbf.lAvgAud 'TTP 10439 - Rerate 21,000,000
                End If
                ReDim Preserve tmQGRP(0 To UBound(tmQGRP) + 1) As Long
                ReDim Preserve tmQCPP(0 To UBound(tmQCPP) + 1) As Long
                ReDim Preserve tmQCPM(0 To UBound(tmQCPM) + 1) As Long
                ReDim Preserve tmQGrimp(0 To UBound(tmQGrimp) + 1) As Long
            Next llLoopOnQtr
            '8 Quarter research totals are stored in tmVRtg, tmVCost, tmVGRP and tmVGrimp to store in each detail record
            ReDim tmWkCntGrps(0 To 0) As WEEKLYGRPS
            If tgSpf.sDemoEstAllowed = "Y" Then     '6-1-04
                mGetCntWkGrps llOverallPopEst
            Else
                mGetCntWkGrps llResearchPop   '2-10-00
            End If
            'Create array of grps, wkly rating, grimps and rates to pass to routine to generate qtrly unique vehicle totals  (detail version)
            'Store the qtr totals in all detail records - it's the only way to get the accurate figures
            'ReDim tmVehQtrList(1 To 1) As VEHQTRLIST
            ReDim tmVehQtrList(0 To 0) As VEHQTRLIST
            '4-10-19 replace ilQtr with llLoopOnQtr
            For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1 Step 1
                llPop = lmPop(ilVehicle)           '11-23'99??  associated vehicles pop
                For llLoopOnQtr = 1 To imMaxQtrs Step 1
                    ReDim tmVRtg(0 To 0) As Integer
                    ReDim tmVCost(0 To 0) As Long
                    ReDim tmVGRP(0 To 0) As Long
                    ReDim tmVGrimp(0 To 0) As Long

                    'get each vehicle for a quarter at a time (detail option)
                    '3/13/04- moved llLnSpots = 0 here
                    llLnSpots = 0
                    If tgSpf.sDemoEstAllowed = "Y" Then     '6-1-04
                        llResearchPop = -1
                    End If
                    For llRchQtr = llLoopOnQtr To UBound(tmLRch) - 1 Step 8
                        '10-20-03
                        If tmLRch(llRchQtr).sType <> "H" Then    '6-6-00 not a package of any kind
                            If tmLRch(llRchQtr).lQSpots <> 0 And tmLRch(llRchQtr).iVefCode = ilVehList(ilVehicle) Then
                                llLnSpots = llLnSpots + tmLRch(llRchQtr).lQSpots
                                If llPop = -1 Then      '6-1-04
                                    llPop = tmLRch(llRchQtr).lSatelliteEst
                                Else
                                    If llPop <> tmLRch(llRchQtr).lSatelliteEst And llPop <> 0 Then
                                        llPop = 0
                                    End If
                                End If
                                tmVRtg(UBound(tmVRtg)) = tmLRch(llRchQtr).iTotalAvgRating
                                tmVCost(UBound(tmVCost)) = tmLRch(llRchQtr).dTotalCost 'TTP 10439 - Rerate 21,000,000
                                tmVGRP(UBound(tmVGRP)) = tmLRch(llRchQtr).lTotalGRP
                                tmVGrimp(UBound(tmVGrimp)) = tmLRch(llRchQtr).lTotalGrimps
                                ReDim Preserve tmVRtg(0 To UBound(tmVRtg) + 1) As Integer
                                ReDim Preserve tmVCost(0 To UBound(tmVCost) + 1) As Long
                                ReDim Preserve tmVGRP(0 To UBound(tmVGRP) + 1) As Long
                                ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) + 1) As Long
                            End If
                        End If
                    Next llRchQtr

                    'If UBound(tmVRtg) > 1 Then
                    If UBound(tmVRtg) > 0 Then
                        'dimensions must be exact sizes
                        ReDim Preserve tmVRtg(0 To UBound(tmVRtg) - 1) As Integer
                        ReDim Preserve tmVCost(0 To UBound(tmVCost) - 1) As Long
                        ReDim Preserve tmVGRP(0 To UBound(tmVGRP) - 1) As Long
                        ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) - 1) As Long
                        gResearchTotals sm1or2PlaceRating, True, llPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, dlTemp, tmCbf.iAvgRate, tmCbf.lVQGrimp, tmCbf.lVQGRP, tmCbf.lVQCPP, tmCbf.lVQCPM, tmCbf.lAvgAud 'TTP 10439 - Rerate 21,000,000
                        tmVehQtrList(ilVehicle).lVQGrimp(llLoopOnQtr - 1) = tmCbf.lVQGrimp
                        tmVehQtrList(ilVehicle).lVQGRP(llLoopOnQtr - 1) = tmCbf.lVQGRP
                        tmVehQtrList(ilVehicle).lVQCPP(llLoopOnQtr - 1) = tmCbf.lVQCPP
                        tmVehQtrList(ilVehicle).lVQCPM(llLoopOnQtr - 1) = tmCbf.lVQCPM
                    End If
                Next llLoopOnQtr
                ReDim Preserve tmVehQtrList(0 To UBound(tmVehQtrList) + 1) As VEHQTRLIST
            Next ilVehicle
            '8 Quarter research totals are stored in tmVRtg, tmVCost, tmVGRP and tmVGrimp to store in each detail record

            If tgSpf.sDemoEstAllowed = "Y" Then
                mGetCntWkVGrps llOverallPopEst, ilVehList()  '2-11-00
                mGetHiddenWkVGrps llOverallPopEst, ilVehList(), ilPkgLineList() '1-24-03
            Else
                mGetCntWkVGrps llResearchPop, ilVehList()  '2-11-00
                mGetHiddenWkVGrps llResearchPop, ilVehList(), ilPkgLineList() '1-24-03
            End If
            'Obtain package totals by line
            '2-20-00
            mBRPkgVehTotals ilPkgLineList(), ilVehList(), ilPkgVehList()

            ReDim ilPkgVehicleForSurveyList(0 To 0) As Integer
            'Create the CBF quarterly records for as many lines as there are
            'improcessflag = 4 are the hidden lines to packages.  Show them for proof.
            For ilLoop = 1 To ilLineSpots - 1
                If imProcessFlag(tmLnr(ilLoop).iLineInx) = 1 Or imProcessFlag(tmLnr(ilLoop).iLineInx) = 3 Or imProcessFlag(tmLnr(ilLoop).iLineInx) = 4 Then  'show current and prev lines depending
                    ilLoop3 = tmLnr(ilLoop).iLineInx            'line index to process
                    tmClf = tgClf(ilLoop3).ClfRec
                    'TTP 8410
                    'If (tlBR.iShowProof) Or (Not tlBR.iShowProof And tmClf.sType <> "H") Then
                    If (tlBR.iShowProof) Or (Not tlBR.iShowProof And tmClf.sType <> "H") Or tgChf.sInstallDefined = "Y" Then
                        'on the option selected.  Mods show current & prev, Full BR shows curent only
                        'Loop thru the array containing type of schedule lines to determine if podcast or not.
                        'if podcast vehicle, do not show avg rtg, cpp or grp column.  Show a hidden line podcast reserach
                        'info only if a mixture of podcast and non podcast vehicle
                        For ilTemp = 0 To UBound(tmPodcast_Info) - 1
                            If tmPodcast_Info(ilTemp).iLine = tmClf.iLine Then
                                If tmPodcast_Info(ilTemp).bShowResearch Then
                                    tmCbf.sMixTypes = ""                'show the research info
                                Else
                                    tmCbf.sMixTypes = "H"               'hide the research info; its podcast
                                End If
                            End If
                        Next ilTemp
                        
                        ilGot1ToPrint = True
                        
                        If Not tlBR.iThisCntMod Then
                            tmCbf.iPctDist = 0              'Pct dist field replaced by Difference flag (0=full)
                        Else
                            tmCbf.iPctDist = 1              '1 = difference only (used to print legend on BR for diff. option)
                        End If
                        tmCbf.sLineSurvey = ""

                        'if package line, gather all the unique book names from the hidden lines that make up pkg
                        If tmClf.sType = "E" Or tmClf.sType = "A" Or tmClf.sType = "O" Then 'for packages, show all the books that make up the package
                            '12-19-13 Get the book references for this package line.  While processing each demo reference for this package, flag it with a -1 so not to look at it again
                            'Create a txr for each unique package vehicle to show on summary
                            slSurvey = ""
                            For ilPkgOrSpot = LBound(tlPkgVehicleForSurveyList) To UBound(tlPkgVehicleForSurveyList) - 1
                                If tlPkgVehicleForSurveyList(ilPkgOrSpot).iPkgVefCode = tmClf.iVefCode Then 'got matching package vehicle; determine the books for all the hidden lines
                                    tmDnfSrchKey.iCode = tlPkgVehicleForSurveyList(ilPkgOrSpot).iHiddenDnfCode
                                    ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If ilRet <> BTRV_ERR_NONE Then
                                        tmDnf.sBookName = " "
                                    End If
                                    If Trim$(slSurvey) = "" Then            'first time for this package
                                        slSurvey = RTrim$(tmDnf.sBookName)
                                    Else
                                        slSurvey = Trim$(slSurvey) & ", " & RTrim$(tmDnf.sBookName)
                                    End If
                                    tlPkgVehicleForSurveyList(ilPkgOrSpot).iPkgVefCode = -1         'flag processed
                                End If
                            Next ilPkgOrSpot
                            
                            If Trim$(slSurvey) <> "" Then
                                tmTxr.lGenTime = lgNowTime
                                tmTxr.iGenDate(0) = igNowDate(0)
                                tmTxr.iGenDate(1) = igNowDate(1)
                                tmTxr.lCsfCode = tgChf.lCode                'contract internal code
                                tmTxr.lSeqNo = -1                                 'distinguish this from the Split Network data that may be created into TXR
                                tmTxr.iGeneric1 = tgChf.iMnfDemo(ilDemo)      '6-2-16 all demo option
                                tmTxr.iType = tmClf.iVefCode
                                tmTxr.sText = Trim$(slSurvey)
                                ilRet = btrInsert(hmTxr, tmTxr, Len(tmTxr), INDEXKEY0)
                            End If
                        Else
                            tmDnfSrchKey.iCode = tmClf.iDnfCode
                            ilRet = btrGetEqual(hmDnf, tmDnf, imDnfRecLen, tmDnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching advt recd
                            If ilRet = BTRV_ERR_NONE Then
                                tmCbf.sLineSurvey = tmDnf.sBookName
                            Else                        '8-17-19 none found, init output field
                                tmCbf.sLineSurvey = ""
                            End If
                        End If
                        mGetComment tmCbf.lLineComment, tmClf.lCxfCode, tlBR.iPropOrOrder       'line comment
                        tmCbf.lIntComment = 0               'in place of internal comment code, use for hidden line reference to pkg line
                        tmCbf.sLineType = tmClf.sType       'in reports - if type "E", then the rate is the weekly rate, not spot rate
                        If tmClf.sType = "H" Then   'hidden line, reference the line # for proof
                            tmCbf.lIntComment = tmClf.iPkLineNo
                        ElseIf tmClf.sType = "A" Or tmClf.sType = "O" Or tmClf.sType = "E" Then
                            tmCbf.lIntComment = -2                  'let crystal know its a package line. Show research totals on same line as spots
                            If (tmLnr(ilLoop).iManyFlts) Then
                                tmCbf.lIntComment = -1              'multiple flights for same sche line, show package research totals on "total line"
                            End If
                        End If
                        If ilDiffOnly Then                      'grimps to calculate % of distibution in Crystal
                            tmCbf.lCntGrimps = llCntGrimps - llPrevCntGrimps
                        Else
                            tmCbf.lCntGrimps = llCntGrimps          'cnt total grimps so % distribution can be obtained in Crystal
                        End If
                        slStr = slStartDate                     'don't destroy start date of this order
                        gUnpackDate tmClf.iStartDate(0), tmClf.iStartDate(1), slNameCode
                        gUnpackDate tmClf.iEndDate(0), tmClf.iEndDate(1), slCode
                        'test for cancel before start which is always start date one day later that end date
                        'This is code to fix up schedule lines that had bad start & end dates in them.  They didnt coincide with the flights or header
                        If ((gDateValue(slNameCode) - gDateValue(slCode)) <> 1) Then      'not cancel before start
                            If gDateValue(slNameCode) < llChfStart Or gDateValue(slNameCode) > llChfEnd Then
                                slNameCode = Format$(llChfStart, "m/d/yy")
                            End If
                            If gDateValue(slCode) < llChfStart Or gDateValue(slCode) > llChfEnd Then
                                slCode = Format$(llChfEnd, "m/d/yy")
                            End If
                        End If

                        '7-8-05 determine if using solo avails, 1st position and preferred days & times
                        tmCbf.iOBBLen = 0
                        tmCbf.iCBBLen = 0
                        tmCbf.s1stPosition = "N"
                        tmCbf.sSoloAvail = "N"
                        tmCbf.sPrefDT = ""
                        If (Asc(tmClf.sOV2DefinedBits) And &H4) = &H4 Then  '1St pos.
                            If tmClf.iPosition > 0 Then
                                tmCbf.s1stPosition = "Y"
                            End If
                        End If
                        If tmClf.sSoloAvail = "Y" Then          'solo avails
                            tmCbf.sSoloAvail = "Y"
                        End If
                        
                        '1-10-14 Audio type from line override
                        If ((Asc(tgSaf(0).sFeatures1) And SHOWAUDIOTYPEONBR) = SHOWAUDIOTYPEONBR) Then      'yes, show audio type on lines
                            tmCbf.sAudioType = tmClf.sLiveCopy
                        Else
                           tmCbf.sAudioType = ""
                        End If

                        'Handle the Cancel Before Start case- where contract end date is prior to contract start date or the dates
                        'are all zero
                        tmCbf.sWeeksInQtr = "3"   '11-14-11 default to 13 week qtr
                        If gDateValue(slEndDate) < gDateValue(slStartDate) Or (gDateValue(slStartDate) = 0 And gDateValue(slEndDate) = 0) Then             'test for cancel before start
                            gPackDate slStr, tmCbf.iDtFrstBkt(0), tmCbf.iDtFrstBkt(1)
                            'ilLoop3 = tmLnr(ilLoop).iLineInx
                            tmCbf.iTotalWks = 0
                            If tmClf.sHideCBS <> "Y" Then
                                mShowCBS ilListIndex, tlSofList(), slDailyExists       '2-24-06, format the CBS line
                                '2-24-06 If this line is the last line to print for the quarter, the quarterly
                                'research numbers need to be updated because crystal will put print these fields
                                 'setup quarterly  Research values for all vehicles
                                tmCbf.lQGRP = tmQGRP(0)
                                tmCbf.lQCPP = tmQCPP(0)
                                tmCbf.lQCPM = tmQCPM(0)
                                tmCbf.lQGrimp = tmQGrimp(0)
                                mPutCntWkGrps 1                 'update the quarters weekly grps
                                tmCbf.sSnapshot = tlBR.sSnapshot

                                igBR_SSLinesExist = True        '12-16-03 flag at least 1 sch line exists to print either detail or summary version
                                If imWeeksPerQtr(1) = 14 Then       '9-28-17 handle CBS and qtr not 13 weeks
                                    tmCbf.sWeeksInQtr = "4"
                                ElseIf imWeeksPerQtr(1) = 12 Then
                                    tmCbf.sWeeksInQtr = "2"
                                End If
                                ilVefInxForCallLetters = gBinarySearchVef(tmCbf.iVefCode)
                                If ilVefInxForCallLetters <> -1 Then        '10-23-06 if new package vehicle entered for proposal, the index wont exist.  The vehicle
                                    gFindStationMkt ilVefInxForCallLetters, tmCbf            '2-21-18 show market name?
                                End If
                                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                            End If
                        ElseIf gDateValue(slCode) < gDateValue(slNameCode) Then  'test cancel before start on line dates
                            slTempDays = slStr      'cancel before start line, put out at least 1 qtr to show CBS
                            gPackDate slTempDays, tmCbf.iDtFrstBkt(0), tmCbf.iDtFrstBkt(1)
                            gPackDate slTempDays, tmCbf.iDtFrstBkt(0), tmCbf.iDtFrstBkt(1)
                            If tmCbf.iAirWks = 0 Then
                                tmCbf.iTotalWks = 0
                            End If
                            If tmClf.sHideCBS <> "Y" Then
                                mShowCBS ilListIndex, tlSofList(), slDailyExists       '2-24-06, format the CBS line
                                '2-24-06 If this line is the last line to print for the quarter, the quarterly
                                'research numbers need to be updated because crystal will put print these fields
                                 'setup quarterly  Research values for all vehicles
                                tmCbf.lQGRP = tmQGRP(0)
                                tmCbf.lQCPP = tmQCPP(0)
                                tmCbf.lQCPM = tmQCPM(0)
                                tmCbf.lQGrimp = tmQGrimp(0)
                                
                                '4-29-10 place the vehicle qtr totals in each record for CBS line
                                'so that totals will be shown properly.  Depending on where the CBS
                                'line is sorted for output, it may get incorrect values if not udpated
                                For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1 Step 1
                                    If tmClf.iVefCode = ilVehList(ilVehicle) Then
                                        Exit For
                                    End If
                                Next ilVehicle
                                ''setup quarterly  Research values by vehicles
                                tmCbf.lVQGRP = tmVehQtrList(ilVehicle).lVQGRP(0)
                                tmCbf.lVQCPP = tmVehQtrList(ilVehicle).lVQCPP(0)
                                tmCbf.lVQCPM = tmVehQtrList(ilVehicle).lVQCPM(0)
                                tmCbf.lVQGrimp = tmVehQtrList(ilVehicle).lVQGrimp(0)
                                
                                mPutCntWkGrps 1             'update the quarters weekly grps
                                tmCbf.sSnapshot = tlBR.sSnapshot

                                igBR_SSLinesExist = True        '12-16-03 flag at least 1 sch line exists to print either detail or summary version
                                If imWeeksPerQtr(1) = 14 Then       '9-28-17 handle CBS and qtr not 13 weeks
                                    tmCbf.sWeeksInQtr = "4"
                                ElseIf imWeeksPerQtr(1) = 12 Then
                                    tmCbf.sWeeksInQtr = "2"
                                End If
                                ilVefInxForCallLetters = gBinarySearchVef(tmCbf.iVefCode)
                                If ilVefInxForCallLetters <> -1 Then        '10-23-06 if new package vehicle entered for proposal, the index wont exist.  The vehicle
                                    gFindStationMkt ilVefInxForCallLetters, tmCbf            '2-21-18 show market name?
                                End If
                                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                            End If
                        Else
                            '4-10-19 replace ilRch with llRchQtr due to subscript out of range
                            llRchQtr = tmLnr(ilLoop).lLRchInx          'get starting index to research data for this line/rate event
                            ilVaryingQtrs = 0                       '11-14-11
                            For ilQtr = 1 To imMaxQtrs Step 1              'create 8 qtrs max
                                llAccumSpotCount = 0                           'gather # spots in qtr
                                dlTemp = 0                          'total $ in qtr'TTP 10439 - Rerate 21,000,000
                                ilTemp = 0                          '3/4/99 flag to detect mod with net result of spots is zero and $ is n/c
                                For ilWeekInx = 1 To imWeeksPerQtr(ilQtr)       '11-14-11
                                    llAccumSpotCount = lmSpotsByWk(ilVaryingQtrs + ilWeekInx, ilLoop) + llAccumSpotCount
                                    If lmSpotsByWk(ilVaryingQtrs + ilWeekInx, ilLoop) < 0 Then
                                        ilTemp = 1
                                    End If
                                    dlTemp = lmratesbywk(ilVaryingQtrs + ilWeekInx, ilLoop) + dlTemp 'TTP 10439 - Rerate 21,000,000
                                    tmCbf.lWeek(ilWeekInx - 1) = lmSpotsByWk(ilVaryingQtrs + ilWeekInx, ilLoop)
                                    tmCbf.lValue(ilWeekInx - 1) = tmLRch(llRchQtr + ilQtr - 1).lWklyGRP(ilWeekInx - 1)
                                Next ilWeekInx

                                'Build all 13 week buckets in the standard month buckets
                                gPackDate slStr, tmCbf.iDtFrstBkt(0), tmCbf.iDtFrstBkt(1)
                                llDate = gDateValue(slStr)
                                
                                '8410 - proceed with collecting monthly buckets of values on Installment contracts
                                'If tgChf.sInstallDefined <> "Y" Then           'installment contract, dont gather from sched lines, need the monthly install $ from SBF
                                    'For ilWeekInx = 1 To 13 Step 1              'detrmine which month bucket the weekly values belong in
                                    For ilWeekInx = 1 To imWeeksPerQtr(ilQtr)
                                        For ilMonthLoop = 1 To 12 Step 1
                                            'if this qtr is beyond 12 months of data, everything goes into Over 12 months bucket
                                            ilPkgOrSpot = tmCbf.lWeek(ilWeekInx - 1) 'set the # of spots this week
                                            If tmClf.sType = "E" Then               'if package is "equal", then the rate is the weekly rate, not spot rate
                                                If ilPkgOrSpot < 0 Then
                                                ilPkgOrSpot = -1
                                                ElseIf ilPkgOrSpot > 0 Then
                                                ilPkgOrSpot = 1
                                                End If
                                            End If

                                            If llDate >= llStdStartDates(13) Then
                                                If tmClf.sType = "E" Then               'if package type is "E", then the rate is the weekly rate, not spot rate
                                                    If tmCbf.lWeek(ilWeekInx - 1) < 0 And tmLnr(ilLoop).lRate < 0 Then    'decrease in $ and spots, keep it negative
                                                        tmCbf.lMonth(12) = tmCbf.lMonth(12) + (-((ilPkgOrSpot * tmLnr(ilLoop).lRate)))
                                                    Else
                                                        tmCbf.lMonth(12) = tmCbf.lMonth(12) + ((ilPkgOrSpot * tmLnr(ilLoop).lRate))
                                                    End If
                                                Else
                                                    If tmCbf.lWeek(ilWeekInx - 1) < 0 And tmLnr(ilLoop).lRate < 0 Then    'decrease in $ and spots, keep it negative
                                                        tmCbf.lMonth(12) = tmCbf.lMonth(12) + (-((tmCbf.lWeek(ilWeekInx - 1) * tmLnr(ilLoop).lRate)))
                                                    Else
                                                        tmCbf.lMonth(12) = tmCbf.lMonth(12) + ((tmCbf.lWeek(ilWeekInx - 1) * tmLnr(ilLoop).lRate))
                                                    End If
                                                End If
                                                'add monthly spot counts
                                                tmCbf.lMonthUnits(12) = tmCbf.lMonthUnits(12) + tmCbf.lWeek(ilWeekInx - 1)
                                                ilMonthLoop = 12            'end loop
                                            Else
                                                If llDate >= llStdStartDates(ilMonthLoop) And llDate < llStdStartDates(ilMonthLoop + 1) Then
                                                    If tmClf.sType = "E" Then       'if package type is "E", then the rate is the weekly rate, not spot rate
                                                        If tmCbf.lWeek(ilWeekInx - 1) < 0 And tmLnr(ilLoop).lRate < 0 Then    'decrease in $ and spots, keep it negative
                                                        tmCbf.lMonth(ilMonthLoop - 1) = tmCbf.lMonth(ilMonthLoop - 1) + (-((ilPkgOrSpot * tmLnr(ilLoop).lRate)))
                                                        Else
                                                        tmCbf.lMonth(ilMonthLoop - 1) = tmCbf.lMonth(ilMonthLoop - 1) + ((ilPkgOrSpot * tmLnr(ilLoop).lRate))
                                                        End If
                                                    Else
                                                        If tmCbf.lWeek(ilWeekInx - 1) < 0 And tmLnr(ilLoop).lRate < 0 Then    'decrease in $ and spots, keep it negative
                                                        tmCbf.lMonth(ilMonthLoop - 1) = tmCbf.lMonth(ilMonthLoop - 1) + (-((tmCbf.lWeek(ilWeekInx - 1) * tmLnr(ilLoop).lRate)))
                                                        Else
                                                        tmCbf.lMonth(ilMonthLoop - 1) = tmCbf.lMonth(ilMonthLoop - 1) + ((tmCbf.lWeek(ilWeekInx - 1) * tmLnr(ilLoop).lRate))
                                                        End If
                                                    End If
                                                    'Add monthly spot counts
                                                    tmCbf.lMonthUnits(ilMonthLoop - 1) = tmCbf.lMonthUnits(ilMonthLoop - 1) + tmCbf.lWeek(ilWeekInx - 1)
                                                    ilMonthLoop = 12            'end loop
                                                End If
                                            End If
                                        Next ilMonthLoop
                                        llDate = llDate + 7
                                    Next ilWeekInx
                                'End If
                                
                                'Output a quarter record if its a mod (where either the spot count or dollars are non-zero), or it a fullcontract
                                'and there has to be spots in the quarter ($ could be zero)
                                If ((ilTemp <> 0 Or llAccumSpotCount <> 0 Or dlTemp <> 0) And (tlBR.iThisCntMod)) Or (llAccumSpotCount <> 0 And Not tlBR.iThisCntMod) Then       '3/4/99 detect a chg for diff only :  net result of spots is zero and $ is n/c'TTP 10439 - Rerate 21,000,000
                                    mBRDaysTimes ilLoop, slDailyExists, tlSofList()   '6-20-03 format days and times for the line
                                    tmCbf.lCPP = tmLRch(llRchQtr + ilQtr - 1).lTotalCPP
                                    tmCbf.lCPM = tmLRch(llRchQtr + ilQtr - 1).lTotalCPM
                                    tmCbf.lGrImp = tmLRch(llRchQtr + ilQtr - 1).lTotalGrimps
                                    tmCbf.iAvgRate = tmLRch(llRchQtr + ilQtr - 1).iTotalAvgRating
                                    tmCbf.lAvgAud = tmLRch(llRchQtr + ilQtr - 1).lTotalAvgAud
                                    tmCbf.lGRP = tmLRch(llRchQtr + ilQtr - 1).lTotalGRP
                                    'setup quarterly  Research values for all vehicles
                                    tmCbf.lQGRP = tmQGRP(ilQtr - 1)
                                    tmCbf.lQCPP = tmQCPP(ilQtr - 1)
                                    tmCbf.lQCPM = tmQCPM(ilQtr - 1)
                                    tmCbf.lQGrimp = tmQGrimp(ilQtr - 1)

                                    For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1 Step 1
                                        If tmCbf.iVefCode = ilVehList(ilVehicle) Then
                                            Exit For
                                        End If
                                    Next ilVehicle
                                    'setup quarterly  Research values by vehicles
                                    tmCbf.lVQGRP = tmVehQtrList(ilVehicle).lVQGRP(ilQtr - 1)
                                    tmCbf.lVQCPP = tmVehQtrList(ilVehicle).lVQCPP(ilQtr - 1)
                                    tmCbf.lVQCPM = tmVehQtrList(ilVehicle).lVQCPM(ilQtr - 1)
                                    tmCbf.lVQGrimp = tmVehQtrList(ilVehicle).lVQGrimp(ilQtr - 1)

                                    tmCbf.sSnapshot = tlBR.sSnapshot
                                    mPutCntWkGrps ilQtr
                                    igBR_SSLinesExist = True      '12-16-03  flag at least 1 sch line found to print detail or summary

                                    If tgChf.sInstallDefined = "Y" Then   'if installment contract, show install $ on summary page
                                        'TTP 8410 - if this is a Hidden line, skip the duplicate check so that Installment months can be built
                                        'gBuildInstallMonths tmClf.iVefCode, ilVehiclesDone(), llStdStartDates(), llInstallBilling(), tlSbf()
                                        gBuildInstallMonths tmClf.iVefCode, ilVehiclesDone(), llStdStartDates(), llInstallBilling(), tlSbf(), tmClf.iPkLineNo > 0
                                        For ilMonthLoop = 1 To 13
                                            tmCbf.lMonth(ilMonthLoop - 1) = llInstallBilling(ilMonthLoop)
                                        Next ilMonthLoop
                                    End If
                                    
                                    tmCbf.sWeeksInQtr = "3"                 'flag in crystal to tell how many week columns to show
                                    If imWeeksPerQtr(ilQtr) = 14 Then
                                        tmCbf.sWeeksInQtr = "4"
                                    ElseIf imWeeksPerQtr(ilQtr) = 12 Then
                                        tmCbf.sWeeksInQtr = "2"
                                    End If
                                    ilVefInxForCallLetters = gBinarySearchVef(tmCbf.iVefCode)
                                    If ilVefInxForCallLetters <> -1 Then        '10-23-06 if new package vehicle entered for proposal, the index wont exist.  The vehicle
                                        gFindStationMkt ilVefInxForCallLetters, tmCbf            '2-21-18 show market name?
                                    End If
                                    
                                    If ((Asc(tgSaf(0).sFeatures4) And ACT1CODES) = ACT1CODES) Then      '2-27-18 using Act1 Codes?, if so, show lineup code reference.  Set the clfcode using cbfPop field .  cbfPop not used anywhere else in contract print report
                                        tmCbf.lPop = 0
                                        'If tmClf.sType <> "H" And Trim$(tmClf.sACT1LineupCode) <> "" Then                                  'hidden lines of package should not show the lineup to avoid too many lines printed
                                        '6/18/21 - ACT1 Task1 - JW - Add Act 1 code to Hidden Lines
                                        '11/23/21 - JW - comment out IF -- Crystal will show the Act1 data, if its there.
                                        'If Trim$(tmClf.sACT1LineupCode) <> "" Then                                   'hidden lines of package should not show the lineup to avoid too many lines printed
                                        'TTP 10382 - Contract report: Option To not show Act1 codes on PDF
                                        If tlBR.iShowAct1 = True Then
                                            tmCbf.lPop = tmClf.lCode                                'schedule line reference
                                        End If
                                    End If
                                    
                                    'TTP 8410 - Insert Pkg vef with Hidden values and set hidden lines with special code for Installment
                                    If tmClf.sType = "H" And tgChf.sInstallDefined = "Y" Then
                                        tmPkgCbf = tmCbf
                                        ilClfInx = mGetPkgLineNoFromHiddenLine(tmClf.iLine)
                                        tmPkgCbf.iVefCode = tgClf(ilClfInx).ClfRec.iVefCode
                                        tmPkgCbf.lLineNo = tgClf(ilClfInx).ClfRec.iLine
                                        tmPkgCbf.iExtra2Byte = 0
                                        tmPkgCbf.lIntComment = 1
                                        tmPkgCbf.sLineType = "Y" 'Insert some package vehicles with hidden vehicle Amounts, use Code "Y" so it can be excluded or counted for specially in the 3 reports: Br, BrSum, BrSumZer
                                        ilRet = btrInsert(hmCbf, tmPkgCbf, imCbfRecLen, INDEXKEY0)
                                        
                                        'Set Hidden lines with Code "X", so they can be excluded in the 3 reports: Br, BrSum, BrSumZer (if Proof/Hidden Lines is unchecked)
                                        If BrSnap.ckcProof.Value <> vbChecked Then
                                            tmCbf.sLineType = "X"
                                        End If
                                    End If
                                    
                                    ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                                    'init monthly $ and units buckets for next flight
                                    For ilMonthLoop = 1 To 14                '11-14-11 adjust for max 14 week qtr
                                        If ilMonthLoop = 14 Then            'weekly buckets, adjust for 14 weeks in qtr
                                            tmCbf.lWeek(ilMonthLoop - 1) = 0
                                            tmCbf.lValue(ilMonthLoop - 1) = 0
                                        End If
                                        If ilMonthLoop <= 13 Then                'monthly bkts
                                            tmCbf.lMonth(ilMonthLoop - 1) = 0
                                            tmCbf.lMonthUnits(ilMonthLoop - 1) = 0
                                            tmCbf.lWeek(ilMonthLoop - 1) = 0
                                        End If
                                    Next ilMonthLoop
                                End If
                                'prepare start date of next quarter
                                gUnpackDate tmCbf.iDtFrstBkt(0), tmCbf.iDtFrstBkt(1), slStr
                                'llDate = gDateValue(slStr) + (13 * 7)       'add 13 weeks for next quarter start date
                                llDate = gDateValue(slStr) + (imWeeksPerQtr(ilQtr) * 7)     '11-14-11 get next qtr start date based on varying # of week per qtr
                                slStr = Format$(llDate, "m/d/yy")
                                ilVaryingQtrs = ilVaryingQtrs + imWeeksPerQtr(ilQtr)
                            Next ilQtr                          'process next quarter
                        End If                              'llfltend < llfltstart
                    Else                        '3-10-06 if ignoring hidden overrides and its a hidden lines, need to flag if applicable
                        If (imHiddenOverride And HIDDENOVERRIDE) = HIDDENOVERRIDE And tmClf.sType = "H" Then
                            mBRDaysTimes ilLoop, slDailyExists, tlSofList()    'the hidden line is not shown, but need to flag at least
                                                        'one output record so research disclaimer can be shown if applicable
                        End If
                    End If                          'ilShowProof
                End If                              'if imProcessFlag = 1 or 3
            Next ilLoop                             '1 to ilLinespots
            If Not ilGot1ToPrint Then               'write out the header only - for a difference option when changing the header,
                                                    'nothing gets shown for lines -  at least show header so that the change reason is printed

                If Not tlBR.iThisCntMod Then
                'If Not ilThisCntMod Then
                    tmCbf.iPctDist = 0              'Pct dist field replaced by Difference flag (0=full)
                    tmCbf.iMnfGroup = 0             'Normal BR (not differences), init flag for Crystal
                Else
                    tmCbf.iPctDist = 1              '1 = difference only (used to print legend on BR for diff. option)
                    tmCbf.iMnfGroup = 1             '1 = nothing found to print; do special eliminating of showing fields in Crystal
                    igBR_SSLinesExist = True      '12-16-03 force output
                End If
                tmCbf.sSnapshot = tlBR.sSnapshot
                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            End If

            '******new 4-23-99 (copied from v4.5)
            'Generate line totals (prepare for vehicle summary)
            
            '3-15-19 process only package vehicles, combining all hidden lines for each packages vehicles with same name
            'ie a package may be used more than once, maybe for different dp or spot lengths
            mGatherPkgVehSummary ilLineSpots, ilPkgVehList(), ilVehList(), llOverallPopEst, llLnSpots, llPopByLine(), llWklySpots(), llWklyRates(), llWklyAvgAud(), llWklyPopEst(), lmSpotsByWk(), lmWklyRates(), lmAvgAud(), lmPopEst(), ilWklyRtg(), llWklyGrimp(), llWklyGRP()

            'Create array of grps, wkly rating, grimps and rates to pass to routine to generate vehicle totals for summary
            '3-15-19 hidden and std vehicle summary
            For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1 Step 1
                llLnSpots = 0
                llPop = lmPop(ilVehicle)        '11-23-99

                ReDim tmLRch(0 To 1) As RESEARCHINFO        '10-30-01
                For ilLoop = 1 To ilLineSpots - 1
                    If imProcessFlag(tmLnr(ilLoop).iLineInx) = 1 Or imProcessFlag(tmLnr(ilLoop).iLineInx) = 3 Or imProcessFlag(tmLnr(ilLoop).iLineInx) = 4 Then  'show current and prev lines depending
                        ilLoop3 = tmLnr(ilLoop).iLineInx            'line index to process
                                                'on the option selected.  Mods show current & prev, Full BR shows curent only
                                                'Move spots, $ and aud to single dimension array for gAvgAudToLnResearch routine
                        tmClf = tgClf(ilLoop3).ClfRec
                        llPop = llPopByLine(ilLoop3 + 1)    '11-23-99
                        ilFoundLine = False
                        '4-10-19 change ilTemp to llTempLrch
                        llTempLRch = UBound(tmLRch)
    
                        If tmClf.sType = "H" Then                   'hidden line to a package
                            If tlBR.iShowProof And ilVehList(ilVehicle) = tmClf.iVefCode Then
                                ilFoundLine = True
                            Else
                                '11-29-04 if not showing hidden lines, see if this hidden vehicle is also a package vehicle to process. if so, it needs to be processed
                                If Not tlBR.iShowProof And ilVehList(ilVehicle) = tmClf.iVefCode Then
                                    For ilHiddenAndConv = LBound(ilPkgVehList) To UBound(ilPkgVehList) - 1
                                        If ilPkgVehList(ilHiddenAndConv) = tmClf.iVefCode Then
                                            ilFoundLine = True
                                            Exit For
                                        End If
                                    Next ilHiddenAndConv
                                End If
                            End If
                        ElseIf tmClf.sType = "S" Then
                            If tmClf.iVefCode = ilVehList(ilVehicle) Then
                                ilFoundLine = True
                            End If
                        Else
                        
                        End If
        
                        If ilFoundLine Then
                            For ilSpots = 1 To MAXWEEKSFOR2YRS Step 1
                                llWklySpots(ilSpots - 1) = lmSpotsByWk(ilSpots, ilLoop)
                                llLnSpots = llLnSpots + llWklySpots(ilSpots - 1)
                                tmLRch(llTempLRch).lQSpots = tmLRch(llTempLRch).lQSpots + llWklySpots(ilSpots - 1)
                                llWklyRates(ilSpots - 1) = lmWklyRates(ilSpots, ilLoop3)
                                llWklyAvgAud(ilSpots - 1) = lmAvgAud(ilSpots, ilLoop3)
                                llWklyPopEst(ilSpots - 1) = lmPopEst(ilSpots, ilLoop3)
                            Next ilSpots
                            'gAvgAudToLnResearch sm1or2PlaceRating, True, llPop, llWklyPopEst(), llWklySpots(), llWklyRates(), llWklyAvgAud(), llTemp, tmCbf.lAvgAud, ilWklyRtg(), tmCbf.iAvgRate, llWklyGrimp(), tmCbf.lGrImp, llWklyGRP(), tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, llPopEst
                            gAvgAudToLnResearch sm1or2PlaceRating, True, llPop, llWklyPopEst(), llWklySpots(), llWklyRates(), llWklyAvgAud(), dlTemp, tmCbf.lAvgAud, ilWklyRtg(), tmCbf.iAvgRate, llWklyGrimp(), tmCbf.lGrImp, llWklyGRP(), tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
                            tmLRch(llTempLRch).iVefCode = ilVehList(ilVehicle)
                            'mRchSameData llTempLRch, llTemp
                            mRchSameData llTempLRch, dlTemp 'TTP 10439 - Rerate 21,000,000
                        End If
                    End If                          'imProcessFlag(tmLnr(ilLoop).iLineInx) = 1
                Next ilLoop

                'Line totals generated for max 2 years each, total all lines from results of each line
                'For Summary Page (1 line per vehicle)
                ReDim tmVRtg(0 To 0) As Integer
                ReDim tmVCost(0 To 0) As Long
                ReDim tmVGRP(0 To 0) As Long
                ReDim tmVGrimp(0 To 0) As Long

                '4-10-19 change ilRchQtr to llRchQTr due to subscript out of range
                For llRchQtr = LBONE To UBound(tmLRch) - 1 Step 1
                    If llLnSpots <> 0 Then
                        tmVRtg(UBound(tmVRtg)) = tmLRch(llRchQtr).iTotalAvgRating
                        tmVCost(UBound(tmVCost)) = tmLRch(llRchQtr).dTotalCost 'TTP 10439 - Rerate 21,000,000
                        tmVGRP(UBound(tmVGRP)) = tmLRch(llRchQtr).lTotalGRP
                        tmVGrimp(UBound(tmVGrimp)) = tmLRch(llRchQtr).lTotalGrimps
                        ReDim Preserve tmVRtg(0 To UBound(tmVRtg) + 1) As Integer
                        ReDim Preserve tmVCost(0 To UBound(tmVCost) + 1) As Long
                        ReDim Preserve tmVGRP(0 To UBound(tmVGRP) + 1) As Long
                        ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) + 1) As Long
                    End If
                Next llRchQtr
                If UBound(tmVRtg) > 0 Then
                    'dimensions must be exact sizes
                    ReDim Preserve tmVRtg(0 To UBound(tmVRtg) - 1) As Integer
                    ReDim Preserve tmVCost(0 To UBound(tmVCost) - 1) As Long
                    ReDim Preserve tmVGRP(0 To UBound(tmVGRP) - 1) As Long
                    ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) - 1) As Long
                    If UBound(tmVRtg) >= 0 Then
                        llPop = lmPop(ilVehicle)   '11-23-99
                        gResearchTotals sm1or2PlaceRating, True, llPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, dlTemp, tmCbf.iAvgRate, tmCbf.lGrImp, tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, tmCbf.lAvgAud 'TTP 10439 - Rerate 21,000,000
                    Else
                        tmCbf.lCPP = 0
                        tmCbf.lCPM = 0
                        tmCbf.lGRP = 0
                        tmCbf.lGrImp = 0
                    End If
                    
                    '3-31-21 Save the total air time gross impressions and cost to combine with any adserver lines.  The CPM and gross impressions will be shown on the Research Summary page
                    'convert the gross impressions to units (to be same with the adserver base) based on the Site Aud feature
                    If tgSpf.sSAudData = "H" Then     'hundreds
                        slStr = "100"
                    ElseIf tgSpf.sSAudData = "N" Then   'tens
                        slStr = "10"
                    ElseIf tgSpf.sSAudData = "U" Then   'units
                        slStr = "1"
                    Else        'tgspf.sSAudData = "T"   'thousands
                        slStr = "1000"
                    End If
    
                    lgAirTimeGross = lgAirTimeGross + dlTemp      'total cost of air time gross'TTP 10439 - Rerate 21,000,000
                    slTempGrimp = gLongToStrDec(tmCbf.lGrImp, 0)    'grimps for line
                    slTempGrimp = gMulStr(slTempGrimp, slStr)          'adjust to be by units
                    sgAirTimeGrimp = gAddStr(sgAirTimeGrimp, slTempGrimp)       'accum total unit grimps for air time

                    tmCbf.iExtra2Byte = 2               'vehicle summary totals
                    tmCbf.iVefCode = ilVehList(ilVehicle)

                    tmCbf.lQGRP = 0
                    tmCbf.lQCPP = 0
                    tmCbf.lQCPM = 0
                    tmCbf.lQGrimp = 0
                    tmCbf.lVQGRP = 0
                    tmCbf.lVQCPP = 0
                    tmCbf.lVQCPM = 0
                    tmCbf.lVQGrimp = 0
                    tmCbf.sSnapshot = tlBR.sSnapshot
                    igBR_SSLinesExist = True      '12-16-03 force output
                    tmCbf.sMixTypes = ""                'default to show vehicle grp, cpp field
                    For ilTemp = 0 To UBound(tmPodcast_Info) - 1
                        If tmPodcast_Info(ilTemp).iVefCode = tmCbf.iVefCode Then
                            'type of line:  P = Podcast, K = package line, H = Podcast Hidden Line, L = Other, not podcast Hidden Line, O = other, not podcast (conventional, selling)
                            'look for package vehicles to determine how to show the vehicle cpp, grp columns
                            If tmPodcast_Info(ilTemp).sType = "K" Then
                                If Not tmPodcast_Info(ilTemp).bShowResearch Then
                                    tmCbf.sMixTypes = "H"
                                End If
                                Exit For
                            Else
                                If tmPodcast_Info(ilTemp).sType = "P" Or tmPodcast_Info(ilTemp).sType = "H" Then        'podcast vehicle not in hidden line (P), or in pkg (H)
                                    tmCbf.sMixTypes = "H"
                                End If
                                Exit For
                            End If
                        End If
                    Next ilTemp

                    ilVefInxForCallLetters = gBinarySearchVef(tmCbf.iVefCode)
                    If ilVefInxForCallLetters <> -1 Then        '10-23-06 if new package vehicle entered for proposal, the index wont exist.  The vehicle
                        gFindStationMkt ilVefInxForCallLetters, tmCbf            '2-21-18 show market name?
                    End If
                    
                    ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
                End If
            Next ilVehicle

            'Create array of grps, wkly rating, grimps and rates to pass to routine to generate contract totals for summary
            ilfirstTime = True
            'ReDim tmLRch(1 To 1) As RESEARCHLIST
            'ReDim tmLRch(1 To 1) As RESEARCHINFO
            ReDim tmLRch(0 To 1) As RESEARCHINFO    'Index zero ignored
            For ilLoop = 1 To ilLineSpots - 1
                If imProcessFlag(tmLnr(ilLoop).iLineInx) = 1 Or imProcessFlag(tmLnr(ilLoop).iLineInx) = 3 Or imProcessFlag(tmLnr(ilLoop).iLineInx) = 4 Then  'show current and prev lines depending
                    ilLoop3 = tmLnr(ilLoop).iLineInx            'line index to process
                    tmClf = tgClf(ilLoop3).ClfRec
                    'on the option selected.  Mods show current & prev, Full BR shows curent only
                    'Move spots, $ and aud to single dimension array for gAvgAudToLnResearch routine
                    'ignore the package lines for summary totals, generate contract research totals on hidden and std lines
                    If tmClf.sType <> "A" And tmClf.sType <> "O" And tmClf.sType <> "E" Then
                        If ilfirstTime Then
                            '4-10-19 replace ilTemp with ilTempLRch due to subscript out of range
                            ilfirstTime = False
                            llTempLRch = UBound(tmLRch)
                            tmLRch(llTempLRch).lQSpots = 0
                            For ilSpots = 1 To MAXWEEKSFOR2YRS Step 1
                                llWklySpots(ilSpots - 1) = 0
                                llWklyRates(ilSpots - 1) = 0
                                llWklyAvgAud(ilSpots - 1) = 0
                                llWklyPopEst(ilSpots - 1) = 0
                            Next ilSpots
                            ilFoundLine = tmClf.iLine
                            ilFoundVeh = tmClf.iVefCode
                        End If
                        If ilFoundLine = tmClf.iLine Then
                            llPop = llPopByLine(ilLoop3 + 1)  '11-23-99    'on the option selected.  Mods show current & prev, Full BR shows curent only
                            For ilSpots = 1 To MAXWEEKSFOR2YRS Step 1
                                llWklySpots(ilSpots - 1) = llWklySpots(ilSpots - 1) + lmSpotsByWk(ilSpots, ilLoop)
                                tmLRch(llTempLRch).lQSpots = tmLRch(llTempLRch).lQSpots + lmSpotsByWk(ilSpots, ilLoop)
                                'llWklyRates(ilSpots) = llWklyRates(ilSpots) + lmWklyRates(ilSpots, ilLoop3)
                                If llWklyRates(ilSpots - 1) = 0 Then
                                    llWklyRates(ilSpots - 1) = lmWklyRates(ilSpots, ilLoop3)
                                End If
                                'llWklyAvgAud(ilSpots) = llWklyAvgAud(ilSpots) + lmAvgAud(ilSpots, ilLoop3)
                                If llWklyAvgAud(ilSpots - 1) = 0 Then
                                    llWklyAvgAud(ilSpots - 1) = lmAvgAud(ilSpots, ilLoop3)
                                    llWklyPopEst(ilSpots - 1) = lmPopEst(ilSpots, ilLoop3)
                                End If
                            Next ilSpots
                        Else
                            gAvgAudToLnResearch sm1or2PlaceRating, True, llPop, llWklyPopEst(), llWklySpots(), llWklyRates(), llWklyAvgAud(), dlTemp, tmCbf.lAvgAud, ilWklyRtg(), tmCbf.iAvgRate, llWklyGrimp(), tmCbf.lGrImp, llWklyGRP(), tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
                            'gAvgAudToLnResearch True, lmPop(ilLoop3), llWklySpots(), llWklyRates(), llWklyAvgAud(), llTemp, tmCbf.lAvgAud, ilWklyRtg(), tmCbf.iAvgRate, llWklyGrimp(), tmCbf.lGrimp, llWklyGRP(), tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM
                            'Update array with results of research routine (didn't use the actual field names in the subroutine call because
                            'the total parameters were too long
                            tmLRch(llTempLRch).iVefCode = ilFoundVeh
                            tmLRch(llTempLRch).dTotalCost = dlTemp 'TTP 10439 - Rerate 21,000,000
                            tmLRch(llTempLRch).lTotalCPP = tmCbf.lCPP
                            tmLRch(llTempLRch).lTotalCPM = tmCbf.lCPM
                            tmLRch(llTempLRch).lTotalGRP = tmCbf.lGRP
                            tmLRch(llTempLRch).lTotalGrimps = tmCbf.lGrImp
                            tmLRch(llTempLRch).iTotalAvgRating = tmCbf.iAvgRate
                            tmLRch(llTempLRch).lTotalAvgAud = tmCbf.lAvgAud
                            ReDim Preserve tmLRch(0 To UBound(tmLRch) + 1)  'Index zero ignored
                            llTempLRch = UBound(tmLRch)
                            tmLRch(llTempLRch).lQSpots = 0
                            'initialize arrays for next sched line
                            For ilSpots = 1 To MAXWEEKSFOR2YRS Step 1
                                llWklySpots(ilSpots - 1) = 0
                                tmLRch(llTempLRch).lQSpots = 0
                                llWklyRates(ilSpots - 1) = 0
                                llWklyAvgAud(ilSpots - 1) = 0
                                llWklyPopEst(ilSpots - 1) = 0
                            Next ilSpots
                            For ilSpots = 1 To MAXWEEKSFOR2YRS Step 1
                                llWklySpots(ilSpots - 1) = llWklySpots(ilSpots - 1) + lmSpotsByWk(ilSpots, ilLoop)
                                tmLRch(llTempLRch).lQSpots = tmLRch(llTempLRch).lQSpots + lmSpotsByWk(ilSpots, ilLoop)
                                'llWklyRates(ilSpots) = llWklyRates(ilSpots) + lmWklyRates(ilSpots, ilLoop3)
                                If llWklyRates(ilSpots - 1) = 0 Then
                                    llWklyRates(ilSpots - 1) = lmWklyRates(ilSpots, ilLoop3)
                                End If
                                'llWklyAvgAud(ilSpots) = llWklyAvgAud(ilSpots) + lmAvgAud(ilSpots, ilLoop3)
                                If llWklyAvgAud(ilSpots - 1) = 0 Then
                                    llWklyAvgAud(ilSpots - 1) = lmAvgAud(ilSpots, ilLoop3)
                                    llWklyPopEst(ilSpots - 1) = lmPopEst(ilSpots, ilLoop3)
                                End If
                            Next ilSpots
                            ilFoundLine = tmClf.iLine
                            ilFoundVeh = tmClf.iVefCode
                            'for next lines population
                            llPop = llPopByLine(ilLoop3 + 1)   '11-23-99    'on the option selected.  Mods show current & prev, Full BR shows curent only
                        End If
                    End If
                End If
            Next ilLoop
            
            'Process the research data for the last line sitting in arrays
            gAvgAudToLnResearch sm1or2PlaceRating, True, llPop, llWklyPopEst(), llWklySpots(), llWklyRates(), llWklyAvgAud(), dlTemp, tmCbf.lAvgAud, ilWklyRtg(), tmCbf.iAvgRate, llWklyGrimp(), tmCbf.lGrImp, llWklyGRP(), tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
            
            'Update array with results of research routine (didn't use the actual field names in the subroutine call because
            'the total parameters were too long
            If UBound(tmLRch) >= llTempLRch Then
                tmLRch(llTempLRch).iVefCode = ilFoundVeh
                tmLRch(llTempLRch).dTotalCost = dlTemp 'TTP 10439 - Rerate 21,000,000
                tmLRch(llTempLRch).lTotalCPP = tmCbf.lCPP
                tmLRch(llTempLRch).lTotalCPM = tmCbf.lCPM
                tmLRch(llTempLRch).lTotalGRP = tmCbf.lGRP
                tmLRch(llTempLRch).lTotalGrimps = tmCbf.lGrImp
                tmLRch(llTempLRch).iTotalAvgRating = tmCbf.iAvgRate
                tmLRch(llTempLRch).lTotalAvgAud = tmCbf.lAvgAud
                ReDim Preserve tmLRch(0 To UBound(tmLRch) + 1)  'Index zero ignored
            End If

            'Line totals generated for max 2 years each, total all lines from results of each line
            ReDim tmVRtg(0 To 0) As Integer
            ReDim tmVCost(0 To 0) As Long
            ReDim tmVGRP(0 To 0) As Long
            ReDim tmVGrimp(0 To 0) As Long
            '3/13/04- moved llLnSpots = 0 here
            llLnSpots = 0
            '4-10-19 change ilRchQtr to llRchQtr due to subscript out of range
            For llRchQtr = LBONE To UBound(tmLRch) - 1 Step 1
'                llLnSpots = 0
                If tmLRch(llRchQtr).lQSpots <> 0 Then   'And tmLRch(ilRchQtr).lTotalGrimps Then
                    llLnSpots = llLnSpots + tmLRch(llRchQtr).lQSpots
                    tmVRtg(UBound(tmVRtg)) = tmLRch(llRchQtr).iTotalAvgRating
                    'tmVCost(UBound(tmVCost)) = tmLRch(llRchQtr).lTotalCost
                    tmVCost(UBound(tmVCost)) = tmLRch(llRchQtr).dTotalCost 'TTP 10439 - Rerate 21,000,000
                    tmVGRP(UBound(tmVGRP)) = tmLRch(llRchQtr).lTotalGRP
                    tmVGrimp(UBound(tmVGrimp)) = tmLRch(llRchQtr).lTotalGrimps
                    ReDim Preserve tmVRtg(0 To UBound(tmVRtg) + 1) As Integer
                    ReDim Preserve tmVCost(0 To UBound(tmVCost) + 1) As Long
                    ReDim Preserve tmVGRP(0 To UBound(tmVGRP) + 1) As Long
                    ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) + 1) As Long
                End If
            Next llRchQtr
            'If UBound(tmVRtg) > 1 Then
            If UBound(tmVRtg) > 0 Then
                'dimensions must be exact sizes
                ReDim Preserve tmVRtg(0 To UBound(tmVRtg) - 1) As Integer
                ReDim Preserve tmVCost(0 To UBound(tmVCost) - 1) As Long
                ReDim Preserve tmVGRP(0 To UBound(tmVGRP) - 1) As Long
                ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) - 1) As Long
                gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, dlTemp, tmCbf.iAvgRate, tmCbf.lGrImp, tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, tmCbf.lAvgAud 'TTP 10439 - Rerate 21,000,000
                tmCbf.iExtra2Byte = 3               'contract totals
                tmCbf.lQGRP = 0
                tmCbf.lQCPP = 0
                tmCbf.lQCPM = 0
                tmCbf.lQGrimp = 0
                tmCbf.lVQGRP = 0
                tmCbf.lVQCPP = 0
                tmCbf.lVQCPM = 0
                tmCbf.lVQGrimp = 0
                tmCbf.sSnapshot = tlBR.sSnapshot
                igBR_SSLinesExist = True      '12-16-03 force output
                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            End If
            mCreateBRSportsComments
#If programmatic <> 1 Then
            gBuildSplitNetStations hmRaf, hmSef, hmTxr, hmShf, hmMkt, hmVLF, hmAtt, tgClf(), tgChf
#End If
        End If                                      'tgChf.imnfDemo(ilDemo) > 0
        If Not tlBR.iAllDemos Then            'use primary demo only
            ilDemo = 4                              'force to get out of demo loop
        End If
        Next ilDemo

        mProcessBrNTR tlBR.iWhichSort, tlBR.iShowNTRBillSummary, tlRvf(), tlSbf(), ilVehiclesDone(), llStdStartDates()                  '12-9-02 create records for NTR summary
        blPodExists = gBuildCPMIDs(hmPcf, tgChf, llStdStartDates(), tmCPM_IDs(), tmCPMSummary())
        If blPodExists Then
            mProcessBR_CPM tlBR.iWhichSort, tlBR.iShowProof, tmCPM_IDs(), tmCPMSummary()      'create the detail IDs, summary by vehicle Research, and vehicle billing vehicle
        End If
        Exit Function
    Else                                                        'close all files, eoj
        mCloseBRFiles
        Erase llPopByLine           '11-23-99
        Erase lmPopPkgByLine        '11-11-05
        Erase tmLnr ', lmMatchingCnts 'lgPrintedCnts
        Erase lmSpotsByWk, imDnfCodes, imProcessFlag        '12-19-13 remove impkgdnfcodes
        Erase lmratesbywk
        Erase lmWklySpots, lmWklyRates, lmAvgAud, lmPopEst
        Erase llStdStartDates, lmPop, tmLRch    ', tmPrevLRch
        Erase ilVehList, tmVCost, tmVRtg, tmVGRP, tmVGrimp
        Erase tmQGRP, tmQCPP, tmQCPM, tmQGrimp, tmVehQtrList
        Erase ilPkgVehList, ilPkgLineList, ilVehList
        Erase tmPkVCost, tmPkVRtg, tmPkVGrimp, tmPkVGRP ', tmPkVehQtrList
        Erase tmPkLnCost, tmPkLnGrimp, tmPkLnGRP  ', tmPkLnQtrList
        Erase tmAnfTable
        Erase imAirWks, ilVehiclesDone, llInstallBilling, tlSofList, lmPopPkg, tlRvf     '12-19-13 remove ilPkgvehicleforsurveylist

        Erase tmWkCntGrps, tmWkCntVGrps, tmWkHiddenVGrps, tmWkPkgVGrps  '11-6-13
        Erase llProjectedTax1, llProjectedTax2, llProjectedFlights      '11-6-13
        Erase tlSbf                                                     '11-6-13
        Erase tmCPM_IDs, tmCPMSummary                                       '12-16-20
        If ilTask = REPORTSJOB Then
            Erase tgClf, tgCff
        End If
        Exit Function
    End If
End Function

'
'               mBRDaystimes - Format Days & Times to print for each
'                   schedule line on the Broadcast Contract printout
'
'               <input> ilLoop - index into the array (tmLNR) containing data to process
'                       slDailyExists - Y if at least one daily exists on the order, else N
'               created:  4/22/98
'               11-19-02 Implement daily buys
'               mBRDaysTimes ilLoop
'
Sub mBRDaysTimes(ilLoop As Integer, slDailyExists As String, tlSofList() As SOFLIST)
    Dim ilRet As Integer
    Dim ilLoop2 As Integer
    Dim ilShowOVDays As Integer
    Dim ilShowOVTimes As Integer
    Dim slStr As String
    Dim ilTemp As Integer
    Dim slPrefStartTime As String
    Dim slPrefEndTime As String
    Dim slPrefDays As String
    Dim slTempDays As String
    Dim ilDays(0 To 6) As Integer
    Dim slPrefDaysOfWk(0 To 6) As String * 1
    Dim slDay As String
    Dim slOVStartTime As String
    Dim slOVEndTime As String
    Dim ilPrefDaysExists As Integer
    Dim ilLoop3 As Integer
    Dim llrunningStartTime As Long
    Dim llrunningEndtime As Long
    Dim ilXMid As Integer
    Dim slOVTemp As String
    Dim ilAnfIndex As Integer
    Dim llLineRef As Long

    tmCbf.lRate = tmLnr(ilLoop).lRate
    tmCbf.sPriceType = tmLnr(ilLoop).sPriceType     '2-23-01 sch line price type (adu, recap, bonus, etc)
    tmCbf.sDysTms = tmLnr(ilLoop).sValidDays
    If slDailyExists = "N" Then        '5-21-03
        tmCbf.sDailyWkly = "0"            'weekly only
    Else
        If tmLnr(ilLoop).sDailyWkly = "D" Then
            tmCbf.sDailyWkly = "1"
        Else
            tmCbf.sDailyWkly = "2"
        End If
    End If

    tmCbf.iLen = tmClf.iLen
    tmCbf.lLineNo = CLng((tmClf.iLine) * CLng(1000)) + tmClf.iCntRevNo
    tmCbf.iVefCode = tmClf.iVefCode
    tmCbf.lRafCode = tmClf.lRafCode '8-30-06

'       4/14/99 Format the schedule line for the Contract printout (BR)
'       Package/Hidden lines, then conventional lines.

    tmCbf.iOBBLen = tmClf.iBBOpenLen
    tmCbf.iCBBLen = tmClf.iBBCloseLen
    tmCbf.sCBS = ""             'flag to indicate to crystal how to show the open/close bb
    If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then    'use closest avail, dont know if its open or close
        tmCbf.sCBS = "C"       'flag to use closest to show on report
    Else                            'open & closes are defined
       tmCbf.sCBS = "S"        'use specific avail
    End If

'        tmCbf.sResort = ""
'        tmCbf.sResortType = ""          '5-31-05
'        If tmClf.sType = "H" Then               'hidden
'            slStr = Trim$(str$(tmCbf.lIntComment))  'package line # stored in this field
'            Do While Len(slStr) < 4         '5-31-05 use 4 digit line #s (vs 3 digit line #s)
'            slStr = "0" & slStr
'            Loop
'            tmCbf.sResort = slStr '& "C"
'            tmCbf.sResortType = "C"         '5-31-05
'        ElseIf tmClf.sType = "A" Or tmClf.sType = "O" Or tmClf.sType = "E" Then     'packages
'            slStr = Trim$(str$(tmClf.iLine))
'            Do While Len(slStr) < 4         '5-31-05
'            slStr = "0" & slStr
'            Loop
'            tmCbf.sResort = slStr '& "A"
'            tmCbf.sResortType = "A"         '5-31-05
'        Else                                    'conventionals, all others (fall after package/hiddens)
'            tmCbf.sResort = "9999"  '~"
'            tmCbf.sResortType = "~"         '5-31-05
'        End If

    llLineRef = tmClf.iLine
    If tmClf.sType = "H" Then
        llLineRef = tmClf.iPkLineNo
    End If
    mSetResortField tmClf.sType, llLineRef               '12-22-20

    tmRdfSrchKey.iCode = tmClf.iRdfCode
    ilRet = btrGetEqual(hmRdf, tmRdf, imRdfRecLen, tmRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching r/c recd
    If ilRet <> BTRV_ERR_NONE Then
        'What do I do here?
    End If
    For ilLoop2 = 0 To 6 Step 1
        'If tmRdf.sWkDays(7, ilLoop2 + 1) = "Y" Then             'is DP is a valid day
        If tmRdf.sWkDays(6, ilLoop2) = "Y" Then             'is DP is a valid day
            If tmLnr(ilLoop).iCffDays(ilLoop2) >= 0 Then         '11-19-02 is flight a valid day? 0=invalid day
                ilShowOVDays = True
                Exit For
            Else
                ilShowOVDays = False
            End If
        End If
    Next ilLoop2
    'Times
    ilShowOVTimes = False
    tmCbf.iRdfDPSort = gFindDPSort(tmMRif(), tmMRDF(), tmClf.iRdfCode, tmClf.iVefCode)
    If tmCbf.iRdfDPSort < 0 Then
        tmCbf.iRdfDPSort = tmClf.iLine
    End If

    If ((tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And ((tmClf.iEndTime(0) <> 1) Or (tmClf.iEndTime(1) <> 0)) Then
        gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slOVStartTime
        gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slOVEndTime
        'tmCbf.sDysTms = Trim$(tmCbf.sDysTms) & " " & Trim$(slStartTime) & "-" & Trim$(slEndTime)
        'tmCbf.irdfDPSort = 0            'override exists, show that instead of DP descrption
        'init the daypart code so it won't print
        ilShowOVTimes = True
        gAdjOverrideTimes tmClf.iVefCode, slOVStartTime, slOVEndTime    'adjust the override times to a specified time zone (site) if applicable
        slOVTemp = slOVStartTime & "-" & slOVEndTime
    Else
        'Add times
        'tmCbf.iRdfDPSort = tmClf.iRdfcode           'DP name
        '3-6-07 determine how to sort the order:  using R/C Items sort code, DP sort code, or Sch Line #
        ilXMid = False
        'if there are multiple segments and it cross midnight, show the earliest start time and xmidnight end time;
        'otherwise the first segments start and end times are shown
        slOVTemp = ""           'init forming of dp segment times
        For ilLoop2 = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1 'Row
            If (tmRdf.iStartTime(0, ilLoop2) <> 1) Or (tmRdf.iStartTime(1, ilLoop2) <> 0) Then
                gUnpackTime tmRdf.iStartTime(0, ilLoop2), tmRdf.iStartTime(1, ilLoop2), "A", "1", slOVStartTime
                gUnpackTime tmRdf.iEndTime(0, ilLoop2), tmRdf.iEndTime(1, ilLoop2), "A", "1", slOVEndTime
                gUnpackTimeLong tmRdf.iEndTime(0, ilLoop2), tmRdf.iEndTime(1, ilLoop2), True, llrunningEndtime
                '5-30-18 chged from <> 7 to <>6 due to 0 based array
                If llrunningEndtime = 86400 And ilLoop2 <> 6 Then     '6-8-10 if the running time is 12m and its not the first entry, process for x-midnite
                    ilXMid = True
                End If
                'For ilLoop3 = ilLoop2 + 1 To 7
                For ilLoop3 = ilLoop2 + 1 To UBound(tmRdf.iStartTime, 2)
                    gUnpackTimeLong tmRdf.iStartTime(0, ilLoop3), tmRdf.iStartTime(1, ilLoop3), False, llrunningStartTime
                     If llrunningStartTime = 0 And llrunningEndtime = 86400 Then
                        If ilXMid Then
                            gUnpackTime tmRdf.iEndTime(0, ilLoop3), tmRdf.iEndTime(1, ilLoop3), "A", "1", slOVEndTime
                            Exit For
                        End If
                    Else
                        gUnpackTimeLong tmRdf.iEndTime(0, ilLoop3), tmRdf.iEndTime(1, ilLoop3), True, llrunningEndtime
                    End If
                Next ilLoop3
                
                '4-12-10 show all times defined in the DP when there is an override
                'If ilXMid Then
                    gAdjOverrideTimes tmClf.iVefCode, slOVStartTime, slOVEndTime        'adjust the times (by option in Site), if x-mid daypart
                'End If
                If slOVTemp = "" Then
                    slOVTemp = slOVStartTime & "-" & slOVEndTime
                Else
                    slOVTemp = slOVTemp & "," & slOVStartTime & "-" & slOVEndTime
                End If
                If ilXMid Then
                    ilShowOVTimes = True
                    Exit For
                End If

                'need to build times (like below code) for multiple times in the rate card
                'ilNoTimes = ilNoTimes + 1
                'If ilNoTimes > UBound(smOrdered, 2) Then
                '    ReDim Preserve smOrdered(1 To 12, 1 To ilNoTimes) As String
                'End If
                'tmCbf.sDysTms = Trim$(tmCbf.sDysTms) & " " & Trim$(slStartTime) & "-" & Trim$(slEndTime)
                'Exit For
            End If
        Next ilLoop2
    End If

    slPrefDays = tmCbf.sDysTms          'default if no preferred days defined and the preferred must be shown
    If ilShowOVDays Or ilShowOVTimes Then
        'tmCbf.sDysTms = RTrim$(tmCbf.sDysTms) & " " & Trim$(slOVTemp)       '4-12-10 show multiple dp times if applicable
        '4-14-11 option to choose what to show as additional description when DP overrides exist
        If tmRdf.sOverridePrtRules = "D" Then   'show DP name
            tmCbf.sDysTms = RTrim$(tmCbf.sDysTms) & " " & Trim$(slOVTemp) & sgCR & sgLF & "(" & Trim$(tmRdf.sName) & ")"     '4-14-11 show the DP name along with the override desc
        ElseIf tmRdf.sOverridePrtRules = "A" Then  'show Avail Name
            ilAnfIndex = gBinarySearchAnf(tmRdf.ianfCode, tmAnfTable())
            If ilAnfIndex <> -1 Then
                'include or exclude the AVail
                If tmRdf.sInOut = "I" Then      'include avail
                    tmCbf.sDysTms = RTrim$(tmCbf.sDysTms) & " " & Trim$(slOVTemp) & sgCR & sgLF & "(" & Trim$(tmAnfTable(ilAnfIndex).sName) & ")"     '4-14-11 show the DP name along with the override desc
               Else                            'exclude avail
                    tmCbf.sDysTms = RTrim$(tmCbf.sDysTms) & " " & Trim$(slOVTemp) & sgCR & sgLF & "(Excl " & Trim$(tmAnfTable(ilAnfIndex).sName) & ")"     '4-14-11 show the DP name along with the override desc
                End If
            Else
                tmCbf.sDysTms = RTrim$(tmCbf.sDysTms) & " " & Trim$(slOVTemp)
            End If
        Else
            tmCbf.sDysTms = RTrim$(tmCbf.sDysTms) & " " & Trim$(slOVTemp)
        End If
        'init the daypart code to it won't print .  Print overrride time instead
        'tmCbf.irdfDPSort = 0       'DP code
        If (imHiddenOverride And HIDDENOVERRIDE) = HIDDENOVERRIDE And tmClf.sType = "H" Then
            tmCbf.sHiddenOverride = "Y"     '3-10-06 show Site Research disclaimer on contract
        End If
    Else
        tmCbf.sDysTms = tmRdf.sName
        'tmCbf.irdfDPSort = tgClf(ilLoop3).ClfRec.iRdfCode           'DP name
    End If

    ilPrefDaysExists = False        'assume preferred days do not exists
    If (tmClf.iPrefStartTime(0) = 1 And tmClf.iPrefStartTime(1) = 0) And (tmClf.sPrefDays(0) = "") Then         'test for preferred days or times
        'no preferred times
        ilRet = ilRet
    Else
        gUnpackTime tmClf.iPrefStartTime(0), tmClf.iPrefStartTime(1), "A", "1", slPrefStartTime
        gUnpackTime tmClf.iPrefEndTime(0), tmClf.iPrefEndTime(1), "A", "1", slPrefEndTime


        If tmClf.sPrefDays(0) <> "" Then
            For ilTemp = 0 To 6
                slPrefDaysOfWk(ilTemp) = ""     'unused for now, reqd for gDaynames parameter
                If tmClf.sPrefDays(ilTemp) = "Y" Then
                    ilDays(ilTemp) = 1
                    ilPrefDaysExists = True
                Else
                    ilDays(ilTemp) = 0
                End If
            Next ilTemp

            If ilPrefDaysExists Then
                slPrefDays = ""
                slTempDays = gDayNames(ilDays(), slPrefDaysOfWk(), 2, slStr)            'slstr not needed when returned
                For ilTemp = 1 To Len(slTempDays) Step 1
                    slDay = Mid$(slTempDays, ilTemp, 1)
                    If slDay <> " " And slDay <> "," Then
                        slPrefDays = Trim$(slPrefDays) & Trim$(slDay)
                    End If
                Next ilTemp
            Else
                'preferred days dont exist, show override days or r/c days if preferred times need to be shown
                If tmLnr(ilLoop).sDailyWkly = "D" Then      'daily vs wkly
                     slPrefDays = ""
                    slTempDays = gDayNames(ilDays(), slPrefDaysOfWk(), 2, slStr)            'slstr not needed when returned
                    For ilTemp = 1 To Len(slTempDays) Step 1
                        slDay = Mid$(slTempDays, ilTemp, 1)
                        If slDay <> " " And slDay <> "," Then
                            slPrefDays = Trim$(slPrefDays) & Trim$(slDay)
                        End If
                    Next ilTemp
                End If
            End If

        End If
        If slPrefStartTime <> "" Or ilPrefDaysExists Then
           If slPrefStartTime = "" Then
                slPrefStartTime = slOVStartTime
                slPrefEndTime = slOVEndTime
            End If
            tmCbf.sPrefDT = Trim$(slPrefDays) & " " + Trim$(slPrefStartTime) & "-" & Trim$(slPrefEndTime)
        End If
     End If
End Sub

'
'           Generate the Quarterly Research totals for packages only
'
'           3-25-08 Hidden lines may include the orig line and the line that has been changed to show
'                   the increase or decrease of spots/$. (Change in $ shows a 2nd line on the printed
'                   contract for differences).  When a decrease in spots occurs, it results in zero grimps, grps, etc.
'                   The Orig line will have the grimps, grps, etc and both the orig and the zero values will
'                   be sent to gResearchTotals; which results in no avg rating when theres a decrease.
'                   For example:  Pkg line with 12 spots and 3 hidden lines, each with 4 spots in the week.
'                   The week is deleted.  Diff only version will show the hidden line with the orig 4 spots in each week
'                   then another line showing the decrease of the 4 spots. This is the line with no grps, cpp, grimps, etc.
'                   All of that is sent to gResearchTotals, which returns no avg rtg and avg aud; but other
'                   reserach info is returned.
Sub mBRPkgQtrTotals(ilPkgLineList() As Integer, ilVehList() As Integer, ilPkgVehList() As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVehicle                                                                             *
'******************************************************************************************
    Dim ilPkg As Integer
    'Dim ilQtr As Integer
    Dim ilTemp As Integer
    'ReDim llPkgWkGrp(1 To 14) As Long       '11-14-11
    ReDim llPkgWkGrp(0 To 14) As Long       '11-14-11. Index zero ignored
    Dim llLnSpots As Long
    'Dim ilRchQtr As Integer
    Dim llResearchPop As Long
    'Dim llTemp As Long
    Dim dlTemp As Double 'TTP 10439 - Rerate 21,000,000
    Dim llQtr As Long               '4-10-19 replace ilQtr due to subscript out of range
    Dim llRchQtr As Long            '4-10-19 replace ilRchQtr due to subscript out of range

    'For Packages Only........
    'Create array of grps, weekly rating, grimps and rates to pass to routine to generate qtrly totals for individual
    'line packages.  (combination of all hidden lines for 1 package line)
    For ilPkg = LBound(ilPkgLineList) To UBound(ilPkgLineList) - 1 Step 1
    For llQtr = 1 To imMaxQtrs Step 1
        ''ReDim tmPkLnRtg(1 To 1) As Integer
        'ReDim tmPkLnCost(1 To 1) As Long
        'ReDim tmPkLnGRP(1 To 1) As Long
        'ReDim tmPkLnGrimp(1 To 1) As Long
        ReDim tmPkLnCost(0 To 0) As Long
        ReDim tmPkLnGRP(0 To 0) As Long
        ReDim tmPkLnGrimp(0 To 0) As Long
        'For ilTemp = 1 To 13                    'init the weekly grps for this package for this qtr
        'ReDim llPkgWkGrp(1 To imWeeksPerQtr(ilQtr)) As Long
        ReDim llPkgWkGrp(0 To imWeeksPerQtr(llQtr)) As Long
        For ilTemp = LBONE To imWeeksPerQtr(llQtr)
            llPkgWkGrp(ilTemp) = 0
        Next ilTemp
        llLnSpots = 0
        If tgSpf.sDemoEstAllowed = "Y" Then     '6-1-04
            llResearchPop = -1
        End If
        'get each package vehicle for a quarter at a time (detail option)
        For llRchQtr = llQtr To UBound(tmLRch) - 1 Step 8
        If tmLRch(llRchQtr).iLineNo = ilPkgLineList(ilPkg) Then
            llLnSpots = llLnSpots + tmLRch(llRchQtr).lQSpots
        End If
        If tmLRch(llRchQtr).lQSpots <> 0 And tmLRch(llRchQtr).iPkLineNo = ilPkgLineList(ilPkg) Then
            'tmPkLnRtg(UBound(tmPkLnRtg)) = tmLRch(ilRchQtr).iTotalAvgRating
            If tgSpf.sDemoEstAllowed = "Y" Then         '6-1-04
                If llResearchPop = -1 Then
                    llResearchPop = tmLRch(llRchQtr).lSatelliteEst
                Else
                    If llResearchPop <> tmLRch(llRchQtr).lSatelliteEst And llResearchPop <> 0 Then
                        llResearchPop = 0
                    End If
                End If
            Else
                '11-11-05 use pkg pop by line (not by unique veh) for detail sch line research
                llResearchPop = lmPopPkgByLine(ilPkg)

                'For ilVehicle = 1 To UBound(ilVehList) - 1 Step 1
                '    llResearchPop = 0
                '    If ilVehList(ilVehicle) = ilPkgVehList(ilPkg) Then
                '        llResearchPop = lmPopPkg(ilVehicle)
                '        Exit For
                '    End If
               'Next ilVehicle
            End If
            'tmPkLnCost(UBound(tmPkLnCost)) = tmLRch(llRchQtr).lTotalCost
            tmPkLnCost(UBound(tmPkLnCost)) = tmLRch(llRchQtr).dTotalCost 'TTP 10439 - Rerate 21,000,000
            tmPkLnGRP(UBound(tmPkLnGRP)) = tmLRch(llRchQtr).lTotalGRP
            tmPkLnGrimp(UBound(tmPkLnGrimp)) = tmLRch(llRchQtr).lTotalGrimps
            ''ReDim Preserve tmPkLnRtg(1 To UBound(tmPkLnRtg) + 1) As Integer
            'ReDim Preserve tmPkLnCost(1 To UBound(tmPkLnCost) + 1) As Long
            'ReDim Preserve tmPkLnGRP(1 To UBound(tmPkLnGRP) + 1) As Long
            'ReDim Preserve tmPkLnGrimp(1 To UBound(tmPkLnGrimp) + 1) As Long
            ReDim Preserve tmPkLnCost(0 To UBound(tmPkLnCost) + 1) As Long
            ReDim Preserve tmPkLnGRP(0 To UBound(tmPkLnGRP) + 1) As Long
            ReDim Preserve tmPkLnGrimp(0 To UBound(tmPkLnGrimp) + 1) As Long
            'For ilTemp = 1 To 13    'gather the weekly grps for this package for this qtr
            For ilTemp = LBONE To imWeeksPerQtr(llQtr)      '11-14-11
                llPkgWkGrp(ilTemp) = llPkgWkGrp(ilTemp) + tmLRch(llRchQtr).lWklyGRP(ilTemp - 1)
            Next ilTemp
        End If
        Next llRchQtr               'loop to get next lines same qtr
        ''If UBound(tmPkLnRtg) > 1 Then
        'If UBound(tmPkLnCost) > 1 Then      '4/7/99 Avg Rating eliminated from computations
        If UBound(tmPkLnCost) > 0 Then      '4/7/99 Avg Rating eliminated from computations
        'dimensions must be exact sizes
        ''ReDim Preserve tmPkLnRtg(1 To UBound(tmPkLnRtg) - 1) As Integer
        'ReDim Preserve tmPkLnCost(1 To UBound(tmPkLnCost) - 1) As Long
        'ReDim Preserve tmPkLnGRP(1 To UBound(tmPkLnGRP) - 1) As Long
        'ReDim Preserve tmPkLnGrimp(1 To UBound(tmPkLnGrimp) - 1) As Long
        ReDim Preserve tmPkLnCost(0 To UBound(tmPkLnCost) - 1) As Long
        ReDim Preserve tmPkLnGRP(0 To UBound(tmPkLnGRP) - 1) As Long
        ReDim Preserve tmPkLnGrimp(0 To UBound(tmPkLnGrimp) - 1) As Long
        'gResearchTotals True, llResearchPop, tmPkLnCost(), tmPkLnRtg(), tmPkLnGrimp(), tmPkLnGRP(), llTemp, tmCbf.iAvgRate, tmCbf.lVQGrimp, tmCbf.lVQGRP, tmCbf.lVQCPP, tmCbf.lVQCPM
        '4/7/99
        'gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmPkLnCost(), tmPkLnGrimp(), tmPkLnGRP(), llLnSpots, llTemp, tmCbf.iAvgRate, tmCbf.lVQGrimp, tmCbf.lVQGRP, tmCbf.lVQCPP, tmCbf.lVQCPM, tmCbf.lAvgAud
        gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmPkLnCost(), tmPkLnGrimp(), tmPkLnGRP(), llLnSpots, dlTemp, tmCbf.iAvgRate, tmCbf.lVQGrimp, tmCbf.lVQGRP, tmCbf.lVQCPP, tmCbf.lVQCPM, tmCbf.lAvgAud 'TTP 10439 - Rerate 21,000,000
        End If
        'Research totals for a quarter for each hidden line of the package has been gathered
        'Place the totals into the package line
        For llRchQtr = llQtr To UBound(tmLRch) - 1 Step 8
        If tmLRch(llRchQtr).iLineNo = ilPkgLineList(ilPkg) And tmLRch(llRchQtr).lQSpots <> 0 Then
            tmLRch(llRchQtr).lTotalCPP = tmCbf.lVQCPP
            tmLRch(llRchQtr).lTotalCPM = tmCbf.lVQCPM
            tmLRch(llRchQtr).lTotalGRP = tmCbf.lVQGRP
            tmLRch(llRchQtr).lTotalGrimps = tmCbf.lVQGrimp
            tmLRch(llRchQtr).iTotalAvgRating = tmCbf.iAvgRate
            tmLRch(llRchQtr).lTotalAvgAud = tmCbf.lAvgAud       '4/7/99
            'tmLRch(llRchQtr).lTotalCost = llTemp
            tmLRch(llRchQtr).dTotalCost = dlTemp 'TTP 10439 - Rerate 21,000,000
            tmLRch(llRchQtr).lSatelliteEst = llResearchPop      '6-1-04

            'For ilTemp = 1 To 13    'gather the weekly grps for this package for this qtr
            For ilTemp = LBONE To imWeeksPerQtr(llQtr)       '11-14-11
                tmLRch(llRchQtr).lWklyGRP(ilTemp - 1) = llPkgWkGrp(ilTemp)
            Next ilTemp
            'the weekly totals only gets updated once for the package
            'For ilTemp = 1 To 13                    'init the weekly grps for this package for this qtr
            For ilTemp = LBONE To imWeeksPerQtr(llQtr)   '11-14-11
                llPkgWkGrp(ilTemp) = 0
            Next ilTemp
            'Exit For
        End If
        Next llRchQtr
    Next llQtr
    'ReDim Preserve tmPkLnQtrList(1 To UBound(tmPkLnQtrList) + 1) As VEHQTRLIST
    Next ilPkg
    
    Erase llPkgWkGrp                '11-4-13
End Sub

'
'           mBRPkgVehTotals ilPkgLineList, ilVehlist, ilPkgVehList
'           Get the package totals by line and place them into
'           the package lines buckets and show on detail line only
'
'
Sub mBRPkgVehTotals(ilPkgLineList() As Integer, ilVehList() As Integer, ilPkgVehList() As Integer)
    Dim ilPkg As Integer
    'Dim ilRchQtr As Integer
    'Dim ilQtr As Integer
    Dim llLnSpots As Long
    Dim ilFound As Integer
    Dim llResearchPop As Long
    Dim ilVehicle As Integer
    'Dim llTemp As Long
    Dim dlTemp As Double 'TTP 10439 - Rerate 21,000,000
    Dim ilWeekly As Integer
    Dim ilWeeklyMinusOne As Integer
    Dim ilUpperGrp As Integer
    Dim ilLoop As Integer
    'ReDim tmWkPkgVGrps(1 To 1) As WEEKLYGRPS
    ReDim tmWkPkgVGrps(0 To 0) As WEEKLYGRPS
    'Dim ilVaryingQtrs As Integer
    Dim llQtr As Long                   '4-10-19    replace ilQtr due to subscript out of range
    Dim llRchQtr As Long                '4-10-19    replace ilRchQtr
    Dim llVaryingQtrs As Long           '4-10-19    replace ilVaryingQtrs

    ilUpperGrp = UBound(tmWkPkgVGrps)   '1
    'For Packages Only (Get the totals by package line)........
    'Create array of grps, weekly rating, grimps and rates to pass to routine to generate qtrly totals for individual
    'line packages.  (combination of all hidden lines for 1 package line)
    For ilPkg = LBound(ilPkgLineList) To UBound(ilPkgLineList) - 1 Step 1
        llVaryingQtrs = 0       '11-14-11
        For llQtr = 1 To imMaxQtrs
            For ilWeekly = 1 To imWeeksPerQtr(llQtr)
                ilWeeklyMinusOne = ilWeekly - 1
            'For ilWeekly = 1 To 13
            'ReDim tmPkLnCost(1 To 1) As Long
            'ReDim tmPkLnGRP(1 To 1) As Long
            'ReDim tmPkLnGrimp(1 To 1) As Long
            ReDim tmPkLnCost(0 To 0) As Long
            ReDim tmPkLnGRP(0 To 0) As Long
            ReDim tmPkLnGrimp(0 To 0) As Long
            llLnSpots = 0
    
            For llRchQtr = llQtr To UBound(tmLRch) - 1 Step 8   '5-16-00
    
                If tmLRch(llRchQtr).iPkLineNo = ilPkgLineList(ilPkg) Then
                llLnSpots = llLnSpots + tmLRch(llRchQtr).lSpots(ilWeeklyMinusOne)
                For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1 Step 1
                    llResearchPop = 0
                    If ilVehList(ilVehicle) = ilPkgVehList(ilPkg) Then
                        If tgSpf.sDemoEstAllowed = "Y" Then
                            llResearchPop = -1
                            Exit For
                        Else
                            'llResearchPop = lmPopPkg(ilVehicle)     '11-24-04   ; removed 3-15-19 , and replace with below
                            llResearchPop = lmPopPkgByLine(ilPkg)           '3-15-19
                            Exit For
                        End If
                    End If
                Next ilVehicle
                'look only for the matching hidden line
                If (tmLRch(llRchQtr).lSpots(ilWeeklyMinusOne)) <> 0 And tmLRch(llRchQtr).iPkLineNo = ilPkgLineList(ilPkg) Then
    
                    tmPkLnCost(UBound(tmPkLnCost)) = tmLRch(llRchQtr).lRates(ilWeeklyMinusOne)
                    tmPkLnGRP(UBound(tmPkLnGRP)) = tmLRch(llRchQtr).lWklyGRP(ilWeeklyMinusOne)
                    tmPkLnGrimp(UBound(tmPkLnGrimp)) = tmLRch(llRchQtr).lWklyGrimp(ilWeeklyMinusOne)
                    If tgSpf.sDemoEstAllowed = "Y" Then
                        If llResearchPop = -1 Then
                            llResearchPop = tmLRch(llRchQtr).lPopEst(ilWeeklyMinusOne)
                        Else
                            If llResearchPop <> tmLRch(llRchQtr).lPopEst(ilWeeklyMinusOne) And llResearchPop <> 0 Then
                                llResearchPop = 0
                            End If
                        End If
                    End If
                    'ReDim Preserve tmPkLnCost(1 To UBound(tmPkLnCost) + 1) As Long
                    'ReDim Preserve tmPkLnGRP(1 To UBound(tmPkLnGRP) + 1) As Long
                    'ReDim Preserve tmPkLnGrimp(1 To UBound(tmPkLnGrimp) + 1) As Long
                    ReDim Preserve tmPkLnCost(0 To UBound(tmPkLnCost) + 1) As Long
                    ReDim Preserve tmPkLnGRP(0 To UBound(tmPkLnGRP) + 1) As Long
                    ReDim Preserve tmPkLnGrimp(0 To UBound(tmPkLnGrimp) + 1) As Long
                End If
                End If
            Next llRchQtr
            'If UBound(tmPkLnCost) > 1 Then      '4/7/99 Avg Rating eliminated from computations
            If UBound(tmPkLnCost) > 0 Then      '4/7/99 Avg Rating eliminated from computations
                'dimensions must be exact sizes
                ''ReDim Preserve tmPkLnRtg(1 To UBound(tmPkLnRtg) - 1) As Integer
                'ReDim Preserve tmPkLnCost(1 To UBound(tmPkLnCost) - 1) As Long
                'ReDim Preserve tmPkLnGRP(1 To UBound(tmPkLnGRP) - 1) As Long
                'ReDim Preserve tmPkLnGrimp(1 To UBound(tmPkLnGrimp) - 1) As Long
                ReDim Preserve tmPkLnCost(0 To UBound(tmPkLnCost) - 1) As Long
                ReDim Preserve tmPkLnGRP(0 To UBound(tmPkLnGRP) - 1) As Long
                ReDim Preserve tmPkLnGrimp(0 To UBound(tmPkLnGrimp) - 1) As Long
                'ilRchQtr = (ilQtr - 1) * 13 + ilWeekly
                llRchQtr = llVaryingQtrs + ilWeekly         '11-14-11
                'only care about the weeks grps
                ilFound = False
                For ilLoop = LBound(tmWkPkgVGrps) To UBound(tmWkPkgVGrps) - 1
                    If tmWkPkgVGrps(ilLoop).iVefCode = ilPkgLineList(ilPkg) Then
                        ilFound = True
                        Exit For
                    End If
                Next ilLoop
                If ilFound Then
                '4/7/99
                    'gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmPkLnCost(), tmPkLnGrimp(), tmPkLnGRP(), llLnSpots, llTemp, tmCbf.iAvgRate, tmCbf.lVQGrimp, tmWkPkgVGrps(ilLoop).lGrps(llRchQtr), tmCbf.lVQCPP, tmCbf.lVQCPM, tmCbf.lAvgAud
                    gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmPkLnCost(), tmPkLnGrimp(), tmPkLnGRP(), llLnSpots, dlTemp, tmCbf.iAvgRate, tmCbf.lVQGrimp, tmWkPkgVGrps(ilLoop).lGrps(llRchQtr), tmCbf.lVQCPP, tmCbf.lVQCPM, tmCbf.lAvgAud 'TTP 10439 - Rerate 21,000,000
                Else
                    tmWkPkgVGrps(ilUpperGrp).iVefCode = ilPkgLineList(ilPkg)
                    'gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmPkLnCost(), tmPkLnGrimp(), tmPkLnGRP(), llLnSpots, llTemp, tmCbf.iAvgRate, tmCbf.lVQGrimp, tmWkPkgVGrps(ilUpperGrp).lGrps(llRchQtr), tmCbf.lVQCPP, tmCbf.lVQCPM, tmCbf.lAvgAud
                    gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmPkLnCost(), tmPkLnGrimp(), tmPkLnGRP(), llLnSpots, dlTemp, tmCbf.iAvgRate, tmCbf.lVQGrimp, tmWkPkgVGrps(ilUpperGrp).lGrps(llRchQtr), tmCbf.lVQCPP, tmCbf.lVQCPM, tmCbf.lAvgAud 'TTP 10439 - Rerate 21,000,000
                    ilUpperGrp = UBound(tmWkPkgVGrps) + 1
                    'ReDim Preserve tmWkPkgVGrps(1 To ilUpperGrp) As WEEKLYGRPS
                    ReDim Preserve tmWkPkgVGrps(0 To ilUpperGrp) As WEEKLYGRPS
                End If
            End If
            'ReDim Preserve tmPkLnQtrList(1 To UBound(tmPkLnQtrList) + 1) As VEHQTRLIST
            Next ilWeekly
            llVaryingQtrs = llVaryingQtrs + imWeeksPerQtr(llQtr)        '11-14-11
        Next llQtr
    Next ilPkg
End Sub

'
'
'               mBuildLRch - Create the current versions line research data based on unique rate & DP or
'                           Create the previous versions line research data
'               <output> tlRch() - Current research array (tmLRch) or Previous research array (tmPrevLRch)
'               <input> llTemp - index into the saved lines research data
'                       ilSavePkVgh - Package vehicle code
'                       llCntGrimps - Total Gross impressions accumulated
'                       llPop - population
'                       ilQtr - quarter that is beging calculated  2=20-00
'       4-10-19 iltemp changed to long llTemp due to subscript out of range
'Sub mBuildLRch(tlLRch() As RESEARCHINFO, llTemp As Long, ilSavePkVeh As Integer, llCntGrimps As Long, llPop As Long, ilQtr As Long)
Sub mBuildLRch(tlLRch() As RESEARCHINFO, llTemp As Long, ilSavePkVeh As Integer, llCntGrimps As Long, llPop As Long)        'remove ilQtr parameter
    'Dim llTotalCost As Long
    Dim dlTotalCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim llPopEst As Long

    If tmClf.sType <> "A" And tmClf.sType <> "O" And tmClf.sType <> "E" Then
        'gAvgAudToLnResearch sm1or2PlaceRating, True, llPop, tlLRch(llTemp).lPopEst(), tlLRch(llTemp).lSpots(), tlLRch(llTemp).lRates(), tlLRch(llTemp).lAvgAud(), llTotalCost, tmCbf.lAvgAud, tlLRch(llTemp).iWklyRating(), tmCbf.iAvgRate, tlLRch(llTemp).lWklyGrimp(), tmCbf.lGrImp, tlLRch(llTemp).lWklyGRP(), tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, llPopEst
        gAvgAudToLnResearch sm1or2PlaceRating, True, llPop, tlLRch(llTemp).lPopEst(), tlLRch(llTemp).lSpots(), tlLRch(llTemp).lRates(), tlLRch(llTemp).lAvgAud(), dlTotalCost, tmCbf.lAvgAud, tlLRch(llTemp).iWklyRating(), tmCbf.iAvgRate, tlLRch(llTemp).lWklyGrimp(), tmCbf.lGrImp, tlLRch(llTemp).lWklyGRP(), tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
        '2-20-00
        'for hiddenlines only, gather the weekly grps for its vehicle totals.  the hidden
        'lines are not grouped with conventional of the same vehicle
        '1-24-03 remove processing of hidden lines here
        'If tmClf.sType = "H" Then
        '   ilFound = False
        '   For ilLoop = 1 To UBound(tmWkHiddenVGrps) - 1
        '       If tmClf.iLine = tmWkHiddenVGrps(ilLoop).iVefCode Then   'the line # is substituted in place of vehicle code for hidden info
        '       ilFound = True
        '       Exit For
        '       End If
        '   Next ilLoop
        '   ilWeekInx = (ilQtr - 1) * 13
        '   If ilFound Then
        '       For ilWeek = 1 To 13
        '       tmWkHiddenVGrps(ilLoop).lGrps(ilWeekInx + ilWeek) = tlLRch(ilTemp).lWklyGRP(ilWeek)
        '        Next ilWeek
        '    Else
        '        For ilWeek = 1 To 13
        '        tmWkHiddenVGrps(UBound(tmWkHiddenVGrps)).lGrps(ilWeekInx + ilWeek) = tlLRch(ilTemp).lWklyGRP(ilWeek)
        '        Next ilWeek
        '        tmWkHiddenVGrps(UBound(tmWkHiddenVGrps)).iVefCode = tmClf.iLine
        '        ReDim Preserve tmWkHiddenVGrps(1 To UBound(tmWkHiddenVGrps) + 1) As WEEKLYGRPS
        '    End If
        'End If

    Else
        'llTotalCost = 0
        dlTotalCost = 0 'TTP 10439 - Rerate 21,000,000
        tmCbf.lAvgAud = 0
        tmCbf.iAvgRate = 0
        tmCbf.lGrImp = 0
        tmCbf.lGRP = 0
        tmCbf.lCPP = 0
        tmCbf.lCPM = 0
        llPopEst = 0            '6-1-04
    End If
    'Update array with results of research routine (didn't use the actual field names in the subroutine call because
    'the total parameters were too long
    tlLRch(llTemp).iVefCode = tmClf.iVefCode
    tlLRch(llTemp).iLineNo = tmClf.iLine
    tlLRch(llTemp).sType = tmClf.sType          's = std, O = order, a=air, H = hidden
    tlLRch(llTemp).iPkLineNo = tmClf.iPkLineNo  'pkg line reference if hidden
    tlLRch(llTemp).iPkvefCode = ilSavePkVeh     'associ pkg vehicle code if hidden line
    'tlLRch(llTemp).lTotalCost = llTotalCost
    tlLRch(llTemp).dTotalCost = dlTotalCost 'TTP 10439 - Rerate 21,000,000
    tlLRch(llTemp).lTotalCPP = tmCbf.lCPP
    tlLRch(llTemp).lTotalCPM = tmCbf.lCPM
    tlLRch(llTemp).lTotalGRP = tmCbf.lGRP
    tlLRch(llTemp).lTotalGrimps = tmCbf.lGrImp
    tlLRch(llTemp).iTotalAvgRating = tmCbf.iAvgRate
    tlLRch(llTemp).lTotalAvgAud = tmCbf.lAvgAud
    tlLRch(llTemp).lSatelliteEst = llPopEst             '6-1-04
    'ReDim Preserve tlLRch(1 To UBound(tlLRch) + 1)
    ReDim Preserve tlLRch(0 To UBound(tlLRch) + 1)
    llCntGrimps = llCntGrimps + tmCbf.lGrImp              'accum contracts total grimps so that % distrib on each line can be calculated in BR
End Sub

'
'
'             mFindMaxDates - Determine the earliest and latest dates of the contract
'                   header.  If Differences option, see if previous revision has an earlier
'                   start date and/or later end date
'
'           <input> slStartDate - Contract header start date
'                   slEnd Date - contract header end date
'                   ilShowStdQtr - true if showing quarters by std months
'                   ilDefaultToQtr - true if start date should be on qtr (for BR, else for Order Audit get
'                   the start month which doesnt have to be on a qtr
'           <output> llchfStart - Start date of quarter (based on contract header start date)
'                   llchfEnd - End date of quarter (based on contract header end date)
'                   ilCurrTotalMonths - Total months of order
'                   llStdStartDates - start dates of each std month (to summarize $ and spots)
'                   ilYear - corp or std Year that this contract belongs in
'                   ilStartMonth - Start month of corp or std year that this contracts starts in
'
'
' COMMENT OUT, Changed to Global Subroutine :  gFindMaxDates
'Sub mFindMaxDates(slStartDate As String, slEndDate As String, llChfStart As Long, llChfEnd As Long, ilCurrStartQtr() As Integer, ilCurrTotalMonths As Integer, llStdStartDates() As Long, ilShowStdQtr As Integer, ilYear As Integer, ilStartMonth As Integer, ilDefaultToQtr As Integer)
'Dim llStartQtr As Long
'Dim llEndQtr As Long
'Dim slDay As String
'Dim ilLoop3 As Integer
'Dim llFltStart As Long
'Dim slStr As String
'Dim ilLoop As Integer
'    If ilShowStdQtr Then
'        'determine current qtr start date
'        llStartQtr = gDateValue(gObtainStartStd(slStartDate))    'get the std start date of the contract
'        llEndQtr = gDateValue(gObtainEndStd(slEndDate))       'get the std end date of the contract
'        'Save the earliest/latest dates from contr header
'        llChfStart = gDateValue(slStartDate)
'        llChfEnd = gDateValue(slEndDate)
'        'Calculate the total number of std airng months for this order
'        'ilCurrStartQtr(0) = 0                       'init to place start date of starting qtr of order
'        'ilCurrStartQtr(1) = 0
'        'ilCurrTotalMonths = 0
'        slDay = gObtainEndStd(Format$(llStartQtr, "m/d/yy"))
'
'        If ilDefaultToQtr = True Then       'earliest date must fall on a qtr start date
'            ilLoop3 = Month(Format$(gDateValue(slDay), "m/d/yy"))
'            'Calculate the first qtr, backup months until start of a qtr month
'            Do While ilLoop3 <> 1 And ilLoop3 <> 4 And ilLoop3 <> 7 And ilLoop3 <> 10
'                slStr = gObtainStartStd(slDay)      'start date of month to backup to previous month
'                llFltStart = gDateValue(slStr) - 1
'                slDay = gObtainEndStd(Format$(llFltStart, "m/d/yy"))
'                ilLoop3 = Month(Format$(gDateValue(slDay), "m/d/yy"))
'            Loop
'        End If
'        'slDay contains start of contract's qtr
'        gPackDate slDay, ilCurrStartQtr(0), ilCurrStartQtr(1)   'save this contracts start month to store in cbf
'        llStartQtr = gDateValue(slDay)
'        Do While llStartQtr <= llEndQtr
'            slStr = gObtainEndStd(Format$(llStartQtr, "m/d/yy"))
'            'Determine what the starting qtr and month is for this order
'            ilLoop3 = Month(Format$(gDateValue(slStr), "m/d/yy"))
'            llStartQtr = gDateValue(slStr) + 1
'            ilCurrTotalMonths = ilCurrTotalMonths + 1         'accum total # of airing std months   (to be stored in cbf)
'        Loop
'        'Calc # std month start and date dates to total the week into
'        'build array of 13 start standard dates
'        For ilLoop = 1 To 37 Step 1         '12-29-06 chged to do 3 yrs for tax creation
'            slDay = gObtainStartStd(slDay)
'            llStdStartDates(ilLoop) = gDateValue(slDay)
'            slDay = gObtainEndStd(slDay)
'            llStartQtr = gDateValue(slDay) + 1                      'increment for next month
'            slDay = Format$(llStartQtr, "m/d/yy")
'        Next ilLoop
'    Else
'        'determine current qtr start date
'        llStartQtr = gDateValue(gObtainStartCorp(slStartDate, False))   'get the std start date of the contract
'        llEndQtr = gDateValue(gObtainEndCorp(slEndDate, False))      'get the std end date of the contract
'        'Save the earliest/latest dates from contr header
'        llChfStart = gDateValue(slStartDate)
'        llChfEnd = gDateValue(slEndDate)
'
'        'slDay = gObtainEndCorp(Format$(llStartQtr + 14, "m/d/yy"), False)      'get to middle of the month to find its true month  #
'        slDay = Format$(llStartQtr + 14, "m/d/yy")
'        ilLoop3 = Month(Format$(gDateValue(slDay), "m/d/yy"))
'        'Calculate the first qtr, backup months until start of a qtr month
'        Do While ilLoop3 <> 1 And ilLoop3 <> 4 And ilLoop3 <> 7 And ilLoop3 <> 10
'            slStr = gObtainStartCorp(slDay, False)     'start date of month to backup to previous month
'            'llFltStart = gDateValue(slStr) - 14
'            slDay = Format$((gDateValue(slStr) - 14), "m/d/yy")
'            'slDay = gObtainEndCorp(Format$(llFltStart, "m/d/yy"), False)
'            ilLoop3 = Month(Format$(gDateValue(slDay), "m/d/yy"))
'        Loop
'        'slDay contains start of contract's qtr
'
'        gPackDate slDay, ilCurrStartQtr(0), ilCurrStartQtr(1)   'save this contracts start month to store in cbf
'        slDay = gObtainStartCorp(slDay, False)
'        llStartQtr = gDateValue(slDay)
'        Do While llStartQtr <= llEndQtr
'            slStr = gObtainEndCorp(Format$(llStartQtr, "m/d/yy"), False)
'            'Determine what the starting qtr and month is for this order
'            ilLoop3 = Month(Format$(gDateValue(slStr), "m/d/yy"))
'            llStartQtr = gDateValue(slStr) + 1
'            ilCurrTotalMonths = ilCurrTotalMonths + 1         'accum total # of airing std months   (to be stored in cbf)
'        Loop
'
'        'Calc # std month start and date dates to total the week into
'        'build array of 13 start corp dates
'        For ilLoop = 1 To 37 Step 1
'            slDay = gObtainStartCorp(slDay, False)
'            llStdStartDates(ilLoop) = gDateValue(slDay)
'            slDay = gObtainEndCorp(slDay, False)
'            llStartQtr = gDateValue(slDay) + 1                      'increment for next month
'            slDay = Format$(llStartQtr, "m/d/yy")
'        Next ilLoop
'    End If
'    slStr = Format$(llStdStartDates(1), "m/d/yy")
'    mGetYearStartMo ilShowStdQtr, slStr, ilYear, ilStartMonth
'End Sub
'
'
'           mGetCntWkGrimps - Calculate the grimps for each week individually
'           for contract totals
'           <input> llResearchPop - population
'           <output> 104 weekly grimps for all lines
'           2-11-00
'
Sub mGetCntWkGrps(llInResearchPop As Long)
    'Dim ilQtr As Integer
    Dim ilWeekly As Integer
    Dim ilWeeklyMinusOne As Integer
    'Dim ilRchQtr As Integer
    Dim llLnSpots As Long
    Dim llCPP As Long
    Dim llCPM As Long
    Dim llAvgAud As Long
    'Dim llTotalCost As Long
    Dim dlTotalCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim llGrimps As Long
    Dim llResearchPop As Long           '6-29-04
    'Dim ilVaryingQtrs As Integer
    Dim llQtr As Long                   '4-10-19    replace ilQtr due to subscript out of range
    Dim llRchQtr As Long                '4-10-19    replace ilRchQtr
    Dim llVaryingQtrs As Long           '4-10-19    replace ilVaryingQtrs
    '2-11-00 Build weekly grimps for all vehicles.  Each line entry is 8 quarters (max 2 years).  Entries are unique
    'for  vehicle/daypart/#spots/line rate
    
    'ReDim tmWkCntGrps(1 To 1) As WEEKLYGRPS
    ReDim tmWkCntGrps(0 To 0) As WEEKLYGRPS
    llVaryingQtrs = 0               '11-14-11
    For llQtr = 1 To imMaxQtrs Step 1              'cycle thru 8 quarters
        'For ilWeekly = 1 To 13              'each qtr has 13 weeks stored
        For ilWeekly = 1 To imWeeksPerQtr(llQtr)
            ilWeeklyMinusOne = ilWeekly - 1
            'ReDim tmVRtg(1 To 1) As Integer
            'ReDim tmVCost(1 To 1) As Long
            'ReDim tmVGRP(1 To 1) As Long
            'ReDim tmVGrimp(1 To 1) As Long
            ReDim tmVRtg(0 To 0) As Integer
            ReDim tmVCost(0 To 0) As Long
            ReDim tmVGRP(0 To 0) As Long
            ReDim tmVGrimp(0 To 0) As Long

            llLnSpots = 0
            If tgSpf.sDemoEstAllowed = "Y" Then
                llResearchPop = -1
            Else
                llResearchPop = llInResearchPop
            End If
            For llRchQtr = llQtr To UBound(tmLRch) - 1 Step 8   'process as many lines that are built
            'bypass the Package lines, the quarterly totals will be obtained from the individual hidden lines because
            'each package line can contain the same data repeated in each flight.  Exclude package lines to avoid
            'duplicating results.

            If (tmLRch(llRchQtr).sType <> "A" And tmLRch(llRchQtr).sType <> "O" And tmLRch(llRchQtr).sType <> "E") Then    'not a package of any kind
                llLnSpots = llLnSpots + tmLRch(llRchQtr).lSpots(ilWeeklyMinusOne)
                If (tmLRch(llRchQtr).lSpots(ilWeeklyMinusOne) <> 0) Then              'only process if spots exist in the qtr
               'tmVRtg(UBound(tmVRtg)) = tmLRch(ilRchQtr).lAvgAud(ilWeekly)
                tmVRtg(UBound(tmVRtg)) = tmLRch(llRchQtr).iWklyRating(ilWeeklyMinusOne) '2-17-00
                tmVCost(UBound(tmVCost)) = tmLRch(llRchQtr).lRates(ilWeeklyMinusOne)
                tmVGRP(UBound(tmVGRP)) = tmLRch(llRchQtr).lWklyGRP(ilWeeklyMinusOne)
                tmVGrimp(UBound(tmVGrimp)) = tmLRch(llRchQtr).lWklyGrimp(ilWeeklyMinusOne)
                If tgSpf.sDemoEstAllowed = "Y" Then
                    If llResearchPop = -1 Then
                        llResearchPop = tmLRch(llRchQtr).lPopEst(ilWeeklyMinusOne)
                    Else
                        If llResearchPop <> tmLRch(llRchQtr).lPopEst(ilWeeklyMinusOne) And llResearchPop <> 0 Then
                            llResearchPop = 0
                        End If
                    End If
                End If
                'ReDim Preserve tmVRtg(1 To UBound(tmVRtg) + 1) As Integer
                'ReDim Preserve tmVCost(1 To UBound(tmVCost) + 1) As Long
                'ReDim Preserve tmVGRP(1 To UBound(tmVGRP) + 1) As Long
                'ReDim Preserve tmVGrimp(1 To UBound(tmVGrimp) + 1) As Long
                ReDim Preserve tmVRtg(0 To UBound(tmVRtg) + 1) As Integer
                ReDim Preserve tmVCost(0 To UBound(tmVCost) + 1) As Long
                ReDim Preserve tmVGRP(0 To UBound(tmVGRP) + 1) As Long
                ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) + 1) As Long
                End If
            End If
            Next llRchQtr
            'If UBound(tmVRtg) > 1 Then
            If UBound(tmVRtg) > 0 Then
                'dimensions must be exact sizes
                'ReDim Preserve tmVRtg(1 To UBound(tmVRtg) - 1) As Integer
                'ReDim Preserve tmVCost(1 To UBound(tmVCost) - 1) As Long
                'ReDim Preserve tmVGRP(1 To UBound(tmVGRP) - 1) As Long
                'ReDim Preserve tmVGrimp(1 To UBound(tmVGrimp) - 1) As Long
                ReDim Preserve tmVRtg(0 To UBound(tmVRtg) - 1) As Integer
                ReDim Preserve tmVCost(0 To UBound(tmVCost) - 1) As Long
                ReDim Preserve tmVGRP(0 To UBound(tmVGRP) - 1) As Long
                ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) - 1) As Long
                'ilRchQtr = (ilQtr - 1) * 13 + ilWeekly
                llRchQtr = llVaryingQtrs + ilWeekly         '11-14-11
                'only care about the weeks grimps
                'gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, llTotalCost, tmCbf.iAvgRate, llGrimps, tmWkCntGrps(1).lGrps(ilRchQtr), llCPP, llCPM, llAvgAud
                'gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, llTotalCost, tmCbf.iAvgRate, llGrimps, tmWkCntGrps(0).lGrps(llRchQtr), llCPP, llCPM, llAvgAud
                gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, dlTotalCost, tmCbf.iAvgRate, llGrimps, tmWkCntGrps(0).lGrps(llRchQtr), llCPP, llCPM, llAvgAud 'TTP 10439 - Rerate 21,000,000
            End If
        Next ilWeekly
        llVaryingQtrs = llVaryingQtrs + imWeeksPerQtr(llQtr) 'accumulate # weeks in qtr to get the index into week processing, no longer hard coded 13 weeks/qtr        '11-14-11
    Next llQtr
End Sub

'
Sub mGetCntWkVGrps(llInResearchPop As Long, ilVehList() As Integer)
    'Dim ilQtr As Integer
    Dim ilWeekly As Integer
    Dim ilWeeklyMinusOne As Integer
    'Dim ilRchQtr As Integer
    Dim llLnSpots As Long
    Dim llCPP As Long
    Dim llCPM As Long
    Dim llAvgAud As Long
    'Dim llTotalCost As Long
    Dim dlTotalCost As Double 'TTP 10439 - Rerate 21,000,000
    Dim llGrimps As Long
    Dim ilVehicle As Integer
    Dim ilUpperGrp As Integer
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim llResearchPop As Long       '6-29-04
    'Dim ilVaryingQtrs As Integer
    Dim llQtr As Long                   '4-10-19    replace ilQtr due to subscript out of range
    Dim llRchQtr As Long                '4-10-19    replace ilRchQtr
    Dim llVaryingQtrs As Long           '4-10-19    replace ilVaryingQtrs

    '2-11-00 Build weekly grps for all vehicles.  Each line entry is 8 quarters (max 2 years).  Entries are unique
    'for  vehicle/daypart/#spots/line rate
    'ReDim tmWkCntVGrps(1 To 1) As WEEKLYGRPS
    'ilUpperGrp = 1
    ReDim tmWkCntVGrps(0 To 0) As WEEKLYGRPS
    ilUpperGrp = UBound(tmWkCntVGrps)   '1
    For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1 Step 1
        llVaryingQtrs = 0
        For llQtr = 1 To imMaxQtrs Step 1              'cycle thru 8 quarters
            'For ilWeekly = 1 To 13              'each qtr has 13 weeks stored
            For ilWeekly = 1 To imWeeksPerQtr(llQtr)
                ilWeeklyMinusOne = ilWeekly - 1
                'ReDim tmVRtg(1 To 1) As Integer
                'ReDim tmVCost(1 To 1) As Long
                'ReDim tmVGRP(1 To 1) As Long
                'ReDim tmVGrimp(1 To 1) As Long
                ReDim tmVRtg(0 To 0) As Integer
                ReDim tmVCost(0 To 0) As Long
                ReDim tmVGRP(0 To 0) As Long
                ReDim tmVGrimp(0 To 0) As Long
        
                llLnSpots = 0
                If tgSpf.sDemoEstAllowed = "Y" Then
                    llResearchPop = -1
                Else
                    llResearchPop = llInResearchPop
                End If
                For llRchQtr = llQtr To UBound(tmLRch) - 1 Step 8   'process as many lines that are built
                    'bypass the Package lines, the quarterly totals will be obtained from the individual hidden lines because
                    'each package line can contain the same data repeated in each flight.  Exclude package lines to avoid
                    'duplicating results.
        
                    If (tmLRch(llRchQtr).sType <> "A" And tmLRch(llRchQtr).sType <> "O" And tmLRch(llRchQtr).sType <> "E") And tmLRch(llRchQtr).iVefCode = ilVehList(ilVehicle) Then    'not a package of any kind
                        If tmLRch(llRchQtr).sType = "S" Then            '8-3-07 bypass the hidden here, dont want to duplicate research for the week
                            llLnSpots = llLnSpots + tmLRch(llRchQtr).lSpots(ilWeeklyMinusOne)
                            If (tmLRch(llRchQtr).lSpots(ilWeeklyMinusOne) <> 0) Then              'only process if spots exist in the qtr
                                'tmVRtg(UBound(tmVRtg)) = tmLRch(ilRchQtr).lAvgAud(ilWeekly)
                                tmVRtg(UBound(tmVRtg)) = tmLRch(llRchQtr).iWklyRating(ilWeeklyMinusOne) '2-17-00
                                tmVCost(UBound(tmVCost)) = tmLRch(llRchQtr).lRates(ilWeeklyMinusOne)
                                tmVGRP(UBound(tmVGRP)) = tmLRch(llRchQtr).lWklyGRP(ilWeeklyMinusOne)
                                tmVGrimp(UBound(tmVGrimp)) = tmLRch(llRchQtr).lWklyGrimp(ilWeeklyMinusOne)
                                If tgSpf.sDemoEstAllowed = "Y" Then
                                    If llResearchPop = -1 Then
                                        llResearchPop = tmLRch(llRchQtr).lPopEst(ilWeeklyMinusOne)
                                    Else
                                        If llResearchPop <> tmLRch(llRchQtr).lPopEst(ilWeeklyMinusOne) And llResearchPop <> 0 Then
                                            llResearchPop = 0
                                        End If
                                    End If
                                End If
                                'ReDim Preserve tmVRtg(1 To UBound(tmVRtg) + 1) As Integer
                                'ReDim Preserve tmVCost(1 To UBound(tmVCost) + 1) As Long
                                'ReDim Preserve tmVGRP(1 To UBound(tmVGRP) + 1) As Long
                                'ReDim Preserve tmVGrimp(1 To UBound(tmVGrimp) + 1) As Long
                                ReDim Preserve tmVRtg(0 To UBound(tmVRtg) + 1) As Integer
                                ReDim Preserve tmVCost(0 To UBound(tmVCost) + 1) As Long
                                ReDim Preserve tmVGRP(0 To UBound(tmVGRP) + 1) As Long
                                ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) + 1) As Long
                            End If
                        End If
                    End If
                Next llRchQtr
                'If UBound(tmVRtg) > 1 Then
                If UBound(tmVRtg) > 0 Then
                    'dimensions must be exact sizes
                    'ReDim Preserve tmVRtg(1 To UBound(tmVRtg) - 1) As Integer
                    'ReDim Preserve tmVCost(1 To UBound(tmVCost) - 1) As Long
                    'ReDim Preserve tmVGRP(1 To UBound(tmVGRP) - 1) As Long
                    'ReDim Preserve tmVGrimp(1 To UBound(tmVGrimp) - 1) As Long
                    ReDim Preserve tmVRtg(0 To UBound(tmVRtg) - 1) As Integer
                    ReDim Preserve tmVCost(0 To UBound(tmVCost) - 1) As Long
                    ReDim Preserve tmVGRP(0 To UBound(tmVGRP) - 1) As Long
                    ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) - 1) As Long
                    'ilRchQtr = (ilQtr - 1) * 13 + ilWeekly
                    llRchQtr = llVaryingQtrs + ilWeekly         '11-14-11
                    'only care about the weeks grps
                    ilFound = False
                    For ilLoop = LBound(tmWkCntVGrps) To UBound(tmWkCntVGrps) - 1
                    If tmWkCntVGrps(ilLoop).iVefCode = ilVehList(ilVehicle) Then
                        ilFound = True
                        Exit For
                    End If
                    Next ilLoop
                    If ilFound Then
                        'gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, llTotalCost, tmCbf.iAvgRate, llGrimps, tmWkCntVGrps(ilLoop).lGrps(llRchQtr), llCPP, llCPM, llAvgAud
                        gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, dlTotalCost, tmCbf.iAvgRate, llGrimps, tmWkCntVGrps(ilLoop).lGrps(llRchQtr), llCPP, llCPM, llAvgAud 'TTP 10439 - Rerate 21,000,000
                    Else
                        tmWkCntVGrps(ilUpperGrp).iVefCode = ilVehList(ilVehicle)
                        'tmWkCntVGrps(ilUpperGrp).iQtr = ilQtr
                        'gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, llTotalCost, tmCbf.iAvgRate, llGrimps, tmWkCntVGrps(ilUpperGrp).lGrps(llRchQtr), llCPP, llCPM, llAvgAud
                        gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, dlTotalCost, tmCbf.iAvgRate, llGrimps, tmWkCntVGrps(ilUpperGrp).lGrps(llRchQtr), llCPP, llCPM, llAvgAud 'TTP 10439 - Rerate 21,000,000
                        ilUpperGrp = UBound(tmWkCntVGrps) + 1
                        'ReDim Preserve tmWkCntVGrps(1 To ilUpperGrp) As WEEKLYGRPS
                        ReDim Preserve tmWkCntVGrps(0 To ilUpperGrp) As WEEKLYGRPS
                    End If
                End If
            Next ilWeekly
            llVaryingQtrs = llVaryingQtrs + imWeeksPerQtr(llQtr)  'accumulate # weeks in qtr to get the index into week processing, no longer hard coded 13 weeks/qtr       '11-14-11
        Next llQtr
    Next ilVehicle
    Erase tmVRtg
    Erase tmVCost
    Erase tmVGRP
    Erase tmVGrimp
End Sub

Sub mGetComment(lmOutComment As Long, lmInComment As Long, ilPropOrOrder As Integer)
    Dim ilRet As Integer
    'Setup comment pointers only if show = yes for each applicable comment
    lmOutComment = 0                     'assume no "comment
    tmCxfSrchKey.lCode = lmInComment      'comment  code
    imCxfRecLen = Len(tmCxf)
    ilRet = btrGetEqual(hmCxf, tmCxf, imCxfRecLen, tmCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching comment recd
    If ilRet = BTRV_ERR_NONE Then
        If ilPropOrOrder Then                        'proposal
        'If RptSelCt!rbcSelCInclude(0).Value Then         'proposal
            If tmCxf.sShProp = "Y" Then         'show comment on proposal
                lmOutComment = lmInComment
            End If
        Else                                    'order/hold
            If tmCxf.sShOrder = "Y" Then        'show it on order
                lmOutComment = lmInComment
            End If
        End If
    End If
End Sub

'
'           mGetYearStartMo - get the Year and Starting Month of the Corp
'               calendar or standard bdcst year based on any date
'
'           <input> slInpDate - Date string to determine year and start month
'                   ilShowStdQtr - true if using std qtrs, else false; 12-2020- changed to 0 = std, 1 = cal, 2 = corp
'           <output> ilYear -  year of corp calendar or std bdcst year
'                   ilStartMonth - start month of corp cal or std bdcst
'
'Sub mGetYearStartMo(ilShowStdQtr As Integer, slInpDate As String, ilYear As Integer, ilStartMonth As Integer)
'Dim ilLoop As Integer
'Dim llStartDate As Long
'Dim llEndDate As Long
'Dim llInpDate As Long
'Dim slTempDate As String
'Dim slYear As String
'Dim slMonth As String
'Dim slDay As String
''    If ilShowStdQtr Then                'using the std bdcst month for output
'      If ilShowStdQtr = 0 Then          '1-7-21
'        ilStartMonth = 1                'standard always starts with Jan
'        slTempDate = gObtainEndStd(slInpDate)        'get std bdcst end date
'        gObtainYearMonthDayStr slTempDate, True, slYear, slMonth, slDay
'        ilYear = Val(slYear)
'        Exit Sub
'    ElseIf ilShowStdQtr = 1 Then            '1-7-21  calendar
'        ilStartMonth = 1                'standard always starts with Jan
'        slTempDate = gObtainEndStd(slInpDate)        'get std bdcst end date
'        gObtainYearMonthDayStr slTempDate, True, slYear, slMonth, slDay
'        ilYear = Val(slYear)
'        Exit Sub
'    Else
'        llInpDate = gDateValue(slInpDate)
'        For ilLoop = LBound(tgMCof) To UBound(tgMCof) - 1 Step 1
'            gUnpackDateLong tgMCof(ilLoop).iStartDate(0, 0), tgMCof(ilLoop).iStartDate(1, 0), llStartDate
'            gUnpackDateLong tgMCof(ilLoop).iEndDate(0, 11), tgMCof(ilLoop).iStartDate(1, 11), llEndDate
'            If llInpDate >= llStartDate And llInpDate <= llEndDate Then
'                ilYear = tgMCof(ilLoop).iYear
'                ilStartMonth = tgMCof(ilLoop).iStartMnthNo
'                Exit Sub
'            End If
'        Next ilLoop
'    End If
'End Sub
Function mOpenBRFiles() As Integer
    Dim ilRet As Integer
    Dim ilValue As Integer

    mOpenBRFiles = 0                        'assume no Open File errors

    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmCHF

        mOpenBRFiles = 1
        Exit Function
    End If
    imCHFRecLen = Len(tmChf)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imClfRecLen = Len(tgClf(0).ClfRec)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imCffRecLen = Len(tgCff(0).CffRec)
    hmAdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imAdfRecLen = Len(tmAdf)
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imAgfRecLen = Len(tmAgf)
    
    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmAnf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
        mOpenBRFiles = 1
        Exit Function
    End If
    imAnfRecLen = Len(tmAnf)

    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imRdfRecLen = Len(tmRdf)
    hmDnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDnf, "", sgDBPath & "Dnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmDnf)
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmDnf
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imDnfRecLen = Len(tmDnf)
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmSlf)
'        ilRet = btrClose(hmDnf)
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmSlf
'        btrDestroy hmDnf
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imSlfRecLen = Len(tmSlf)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmMnf)
'        ilRet = btrClose(hmSlf)
'        ilRet = btrClose(hmDnf)
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmMnf
'        btrDestroy hmSlf
'        btrDestroy hmDnf
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imMnfRecLen = Len(tmMnf)
    hmUrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmUrf, "", sgDBPath & "Urf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmUrf)
'        ilRet = btrClose(hmMnf)
'        ilRet = btrClose(hmSlf)
'        ilRet = btrClose(hmDnf)
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmUrf
'        btrDestroy hmMnf
'        btrDestroy hmSlf
'        btrDestroy hmDnf
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imUrfRecLen = Len(tmUrf)
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmSof)
'        ilRet = btrClose(hmUrf)
'        ilRet = btrClose(hmMnf)
'        ilRet = btrClose(hmSlf)
'        ilRet = btrClose(hmDnf)
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmSof
'        btrDestroy hmUrf
'        btrDestroy hmMnf
'        btrDestroy hmSlf
'        btrDestroy hmDnf
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imSofRecLen = Len(tmSof)
    hmCbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCbf, "", sgDBPath & "Cbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmCbf)
'        ilRet = btrClose(hmSof)
'        ilRet = btrClose(hmUrf)
'        ilRet = btrClose(hmMnf)
'        ilRet = btrClose(hmSlf)
'        ilRet = btrClose(hmDnf)
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmCbf
'        btrDestroy hmSof
'        btrDestroy hmUrf
'        btrDestroy hmMnf
'        btrDestroy hmSlf
'        btrDestroy hmDnf
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imCbfRecLen = Len(tmCbf)
    hmCxf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmCxf)
'        ilRet = btrClose(hmCbf)
'        ilRet = btrClose(hmSof)
'        ilRet = btrClose(hmUrf)
'        ilRet = btrClose(hmMnf)
'        ilRet = btrClose(hmSlf)
'        ilRet = btrClose(hmDnf)
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmCxf
'        btrDestroy hmCbf
'        btrDestroy hmSof
'        btrDestroy hmUrf
'        btrDestroy hmMnf
'        btrDestroy hmSlf
'        btrDestroy hmDnf
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imCxfRecLen = Len(tmCxf)

    hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmDrf)
'        ilRet = btrClose(hmCbf)
'        ilRet = btrClose(hmSof)
'        ilRet = btrClose(hmUrf)
'        ilRet = btrClose(hmMnf)
'        ilRet = btrClose(hmSlf)
'        ilRet = btrClose(hmDnf)
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmDrf
'        btrDestroy hmCxf
'        btrDestroy hmCbf
'        btrDestroy hmSof
'        btrDestroy hmUrf
'        btrDestroy hmMnf
'        btrDestroy hmSlf
'        btrDestroy hmDnf
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    imDrfRecLen = Len(tmDrf)
    hmTChf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmTChf)
'        ilRet = btrClose(hmDrf)
'        ilRet = btrClose(hmCbf)
'        ilRet = btrClose(hmSof)
'        ilRet = btrClose(hmUrf)
'        ilRet = btrClose(hmMnf)
'        ilRet = btrClose(hmSlf)
'        ilRet = btrClose(hmDnf)
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmTChf
'        btrDestroy hmDrf
'        btrDestroy hmCxf
'        btrDestroy hmCbf
'        btrDestroy hmSof
'        btrDestroy hmUrf
'        btrDestroy hmMnf
'        btrDestroy hmSlf
'        btrDestroy hmDnf
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
        mOpenBRFiles = 1
        Exit Function
    End If
    hmDpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmDpf)
'        ilRet = btrClose(hmTChf)
'        ilRet = btrClose(hmDrf)
'        ilRet = btrClose(hmCbf)
'        ilRet = btrClose(hmSof)
'        ilRet = btrClose(hmUrf)
'        ilRet = btrClose(hmMnf)
'        ilRet = btrClose(hmSlf)
'        ilRet = btrClose(hmDnf)
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmDpf
'        btrDestroy hmTChf
'        btrDestroy hmDrf
'        btrDestroy hmCxf
'        btrDestroy hmCbf
'        btrDestroy hmSof
'        btrDestroy hmUrf
'        btrDestroy hmMnf
'        btrDestroy hmSlf
'        btrDestroy hmDnf
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
'
        mOpenBRFiles = 1
        Exit Function
    End If

    hmDef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmDef)
'        ilRet = btrClose(hmDpf)
'        ilRet = btrClose(hmTChf)
'        ilRet = btrClose(hmDrf)
'        ilRet = btrClose(hmCbf)
'        ilRet = btrClose(hmSof)
'        ilRet = btrClose(hmUrf)
'        ilRet = btrClose(hmMnf)
'        ilRet = btrClose(hmSlf)
'        ilRet = btrClose(hmDnf)
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmDef
'        btrDestroy hmDpf
'        btrDestroy hmTChf
'        btrDestroy hmDrf
'        btrDestroy hmCxf
'        btrDestroy hmCbf
'        btrDestroy hmSof
'        btrDestroy hmUrf
'        btrDestroy hmMnf
'        btrDestroy hmSlf
'        btrDestroy hmDnf
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
'
        mOpenBRFiles = 1
        Exit Function
    End If

    hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
'        ilRet = btrClose(hmSbf)
'        ilRet = btrClose(hmDef)
'        ilRet = btrClose(hmDpf)
'        ilRet = btrClose(hmTChf)
'        ilRet = btrClose(hmDrf)
'        ilRet = btrClose(hmCbf)
'        ilRet = btrClose(hmSof)
'        ilRet = btrClose(hmUrf)
'        ilRet = btrClose(hmMnf)
'        ilRet = btrClose(hmSlf)
'        ilRet = btrClose(hmDnf)
'        ilRet = btrClose(hmRdf)
'        ilRet = btrClose(hmAgf)
'        ilRet = btrClose(hmAdf)
'        ilRet = btrClose(hmCff)
'        ilRet = btrClose(hmClf)
'        ilRet = btrClose(hmCHF)
'        btrDestroy hmSbf
'        btrDestroy hmDef
'        btrDestroy hmDpf
'        btrDestroy hmTChf
'        btrDestroy hmDrf
'        btrDestroy hmCxf
'        btrDestroy hmCbf
'        btrDestroy hmSof
'        btrDestroy hmUrf
'        btrDestroy hmMnf
'        btrDestroy hmSlf
'        btrDestroy hmDnf
'        btrDestroy hmRdf
'        btrDestroy hmAgf
'        btrDestroy hmAdf
'        btrDestroy hmCff
'        btrDestroy hmClf
'        btrDestroy hmCHF
'
        mOpenBRFiles = 1
        Exit Function
    End If

    '4-23-13 moved from processing for split networks, always open for the demo names to create the string of books names for packages when different across hidden lines
    hmTxr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmTxr, "", sgDBPath & "Txr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mOpenBRFiles = -1
    End If
    If ((Asc(tgSpf.sUsingFeatures3) And TAXONAIRTIME) = TAXONAIRTIME) Or ((Asc(tgSpf.sUsingFeatures3) And TAXONNTR) = TAXONNTR) Then
        hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmRvf, "", sgDBPath & "Rvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
'            ilRet = btrClose(hmRvf)
'            ilRet = btrClose(hmDef)
'            ilRet = btrClose(hmDpf)
'            ilRet = btrClose(hmTChf)
'            ilRet = btrClose(hmDrf)
'            ilRet = btrClose(hmCbf)
'            ilRet = btrClose(hmSof)
'            ilRet = btrClose(hmUrf)
'            ilRet = btrClose(hmMnf)
'            ilRet = btrClose(hmSlf)
'            ilRet = btrClose(hmDnf)
'            ilRet = btrClose(hmRdf)
'            ilRet = btrClose(hmAgf)
'            ilRet = btrClose(hmAdf)
'            ilRet = btrClose(hmCff)
'            ilRet = btrClose(hmClf)
'            ilRet = btrClose(hmCHF)
'            btrDestroy hmRvf
'            btrDestroy hmDef
'            btrDestroy hmDpf
'            btrDestroy hmTChf
'            btrDestroy hmDrf
'            btrDestroy hmCxf
'            btrDestroy hmCbf
'            btrDestroy hmSof
'            btrDestroy hmUrf
'            btrDestroy hmMnf
'            btrDestroy hmSlf
'            btrDestroy hmDnf
'            btrDestroy hmRdf
'            btrDestroy hmAgf
'            btrDestroy hmAdf
'            btrDestroy hmCff
'            btrDestroy hmClf
'            btrDestroy hmCHF
            mOpenBRFiles = 1
            Exit Function
        End If

        hmPhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmPhf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
'            ilRet = btrClose(hmPhf)
'            ilRet = btrClose(hmRvf)
'            ilRet = btrClose(hmDef)
'            ilRet = btrClose(hmDpf)
'            ilRet = btrClose(hmTChf)
'            ilRet = btrClose(hmDrf)
'            ilRet = btrClose(hmCbf)
'            ilRet = btrClose(hmSof)
'            ilRet = btrClose(hmUrf)
'            ilRet = btrClose(hmMnf)
'            ilRet = btrClose(hmSlf)
'            ilRet = btrClose(hmDnf)
'            ilRet = btrClose(hmRdf)
'            ilRet = btrClose(hmAgf)
'            ilRet = btrClose(hmAdf)
'            ilRet = btrClose(hmCff)
'            ilRet = btrClose(hmClf)
'            ilRet = btrClose(hmCHF)
'            btrDestroy hmPhf
'            btrDestroy hmRvf
'            btrDestroy hmDef
'            btrDestroy hmDpf
'            btrDestroy hmTChf
'            btrDestroy hmDrf
'            btrDestroy hmCxf
'            btrDestroy hmCbf
'            btrDestroy hmSof
'            btrDestroy hmUrf
'            btrDestroy hmMnf
'            btrDestroy hmSlf
'            btrDestroy hmDnf
'            btrDestroy hmRdf
'            btrDestroy hmAgf
'            btrDestroy hmAdf
'            btrDestroy hmCff
'            btrDestroy hmClf
'            btrDestroy hmCHF
            mOpenBRFiles = 1
            Exit Function
        End If


        ilRet = gObtainTrf()
        If Not ilRet Then
            mOpenBRFiles = -1
        End If
        Exit Function
    End If
    ilValue = Asc(tgSpf.sSportInfo)
    If (ilValue And USINGSPORTS) = USINGSPORTS Then 'Using Sports, open game files
        hmCgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmCgf, "", sgDBPath & "Cgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
'            ilRet = btrClose(hmCgf)
'            ilRet = btrClose(hmDef)
'            ilRet = btrClose(hmDpf)
'            ilRet = btrClose(hmTChf)
'            ilRet = btrClose(hmDrf)
'            ilRet = btrClose(hmCbf)
'            ilRet = btrClose(hmSof)
'            ilRet = btrClose(hmUrf)
'            ilRet = btrClose(hmMnf)
'            ilRet = btrClose(hmSlf)
'            ilRet = btrClose(hmDnf)
'            ilRet = btrClose(hmRdf)
'            ilRet = btrClose(hmAgf)
'            ilRet = btrClose(hmAdf)
'            ilRet = btrClose(hmCff)
'            ilRet = btrClose(hmClf)
'            ilRet = btrClose(hmCHF)
'            ilRet = btrClose(hmRvf)
'            ilRet = btrClose(hmPhf)
'            ilRet = btrClose(hmSbf)
'            btrDestroy hmCgf
'            btrDestroy hmDef
'            btrDestroy hmDpf
'            btrDestroy hmTChf
'            btrDestroy hmDrf
'            btrDestroy hmCxf
'            btrDestroy hmCbf
'            btrDestroy hmSof
'            btrDestroy hmUrf
'            btrDestroy hmMnf
'            btrDestroy hmSlf
'            btrDestroy hmDnf
'            btrDestroy hmRdf
'            btrDestroy hmAgf
'            btrDestroy hmAdf
'            btrDestroy hmCff
'            btrDestroy hmClf
'            btrDestroy hmCHF
'            btrDestroy hmRvf
'            btrDestroy hmPhf
'            btrDestroy hmSbf
'
            mOpenBRFiles = 1
            Exit Function
        End If
        imCgfRecLen = Len(tmCgf)

        hmGsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmGsf, "", sgDBPath & "GSf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
'            ilRet = btrClose(hmGsf)
'            ilRet = btrClose(hmCgf)
'            ilRet = btrClose(hmDef)
'            ilRet = btrClose(hmDpf)
'            ilRet = btrClose(hmTChf)
'            ilRet = btrClose(hmDrf)
'            ilRet = btrClose(hmCbf)
'            ilRet = btrClose(hmSof)
'            ilRet = btrClose(hmUrf)
'            ilRet = btrClose(hmMnf)
'            ilRet = btrClose(hmSlf)
'            ilRet = btrClose(hmDnf)
'            ilRet = btrClose(hmRdf)
'            ilRet = btrClose(hmAgf)
'            ilRet = btrClose(hmAdf)
'            ilRet = btrClose(hmCff)
'            ilRet = btrClose(hmClf)
'            ilRet = btrClose(hmCHF)
'            ilRet = btrClose(hmRvf)
'            ilRet = btrClose(hmPhf)
'            ilRet = btrClose(hmSbf)
'            btrDestroy hmGsf
'            btrDestroy hmCgf
'            btrDestroy hmDef
'            btrDestroy hmDpf
'            btrDestroy hmTChf
'            btrDestroy hmDrf
'            btrDestroy hmCxf
'            btrDestroy hmCbf
'            btrDestroy hmSof
'            btrDestroy hmUrf
'            btrDestroy hmMnf
'            btrDestroy hmSlf
'            btrDestroy hmDnf
'            btrDestroy hmRdf
'            btrDestroy hmAgf
'            btrDestroy hmAdf
'            btrDestroy hmCff
'            btrDestroy hmClf
'            btrDestroy hmCHF
'            btrDestroy hmRvf
'            btrDestroy hmPhf
'            btrDestroy hmSbf
            mOpenBRFiles = 1
            Exit Function
        End If
        imGsfRecLen = Len(tmGsf)
    End If
    
    ilValue = Asc(tgSaf(0).sFeatures8)
    If ((ilValue And PODCASTCPMTAG) = PODCASTCPMTAG) Or ((ilValue And PODCASTCPMTAG) = PODCASTCPMTAG) Then   'using cpm podcasts
        'open  podcast cpm file
        hmPcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmPcf, "", sgDBPath & "Pcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mOpenBRFiles = -1
        End If
    End If

    If ((Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) <> SPLITNETWORKS) Then
        If ((Asc(tgSaf(0).sFeatures4) And MKTNAMEONBR) = MKTNAMEONBR) Then      '8-9-18 build stations and markets when not using split networks and using show mktname on contract
            ilRet = gObtainAllStations()
            ilRet = gObtainMarkets()
        End If
        Exit Function
    Else
        'open the necessary split network files
        hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mOpenBRFiles = -1
        End If

        hmSef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmSef, "", sgDBPath & "Sef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mOpenBRFiles = -1
        End If
'4-23-13 move so that txr is always open; also used for long package demo book names on the summary when more than 1 hidden line have different book names in the package'        hmTxr = CBtrvTable(ONEHANDLE) 'CBtrvObj()
'        ilRet = btrOpen(hmTxr, "", sgDBPath & "Txr.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
'        If ilRet <> BTRV_ERR_NONE Then
'            mOpenBRFiles = -1
'        End If

        hmShf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmShf, "", sgDBPath & "Shtt.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mOpenBRFiles = -1
        End If

        hmMkt = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmMkt, "", sgDBPath & "Mkt.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mOpenBRFiles = -1
        End If
        
        If ((Asc(tgSaf(0).sFeatures4) And MKTNAMEONBR) = MKTNAMEONBR) Then      '2-22-18 build stations and markets
            ilRet = gObtainAllStations()
            ilRet = gObtainMarkets()
        End If

        hmVLF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmVLF, "", sgDBPath & "Vlf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mOpenBRFiles = -1
        End If

        hmAtt = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmAtt, "", sgDBPath & "Att.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mOpenBRFiles = -1
        End If
    End If

End Function

'       mPutCntWkGrps - Store the overall weekly contract grps into
'               all detail lines, built and stored into tmWkCntGrps
'       2-18-00
Sub mPutCntWkGrps(ilQtr As Integer)
    Dim ilLoop As Integer
    Dim ilWeekInx As Integer
    Dim ilVehicle As Integer
    Dim ilFound As Integer
    ilFound = False
    If tmClf.sType = "H" Then       '2-20-00hidden lines have their own weekly grp totals because
         For ilVehicle = LBound(tmWkHiddenVGrps) To UBound(tmWkHiddenVGrps) - 1
            'If tmClf.iLine = tmWkHiddenVGrps(ilVehicle).iVefCode Then   'line # is stored in palce of vehicle in the weekly grp array
            '1-24-03
            If tmClf.iVefCode = tmWkHiddenVGrps(ilVehicle).iVefCode And tmClf.iPkLineNo = tmWkHiddenVGrps(ilVehicle).iPkLineNo Then   'line # is stored in palce of vehicle in the weekly grp array
                ilFound = True
                Exit For
            End If
        Next ilVehicle 'they are not combined with other lines for the same vehicle
    ElseIf tmClf.sType = "S" Then
        For ilVehicle = LBound(tmWkCntVGrps) To UBound(tmWkCntVGrps) - 1
            If tmCbf.iVefCode = tmWkCntVGrps(ilVehicle).iVefCode Then
                ilFound = True
                Exit For
            End If
        Next ilVehicle
    Else
        For ilVehicle = LBound(tmWkPkgVGrps) To UBound(tmWkPkgVGrps) - 1
            If tmClf.iLine = tmWkPkgVGrps(ilVehicle).iVefCode Then  'get the matching package grps from array
            ilFound = True
            Exit For
            End If
        Next ilVehicle

    End If
    'ilWeekInx = (ilQtr - 1) * 13 + 1           '11-14-11
    ilWeekInx = 0
    For ilLoop = 1 To ilQtr - 1
        ilWeekInx = ilWeekInx + imWeeksPerQtr(ilLoop)
    Next ilLoop
    ilWeekInx = ilWeekInx + 1
    
    'For ilLoop = 1 To 13
    For ilLoop = 1 To imWeeksPerQtr(ilQtr)          '11-14-11
        'tmCbf.lWkCntGrp(ilLoop - 1) = tmWkCntGrps(1).lGrps(ilWeekInx + ilLoop - 1) 'only 1 array for the contract weekly totals
        tmCbf.lWkCntGrp(ilLoop - 1) = tmWkCntGrps(0).lGrps(ilWeekInx + ilLoop - 1) 'only 1 array for the contract weekly totals
        If ilFound Then
            If tmClf.sType = "H" Then
                tmCbf.lWkVehGrp(ilLoop - 1) = tmWkHiddenVGrps(ilVehicle).lGrps(ilWeekInx + ilLoop - 1)
            ElseIf tmClf.sType = "S" Then
                tmCbf.lWkVehGrp(ilLoop - 1) = tmWkCntVGrps(ilVehicle).lGrps(ilWeekInx + ilLoop - 1)
            Else
                tmCbf.lWkVehGrp(ilLoop - 1) = tmWkPkgVGrps(ilVehicle).lGrps(ilWeekInx + ilLoop - 1)
            End If
        Else
            tmCbf.lWkVehGrp(ilLoop - 1) = 0
        End If
    Next ilLoop
    If ilLoop = 13 And imWeeksPerQtr(ilQtr) = 13 Then       '11-14-11 if processing 13th week in qtr, clear out the 14th week if
        'tmCbf.lWkVehGrp(14) = 0         'initialize 14th bucket when only 13 weeks exist in qtr
        tmCbf.lWkVehGrp(13) = 0         'initialize 14th bucket when only 13 weeks exist in qtr
    ElseIf ilLoop = 12 And imWeeksPerQtr(ilQtr) = 12 Then        'very rare occurence when qtr is 12 weeks
        'tmCbf.lWkVehGrp(13) = 0
        'tmCbf.lWkVehGrp(14) = 0
        tmCbf.lWkVehGrp(12) = 0
        tmCbf.lWkVehGrp(13) = 0
    End If
    
End Sub

'
'               Move in common data into Research arrays
'
'               mRchSameData  llIndex, llCost
'               <input>  llIndex - (4-10-19) changed from ilIndex) index into Research array
'                        llCost - total cost of period
'
'Sub mRchSameData(llIndex As Long, llCost As Long)
Sub mRchSameData(llIndex As Long, dlCost As Double) 'TTP 10439 - Rerate 21,000,000
    tmLRch(llIndex).dTotalCost = dlCost 'TTP 10439 - Rerate 21,000,000
    tmLRch(llIndex).lTotalCPP = tmCbf.lCPP
    tmLRch(llIndex).lTotalCPM = tmCbf.lCPM
    tmLRch(llIndex).lTotalGRP = tmCbf.lGRP
    tmLRch(llIndex).lTotalGrimps = tmCbf.lGrImp
    tmLRch(llIndex).iTotalAvgRating = tmCbf.iAvgRate
    tmLRch(llIndex).lTotalAvgAud = tmCbf.lAvgAud
    'ReDim Preserve tmLRch(1 To UBound(tmLRch) + 1)
    ReDim Preserve tmLRch(0 To UBound(tmLRch) + 1)      'Index zero ignored
End Sub

'
'                   mSetupBrHdr - Read advertiser, agency and salesperson records once per contract
'                       and place into the tmCBF (prepass record)
'
'
'
Sub mSetupBrHdr(ilWhichSort As Integer)
    Dim ilRet As Integer
    'Build the advertiser, agency, slsp, and sort fields specifications that need to be built
    'only once per contract.
    '
    tmCbf.iAdfCode = tgChf.iAdfCode
    tmAdfSrchKey.iCode = tgChf.iAdfCode
    ilRet = btrGetEqual(hmAdf, tmAdf, imAdfRecLen, tmAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching advt recd
    If ilRet <> BTRV_ERR_NONE Then
        tmCbf.iAdfCode = 0
        tmAdf.sName = "Missing"
    End If
    If tgChf.iAgfCode > 0 Then                          'contract has an agency
        tmCbf.iAgfCode = tgChf.iAgfCode
        tmAgfSrchKey.iCode = tgChf.iAgfCode
        ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching agency recd
        If ilRet <> BTRV_ERR_NONE Then
            tmAgf.sName = "Missing"
            tmCbf.iAgfCode = 0
        End If
    End If
    '10/27/14: set 1 or 2 place rating
    sm1or2PlaceRating = gSet1or2PlaceRating(tgChf.iAgfCode)
    If tgChf.iSlfCode(0) > 0 Then
        tmCbf.iSlfCode = tgChf.iSlfCode(0)
        tmSlfSrchKey.iCode = tgChf.iSlfCode(0)
        ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching Slsp recd
        If ilRet <> BTRV_ERR_NONE Then
            tmSlf.sFirstName = "Mising"
            tmSlf.sLastName = "Missing"
            tmCbf.iSlfCode = 0
        End If
    End If
    If ilWhichSort = 0 Then              'advt/cont
    'If RptSelCt!rbcSelCSelect(0).Value Then              'advt/cont
        tmCbf.sSortField1 = ""
        tmCbf.sSortField2 = ""
    ElseIf ilWhichSort = 1 Then           'agency
    'ElseIf RptSelCt!rbcSelCSelect(1).Value Then          'agency
        If tgChf.iAgfCode > 0 Then                      'has an agy (vs. direct)
            tmCbf.sSortField1 = tmAgf.sName
        Else
            tmCbf.sSortField1 = tmAdf.sName             'direct
        End If
        tmCbf.sSortField2 = tmAdf.sName
    Else                                                'salesperson
        tmCbf.sSortField1 = tmSlf.sFirstName & "" & tmSlf.sLastName
        tmCbf.sSortField2 = ""
    End If
    'format all fields that need to be retrieved from CHF and
    'place into CBF so that snapshot will work since there is
    'no actual contract written to disk
    tmCbf.iPropVer = tgChf.iPropVer
    tmCbf.iCntRevNo = tgChf.iCntRevNo
    tmCbf.iExtRevNo = tgChf.iExtRevNo
    'tmCbf.sType = tgChfCT.sType            'replace with Acq/Not Acq flag
    tmCbf.sStatus = tgChf.sStatus
    'tmCbf.sBuyer = tgChf.sBuyer        '1-20-09 buyer name will be obtained from pnf
    If (Asc(tgSaf(0).sFeatures6) And SIGNATUREONPROPOSAL) = SIGNATUREONPROPOSAL Then 'Print signature line on proposals      '6-14-19 if feature set to show signature line on Proposals, set falg
                                                                    'the field cbfdrfcode isnt used in crystal
        tmCbf.lDrfCode = 1              'show  signature line on working
    Else
        tmCbf.lDrfCode = 0              'do not show signature line except on the orders and complete prop
    End If
    tmCbf.sProduct = tgChf.sProduct
    tmCbf.iStartDate(0) = tgChf.iStartDate(0)
    tmCbf.iStartDate(1) = tgChf.iStartDate(1)
    tmCbf.iEndDate(0) = tgChf.iEndDate(0)
    tmCbf.iEndDate(1) = tgChf.iEndDate(1)
    tmCbf.iPctTrade = tgChf.iPctTrade
    tmCbf.sAgyCTrade = tgChf.sAgyCTrade
    If tgChf.sStatus = "H" Or tgChf.sStatus = "O" Or tgChf.sStatus = "G" Or tgChf.sStatus = "N" Then
        tmCbf.iPropOrdDate(0) = tgChf.iOHDDate(0)
        tmCbf.iPropOrdDate(1) = tgChf.iOHDDate(1)
        tmCbf.iPropOrdTime(0) = tgChf.iOHDTime(0)
        tmCbf.iPropOrdTime(1) = tgChf.iOHDTime(1)
    Else
        tmCbf.iPropOrdDate(0) = tgChf.iPropDate(0)
        tmCbf.iPropOrdDate(1) = tgChf.iPropDate(1)
        tmCbf.iPropOrdTime(0) = tgChf.iPropTime(0)
        tmCbf.iPropOrdTime(1) = tgChf.iPropTime(1)
    End If
    tmCbf.lContrNo = tgChf.lCntrNo
    
    tmCbf.lGenTime = lgNowTime
    tmCbf.iGenDate(0) = igNowDate(0)
    tmCbf.iGenDate(1) = igNowDate(1)
    '2-16-13 Update the user ID into the record to filter along with gendate and time
    tmCbf.iUrfCode = tgUrf(0).iCode
    
    tmCbf.lChfCode = tgChf.lCode                'contract internal code

        If imDiffExceeds104Wks Then             '3-9-06 this difference only exceeds 104 weeks.  its an old contract with
                                           'an invalid end date
    tmCbf.sSurvey = "Incomplete: exceeds 104+ weeks"
    End If
End Sub

'           gOrderAudit
'           Create prepass for Order Audit report which shows information
'           by vehicle (separated air time vs NTR):
'           Gross, Agy Comm, Net, Merch $, Promotions, Acquisition & Net-Net
'           Summary of monthly net-net shown.\
'           Rate card is retrieved to get the rates
'           Process one contract at a time, building array of unique air time vehicles and
'           NTR vehicles containing vheicle, ntr/air time, gross amt, agy comm, net, merchan $,
'               Promo $, acquistion costs, and T-net amts.  Also include the r/c amts
'           <input>  tlSBf() - array of NTR records
'           <output>  llStdStartDates() - array of 37 bdcst start months
'           CbfGenTime - generation time
'           CbfGenDAte = generation date
'           cbfCntrno = contract #
'           CbfVefCode - vehicle code
'           CbfType - A = air time, N = ntr
'           CbfMnfGroup = mnfcode for NTR type (0 for Air Time entries)
'           cbfStartDate - start date of contract
'           cbfEndDate - end date of contract
'           cbfDtFrstBkt = entered date
'           cbfdnfCode - rate card code
'           cbfSlfCode - salesperson code (1st slsp code)
'           CbfCurrMod - 1st slsp revenue split     11-14-08
'           cbfMonthUnits - 2nd-10th slsp code      11-14-08
'           cbflWkCntGrp - 2nd - 10th slsp revenue split  11-14-08
'           cbfAdfCode - advertiser code
'           cbfAgfcode - agency code
'           cbfProduct = product description
'           cbfMonth(1 - 12) = 12 monthly t-net $ (the entire contract may require more than 1 record if
'                       contract exceeds 1 year)
'           cbfExtra2Byte = sequence # for multiple records to contain monthly $ over 1 year.
'                           required only for the monthly summary records
'           cbfCntGrimps = total spots / contract (entire order)
'           vehicle totals: cbfValue(1) - Gross, (2) = agy comm, (3) = net, (4) = merchandising
'                                   (5) = promotions, (6) = acquisition, (7) = tnet, (8) = spot count
'           cbfRdfDPSort : 0 = detail (vehicle totals), 1 = summary (monthly totals)
'           CbfPop - contract R/C total for all spots (obtained from flight)
'           cbfMaxMonths - # months to print per summary record
'           cbfTotalWks- year of first date to print (1990....2010, etc)
'           cbfAirWks - first month to print (1-12)
'           cbfOtherComments - Other comments code
'           cbfIntComments - Internal comments code
'           cbfCancComment - cancellation comments code
'           cbfChgRComment - change reason comments code
'           cbfPromoComment - promo comment code
'           cbfMerchComment - merchandising comment code
'       Crystal sort:  Advertiser name, contract #, record type (0 or 1), vehicle name, sequence #
Public Function gOrderAudit(llStdStartDates() As Long, tlSbf() As SBF) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  tlRvfList                                                                             *
'******************************************************************************************
    Dim ilRet As Integer
    Dim ilClf As Integer
    'Dim llProject(1 To 36) As Long
    'Dim llProjectRC(1 To 36) As Long
    'Dim llProjectAcq(1 To 36) As Long
    'Dim llProjectSpots(1 To 36) As Long
    Dim llProject(0 To 36) As Long          'Index zero ignored
    Dim llProjectRC(0 To 36) As Long        'Index zero ignored
    Dim llProjectAcq(0 To 36) As Long       'Index aero ignored
    Dim llProjectSpots(0 To 36) As Long     'Index zero ignored
    Dim slStartDate As String       'contract earliest date from header
    Dim slEndDate As String         'contract latest date from header
    Dim llChfStart As Long
    Dim llChfEnd As Long
    Dim ilShowStdQtr As Integer     '12-19-20 0 = std, only calendar this report uses
    Dim llStartDate As Long         'contract earliest date from header
    Dim llEndDate As Long           'contract latest date from header
    Dim ilCurrStartQtr(0 To 1) As Integer
    Dim ilCurrTotalMonths As Integer
    Dim ilCorpStdYear As Integer
    Dim ilStartMonth As Integer
    Dim ilFound As Integer
    Dim ilLoopOnVef As Integer
    Dim ilLoopOnMonth As Integer
    Dim tlTranType As TRANTYPES
    Dim ilAgyComm As Integer
    Dim ilLoopOnRvf As Integer
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim llAmt As Long
    Dim ilLoopOnNTR As Integer
    Dim ilMonth As Integer
    Dim ilUpper As Integer
    Dim blDefaultToQtr As Boolean

    gOrderAudit = 0
    ilRet = mOpenBRFiles()                      'open all BR files
    If ilRet <> BTRV_ERR_NONE Then              'any error, exit and quit
        Screen.MousePointer = vbDefault
        gOrderAudit = -1                         'return and close all files
        mCloseBRFiles
        Exit Function
    End If
    'convert contracts earliest/latest dates
    gUnpackDate tgChf.iStartDate(0), tgChf.iStartDate(1), slStartDate
    If slStartDate = "" Then
        llStartDate = 0
    Else
        llStartDate = gDateValue(slStartDate)
    End If
    gUnpackDate tgChf.iEndDate(0), tgChf.iEndDate(1), slEndDate
    If slEndDate = "" Then
        llEndDate = 0
    Else
        llEndDate = gDateValue(slEndDate)
    End If

    'default end date to end date of the standard month.
    'if the end date of the contract falls within the middle of the month, the merchandising
    'maynot be picked up if its tran date is the end date of the std month
    'i.e.  start/end date of contract:  5/26/08 - 6/1/08.  The merch is entered for 6/29/08
    'and it gets bypassed
    slEndDate = gObtainEndStd(Format$(llEndDate, "m/d/yy"))


'        ilShowStdQtr = False        'show the monthly breakout by the actual start month of the order
    
    ilCurrStartQtr(0) = 0                       'init to place start date of starting qtr of order
    ilCurrStartQtr(1) = 0
    ilCurrTotalMonths = 0
    ilShowStdQtr = 0             '12-19-20 show in std bdcst months only, changed from true
    'Set up earliest/latest dates of contr, set to std dates.  Set array of starting bdcst months for summary page
    blDefaultToQtr = False
'        mFindMaxDates slStartDate, slEndDate, llChfStart, llChfEnd, ilCurrStartQtr(), ilCurrTotalMonths, llStdStartDates(), ilShowStdQtr, ilCorpStdYear, ilStartMonth, ilDefaultToQtr
    gFindMaxDates slStartDate, slEndDate, llChfStart, llChfEnd, ilCurrStartQtr(), ilCurrTotalMonths, llStdStartDates(), ilShowStdQtr, ilCorpStdYear, ilStartMonth, blDefaultToQtr

    If tgChf.iAgfCode > 0 Then          'agy commissionable
        'determine the agency commission, could be differnt than 15% so need to get the agy
        If tgChf.iAgfCode > 0 Then                          'contract has an agency
            tmAgfSrchKey.iCode = tgChf.iAgfCode
            ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching agency recd
            If ilRet <> BTRV_ERR_NONE Then
                ilAgyComm = 10000
            Else
                ilAgyComm = 10000 - tmAgf.iComm   'i.e 10000 - 1500 = 8500
            End If
        End If
    Else                                'direct
        ilAgyComm = 10000                            '10000
    End If

    ReDim tmAuditInfo(0 To 0) As AUDITINFO
    'ReDim lmMonthlyTNet(1 To 36) As Long    'max 3 years monthly tnet values for a single contract
    ReDim lmMonthlyTNet(0 To 36) As Long    'max 3 years monthly tnet values for a single contract. Index zero ignored
    'loop thru the air time lines to gather info
    'get projected actual rates, # spots, acquisition $ and R/c $
    lmContractSpots = 0                 'init total contract spot count
    ilUpper = 0
    For ilClf = LBound(tgClf) To UBound(tgClf) - 1
        tmClf = tgClf(ilClf).ClfRec
        If tmClf.sType <> "O" And tmClf.sType <> "A" And tmClf.sDelete = "N" Then       'ignore package lines
            gBuildFlightInfo ilClf, llStdStartDates(), 1, 36, llProject(), llProjectSpots(), llProjectRC(), llProjectAcq(), 1, tgClf(), tgCff()

            ilFound = False
            'see if this vehicle has already been processed and is in internal table
            For ilLoopOnVef = LBound(tmAuditInfo) To UBound(tmAuditInfo)
                If tmAuditInfo(ilLoopOnVef).iVefCode = tmClf.iVefCode And tmAuditInfo(ilLoopOnVef).iMnfCode = 0 Then
                    ilFound = True
                    Exit For
                End If
            Next ilLoopOnVef
            If Not ilFound Then
                tmAuditInfo(ilUpper).iVefCode = tmClf.iVefCode
                tmAuditInfo(ilUpper).iType = 0            'air time
                tmAuditInfo(ilUpper).iMnfCode = 0
                'vehicle totals
                For ilLoopOnMonth = 1 To 36
                    tmAuditInfo(ilUpper).lGross = tmAuditInfo(ilUpper).lGross + llProject(ilLoopOnMonth)
                    tmAuditInfo(ilUpper).lRateCard = tmAuditInfo(ilUpper).lRateCard + llProjectRC(ilLoopOnMonth)
                    tmAuditInfo(ilUpper).lNet = tmAuditInfo(ilUpper).lNet + ((llProject(ilLoopOnMonth) * CDbl(ilAgyComm)) / 10000)
                    tmAuditInfo(ilUpper).lAgyComm = tmAuditInfo(ilUpper).lAgyComm + (llProject(ilLoopOnMonth) - ((llProject(ilLoopOnMonth) * CDbl(ilAgyComm)) / 10000))
                    tmAuditInfo(ilUpper).lAcquisition = tmAuditInfo(ilUpper).lAcquisition + llProjectAcq(ilLoopOnMonth)
                    tmAuditInfo(ilUpper).lSpots = tmAuditInfo(ilUpper).lSpots + llProjectSpots(ilLoopOnMonth)
                Next ilLoopOnMonth
                ReDim Preserve tmAuditInfo(0 To ilUpper + 1) As AUDITINFO
                ilUpper = ilUpper + 1
            Else                'already an entry in array for the vehicle
                'vehicle totals
                For ilLoopOnMonth = 1 To 36
                    tmAuditInfo(ilLoopOnVef).lGross = tmAuditInfo(ilLoopOnVef).lGross + llProject(ilLoopOnMonth)
                    tmAuditInfo(ilLoopOnVef).lRateCard = tmAuditInfo(ilLoopOnVef).lRateCard + llProjectRC(ilLoopOnMonth)
                    tmAuditInfo(ilLoopOnVef).lNet = tmAuditInfo(ilLoopOnVef).lNet + ((llProject(ilLoopOnMonth) * CDbl(ilAgyComm)) / 10000)
                    tmAuditInfo(ilLoopOnVef).lAgyComm = tmAuditInfo(ilLoopOnVef).lAgyComm + (llProject(ilLoopOnMonth) - ((llProject(ilLoopOnMonth) * CDbl(ilAgyComm)) / 10000))
                    tmAuditInfo(ilLoopOnVef).lAcquisition = tmAuditInfo(ilLoopOnVef).lAcquisition + llProjectAcq(ilLoopOnMonth)
                    tmAuditInfo(ilLoopOnVef).lSpots = tmAuditInfo(ilLoopOnVef).lSpots + llProjectSpots(ilLoopOnMonth)
                Next ilLoopOnMonth
            End If

            'Contract monthly t-net values; NTR is excluded from Tnet calculation
            For ilMonth = 1 To 36
                    lmContractSpots = lmContractSpots + llProjectSpots(ilMonth)
                    'gross * agy comm (10000-1500) - acq (get net minus acquisition for tnet value)
                    lmMonthlyTNet(ilMonth) = lmMonthlyTNet(ilMonth) + ((llProject(ilMonth) * CDbl(ilAgyComm)) / 10000) - llProjectAcq(ilMonth)  'net, and subtr acq cost
            Next ilMonth

            For ilMonth = 1 To 36
                llProject(ilMonth) = 0
                llProjectRC(ilMonth) = 0
                llProjectAcq(ilMonth) = 0
                llProjectSpots(ilMonth) = 0
            Next ilMonth
        End If
    Next ilClf

    'determine net amount for the R/C contract total
    For ilLoopOnVef = 0 To UBound(tmAuditInfo) - 1
        tmAuditInfo(ilLoopOnVef).lRateCard = (tmAuditInfo(ilLoopOnVef).lRateCard * CDbl(ilAgyComm)) / 100
    Next ilLoopOnVef

    'get the merchandising and promotions records from RVF/PHF if its not a new contract entry
    If tgChf.lCntrNo <> 0 Then      'lack of contract # means new contract entry
        tlTranType.iAdj = False              'look only for adjustments in the History & Rec files
        tlTranType.iInv = True
        tlTranType.iWriteOff = False
        tlTranType.iPymt = False
        tlTranType.iCash = True
        tlTranType.iTrade = True
        tlTranType.iMerch = True
        tlTranType.iPromo = True
        tlTranType.iNTR = False
        ReDim tmRvfList(0 To 0) As RVF
        'retrieve merchandising and promotion records from receivables/history
#If programmatic <> 1 Then
        ilRet = gObtainPhfRvfbyCntr(BrSnap, tgChf.lCntrNo, slStartDate, slEndDate, tlTranType, tmRvfList())
#End If
        'loop thru the merch/promo records and accumulate the $
        For ilLoopOnRvf = LBound(tmRvfList) To UBound(tmRvfList) - 1
            tmRvf = tmRvfList(ilLoopOnRvf)
            For ilLoopOnVef = LBound(tmAuditInfo) To UBound(tmAuditInfo) - 1
                If tmAuditInfo(ilLoopOnVef).iVefCode = tmRvf.iAirVefCode Then
                    gUnpackDateLong tmRvf.iTranDate(0), tmRvf.iTranDate(1), llDate
                    For ilLoop = 1 To 36
                        If llDate >= llStdStartDates(ilLoop) And llDate < llStdStartDates(ilLoop + 1) Then  'tran date falls within one of the months
                            gPDNToLong tmRvf.sNet, llAmt
                            If tmRvf.sCashTrade = "M" Then      'merchandising
                                tmAuditInfo(ilLoopOnVef).lMerch = tmAuditInfo(ilLoopOnVef).lMerch + llAmt       'vheicle total merch acquisition costs
                                tmAuditInfo(ilLoopOnVef).lAcquisition = tmAuditInfo(ilLoopOnVef).lAcquisition + tmRvf.lAcquisitionCost  'accum vehicle acq costs
                                lmMonthlyTNet(ilLoop) = lmMonthlyTNet(ilLoop) - llAmt                           'tNet by Month, subtract the acq. costs
                           ElseIf tmRvf.sCashTrade = "P" Then  'promotions
                                tmAuditInfo(ilLoopOnVef).lPromo = tmAuditInfo(ilLoopOnVef).lPromo + llAmt       'vehicle total promo acquistion costs
                                tmAuditInfo(ilLoopOnVef).lAcquisition = tmAuditInfo(ilLoopOnVef).lAcquisition + tmRvf.lAcquisitionCost  'accum vehicle acq costs
                                lmMonthlyTNet(ilLoop) = lmMonthlyTNet(ilLoop) - llAmt                           'tNet by Month, subtract the acq. costs
                            End If
                            Exit For
                        End If
                    Next ilLoop
                    Exit For
                End If
            Next ilLoopOnVef
        Next ilLoopOnRvf
    End If

    'process NTR; NTR is excluded from Tnet calculation
    ilUpper = UBound(tmAuditInfo)
    For ilLoopOnNTR = LBound(tlSbf) To UBound(tlSbf) - 1
        tmSbf = tlSbf(ilLoopOnNTR)
        If tmSbf.sAgyComm <> "Y" Then
            ilAgyComm = 10000
        Else
            'commissionable if theres an agency
            If tgChf.iAgfCode = 0 Then
                ilAgyComm = 10000
            Else
                ilAgyComm = 10000 - tmAgf.iComm
            End If
        End If
        ilFound = False

        gUnpackDateLong tmSbf.iDate(0), tmSbf.iDate(1), llDate
        For ilLoop = 1 To 36
            If llDate >= llStdStartDates(ilLoop) And llDate < llStdStartDates(ilLoop + 1) Then  'tran date falls within one of the months
                lmMonthlyTNet(ilLoop) = lmMonthlyTNet(ilLoop) + (((tmSbf.lGross * tmSbf.iNoItems) * CDbl(ilAgyComm)) / 10000) - (tmSbf.lAcquisitionCost * tmSbf.iNoItems) 'net, and subtr acq cost
            End If
        Next ilLoop
        For ilLoopOnVef = LBound(tmAuditInfo) To UBound(tmAuditInfo) - 1

            If tmAuditInfo(ilLoopOnVef).iVefCode = tmSbf.iAirVefCode And tmAuditInfo(ilLoopOnVef).iMnfCode = tmSbf.iMnfItem Then
                For ilLoop = 1 To 36
                    If llDate >= llStdStartDates(ilLoop) And llDate < llStdStartDates(ilLoop + 1) Then  'tran date falls within one of the months
                        ilFound = True
                        tmAuditInfo(ilLoopOnVef).lGross = tmAuditInfo(ilLoopOnVef).lGross + (tmSbf.lGross * tmSbf.iNoItems)

                        tmAuditInfo(ilLoopOnVef).lNet = tmAuditInfo(ilLoopOnVef).lNet + (((tmSbf.lGross * tmSbf.iNoItems) * CDbl(ilAgyComm)) / 10000)       'net
                        tmAuditInfo(ilLoopOnVef).lAgyComm = tmAuditInfo(ilLoopOnVef).lAgyComm + ((tmSbf.lGross * tmSbf.iNoItems) - (((tmSbf.lGross * tmSbf.iNoItems) * CDbl(ilAgyComm)) / 10000))  'agy comm = gross - net
                        tmAuditInfo(ilLoopOnVef).lAcquisition = tmAuditInfo(ilLoopOnVef).lAcquisition + (tmSbf.lAcquisitionCost * tmSbf.iNoItems)
                        tmAuditInfo(ilLoopOnVef).lSpots = 0     'spot count doesnt apply for NTR
                        'lmMonthlyTNet(ilLoop) = lmMonthlyTNet(ilLoop) + (((tmSbf.lGross * tmSbf.iNoItems) * CDbl(ilAgyComm)) / 10000) - (tmSbf.lAcquisitionCost * tmSbf.iNoItems) 'net, and subtr acq cost
                    End If
                Next ilLoop
                Exit For
            End If
        Next ilLoopOnVef
        If Not ilFound Then
            tmAuditInfo(ilUpper).iVefCode = tmSbf.iAirVefCode
            tmAuditInfo(ilUpper).iType = 1           'NTR
            tmAuditInfo(ilUpper).iMnfCode = tmSbf.iMnfItem      'item type mnf code
            tmAuditInfo(ilUpper).lGross = tmSbf.lGross * tmSbf.iNoItems
            tmAuditInfo(ilUpper).lNet = (((tmSbf.lGross * tmSbf.iNoItems) * CDbl(ilAgyComm)) / 10000)
            tmAuditInfo(ilUpper).lAgyComm = ((tmSbf.lGross * tmSbf.iNoItems) - (((tmSbf.lGross * tmSbf.iNoItems) * CDbl(ilAgyComm)) / 10000))
            tmAuditInfo(ilUpper).lAcquisition = tmSbf.lAcquisitionCost * tmSbf.iNoItems
            tmAuditInfo(ilUpper).lSpots = 0         'spot counts dont apply for NTR
            'lmMonthlyTNet(ilLoop) = lmMonthlyTNet(ilLoop) + (((tmSbf.lGross * tmSbf.iNoItems) * CDbl(ilAgyComm)) / 10000) - (tmSbf.lAcquisitionCost * tmSbf.iNoItems) 'net, and subtr acq cost

            ilUpper = ilUpper + 1
            ReDim Preserve tmAuditInfo(0 To ilUpper) As AUDITINFO


        End If
    Next ilLoopOnNTR

    Erase tmRvfList, tlSbf
End Function

'
'       close files for BR and order audit report
'
Public Sub mCloseBRFiles()
    Dim ilRet As Integer
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmUrf)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmDnf)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmTChf)
    ilRet = btrClose(hmCbf)
    ilRet = btrClose(hmCxf)
    ilRet = btrClose(hmDrf)
    ilRet = btrClose(hmDpf)
    ilRet = btrClose(hmDef)
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmRvf)
    ilRet = btrClose(hmPhf)
    ilRet = btrClose(hmCgf)
    ilRet = btrClose(hmGsf)
    ilRet = btrClose(hmRaf)
    ilRet = btrClose(hmSef)
    ilRet = btrClose(hmTxr)
    ilRet = btrClose(hmShf)
    ilRet = btrClose(hmMkt)
    ilRet = btrClose(hmVLF)
    ilRet = btrClose(hmAtt)
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmPcf)
    ilRet = btrClose(hmThf)
    ilRet = btrClose(hmTif)

    btrDestroy hmSof
    btrDestroy hmUrf
    btrDestroy hmMnf
    btrDestroy hmSlf
    btrDestroy hmRdf
    btrDestroy hmDnf
    btrDestroy hmAgf
    btrDestroy hmAdf
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmCHF
    btrDestroy hmTChf
    btrDestroy hmCbf
    btrDestroy hmCxf
    btrDestroy hmDrf
    btrDestroy hmDpf
    btrDestroy hmDef
    btrDestroy hmSbf
    btrDestroy hmRvf
    btrDestroy hmPhf
    btrDestroy hmCgf
    btrDestroy hmGsf
    btrDestroy hmRaf
    btrDestroy hmSef
    btrDestroy hmTxr
    btrDestroy hmShf
    btrDestroy hmMkt
    btrDestroy hmVLF
    btrDestroy hmAtt
    btrDestroy hmAnf
    btrDestroy hmPcf
    btrDestroy hmThf
    btrDestroy hmTif
    Exit Sub
End Sub

'           gOrderAuditWrite
'           Write CBF records built from internal memory arrays
'           to create report for Order Audit Report
'           Create array from array tmAuditInfo
'
Public Sub gOrderAuditWrite(llStdStartDates() As Long)
    Dim ilLoop As Integer
    Dim ilSeq As Integer        'max 3 records per contract for monthly totals
    Dim ilTemp As Integer
    Dim ilMaxMonths As Integer
    Dim il12Months As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim slYear As String
    Dim slMonth As String
    Dim slDay As String

    tmCbf.lGenTime = lgNowTime          '10-30-01
    tmCbf.iGenDate(0) = igNowDate(0)
    tmCbf.iGenDate(1) = igNowDate(1)
    tmCbf.lContrNo = tgChf.lCntrNo               'contract #
    tmCbf.iAdfCode = tgChf.iAdfCode
    tmCbf.sProduct = Trim(tgChf.sProduct)
    tmCbf.iAgfCode = tgChf.iAgfCode
    tmCbf.iSlfCode = tgChf.iSlfCode(0)
    'show all 10 slsp splits, and slsp revenue splits
    tmCbf.lCurrMod = tgChf.lComm(0)         '1st slsp revenue split %
    For ilLoop = 1 To 9                         'slsp splits
        tmCbf.lMonthUnits(ilLoop - 1) = tgChf.iSlfCode(ilLoop)
        tmCbf.lWkCntGrp(ilLoop - 1) = tgChf.lComm(ilLoop)     'rev splits
    Next ilLoop

    tmCbf.lOtherComment = tgChf.lCxfCode      'other commnets
    tmCbf.lIntComment = tgChf.lCxfInt           'internal comments
    tmCbf.lCancComment = tgChf.lCxfCanc         'cancel comments
    tmCbf.lPromoComment = tgChf.lCxfProm        'promo comments
    tmCbf.lMerchComment = tgChf.lCxfMerch       'merchandising comments
    tmCbf.lChgRComment = tgChf.lCxfChgR         'change reason comments
    tmCbf.iExtra2Byte = 1                       'init seq #
    tmCbf.iDnfCode = tgChf.iRcfCode             'rate card code
    tmCbf.iStartDate(0) = tgChf.iStartDate(0)
    tmCbf.iStartDate(1) = tgChf.iStartDate(1)
    tmCbf.iEndDate(0) = tgChf.iEndDate(0)
    tmCbf.iEndDate(1) = tgChf.iEndDate(1)
    tmCbf.lCntGrimps = lmContractSpots          'total spots/contract
    tmCbf.iDtFrstBkt(0) = tgChf.iOHDDate(0)     'date entered
    tmCbf.iDtFrstBkt(1) = tgChf.iOHDDate(1)
    For ilLoop = LBound(tmAuditInfo) To UBound(tmAuditInfo) - 1
        tmCbf.iVefCode = tmAuditInfo(ilLoop).iVefCode
        tmCbf.lPop = tmAuditInfo(ilLoop).lRateCard
        'tmCbf.lValue(1) = tmAuditInfo(ilLoop).lGross
        'tmCbf.lValue(2) = tmAuditInfo(ilLoop).lAgyComm
        'tmCbf.lValue(3) = tmAuditInfo(ilLoop).lNet
        'tmCbf.lValue(4) = tmAuditInfo(ilLoop).lMerch
        'tmCbf.lValue(5) = tmAuditInfo(ilLoop).lPromo
        'tmCbf.lValue(6) = tmAuditInfo(ilLoop).lAcquisition
        'tmCbf.lValue(7) = tmAuditInfo(ilLoop).lNet - tmAuditInfo(ilLoop).lMerch - tmAuditInfo(ilLoop).lPromo - tmAuditInfo(ilLoop).lAcquisition
        'tmCbf.lValue(8) = tmAuditInfo(ilLoop).lSpots
        
        tmCbf.lValue(0) = tmAuditInfo(ilLoop).lGross
        tmCbf.lValue(1) = tmAuditInfo(ilLoop).lAgyComm
        tmCbf.lValue(2) = tmAuditInfo(ilLoop).lNet
        tmCbf.lValue(3) = tmAuditInfo(ilLoop).lMerch
        tmCbf.lValue(4) = tmAuditInfo(ilLoop).lPromo
        tmCbf.lValue(5) = tmAuditInfo(ilLoop).lAcquisition
        tmCbf.lValue(6) = tmAuditInfo(ilLoop).lNet - tmAuditInfo(ilLoop).lMerch - tmAuditInfo(ilLoop).lPromo - tmAuditInfo(ilLoop).lAcquisition
        tmCbf.lValue(7) = tmAuditInfo(ilLoop).lSpots
        tmCbf.iRdfDPSort = 0                'detail record type
        If tmAuditInfo(ilLoop).iType = 0 Then       'air time
            tmCbf.sType = "A"
            tmCbf.iMnfGroup = 0                     'ntr mnf code does not apply on air time
        Else
            tmCbf.sType = "N"                   'ntr
            tmCbf.iMnfGroup = tmAuditInfo(ilLoop).iMnfCode
        End If
        tmCbf.iExtra2Byte = 0               'seq # n/a for detail record, there is only 1 per vehicle
        ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
    Next ilLoop

    For ilTemp = 1 To 12
        tmCbf.lMonth(ilTemp - 1) = 0
    Next ilTemp
    For ilMaxMonths = 36 To 1 Step -1        'determine how many months to print, finding the latest month with $
        If lmMonthlyTNet(ilMaxMonths) <> 0 Then
            Exit For
        End If
    Next ilMaxMonths
    'always create at least 1 record / contract
    If ilTemp < 1 Then
        ilTemp = 1
    End If
    'determine # of records (12 months per record)
    If ilMaxMonths Mod 12 <> 0 Then          'remainder, need extra record
        ilTemp = (ilMaxMonths \ 12) + 1
    Else
        ilTemp = ilMaxMonths / 12
    End If
    'tmCbf.iTotalMonths = ilTemp         'total months to print
    tmCbf.iRdfDPSort = 1                'summary montly totals type
    'create up to 3 records for the monthly $
    For ilSeq = 1 To ilTemp
        tmCbf.iExtra2Byte = ilSeq
        'get the start date of the 2nd month to print, then subtract 1 day from it to get the end date of
        'the first month.  The end date of the bdcst month will always fall within the calendar (jan, feb...) month
        'gPackDateLong llStdStartDates((ilSeq - 1) * 12 + 1), tmCbf.iDtFrstBkt(0), tmCbf.iDtFrstBkt(1)
        slDate = Format$(llStdStartDates((ilSeq - 1) * 12 + 2) - 1, "m/d/yy")
        gObtainYearMonthDayStr slDate, True, slYear, slMonth, slDay
        tmCbf.iAirWks = Val(slMonth)            'first month to print
        tmCbf.iTotalWks = Val(slYear)
        For il12Months = 1 To 12
            tmCbf.lMonth(il12Months - 1) = lmMonthlyTNet((ilSeq - 1) * 12 + il12Months)
            tmCbf.lWeek(il12Months - 1) = llStdStartDates((ilSeq - 1) * 12 + il12Months)
        Next il12Months
        If ilMaxMonths <= 12 Then
                tmCbf.iTotalMonths = ilMaxMonths
            Else
                tmCbf.iTotalMonths = 12
                ilMaxMonths = ilMaxMonths - 12
            End If
        ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
    Next ilSeq
    Erase tmAuditInfo
    mCloseBRFiles
    Exit Sub
End Sub

'
'           mSetResortField - tmCbf.sResort is used in sorting the output in the proposal/contract print.
'           This has to be set for all lines, including Cancel Before start lines
'           '12-22-20  Modify to handle resort flags for both schedule lines and CPM Line IDs
Public Sub mSetResortField(slType As String, llLineRef As Long)
    Dim slStr As String

    tmCbf.sResort = ""
    tmCbf.sResortType = ""          '5-31-05
'        If tmClf.sType = "H" Then               'hidden
    If Trim$(slType) = "H" Then               'hidden
'            slStr = Trim$(str$(tmCbf.lIntComment))  'package line # stored in this field
        slStr = Trim$(str$(llLineRef))  'package line # stored in this field
        Do While Len(slStr) < 4         '5-31-05 use 4 digit line #s (vs 3 digit line #s)
        slStr = "0" & slStr
        Loop
        tmCbf.sResort = slStr '& "C"
        tmCbf.sResortType = "C"         '5-31-05
'        ElseIf tmClf.sType = "A" Or tmClf.sType = "O" Or tmClf.sType = "E" Then     'packages
    ElseIf Trim$(slType) = "A" Or Trim$(slType) = "O" Or Trim$(slType) = "E" Or Trim$(slType) = "P" Then     'packages: P = package for cpm; all others package for spots
        slStr = Trim$(str$(llLineRef))
        Do While Len(slStr) < 4         '5-31-05
        slStr = "0" & slStr
        Loop
        tmCbf.sResort = slStr '& "A"
        tmCbf.sResortType = "A"         '5-31-05
    Else                                    'conventionals, all others (fall after package/hiddens)
        tmCbf.sResort = "9999"  '~"
        tmCbf.sResortType = "~"         '5-31-05
    End If
End Sub

'           3-15-19 Gather all the hidden lines for each unique package (1 pkg may be used more than once)
Public Sub mGatherPkgVehSummary(ilLineSpots As Integer, ilPkgVehList() As Integer, ilVehList() As Integer, llOverallPopEst As Long, llLnSpots As Long, llPopByLine() As Long, llWklySpots() As Long, llWklyRates() As Long, llWklyAvgAud() As Long, llWklyPopEst() As Long, lmSpotsByWk() As Long, lmWklyRates() As Long, lmAvgAud() As Long, lmPopEst() As Long, ilWklyRtg() As Integer, llWklyGrimp() As Long, llWklyGRP() As Long)
    Dim ilLoop As Integer
    Dim ilLoop3 As Integer
    Dim ilFoundLine As Integer
    Dim ilTemp As Integer
    Dim llPop As Long
    Dim ilVehicle As Integer
    Dim ilSavePkgVeh As Integer
    Dim ilPkg As Integer
    Dim ilSpots As Integer
    Dim ilLoopOnPkgVehList As Integer
    Dim ilLoopOnVehList As Integer
    Dim ilPkgLines() As Integer
    Dim ilLoopOnPkgLines As Integer
    Dim blGotHiddenLineForPkg As Boolean
    'Dim llTemp As Long
    Dim dlTemp As Double 'TTP 10439 - Rerate 21,000,000
    Dim llPopEst As Long
    Dim llResearchPop As Long
    Dim blItsAPkg As Boolean
    'Dim ilRchQtr As Integer
    Dim ilVefInxForCallLetters As Integer
    Dim ilRet As Integer
    Dim llTempLRch As Long          '4-10-19 replace ilTemp due to subscript out of range
    Dim llRchQtr As Long            '4-01-19

'            ReDim tmLRch(0 To 1) As RESEARCHINFO        '10-30-01. Index zero ignored
    For ilLoopOnVehList = 0 To UBound(ilVehList) - 1        'loop on unique vehicle codes
        blItsAPkg = False
        For ilLoopOnPkgVehList = 0 To UBound(ilPkgVehList) - 1        'tet to see if the vehicle is a package.  if not, ignore it
            If ilVehList(ilLoopOnVehList) = ilPkgVehList(ilLoopOnPkgVehList) Then
                blItsAPkg = True
                Exit For
            End If
        Next ilLoopOnPkgVehList                     'ilLoopOnPkgVehList = 0 To UBound(ilPkgVehList) - 1

        If blItsAPkg Then          'bypass vehicles that are not packages
            ReDim ilPkgLines(0 To 0) As Integer
            ReDim tmLRch(0 To 1) As RESEARCHINFO        '10-30-01. Index zero ignored
            llLnSpots = 0
            llResearchPop = -1
            For ilPkg = 0 To UBound(tgClf) - 1 Step 1       'determine which line #s are using the same pkg vehicle
                If ilPkgVehList(ilLoopOnPkgVehList) = tgClf(ilPkg).ClfRec.iVefCode Then    'find the assoc pkg vehicle name
                    ilPkgLines(UBound(ilPkgLines)) = tgClf(ilPkg).ClfRec.iLine              'save the line #s for this matching package vehicle, could be used more than one and need to combine them all
                    ReDim Preserve ilPkgLines(0 To UBound(ilPkgLines) + 1) As Integer
                End If
            Next ilPkg
            'Now search for all the matching hidden lines using this package name and process its line for the package vehicle summary
            For ilLoop = 1 To ilLineSpots - 1
                If imProcessFlag(tmLnr(ilLoop).iLineInx) = 1 Or imProcessFlag(tmLnr(ilLoop).iLineInx) = 3 Or imProcessFlag(tmLnr(ilLoop).iLineInx) = 4 Then  'show current and prev lines depending
                    'on the option selected.  Mods show current & prev, Full BR shows curent only
                    'Move spots, $ and aud to single dimension array for gAvgAudToLnResearch routine
                    ilLoop3 = tmLnr(ilLoop).iLineInx            'line index to process
                    tmClf = tgClf(ilLoop3).ClfRec

                    ilFoundLine = False
                    llTempLRch = UBound(tmLRch)
                    blGotHiddenLineForPkg = False
                    For ilLoopOnPkgLines = 0 To UBound(ilPkgLines) - 1
                        If tmClf.iPkLineNo = ilPkgLines(ilLoopOnPkgLines) Then
                            blGotHiddenLineForPkg = True
                            Exit For
                        End If
                    Next ilLoopOnPkgLines

                    If blGotHiddenLineForPkg Then
                        'find the hidden vehicle for its population
                        llPop = 0
                        For ilVehicle = LBound(ilVehList) To UBound(ilVehList) - 1 Step 1
                            If tmClf.iVefCode = ilVehList(ilVehicle) Then
                                llPop = llPopByLine(ilLoop3 + 1)    '11-23-99
                                If tgSpf.sDemoEstAllowed <> "Y" Then            'this if for the researchtotal population
                                    If llResearchPop = -1 Then
                                        If llPop <> 0 Then
                                            llResearchPop = llPop
                                        End If
                                    Else                                    'population already set with at least 1 lines population
                                        If llResearchPop <> llPop And llPop <> 0 Then
                                            llResearchPop = 0                   'found at least 1 vehicle with diff pop
                                        End If
                                    End If
                                End If
                                Exit For
                            End If
                        Next ilVehicle
                        
                        'Need overall vehicle population based on if there were different books across the package
                        ilSavePkgVeh = ilVehList(ilLoopOnVehList)
                       
                        'create the array for all schedule lines
                        For ilSpots = 1 To MAXWEEKSFOR2YRS
                            llWklySpots(ilSpots - 1) = lmSpotsByWk(ilSpots, ilLoop)
                            llLnSpots = llLnSpots + lmSpotsByWk(ilSpots, ilLoop)        'accum for research total pkg vehicle
                            tmLRch(llTempLRch).lQSpots = tmLRch(llTempLRch).lQSpots + llWklySpots(ilSpots - 1)
                            llWklyRates(ilSpots - 1) = lmWklyRates(ilSpots, ilLoop3)
                            llWklyAvgAud(ilSpots - 1) = lmAvgAud(ilSpots, ilLoop3)
                            llWklyPopEst(ilSpots - 1) = lmPopEst(ilSpots, ilLoop3)
                        Next ilSpots
                        'gAvgAudToLnResearch True, llPop, llWklySpots(), llWklyRates(), llWklyAvgAud(), llTemp, tmCbf.lAvgAud, ilWklyRtg(), tmCbf.iAvgRate, llWklyGrimp(), tmCbf.lGrimp, llWklyGRP(), tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM
                        'gAvgAudToLnResearch sm1or2PlaceRating, True, llPop, llWklyPopEst(), llWklySpots(), llWklyRates(), llWklyAvgAud(), llTemp, tmCbf.lAvgAud, ilWklyRtg(), tmCbf.iAvgRate, llWklyGrimp(), tmCbf.lGrImp, llWklyGRP(), tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, llPopEst
                        gAvgAudToLnResearch sm1or2PlaceRating, True, llPop, llWklyPopEst(), llWklySpots(), llWklyRates(), llWklyAvgAud(), dlTemp, tmCbf.lAvgAud, ilWklyRtg(), tmCbf.iAvgRate, llWklyGrimp(), tmCbf.lGrImp, llWklyGRP(), tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, llPopEst 'TTP 10439 - Rerate 21,000,000
                        tmLRch(llTempLRch).iVefCode = tmClf.iVefCode
                        tmLRch(llTempLRch).iLineNo = tmClf.iLine
                        tmLRch(llTempLRch).sType = tmClf.sType          's = std, O = order, a=air, H = hidden
                        tmLRch(llTempLRch).iPkLineNo = tmClf.iPkLineNo  'pkg line reference if hidden
                        tmLRch(llTempLRch).iPkvefCode = ilSavePkgVeh     'associ pkg vehicle code if hidden line
                        'mRchSameData llTempLRch, llTemp
                        mRchSameData llTempLRch, dlTemp 'TTP 10439 - Rerate 21,000,000
    
                        '5-28-04 If using research estimates, see if different estimates across line and or vehicle
                        If tgSpf.sDemoEstAllowed = "Y" Then
                            llPopByLine(ilLoop3 + 1) = llPopEst
                            If lmPop(ilVehicle) = 0 Then            'same vehicle found more than once, if pop already stored,
                                                                    'dont wipe out with a possible non-population value
                                lmPop(ilVehicle) = llPopEst           'associate the population with the vehicle
                            End If
                            If llResearchPop = -1 And llPopEst <> 0 Then          'first time, llResearchPop is for the summary records (-1 first time thru, 0 = different books across vehicles)
                                llResearchPop = llPopEst
                            Else
                                If (llResearchPop <> 0) And (llResearchPop <> llPopEst) And (llPopEst <> 0) Then      'test to see if this pop is different that the prev one.
                                    llResearchPop = 0                                           'if different pops, calculate the contract  summary different
                                    If llPopEst <> lmPop(ilVehicle) Then  '11-30-99
                                        lmPop(ilVehicle) = -1          '11-30-99
                                    End If                             '11-30-99
                                Else
                                    'if current line has population, but there was already a different across
                                    'lines in pop, dont save new one
                                    If llPopEst <> 0 And (llResearchPop <> 0 And llResearchPop <> -1) Then   '2/1/99
                                        llResearchPop = llPopEst
                                    Else        '5-8-02
                                        If lmPop(ilVehicle) <> llPopEst And lmPop(ilVehicle) <> -1 Then
                                            lmPop(ilVehicle) = -1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If                      'blGotHiddenLineForPkg
                End If                          'imProcessFlag(tmLnr(ilLoop).iLineInx) = 1
            Next ilLoop                         'loop on all line entries
            
            'Line totals generated for max 2 years each, total all lines from results of each line
            'For Summary Page (1 line per vehicle)
            'ReDim tmVRtg(1 To 1) As Integer
            'ReDim tmVCost(1 To 1) As Long
            'ReDim tmVGRP(1 To 1) As Long
            'ReDim tmVGrimp(1 To 1) As Long
            ReDim tmVRtg(0 To 0) As Integer
            ReDim tmVCost(0 To 0) As Long
            ReDim tmVGRP(0 To 0) As Long
            ReDim tmVGrimp(0 To 0) As Long
            For llRchQtr = LBONE To UBound(tmLRch) - 1 Step 1
'                        If (ilVehList(ilVehicle) = tmLRch(ilRchQtr).iVefCode Or ilVehList(ilVehicle) = tmLRch(ilRchQtr).iPkvefCode) And tmLRch(ilRchQtr).lQSpots <> 0 Then
'                            If (tlBR.iShowProof And tmLRch(ilRchQtr).sType = "H") Or (tmLRch(ilRchQtr).sType = "S") Or (tmLRch(ilRchQtr).sType = "H" And ilVehList(ilVehicle) = tmLRch(ilRchQtr).iPkvefCode) Then
                        tmVRtg(UBound(tmVRtg)) = tmLRch(llRchQtr).iTotalAvgRating
                        'tmVCost(UBound(tmVCost)) = tmLRch(llRchQtr).lTotalCost
                        tmVCost(UBound(tmVCost)) = tmLRch(llRchQtr).dTotalCost 'TTP 10439 - Rerate 21,000,000
                        tmVGRP(UBound(tmVGRP)) = tmLRch(llRchQtr).lTotalGRP
                        tmVGrimp(UBound(tmVGrimp)) = tmLRch(llRchQtr).lTotalGrimps
                        ReDim Preserve tmVRtg(0 To UBound(tmVRtg) + 1) As Integer
                        ReDim Preserve tmVCost(0 To UBound(tmVCost) + 1) As Long
                        ReDim Preserve tmVGRP(0 To UBound(tmVGRP) + 1) As Long
                        ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) + 1) As Long
'                            End If
'                        End If
            Next llRchQtr
            'If UBound(tmVRtg) > 1 Then
            If UBound(tmVRtg) > 0 Then
                'dimensions must be exact sizes
                ReDim Preserve tmVRtg(0 To UBound(tmVRtg) - 1) As Integer
                ReDim Preserve tmVCost(0 To UBound(tmVCost) - 1) As Long
                ReDim Preserve tmVGRP(0 To UBound(tmVGRP) - 1) As Long
                ReDim Preserve tmVGrimp(0 To UBound(tmVGrimp) - 1) As Long
                'If UBound(tmVRtg) >= 1 Then
                If UBound(tmVRtg) >= 0 Then
                    'gResearchTotals True, llPop, tmVCost(), tmVRtg(), tmVGrimp(), tmVGRP(), llTemp, tmCbf.iAvgRate, tmCbf.lGrimp, tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM
                    '4/7/99
         
                    If tgSpf.sDemoEstAllowed = "Y" Then         '6-1-04
                        'gResearchTotals sm1or2PlaceRating, True, llOverallPopEst, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, llTemp, tmCbf.iAvgRate, tmCbf.lGrImp, tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, tmCbf.lAvgAud
                        gResearchTotals sm1or2PlaceRating, True, llOverallPopEst, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, dlTemp, tmCbf.iAvgRate, tmCbf.lGrImp, tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, tmCbf.lAvgAud 'TTP 10439 - Rerate 21,000,000
                    Else
'                                If lmPopPkg(ilVehicle) <> 0 Then        '11-24-04
'                                    llPop = lmPopPkg(ilVehicle)
'                                Else
'                                    llPop = lmPop(ilVehicle)   '11-23-99????
'                                End If
                        'gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, llTemp, tmCbf.iAvgRate, tmCbf.lGrImp, tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, tmCbf.lAvgAud
                        gResearchTotals sm1or2PlaceRating, True, llResearchPop, tmVCost(), tmVGrimp(), tmVGRP(), llLnSpots, dlTemp, tmCbf.iAvgRate, tmCbf.lGrImp, tmCbf.lGRP, tmCbf.lCPP, tmCbf.lCPM, tmCbf.lAvgAud 'TTP 10439 - Rerate 21,000,000
                    End If
                Else
                    tmCbf.lCPP = 0
                    tmCbf.lCPM = 0
                    tmCbf.lGRP = 0
                    tmCbf.lGrImp = 0
                End If


                tmCbf.iVefCode = ilVehList(ilLoopOnVehList)
                tmCbf.lQGRP = 0
                tmCbf.lQCPP = 0
                tmCbf.lQCPM = 0
                tmCbf.lQGrimp = 0
                tmCbf.lVQGRP = 0
                tmCbf.lVQCPP = 0
                tmCbf.lVQCPM = 0
                tmCbf.lVQGrimp = 0
                igBR_SSLinesExist = True      '12-16-03 force output
                tmCbf.iExtra2Byte = 2               'vehicle summary totals
                tmCbf.sMixTypes = ""            'default to show vehicle grp, cpp because not a podcast
                For ilTemp = 0 To UBound(tmPodcast_Info) - 1
                    If tmPodcast_Info(ilTemp).iVefCode = tmCbf.iVefCode Then
                        'type of line:  P = Podcast, K = package line, H = Podcast Hidden Line, L = Other, not podcast Hidden Line, O = other, not podcast (conventional, selling)
                        'look for package vehicles to determine how to show the vehicle cpp, grp columns
                        If tmPodcast_Info(ilTemp).sType = "K" Then
                            If Not tmPodcast_Info(ilTemp).bShowResearch Then
                                tmCbf.sMixTypes = "H"
                            End If
                            Exit For
                        Else
                            If tmPodcast_Info(ilTemp).sType = "P" Or tmPodcast_Info(ilTemp).sType = "H" Then        'podcast vehicle not in hidden line (P), or in pkg (H)
                                tmCbf.sMixTypes = "H"
                            End If
                            Exit For
                        End If
                    End If
                Next ilTemp
                ilVefInxForCallLetters = gBinarySearchVef(tmCbf.iVefCode)
                If ilVefInxForCallLetters <> -1 Then        '10-23-06 if new package vehicle entered for proposal, the index wont exist.  The vehicle
                    gFindStationMkt ilVefInxForCallLetters, tmCbf            '2-21-18 show market name?
                End If
                ilRet = btrInsert(hmCbf, tmCbf, imCbfRecLen, INDEXKEY0)
            End If                              'UBound(tmVRtg) > 0
        End If                                  'blItsAPkg
    Next ilLoopOnVehList                            'ilLoopOnVehList = 0 To UBound(ilVehList) - 1
    Exit Sub
End Sub

Public Sub mProcessBR_CPM(ilWhichSort As Integer, ilShowProof As Integer, tlCPM_IDs() As CPM_BR, tlCPMSummary() As CPM_BR)
    Dim ilRet As Integer
    If tgChf.sAdServerDefined = "Y" Then
        mSetupBrHdr ilWhichSort                                     'build the advt, agy, slsp and sort fields specs that ned to be built
        
        gWriteBR_CPM hmCbf, hmRdf, tgChf, tmCbf, ilShowProof, tlCPM_IDs(), tmCPMSummary()
    End If
    Exit Sub
End Sub

'8410 - get the Package Line # from a Hidden Line #
Function mGetPkgLineNoFromHiddenLine(ilHiddenLineNo As Integer) As Integer
    Dim ilLoop As Integer
    Dim ilClfInx As Integer
    mGetPkgLineNoFromHiddenLine = 0
    For ilLoop = 1 To UBound(tmLnr) - 1
        ilClfInx = tmLnr(ilLoop).iLineInx            'line index to process
        tmClf = tgClf(ilClfInx).ClfRec
        If tmClf.iLine = ilHiddenLineNo Then
            mGetPkgLineNoFromHiddenLine = ilClfInx
            Exit For
        End If
    Next ilLoop
End Function

