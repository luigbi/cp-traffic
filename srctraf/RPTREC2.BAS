Attribute VB_Name = "RPTREC2"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of RPTREC2.BAS on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptRec2.BAS    Report Definitions
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions for crystal intermediate files
Option Explicit
Option Compare Text
'********************************************************
'
'Contract Report (BR) file definition
'
'   2/18/00 Add 13 week grp buckets by vehicle & contract,
'           change unused from 15 to 20
'
'*********************************************************
'Cbf record layout
Type CBF

    'iGenTime(0 To 1) As Integer   'Generation Time by budget
    iGenDate(0 To 1) As Integer   'Generation Date by budget
    '10-10-01                   'Date & time were swapped, make date first, followed by time to be
                                'consistent with all other prepass records
    lGenTime As Long            'generation timelChfCode As Long        'Contract Code
    lChfCode As Long
    lLineNo As Long         'Line #
    iVefCode As Integer    'Vehicle code
    sDysTms As String * 80  '11-20-02 Dates/Time, changed from 40 to 80 on 12/9/02
    iDtFrstBkt(0 To 1) As Integer   'Date of first bucket
    'lWeek(1 To 14) As Long      'Spot count (or rates) for week
    lWeek(0 To 13) As Long      'Spot count (or rates) for week
    iLen As Integer             'Spot Length
    lRate As Long               'Rate
    lGRP As Long                'GRPS
    lCPP As Long                'CPP
    lCPM As Long                'cpm
    lAvgAud As Long
    iMnfGroup As Integer        'Vehicle Group sort
    iVehSort As Integer         'Vehicle sort
    iRdfDPSort As Integer       'Daypart sort
    iAirWks As Integer          'Actual airing weeks
    iTotalWks As Integer        'total contracts weeks
    sSortField1 As String * 20  'Sort field 1
    sSortField2 As String * 20  'sort field 2
    iPctDist As Integer         '% distribution
    lContrNo As Long            'contract number
    iAvgRate As Integer         'Average rating
    lGrImp As Long              'Gross impressions
    sSurvey As String * 30      'Survey book names from all sch lines
    lIntComment As Long         'internal comment code if Show = yes
    lCancComment As Long        'cancellation comment code if show = yes
    lChgRComment As Long        'change reason comment code if show = ye
    lOtherComment As Long       'other comment code if show = yes
    lMerchComment As Long       'merchandising comment code
    lPromoComment As Long       'promotions comment code
    lLineComment As Long        'line comment
    iOurShare As Integer        '% of buy for "us" (xxx)
    lDrfCode As Long            'demo research code
    iDnfCode As Integer         'research book name
    lPop As Long                'population
    'lValue(1 To 14) As Long     'avg audience
    lValue(0 To 13) As Long     'avg audience
    sDemos As String * 30       'max 4 demo names from hdr
    sLineSurvey As String * 30  'line survey name
    iStartQtr(0 To 1) As Integer   'Start date of first qtr oforder
    iTotalMonths As Integer     'total # airing months of order
    'lMonth(1 To 13) As Long   'gross $ per month, anything over 12 month all goes in 13th bucket
    lMonth(0 To 12) As Long   'gross $ per month, anything over 12 month all goes in 13th bucket
    'lMonthUnits(1 To 13) As Long 'total spots per std month, anything over 12 months goes into 13th bucket
    lMonthUnits(0 To 12) As Long 'total spots per std month, anything over 12 months goes into 13th bucket
    lCntGrimps As Long          'total grimps for entire contract
    'iCurrModSpots As Integer   'total spot count for current version
    iOBBLen As Integer          '3-28-05 chg from unused to Open bb len
    lCurrMod As Long           'total $ for current version
    iExtra2Byte As Integer      ' (br type 0 = detail, 2 = vehicle totals entire cnt, 3 = cnt totals)
    lExtra4Byte As Long         '
    lQGRP As Long               'grps for all veh in qtr
    lQCPP As Long               'cpps for all veh in qtr
    lQCPM As Long               'cpms for all veh in qtr
    lQGrimp As Long             'grimps for all veh in qtr
    lVQGRP As Long              'grps for one veh by qtr
    lVQCPP As Long              'cpps for one veh by qtr
    lVQCPM As Long              'cpms for one veh by qtr
    lVQGrimp As Long            'grimps for one veh by qtr
    iAdfCode As Integer         'advt code
    iAgfCode As Integer         'agency code
    iSlfCode As Integer         'slsp code
    iPropVer  As Integer        'Proposal version #
    iCntRevNo As Integer        'Internal Rev #
    iExtRevNo As Integer     'External Rev #
    iPropOrdDate(0 To 1) As Integer 'Proposal or Order date entered
    iPropOrdTime(0 To 1) As Integer 'Proposal or Order time entered
    iStartDate(0 To 1) As Integer   'Contract start date
    iEndDate(0 To 1) As Integer     'Contract end date
    sType As String * 1             'Contract type
    sStatus As String * 1           'contract status
    iPctTrade As Integer            'Percent of trade (1 - 100)
    sProduct As String * 35         'contract product description
    sBuyer As String * 20           'contract buyer
    sAgyCTrade As String * 1        'Agency commissionable on trade (Y/N)
    sSnapshot As String * 1         'S = snapshot (not saved), P = snapshot (previous version), C = (snapshot, curent version), blank = from reports
    'iRemoteID As Integer        'cntr# = cntr#:remoteID
    sUnused As String * 2       'Unused
    sCBS As String * 1          'C = Cancel before start line
    sLineType As String * 1     'Line type (A,O,E = package, H = hidden, S = standard
    sResort As String * 4       '4/14/99 sorting for package/hidden & conventionals:  pkg = xxxA, hidden = xxxC, conv = 999~
                                'xxxx = line # of package
    'lWkVehGrp(1 To 14) As Long  '2-18-00 weekly vehicle grps for a quarter
    lWkVehGrp(0 To 13) As Long  '2-18-00 weekly vehicle grps for a quarter
    'lWkCntGrp(1 To 14) As Long  '2-18-00 weekly contract grps for a quarter
    lWkCntGrp(0 To 13) As Long  '2-18-00 weekly contract grps for a quarter
    iTime(0 To 1) As Integer    'Generic Time by budget
    iProdPct As Integer         '1-16-01
    sPriceType As String * 1    '2-23-01 added for sch line price type description (ADU, recap, bonus, etc)
    lCurrModSpots As Long   'total spot count for current version
    lCxfComment As Long     'Insertion or contract comment
    sDailyExists As String * 1   '5-21-03 Y if at least one daily line exists on this contract
    sDailyWkly As String * 1     '5-20-03 0 = no daily exist, line is wkly, 1 = dailies exist, line is daily, 2 = dailies exist, line is weekly
    iCBBLen As Integer          '3-28-05 chg unused from 3 to 1 Closed BB length
    sResortType As String * 1   '5/31/05 Letter to separate Package from Hidden from Conventional
                                '5-20-03 chg from 5 to 3 unused
                                'unused  (4/14/99 chg from 19 to 15), 2-18-00 chged from 15 to 20
                                'chged from 20 to 16, added generic time field)
                                'unused (1/16/01 chg from 16 to 14) add prodpct
                                'unused (2/23-01 chg from 14 to 13) to add price type description
    s1stPosition As String * 1  '1st Position (Y or N)
    sSoloAvail As String * 1    'Solo Avail (Y or N)
    sPrefDT As String * 40      'Preferred Days-Time
    sHiddenOverride As String * 1 '3-10-06 associated with feature in Site
    lBarterCefCode As Long      'Barter insertion comment
    lRafCode As Long            'show regions as line remarks 8-30-06
    lTax1 As Long
    lTax2 As Long
    sWeeksInQtr As String * 1   '3=13 weeks in Qtr; 4=14 weeks in Qtr
    sPkgLnRateOnBr        As String * 1      ' Package Line flight Rate on
                                             ' Contract printout (Y/N).  Test
                                             ' for Y
    iUrfCode As Integer       'User Reference code
    sAudioType As String * 1    'Audio type from sch line override box
    lCashCommAmt As Long    'Cash Commission Amount
    lTradeCommAmt As Long    'Trade Commission Amount
    lPkLnComment As Long    'Package Line Comment
    iAffMktCode As Integer  'Market code obtained from Station (when station are vehicles)
    'sUnused               As String * 2     ' Unused,
    sRschColHdr As String * 1   'Research Columns Title Headers (avg rrating; GRP and Cpp): Blank = Show; Non-Blank = Hide
    sMixTypes As String * 1     'Resaerch Values within the Detail lines. Blank = Show Value; Non-Blank = Hide Value
End Type
'Cbf key record layout- use INTKEY0
Type CBFKEY0
    iGenDate(0 To 1) As Integer   'Generation Date
    '10-10-01
    lGenTime As Long
    'iGenTime(0 To 1) As Integer   'Generation Time
End Type

Type CBFKEY1
    iGenDate(0 To 1) As Integer   'Generation Date
    '10-10-01
    lGenTime As Long
    'iGenTime(0 To 1) As Integer   'Generation Time
    iUrfCode As Integer       'User Reference code
End Type

Type VEHQTRLIST                 'return values from vehicle by qtr for research
    'lVQGRP(1 To 8) As Long
    'lVQCPP(1 To 8) As Long
    'lVQCPM(1 To 8) As Long
    'lVQGrimp(1 To 8) As Long
    lVQGRP(0 To 8) As Long  'Index zero ignored
    lVQCPP(0 To 8) As Long  'Index zero ignored
    lVQCPM(0 To 8) As Long  'Index zero ignored
    lVQGrimp(0 To 8) As Long  'Index zero ignored
End Type
'Line Rate structure
Type LNR
    iLineInx As Integer         'Line index # built from mObtain Cnt
    lRate As Long               'flight rate
    'sValidDays As String * 7    'Y for valid airing day, else N
    sValidDays As String * 40   '5-20-03 chged from 20 to 30 for 4 char/day (plus blank) on dailies
                                '11-20-02 Y for valid airing day, else N
    iCffDays(0 To 6) As Integer   'bit map days of the week from flight
    'iLRchInx As Integer           'index into research arrays (LRch)
    lLRchInx As Long                '4-10-19 subscript out of range
    iManyFlts As Integer          'false = only 1 flight this line, true = more than 1 flight this line
    sPriceType As String * 1       '2-23-01  sch line price type (adu, recap, bonus, etc)
    sDailyWkly As String * 1        '5-20-03 D = daily, W = weekly .  flag to test for output on printed contract
End Type
Type WEEKLYGRPS                   '2-11-00
    iVefCode As Integer
    iPkLineNo As Integer          '1-24-03
    'lGrps(1 To 105) As Long       '2 years weekly gross rating points for vehicle & contract weekly totals
    lGrps(0 To 105) As Long       '2 years weekly gross rating points for vehicle & contract weekly totals. Index zero ignored
End Type

Type INSERTINFO                 '12-12-08 for Insertion Orders to get the barter comment code
    iVefCode As Integer
    lBarterCefCode As Long
    iDtFrstBkt(0 To 1) As Integer          '6-30-15 start of quarter
End Type

'5-20-15 List of unique vehicle/contract for separate Insertion Order pdfs to be emailed
Type EMAIL_PDFS
    iVefCode As Integer
    sVefName As String * 40
    lCntrNo As Long
    lChfCode As Long
    sStartDate As String * 10   'contract start date
    sEndDate As String * 10     'contract end date
    sAdvtName As String * 30
    iCntRevNo As Integer
    sAgyEstNo As String * 20        '8-20-15 chfagyestno + chftitle
    sProduct As String * 35
    lEmfCode As Long
    sResponseDate As String * 10
End Type

Type PODCAST_INFO                       '3-9-18
    iVefCode As Integer
    iLine As Integer                    'sch line #
    sType As String * 1                 'type of line:                                 'type of line:  P = Podcast, K = package line, H = Podcast Hidden Line, L = Other, not podcast Hidden Line, O = other, not poscast (conventional, selling)
    iPkgRefLine As Integer              'pkg line reference, else 0
    bShowResearch As Boolean            'indicator whether to show research info or not:  true = show avg rtg, cpp or grp values
End Type

Public tgEmail_PDFs() As EMAIL_PDFS         '5-20-15
Public tgAllStations() As SHTTINFO          '2-26-18
Public sgAllStationsStamp As String

'TTP 10549 - Learfield Cloud printing 911, Crystal Crashes, Use Adobe
Public sgReportFilename As String
Public sgReportTempFolder As String
Public bgUseAdobe As Boolean

'TTP 10549 - Learfield Cloud printing 911, Crystal Crashes, Use Adobe
Public Sub ShellAdobeAndDelete(slPDFFileName As String)
    On Error Resume Next
    Dim oShell As Object
    Dim ilRet As Integer
    If InStr(1, slPDFFileName, ".pdf") = 0 Then
        slPDFFileName = slPDFFileName & ".pdf"
    End If
    
    'Shell PDF file (open with Adobe viewer)
    Set oShell = CreateObject("Wscript.Shell")
    ilRet = oShell.Run("""" & slPDFFileName & """", 3, True)

    'PDF is now Closed...  Now Delete the temp pdf file(s)
    slPDFFileName = Replace(slPDFFileName, ".pdf", "*.pdf")
    Kill slPDFFileName
    
    sgReportFilename = "" 'Reset filename so that next report ran should set this var to be sent to Adobe
    On Error GoTo 0
End Sub

