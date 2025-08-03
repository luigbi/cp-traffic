Attribute VB_Name = "RPTREC"

'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptrec.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptRec.BAS    Report Definitions
'
' Release: 1.0
'
' Description:
'   This file contains the record definitions for crystal intermediate files
'  All Temporary prepass files should be defined in Vardef.bas:
'           sgTdbNames - increase dimension a (ie. jsr,cbf, cpr, grf, anf, etc)
'           Initsubs.bas - procedure gInitGlobvar - (add the filename)
'           this table is used in crpevar.bas to use temporary directory
'           name instead of database name (gOpenPrtJob)
'
'           3/15/98 Change size of avrDPDays from 7 to 20
'           8/5/98 Added new snapshot fields to CBF
'           10/1/98 Added major & minor sort codes & unused to AVR
'           11/19/98 Added IVR file definitions
'           5/21/99 Added more fields to cnttypes structure
'           1-30-00 Reduce unused from 20 to 18, insert majorvgrp 2 bytes
'           7-5-00 AVR.btr: increase quarterly buckets from 13 to 14
'                   (overall record size increased)
Option Explicit
Option Compare Text

Public tgRptSelAdvertiserCode() As SORTCODE

Type SEQSORTTYPE
    sKey As String * 20
    lSdfCode As Long
    iSeqNo As Integer
End Type

'3-25-14 sdfsortbyline & spottypes moved from rptextra.bas
Type SDFSORTBYLINE
    sKey As String * 20         'line ID
    tSdf As SDF
End Type

Type IBFSORTBYLINE
    sKey As String * 20         'line ID
    tIbf As IBF
End Type

Type SPOTTYPES
    iSched As Integer
    iMissed As Integer
    iMG As Integer
    iOutside As Integer
    iHidden As Integer
    iCancel As Integer
    iFill As Integer
    iOpen As Integer        '1-18-11
    iClose As Integer       '1-18-11
End Type


'7-6-15 moved from rptgen.bas so modules can be removed from traffic
Type TYPESORT
    sKey As String * 100
    lRecPos As Long
End Type

'******************************************************************************
' afr Record Definition
'
'******************************************************************************
Type AFR
    iGenDate(0 To 1)      As Integer         ' Generation date
    lGenTime              As Long            ' Generation time
    lAstCode              As Long            ' Ast code
    iRegionCopyExists     As Integer         ' Regional copy exists for spot
    sCart                 As String * 6      ' Region cart (media code +
                                             ' inventory #)
    sProduct              As String * 35     ' Region Product
    sISCI                 As String * 20     ' Region ISCI code
    sCreative             As String * 30     ' Region creative title
    lCrfCsfcode           As Long            ' regional copy script or comment
    lAttCode              As Long            ' Agreement code
    sCompliant            As String * 1      ' Spot Compliant. Y or N.  Spot air
                                             ' date and time within pledge date
                                             ' and time
    iSeqNo                As Integer         ' Sequence number
    sAdfName              As String * 30     ' Advertiser name
    sProdName             As String * 35     ' Contract Product Name
    iMissReplDate(0 To 1) As Integer         ' Missed Replacement date
    iMissReplTime(0 To 1) As Integer         ' Missed Replacement Time
    iMissedMnfCode        As Integer         ' Missed Reason reference code
    sLinkStatus           As String * 1      ' Link status for Missed spot with
                                             ' MG and Replaced spot with
                                             ' Replacement (M=Missed with MG,
                                             ' R=Replaced with Replacement)
    sCallLetters          As String * 10     ' Call Letters plus band
    iPledgeDate(0 To 1)   As Integer         ' Pledge date
    iPledgeStartTime(0 To 1) As Integer      ' Pledge start time
    iPledgeEndTime(0 To 1) As Integer        ' Pledge end time
    iPledgeStatus          As Integer        ' pledge status
    sSplitNet              As String * 1     ' Split Network (blank = no , p = primary, s = secondary, test for P or S
    lID                    As Long           ' line #
    iLen                   As Integer        ' spot length
    sSpotType              As String * 1     ' spot id (0 = sch; 1-mg;2-filled,3-outside, 4=?, 5 = added; 6 = open bb, 7=closebb
    sUnused               As String * 12      ' Unused
End Type


Type AFRKEY0
    iGenDate(0 To 1)      As Integer
    lGenTime              As Long
End Type
'
'********************************************************
'
'Import Contract file definition
'
'*********************************************************
Type ICF
    iDate(0 To 1) As Integer    'Import system date
    iTime(0 To 1) As Integer    'Import system time
    iSeqNo As Integer           'Sequence number to kept records in order
    sType As String * 1         '0=Import file definition; 1=Header; 2=Line/flight; 9=Total
    lCntrNo As Long             'Contract number
    sAdvtName As String * 30    'Advertiser name (header only)
    sProduct As String * 20     'Product (header only)
    sGross As String * 6        'Gross amount for header
    iLineNo As Integer          'Line number
    iFlightNo As Integer        'Flight
    sFlightDates As String * 17 'Flight dates xx/xx/xx-xx/xx/xx
    sStatus As String * 1       'Status of input: A=Accepted; T=Contract totally rejected
    sErrorMess As String * 60   'Error message if contract rejected
    iUrfCode As Integer         'User code
End Type
'********************************************************
'
'Copy Report file definition
'
'*********************************************************
Type CPR
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long            'rept gen time
    'iGenTime(0 To 1) As Integer 'Generation Time
    iVefCode As Integer
    iSpotDate(0 To 1) As Integer 'Spot Date
    iSpotTime(0 To 1) As Integer 'Spot Time
    iAdfCode As Integer         'Advertiser Code
    lCntrNo As Long             'Contract #
    iLineNo As Integer          'Line #
    iLen As Integer             'Spot Length
    sProduct As String * 35     'Contract Product or Copy Product
    sZone As String * 3         'Zone
    sCartNo As String * 12
    sISCI As String * 20
    sCreative As String * 30
    sStatus As String * 1       'S=Scheduled
    iRemoteID As Integer        'cntr# = cntr#:remoteID
    iMissing As Integer         '12-7-99 (from unused)
    iUnassign As Integer        '12-7-99 (from unused)
    iReady As Integer           '12-7-99 (from unused)
    lHd1CefCode As Long         'VOF header comments
    lFt1CefCode As Long         'vof footer comments 1 of 2
    lFt2CefCode As Long         'vof footer comments 2 of 2
    sLive As String * 1         '11-16-05   live flag for spot
    sUnused As String * 1     'unused 12-7-99 chged from 20 to 14; 6-20-00 unused chged from 14 to 2
                                '11-16-05 chgned from 2 to 1 for live flag
End Type
'Cpr key record layout
Type CPRKEY0
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long
    'iGenTime(0 To 1) As Integer 'Generation Time
End Type
'********************************************************
'
'Spot Projection file definition
'
'*********************************************************
Type JSR
    iGenDate(0 To 1) As Integer 'Generation Date
    lGenTime As Long            '10-10-01
    'iGenTime(0 To 1) As Integer 'Generation Time
    iPdStartDate(0 To 1) As Integer 'First Period Start Date
    sPdType As String * 1   'W=Weekly; S=Standard; F=Corporate, C = calendar
    lCntrNo As Long         'contract #
    iAdfCode As Integer
    iVefCode As Integer
    iSlfCode As Integer
    sCorTFlag As String * 1 'C=Cash Record; T=Trade Record
    iCorTPct As Integer     '% Cash or % Trade (xxx)
    'lSchDollars(1 To 13) As Long       'Scheduled Dollar Buckets (xx,xxx,xxx.xx)
    'lMGODollars(1 To 13) As Long       'MG and Outside Dollar Buckets (xx,xxx,xxx.xx)
    'lMCHDollars(1 To 13) As Long       'Missed; Cancelled; Hidden Dollar Buckets (xx,xxx,xxx.xx)
    lSchDollars(0 To 12) As Long       'Scheduled Dollar Buckets (xx,xxx,xxx.xx)
    lMGODollars(0 To 12) As Long       'MG and Outside Dollar Buckets (xx,xxx,xxx.xx)
    lMCHDollars(0 To 12) As Long       'Missed; Cancelled; Hidden Dollar Buckets (xx,xxx,xxx.xx)
    lCntrCode As Long                   'contract code
    sAgyCTrade As String * 1            'if trade, Y if comm, else N
    iAgfCode As Integer                 'agy code (to know if commissionable or direct)
    iRemoteID As Integer        'cntr# = cntr#:remoteID
    iMajorVGrp As Integer       '1-30-00
    sProjType As String * 1     'A = air time, N = NTR
    sComm As String * 1         'Commissionable for NTR
    sUnused As String * 16     'unused chg from 18 to 16 6-28-07
End Type
'Jsr key record layout
Type JSRKEY0
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long            'generation time
    'iGenTime(0 To 1) As Integer 'Generation Time
End Type
'********************************************************
'
'Quarterly Avails file definition
'
'   8/10/99 DH Add Proposal counts for 30/60
'
'*********************************************************
Type AVR
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long            'generation time
    'iGenTime(0 To 1) As Integer 'Generation Time
    iVefCode As Integer
    iDay As Integer             '0=Monday, 1=Tuesday,..6=Sunday
    iQStartDate(0 To 1) As Integer 'Start Date of quarter
    iFirstBucket As Integer     '1 thru 13, first bucket with data
    sBucketType As String * 1   'I=Inventory, A=Avail, S=Sold
    iRdfCode As Integer     'RdfCode (in DDF called avrrpfCode, Currently rdfSortCode stored into this field, 12/11/96)
    sInOut As String * 1    'N=N/A; I=Book into avail specified; O=Exclude avail specified
    ianfCode As Integer     'Avail name code if sInOut = I or O
    iDPStartTime(0 To 1) As Integer 'Daypart start time
    iDPEndTime(0 To 1) As Integer   'Daypart end time
    sDPDays As String * 40       '11-22-02 chg from 20 to 40,Daypart days: Y or N for each day for Monday thru Sunday
    sNot30Or60 As String * 1    'For Match Units only (Y or N, N=avails or spot of length not equal to 30 or 60)
'    i30Count(1 To 14) As Integer   'Count of Inventory or Avails or Sold for 30sec (for qtrly detail this contains sold values)
'    i60Count(1 To 14) As Integer   'Count of Inventory or Avails or Sold for 60sec (for qtrly detail this contains sold values)
'    i30InvCount(1 To 14) As Integer   'Count of Inventory for 30sec
'    i60InvCount(1 To 14) As Integer   'Count of Inventory for 60sec
'    'the remaining variables are for the qtrly detail report
'    i30Hold(1 To 14) As Integer     'count of hold units sold for 30sec
'    i60Hold(1 To 14) As Integer     'count of hold units sold for 60sec
'    i30Reserve(1 To 14) As Integer     'count of reserve units sold for 30sec
'    i60Reserve(1 To 14) As Integer     'count of reserve units sold for 60sec
'    i30Avail(1 To 14) As Integer     'count of available units sold for 30sec
'    i60Avail(1 To 14) As Integer     'count of available units sold for 60sec
'    iRdfSortCode As Integer         'Currently sort code is stored into rdfcode (12/11/96)
'    lRate(1 To 14) As Long       'holds weekly rates from rif.btr
'    lMonth(1 To 3) As Long       'holds monthly rates from the 13 week buckets
'    imnfMajorCode As Integer     'major mnf sort code
'    imnfMinorCode As Integer     'minor mnf sort code
'    i30Prop(1 To 14) As Integer  'count of 30 proposals
'    i60Prop(1 To 14) As Integer  'count of 60 proposals
'    iWksInQtr As Integer         'wks in quarter, used for Crystal to show date headers across quarters
'    sUnused As String * 20       'unused
    i30Count(0 To 13) As Integer   'Count of Inventory or Avails or Sold for 30sec (for qtrly detail this contains sold values)
    i60Count(0 To 13) As Integer   'Count of Inventory or Avails or Sold for 60sec (for qtrly detail this contains sold values)
    i30InvCount(0 To 13) As Integer   'Count of Inventory for 30sec
    i60InvCount(0 To 13) As Integer   'Count of Inventory for 60sec
    'the remaining variables are for the qtrly detail report
    i30Hold(0 To 13) As Integer     'count of hold units sold for 30sec
    i60Hold(0 To 13) As Integer     'count of hold units sold for 60sec
    i30Reserve(0 To 13) As Integer     'count of reserve units sold for 30sec
    i60Reserve(0 To 13) As Integer     'count of reserve units sold for 60sec
    i30Avail(0 To 13) As Integer     'count of available units sold for 30sec
    i60Avail(0 To 13) As Integer     'count of available units sold for 60sec
    iRdfSortCode As Integer         'Currently sort code is stored into rdfcode (12/11/96)
    lRate(0 To 13) As Long       'holds weekly rates from rif.btr
    lMonth(0 To 2) As Long       'holds monthly rates from the 13 week buckets
    imnfMajorCode As Integer     'major mnf sort code
    imnfMinorCode As Integer     'minor mnf sort code
    i30Prop(0 To 13) As Integer  'count of 30 proposals
    i60Prop(0 To 13) As Integer  'count of 60 proposals
    iWksInQtr As Integer         'wks in quarter, used for Crystal to show date headers across quarters
    sUnused As String * 20       'unused
End Type
'Avr key record layout
Type AVRKEY0
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long            'generation time
    'iGenTime(0 To 1) As Integer 'Generation Time
End Type
'*****************************************************************************
'
'           Generic Report Record Definitions
'           Modified 12/3/98:  remove 4 bytes from unused,
'                   replace with generic time field
'                   6/16/00 remove 4 bytes from unused GRF(unused 16 to 12),
'                   replace with another time field
'           8-27-01 replaced unused with 4 byte long
'******************************************************************************
Type GRF
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long             'rept gen time
    'iGenTime(0 To 1) As Integer 'Generation Time
    iVefCode As Integer         'vehicle code
    iSofCode As Integer         'Office code
    iSlfCode As Integer         'slsp code
    iRdfCode As Integer         'DP code
    iAdfCode As Integer         'advt code
    lChfCode As Long            'contr # or code
    iStartDate(0 To 1) As Integer   'Start date of budget record
    sDateType As String * 1     'date type (general field)
    iYear As Integer
    sBktType As String * 1      'generic field
    iCode2 As Integer           'generic code pointer
    lCode4 As Long              'generic code pointer 4 bytes
    iDate(0 To 1) As Integer    'generic date field
    sGenDesc As String * 40     'generic string field
    'lDollars(1 To 18) As Long '$ in period 1-14  (addl 4 for qtr totals)
    lDollars(0 To 17) As Long '$ in period 1-14  (addl 4 for qtr totals)
    'iPerGenl(1 To 18) As Integer    '14 generic fields to match dollar buckets
    iPerGenl(0 To 17) As Integer    '14 generic fields to match dollar buckets
                                   'anything you want as integers (addl 4 for qtrs)
    'iDateGenl(0 To 1, 1 To 18) As Integer   '14 generic dates
    iDateGenl(0 To 1, 0 To 17) As Integer   '14 generic dates
    iRemoteID As Integer        'cntr# = cntr#:remoteID
    'sUnused As String * 20      'unused , replaced 12/3/98 with Time/unused
    iTime(0 To 1) As Integer    'generic time field
    iMissedTime(0 To 1) As Integer    'missed time
    lLong As Long               'another long variable
    'sUnused As String * 12      'unused, chged 6-16-00 to 12(from 16) for new time field
    sUnused As String * 8        'unused chged 8-27-01 from 12 to 8 for new long field
End Type
'Grfkey record layout
Type GRFKEY0
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long            'rept gen time
    'iGenTime(0 To 1) As Integer 'Generation Time
End Type
''printed Invoice structure
'Type IVR
'    iGenDate(0 To 1) As Integer 'generation date (key)
'    iGenTime(0 To 1) As Integer 'generation time (key)
'    lSpotKeyNo As Long          'sequential spot #
'    iType As Integer            '0 = spot, 1 = bonus, 2 = subtotal, 3 = total
'    sTitle1 As String * 15      'Title Line 1 of 4
'    sTitle2 As String * 15      'Title Line 2 of 4
'    sTitle3 As String * 15      'Title Line 3 of 4
'    sTitle4 As String * 15      'Title Line 4 of 4
'    lChfCode As Long            'Contract code
'    iInvDate(0 To 1) As Integer 'Invoice date
'    lInvNo As Long              'Invoice #
'    sTerms As String * 30       'Terms
'    iShowInvType As Integer     'Invoice type message
'    sCashTrade As String * 1    'C = cash, t = trade
'    iCTSplit As Integer         'cash/trade split percent
'    sAddr1 As String * 40       'agy/advt address line 1 of 5
'    sAddr2 As String * 40       'agy/advt address line 2 of 5
'    sAddr3 As String * 40       'agy/advt address line 3 of 5
'    sAddr4 As String * 40       'agy/advt address line 4 of 5
'    sAddr5 As String * 40       'agy/advt address line 5 of 5
'    sPayAddr1 As String * 40    'pay addr line 1 of 4
'    sPayAddr2 As String * 40    'pay addr line 2 of 4
'    sPayAddr3 As String * 40    'pay addr line 3 of 4
'    sPayAddr4 As String * 40    'pay addr line 4 of 4
'    iLineNo As Integer          'Line #
'    sOVehName As String * 20    'Ordered vehicle name
'    sODPName As String * 20     'Ordered daypart name
'    iLen As Integer             'spot length
'    sODays As String * 20       'Ordered days (m-s, s-s....)
'    iWkNo As Integer            'week #
'    iONoSpots As Integer        'ordered # spots
'    sORate As String * 14       'Ordered rate
'    sADayDate As String * 20    'Aired day, date of spot
'    sATime As String * 12       'aired time of spot
'    sARate As String * 14       'aired spot rate
'    sACopy1 As String * 45      'isci, product for spot (1 of 4)
'    sACopy2 As String * 45      'isci, product for spot (2 of 4)
'    sACopy3 As String * 45      'isci, product for spot (3 of 4)
'    sACopy4 As String * 45      'isci, product for spot (4 of 4)
'    sAVehname As String * 20    'Aired Vehicle name
'    sRRemark As String * 20     'reconciliation remark
'    sRAmount As String * 14     'reconciliation rate (xxxxx.xx)
'    iPctComm As Integer         '% commission
'    iOTotalSpots As Integer     'Ordered Total Spots
'    lOTotalGross As Long        'Ordered Total Gross
'    iATotalSpots As Integer     'Aired Total # spots
'    lATotalGross As Long        'Aired Total gross
'    lRTotalGross As Long        'reconciliation total gross
'    lComment1 As Long           'comment 1 of 4
'    lComment2 As Long           'comment 2 of 4
'    lComment3 As Long           'comment 3 of 4
'    lComment4 As Long           'comment 4 of 4
'    sEDIComment As String * 60  'EDI comment
'    sKey As String * 160        'Sort key for air invoice
'    sUnused As String * 30      'unused
'End Type
'Struture of arrays built for Tie Out report
Type TIEOUT
    sType As String * 1                 ' V = vehicle, O = office
    iCode As Integer                    'code of vehicle or office
                                        'the following arrays have 12 months plus 1 total year, plus 4 qtrs
    'Index zero ignored with each of the arrays below
    lTYPlan(0 To 17) As Long            'This years direct plan
    lTYSplitPlan(0 To 17) As Long       'This years split plan
    lTYOrders(0 To 17) As Long          'This years orders figures
    lTYNewBus(0 To 17) As Long          'this years new busines figures
    lLYOrders(0 To 17) As Long          'last years business figures
    lSplits(0 To 17) As Long            'This years total split contracts
    lSplitsIn(0 To 17) As Long          'this years split in figures
    lSplitsOut(0 To 17) As Long         'this years split out figures
    lOOBSplits(0 To 17) As Long
End Type
'Type CTFLIST
'    CtfRec As CTF                'image of contract summary records
'End Type
'Receivables report record layout
Type RVR
    lCode As Long       'auto increment
    iAgfCode As Integer 'Agency code number (0 if advertiser is direct)
    iAdfCode As Integer 'Advertiser code number
    lPrfCode As Long    'Product code
    iSlfCode As Integer 'Salesperson code number (if direct: default to advt otherwisw default to agency)
    lCntrNo As Long     'Contract Number
    lInvNo As Long      'Invoice number
    lRefInvNo As Long   'Reference invoice number
    iAirVefCode As Integer 'Vehicle Code of Airing vehicle
    lUnused2 As Long    'was Check number
    iTranDate(0 To 1) As Integer  'Transaction Date of Rate Card or zero if not superseded
    sTranType As String * 2   '
    sAction As String * 1   '
    sGross As String * 6 'Gross amount (xx,xxx,xxx.xx)
    sNet As String * 6   'Net amount (xx,xxx,xxx.xx)
    iAgePeriod As Integer   'Aging period
    iAgingYear As Integer   'Aging year
    sCashTrade As String * 1    'C=Cash; T=Trade
    iPurgeDate(0 To 1) As Integer  'Purge Date or zero if not purged
    iUrfCode As Integer 'Last user who modified receivable
    iBillVefCode As Integer    'Vehicle code of billing vehicle
    iPkLineNo As Integer       'Package line no used to combine transactions
    iInvDate(0 To 1) As Integer 'Invoice date
    iDateEntrd(0 To 1) As Integer   'Date entered
    iRemoteID As Integer        'cntr# = cntr#:remoteID
    iMnfGroup As Integer
    lTax1 As Long          'sales tax #1 1-3-02
    lTax2 As Long          'sales tax #2 1-3-02
    iMnfItem As Integer     '9-16-02 billing type type
    sInCollect As String * 1 '10-07-02
    lCefCode As Long        '10-14-02 Transaction Comment (for On Account)
    lSbfCode As Long        'Pointer to SBF if REP or NTR invoicing (added 12/13/02)
    lAcquisitionCost As Long    'Acquisition Cost (xxxxxx.xx)
    iBacklogTrfCode As Integer  'Tax table code for backlog NTR
    sType As String * 1         'Type of record: Blank = Aired and NTR dollars for contracts w/o Installment;
                                'I=Installment amount from contract; A=Air and NTR amount for contract with installment
    sInvoiceUndone As String * 1    'Invoice Undone (N/Y).  Test for Y
    iPnfBuyer As Integer
    lGsfCode              As Long            ' Sport Event Reference code
    sCheckNo              As String * 10     ' Check Number. replaced Long with string
    lPcfCode As Long            ' Podcast Ad Server contract reference
                                             ' code
    sUnused               As String * 6     ' Unused
    imnfVefGroup As Integer '7-11-02 reduced unused to 8 and inserted
    'sUnused As String * 4      '9-16-02 chg from 8 to 5,7-11-02 reduced from 10 to 8, unused, 1-3-02 chged from 20 to 12 for sls tax amts
    '                            '10-7-02 chged from 5 to 4
    iGenDate(0 To 1) As Integer 'Report generation date
    '10-10-01
    lGenTime As Long            'generation time
    'iGenTime(0 To 1) As Integer 'report generation time
    sSource As String * 1       'H = came from PHF (billed internal), X = came from PHF (billed external), R = came from RVF (billed internal)
    imnfOwner As Integer        'group owner code
    iMnfSSCode As Integer       'Sales source code for participant
    'imnfGroup As Integer        'Vehicle group code for paticipant
    iProdPct As Integer         'participant % (xxx.xx)
    lDistAmt As Long            'distribution amt (xxxxxxxx.xx)
End Type
' Receivables record layout
Type RVRKEY0
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long            'generation time
    'iGenTime(0 To 1) As Integer 'Generation Time
End Type

Type SCR
    lCode As Long                   'auto code
    iGenDate(0 To 1) As Integer     'Report generation date
    lGenTime As Long                'report generation time
    'iStrLen As Integer  'String length (required by LVar)
    'sScript As String * 5002   'Last two bytes after the comment must be 0
    sScript As String * 5004   'Last bytes after the comment must be 0
End Type

Type SCRKEY1
    iGenDate(0 To 1) As Integer   'Generation Date
    lGenTime As Long
End Type

'Seven Day Report record layout (mainly for spots)
Type SVR
    iGenDate(0 To 1) As Integer     'Report generation date
    '10-10-01
    lGenTime As Long                'report generation time
    'iGenTime(0 To 1) As Integer     'report generation time
    iVefCode As Integer             'vehicle code
    sZone As String * 3             'time zone (est, pst, etc)
    iType As Integer                '1=pgm,2=comment,3=avail,4=spot, 5=time zone cpy
    iStartofWk(0 To 1) As Integer   'start date of week
    iAirTime(0 To 1) As Integer     'air time
    iDPStartTime(0 To 1) As Integer 'DP start time
    iSeq As Integer                 'sequence # if same times
    iPosition As Integer            'position of spot within break
    sSpotID(0 To 6) As String * 5   'Spot ID for Mon - Sun
    iBreak(0 To 6) As Integer       'break # (mon-sun)
    'iPos(1 To 7) As Integer         'position within break (mon-sun)
    iLen(0 To 6) As Integer         'length (mon - sun)
    'iAirDate(0 To 1, 1 To 7) As Integer     'air date  (mon - sun)
    iAdfCode(0 To 6) As Integer     'advt code (mon-sun)
    sProduct(0 To 6) As String * 30 'product description (mon-sun),chg 11-5-99 from 20 to 30
    sProgramInfo As String * 30     'added 11-5-99 for L37
    'ienfcode(1 To 7) As Integer     'event name code
    'ianfCode(1 To 7) As Integer     'avail name code
    'icifCode(1 To 7) As Integer     'copy pointer
    lHd1CefCode As Long       '9-12-00 Header comment 1 from vehicle options
    lFt1CefCode As Long       '9-12-00 Footer comment 1 from vehicle options
    lFt2CefCode As Long       '9-12-00 Footer comment 2 from vehicle options
    'lRefCode(1 To 7) As Long            ' General Reference Code
    lRefCode(0 To 6) As Long            ' General Reference Code
    sUnused As String * 20
End Type
Type SOFLIST
    iSofCode As Integer         'Selling Office code
    iMnfSSCode As Integer          'associated sales source
End Type
Type MNFLIST
    iMnfCode As Integer         'mnf internal code
    iBillMissMG As Integer      'missed reason billing rules
End Type
Type ACTLIST                    'Sales Activity report
    iAdfCode As Integer         'advertiser code
    iSlfCode As Integer         'slsp code
    iPotnCode As Integer        'potential mnf code (A,B,C)
    lCxfChgR As Long            'change reason code
    iWeekFlag As Integer        '0= previous week data only, 1 = current week  only, 2=both
    lAmount As Long             '$ difference or new amount
End Type
Type ADJUSTLIST                 'Billed & Booked - MGs where they air list
                                'Projection scenarios
    iVefCode As Integer         'For Billed & Booked:vehicle code where $ are moved
                                'For Scenarios - mnf Code for Potential A , B, or C
    'lProject(1 To 13) As Long   'For B & B: mg $
    'lProject(0 To 13) As Long   'For B & B: mg $. Index zero ignored
    lProject(0 To 24) As Long   'For B & B: mg $. Index zero ignored.
                                'For Scenarios: inx 1 = most likely %, inx 2 = optm %, inx 3 = pesm %
                                'Also used in RAB Export for AdServer (RAB can go to 24 months) - TTP 10663 - RAB broadcast export: subscript error on 16 month export.
    iSlsCommPct As Integer         'for NTR to indicate if commissionable on each NTR line.  If 0, no comm; otherwise % xx.xx
    sAgyComm As String * 1          'Y/N for NTR for agy commissionable flag
    iSortCode As Integer        '4-19-05 sort field for generalized sort field comparisons
    'lProjectTrade(1 To 13) As Long  '11-18-05 required for Sales Activity to maintain cash and trade $ by the requested gross or net option
    lProjectTrade(0 To 13) As Long  '11-18-05 required for Sales Activity to maintain cash and trade $ by the requested gross or net option. Index zero ignored
    'lAcquisitionCost(1 To 13) As Long   '6-9-06 for net-net reports to subtract out
    lAcquisitionCost(0 To 13) As Long   '6-9-06 for net-net reports to subtract out. Index zero ignored
    iMnfItem As Integer             '8-10-06  need to separate hard cost NTR on B&B
    iIsItHardCost As Integer        'true if hard cost item
    iSlfCode As Integer             '4-5-11 Sales Activity split slsp
    iNTRInd As Integer              '09/29/2020 TTP # 9952 - If splitting NTR, NTR will be it's own record.  In this event NTR is indicated as 1
    lOrderedCPMCost As Long         '1-21-21
    lBilledCPMCost As Long          '1-21-21
    lPodCode As Long                'ad server internal ID (RE: PATCH-Test(B2402):v8.1 Test Traffic & Affiliate, 6/21/23) - TTP 10761
    iPodCPMID As Integer           'Ad Server Line #
End Type
Type ANR                        '4-24-01 increase all arrays from 13 to 18, add 20 byte unused
    iGenDate(0 To 1) As Integer 'Generation Date
    '10-10-01
    lGenTime As Long            'generation time
    'iGenTime(0 To 1) As Integer 'Generation Time
    iVefCode As Integer         'vehicle code
    iRdfCode As Integer         'DP code
    iYear As Integer            'year processing
    iEffectiveDate(0 To 1) As Integer
    iMnfBudget As Integer       'budget name code
'    lBudget(1 To 18) As Long    'budget $
'    lSold(1 To 18) As Long      'sold $
'    lPriceNeeded(1 To 18) As Long   'price needed to make plan
'    lRCPrice(1 To 18) As Long       'rc price
'    lInv(1 To 18) As Long       'inventory
'    lMinPrice(1 To 18) As Long  'Minimum spot rate
'    lMaxPrice(1 To 18) As Long  'max spot rate
'    lUpfPrice(1 To 18) As Long  'upfront spot rate
'    lScatPrice(1 To 18) As Long 'scatter spot rate
'    iPctSellout(1 To 18) As Integer '% sold , xxx
'    lValuation(1 To 18) As Long     '?
    lBudget(0 To 17) As Long    'budget $
    lSold(0 To 17) As Long      'sold $
    lPriceNeeded(0 To 17) As Long   'price needed to make plan
    lRCPrice(0 To 17) As Long       'rc price
    lInv(0 To 17) As Long       'inventory
    lMinPrice(0 To 17) As Long  'Minimum spot rate
    lMaxPrice(0 To 17) As Long  'max spot rate
    lUpfPrice(0 To 17) As Long  'upfront spot rate
    lScatPrice(0 To 17) As Long 'scatter spot rate
    iPctSellout(0 To 17) As Integer '% sold , xxx
    lValuation(0 To 17) As Long     '?
    iSlfCode As Integer         'slsp code
    iExtra1 As Integer          'extra 2 bytes
    lExtra2 As Long             'extra 4 bytes
    sUnused As String * 20
End Type
'AnrKey record layout
Type BOOKGEN
    lPop As Long
    'sKey As String * 200
    iRdfCode As Integer
    iMnfDemo As Integer
    'sVehName As String * 30
    'sDPName As String * 100
    iVefCode As Integer
    lAvgPrice As Long
    iAvgRating As Integer
    lAvgAud As Long
    lCPP As Long
    lCPM As Long
    iStatus As Integer
End Type
Type SLSLIST
    iVefCode As Integer         'vehicle code
    lPlan As Long               'Plan $
    lTYAct As Long              'actuals (Orders on Books OOB)
    lTYActHold As Long          'For Sales Ana: Actual w/Holds
    lProj As Long               'For Sales Ana: Projection
    lLYWeek As Long             'Sales Ana:  Last Years Week actuals
    lLYAct As Long              'Sales Ana:  Last Years actual
    lPotCodeA As Long           'Act/Proj: Potential code "A" $
    lPotCodeB As Long           'Act/Proj: Potential code "B" $
    lPotCodeC As Long           'Act/Proj: Potential code "C" $
    lMostLike As Long           'Act/Proj: Adjusted most likely $
    iAdfCode As Integer         'Act/Proj: Advertiser code
    iSlfCode As Integer         'Act/Proj: Salesperson code
End Type
Type CPPCPMLIST                 'cpp or cpp by demo/vehicle/advt/product
    sKey As String * 58         'key string:  demo code(5), vef code(5), adf code(5), sofcode (5), length(3), product(35)
                                'each field separated by "|"
    iDnfCode As Integer         'book name (if diff books for same demo, 0 the population
    'lPop(1 To 5) As Long        '6-1-04 changd to array for 4 qtrs and total year
    lPop(0 To 5) As Long        '6-1-04 changd to array for 4 qtrs and total year. Index zero ignored
    iSpots As Integer
    lGross As Long
    lAvgRate As Long
     'the following fields are for 4 quarters and a year total
    'lCPP(1 To 5) As Long
    'lCPM(1 To 5) As Long
    'lCost(1 To 5) As Long
    'iRtg(1 To 5) As Integer
    'lGrImp(1 To 5) As Long
    'lGRP(1 To 5) As Long
    'lAvgAud(1 To 5) As Long
    'Index zero ignored below
    lCPP(0 To 5) As Long
    lCPM(0 To 5) As Long
    lCost(0 To 5) As Long
    iRtg(0 To 5) As Integer
    lGrImp(0 To 5) As Long
    lGRP(0 To 5) As Long
    lAvgAud(0 To 5) As Long
End Type
Type CNTTYPES                   'true or false to include the following contract or spot types
    iHold As Integer
    iOrder As Integer
    iMissed As Integer
    iXtra As Integer
    iTrade As Integer
    iCash As Integer
    iNC As Integer
    iReserv As Integer
    iRemnant As Integer
    iStandard As Integer
    iDR As Integer
    iPI As Integer
    iPSA As Integer
    iPromo As Integer
    iRated As Integer           '10-9-00 also used as Include pgms for spot report
    iNonRAted As Integer        '10-9-00 also used as include comments for spot report
    iSuburban As Integer        '10-9-00 also used as include only open avails in spot report
    sAvailType As String * 1    'S=sellout, A=avails, P = %sellout, I=Inventory
    iOrphan As Integer          'true if using orphan category
    iDayOption As Integer       '0 = Daypart , 1 = Days in Dayparts, 2 = Daypart in Days
    iBuildKey As Integer        'true if Quarterly Booked - build addl tables by line
    iDetail As Integer          'true if detail (separate totals for reserves, holds, sold, etc); false if summary (bottom line availability)
    iShowReservLine As Integer  'true if separate reserved from sold
    iNetwork As Integer         '9-26-00  For stations, include Network spots
    iCharge As Integer          '9-26-00  Include charge spots (non zero)
    iZero As Integer            '9-26-99  Include 0.00 spots
    iADU As Integer             '9-26-00 Include ADU spots
    iBonus As Integer           '9-26-00 Include Bonus spots
    iFill As Integer            '9-26-00 Include fill spots
    iRecapturable As Integer    '9-26-00 Include recapturable spots
    iSpinoff As Integer         '9-26-00 Include spinoff spots
    iMG As Integer              '5-8-08 Incluce MG spot rate
    iFixedTime As Integer       '9-26-00 Include fixed time buys
    iSponsor As Integer         '9-26-00 Include named avail buys
    iDP As Integer              '9-26-00 Include Daypart buys
    iROS As Integer             '9-26-00 Include ROS buys
    iLenHL(0 To 9) As Integer   '9-26-00 4 spot lengths to highlight; 1-09-08 expand for future use to max spot lengths allowed in vehicle
    iValidDays(0 To 6) As Integer '9-26-00 Integer 0 to 6 representing valid days to use
    lRate As Long               '10-9-00 spot rate
    iWorking As Integer         '4-18-02 working props
    iComplete As Integer        '4-18-02 complete props
    iIncomplete As Integer      '4-18-02 incomplete props
    iCntrSpots As Integer           '9-17-04 include contract spots (vs network or feed spots)
    iFeedSpots As Integer           '9-17-04 include network feed spots (vs contract spots)
    iFirm As Integer                '1-16-08 include game type firm
    iTentative As Integer           '1-16-08 include tentative game type
    iPostpone As Integer            '1-16-08 include postponed game type
    iCancelled As Integer           '1-16-08 include cancelled game type
    iNTR As Integer             '3-18-11 include NTR
    iAirTime As Integer         '3-18-11 include Air time (vs NTR)
    iHardCost As Integer        '3-18-11 include hard cost
    iPolit As Integer           '3-18-11 politicals
    iNonPolit As Integer        '3-18-11 non-politicals
    iRep As Integer             '3-18-11 REP
End Type
Type TRANTYPES
    iInv As Integer             'include "I" trans types
    iAdj As Integer             'include "A" trans types
    iWriteOff As Integer        'include "W" trans types
    iPymt As Integer            'include "P" trans types
    iCash As Integer            'include Cash
    iTrade As Integer           'include Trade
    iMerch As Integer           'include merchandising
    iPromo As Integer           'include Promotions
    iNTR As Integer             'include NTR
    iAirTime As Integer         'include Air time (vs NTR)
    iHardCost As Integer        '4-12-07 include hard cost
End Type

'Type SBFTypes                   '9-30-02
'    iNTR As Integer             'include "I" SBFTypes
'    iInstallment As Integer     'include "F" SBFTypes
'    iImport As Integer          'include "T" SBFTypes
'End Type

Type splitinfo                  '9-30-02
    iMatchSSCode As Integer     'Sales source
    iStartCorT As Integer       'Include cash = 1
    iEndCorT As Integer        'include trade = 2
    iFirstProjInx As Integer   'first month index for projection data
    iUseSlsComm As Integer      'use sls commission (true/false)
    iVefCode As Integer         'vehicle code to process
    iDateFlag As Integer        '0 = no comparison, i.e. Billed & Booked (base dates vs comparison dates)
'                                Sales Comparison - 1 = base date
'                                Sales Comparioson - 2 = comparison date
    iNewCode As Integer         'mnf code designated as NEW business
    sPctTrade As String         'Trade pct
    sCashAgyComm As String      'Agency comm pct
    sTradeAgyComm As String     '12-23-02 NTR could have different agy flag than air time
    iNTRSlspComm As Integer     '2-2-04 NTR slsp comm.  If not using sub-companies for commissions, use the slsp comm
                                'in NTR; otherwise use slsp comm from SBF
    sNTR As String * 1          'flag to indiate NTR (Y/N)
    iHardCost As Integer        'hard cost (true/false)
    iMnfNTRItemCode As Integer  '11-06-06 NTR Item type code
End Type
'********************************************************
'
'Text Report file definition
'
'*********************************************************
Type TXR
    iGenDate(0 To 1) As Integer 'Generation Date
    'l10-10-01
    lGenTime As Long            'generation time
    'iGenTime(0 To 1) As Integer 'Generation Time
    lSeqNo As Long              'Sequence number to retain order of records in Cyrstal
    iType As Integer            'Record Type:
                                '  For Export Station Feed (Copy Export)
                                '    1= Rotation header line
                                '    2= Copy line
                                '    3= Comment and instruction line
    sText As String * 200       'Text to be displayed
    lCsfCode As Long            'Comment reference
    iGeneric1 As Integer        'Generic value
    lGenericLong As Long        'Generic long value
    sUnused As String * 10      'unused
End Type
'Txr key record layout
Type TXRKEY0
    iGenDate(0 To 1) As Integer 'Generation Date
    lGenTime As Long        '10-10-01 generation time
    'iGenTime(0 To 1) As Integer 'Generation Time
End Type

Type NetInvByWeek
    iVefCode As Integer             'vehicle code
    iYear As Integer                'year (i.e. 2005)
    sInvWkYear As String * 1        'Inv by wk (W) or Yr (Y)
    sAllowRollover As String * 1    'allow unused inventory to rollover (Y/N)
    lTotalYear As Long              'all 53 weeks added
    'lInvCount(1 To 53) As Long      'time in minutes/seconds
    lInvCount(0 To 53) As Long      'time in minutes/seconds. Index zero ignored
End Type

Type PRICETYPES
    iCharge As Integer          'include spot cost Charge lines
    iZero As Integer            'include .00 lines
    iADU As Integer             'include ADU lines
    iBonus As Integer           'include bonus lines
    iNC As Integer              'include NC lines
    iRecap As Integer           'include recapturable lines
    iSpinoff As Integer         'include spinoff lines
    iMG As Integer              'include MG line rates
End Type

'******************************************************************************
' UOR_User_Option_Rpt Record Definition
'
'******************************************************************************
Type UOR
    iGenDate(0 To 1)      As Integer         ' Generation Date
    lGenTime              As Long            ' Generation Time
    tUor As URF
'    iUrfCode              As Integer         ' URF reference code.  All field
'                                             ' except the genDate and GenTime
'                                             ' must match the URF definitions
'    sName                 As String * 20
'    iVefCode              As Integer
'    sRept                 As String * 20
'    sPassword             As String * 20
'    iSlfCode              As Integer
'    sClnMoYr              As String * 1
'    sClnType              As String * 1
'    sClnLayout            As String * 1
'    iClnLeft              As Integer
'    iClnTop               As Integer
'    iClcLeft              As Integer
'    iClcTop               As Integer
'    sWin(0 To 70)         As String * 1
'    sGrid                 As String * 1
'    sPrice                As String * 1
'    sCredit               As String * 1
'    sPayRate              As String * 1
'    sMerge                As String * 1
'    sDelete               As String * 1
'    sHideSpots            As String * 1
'    sChgBilled            As String * 1
'    sChgCntr              As String * 1
'    sChgCrRt              As String * 1
'    sBouChk               As String * 1
'    sReprintLogAlert      As String * 1
'    sIncompAlert          As String * 1
'    sCompAlert            As String * 1
'    sSchAlert             As String * 1
'    sHoldAlert            As String * 1
'    sRateCardAlert        As String * 1
'    sResearchAlert        As String * 1
'    sAvailAlert           As String * 1
'    sCrdChkAlert          As String * 1
'    sDeniedAlert          As String * 1
'    sCrdLimitAlert        As String * 1
'    sMoveAlert            As String * 1
'    iMnfHubCode           As Integer
'    sChgLnBillPrice       As String * 1
'    sShowNRMsg            As String * 1
'    sUnused2              As String * 12
'    sWorkToDead           As String * 1
'    sWorkToComp           As String * 1
'    sWorkToHold           As String * 1
'    sWorkToOrder          As String * 1
'    sCompToIncomp         As String * 1
'    sCompToDead           As String * 1
'    sCompToHold           As String * 1
'    sCompToOrder          As String * 1
'    sIncompToDead         As String * 1
'    sIncompToComp         As String * 1
'    sIncompToHold         As String * 1
'    sIncompToOrder        As String * 1
'    sDeadToWork           As String * 1
'    sHoldToOrder          As String * 1
'    iPasswordDate(0 To 1) As Integer
'    sChangeCSIDate        As String * 1
'    sAllowInvDisplay      As String * 1
'    sUnused3              As String * 6
'    sResvType             As String * 1
'    sRemType              As String * 1
'    sDRType               As String * 1
'    sPiType               As String * 1
'    sPSAType              As String * 1
'    sPromoType            As String * 1
'    sRefResvType          As String * 1
'    iSnfCode              As Integer
'    sUseComputeCMC        As String * 1
'    iGroupNo              As Integer
'    sReviseCntr           As String * 1
'    sBlockRU              As String * 1
'    sRCView               As String * 1
'    sRegionCopy           As String * 1
'    sPDFDrvChar           As String * 1
'    iPDFDnArrowCnt        As Integer
'    sPrtDrvChar           As String * 1
'    iPrtDnArrowCnt        As Integer
'    sPrtNameAltKey        As String * 1
'    iPrtNoEnterKey        As Integer
'    sChgPrices            As String * 1
'    sSpotFont             As String * 1
'    sActFlightButton      As String * 1
'    sAvailSettings        As String * 1
'    iRemoteUserID         As Integer
'    iRemoteID             As Integer
'    iAutoCode             As Integer
'    iSyncDate(0 To 1)     As Integer
'    isyncTime(0 To 1)     As Integer
'    sOldPassword1         As String * 20
'    sOldPassword2         As String * 20
'    sOldPassword3         As String * 20
'    sSportPropOnly        As String * 1
'    lEMailCefCode         As Long
'    sPhoneNo              As String * 25
'    sCity                 As String * 50
'    sAllowedToBlock       As String * 1
End Type


Type UORKEY0
    iGenDate(0 To 1)      As Integer
    lGenTime              As Long
End Type

Public Type RQF
    lCode                 As Long
    sPriority             As String * 1      ' L=Low, N=Normal, H=High
    iPrintCopies          As Integer         ' # of copies to print
    sReportName           As String * 20
    sRunType              As String * 1      ' N=Now,D=Daily,W=Weekly,F=First
                                             ' day after month end
    sReportSource         As String * 1      ' N=No pre-pass P=pre-pass
    sReportType           As String * 1      ' T=Traffic A=Affiliate
    sOutputType           As String * 1      ' D=Display P=Print S=Save to file
    iOutputSaveType       As Integer         ' 0=pdf 1=excel 2=word 3=text 4=csv
                                             ' 5=rtf
    sOutputFileName       As String * 200    ' can include path
    sRunMode              As String * 1      ' C=Client S=Server
    sRunTime              As String * 11
    iRunDay               As Integer         ' 1 = monday -> 7= Sunday
    sLastDateRun          As String * 10
    lPrePassDate          As Long
    lPrePassTime          As Long
    lEnteredDate          As Long
    lEnteredTime          As Long
    sUserName             As String * 20     'User Name
    sDisposition          As String * 1      ' E=erase when done, R=retain when done
    sCompleted            As String * 1      ' N, Y, P (Processing) and E(Report completd but had error)
    lConnection           As Long            '0 is api call, 1 is odbc
    'sUnused               As String * 20
    lRqfCode              As Long            ' Used to obtain the Multi-Report
                                             ' RQF records.  Master (or Parent)
                                             ' RQF Code stored into each
                                             ' Multi-Report.
    iMultiReportSeqNo     As Integer         ' Sequence number of the
                                             ' multi-reports.
    sPCMACAddr            As String * 20     ' MAC Address
    sUnused               As String * 14
End Type


Public Type RQFKEY0 'VBC NR
    lCode                 As Long 'VBC NR
End Type 'VBC NR
' Dan M 10/21/09 these won't be used in traffic
Public Type RQFKEY1 'VBC NR
    sRunMode              As String * 1 'VBC NR
    sCompleted            As String * 1
    sPriority             As String * 1 'VBC NR
    lEnteredDate          As Long 'VBC NR
    lEnteredTime          As Long 'VBC NR
End Type 'VBC NR
Public Type RQFKEY2 'VBC NR
    lRqfCode              As Long
    iMultiReportSeqNo     As Integer
End Type 'VBC NR
Public Type RQFKEY3
    sUserName             As String * 20
    lEnteredDate          As Long
    lEnteredTime          As Long
End Type

Public Type RFF
    lCode                 As Long
    lRqfCode              As Long
    iSequenceNumber       As Integer
    sFormulaName          As String * 40
    sType                 As String * 1      ' F=formula field, R= record selection, A= ado recordset filename, M= Multi-report, P= Pre-pass report selection crtieria
    sFormulaValue         As String * 255
    lRffCode              As Long            ' 0 unless is a child of extended value; parent code acts as foreign key
    lExtendExists         As Long            ' 0 = no 1=yes.  Only rff to be split (parent) gets a 1
    sUnused               As String * 20
End Type


Public Type RFFKEY0 'VBC NR
    lCode                 As Long 'VBC NR
End Type 'VBC NR

Public Type RFFKEY1 'VBC NR
    lRqfCode              As Long 'VBC NR
    sType                 As String * 1 'VBC NR
    iSequenceNumber       As Integer 'VBC NR
End Type 'VBC NR
Public tgrff() As RFF
Public tgRffExtended() As RFF


Type AVAILCOUNT                 'handles reqd to access avails & spots (Gather avails for avails reports)
     hVef As Integer
     hVpf As Integer
     hVsf As Integer
     hChf As Integer
     hClf As Integer
     hCff As Integer
     hSdf As Integer
     hSsf As Integer
     hSmf As Integer
     hLcf As Integer
     hFsf As Integer
     hAnf As Integer
     iVefCode As Integer                'vehicle to process
     iVpfIndex As Integer               'index to vehicle options
     lSDate As Long                     'start date to gather
     lEDate As Long                     'end date to gather
     iFirstBkt As Integer               '1st bucket to start gathering in qtr
     iSpare1 As Integer
     iSpare2 As Integer
     lSpare1 As Long
     lSpare2 As Long
End Type
