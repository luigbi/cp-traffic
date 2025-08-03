Attribute VB_Name = "RptCrParPay"
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: RptCrParPay.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the generation of Participant Payables report
Option Explicit
Option Compare Text


Dim tlChfAdvtExt() As CHFADVTEXT
Dim tlMMnf() As MNF                    'array of MNF records for specific type
Dim tmPifKey() As PIFKEY          'array of vehicle codes and start/end indices pointing to the participant percentages
                                        'i.e Vehicle XYZ has 2 sales sources, each with 3 participants.  That will be a total of
                                        '6 entries.  Vehicle XYZ points to lo index equal to 1, and a hi index equal to 6; the
                                        'next vehicle will be a lo index of 7, etc.
Dim tmPifPct() As PIFPCT          'all vehicles and all percentages from PIF

Dim tmCntAllYear() As ALLPIFPCTYEAR      'all participant % for all vehicles for a contract for 12 months (1 or more vehicles each could have
                                         '1 or more participants
Dim tmOneVehAllYear() As ALLPIFPCTYEAR       'ss mnf code, mnfgroup, 1 year percentages for 1 vehicle (1 or more participants)
Dim tmOnePartAllYear As ONEPARTYEAR      'ss mnf code, mnfgroup, 1 years percentages for 1 participant


Dim lmSingleCntr As Long                'selective contract # to process
Dim lmSingleUserCntr As Long            'selective user entered contract
Dim smGrossOrNet As String              'G = gross, N = net
Dim imVehicle As Integer                'true if vehicle option
Dim imVehGroup As Integer
Dim imMajorSet As Integer               'major vehicle group selection
Dim imMinorSet As Integer               'minor vehicle group selection
Dim lmProject(0 To 13) As Long          'projection $. Index zero ignored
Dim lmAcquisition(0 To 13) As Long      'acquisition $ to be adjusted (subtracted) from contracts. Index zero ignored
Dim lmAcquisitionNet(0 To 13) As Long   'Index zero ignored
Dim smAirOrder As String * 1         'Site Pref: inv all contracts as aired or ordered (A or O)
Dim tmAdjust() As ADJUSTLIST         'list of MGs $ and vehicles moved to
Dim tmSBFAdjust() As ADJUSTLIST
Dim imUpperAdjust As Integer         'running count of count of MGs built per contract

Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey As SDFKEY0            'SDF record image (key 3)
Dim tmSdfSrchKey1 As SDFKEY1            'SDF record image (key 1)
Dim tmSdfSrchKey2 As SDFKEY2            'SDF record image (key 2)
Dim tmSdfSrchKey4 As SDFKEY4
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdf As SDF
Dim tmSpotTypes As SPOTTYPES     'spot types to include
Dim tmSdfSrchKey3 As LONGKEY0     'SDF record image (SDF code as keyfield)
Dim hmSmf As Integer            'MG and outside Times file handle
Dim tmSmf As SMF                'RPF record image
Dim tmSmfSrchKey As SMFKEY0            'SMF record image
Dim tmSmfSrchKey1 As LONGKEY0           'SMF key1
Dim tmSmfSrchKey4 As SMFKEY4
Dim imSmfRecLen As Integer        'RPF record length
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim tmChfSrchKey1 As CHFKEY1            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim tmClfSrchKey As CLFKEY0     'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim lmActPrice As Long          'actual spot price from cff, or the acquisition cost if Insrtion Order

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
Dim tmSofSrchKey As INTKEY0     'SoF key image
Dim tmSof As SOF
Dim hmSlf As Integer            'Salesperson file handle
Dim imSlfRecLen As Integer      'SLF record length
Dim tmSlfSrchKey As INTKEY0     'SLF key image
Dim tmSlf As SLF

Dim hmUrf As Integer            'User file handle
Dim imUrfRecLen As Integer      'URF record length
Dim tmUrf As URF
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnfSrchKey As INTKEY0
Dim tmMnf As MNF
Dim tmMnfList() As MNFLIST        'array of mnf codes for Missed reasons and billing rules

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim tmVefSrchKey As INTKEY0            'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer            'Virtual Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length

Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length

Dim imStandard As Integer
Dim imRemnant As Integer    'True=Include Remnant
Dim imReserv As Integer  'true = include reservations
Dim imDR As Integer     'True =Include Direct Response
Dim imPI As Integer     'True=Include per Inquiry
Dim imPSA As Integer    'True=Include PSA
Dim imPromo As Integer  'True=Include Promo
Dim imHold As Integer   'true = include hold contracts
Dim imTrade As Integer  'true = include trade contracts
Dim imAirTime As Integer '11-25-02 true = include air time (vs NTR)
Dim imNTR As Integer    '11-25-02 true = include NTR
Dim imHardCost As Integer   '3-21-05 includee hard cost
Dim imOrder As Integer  'true = include Complete order contracts
Dim imFeed As Integer   '7-20-04 Include network (vs local)
Dim imAdjustAcquisition As Integer      '6-8-06
Dim imInclPolit As Integer             '10-2-06 include politicals for B & B / sales Comparison
Dim imInclNonPolit As Integer           '10-2-06 incl non-politicals for B & B/Sales comparison
Dim imInclAdj As Integer

'  Receivables File
Dim hmRvf As Integer        'receivables file handle
Dim tmRvf As RVF            'RVF record image
Dim imRvfRecLen As Integer  'RVF record length

Dim hmPhf As Integer        'receivables file handle

Dim hmSbf As Integer        'SBF file handle
Dim tmSbf As SBF            'SBF record image
Dim tmSbfSrchKey As LONGKEY0 'SBF key record image
Dim imSbfRecLen As Integer  'SBF record length


'Dim tlSofList() As SOFLIST    'list of selling office codes and sales sourcecodes
Dim tmSofList() As SOFLIST    'list of selling office codes and sales sourcecodes

'4-20-00 Buffers for Billed & Booked by Commissions
Dim lgNetDollars(0 To 18) As Long   'Index zero ignored
Dim lgCommDollars(0 To 18) As Long  'Index zero ignored
Dim lgNetNoSplit(0 To 18) As Long   'index zero ignored
'Dim lg12MonthTotal As Long
Dim dg12MonthTotal As Double            '2-28-08
Dim lgGrsDollars(0 To 18) As Long       '2-26-01. Index zero ignored
'Dim lg12MonthAcquisition As Long         '9-20-06
Dim dg12MonthAcquisition As Double      '2-28-08
Dim lgNetCollectNoSplit(0 To 18) As Long     'net payment collected


Dim tmTranTypes As TRANTYPES            '12-29-06
Dim tmSBFType As SBFTypes
Dim imUseCodes() As Integer             'codes to include or exclude (for agy, advt, vehicle, etc)
Dim imUsevefcodes() As Integer        'array of vehicle codes to include/exclude
Dim imInclVefCodes As Integer               'flag to incl or exclude vehicle codes

'7-16-08 ntr/hard cost
Dim tmMnfNtr() As MNF
Const NOT_SELECTED = 0
Dim imAcqLoInx As Integer
Dim imAcqHiInx As Integer

Dim imTempCount As Integer
Private UniqueInfo_rst As ADODB.Recordset

Type PAYBY_BREAKOUT
    lCntrNo As Long         'contract # for detail
    lCode As Long           'internal contract code
    iMnfPartCode As Integer 'participant name
    iPartSharePct As Integer    'participant %
    iVefCode As Integer         'vehicle code
    iAdfCode As Integer         'advt code
    sProductName As String * 30 'product from contract
    '12 monthly buckets plus a total year
    lGross(0 To 12) As Long     'gross sales from order (no splits)
    lNet(0 To 12) As Long       'net sales (no splits)
    lCollected(0 To 12) As Long 'cash collected (PI, journal entries "W"), exlcude PO
    lOutstanding(0 To 12) As Long   'Net sales minus net collected
    lPartner(0 To 12) As Long       'Partner collected (collected * partner %)- what partner is due based on what collected so far
    lOwed(0 To 12) As Long          'Partner Outstanding (partner pct * net sales) minus Partner collected
    lPartnersWorth(0 To 12) As Long 'amt partner should get in entirety
End Type

Private tmPayByContract() As PAYBY_BREAKOUT     'detail contract totals
Private tmPayByVehicle() As PAYBY_BREAKOUT      'vehicle totals
Private tmPayByPartner() As PAYBY_BREAKOUT      'participant totals
Private tmPayByFinal() As PAYBY_BREAKOUT        'final totals

Const TOTAL_BYCNT = 1
Const TOTAL_BYVEHICLE = 2
Const TOTAL_BYPARTNER = 3
Const TOTAL_BYFINAL = 4
Const CAT_GROSS = 10                'gross sales , no splits (totals useless)
Const CAT_NET = 20                  'net sales, no splits, see Net() in tables
Const CAT_COLLECTED = 30            'net collected (orig sales, no split), see Collected() in tables
Const CAT_OUTSTANDING = 40          'net outstanding (orig sales, no split), see Outstanding() in tables
Const CAT_PARTNER = 50              'participant share, see PartnerWorth() in tables
Const CAT_PARTNER_COLLECT = 51      'participant share collected, see Partner() in tables
Const CAT_OWED = 60                 'participant share owed , see Owed() in tables

Sub mParPaySbfAdjustForInstall(tlSbf() As SBF, llStdStartDates() As Long, ilFirstProjInx As Integer, llStartAdjust As Long, llEndAdjust As Long, ilHowManyPer As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilIsItHardCost                                                                        *
'******************************************************************************************
    Dim slDate As String
    Dim llDate As Long
    Dim ilMonthInx As Integer
    Dim ilFoundMonth As Integer
    Dim ilFoundVef As Integer
    Dim ilTemp As Integer
    'TTP 10855 - prevent overflow due to too many NTR items
    Dim llSBFLoop As Long
    'Dim ilSBFLoop As Integer
    Dim ilFoundOption As Integer

    For llSBFLoop = LBound(tlSbf) To UBound(tlSbf) - 1
        tmSbf = tlSbf(llSBFLoop)
        gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), slDate
        llDate = gDateValue(slDate)
        If llDate > llEndAdjust Then
            Exit Sub
        End If
        'SBF is OK with dates, adjust the $

        'determine if hard cost to be included for Billed & Booked; other reports dont have option or is
        'interspersed with non hard cost
        ilFoundOption = True                'default to include the NTR if no options exist and its not B & B or Recap report
            If tlSbf(llSBFLoop).sTranType <> "F" Then            'installment
                ilFoundOption = False
            End If

        If ilFoundOption Then
            ilFoundMonth = False
            For ilMonthInx = ilFirstProjInx To ilHowManyPer Step 1         'loop thru months to find the match
                If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                    ilFoundMonth = True
                    Exit For
                End If
            Next ilMonthInx

            ilFoundVef = False
            'setup vehicle that spot was moved to
            For ilTemp = LBound(tmSBFAdjust) To UBound(tmSBFAdjust) - 1 Step 1
                If tmSBFAdjust(ilTemp).iVefCode = tlSbf(llSBFLoop).iBillVefCode Then
                    ilFoundVef = True
                    Exit For
                End If
            Next ilTemp
            If Not (ilFoundVef) Then
                ilTemp = UBound(tmSBFAdjust)
                tmSBFAdjust(ilTemp).iVefCode = tlSbf(llSBFLoop).iBillVefCode
                'Installments from SBF need to set the agy Commission flag since its not coming from the SBF record,
                'but from the contract header
                tmSBFAdjust(ilTemp).iSlsCommPct = 0                         '4-25-08
                If tgChfCT.iAgfCode > 0 Then
                    tmSBFAdjust(ilTemp).sAgyComm = "Y"   '4-25-08
                Else
                    tmSBFAdjust(ilTemp).sAgyComm = "N"
                End If
                tmSBFAdjust(ilTemp).iIsItHardCost = False                   '4-25-08
                tmSBFAdjust(ilTemp).iMnfItem = 0                           '4-25-08


'                tmSBFAdjust(ilTemp).iSlsCommPct = tlSbf(ilSBFLoop).iCommPct
'                tmSBFAdjust(ilTemp).sAgyComm = tlSbf(ilSBFLoop).sAgyComm
'                tmSBFAdjust(ilTemp).iIsItHardCost = ilIsItHardCost              '8-10-06
'                tmSBFAdjust(ilTemp).iMnfItem = tlSbf(ilSBFLoop).iMnfItem        '8-10-06

                If ilFoundMonth Then
                    tmSBFAdjust(ilTemp).lProject(ilMonthInx) = tmSBFAdjust(ilTemp).lProject(ilMonthInx) + tlSbf(llSBFLoop).lGross
                    tmSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) = tmSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) + (tlSbf(llSBFLoop).lAcquisitionCost)
                End If
                ReDim Preserve tmSBFAdjust(0 To UBound(tmSBFAdjust) + 1) As ADJUSTLIST
            Else
                If ilFoundMonth Then
                    tmSBFAdjust(ilTemp).lProject(ilMonthInx) = tmSBFAdjust(ilTemp).lProject(ilMonthInx) + tlSbf(llSBFLoop).lGross
                    tmSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) = tmSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) + (tlSbf(llSBFLoop).lAcquisitionCost)
                End If
            End If
        End If
    Next llSBFLoop
    Exit Sub
End Sub

'                   mParPaySbfAdjustForNTR -  determine where item billing $ goes (Billed & Booked)
'                   <input> tlSbf() - array of SBF records to process
'                           llstdstartDates() - array of 13 start dates of the 12 months to gather
'                           ilFirstProjInx - index of first month to start projection (earlier is from receivables)
'                           llStartAdjust -  Earliest date to start searching for missed, etc.
'                           llEndAdjust - latest date to stop searchng for missed, etc.
'                           ilHowManyPer - # of periods to gather
'                                           (for billed & booked its 12,
'                                           for Sales Comparisons its max 3)
'
Sub mParPaySbfAdjustForNTR(tlSbf() As SBF, llStdStartDates() As Long, ilFirstProjInx As Integer, llStartAdjust As Long, llEndAdjust As Long, ilHowManyPer As Integer)
    Dim slDate As String
    Dim llDate As Long
    Dim ilMonthInx As Integer
    Dim ilFoundMonth As Integer
    Dim ilFoundVef As Integer
    Dim ilTemp As Integer
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSBFLoop As Integer
    Dim llSBFLoop As Long
    Dim ilIsItHardCost As Integer
    Dim ilFoundOption As Integer

    For llSBFLoop = LBound(tlSbf) To UBound(tlSbf) - 1
        tmSbf = tlSbf(llSBFLoop)
        gUnpackDate tmSbf.iDate(0), tmSbf.iDate(1), slDate
        llDate = gDateValue(slDate)
        If llDate > llEndAdjust Then
            Exit Sub
        End If
        'SBF is OK with dates, adjust the $

        'determine if hard cost to be included for Billed & Booked; other reports dont have option or is
        'interspersed with non hard cost
        ilFoundOption = True                'default to include the NTR if no options exist and its not B & B or Recap report
        ilIsItHardCost = gIsItHardCost(tlSbf(llSBFLoop).iMnfItem, tlMMnf())

        If ilIsItHardCost Then              'hard cost item
            If Not imHardCost Then
                ilFoundOption = False
            End If
        Else                                'normal NTR
            If Not imNTR Then
                ilFoundOption = False
            End If
        End If

        If ilFoundOption Then
            ilFoundMonth = False
            For ilMonthInx = ilFirstProjInx To ilHowManyPer Step 1         'loop thru months to find the match
                If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                    ilFoundMonth = True
                    Exit For
                End If
            Next ilMonthInx

            ilFoundVef = False
            'setup vehicle that spot was moved to
            For ilTemp = LBound(tmSBFAdjust) To UBound(tmSBFAdjust) - 1 Step 1
                If tmSBFAdjust(ilTemp).iVefCode = tlSbf(llSBFLoop).iBillVefCode And tmSBFAdjust(ilTemp).iSlsCommPct = tlSbf(llSBFLoop).iCommPct And tmSBFAdjust(ilTemp).sAgyComm = tlSbf(llSBFLoop).sAgyComm And tmSBFAdjust(ilTemp).iMnfItem = tlSbf(llSBFLoop).iMnfItem Then      '8-10-06
                    ilFoundVef = True
                    Exit For
                End If
            Next ilTemp
            If Not (ilFoundVef) Then
                ilTemp = UBound(tmSBFAdjust)
                tmSBFAdjust(ilTemp).iVefCode = tlSbf(llSBFLoop).iBillVefCode
                tmSBFAdjust(ilTemp).iSlsCommPct = tlSbf(llSBFLoop).iCommPct
                tmSBFAdjust(ilTemp).sAgyComm = tlSbf(llSBFLoop).sAgyComm
                tmSBFAdjust(ilTemp).iIsItHardCost = ilIsItHardCost              '8-10-06
                tmSBFAdjust(ilTemp).iMnfItem = tlSbf(llSBFLoop).iMnfItem        '8-10-06

                If ilFoundMonth Then
                    tmSBFAdjust(ilTemp).lProject(ilMonthInx) = tmSBFAdjust(ilTemp).lProject(ilMonthInx) + (tlSbf(llSBFLoop).lGross * tlSbf(llSBFLoop).iNoItems)
                    tmSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) = tmSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) + (tlSbf(llSBFLoop).lAcquisitionCost * tlSbf(llSBFLoop).iNoItems)
                End If
                ReDim Preserve tmSBFAdjust(0 To UBound(tmSBFAdjust) + 1) As ADJUSTLIST
            Else
                If ilFoundMonth Then
                    tmSBFAdjust(ilTemp).lProject(ilMonthInx) = tmSBFAdjust(ilTemp).lProject(ilMonthInx) + (tlSbf(llSBFLoop).lGross * tlSbf(llSBFLoop).iNoItems)
                    tmSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) = tmSBFAdjust(ilTemp).lAcquisitionCost(ilMonthInx) + (tlSbf(llSBFLoop).lAcquisitionCost * tlSbf(llSBFLoop).iNoItems)
                End If
            End If
        End If
    Next llSBFLoop
    Exit Sub
End Sub

'           gCloseBOBFilesCt - Close all applicable files for
'                       Billed and booked, Slsp Comm Proj,
'                       and Sales Comparison reports
'
Sub gParPayCloseFiles()
    Dim ilRet As Integer
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmRvf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmSbf)
    ilRet = btrClose(hmVsf)
    btrDestroy hmGrf
    btrDestroy hmRvf
    btrDestroy hmCHF
    btrDestroy hmSlf
    btrDestroy hmVef
    btrDestroy hmMnf
    btrDestroy hmSof
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmAgf
    btrDestroy hmSdf
    btrDestroy hmSmf
    btrDestroy hmSbf
    btrDestroy hmVsf

    Erase tmPifKey, tmPifPct
    Erase tmCntAllYear, tmOneVehAllYear
End Sub

'*****************************************************************************************
'
'                            1st part builds actual data from PHF & RVF.
'                            Gathers all "I" and "A" transactions and places
'                            the $ in the appropriate standard month. If selecting
'                            by owner, the $ are split by the vehicle participation.
'                            If running by slsp, split by commission split %.
'                   <Input>  llStdStartDates - array of max 13 start dates, denoting
'                                              start date of each period to gather
'                            llLastBilled - Date of last invoice period
'                            ilLastbilledInx - Index into llStdStartDates of period last
'                                           invoiced
'                            ilHowManyPer - # of periods to gather
'                                           (currently, always 12 for Participant Payables)
'                            blNewContract - if new, initialize the array that stores the gathered data.
'                                           This routine is entered twice per contract (once for PHF/RVF & once for chf)

'
'   GRF fields:
'   grfgenDate - generation date
'   grfGenTime - generation time
'   grfChfCode - contract code
'   grfDateGenl - Date billed or paid
'   grfVefCode - vehicle code (airing or billing)
'   grfAdfCode = advertiser code
'   grfSlfCode = salesperson code
'   grfSofCode = Sales Source
'   grfCode2 - Unused
'   grfCode4 - Unused
'   grfYear - Unused
'   grfDateType = C = cash, T = trade
'   grfDollars(1-12) 12 months $
'   grfDollars(13) total year $ (gross or net)
'   grsDollars(14-17) Q1 - Q4 $
'   grfDollars(18) - year total $ always net
'   grfPerGenl(1) - unused
'   grfPerGenl(2) - Unused
'   GrfPerGenl(3) - Unused
'   GrfPerGenl(4) - major vehicle group
'   grfPerGenl(5) - Flag for sorting  1 = gross, 2 = net, 3= collected, 4 = outstanding, 5 = partner net , 6 = owed to partner
'   grfPerGenl(7) -Participant % (no longer used since the calculation must be done in prepass, not in .rpt
'   grfPerGenl(8) - Unused
'   grfPerGenl(9)-  Is it hard cost (true/false)
'   grfPerGenl(10) - 4-20-05 subsort field for B & B Recap:  1 = airtime, 2 = NTR , 3 = hard cost NTR
'   grfPerGenl(10) - 11-06-06 changed to:  0 = not NTR or Hard cost, > 0 = MNF item code for NTR
'                   if NTR, flag in NTR indicates if agy sales, direct sales or NTR sales.  If
'                   agy or direct sales, embed them in those categories and not with NTR.
'   grfPerGenl(11) - Unused

'   grfPerGenl(12) - Unused
'   grfPerGenl(13) - Unused
'   grfPerGenl(14) - Unused
'   grfPerGenl(18) - # periods to print
'
'*****************************************************************************************
Sub gCRParPay_Past(llStdStartDates() As Long, llLastBilled As Long, ilLastBilledInx As Integer, ilHowManyPer As Integer, blNewContract As Boolean)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCurrentRecd                 slNameCode                    ilFound                   *
'*  slAmount                      slDollar                      slPct                     *
'*  ilLoopSlsp                    ilSearchSlf                   llAcquisitionCost         *
'*  llTempStart                   llTempEnd                     ilLoopOnVG                *
'*                                                                                        *
'******************************************************************************************

    Dim ilRet As Integer
    Dim illoop  As Integer
    Dim slCode As String
    Dim slStr As String
    Dim llAmt As Long
    Dim ilMonthNo As Integer
    Dim llDate As Long
    Dim ilTemp As Integer
    Dim ilFoundMonth As Integer
    Dim llGrossDollar As Long
    Dim llTransGross As Long                'for participant splits, orig.  gross amt of trans.
    Dim llNetDollar As Long
    Dim llCommDollar As Long                   '3/11/99 for new comm rept  (comm amt for slsp)
    Dim llTransNet As Long                  'for participants splits, orig net amt of trans.  For split participants, as each participant share iscalculated,
                                            'it is subtracted from orig trans amt.  Any remaining pennies is added to last participant
    Dim llOrigTransNet As Long              'Orig net $ from transaction, no splits
    Dim ilFoundOne As Integer
    Dim ilFoundOption As Integer
    Dim ilMatchCntr As Integer              'selectivity on holds, & contr types (remnants, PIs, etc)
    Dim ilHowMany As Integer                'times to loop - up to 10 records created per trans if by slsp with splits,
                                            'or just 1 per transaction if vehicle or advt option
    Dim ilMatchSSCode As Integer            'Sales source to process
    Dim ilHowManyDefined As Integer         '# of percentages (slsp for each contract, or participants for each vehicle)
    Dim llProcessPct As Long                '% of slsp split or vehicle owner split (else 100%)
    Dim slCommPct As String
    Dim ilmaxTypes As Integer               '# of records to create for Billed & booked vs Slsp comm rept
                                            'For Billed &booked:  1 for gross, 2 for net)
                                            '1 if comm rept and the trans is a merch or promotion, 3 if cash or trade transaction for comm rept)
    Dim ilMinTypes As Integer               'set to 1,2,or 3: For Billed &booked:  1 for gross, 2 for net
                                            'for slsp comm:  1 if its a cash or trade trans, 3 if its a merch or promo trans.  Only create
                                            'the comm , dont affect gross or net values
    Dim ilTypes As Integer
    Dim ilReverseFlag As Integer            'reverse sign of amount before updating record
    Dim ilLoopType As Integer
    Dim llLoopType As Long                  '7-29-09
    Dim ilFoundType As Integer
'    Dim tlTranType As TRANTYPES
    Dim ilRelativeIndex As Integer          'relative index for slsp% split or participant split
    ReDim tlRvf(0 To 0) As RVF

    'ReDim tlMnf(1 To 1) As MNF
    ReDim tlMnf(0 To 0) As MNF
    Dim ilLoopRecv As Integer                   '9-16-01

    Dim ilMatchSOFCode As Integer
    Dim ilValidTran As Integer                  'valid trans for test of  NTR or AirTime
    Dim ilTempPeriods As Integer                'process # of periods based on user entered, or last period invoiced index, whichever is less
                                                'for example, last period invoiced index is 6 (June) but user requests 3 periods to process.
    'Dim llTempSTdStartDates(1 To 13) As Long    'monthly std dates for Sales Comparison
    Dim llTempStdStartDates(0 To 13) As Long    'monthly std dates for Sales Comparison. Index zero ignored
    Dim ilOKtoSeeVeh As Integer                 '11-13-03 flag to determine if user allowed to see vehicle
    Dim llRvfLoop As Long                       '2-11-05
    Dim ilIsItHardCost As Integer
    ''Dim ilProdPct(1 To 8) As Integer            '7-06-05
    ''Dim ilMnfGroup(1 To 8) As Integer           '7-06-05
    ''Dim ilMnfSSCode(1 To 8) As Integer
    'ReDim ilProdPct(1 To 1) As Integer            '5-1-07  Participant share
    'ReDim ilMnfGroup(1 To 1) As Integer           '5-1-07  Participants
    'ReDim ilMnfSSCode(1 To 1) As Integer          '5-1-07  Particpant sales source
    ReDim ilProdPct(0 To 1) As Integer            '5-1-07  Participant share
    ReDim ilMnfGroup(0 To 1) As Integer           '5-1-07  Participants
    ReDim ilMnfSSCode(0 To 1) As Integer          '5-1-07  Particpant sales source
    Dim ilSaveLastBilledInx As Integer
    Dim llDateAdjust As Long                        '1-08-08  adjustment for end of last bill date when other than std requested.
                                                    'i.e. last bill date = 11/25/07.  request corp which month of Nov ends 12/2/07.
                                                    'when its detected that the corp end date extends past the last billed date,
                                                    'its assumed that corp nov is in the future and stops.
    Dim slLastBilledDate As String              '1-08-08
    Dim slTemp As String
    Dim ilIncludeVehicleGroup As Integer    '11-21-08
    Dim llContrCode As Long                 '8-19-09
    Dim llTransEnteredDate As Long          '8-19-09
    Dim blAcqOK As Boolean
    Dim slEarliestDate As String
    Dim slLatestDate As String
    Dim blThisIsBilling As Boolean
    Dim ilCategory As Integer
    Dim blBilledInFuture As Boolean         'contr invoiced, but last billing date not set in site; need to ignore these or double numbers
    Dim blPennyVariance As Boolean
    ilSaveLastBilledInx = ilLastBilledInx
    slLastBilledDate = Format$(llLastBilled, "m/d/yy")


'    'build array of selling office codes and their sales sources.
'    'need to get the sales source from contracts slsp in order to find the correct participant entry
'    ilTemp = 0
'    ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'    Do While ilRet = BTRV_ERR_NONE
'        ReDim Preserve tmSofList(0 To ilTemp) As SOFLIST
'        tmSofList(ilTemp).iSofCode = tmSof.iCode
'        tmSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
'        ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        ilTemp = ilTemp + 1
'    Loop
'    ilRet = gObtainSlf(RptSelParPay, hmSlf, tlSlf())
'
'    'build array of vehicles to include or exclude
'    gObtainCodesForMultipleLists 0, tgCSVNameCode(), imInclVefCodes, imUsevefcodes(), RptSelParPay
'    gAddDormantVefToExclList imInclVefCodes, tgMVef(), imUsevefcodes()          '8-4-17 if excluding vehicles, make sure dormant ones exluded since
'                                                                                'they wont be in original list
'    ilUsePkg = False
'    'default all other options that are not asked for the Sales Placement report
'    ilStd = True                                'Assume std reporting (vs corporate)
'    igPeriods = 12
'
'    ilWhichDate = 0                     'assume no pacing, use trandate to retieve IN/ANs
'    tlTranType.iAdj = True              'look only for adjustments in the History & Rec files
'    tlTranType.iInv = False
'    tlTranType.iWriteOff = False
'    tlTranType.iPymt = False
'    tlTranType.iCash = True
'    tlTranType.iTrade = False
'    tlTranType.iMerch = False
'    tlTranType.iPromo = False
'
'    tlTranType.iNTR = False         '9-17-02
'    If imNTR Or imHardCost Then                   '4-25-05 need to gather the NTR for adjustments too
'        tlTranType.iNTR = True
'    End If


    llDateAdjust = 0                                '1-08-08

    ilmaxTypes = 3

    If blNewContract Then
'        ReDim tmPayByContract(0 To 0) As PAYBY_BREAKOUT

        For illoop = 1 To 13
            llTempStdStartDates(illoop) = llStdStartDates(illoop)
        Next illoop

    End If


'    For ilLoopRecv = 1 To ilTempPeriods       'loop for as many months that have been billed.  When RVF and PHF accessed separately, exceeded
                                             '32000 (prior to VB6)
       
       slEarliestDate = Format$(llTempStdStartDates(1), "m/d/yy")           'get the start date of month to process
       'slCode = Format$(llTempSTdStartDates(ilLoopRecv + 1) - 1, "m/d/yy")  'get the end date of the month to process
       slLatestDate = "12/31/2069"             'default to get all dates because collections need to be based on ageing period & year falling with the user requested dates
       

       ReDim tlRvf(0 To 0) As RVF

       ilRet = gObtainPhfRvfbyCntr(RptSelParPay, lmSingleCntr, slEarliestDate, slLatestDate, tmTranTypes, tlRvf())    'get all phfrvf by contract starting from earliest request date (routine uses trandate, default to year 2069 for end date
                                                                                 'because collections need to be by ageing date
       '3-20-18 trouble sending contract # 0 because the routine that gets rvf/phf by contract # assumes 0 means get ALL, not any selective contract
       'if lmsinglecntr was -1, it is processing contract #0.  -1 was used for routine gObtainPhfRvfbyCntr to indicate match on cntr #0, not ALL
       'change lmsinglecntr back to 0 and continue process as normal
       If lmSingleCntr = -1 Then
           lmSingleCntr = 0
       End If
       If ilRet = 0 Then
           Exit Sub
       End If
       For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1 Step 1
           tmRvf = tlRvf(llRvfLoop)
           ilValidTran = mParPayNTRTestRVF(tlMMnf(), ilIsItHardCost)

           'for billing rendered, use the transaction date
           'for Collections, use the ageing period & year
           llAmt = mParPayWhichAmtToLong(tmRvf.sNet)            'use receivables gross/net or acquisition cost
           blBilledInFuture = False
   
           If tmRvf.sTranType = "HI" Or tmRvf.sTranType = "IN" Or tmRvf.sTranType = "AN" Then       'use tran date
               gUnpackDate tmRvf.iTranDate(0), tmRvf.iTranDate(1), slStr
               blThisIsBilling = True
               llDate = gDateValue(slStr)
               If llDate >= llStdStartDates(ilLastBilledInx + 1) Then     'trans invoiced, but last bill date not set yet
                   blBilledInFuture = True
               End If
           ElseIf tmRvf.sTranType = "PI" Or Left$(tmRvf.sTranType, 1) = "W" Then
       
               slCode = Trim$(str$(tmRvf.iAgePeriod) & "/15/" & Trim$(str$(tmRvf.iAgingYear)))
               slStr = gObtainEndStd(slCode)
               llDate = gDateValue(slStr)
               blThisIsBilling = False
           End If
           If ((tmRvf.sTranType <> "PO") And (llDate <= llStdStartDates(13))) And Not blBilledInFuture Then

               'get contract from history or rec file
               tmChfSrchKey1.lCntrNo = tmRvf.lCntrNo
               tmChfSrchKey1.iCntRevNo = 32000
               tmChfSrchKey1.iPropVer = 32000
               ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE) 'get matching contr recd

               '9-19-06 when there is no sch lines and need to process merchandising for t-net, the contract may not be scheduled (schstatus = "N");  need to process those contracts
               Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = tmRvf.lCntrNo) And (tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M" And tmChf.sSchStatus <> "N")
                    ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
               Loop

               '4-20-00 mFakeChf                                'if contract headr not found, setup fake header
               gFakeChf tmRvf, tmChf

               ilMatchCntr = mParPayTestTypes()           'test contract types against user requested types to include

               '4-20-00   Contract # selectivity
               If lmSingleCntr > 0 And (lmSingleCntr <> tmRvf.lCntrNo) Then
                   ilMatchCntr = False
               End If

               ilOKtoSeeVeh = gCntrOkForUser(hmVsf, tgUrf(0).iSlfCode, tmChf.lVefCode, tmChf.iSlfCode())

               If (tmRvf.sCashTrade = "M" Or tmRvf.sCashTrade = "P") Then      'ignore merchandising and promotions
                   ilMatchCntr = False
               End If
               'find the sales source and office;
'               'obtain vehicle group if applicable, see if sales source is the selected
               ilFoundOption = mParPaySSCodeRVFSelect(tmSofList(), ilMatchSSCode, ilMatchSOFCode)

               'error return not tested because a phoney record is created if contr not found
               If ilMatchCntr And ilFoundOption And ilOKtoSeeVeh Then       '11-13-03

                   mParPayFormatGrfFromRVF ilIsItHardCost            'format basic fields of prepass record

                   'determine the month that this transaction falls within
                   ilFoundMonth = False
                   'Site pref tested, trans date (entered date) or ageing month/year tested, depending on site pref.
                   For ilMonthNo = 1 To ilHowManyPer Step 1         'loop thru months to find the match
                       If llDate >= llStdStartDates(ilMonthNo) And llDate < llStdStartDates(ilMonthNo + 1) Then
                           ilFoundMonth = True
                           Exit For
                       End If
                   Next ilMonthNo
                   If ilFoundMonth Then
                       llTransGross = mParPayWhichAmtToLong(tmRvf.sGross)            'use receivables gross/net or acquisition cost
                       llTransNet = mParPayWhichAmtToLong(tmRvf.sNet)            'use receivables gross/net or acquisition cost

                       ilReverseFlag = mParPayReverseSign(llTransNet, llTransGross)
                       llOrigTransNet = llTransNet                                 'retain orig net amt since llTransNet will be altered with each split participant
                       'obtain the sales source for major sort of business booked reports
                       '7-1-14 speed up and use binary search
                       ilRet = gBinarySearchSlf(tmChf.iSlfCode(0))
                       If ilRet = -1 Then          'not found
                           tmSlf.iRemUnderComm = 0
                           tmSlf.iSofCode = 0
                       Else
                           tmSlf = tgMSlf(ilRet)
                       End If

                        'if NTR, get that commission instead
                       If tmRvf.iMnfItem > 0 Then          'this indicates NTR
                           'retrieve the associated NTR record from SBF
                           tmSbfSrchKey.lCode = tmRvf.lSbfCode  '12-16-02
                           ilRet = btrGetEqual(hmSbf, tmSbf, imSbfRecLen, tmSbfSrchKey, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                           If ilRet <> BTRV_ERR_NONE Then          '7-28-03 insure the previous % isnt used
                               tmSbf.iCommPct = 0
                           End If
                       End If

                       'associated sales source
                       For illoop = LBound(tmSofList) To UBound(tmSofList)
                           If tmSofList(illoop).iSofCode = tmSlf.iSofCode Then
                               ilMatchSSCode = tmSofList(illoop).iMnfSSCode          'Sales source
                               Exit For
                           End If
                       Next illoop

                       ' Build ilMnfSSCode(), ilMnfGroup(), ilProdPct(), tmPifKey(), tmPifPct()
                       mParPayGetHowManyPast ilMatchSSCode, ilHowMany, ilHowManyDefined, ilMnfSSCode(), ilMnfGroup(), ilProdPct(), blThisIsBilling       '7-6-05

                       'For Advt & Vehicle, create 1 record per transaction
                       'For Owner , create as many as 3 records per transaction.  (up to 3 owners per vehicle)
                       'For slsp, create as many as 10 records per trans (up to 10 split slsp) per trans.
                       For illoop = 0 To ilHowMany - 1 Step 1        'loop based on report option
                           For ilTemp = 1 To 18 Step 1               'init the years $ buckets
                               tmGrf.lDollars(ilTemp - 1) = 0
                           Next ilTemp
                           
                           If ilMnfSSCode(illoop + 1) = ilMatchSSCode Then
                               'llProcessPct = 1000000     '9-27-17
                               llProcessPct = ilProdPct(illoop + 1)    '9-27-17
                               llProcessPct = llProcessPct * 100
                               slCommPct = gIntToStrDec(ilProdPct(illoop + 1), 2)
                           Else
                               If ilMnfSSCode(1) = 0 Then      'no sales source was found for this vehicle, default to 100%
                                   llProcessPct = 1000000
                                   slCommPct = ".00"
                               Else                            'done with the splits
                                   llProcessPct = 0
                                   slCommPct = ".00"
                               End If
                           End If
                          
                           If llProcessPct > 0 Then
                             'if slsp or owner, check to see if this split should be included

                               ilRelativeIndex = mParPayPastSetupKeyForReport(ilMnfGroup(), illoop)

                               tmGrf.iSofCode = ilMatchSSCode          'sales source
                               mParPayConvertAndSplitPast llProcessPct, slCommPct, llTransNet, llNetDollar, llCommDollar, ilRelativeIndex, ilHowManyDefined, ilReverseFlag
                               ilMinTypes = 1          'detrmine the # of records to create if billed & booked vs commissions
                               
                               ilIncludeVehicleGroup = True
                               '8-6-10 option swapped vehicle w/split slsp subtotals with split slsp sort w/vehicle subtotals
                               'need to get the vehicle group selection if applicable
                               gGetVehGrpSets tmGrf.iVefCode, imMinorSet, imMajorSet, tmGrf.iPerGenl(2), tmGrf.iPerGenl(3)   'Genl(3) = minor sort code, genl(4) = major sort code
                               
                               'tmGrf.iPerGenl(4) = tmVef.iMnfGroup(ilLoop + 1)
                               If ilMnfGroup(illoop + 1) <> 0 Then                 '4-5-16 handle case were sales source not defined with vehicle; include so things balance
                                   'tmGrf.iPerGenl(4) = ilMnfGroup(ilLoop + 1)          '5-2-07
                                   tmGrf.iPerGenl(3) = ilMnfGroup(illoop + 1)        '5-2-07
                               Else
                                   'tmGrf.iPerGenl(4) = 0                           '4-14-16 no participant exist for sales source
                                   tmGrf.iPerGenl(3) = 0                           '4-14-16 no participant exist for sales source
                               End If
                                  
                               If tmGrf.iPerGenl(3) > 0 Then
                                   ilIncludeVehicleGroup = mParPayTestVGItem()
                               End If
                               blPennyVariance = False
                               If llNetDollar = 0 And llTransNet > 0 Then      'if the net is $0, and theres pennies left; it cannot split, so the original Net Collected needs to be accounted for
                                   llNetDollar = llTransNet
                                   blPennyVariance = True
                               End If
                               If (llNetDollar <> 0) And (ilIncludeVehicleGroup) Then
                                      If blThisIsBilling Then     'accum gross, net
                                       'tmGrf.lDollars(ilMonthNo - 1) = llGrossDollar
                                       tmGrf.lDollars(ilMonthNo - 1) = llTransGross        '10-6-17
                                       If ilReverseFlag Then         'reverse the sign (math always with positive amts)
                                           tmGrf.lDollars(ilMonthNo - 1) = -tmGrf.lDollars(ilMonthNo - 1)
                                       End If
                                       mParPayAccumPayByContract CAT_GROSS, tmGrf.iVefCode, tmGrf.iPerGenl(3), ilProdPct(illoop + 1)
                                       For ilTemp = 0 To 12
                                           tmGrf.lDollars(ilTemp) = 0
                                       Next ilTemp
                                       
                                       'Total Net (no splits)
                                       tmGrf.lDollars(ilMonthNo - 1) = llOrigTransNet
                                       If ilReverseFlag Then              'reverse the sign (math always with positive amts)
                                           tmGrf.lDollars(ilMonthNo - 1) = -tmGrf.lDollars(ilMonthNo - 1)
                                       End If
                                       mParPayAccumPayByContract CAT_NET, tmGrf.iVefCode, tmGrf.iPerGenl(3), ilProdPct(illoop + 1)
                                       For ilTemp = 0 To 12
                                           tmGrf.lDollars(ilTemp) = 0
                                       Next ilTemp
                                       
                                       'Net Partner (participant) split
                                       tmGrf.lDollars(ilMonthNo - 1) = llNetDollar
                                       If ilReverseFlag Then              'reverse the sign (math always with positive amts)
                                           tmGrf.lDollars(ilMonthNo - 1) = -tmGrf.lDollars(ilMonthNo - 1)
                                       End If
                                       'mParPayAccumPayByContract CAT_NET, tmGrf.iVefCode, tmGrf.iPerGenl(3), ilProdPct(ilLoop + 1)
                                       mParPayAccumPayByContract CAT_PARTNER, tmGrf.iVefCode, tmGrf.iPerGenl(3), ilProdPct(illoop + 1)
                                       For ilTemp = 0 To 12
                                           tmGrf.lDollars(ilTemp) = 0
                                       Next ilTemp
                                       
                                       
                                   Else                        'accum collected
                                       'Total payment collected (no splits)
                                       tmGrf.lDollars(ilMonthNo - 1) = llOrigTransNet
                                       If ilReverseFlag Then              'reverse the sign (math always with positive amts)
                                           tmGrf.lDollars(ilMonthNo - 1) = -tmGrf.lDollars(ilMonthNo - 1)
                                       End If
                                       mParPayAccumPayByContract CAT_COLLECTED, tmGrf.iVefCode, tmGrf.iPerGenl(3), ilProdPct(illoop + 1)
                                       For ilTemp = 0 To 12
                                           tmGrf.lDollars(ilTemp) = 0
                                       Next ilTemp

                                       If Not (blPennyVariance) Then
                                           tmGrf.lDollars(ilMonthNo - 1) = llNetDollar
                                           If ilReverseFlag Then              'reverse the sign (math always with positive amts)
                                               tmGrf.lDollars(ilMonthNo - 1) = -tmGrf.lDollars(ilMonthNo - 1)
                                           End If
                                           mParPayAccumPayByContract CAT_PARTNER_COLLECT, tmGrf.iVefCode, tmGrf.iPerGenl(3), ilProdPct(illoop + 1)
                                           For ilTemp = 0 To 12
                                               tmGrf.lDollars(ilTemp) = 0
                                           Next ilTemp
                                       End If
'
'                                        slTemp = gLongToStrDec(llNetDollar, 2)
'                                        slStr = gMulStr(slCommPct, slTemp)
'                                        llCommDue = Val(gRoundStr(slStr, "01.", 0))
'                                        tmGrf.lDollars(ilMonthNo - 1) = llCommDue
'                                        If ilReverseFlag Then              'reverse the sign (math always with positive amts)
'                                            tmGrf.lDollars(ilMonthNo - 1) = -tmGrf.lDollars(ilMonthNo - 1)
'                                        End If
'
'                                        mParPayAccumPayByContract CAT_OWED, tmGrf.iVefCode, tmGrf.iPerGenl(3), ilProdPct(ilLoop + 1)
'                                        For ilTemp = 0 To 12
'                                            tmGrf.lDollars(ilTemp) = 0
'                                        Next ilTemp
                                   End If

                               End If
                           End If                              'llProcessPct > 0
'                            ilLoop = ilHowMany
                       Next illoop                         'For ilLoop = 0 To ilHowMany - 1 Step 1
                   End If                                  'ilFoundMonth And ilMonthNo <= igPeriods
               End If                                      'ilMatchCntr And ilFoundOption And ilOKtoSeeVeh
           End If                                          '(Not (tmRvf.sTranType <> "PO") And (llDate >= llStdStartDates(12)))
       Next llRvfLoop                                      'For llRvfLoop = LBound(tlRvf) To UBound(tlRvf) - 1 Step 1
'    Next ilLoopRecv

'    'write out the contracts detail 6 records per vehicle
'    For ilLoopRecv = LBound(tmPayByContract) To UBound(tmPayByContract) - 1
'        tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
'        tmGrf.iGenDate(1) = igNowDate(1)
'        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
'        tmGrf.lGenTime = lgNowTime
'        tmGrf.lChfCode = tmChf.lCode           'contr internal code
'        tmGrf.lCode4 = lmSingleCntr             'contract #
'        tmGrf.iVefCode = tmPayByContract(ilLoopRecv).iVefCode
'        tmGrf.iCode2 = tmPayByContract(ilLoopRecv).iMnfPartCode
'        tmGrf.iAdfCode = tmChf.iAdfCode
'        tmGrf.iPerGenl(0) = TOTAL_BYCNT
'
'        tmGrf.iPerGenl(1) = CAT_GROSS
'        For ilTemp = 0 To 12
'            tmGrf.lDollars(ilTemp) = tmPayByContract(ilLoopRecv).lGross(ilTemp)
'        Next ilTemp
'        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
'
'        tmGrf.iPerGenl(1) = CAT_NET
'        For ilTemp = 0 To 12
'            tmGrf.lDollars(ilTemp) = tmPayByContract(ilLoopRecv).lNet(ilTemp)
'        Next ilTemp
'        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
'
'        tmGrf.iPerGenl(1) = CAT_COLLECTED
'        For ilTemp = 0 To 12
'            tmGrf.lDollars(ilTemp) = tmPayByContract(ilLoopRecv).lCollected(ilTemp)
'        Next ilTemp
'        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
'
'
'        tmGrf.iPerGenl(1) = CAT_OUTSTANDING
'        For ilTemp = 0 To 12
'            tmGrf.lDollars(ilTemp) = tmPayByContract(ilLoopRecv).lNet(ilTemp) - tmPayByContract(ilLoopRecv).lCollected(ilTemp)
'            tmPayByContract(ilLoopRecv).lOutstanding(ilTemp) = tmPayByContract(ilLoopRecv).lNet(ilTemp) - tmPayByContract(ilLoopRecv).lCollected(ilTemp)
'        Next ilTemp
'        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
'
'        tmGrf.iPerGenl(1) = CAT_PARTNER
'        For ilTemp = 0 To 12
'            tmGrf.lDollars(ilTemp) = tmPayByContract(ilLoopRecv).lPartner(ilTemp)
'        Next ilTemp
'        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
'
'        tmGrf.iPerGenl(1) = CAT_OWED
'        For ilTemp = 0 To 12
'            tmGrf.lDollars(ilTemp) = tmPayByContract(ilLoopRecv).lOwed(ilTemp)
'        Next ilTemp
'        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
'    Next ilLoopRecv
    Erase tlRvf
    Exit Sub
End Sub

'           gParPayOpenFiles - open files applicable to Participant Payables report
'
Function gParPayOpenFiles() As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilError As Integer
    Dim slStamp As String
    Dim illoop As Integer
    Dim slTemp As String
    
    ilError = False
    
    On Error GoTo gParPayOpenFilesErr
    '1-27-07 implement ntr selectivity in sales compare
    hmSbf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSbfRecLen = Len(tmSbf)
    
    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imGrfRecLen = Len(tmGrf)
    hmRvf = CBtrvTable(ONEHANDLE) 'CBtrvObj()            'read History files using RVF handles and buffers
    ilRet = btrOpen(hmRvf, "", sgDBPath & "Phf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imRvfRecLen = Len(tmRvf)
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imCHFRecLen = Len(tmChf)
    
    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSlfRecLen = Len(tmSlf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imVefRecLen = Len(tmVef)
    
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imMnfRecLen = Len(tmMnf)
    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSofRecLen = Len(tmSof)
    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imClfRecLen = Len(tmClf)
    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imCffRecLen = Len(tmCff)
    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imAgfRecLen = Len(tmAgf)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSdfRecLen = Len(tmSdf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imSmfRecLen = Len(tmSmf)
    
    'open VSF if contract needs to check whether user can see the contract
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilError = ilRet
    End If
    imVsfRecLen = Len(tmVsf)
    On Error GoTo 0
    
    imAdjustAcquisition = False                 'this is for the net-net or triple-net options
    
    imInclPolit = True
    imInclNonPolit = True
    
    imVehicle = False
    imVehGroup = False
    imAirTime = True        '11-25-02   unless asked, always include the air time
    imNTR = False           '11-25-02
    imHardCost = False      '3-21-05
    'do not adjust acquisition for the 4-line version (gross/net/comm/net-net), so it truly it a net-net report.
    'if acquistion is to be added later, it needs to show a new line with the acquistion costs
    'i.e. Lines will be:  gross, net, acq $, %comm, t-net
    'imAdjustAcquisition = True
    imVehicle = True
    
    '10-2-06
    imInclPolit = gSetCheck(RptSelParPay!ckcAllTypes(15).Value)       'include politicals
    imInclNonPolit = gSetCheck(RptSelParPay!ckcAllTypes(16).Value)       'include non-politicals
    
    ilRet = gObtainMnfForType("I", slStamp, tlMMnf())        'NTR Item types
    
    imHold = False
    imOrder = False
    imStandard = False
    imReserv = False
    imRemnant = False
    imDR = False
    imPI = False
    imPSA = False
    imPromo = False
    imTrade = False
    imNTR = False           '11-25-02
    imAirTime = True        '11-25-02  normal reports should always include Air Time (vs NTR) unless there is selectivity to
                            'exclude the Air Time
    If RptSelParPay!ckcAllTypes(0).Value = vbChecked Then
        imHold = True
    End If
    If RptSelParPay!ckcAllTypes(1).Value = vbChecked Then
        imOrder = True
    End If
    If RptSelParPay!ckcAllTypes(3).Value = vbChecked Then  'include std cntrs?
        imStandard = True
    End If
    If RptSelParPay!ckcAllTypes(4).Value = vbChecked Then   'include reserves?
        imReserv = True
    End If
    If RptSelParPay!ckcAllTypes(5).Value = vbChecked Then   'include remnants?
        imRemnant = True
    End If
    If RptSelParPay!ckcAllTypes(6).Value = vbChecked Then   'direct response?
        imDR = True
    End If
    If RptSelParPay!ckcAllTypes(7).Value = vbChecked Then   'per inquiry?
        imPI = True
    End If
    If RptSelParPay!ckcAllTypes(8).Value = vbChecked Then   'psa?
        imPSA = True
    End If
    If RptSelParPay!ckcAllTypes(9).Value = vbChecked Then  'promo?
        imPromo = True
    End If
    If RptSelParPay!ckcAllTypes(10).Value = vbChecked Then   'trades?
        imTrade = True
    End If
    
    If RptSelParPay!ckcAllTypes(11).Value = vbUnchecked Then   'Air Time?
        imAirTime = False
    End If
    'ckcAllTypes(12) = REP, hidden for now
    If RptSelParPay!ckcAllTypes(13).Value = vbChecked Then   'NTR
        imNTR = True
    End If
    If RptSelParPay!ckcAllTypes(14).Value = vbChecked Then   'hard cost
        imHardCost = True
    End If
    
    slStr = RptSelParPay!edcContract.Text     'selective contract
    
    If slStr = "" Then
        lmSingleUserCntr = 0
    Else
        lmSingleUserCntr = CLng(slStr)
    End If
    
    
    lgSTime1 = 0
    lgSTime2 = 0
    lgSTime3 = 0
    lgSTime4 = 0
    lgSTime5 = 0
    lgSTime6 = 0
    
    lgETime1 = 0
    lgETime2 = 0
    lgETime3 = 0
    lgETime4 = 0
    lgETime5 = 0
    lgETime6 = 0
    
    lgTtlTime1 = 0
    lgTtlTime2 = 0
    lgTtlTime3 = 0
    lgTtlTime4 = 0
    lgTtlTime5 = 0
    lgTtlTime6 = 0
    
    Exit Function
    
gParPayOpenFilesErr:
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmRvf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmSlf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmSof)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmSmf)
    btrDestroy hmGrf
    btrDestroy hmRvf
    btrDestroy hmCHF
    btrDestroy hmSlf
    btrDestroy hmVef
    btrDestroy hmMnf
    btrDestroy hmSof
    btrDestroy hmClf
    btrDestroy hmCff
    btrDestroy hmAgf
    btrDestroy hmSdf
    btrDestroy hmSmf
    gParPayOpenFiles = ilError
    Exit Function
End Function

'                   mParPayAdjust - Subtract missed spots (missed, cancel, hidden)
'                                Count MGs where they air
'                               Billed and booked report
'                   <input> llstdstartDates() - array of 13 start dates of the 12 months to gather
'                           ilFirstProjInx - index of first month to start projection (earlier is from receivables)
'                           llStartAdjust -  Earliest date to start searching for missed, etc.
'                           llEndAdjust - latest date to stop searchng for missed, etc.
'                           ilPkgPrice - true if retrieving rate from the matching clf and cff
'                                        false if retrieving rate from the package clf of the hidden clf (aired billing)
'                           ilSubMissed - true if subtract $ from missed vehicle for missed, cancel, hidden spots
'                           ilCountMGs - true if subtract $ from missed vehicle, move $ to makegood vehicle
'                           ilHowManyPer - # of periods to gather
'                                           (for billed & booked its 12,
'                                           for Sales Comparisons its max 3)
'                           ilAdjustMissedForMG - true to subtract out the missed part of the mg,
'                                       false to ignore the missed part of the MG.
'                                       When gathering spots, accumlating spot $ will be short if
'                                       the missed part is subtracted out
'
'       3/97 dh Comment out code to test the MNF Missed Reasons - field contains
'               value whether or not to bill the mg, bill the missed, etc.
'       3-15-00 Speed up the gathering of SMF spots; obtain the flight instead of going through the
'               generalized routines (which reads & rereads SMF) and could also have caused looping when
'               an SMF error existed.
'               Also, when retrieving the SMF, since the key is missed date, start searching for the spots
'               based on the start of the reporting period minus 2 months, as opposed to the beginning (zero).
'       8-17-06 count mg where they air is not checked for air time billing.  It was always included
'       11-12-07 add flag to test to adjust the missed portion of the makegood
Sub mParPayAdjust(llStdStartDates() As Long, ilFirstProjInx As Integer, llStartAdjust As Long, llEndAdjust As Long, ilPkgPrice As Integer, ilSubMissed As Integer, ilCountMGs As Integer, ilHowManyPer As Integer, ilAdjustMissedForMG As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilListIndex                   ilFoundOption                                           *
'******************************************************************************************
    Dim slDate As String
    ReDim ilDate(0 To 1) As Integer           'converted date for earliest start date for sdf keyread
    Dim ilRet As Integer
    Dim llDate As Long
    Dim ilMonthInx As Integer
    Dim ilFoundMonth As Integer
    Dim ilDoAdjust As Integer
    Dim slPrice As String           'rate from flight
    Dim ilSavePkLineNo As Integer
    Dim illoop As Integer
    Dim llActPrice As Long          'rate from flight
    Dim ilFoundVef As Integer
    Dim ilTemp As Integer
    
    Dim ilAcqCommPct As Integer
    Dim ilAcqLoInx As Integer
    Dim ilAcqHiInx As Integer
    Dim llAcqNet As Long
    Dim llAcqComm As Long
    Dim blAcqOK As Boolean
    Dim ilTempVefCode As Integer            '8-3-16
    
    '   rules to subtract missed spots
    '
    '                     Case 1 "S"        Case 2 "O"              Case 3 "A"
    '                   Bill as Order    Bill as Order             As Aired
    '                   update Order      Update aired             Update aired
    'Package Lines      ignore missed    Ignore missed             Ignore missed
    'Hidden Lines       ignore missed    Ignore missed             ignore missed
    'Standard Lines     ignore missed    Answer from user input    Answer from user input
    '
    'If tmClf.sType <> "O" And tmClf.sType <> "A" And ilSubMissed And smAirOrder <> "S" Then   'possibly adjust missed for standard lines                             'subtract missed, cancel, hidden spots?
                                                            'hidden & package lines should never subtract missed spots
    If tmClf.sType = "S" And ilSubMissed And smAirOrder <> "S" Then   'possibly adjust missed for standard lines                             'subtract missed, cancel, hidden spots?
        tmSdfSrchKey.iVefCode = tmClf.iVefCode
        tmSdfSrchKey.lChfCode = tgChfCT.lCode
        tmSdfSrchKey.iLineNo = tmClf.iLine
        tmSdfSrchKey.lFsfCode = 0
        slDate = Format$(llStartAdjust, "m/d/yy")
        gPackDate slDate, ilDate(0), ilDate(1)
        tmSdfSrchKey.iDate(0) = ilDate(0)
        tmSdfSrchKey.iDate(1) = ilDate(1)
        tmSdfSrchKey.sSchStatus = ""
        tmSdfSrchKey.iTime(0) = 0
        tmSdfSrchKey.iTime(1) = 0
        ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = tmClf.iVefCode) And (tmSdf.lChfCode = tgChfCT.lCode) And (tmSdf.iLineNo = tmClf.iLine)
            'If (tmSdf.sSchStatus = "C" Or tmSdf.sSchStatus = "M" Or tmSdf.sSchStatus = "U" Or tmSdf.sSchStatus = "R" Or slPass = "H") Then
            If (tmSdf.sSchStatus = "C" Or tmSdf.sSchStatus = "M" Or tmSdf.sSchStatus = "U" Or tmSdf.sSchStatus = "R" Or tmSdf.sSchStatus = "H") Then
                gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
                llDate = gDateValue(slDate)
                If llDate > llEndAdjust Then
                    Exit Do
                End If
                'spot is OK, adjust the $
                ilFoundMonth = False
                For ilMonthInx = ilFirstProjInx To ilHowManyPer Step 1         'loop thru months to find the match
                    If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                        ilFoundMonth = True
                        Exit For
                    End If
                Next ilMonthInx
                If ilFoundMonth Then
                    'Temporarily commented out to bypass testing the Missed reason flags for billing,
                    'Always go thru the adjustments
                     ilDoAdjust = True
                    'found a month that it falls within, should the missed be billed?
                    'ilDoAdjust = False
                    'For ilLoop = LBound(tmMnfList) To UBound(tmMnfList) Step 1
                    '    If ((tmMnfList(ilLoop).iMnfCode = tmSdf.iMnfMissed And tmMnfList(ilLoop).iBillMissMG <= 1) Or (tmSdf.iMnfMissed = 0)) Then
                    '    '1=bill mg, nc missed , 0 = nothing answered, default to same as 1
                    '        ilDoAdjust = True
                    '        Exit For
                    '    End If
                    'Next ilLoop
                End If
                If ilDoAdjust Then
                    If ilPkgPrice Then
                        '03-3-01 ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
                        ilRet = gGetRate(tmSdf, tmClf, hmCff, tmSmf, tmCff, slPrice)
                    Else                        'fake it out to get price from package (instead of hidden line)
                        ilSavePkLineNo = tmClf.iLine    'save the contents of the line id since it still may be
                                                        'needed
                        tmClf.iLine = tmClf.iPkLineNo
                        '03-13-01 ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
                        ilRet = gGetRate(tmSdf, tmClf, hmCff, tmSmf, tmCff, slPrice)
                        tmClf.iLine = ilSavePkLineNo
                    End If
    
                    If (InStr(slPrice, ".") <> 0) Then        'found spot cost
                    
                        llAcqNet = tmClf.lAcquisitionCost
                
                        llActPrice = gStrDecToLong(slPrice, 2)
                        lmProject(ilMonthInx) = lmProject(ilMonthInx) - llActPrice     'subtr missed $, since they won't be invoiced
                        '8-3-15 determine if acq is commissionable
                        lmAcquisition(ilMonthInx) = lmAcquisition(ilMonthInx) - llAcqNet
                    End If
    
                End If
            End If                          'sschstatus = C, M, H, U, R
            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop                                        'while BTRV_err none and contracts & lines match
    End If
    'Rules for Count mgs where they air
    '                     Case 1 "S"        Case 2 "O"              Case 3 "A"
    '                   Bill as Order    Bill as Order             As Aired
    '                   update Order      Update aired             Update aired
    'Package Lines      ignore MGs       Ignore mgs                Same as Case 2 (hidden Line)
    'Hidden Lines       ignore mgs       Count MGs where they      Same as Case 2 (hidden line)
    '                                    air($ sub from month of
    '                                    missed spots vehicle and
    '                                    moved to the msised month
    '                                    of the mgs spots vehicle.
    'Standard Lines     ignore mgs       Same as case 2 except      Same as Case 2 (std Line)
    '                                    $ in month vehicle moved to.
    'If (smAirOrder = "O" And tmClf.sType = "H") Or (smAirOrder = "O" And tmClf.sType = "S" And ilCountMGs) Or (smAirOrder = "A" And ilCountMGs) Then        '8-17-06 test to count mg where they air for aired billing
    If (smAirOrder = "O" And tmClf.sType = "H") Or (smAirOrder = "O" And tmClf.sType = "S" And ilCountMGs) Or (smAirOrder = "A") Then         '8-3-16 if bill as aired, test mg/out for where they should show
        'Do the "Outs" and "MGs"
        tmSmfSrchKey.lChfCode = tgChfCT.lCode
        tmSmfSrchKey.iLineNo = tmClf.iLine
        slDate = Format$(llStartAdjust - 60, "m/d/yy")   'cant use start of report period because when looking for the SMF by missed date key, the missed spot
                                                        'could be prior to the reporting period , so back it up 2 months
        gPackDate slDate, ilDate(0), ilDate(1)         '3-13-01
        tmSmfSrchKey.iMissedDate(0) = ilDate(0)
        tmSmfSrchKey.iMissedDate(1) = ilDate(1)
        '2-2-11 rechange to use start of sched line to find all the makegoods
        tmSmfSrchKey.iMissedDate(0) = tmClf.iStartDate(0)              '3-13-01 use start of period minus 2 months to adjust for missed spots
        tmSmfSrchKey.iMissedDate(1) = tmClf.iStartDate(1)
        ilRet = btrGetGreaterOrEqual(hmSmf, tmSmf, imSmfRecLen, tmSmfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        Do While (ilRet = BTRV_ERR_NONE) And (tmSmf.lChfCode = tgChfCT.lCode) And (tmSmf.iLineNo = tmClf.iLine)
            'test dates later in SDF
            'Find associated spot in SDF
            If (tmSmf.sSchStatus = "O" Or tmSmf.sSchStatus = "G") Then
                tmSdfSrchKey3.lCode = tmSmf.lSdfCode
                ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                If (ilRet = BTRV_ERR_NONE) Then
                    'spot is OK, assume to adjust $ to where spot was aired  (as aired billing)
                    gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slDate
    
                    'use Month spot moved to, or month missed from?
    '                If (smAirOrder = "O" And tmClf.sType = "H") Or (smAirOrder = "A" And tmClf.sType <> "S") Then
                    '8-3-16 test as aired billing : if standard line, is option to show mg where it was originally ordered? (for std lines: show mg where they air is Unchecked)
                    ilTempVefCode = tmSdf.iVefCode
                    If Not ilCountMGs Then      'if not counting mgs where they air; always use the original date
                        gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slDate
                        ilTempVefCode = tmClf.iVefCode
                    Else
                        If (smAirOrder = "O" And tmClf.sType = "H") Or (smAirOrder = "A" And tmClf.sType <> "S") Then
                            'use original missed date
                            gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slDate
                        End If
                    End If
                    llDate = gDateValue(slDate)
                    ilFoundMonth = False
    
                    For ilMonthInx = ilFirstProjInx To ilHowManyPer Step 1         'loop thru months to find the match
                        If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                            ilFoundMonth = True
                            Exit For
                        End If
                    Next ilMonthInx
                    ilDoAdjust = False
                    If ilFoundMonth Then
                        '4-7-17 if adjusting  makegoods where they air, only add it if  billing as aired
                        'if billing as ordered, adjustments overstates the $ in the future
                        If smAirOrder = "A" Then
                            ilDoAdjust = True
                        Else        'as ordered
                            'If llDate > llStdStartDates(ilFirstProjInx) Then   'date of mg in the future, already here because as ordered billing; do not adjust since already added in for past from receivables
                             gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slDate
                            '6-6-17 sldate test changed from > to >=
                             If ilMonthInx >= ilFirstProjInx And gDateValue(slDate) >= llStdStartDates(ilFirstProjInx) Then       'month found in the future, check if orig missed date is in the future or has been invoiced
                                'was missed portion already invoiced?
                                ilDoAdjust = True
                            Else
                                'Always go thru the adjustments
                                 ilDoAdjust = ilDoAdjust
                            End If
                        End If
                    End If
                    If ilDoAdjust Then
                        If ilPkgPrice Then
                            '03-13-01 ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
                            ilRet = gGetRate(tmSdf, tmClf, hmCff, tmSmf, tmCff, slPrice)
                        Else                        'fake it out to get price from package (instead of hidden line)
                            ilSavePkLineNo = tmClf.iLine    'save the contents of the line id since it still may be
                                                            'needed
                            tmClf.iLine = tmClf.iPkLineNo
                            'ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
                            ilRet = gGetRate(tmSdf, tmClf, hmCff, tmSmf, tmCff, slPrice)
                            tmClf.iLine = ilSavePkLineNo
                        End If
    
                        If (InStr(slPrice, ".") <> 0) Then         'found spot cost and its a selected vehicle
    
                            ilFoundVef = False
                            'setup vehicle that spot was moved to
                            '8-3-16 if standard line and show mg where they air is unchecked, use original vehicle
    
                            For ilTemp = LBound(tmAdjust) To UBound(tmAdjust) - 1 Step 1
                                If tmAdjust(ilTemp).iVefCode = ilTempVefCode Then
                                    ilFoundVef = True
                                    Exit For
                                End If
                            Next ilTemp
                            If Not (ilFoundVef) Then
                                ReDim Preserve tmAdjust(0 To imUpperAdjust) As ADJUSTLIST
                                tmAdjust(imUpperAdjust).iVefCode = ilTempVefCode
                                ilTemp = imUpperAdjust
                                imUpperAdjust = imUpperAdjust + 1
                            End If
    
                            'mg $ - if ordered (update ordered or aired), put mg in same month it was ordered
                            'if update as aired, put mg where it ran
                            llActPrice = gStrDecToLong(slPrice, 2)
                            'lmProject(ilMonthInx) = lmProject(ilMonthInx) + llActPrice     'add back in the mg that is invoiced
                            
                           
                            llAcqNet = tmClf.lAcquisitionCost
                            tmAdjust(ilTemp).lProject(ilMonthInx) = tmAdjust(ilTemp).lProject(ilMonthInx) + llActPrice    'add back in the mg that is invoiced
                            tmAdjust(ilTemp).lAcquisitionCost(ilMonthInx) = tmAdjust(ilTemp).lAcquisitionCost(ilMonthInx) + llAcqNet
                            'now do the missed portion of the mg
                            gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slDate
                            llDate = gDateValue(slDate)
                            ilFoundMonth = False
                            For ilMonthInx = 1 To ilHowManyPer Step 1         'loop thru months to find the match
                                If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                                    ilFoundMonth = True
                                    Exit For
                                End If
                            Next ilMonthInx
                            If ilFoundMonth And ilAdjustMissedForMG Then    'subtract out the missed portion of the mg?
                                If ilMonthInx >= ilFirstProjInx Then            'only adjust if its in the future
                                    lmProject(ilMonthInx) = lmProject(ilMonthInx) - llActPrice    'adjust the missed portion of the mg
                                    lmAcquisition(ilMonthInx) = lmAcquisition(ilMonthInx) - tmClf.lAcquisitionCost
                                End If
                            End If
                        End If
                    Else                    'month not found for mg, find missed part of adjustment
                        gUnpackDate tmSmf.iMissedDate(0), tmSmf.iMissedDate(1), slDate
                        llDate = gDateValue(slDate)
                        ilFoundMonth = False
            '**********
                        For ilMonthInx = ilFirstProjInx To ilHowManyPer Step 1         'loop thru months to find the match
                            If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                                ilFoundMonth = True
                                Exit For
                            End If
                        Next ilMonthInx
                        If ilFoundMonth And ilAdjustMissedForMG Then    'subtract out the missed portion of the mg?
                            'If tmMnfList(ilLoop).iBillMissMG <= 1 Or tmSdf.iMnfMissed = 0 Then     'nc missed, bill mg
                            ilDoAdjust = False
                            For illoop = LBound(tmMnfList) To UBound(tmMnfList) Step 1
                                'If ((tmMnfList(ilLoop).iMnfCode = tmSdf.iMnfMissed) And (tmMnfList(ilLoop).iBillMissMG <= 1 Or tmMnfList(ilLoop).iBillMissMG = 3)) Or (tmSdf.iMnfMissed = 0) Then
                                '1=bill mg, nc missed  , 3 = bill both missed & mg
                                    ilDoAdjust = True
                                    Exit For
                                'End If
                            Next illoop
                            If ilDoAdjust Then
                                If ilPkgPrice Then
                                    '03-13-01 ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
                                    ilRet = gGetRate(tmSdf, tmClf, hmCff, tmSmf, tmCff, slPrice)
                                Else                        'fake it out to get price from package (instead of hidden line)
                                    ilSavePkLineNo = tmClf.iLine    'save the contents of the line id since it still may be
                                                                    'needed
                                    tmClf.iLine = tmClf.iPkLineNo
                                    '03-13-01 ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)
                                    ilRet = gGetRate(tmSdf, tmClf, hmCff, tmSmf, tmCff, slPrice)
                                    tmClf.iLine = ilSavePkLineNo
                                End If
                                If (InStr(slPrice, ".") <> 0) Then         'found spot cost
                                    llActPrice = gStrDecToLong(slPrice, 2)
                                    If ilMonthInx >= ilFirstProjInx Then            'only adjust if its in the future
                                        lmProject(ilMonthInx) = lmProject(ilMonthInx) - llActPrice    'adjust the missed portion of the mg
                                        lmAcquisition(ilMonthInx) = lmAcquisition(ilMonthInx) - tmClf.lAcquisitionCost
                                    End If
                                End If
                            End If          'iladoadjust
                        End If              'ilfoundmonth and iladjustmissedformg
                    End If                      'if doadjust
                End If                          'btrv_err_none
            End If                              'schstatus = O,G
            ilRet = btrGetNext(hmSmf, tmSmf, imSmfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    DoEvents
End Sub

'                   Billed & Booked - determine how many slsp to split $,
'                       or how many participants (owners to split $)
'
'                   mParPayGetHowManyPast ilMatchSsCode, ilHowMany, ilHowManyDefined
'                   <input> ilMatchSSCode - valid only for owner option
'                   <output> ilHowMany - Total times to loop for splits
'                            ilHowManyDefined - # of elements having values to split
'                            (i.e. There may be only 2 slsp defined, but they are not
'                            in sequence in the arrays - element 1 used, element 2 unused, element 3 used)
'                   4-5-01 obtain participants rev share for net-netoption
'
Sub mParPayGetHowManyPast(ilMatchSSCode As Integer, ilHowMany As Integer, ilHowManyDefined As Integer, ilMnfSSCode() As Integer, ilMnfGroup() As Integer, ilProdPct() As Integer, blThisIsBilling As Boolean)  '7-6-05
    Dim illoop As Integer
    Dim ilRet As Integer
    'ReDim ilMnfSSCode(1 To 1) As Integer
    'ReDim ilMnfGroup(1 To 1) As Integer
    'ReDim ilProdPct(1 To 1) As Integer
    'Index zero ignored in arrays below
    ReDim ilMnfSSCode(0 To 1) As Integer
    ReDim ilMnfGroup(0 To 1) As Integer
    ReDim ilProdPct(0 To 1) As Integer
    Dim ilUse100pct As Integer              '8-21-07 use 100% of participant share if rvfmnfgroup is present
    Dim slStr As String
    Dim slCode As String
    Dim llDate As Long
    Dim ilDate(0 To 1) As Integer
    
    ilHowMany = 0
    ilHowManyDefined = 0
    
    'get the vehicle for this transaction
    '7-1-14 do not need vef here, using all the vefcode from other tables
    '            tmVefSrchKey.iCode = tmRvf.iAirVefCode
    '            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '7-6-05 create separate array of vehicle sales source, particpants & percentages since it maybe
    'altered if the transaction has already been split
    ilUse100pct = True             '8-21-07 dont use 100% for participant share, search for the participants %
    If blThisIsBilling Then
        gInitPartGroupAndPcts tmRvf.iAirVefCode, ilMatchSSCode, 0, ilMnfSSCode(), ilMnfGroup(), ilProdPct(), tmRvf.iTranDate(), tmPifKey(), tmPifPct(), ilUse100pct
    Else
        'collections need to use the ageing date
        slCode = Trim$(str$(tmRvf.iAgePeriod) & "/15/" & Trim$(str$(tmRvf.iAgingYear)))
        slStr = gObtainEndStd(slCode)
        llDate = gDateValue(slStr)
        gPackDateLong llDate, ilDate(0), ilDate(1)
        gInitPartGroupAndPcts tmRvf.iAirVefCode, ilMatchSSCode, 0, ilMnfSSCode(), ilMnfGroup(), ilProdPct(), ilDate(), tmPifKey(), tmPifPct(), ilUse100pct
    End If
    'For ilLoop = LBound(ilMnfSSCode) To UBound(ilMnfSSCode) Step 1
    For illoop = 1 To UBound(ilMnfSSCode) Step 1
        'ilHowManyDefined = 1       '9-27-17
        ilHowManyDefined = UBound(ilMnfSSCode)
    Next illoop
    'ilHowMany = 1      '9-27-17
    ilHowMany = UBound(ilMnfSSCode)
End Sub

'               mParPayRvfSelect - Billed & Booked: test selection of
'                   advertiser, vehicle or agency against the Phf/Rvf file
'
'               mBobSelect return = true if rvf record is valid
'                                      false if ignore record (not a selected one)
'               <input> ilUsePkg  - true if use package (billing vehicle) vs airing (hidden) vehicle
'
'                       ilSaveSS - Sales source to test for match if selectives one chosen
'                       ilSaveSOF - sales office code to test if selective ones chosen
Function mParPayRvfSelect(ilSaveSS As Integer, ilSaveSof As Integer) As Integer
    Dim ilFoundOption As Integer
    Dim ilTemp As Integer
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slCode As String

    ilFoundOption = False
 
    ilFoundOption = gTestIncludeExclude(tmRvf.iAirVefCode, imInclVefCodes, imUsevefcodes())
    If Not ilFoundOption Then
        mParPayRvfSelect = ilFoundOption
        Exit Function
    End If

    'if the record is valid for inclusion/exclusion of politicals and non-politicals
    If gIsItPolitical(tmRvf.iAdfCode) Then          'its a political, include this contract?
         If Not imInclPolit Then
            ilFoundOption = False
        End If
    Else                                                'not a political advt, include this contract?
         If Not imInclNonPolit Then
            ilFoundOption = False
        End If
    End If

    mParPayRvfSelect = ilFoundOption
End Function

'                   Billed & Booked - test contract types against
'                       the user requested types to include
'
'                   Return - true if valid contract to use
'                            false to ignore the contract
'
'           8-9-06 Make NTR a type of its own.  Do not check the contract
'           types as a condition for including/excluding NTR and Hard cost
Function mParPayTestTypes() As Integer
    Dim ilMatchCntr As Integer
    ilMatchCntr = True
    If tmChf.lCntrNo <> tmRvf.lCntrNo Or tmChf.sSchStatus <> "F" And tmChf.sSchStatus <> "M" And tmChf.sSchStatus <> "N" Or (lmSingleCntr > 0 And lmSingleCntr <> tmRvf.lCntrNo) Then
        ilMatchCntr = False
    End If

    If tmRvf.iMnfItem = 0 Then              'NTR are its own type, dont check contract types
        If tmChf.sStatus = "H" And Not imHold Then    'include holds
            ilMatchCntr = False
        End If
        If tmChf.sStatus = "O" And Not imOrder Then    'include orders
            ilMatchCntr = False
        End If
        If tmChf.sType = "C" And Not imStandard Then  'include std cntrs?
            ilMatchCntr = False
        End If
        If tmChf.sType = "V" And Not imReserv Then  'include reserves?
            ilMatchCntr = False
        End If
        If tmChf.sType = "T" And Not imRemnant Then  'include remnants?
            ilMatchCntr = False
        End If
        If tmChf.sType = "R" And Not imDR Then  'direct response?
            ilMatchCntr = False
        End If
        If tmChf.sType = "Q" And Not imPI Then  'per inquiry?
            ilMatchCntr = False
        End If
        If tmChf.sType = "S" And Not imPSA Then  'psa?
            ilMatchCntr = False
        End If
        If tmChf.sType = "M" And Not imPromo Then  'promo?
            ilMatchCntr = False
        End If
        'If tmChf.iPctTrade = 100 And Not imTrade Then  'trades?
        If tmRvf.sCashTrade = "T" And Not imTrade Then  'trades?
            ilMatchCntr = False
        End If
    End If
    mParPayTestTypes = ilMatchCntr
End Function

'                   mParPayBuildFlights - Loop through the flights of the schedule line
'                                   and build the projections dollars into lmprojmonths array
'                   <input> ilclf = sched line index into tgClfCt
'                           llStdStartDates() - 13 std month start dates
'                           ilFristProjInx - index of 1st month to start projecting
'                           ilHowManyPer - # entries containing a date to test in date array
'                   <output> lmProject = array of 12 months data corresponding to
'                                           12 std start months
'
'                   Returning:  lmProject - gross actual spot cost or gross acquisition costs
'                               lmAcquisition - gross acquisition costs
'                               lmAcquisitionNet - net acquisition costs if varying acq commissions; otherwise 0
'           6-8-08 subtract out acquisition $ from schedule line for net-net & triple net options
Sub mParPayBuildFlights(ilClf As Integer, llStdStartDates() As Long, ilFirstProjInx As Integer, ilHowManyPer As Integer)
    Dim ilCff As Integer
    Dim slStr As String
    Dim llFltStart As Long
    Dim llFltEnd As Long
    Dim illoop As Integer
    Dim llDate As Long
    Dim llDate2 As Long
    Dim llSpots As Long
    Dim ilTemp As Integer
    Dim llStdStart As Long
    Dim llStdEnd As Long
    Dim ilMonthInx As Integer
    Dim ilAcqCommPct As Integer
    Dim ilAcqLoInx As Integer
    Dim ilAcqHiInx As Integer
    Dim llAcqNet As Long
    Dim llAcqComm As Long
    Dim blAcqOK As Boolean

    llStdStart = llStdStartDates(ilFirstProjInx)
    llStdEnd = llStdStartDates(ilHowManyPer)
    ilCff = tgClfCT(ilClf).iFirstCff
    Do While ilCff <> -1
        tmCff = tgCffCT(ilCff).CffRec

        'first decide if its Cancel Before Start
        gUnpackDate tmCff.iStartDate(0), tmCff.iStartDate(1), slStr
        llFltStart = gDateValue(slStr)
        gUnpackDate tmCff.iEndDate(0), tmCff.iEndDate(1), slStr
        llFltEnd = gDateValue(slStr)
        If llFltEnd < llFltStart Then
            Exit Sub
        End If
        'gUnpackDate tmcff.iStartDate(0), tmcff.iStartDate(1), slStr
        'llFltStart = gDateValue(slStr)
        'backup start date to Monday
        illoop = gWeekDayLong(llFltStart)
        Do While illoop <> 0
            llFltStart = llFltStart - 1
            illoop = gWeekDayLong(llFltStart)
        Loop
        'gUnpackDate tmcff.iEndDate(0), tmcff.iEndDate(1), slStr
        'llFltEnd = gDateValue(slStr)
        'the flight dates must be within the start and end of the projection periods,
        'not be a CAncel before start flight, and have a cost > 0
        If (llFltStart < llStdEnd And llFltEnd >= llStdStart) And (llFltEnd >= llFltStart And (tmCff.lActPrice > 0 Or tmClf.lAcquisitionCost <> 0)) Then        '9-20-06
            'only retrieve for projections, anything in the past has already
            'been invoiced and has been retrieved from history or receiv files
            'adjust the gather dates from flights: use flight start date or requested start date, whichever is later
            If llStdStart > llFltStart Then
                llFltStart = llStdStart
            End If
            'use flight end date or requsted end date, whichever is lesser
            If llStdEnd < llFltEnd Then
                llFltEnd = llStdEnd
            End If
    
            For llDate = llFltStart To llFltEnd Step 7
                'Loop on the number of weeks in this flight
                'calc week into of this flight to accum the spot count
                If tmCff.sDyWk = "W" Then            'weekly
                    llSpots = tmCff.iSpotsWk + tmCff.iXSpotsWk
                Else                                        'daily
                    If illoop + 6 < llFltEnd Then           'we have a whole week
                        llSpots = tmCff.iDay(0) + tmCff.iDay(1) + tmCff.iDay(2) + tmCff.iDay(3) + tmCff.iDay(4) + tmCff.iDay(5) + tmCff.iDay(6)
                    Else
                        llFltEnd = llDate + 6
                        If llDate > llFltEnd Then
                            llFltEnd = llFltEnd       'this flight isn't 7 days
                        End If
                        For llDate2 = llDate To llFltEnd Step 1
                            ilTemp = gWeekDayLong(llDate2)
                            llSpots = llSpots + tmCff.iDay(ilTemp)
                        Next llDate2
                    End If
                End If
                'determine month that this week belongs in, then accumulate the gross and net $
                'currently, the projections are based on STandard bdcst
                For ilMonthInx = ilFirstProjInx To 12 Step 1         'loop thru months to find the match
                    If llDate >= llStdStartDates(ilMonthInx) And llDate < llStdStartDates(ilMonthInx + 1) Then
                        
                        If tgClfCT(ilClf).ClfRec.sType = "E" Then            'pckage pricing where spot rate equals the entire price for the week
                            lmProject(ilMonthInx) = lmProject(ilMonthInx) + tmCff.lActPrice
                            lmAcquisition(ilMonthInx) = lmAcquisition(ilMonthInx) + (tmClf.lAcquisitionCost)
                        Else
                            lmProject(ilMonthInx) = lmProject(ilMonthInx) + (llSpots * tmCff.lActPrice)         'gross actual spot price
                            lmAcquisition(ilMonthInx) = lmAcquisition(ilMonthInx) + (tgClfCT(ilClf).ClfRec.lAcquisitionCost * llSpots)   'acq gross
                            Exit For
                        End If
                    End If
                    Next ilMonthInx
                Next llDate                                     'for llDate = llFltStart To llFltEnd
        End If                                          '
        ilCff = tgCffCT(ilCff).iNextCff                   'get next flight record from mem
    Loop                                            'while ilcff <> -1
    Exit Sub
End Sub

'**********************************************************************
'
'            Billed & Booked Report - Gather Projection data
'            mParPayBuildProj
'
'           BILLED AND BOOKED
'           BILLED AND BOOKED RECAP
'           BILLED AND BOOKED COMPARISIONS
'           SALES COMPARISON
'           COMMISSION PROJECTIONS
'
'            Loop thru contracts within date last std bdcst billing
'            through the end of the year and build up to 12 periods
'
'                   <Input>  llStdStartDates - array of max 13 start dates, denoting
'                                              start date of each period to gather
'                            llLastBilled - Date of last invoice period
'                            ilHowManyPer - # of periods to gather
'                                           (for billed & booked its 12,

'               3/23/98 DH changed reference of Under/Over slsp comm from long to integer
'    12/10/98:  DH (Gather all makegoods)-
'    Back up the start date for gathering of contracts by one month.  This will include
'    the case where the contract has ended, but makegood spots are in the next month.
'    If the contract spanned the first quarter requested, the makegood spot outside the
'    end date of the contract was gathered.  but if the contract was outside the requested
'    first quarter, the makegood was not gathered.
'    6/14/99 Fix projections by package lines (note:  the pkg vehicles must have their
'           sales source, owner , sort code defined)
'    4-20-98 DH New commission structure with new scf file by vehicle/slsp
'    8-4-00 Split participants when option by vehicle (by option)
'   11-13-03 test if user allowed to see vehicle
'   1-31-04 fix subscript out of range for B& B Sales Commission Projection.  It was processing as
'           tho it needed to use participant index (which only goes thru 8), and slsp index as
'           being passed (which is 10)
'   11-21-05 Add Pacing to B & B
'   4-25-06 add vehicle option to sales comparison
'   6-8-06 Adjust acquisition $ (triple-net for vehicle option & slsp/vehicle option,net-net for vehicle/participant option
'   10-2-06 option to incl/excl polit & non politic
'   5-10-07 implement new design of participants
'
'   GRF fields:
'   grfgenDate - generation date
'   grfGenTime - generation time
'   grfChfCode - contract code
'   grfDateGenl - Date billed or paid
'   grfVefCode - vehicle code (airing or billing)
'   grfAdfCode = advertiser code
'   grfSlfCode = salesperson code
'   grfSofCode = Sales Source
'   grfCode2 - Sales Comparison:  (major or primary sort selection) business category or production cod or agency code or vehicle goup
'   grfCode4 - Sales Commission
'   grfYear - split salesperson/owner flag
'   grfDateType = C = cash, T = trade
'   grfDollars(1-12) 12 months $
'   grfDollars(13) total year $ (gross or net)
'   grsDollars(14-17) Q1 - Q4 $
'   grfDollars(18) - year total $ always net
'   grfPerGenl(1) - 0 = no comparisons , 1 = base date, 2 = comparison date, 3 = last year actual for pacing sales comparison
'   grfPerGenl(2) - new/return flag
'   GrfPerGenl(3) - Minor vehicle group
'   GrfPerGenl(4) - major vehicle group
'   grfPerGenl(5) - Flag for sorting (Net/Net Billed & Booked) 1 = gross, 2 = net, 3= comm
'   grfPerGenl(7) -Participant %
'   grfPerGenl(8) - # periods
'   grfPerGenl(9)-  Is it hard cost (true/false) 3-22-05
'   grfPerGenl(10) - 4-20-05 subsort field for B & B Recap:  1 = airtime, 2 = NTR , 3 = hard cost NTR
'   grfPerGenl(10) - 11-06-06 changed to:  0 = not NTR or Hard cost, > 0 = MNF item code for NTR
'                   if NTR, flag in NTR indicates if agy sales, direct sales or NTR sales.  If
'                   agy or direct sales, embed them in those categories and not with NTR.
'   grfPerGenl(11) - B & B recap 11-06-06 0 = direct, 1 = airtime, 2 = NTR, 3 = political, 4 = Hard cost
'   grfPerGenl(12) - Sales Comparison:  (minor or sub sort selection) business category or production cod or agency code or vehicle goup
'   grfPerGenl(13) - For B & B Comparison: 0 = data from rvf/phf & contracts, 1 = budgets
'   grfPerGenl(14) - Budget mnfcode selected for BOB Comparisons
'   grfPerGenl(18) - # periods to print
'Sales Comparison selection:
'   (imPrimSort) cbcSet1 Index: 0 = Advt, 1 = Agy, 2 = Bus Cat, 3 = Prod Prot, 4 = Slsp, 5 = VEhicle, 6 = Vehicle group
'   (imSubSort) cbcSet2 Index:  0 = none, 1 = Advt, 2 = Agy, 3 = Bus Cat, 4 = Prod Prot, 5 = Slsp, 6 = VEhicle, 7 = Vehicle group
'   lbcSelection(1)= agency,  lbcSelection(2) = Slsp, lbcSelection(3) = Bus Cat, lbcSelection(12) = Vehicle group (single select), 3-18-16 chg from lbcselection(4),
'   lbcSelection(5) = Advt, lbcSelection(6) = vehicle, lbcSelection(7) = prod protection
'******************************************************************************************************
Function mParPayBuildProj(llStdStartDates() As Long, llLastBilled As Long, ilHowManyPer As Integer, blNewContract As Boolean)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slNameCode                    ilIsItHardCost                llDate1                   *
'*  llDate2                       tmOneVefAllYear                                         *
'******************************************************************************************
    Dim ilRet As Integer
    Dim slStdStart As String            'start date to gather (std start)
    Dim slStdEnd As String              'end date to gather (end of std year)
    Dim slTempStart As String           'Start date of period requested (minus 1 month) to handle makegoods outside
                                        'end date of contrct
    Dim slTempEnd As String            'end date of period + 3 months to ensure that pacing changes have been included
    Dim slCntrStatus As String          'list of contract status to gather (working, order, hold, etc)
    Dim slCntrType As String            'list of contract types to gather (Per inq, Direct Response, Remnants, etc)
    Dim ilHOState As Integer            'which type of HO cntr states to include (whether revisions should be included)
    Dim llContrCode As Long
    Dim ilCurrentRecd As Integer
    Dim ilFound As Integer
    Dim ilFoundOption As Integer
    Dim illoop As Integer
    Dim ilHowMany As Integer            '# times to loop and process line (up to 10 times for split slsp),
                                        'up to 3 times for 3 diff. owners per vehicle, else loop once per
                                        'flight (or sch line)
    Dim ilHowManyDefined As Integer     '# of percentages (slsp for each contract, or participants for each vehicle)
    Dim ilMatchSSCode As Integer        'Sales Source code to process for participants
    Dim ilClf As Integer                'loop count for lines
    Dim slCode As String
    Dim ilTemp As Integer
    Dim ilScratch As Integer
    Dim llStdStart As Long              'requested start date to gather (serial date)
    Dim llStdEnd As Long                'requested end date to gather (serial date)
    Dim ilCorT As Integer
    Dim ilStartCorT As Integer
    Dim ilEndCorT As Integer
    Dim slCashAgyComm As String
    Dim slPctTrade As String
    'Dim llLastBilled As Long            'last date billed as starting point of projections
    Dim ilFoundOne As Integer
    Dim llProcessPct As Long             'percent of slsp split, ownership, else 100.0000%
    'ReDim lmProjMonths(1 To 13) As Long    'jan-dec + total all months
    'ReDim llStdStartDates(1 To 13) As Long
    Dim ilFirstProjInx As Integer
    'Dim slGrossOrNet As String             'commissions calc from G = Gross or N = net
    Dim slCommPct As String                 'for comm proj only - slsp comm % xx.xxxx
    Dim ilSubMissed As Integer             'subtract misses (misses, cancel, hidden)
    Dim ilCountMGs As Integer                'count mgs where they air
    Dim ilAdjust As Integer
    Dim ilVehLoop As Integer                'loop for MGs by vehicle
    'Dim llTempProject(1 To 12) As Long      '12 months projection or MG $
    'Dim llTempAcquisition(1 To 12) As Long
    Dim llTempProject(0 To 12) As Long      '12 months projection or MG $
    Dim llTempAcquisition(0 To 12) As Long
    'Dim llAcquisitionNet(1 To 12) As Long
    'Dim llAcquisitionComm(1 To 12) As Long
    
    Dim ilNewCode As Integer                    'Revenue area "New" mnf code to show designation on report
    'Dim llSingleCntr As Long                'user input single contract #
    'ReDim tlSlf(0 To 0) As SLF                  '4-20-00
    'Dim tlSBFType As SBFTypes
    Dim tlSplitInfo As splitinfo
    ReDim tlSbf(0 To 0) As SBF                '9-30-02
    Dim ilOKtoSeeVeh As Integer                 '11-13-03 flag to detrmine if user allowed to see vehicle
                        'else false to use revenue share to split sls $
    Dim ilAdjustDays As Integer             '11-21-05
    ReDim ilValidDays(0 To 6) As Integer
    Dim tlPriceTypes As PRICETYPES
    Dim ilAdjustMissedForMG As Integer  'flag to adjust the missed spots for makegoods.
                                        'for all B & B options except Cal (spots), the
                                        'missed part of the mg is subtracted.  When
                                        'gathering spots, ignore the missed part
    Dim ilSaveNTRFlag As Integer
    Dim ilContinue As Integer
    Dim ilSlfRecd As Integer
    Dim ilWriteForEachLine As Integer       '3-25-13 Slsp option only, write out prepass record each sched line, or on entire contract
    
'    ilRet = gObtainVef()                  '4-2-00 buildglobal vehicle table
'    If ilRet = 0 Then
'        btrDestroy hmVef
'        btrDestroy hmCHF
'        btrDestroy hmSlf
'        btrDestroy hmRvf
'        btrDestroy hmGrf
'        btrDestroy hmVef
'        Exit Function
'    End If
'    ilRet = gObtainSlf(RptSelParPay, hmSlf, tmSlfList())

'    'build array of selling office codes and their sales sources.  This is the most major sort
'    'in the Business Booked reports
'    ilTemp = 0
'    ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
'    Do While ilRet = BTRV_ERR_NONE
'        ReDim Preserve tmSofList(0 To ilTemp) As SOFLIST
'        tmSofList(ilTemp).iSofCode = tmSof.iCode
'        tmSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
'        ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        ilTemp = ilTemp + 1
'    Loop


    ilSubMissed = gSetCheck(RptSelParPay!ckcSubUnresolveMiss.Value)            'subt misses
    ilCountMGs = gSetCheck(RptSelParPay!ckcCountMGWhereAir.Value)             'count mgs
    smAirOrder = tgSpf.sInvAirOrder         'bill as ordred, aired
    If smAirOrder = "S" Then     'bill as ordered (update as order), no adjustments for makegoods/missed
        ilAdjust = False                    'always ignore missed and makegoods
    Else
        ilAdjust = True                     'missed or mg must be adjusted
    End If


    ReDim llCalSpots(0 To 0) As Long        'init buckets for daily calendar values
    ReDim llCalAmt(0 To 0) As Long
    ReDim llCalAcqAmt(0 To 0) As Long
    ReDim llCalAcqNetAmt(0 To 0) As Long

    ilAdjustMissedForMG = True
           
    ilFirstProjInx = 0
    'determine first month to project based on the the billing period
    For illoop = 1 To ilHowManyPer Step 1
        If llLastBilled > llStdStartDates(illoop) And llLastBilled < llStdStartDates(illoop + 1) Then
            ilFirstProjInx = illoop + 1
            slStdStart = Format$(llStdStartDates(ilFirstProjInx), "m/d/yy")
            Exit For
        End If
    Next illoop
    If ilFirstProjInx = 0 Then
        ilFirstProjInx = 1                          'all projections, no actuals
    End If
    If llLastBilled >= llStdStartDates(ilHowManyPer + 1) Then   'all data was in the past only, dont do contracts
        Exit Function
    End If

    
    If Not imNTR And Not imHardCost Then           '3-22-05 is ntr or hard cost selected?
        tmSBFType.iNTR = False           'no, ignore NTR projected $
    End If
 
    llStdStart = llStdStartDates(ilFirstProjInx)  'first date to project
    llStdEnd = llStdStartDates(ilHowManyPer + 1)                'end date to project
    slStdStart = Format$(llStdStart, "m/d/yy")
    slStdEnd = Format$(llStdEnd, "m/d/yy")

'    slCntrStatus = ""                 'statuses: hold, order, unsch hold, uns order
'    If imHold Then                  'exclude holds and uns holds
'        slCntrStatus = "HG"             'include orders and uns orders
'    End If
'    If imOrder Then                  'exclude holds and uns holds
'        slCntrStatus = slCntrStatus & "ON"             'include orders and uns orders
'    End If
'    slCntrType = ""
'    If RptSelParPay!ckcSelC5(0).Value = vbChecked Then
'        slCntrType = "C"
'    End If
'    If RptSelParPay!ckcSelC5(1).Value = vbChecked Then
'        slCntrType = slCntrType & "V"
'    End If
'    If RptSelParPay!ckcSelC5(2).Value = vbChecked Then
'        slCntrType = slCntrType & "T"
'    End If
'    If RptSelParPay!ckcSelC5(3).Value = vbChecked Then
'        slCntrType = slCntrType & "R"
'    End If
'    If RptSelParPay!ckcSelC5(4).Value = vbChecked Then
'        slCntrType = slCntrType & "Q"
'    End If
'    If RptSelParPay!ckcSelC5(5).Value = vbChecked Then
'        slCntrType = slCntrType & "S"
'    End If
'    If RptSelParPay!ckcSelC5(6).Value = vbChecked Then
'        slCntrType = slCntrType & "M"
'    End If
'    If slCntrType = "CVTRQSM" Then          'all types: PI, DR, etc.  except PSA(p) and Promo(m)
'        slCntrType = ""                     'blank out string for "All"
'    End If
'    ilHOState = 2                       'get latest orders & revisions  (HOGN plus any revised orders WCI)
    'build table (into tlchfadvtext) of all contracts that fall within the dates required
    'Back up the start date for gathering of contracts by one month.  This will include
    'the case where the contract has ended, but makegood spots are in the next month.
    'If the contract spanned the first quarter requested, the makegood spot outside the
    'end date of the contract was gathered.  but if the contract was outside the requested
    'first quarter, the makegood was not gathered.

    ilAdjustDays = 30       'adjustment for the days back to retrieve contracts
 
    ReDim tmCntAllYear(0 To 0) As ALLPIFPCTYEAR

    slTempStart = Format$((gDateValue(slStdStart) - ilAdjustDays), "m/d/yy")
    slTempEnd = Format$((gDateValue(slStdEnd) + 90), "m/d/yy")
'    ilRet = gObtainCntrForDate(RptSelParPay, slTempStart, slTempEnd, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())
    
    For illoop = 1 To 13 Step 1
        lmProject(illoop) = 0
        lmAcquisition(illoop) = 0
        lmAcquisitionNet(illoop) = 0
    Next illoop
    
    'only 1 contract processed at a time for phf, rvf & chf
    ReDim tlChfAdvtExt(0 To 0) As CHFADVTEXT
    tmChfSrchKey1.lCntrNo = lmSingleCntr
    tmChfSrchKey1.iCntRevNo = 32000
    tmChfSrchKey1.iPropVer = 32000
    ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = lmSingleCntr)
        If (tmChf.sDelete <> "Y") And (tmChf.sStatus = "H" Or tmChf.sStatus = "O" Or tmChf.sStatus = "G" Or tmChf.sStatus = "N") Then
            tlChfAdvtExt(0).lCode = tmChf.lCode
            tlChfAdvtExt(0).iAdfCode = tmChf.iAdfCode
            ReDim Preserve tlChfAdvtExt(0 To 1) As CHFADVTEXT
            Exit Do
        End If
        ilRet = btrGetNext(hmCHF, tmChf, imCHFRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1                                          'loop while llCurrentRecd < llRecsRemaining
        ilFound = True
        'if the record is valid for inclusion, see if Sales Comparison asks for Political only
        If gIsItPolitical(tlChfAdvtExt(ilCurrentRecd).iAdfCode) Then          'its a political, include this contract?
             If Not imInclPolit Then
                ilFound = False
            End If
        Else                                                'not a political advt, include this contract?
             If Not imInclNonPolit Then
                ilFound = False
            End If
        End If

        If (ilFound) Then
            llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
            ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChfCT, tgClfCT(), tgCffCT(), False)  '8-28-12 do not sort by special dp order
            If (tgChfCT.iPctTrade < 100) Or (tgChfCT.iPctTrade = 100 And imTrade) Then         'all trade contr, include?
                'obtain agency for commission
                If (tgChfCT.iAgfCode > 0) Then   'if direct advert , dont take any agency comm out
                    tmAgfSrchKey.iCode = tgChfCT.iAgfCode
                    ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching agy recd
                    If ilRet = BTRV_ERR_NONE Then
                        slCashAgyComm = gIntToStrDec(tmAgf.iComm, 2)
                    End If          'ilret = btrv_err_none
                Else
                    slCashAgyComm = ".00"
                End If              'iagfcode > 0

                slPctTrade = gIntToStrDec(tgChfCT.iPctTrade, 0)
                If tgChfCT.iPctTrade = 0 Then                     'setup loop to do cash & trade
                    ilStartCorT = 1
                    ilEndCorT = 1
                ElseIf tgChfCT.iPctTrade = 100 Then
                    ilStartCorT = 2
                    ilEndCorT = 2
                Else
                    ilStartCorT = 1
                    If imTrade Then             'Include trades?
                        ilEndCorT = 2
                    Else
                        ilEndCorT = 1
                    End If
                End If
 
                If tgChfCT.iSlfCode(0) <> tmSlf.iCode Then        'only read slsp recd if not in mem already
                    '7-2-14 use binarysearch for speed
                    ilRet = gBinarySearchSlf(tgChfCT.iSlfCode(0))
                    If ilRet = -1 Then              'not found
                        tmSlf.iSofCode = 0
                    Else
                        tmSlf = tgMSlf(ilRet)
                    End If
                End If                                          'table of selling offices built into memory with its
                                                            'associated sales source
                For illoop = LBound(tmSofList) To UBound(tmSofList)
                    If tmSofList(illoop).iSofCode = tmSlf.iSofCode Then
                        ilMatchSSCode = tmSofList(illoop).iMnfSSCode          'Sales source
                        Exit For
                    End If
                Next illoop

                ReDim tmAdjust(0 To 0) As ADJUSTLIST             'prepare list of mgs
                imUpperAdjust = 0

                'only go thru SBF if NTR should be included and the contract header has it flagged as an NTR order
                '1-27-07 implement NTR/hard cost into Sales Comparison
                    ilContinue = False
                    ilSaveNTRFlag = tmSBFType.iNTR
                    ReDim tmSBFAdjust(0 To 0) As ADJUSTLIST             'build new for every contract

                    If (tgChfCT.sInstallDefined = "Y" And (Asc(tgSpf.sUsingFeatures6) And INSTALLMENTREVENUEEARNED) <> INSTALLMENTREVENUEEARNED) Then     'its an install contract, if install method is bill as aired, get the NTRs from SBF                                                                 'if install bill method is Invoiced, get the installment records from SBF
                        'install method is invoiced, get the future from Installment records
                        tmSBFType.iNTR = False
                        tmSBFType.iInstallment = True
                        ilContinue = True
                        ilRet = gObtainSBF(RptSelParPay, hmSbf, tgChfCT.lCode, slStdStart, slStdEnd, tmSBFType, tlSbf(), 0) '11-28-06 add last parm to indicate which key to use
                        'Build array of the vehicles and their NTR $ into tmSBFAdjust array
                        mParPaySbfAdjustForInstall tlSbf(), llStdStartDates(), ilFirstProjInx, llStdStart, llStdEnd, igPeriods

                    Else
                        If ((tgChfCT.sNTRDefined = "Y") And (imNTR = True Or imHardCost = True)) Then
                            ilRet = gObtainSBF(RptSelParPay, hmSbf, tgChfCT.lCode, slStdStart, slStdEnd, tmSBFType, tlSbf(), 0) '11-28-06 add last parm to indicate which key to use
                            'Build array of the vehicles and their NTR $ into tmSBFAdjust array
                            mParPaySbfAdjustForNTR tlSbf(), llStdStartDates(), ilFirstProjInx, llStdStart, llStdEnd, igPeriods
                            ilContinue = True
                        End If
                    End If
                    If ilContinue Then
                        'ReDim tmSBFAdjust(0 To 0) As ADJUSTLIST             'build new for every contract
                        'ilRet = gObtainSBF(RptSelParPay, hmSbf, tgChfCT.lCode, slStdStart, slStdEnd, tlSBFType, tlSbf(), 0) '11-28-06 add last parm to indicate which key to use
                        'Build array of the vehicles and their NTR $ into tmSBFAdjust array
                        'mParPaySbfAdjustForNTR tlSbf(), llStdStartDates(), ilFirstProjInx, llStdStart, llStdEnd, ilHowManyPer, ilListIndex, tlMnf()
                        'loop on sbf to process each $
                        For ilVehLoop = LBound(tmSBFAdjust) To UBound(tmSBFAdjust) - 1      '11-06-06 chg to ubound -1
                            ilFoundOption = True                    'assume everything included
                              ilFoundOption = False
                              'If gFilterLists(tmSBFAdjust(ilVehLoop).iVefCode, imIncludeCodes, imUseCodes()) Then
                              If gFilterLists(tmSBFAdjust(ilVehLoop).iVefCode, imInclVefCodes, imUsevefcodes()) Then
                                  ilFoundOption = True
                              End If
                         
                              If ilFoundOption Then
                                  tlSplitInfo.iMatchSSCode = ilMatchSSCode
                                  tlSplitInfo.iStartCorT = ilStartCorT
                                  tlSplitInfo.iEndCorT = ilEndCorT
                                  tlSplitInfo.iFirstProjInx = ilFirstProjInx
                                  tlSplitInfo.iVefCode = tmSBFAdjust(ilVehLoop).iVefCode
                                  tlSplitInfo.iNewCode = ilNewCode
                                  tlSplitInfo.sPctTrade = slPctTrade
                                  If ((tgChfCT.sInstallDefined = "Y" And (Asc(tgSpf.sUsingFeatures6) And INSTALLMENTREVENUEEARNED) = INSTALLMENTREVENUEEARNED)) Or tgChfCT.sInstallDefined <> "Y" Then     'its an install contract, if install method is bill as aired (separate inv from revenue), get the NTRs from SBF                                                                 'if install bill method is Invoiced, get the installment records from SBF
                                      'this is revenue for NTR / Hard Cost in the future
                                      tlSplitInfo.sTradeAgyComm = tmSBFAdjust(ilVehLoop).sAgyComm
                                      If tmSBFAdjust(ilVehLoop).sAgyComm = "Y" Then
                                          tlSplitInfo.sCashAgyComm = slCashAgyComm
                                      Else
                                          tlSplitInfo.sCashAgyComm = "00.00"
                                      End If
                                      tlSplitInfo.iNTRSlspComm = tmSBFAdjust(ilVehLoop).iSlsCommPct      '2-2-04 NTR slsp comm if not using sub-companies.
                                      tlSplitInfo.sNTR = "Y"
                                      tlSplitInfo.iHardCost = tmSBFAdjust(ilVehLoop).iIsItHardCost      'true false if hard cost
                                      tlSplitInfo.iMnfNTRItemCode = tmSBFAdjust(ilVehLoop).iMnfItem       '11-06-06 NTR Item type from MNF
                                  Else           'otherwise its billed as invoiced and need to get future from NTR and lines
                                      'this is SBF records for Installment in the future
                                      'tlSplitInfo.sCashAgyComm = slCashAgyComm
                                      If tmSBFAdjust(ilVehLoop).sAgyComm = "Y" Then
                                          tlSplitInfo.sCashAgyComm = slCashAgyComm
                                      Else
                                          tlSplitInfo.sCashAgyComm = "00.00"
                                      End If
                                      tlSplitInfo.sTradeAgyComm = tgChfCT.sAgyCTrade        '12-23-02
                                      tlSplitInfo.iNTRSlspComm = 0            '2-2-04
                                      tlSplitInfo.sNTR = "N"
                                      tlSplitInfo.iHardCost = False
                                      tlSplitInfo.iMnfNTRItemCode = 0         '11-06-06 not an NTR
                                  End If

                                  For illoop = 1 To 12
                                      lmProject(illoop) = tmSBFAdjust(ilVehLoop).lProject(illoop)
                                      lmAcquisition(illoop) = tmSBFAdjust(ilVehLoop).lAcquisitionCost(illoop)
                                  Next illoop
                                  '2-3-03 fake out the schedule line since there isnt a schedule line that contains vefcode
                                  tmClf.iVefCode = tlSplitInfo.iVefCode   'other subrtn use the vehicle code from line (SBF doesnt have sch lines)
                                  tmClf.sType = "S"       'standard line
                                  'get the percent splits
                                  mParPaySetupPctOfSplit tlSplitInfo, llStdStartDates(), tmCntAllYear(), tmOneVehAllYear()
                              End If
                    
                        Next ilVehLoop
                    End If
                    tmSBFType.iNTR = ilSaveNTRFlag
                    tmSBFType.iInstallment = False  'installment flag to retrieve SBF records are only set when getting installment records for the future based on install method "Invoiced"
        

                'Insure the common monthly buckets are initialized for the schedule lines
                For illoop = 1 To 12
                    lmProject(illoop) = 0
                    lmAcquisition(illoop) = 0
                    lmAcquisitionNet(illoop) = 0
                Next illoop
                '2-27-08 process the schedule lines if air time included and its not a installment contract, or
                'air time included and its an installment method is Aired (separate invoicing from revenue); get the
                'future from schedule lines.  If installment whose method is invoiced (inv = revenue),
                If (imAirTime And tgChfCT.sInstallDefined <> "Y") Or (imAirTime And tgChfCT.sInstallDefined = "Y" And (Asc(tgSpf.sUsingFeatures6) And INSTALLMENTREVENUEEARNED) = INSTALLMENTREVENUEEARNED) Then                    '11-25-02 include air time (vs NTR)?

                    'build the vehicles participant tables for this contract only into tmCntAllYear, if by owner
                        gInitCntPartYear hmVsf, tgChfCT, ilMatchSSCode, llStdStartDates(), tmCntAllYear(), tmPifKey(), tmPifPct()
                    'End If

                    For ilClf = LBound(tgClfCT) To UBound(tgClfCT) - 1 Step 1
                        tmClf = tgClfCT(ilClf).ClfRec
                        ilFoundOption = False
                        tmGrf.iVefCode = tmClf.iVefCode
                        If gFilterLists(tmClf.iVefCode, imInclVefCodes, imUsevefcodes()) Then
                            If tmClf.sType <> "A" And tmClf.sType <> "O" And tmClf.sType <> "E" Then
                                ilFoundOption = True
                            End If
                        End If
'                        If (ilFoundOption) Then
'                            'use hidden lines
'                            If tmClf.sType <> "A" And tmClf.sType <> "O" And tmClf.sType <> "E" Then
'                                mParPayBuildFlights ilClf, llStdStartDates(), ilFirstProjInx, ilHowManyPer + 1
'                            End If
'
'                            'lmProject - gross actual spot cost or gross acquisition costs
'                            'lmAcquisition - gross acquisition costs
'                            'lmAcquisitionNet - net acquisition costs if varying acq commissions; otherwise 0
'                            'gSetAcqFieldsForOutput imAdjustAcquisition, imUseAcquisitionCost, smGrossOrNet, lmProject(), lmAcquisition(), lmAcquisitionNet()
'
'
''                            'Adjust with misses and makegoods?
''                            If ilAdjust Then
''                                mParPayAdjust llStdStartDates(), ilFirstProjInx, llStdStart, llStdEnd, True, ilSubMissed, ilCountMGs, ilHowManyPer, ilAdjustMissedForMG
''                            End If
'
'                        End If
'                        If (ilClf = UBound(tgClfCT) - 1) Then
                            If ilFoundOption Then             'matching vehicle or owner found
                              'use hidden lines
                              'If tmClf.sType <> "A" And tmClf.sType <> "O" And tmClf.sType <> "E" Then

                                  mParPayBuildFlights ilClf, llStdStartDates(), ilFirstProjInx, ilHowManyPer + 1
                                'Adjust with misses and makegoods?
                                If ilAdjust Then
                                    mParPayAdjust llStdStartDates(), ilFirstProjInx, llStdStart, llStdEnd, True, ilSubMissed, ilCountMGs, ilHowManyPer, ilAdjustMissedForMG
                                End If
                                tlSplitInfo.iMatchSSCode = ilMatchSSCode
                                tlSplitInfo.iStartCorT = ilStartCorT
                                tlSplitInfo.iEndCorT = ilEndCorT
                                tlSplitInfo.iFirstProjInx = ilFirstProjInx
                                tlSplitInfo.iVefCode = tmClf.iVefCode
                                tlSplitInfo.iNewCode = ilNewCode
                                tlSplitInfo.sPctTrade = slPctTrade
                                tlSplitInfo.sCashAgyComm = slCashAgyComm
                                tlSplitInfo.sTradeAgyComm = tgChfCT.sAgyCTrade        '12-23-02
                                tlSplitInfo.iNTRSlspComm = 0            '2-2-04
                                tlSplitInfo.sNTR = "N"
                                tlSplitInfo.iHardCost = False
                                tlSplitInfo.iMnfNTRItemCode = 0         '11-06-06 not an NTR

                                mParPaySetupPctOfSplit tlSplitInfo, llStdStartDates(), tmCntAllYear(), tmOneVehAllYear()
                             'End If                         'ilfoundoption
                            End If

                            'process the makegoods for this line - there can be multiple vehicles that spots were moved  to.  Each
                            'new vehicle must be retrieved and it's splits determined
                            For ilVehLoop = 0 To imUpperAdjust - 1 Step 1

                                ilFoundOption = False                    'assume everything included
                                If gFilterLists(tmAdjust(ilVehLoop).iVefCode, imInclVefCodes, imUsevefcodes()) Then
                                    ilFoundOption = True
                                End If

                                If tmAdjust(ilVehLoop).iVefCode > 0 And ilFoundOption Then
                                    ilHowMany = 0
                                    ilHowManyDefined = 0
                                    '4-20-00
                                    mParPayGetHowManyFuture tmAdjust(ilVehLoop).iVefCode, ilHowMany, ilHowManyDefined, ilMatchSSCode, llStdStartDates()          '8-4-00 determine the number of owner or slsp splits
                                    For illoop = 0 To ilHowMany - 1 Step 1
                                        
                                        If tmOneVehAllYear(illoop).AllYear.iSSMnfCode = ilMatchSSCode Then
                                            llProcessPct = -1
                                            tmGrf.iCode2 = tmOneVehAllYear(illoop).AllYear.iMnfGroup 'participant
                                        Else                            '4-6-16 handle case where sales source doesnt exist with vehicle
                                            llProcessPct = 0
                                            tmGrf.iCode2 = 0                'no owner (participant)
                                        End If

'                                        If tmOneVehAllYear(ilLoop).AllYear.iSSMnfCode = ilMatchSSCode Then
'                                            llProcessPct = 1000000
'
'                                            tmOnePartAllYear = tmOneVehAllYear(ilLoop).AllYear          '5-11-09 participant percentages for the year
'                                            If tmOnePartAllYear.iMnfGroup <> 0 Then             '4-6-16 handle case where participant not defined with vehicle
'                                                If tmOnePartAllYear.iOwnerByDate <> 1 Then
'                                                    llProcessPct = 0
'                                                End If
'                                            End If
'                                        Else
'                                            llProcessPct = 0
'                                        End If
                                       

                                        If llProcessPct <> 0 Then               '-1 indicates owner where the % must be retrieved by date
                                            ilFoundOne = mFoundSelection(illoop, tmAdjust(ilVehLoop).iVefCode)         '8-4-00 see if selective slsp, owner, vehicle chosen
                                            If ilFoundOne Then
                                                For ilCorT = ilStartCorT To ilEndCorT Step 1
                                                    For ilTemp = 1 To 12
                                                        llTempProject(ilTemp) = tmAdjust(ilVehLoop).lProject(ilTemp)
                                                        llTempAcquisition(ilTemp) = tmAdjust(ilVehLoop).lAcquisitionCost(ilTemp)
                                                    Next ilTemp
                                                    tmGrf.iVefCode = tmAdjust(ilVehLoop).iVefCode
                                                    tmOnePartAllYear = tmOneVehAllYear(illoop).AllYear
                                                
                                                    'If tmOnePartAllYear.iOwnerByDate = 1 Or tmOnePartAllYear.iMnfGroup = 0 Then        '9-27-17
                                                        mParPayLoopOnMonth llTempProject(), llTempAcquisition(), ilCorT, llProcessPct, slPctTrade, slCashAgyComm, slCommPct, smGrossOrNet, ilHowManyDefined, ilMatchSSCode, ilNewCode, 1, igPeriods, tmOnePartAllYear, tgChfCT.sAgyCTrade
                                                    'End If
                                                    

                                                Next ilCorT
                                            End If              'ilfoundOne
                                        End If                  'llprocesspct >0
                                        
                                    Next illoop                 '0 to ilhowmany-1 step 1
                                End If
                            Next ilVehLoop                  '0 To imUpperAdjust - 1
                           
                            For ilScratch = 1 To 13 Step 1    'init projection buckets for next line or contract
                                lmProject(ilScratch) = 0
                                lmAcquisition(ilScratch) = 0
                                lmAcquisitionNet(ilScratch) = 0
                            Next ilScratch
                            'ReDim Preserve tmAdjust(0 To 0) As ADJUSTLIST            'prepare list of mgs
                             ReDim tmAdjust(0 To 0) As ADJUSTLIST             'prepare list of mgs
                            imUpperAdjust = 0
                            'init the acquistion costs
                            For ilScratch = 0 To UBound(llCalSpots) - 1
                                llCalSpots(ilScratch) = 0        'init buckets for daily calendar values
                                llCalAmt(ilScratch) = 0
                                llCalAcqAmt(ilScratch) = 0
                                llCalAcqNetAmt(ilScratch) = 0
                            Next ilScratch
                      
                        'End If                          'if (owner), slsp,  or (vehicle) or end of contract
                        'End If                          'ilFoundoption and <> H
                    Next ilClf                          'next schedule line - for advt & slsp the entire contr is accumulated before writing to GRF
                End If                              'include Air Time vs NTR
            End If                                  'exclude all promotions, merchandising - only include trade if requested
        End If                                      'ilfound
        
    Next ilCurrentRecd
'    Erase tmSofList, tmMnfList, tlSbf
'    Erase tmAdjust, tmSBFAdjust
'    Erase imUsevefcodes
'    Erase imUseCodes
    Exit Function
End Function

'                   mFoundSelection - Determine if a selective slsp,
'                                   owner, vehicle selected
'                  <Input>  ilWhichSelection - index to slsp or owner processing
'                                   (because there can be splits)
'                  Return - true - if selection found
'       11-23-03 Selective slsp not producing correct results when a slsp with 0% entered
Private Function mFoundSelection(ilWhichSelection As Integer, ilVefCode As Integer) As Integer    '8-4-00
    Dim illoop As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    Dim ilFoundOne As Integer
    ilFoundOne = True
    
    tmGrf.iVefCode = tmClf.iVefCode
    
    '7-1-14 move obtaining vehicle only when needed; use binary search for speed up
    If tmVef.iCode <> tmClf.iVefCode Then
        ilRet = gBinarySearchVef(tmClf.iVefCode)
        If ilRet = -1 Then          'not found
            tmVef.iCode = 0
            tmVef.iOwnerMnfCode = 0
        End If
    End If
    tmGrf.iCode2 = tmVef.iOwnerMnfCode       'pick up first veh group
    tmGrf.iSlfCode = tgChfCT.iSlfCode(0)      'use primary one onlyl
    
    mFoundSelection = ilFoundOne
End Function

'                       Billed and Booked subroutine
'                       mParPayLoopOnMonths - Loop and calculate the gross, net &/or commission
'                                   Insert a projected record in GRF or
'                                    a MG record in GRF so that any
'                                    spot moved from one vehicle to another
'                                    will be reflected in the correct vehicle
'                                    if updating as aired
'                       <input> llTempProject - 12 months of projected $ (from
'                                               contract line or MG $
'                               ilCorT - 1 = Cash , 2 = Trade processing, 3 = Hard Costs
'                               llProcessPct - % share of owners or slsp
'                               slPctTrade  - % of trade
'                               slCashAgyComm - agency comm %
'                               slCommPct - slsp % commission (Billed & booked commission report)
'                               slGrossOrNet - G = gross report, N = net report, B = net-net for Vehicle option, T = triple net
'                                               for slsp/vehicle option, or D = net-net for vehicle/participant option
'                               ilHowManyDefined - split transactions (i.e. slsp splits or owners splits)
'                               ilMatchSSCode - Sales Source code
'                               ilNewCode = mnf Revenue area "New" code
'                               ilStartMonth - start month to process   4-4-00, all billed &booked except slsp comm will proces all 12 months at a time
'                                       slsp comm retrieves a possible different  comm percentage for each month
'                                       (if neg, first time thru--initialize buckets)
'                               ilHI - end month to process, after 12th period, insert record in grf
'                               ilCurrMnfGroup - if running partcipant veh group with veh/participant option, use the participant that is associated with the
'                                               split, instead of the first one for the owner
'                               slTradeAgyComm - agy commission for trade for air time or for NTR (NTR could be different than whats defined for airtime)
'                               ilIsItHardCost - true or false
'                               5/7/98 dh - add Show New/Return flag to grf.ipergenl(2)
'
'                               5/10/98 dh - change Billed & booked Commission to keep track of net dollars (vs gross)
'                                            in grf.lDollars(18) so that overage can be computed
'                               4-4-00 dh implement new commission structure where slsp can have different commissions
'                                           based on vehicles by start & end dates
Sub mParPayLoopOnMonth(llTempProject() As Long, llTempAcquisition() As Long, ilCorT As Integer, llProcessPct As Long, slPctTrade As String, slCashAgyComm As String, slCommPct As String, slGrossOrNet As String, ilHowManyDefined As Integer, ilMatchSSCode As Integer, ilNewCode As Integer, ilStartMonth As Integer, ilHi As Integer, tlOnePartAllYear As ONEPARTYEAR, slTradeAgyComm As String)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilMonInx                      ilLoopOnVG                                              *
'******************************************************************************************
    
    Dim ilTemp As Integer
    'Dim ll12MonthTotal As Long     '4-4-00
    Dim slAmount As String
    Dim slSharePct As String
    Dim slStr As String
    Dim slCode As String
    Dim slDollar As String
    Dim slNet As String
    Dim ilRet As Integer
    Dim ilTypes As Integer
    Dim ilmaxTypes As Integer
    Dim ilLo As Integer
    Dim ilOwnerFlag As Integer
    Dim ilIncludeVehicleGroup As Integer
    Dim ilPartPct(0 To 12) As Integer
    Dim ilClear As Integer
    Dim slNetNoSplit As String
    Dim slAmountNoSplit As String

'    ilOwnerFlag = False
'    If llProcessPct < 0 Then        'if negative value, the % is a participant split and the splits
'                                    'comes from the dated participant table
'        ilOwnerFlag = True
'    End If
    ilLo = ilStartMonth
    If ilLo < 0 Or ilLo = 1 Then            'first time, init buckets (from commissions) or from billed and booked
        For ilTemp = 1 To 18 Step 1             'init the years $ buckets for the contract
            tmGrf.lDollars(ilTemp - 1) = 0
            lgNetDollars(ilTemp) = 0
            lgGrsDollars(ilTemp) = 0            '2-26-01
            lgCommDollars(ilTemp) = 0
            lgNetNoSplit(ilTemp) = 0
            lgNetCollectNoSplit(ilTemp) = 0
        Next ilTemp
        dg12MonthTotal = 0                      'total of 12 months buckets
        dg12MonthAcquisition = 0                '9-20-06
        If ilLo < 0 Then
            ilLo = -ilLo                'get back to positive for loop
        End If
    End If

    For ilTemp = ilLo To ilHi               '4-4-00 loop on # buckets to process.  If Slsp billed & booked comm, only processing one bucket at a time
                            'because each month could have a different slsp comm percent
        llProcessPct = tlOnePartAllYear.iPct(ilTemp)

        llProcessPct = llProcessPct * 100
        slAmount = gLongToStrDec(llTempProject(ilTemp), 2)
        slAmountNoSplit = gLongToStrDec(llTempProject(ilTemp), 0)
        slSharePct = gLongToStrDec(llProcessPct, 4)                 'slsp or owner split share in %, else 100.0000%
        slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
        slStr = gRoundStr(slStr, "1", 0)

        If ilCorT = 1 Or ilCorT = 3 Then                 'all cash commissionable, or hard costs
            slCode = gSubStr("100.", slPctTrade)
            slDollar = gDivStr(gMulStr(slStr, slCode), "100")              'slsp gross
            slDollar = gRoundStr(slDollar, "1", 0)
            slNet = gRoundStr(gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "100.00"), ".01", 0)
            slNetNoSplit = gRoundStr(gDivStr(gMulStr(slAmountNoSplit, gSubStr("100.00", slCashAgyComm)), "100.00"), ".01", 0)

            tmGrf.sDateType = "C"
        Else
            If ilCorT = 2 Then                'at least cash is commissionable
                slCode = gIntToStrDec(tgChfCT.iPctTrade, 0)
                slDollar = gDivStr(gMulStr(slStr, slCode), "100")
                slDollar = gRoundStr(slDollar, "1", 0)
                'slNet = slDollar                'assume no commissions on trade
                'If tgChfCT.iAgfCode > 0 And tgChfCT.sAgyCTrade = "Y" Then
                If tgChfCT.iAgfCode > 0 And slTradeAgyComm = "Y" Then

                    slNet = gRoundStr(gDivStr(gMulStr(slDollar, gSubStr("100.00", slCashAgyComm)), "100.00"), "1", 0)
                    slNetNoSplit = gRoundStr(gDivStr(gMulStr(slAmountNoSplit, gSubStr("100.00", slCashAgyComm)), "100.00"), ".01", 0)

                Else
                    slNet = slDollar    'no commission , net is same as gross
                    slNetNoSplit = slAmount     'no split net amount
                End If
                tmGrf.sDateType = "T"
            End If
        End If


            tmGrf.lDollars(ilTemp - 1) = tmGrf.lDollars(ilTemp - 1) + Val(slDollar)
            lgNetDollars(ilTemp) = lgNetDollars(ilTemp) + Val(slNet)
            'lgGrsDollars(ilTemp) = lgGrsDollars(ilTemp) + Val(slDollar)
            lgGrsDollars(ilTemp) = lgGrsDollars(ilTemp) + llTempProject(ilTemp)
            dg12MonthTotal = dg12MonthTotal + Val(slNet)        'accum the years total
            dg12MonthAcquisition = dg12MonthAcquisition + Val(slNet)  '9-20-06 accum total acquisition $
            lgNetNoSplit(ilTemp) = lgNetNoSplit(ilTemp) + Val(slNetNoSplit)
            
'            'need to create extra record to maintain the varying months of different participant %
'            slCommPct = gIntToStrDec(tlOnePartAllYear.iPct(ilTemp), 2)
'            slDollar = slNet                    'put back the net into temp field to calc the net comm $
'            slNet = gRoundStr(gDivStr(gMulStr(slDollar, slCommPct), "100.00"), ".01", 0)   'Mult slsp comm * $, then round
'            lgCommDollars(ilTemp) = lgCommDollars(ilTemp) + Val(slNet)
            lgCommDollars(ilTemp) = lgCommDollars(ilTemp) + Val(slNet)
            
'            '5-11-09 if VEhicle/Gross Net option, when there is an owner change, it cannot show the gross $ for each month of
'            'for each owner.
            If tlOnePartAllYear.iPct(ilTemp) = 0 Then
                lgNetDollars(ilTemp) = 0
                lgGrsDollars(ilTemp) = 0
                lgNetNoSplit(ilTemp) = 0
                lgNetCollectNoSplit(ilTemp) = 0    'net payment collected
            End If
            
        
    Next ilTemp

    'contract complete for cash and or trade values, write out contract
    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for retrieval/removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    'tmGrf.iGenTime(0) = igNowTime(0)        'todays time used for retrieval/removal of records
    'tmGrf.iGenTime(1) = igNowTime(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    'tmGrf.iPerGenl(13) = 0                  'flag to indicate data from contracts (vs budget info for B & B Comparison)
    tmGrf.iPerGenl(12) = 0                  'flag to indicate data from contracts (vs budget info for B & B Comparison)
    tmGrf.sBktType = "C"                    'flag this as a contract recd (vs history, recv. or projection)
    tmGrf.lChfCode = tgChfCT.lCode            'contr internal code
    tmGrf.iAdfCode = tgChfCT.iAdfCode
    'tmGrf.iPerGenl(15) = 0                  '4-11-16 flag to indicate vehicle option with participant vg (different than with other VG options
    tmGrf.iPerGenl(14) = 0                  '4-11-16 flag to indicate vehicle option with participant vg (different than with other VG options
                                            'because this particpant option doesnt split and uses 100% for owner)
    tmGrf.iSofCode = ilMatchSSCode              'sales source
      'If (lg12MonthTotal > 0) Or ((lg12MonthAcquisition <> 0) And (smGrossOrNet = "T")) Or (lg12MonthAcquisition <> 0) Then                      'if all zeros don't write it out

    ilIncludeVehicleGroup = True
    '8-6-10 vehicle w/split slsp subtotals swapped to do slsp w/vehicle subtotals.
    'need to get the vehicle group if application
        gGetVehGrpSets tmGrf.iVefCode, imMinorSet, imMajorSet, tmGrf.iPerGenl(2), tmGrf.iPerGenl(3)   'Genl(3) = minor sort code, genl(4) = major sort code
        'major sort should be Participant (immajorsort)
        If tlOnePartAllYear.iMnfGroup <> 0 Then
            tmGrf.iPerGenl(3) = tlOnePartAllYear.iMnfGroup  '4-5-16 handle case were sales source not defined with vehicle; include so things balance
        Else
            tmGrf.iPerGenl(3) = 0                           '4-14-16 no participant exists for sales source
        End If
        
        If tmGrf.iPerGenl(2) > 0 Or tmGrf.iPerGenl(3) > 0 Then          '4-14-16
            ilIncludeVehicleGroup = mParPayTestVGItem()
        End If

    '4-26-13 test for option to include $0 contracts (currently only b&b has option)
    If (dg12MonthTotal > 0) And (ilIncludeVehicleGroup = True) Then                         'if all zeros don't write it out
        'need to process one month at a timein case percentage change
        For ilTemp = 1 To 12
        'For ilTemp = 0 To 11        '0 index based
            'For ilClear = 1 To 12
            For ilClear = 0 To 11       '0 index based
                tmGrf.lDollars(ilClear) = 0
            Next ilClear
'            tmGrf.lDollars(ilTemp) = lgGrsDollars(ilTemp + 1)   '0 based
            tmGrf.lDollars(ilTemp - 1) = lgGrsDollars(ilTemp)  '1-4-18 0 based
            mParPayAccumPayByContract CAT_GROSS, tmGrf.iVefCode, tmGrf.iPerGenl(3), tlOnePartAllYear.iPct(ilTemp)
'            tmGrf.lDollars(ilTemp) = lgNetDollars(ilTemp + 1)
            tmGrf.lDollars(ilTemp - 1) = lgNetDollars(ilTemp)       '1-4-18 0 based
            mParPayAccumPayByContract CAT_PARTNER, tmGrf.iVefCode, tmGrf.iPerGenl(3), tlOnePartAllYear.iPct(ilTemp)
'            tmGrf.lDollars(ilTemp) = lgNetNoSplit(ilTemp + 1)
            tmGrf.lDollars(ilTemp - 1) = lgNetNoSplit(ilTemp)       '1-4-18 0 based
            mParPayAccumPayByContract CAT_NET, tmGrf.iVefCode, tmGrf.iPerGenl(3), tlOnePartAllYear.iPct(ilTemp)
            '.lDollars(ilTemp) = lgCommDollars(ilTemp + 1)
        Next ilTemp
        
'            ilmaxTypes = 3                  'create both Gross & Net records & owners share
'
'            For ilTypes = 1 To ilmaxTypes             '
'                            '2 = net, 3 = commission.  If not slsp comm, process once for orders on books
'                'if iltype = 1 and ilMaxTypes <> 1 (slsp comm), then the orders on books values are already in tmGRF array  (which should be gross values)
'
'                If ilTypes = 1 And ilmaxTypes <> 1 Then     '2-26-01
'                    For ilTemp = 1 To 18
'                        tmGrf.lDollars(ilTemp - 1) = lgGrsDollars(ilTemp)
'                    Next ilTemp
'                ElseIf ilTypes = 2 Then      'process net
'                    For ilTemp = 1 To 18
'                        tmGrf.lDollars(ilTemp - 1) = lgNetDollars(ilTemp)
'                    Next ilTemp
'                ElseIf ilTypes = 3 Then      'process comm
'                    'processing for commissions
'                    For ilTemp = 1 To 18
'                        tmGrf.lDollars(ilTemp - 1) = lgCommDollars(ilTemp)
'                    Next ilTemp
'                End If
'                For ilTemp = 1 To 12 Step 1
'                    'round the 12 values
'                    slAmount = gLongToStrDec(tmGrf.lDollars(ilTemp - 1), 2)
'                    slAmount = gMulStr("100", slAmount)                       ' gross portion of possible split
'                    tmGrf.lDollars(ilTemp - 1) = Val(slAmount)
'                Next ilTemp
'
'                tmGrf.iPerGenl(4) = ilTypes                     'update for sorting of Slsp comm rept
'                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
'            Next ilTypes

    End If
End Sub

'******************************************************************************
'
'
'                   mParPayGetHowManyFuture - Obtain the number of participant splits
'                   <input>  ilMatchSSCode - for ownership report- find matching
'                                   sales source in the vehicle
'                   <output> ilHowMany - total number of splits there can possibly be
'                                   (for owner always 8, for slsp its the number with
'                                   commission split percent defined, all others its 1)
'                           ilHowManyDefined - total number of splits actually having
'                                   values (for slsp defaults to 1 if not defined and
'                                   he gets 100%; for owner the number is that found
'                                   with the matching sales source; all others its 0)
'                           ilSlfCode() - slsp codes to gather splits for
'                           llSlfCode() - slsp splits percentages for each slsp to process
'
'      Created:  DH 3/97
'      4-20-00 New commission structure by vehicle & slsp & date
'      8-4-00 Option to split participants for vehicle option
'      4-5-01 DH Obtain participants share % for net-net version
'
'******************************************************************************
Sub mParPayGetHowManyFuture(ilVefCode As Integer, ilHowMany As Integer, ilHowManyDefined As Integer, ilMatchSSCode As Integer, llStdStartDates() As Long)    '8-4-00
    Dim illoop As Integer
    Dim ilRet As Integer
    Dim ilFound As Integer

    ilFound = gInitVehAllYearPcts(ilMatchSSCode, ilVefCode, tmCntAllYear(), tmOneVehAllYear()) 'only those participants for the contracts matching SS have
                                     'been built into the tmCntAllYear array.  Each entry is for a different participant of the sales source
    If Not ilFound Then     'no matching vehicle found,must be a mg or outside vehicle that is not on the contract
        gGetOneVehAllYearForMG ilVefCode, llStdStartDates(), ilMatchSSCode, tmPifKey(), tmPifPct(), tmCntAllYear()
        ilFound = gInitVehAllYearPcts(ilMatchSSCode, ilVefCode, tmCntAllYear(), tmOneVehAllYear()) 'only those participants for the contracts matching SS have
        If Not ilFound Then         'force to 100%
            tmOneVehAllYear(0).AllYear.iVefCode = ilVefCode
            tmOneVehAllYear(0).AllYear.iSSMnfCode = ilMatchSSCode
            For illoop = 1 To 12
                tmOneVehAllYear(0).AllYear.iPct(illoop) = 10000
            Next illoop
            tmOneVehAllYear(0).AllYear.iMnfGroup = 0        'unknown participant
        End If
    End If

    'get the vehicle for this transaction
    ilHowManyDefined = UBound(tmOneVehAllYear)
    ilHowMany = UBound(tmOneVehAllYear)
    Exit Sub
End Sub

'           mParPaySetupPctOfSplit - Split the revenue (if required) for the Schedule or SBF.  Create as many
'               GRF records required for splits
'           <input> tlSplitInfo - TYPE array containing  common parameters required when this
'                               code was made into subroutine
'                   llStdStartDates() - array of standard start dates for the year
'                   tlSCF() - salespeople commissions
'                   tlSlfIndex() - array of start/end indices pointing to tlSCF
'                   tlCntAllYear() - array of all participants for all vehicles for 1 contract
Public Sub mParPaySetupPctOfSplit(tlSplitInfo As splitinfo, llStdStartDates() As Long, tlCntAllYear() As ALLPIFPCTYEAR, tlOneVehAllYear() As ALLPIFPCTYEAR)
    Dim ilHowMany As Integer
    Dim ilHowManyDefined As Integer
    Dim illoop As Integer
    Dim ilSlfRecd As Integer
    Dim ilFoundOne As Integer
    Dim ilLoopSlsp As Integer
    Dim slNameCode As String
    Dim ilRet As Integer
    Dim slCode As String
    Dim llProcessPct As Long
    Dim ilTemp As Integer
    'Dim llTempProject(1 To 12) As Long
    'Dim llTempAcquisition(1 To 12) As Long
    Dim llTempProject(0 To 12) As Long  'Index zero ignored
    Dim llTempAcquisition(0 To 12) As Long  'Index zero ignored
    Dim ilCorT As Integer
    Dim ilStartCorT As Integer
    Dim ilEndCorT As Integer
    Dim ilVefCode As Integer
    Dim ilFirstProjInx As Integer
    Dim ilNewCode As Integer
    Dim slPctTrade As String
    Dim slCashAgyComm As String
    Dim ilMatchSSCode As Integer
    Dim slCommPct As String
    Dim slTradeAgyComm As String    '12-23-02
    Dim ilSlspNTRComm As Integer
    Dim slNTR As String * 1
    Dim ilIsItHardCost As Integer
    Dim ilSaveCorT As Integer
    Dim ilMnfItem As Integer
    Dim ilNTRLoop As Integer

    ilMatchSSCode = tlSplitInfo.iMatchSSCode
    ilStartCorT = tlSplitInfo.iStartCorT
    ilEndCorT = tlSplitInfo.iEndCorT
    ilFirstProjInx = tlSplitInfo.iFirstProjInx
    ilVefCode = tlSplitInfo.iVefCode
    ilNewCode = tlSplitInfo.iNewCode
    slPctTrade = tlSplitInfo.sPctTrade
    slCashAgyComm = tlSplitInfo.sCashAgyComm
    slTradeAgyComm = tlSplitInfo.sTradeAgyComm
    ilSlspNTRComm = tlSplitInfo.iNTRSlspComm        '2-2-04
    slNTR = tlSplitInfo.sNTR
    ilIsItHardCost = tlSplitInfo.iHardCost
    ilMnfItem = tlSplitInfo.iMnfNTRItemCode         '11-06-06

    ilHowMany = 0
    ilHowManyDefined = 0

'*********** probably do notneed this        ilMnfSubCo = gGetSubCmpy(tgChfCT, imSlfCode(), lmSlfSplit(), ilVefCode, ilUseSlsComm, lmSlfSplitRev())
    'if owner, it will build array of one vehicles participants in tmOneVehAllYear()
    mParPayGetHowManyFuture ilVefCode, ilHowMany, ilHowManyDefined, ilMatchSSCode, llStdStartDates()      '8-4-00 4-6-00determine the number of owner or slsp splits

    For illoop = 0 To ilHowMany - 1 Step 1
'            If tlOneVehAllYear(ilLoop).AllYear.iSSMnfCode = ilMatchSSCode Then
'                 llProcessPct = 1000000
'                 tmOnePartAllYear = tlOneVehAllYear(ilLoop).AllYear          '5-11-09 participant percentages for the year
'                 If tmOnePartAllYear.iMnfGroup <> 0 Then       '4-6-16 handle case where owner isnt defined with vehicle
'                     '8-2-09, got the matching sales source so if its the first time thru, this is the owner
'                     'If tmOnePartAllYear.iOwnerSeq <> 1 Then
'                     If tmOnePartAllYear.iOwnerByDate <> 1 Then
'                         llProcessPct = 0
'                     End If
'                 End If
'             Else
'                llProcessPct = 0
'            End If

            'only the matching sales sources and vehicle and participants have been build into tmOneVehAllYear()
            'If tmVef.iMnfSSCode(ilLoop + 1) = ilMatchSSCode Then
'               If tlOneVehAllYear(ilLoop + 1).AllYear.iSSMnfCode = ilMatchSSCode Then
            If tlOneVehAllYear(illoop).AllYear.iSSMnfCode = ilMatchSSCode Then
                'llProcessPct = tmVef.iProdPct(ilLoop + 1)            'make it xxx.xxxx
                'llProcessPct = llProcessPct * 100
                'tmGrf.iCode2 = tmVef.iMnfGroup(ilLoop + 1)      '8-4-00 participant
'                    tmGrf.iCode2 = tlOneVehAllYear(ilLoop + 1).AllYear.iMnfGroup 'participant
                 tmGrf.iCode2 = tlOneVehAllYear(illoop).AllYear.iMnfGroup  'participant
                 llProcessPct = -1          'flag to indicate need to process percentages by date (not just 1 percent for all dates applicable)
            Else
                 If tlOneVehAllYear(1).AllYear.iSSMnfCode = 0 Then       'no sales source was found
                    llProcessPct = 1000000                               'take the full amount so it wont be excluded from report
                 Else
                    llProcessPct = 0
                 End If
            End If

        If llProcessPct <> 0 Then        '-1 indicates to process percentages by date (by owner or participant/vehicle option); non 0 indicates a possible slsp split,
                                         '0 indicates to not go thru, nothing to calculate
            'if split slsp comm exist and selective slsp requested, don't show all slsp on this order
            ilFoundOne = mFoundSelection(illoop, ilVefCode)
            If ilFoundOne Then

                
                For ilTemp = 1 To 12
                    llTempProject(ilTemp) = lmProject(ilTemp)
                    llTempAcquisition(ilTemp) = lmAcquisition(ilTemp)
                Next ilTemp
                For ilCorT = ilStartCorT To ilEndCorT Step 1
                    'mBOBNewInsert llTempProject(), ilCorT, llProcessPct, slPctTrade, slCashAgyComm, slCommPct, smGrossOrNet, ilHowManydefined, ilMatchSSCode, ilDateFlag, ilNewCode
                    '4-20-00
                    'tmGrf.iPerGenl(6) = 0       'flag to indicate that at least one vehicle does not have a sub-company defined and
                    tmGrf.iPerGenl(5) = 0       'flag to indicate that at least one vehicle does not have a sub-company defined and
                                'there is at least one slsp in the header with a sub-co (they all must have it, or none at all)
                    'if this vehicle doesnt have a subcompany defined at there is at least one slsp
                    'in the header that has a sub-company defined, the commissions may not be correct.
                    'i.e.  Slsp A  Sub-co A  Slsp Share = 100
                    '      Slsp A  Sub-co B  Slsp Share = 50%
                    '      At least one line doesn't have a sub-company defined, therefore the splits will occur for Slsp A twice.

                    '4-20-05 determine if air time, ntr or hard cost for subsorting  in B & B Recap report
                    'tmGrf.iPerGenl(10) = 0               'assume not an NTR
                    tmGrf.iPerGenl(9) = 0               'assume not an NTR
                    
                    '8-26-09 separate politicals on Sales comparison for contracts, future (vs past)
                     If mCheckAdvPolitical(tgChfCT.iAdfCode) Then      'test for political
                        'tmGrf.iPerGenl(11) = 4                          'political, keep separate from direct, agy, ntr and H/C
                        'tmGrf.iPerGenl(10) = 0                          'NTR subsort doesnt apply
                        tmGrf.iPerGenl(10) = 4                          'political, keep separate from direct, agy, ntr and H/C
                        tmGrf.iPerGenl(9) = 0                          'NTR subsort doesnt apply
                     Else                                               'not political
                        If tgChfCT.iAgfCode = 0 Then
                          'tmGrf.iPerGenl(11) = 0                        'direct
                          'tmGrf.iPerGenl(10) = 0                        'subsort dont apply for directs contracts (not NTR or HC) B & B Recap
                          tmGrf.iPerGenl(10) = 0                        'direct
                          tmGrf.iPerGenl(9) = 0                        'subsort dont apply for directs contracts (not NTR or HC) B & B Recap
                        Else
                          'tmGrf.iPerGenl(11) = 1                        'agency
                          'tmGrf.iPerGenl(10) = 0
                          tmGrf.iPerGenl(10) = 1                        'agency
                          tmGrf.iPerGenl(9) = 0
                        End If
                     End If
                      
                    If slNTR = "Y" Then                  'NTR?
                         'tmGrf.iPerGenl(10) = ilMnfItem      '11-06-06  If B&B Recap, subsort by NTR types
                         tmGrf.iPerGenl(9) = ilMnfItem      '11-06-06  If B&B Recap, subsort by NTR types
                         If ilIsItHardCost Then
                             ''tmGrf.iPerGenl(10) = 3
                             'tmGrf.iPerGenl(11) = 4           '11-06-06 For Billed & Booked Recap, this is to keep Hard Costs separate from Direct/Agy/NTR/Polits
                             tmGrf.iPerGenl(10) = 4           '11-06-06 For Billed & Booked Recap, this is to keep Hard Costs separate from Direct/Agy/NTR/Polits
                         Else
                             ''tmGrf.iPerGenl(10) = 2
                             ''6-17-08 Determine Agency Sales, Direct Sales or just NTR
                             'tmGrf.iPerGenl(11) = 2           '11-06-06 For Billed & Booked Recap, this is to keep NTR separate from Direct/Agy/Polits
                             ''tmGrf.iPerGenl(10) = 2
                             ''6-17-08 Determine Agency Sales, Direct Sales or just NTR
                             tmGrf.iPerGenl(10) = 2           '11-06-06 For Billed & Booked Recap, this is to keep NTR separate from Direct/Agy/Polits

                             'For NTR, determine if its flagged as direct or agy sales.  If so, place with agy or direct and not with NTRs
                             For ilNTRLoop = LBound(tlMMnf) To UBound(tlMMnf) - 1
                                 If (ilMnfItem = tlMMnf(ilNTRLoop).iCode) Then
                                     If (Trim(tlMMnf(ilNTRLoop).sUnitType) = "A") Then     'Agency sales
                                         'tmGrf.iPerGenl(11) = 1
                                         'tmGrf.iPerGenl(10) = 0      'force to no MNF code since it belongs in agency sales
                                         tmGrf.iPerGenl(10) = 1
                                         tmGrf.iPerGenl(9) = 0      'force to no MNF code since it belongs in agency sales
                                     ElseIf (Trim(tlMMnf(ilNTRLoop).sUnitType) = "D") Then 'direct sales
                                         'tmGrf.iPerGenl(11) = 0
                                         'tmGrf.iPerGenl(10) = 0      'force to no MNF code since it belongs in direct sales
                                         tmGrf.iPerGenl(10) = 0
                                         tmGrf.iPerGenl(9) = 0      'force to no MNF code since it belongs in direct sales
                                     Else
                                         'tmGrf.iPerGenl(11) = 2      'normal NTR
                                         tmGrf.iPerGenl(10) = 2      'normal NTR
                                     End If
                                     Exit For
                                 End If
                             Next ilNTRLoop
                             'tmGrf.iPerGenl(11) = 2           '11-06-06 For Billed & Booked Recap, this is to keep NTR separate from Direct/Agy/Polits
                         End If

                     End If

                    
                     ilSaveCorT = ilCorT         'for the billed & booked, force hard cost items to end of report as separte category to cash & trade
                     If ilIsItHardCost = True Then
                         ilCorT = 3
                     End If
                     tmOnePartAllYear = tlOneVehAllYear(illoop).AllYear          'participant percentages for the year
                      
                     '5-11-09  If Vehicle Gross/Net option, need to make sure to get the correct owner % in case the owner has changed
                     '8-2-09 matched on the sales source already, if first time thru this is the owner
                     'If tmOnePartAllYear.iOwnerSeq = 1 Then
                     'If tmOnePartAllYear.iOwnerByDate = 1 Or tmOnePartAllYear.iMnfGroup = 0 Then       '4-6-16 handle case where owner isnt defined with vehicle
                         mParPayLoopOnMonth llTempProject(), llTempAcquisition(), ilCorT, llProcessPct, slPctTrade, slCashAgyComm, slCommPct, smGrossOrNet, ilHowManyDefined, ilMatchSSCode, ilNewCode, 1, igPeriods, tmOnePartAllYear, slTradeAgyComm
                     'End If
                     ilCorT = ilSaveCorT
                Next ilCorT                             'process cash or trade portion
            End If                          'ilFoundOne
        End If                  'llProcessPct
    Next illoop                 'Possibly process mutiple slsp or owners per line.  If advt or vehicle, fall thru after one loop
    Exit Sub
End Sub

'           mParPaySSCodeRVFSelect - Billed & Booked reports: find the sales source and office;
'                     obtain vehicle group if applicable, see if sales source is the selected
'           <input>
'                   tmSofList - array of sls offices
'           <output> ilMatchSSCode - Sales source mnf code
'                    ilMatchSofCode - sales office mnf code
'           <return> true if transaction passes filters
'
Public Function mParPaySSCodeRVFSelect(tmSofList() As SOFLIST, ilMatchSSCode As Integer, ilMatchSOFCode As Integer) As Integer
    Dim ilSearchSlf As Integer
    Dim ilTemp As Integer
    Dim ilFoundOption As Integer
    'Determine Office and Sales Source for later filtering if option by SS/Market
    ilMatchSSCode = 0
    '7-1-14 replace search of internal table to use binarysearch of SLF
    ilSearchSlf = gBinarySearchSlf(tmRvf.iSlfCode)
    If ilSearchSlf = -1 Then                'not found
        ilMatchSOFCode = 0
    Else
        For ilTemp = 0 To UBound(tmSofList) Step 1
            If tmSofList(ilTemp).iSofCode = tgMSlf(ilSearchSlf).iSofCode Then       '7-1-14
                ilMatchSSCode = tmSofList(ilTemp).iMnfSSCode
                ilMatchSOFCode = tmSofList(ilTemp).iSofCode
                Exit For
            End If
        Next ilTemp
    End If
    
    gGetVehGrpSets tmRvf.iAirVefCode, imMinorSet, imMajorSet, tmGrf.iPerGenl(2), tmGrf.iPerGenl(3)   'Genl(3) = minor sort code, genl(4) = major sort code
    ilFoundOption = mParPayRvfSelect(ilMatchSSCode, ilMatchSOFCode)                  '6/14/99 see if this is a valid transaction to process
    mParPaySSCodeRVFSelect = ilFoundOption        'return the results
    Exit Function
End Function

'
'           mParPayNTRTestRVF - Billed and Booked reports:  test NTR/Hard Cost selectivity
'           <input>  tlMnf - array of NTR items
'           <output> ilIsItHardCost : true if hard cost NTR item
'           <Return> true if include transaction
'
Public Function mParPayNTRTestRVF(tlMnf() As MNF, ilIsItHardCost As Integer) As Integer
    Dim ilValidTran As Integer
    ilValidTran = False
    'test for inclusion/exclusion of ntrs, hard cost & air time transactions
    'If (imNTR And tmRvf.iMnfItem > 0) Or (imAirTime And tmRvf.iMnfItem = 0) Or (imHardCost And tmRvf.iMnfItem > 0) Then
    '    ilValidTran = True
    'End If

    ilIsItHardCost = False
    ilIsItHardCost = gIsItHardCost(tmRvf.iMnfItem, tlMnf())
    If tmRvf.iMnfItem > 0 Then          'ntr of some kind
        If ilIsItHardCost Then          'hard cost, Include?
            If imHardCost Then
                ilValidTran = True
            End If
        Else                            'non-hard cost
            If imNTR Then               'include non-hard cost?
                ilValidTran = True
            End If
        End If
    Else                                'air time
        If imAirTime Then               'include air time?
            ilValidTran = True
        End If
    End If
    '2-7-17 ignore all transactions that have been undone (rvfInvoiceUndone = "Y")
    If tmRvf.sInvoiceUndone = "Y" Then
        ilValidTran = False
    End If
    mParPayNTRTestRVF = ilValidTran
    Exit Function
End Function

'       mParPayFormatGrfFromRVF - build fields into GRF prepass record
'       <input>
'               ilIsItHardCost - true if item is a hard cost item
'
Public Sub mParPayFormatGrfFromRVF(ilIsItHardCost As Integer)
    Dim ilNTRLoop As Integer
    'format remainder of record
    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime
    tmGrf.lChfCode = tmChf.lCode           'contr internal code

    'tmGrf.iDateGenl(0, 1) = tmRvf.iTranDate(0)    'date billed or paid
    'tmGrf.iDateGenl(1, 1) = tmRvf.iTranDate(1)
    tmGrf.iDateGenl(0, 0) = tmRvf.iTranDate(0)    'date billed or paid
    tmGrf.iDateGenl(1, 0) = tmRvf.iTranDate(1)

    tmGrf.iVefCode = tmRvf.iAirVefCode
    tmGrf.iAdfCode = tmChf.iAdfCode
    tmGrf.sDateType = tmRvf.sCashTrade          'Type field used for C = Cash, T = Trade
    If tmRvf.sCashTrade = "M" Or tmRvf.sCashTrade = "P" Then        'merch and promotions are for the adjustments, they
                                                                    'wont be categorized on the report by merch or promotions.
        tmGrf.sDateType = "C"
    End If

    ''tmGrf.iPerGenl(10) = 1                      'assume air time
    'tmGrf.iPerGenl(10) = 0               '11-06-06assume not an NTR
    tmGrf.iPerGenl(9) = 0               '11-06-06assume not an NTR
    'Billed & Booked Recap uses the agency vs direct and NTR subsorts
    'Sales Comparison will use political flag to keep separate
    If mCheckAdvPolitical(tmRvf.iAdfCode) Then          'its a political, include this contract?
        'tmGrf.iPerGenl(11) = 4                  'political, keep seperate from direct, agy, ntr and H/C
        'tmGrf.iPerGenl(10) = 0                  'NTR subsort doesnt apply
        tmGrf.iPerGenl(10) = 4                  'political, keep seperate from direct, agy, ntr and H/C
        tmGrf.iPerGenl(9) = 0                  'NTR subsort doesnt apply
    Else            'not political, check for agency, direct, etc.
        If tmRvf.iAgfCode = 0 Then
            'tmGrf.iPerGenl(11) = 0          'direct
            'tmGrf.iPerGenl(10) = 0          'subsorts dont apply for direct contracts (not NTR or HC)
            tmGrf.iPerGenl(10) = 0          'direct
            tmGrf.iPerGenl(9) = 0          'subsorts dont apply for direct contracts (not NTR or HC)
        Else
            'tmGrf.iPerGenl(11) = 1          'agency
            'tmGrf.iPerGenl(10) = 0          'subsorts dont apply for agency contracts (not NTR or HC)
            tmGrf.iPerGenl(10) = 1          'agency
            tmGrf.iPerGenl(9) = 0          'subsorts dont apply for agency contracts (not NTR or HC)
        End If
        'if NTR of some sort, it gets separated from the directs & agency
        If ilIsItHardCost Then                      'if this is a hard cost item, override the cash/trade flag so hard costs fall at end of report
            tmGrf.sDateType = "Z"                   'force hard costs to end fo report
            ''tmGrf.iPerGenl(10) = 3                  'hard cost for flag to subsort
            'tmGrf.iPerGenl(11) = 4                  '11-06-06 flag to indicate hard cost
            'tmGrf.iPerGenl(10) = tmRvf.iMnfItem     'hard cost type for subtotals on B & B Recap
            ''tmGrf.iPerGenl(10) = 3                  'hard cost for flag to subsort
            tmGrf.iPerGenl(10) = 4                  '11-06-06 flag to indicate hard cost
            tmGrf.iPerGenl(9) = tmRvf.iMnfItem     'hard cost type for subtotals on B & B Recap
        Else
            If tmRvf.iMnfItem > 0 Then              'is this an NTR?

                'tmGrf.iPerGenl(11) = 2           '11-06-06 For Billed & Booked Recap, this is to keep NTR separate from Direct/Agy/Polits
                'tmGrf.iPerGenl(10) = tmRvf.iMnfItem
                tmGrf.iPerGenl(10) = 2           '11-06-06 For Billed & Booked Recap, this is to keep NTR separate from Direct/Agy/Polits
                tmGrf.iPerGenl(9) = tmRvf.iMnfItem

                'For NTR, determine if its flagged as direct or agy sales.  If so, place with agy or direct and not with NTRs
                For ilNTRLoop = LBound(tlMMnf) To UBound(tlMMnf) - 1
                    If (tmRvf.iMnfItem = tlMMnf(ilNTRLoop).iCode) Then
                        If (Trim(tlMMnf(ilNTRLoop).sUnitType) = "A") Then     'Agency sales
                            'tmGrf.iPerGenl(11) = 1
                            'tmGrf.iPerGenl(10) = 0          'no NTR item, force to agency sales
                            tmGrf.iPerGenl(10) = 1
                            tmGrf.iPerGenl(9) = 0          'no NTR item, force to agency sales
                        ElseIf (Trim(tlMMnf(ilNTRLoop).sUnitType) = "D") Then 'direct sales
                            'tmGrf.iPerGenl(11) = 0
                            'tmGrf.iPerGenl(10) = 0          'no NTR item, force to direct sales
                            tmGrf.iPerGenl(10) = 0
                            tmGrf.iPerGenl(9) = 0          'no NTR item, force to direct sales
                        Else
                            'tmGrf.iPerGenl(11) = 2      'normal NTR
                            'tmGrf.iPerGenl(10) = tmRvf.iMnfItem
                            tmGrf.iPerGenl(10) = 2      'normal NTR
                            tmGrf.iPerGenl(9) = tmRvf.iMnfItem
                        End If
                        Exit For
                    End If
                Next ilNTRLoop
                'tmGrf.iPerGenl(11) = 2                  'NTR item
                'tmGrf.iPerGenl(10) = tmRvf.iMnfItem
            End If
        End If
    End If
    Exit Sub
End Sub

'           mParPayPastSetupKeyForReport - update the key field for common sorting in Crystal reports
'       <input>
'               ilSlfCode() - array of slsp codes for splits
'               ilMnfGroup() - arry of participant splits
'               ilLoop - index to slsp or participant that is being processed
'       <return>  first index of slsp or participant being processed (i.e. if participants,
'               there may be several sales sources, so the first participant of the sales
'               source may not be 1 the first in the list.
Public Function mParPayPastSetupKeyForReport(ilMnfGroup() As Integer, illoop As Integer) As Integer
    Dim ilRet As Integer
    Dim ilRelativeIndex As Integer
    Dim ilVefCode As Integer
    ilRelativeIndex = illoop                    '2-15-00
    
    ilVefCode = tmRvf.iAirVefCode
    
    'tmGrf.iCode2 = ilMnfGroup(1)             'pick up first veh group
    tmGrf.iCode2 = ilMnfGroup(illoop + 1)
    
    mParPayPastSetupKeyForReport = ilRelativeIndex
    Exit Function
End Function

'           mParPayConvertAndSplitPast - convert Receivables packed decimal to string for
'           math computations
'           <input>  llProcessPct - revenue share %
'                    slCommPct - comm. slsp % share
'                    llNetDollars - running total of net $ split
'                    ilRelativeIndex -
'                    ilHowManyDefined - # of splits defined
'                    ilReverseFlag - true to negate and subtract $
'           <oupput>
'                    llNetDollars - running total of net $ split
'                    llcommDollars - calc. $ for slsp comm
Public Sub mParPayConvertAndSplitPast(llProcessPct As Long, slCommPct As String, llTransNet As Long, llNetDollar As Long, llCommDollar As Long, ilRelativeIndex As Integer, ilHowManyDefined As Integer, ilReverseFlag As Integer)
    Dim slPct As String
    Dim slAmount As String
    Dim slAcquisitionCost As String
    Dim slDollar As String
    Dim llAcquisitionCost As Long
    Dim llAcqNet As Long
    Dim llAcqComm As Long

    slPct = gLongToStrDec(llProcessPct, 4)           'slsp split share in % or Owner pct.  If advt or vehicle
                                                                                    'options, slsp is force to100%
'            slAmount = mParPayWhichAmtToStr(tmRvf.sGross, 2)     'determine whether to use the net value from rvf/phf or acquisition cost
'            slDollar = gMulStr(slPct, slAmount)                 'slsp gross portion of possible split
'            llGrossDollar = Val(gRoundStr(slDollar, "01.", 0))      '10-6-17 Gross on report isnt the split, it total sales from order
'            llTransGross = llTransGross - llGrossDollar          'take away the split amount from the total to see whats remaining.
'            If ilRelativeIndex = ilHowManyDefined - 1 Then       'last slsp or participant processed?  Handle
'                                                        'extra pennies.
'                llGrossDollar = llGrossDollar + llTransGross    'last recd gets left over pennies
'            End If

    slAmount = mParPayWhichAmtToStr(tmRvf.sNet, 2)     'determine whether to use the net value from rvf/phf or acquisition cost
    slDollar = gMulStr(slPct, slAmount)                 'slsp net portion of possible split
    llNetDollar = Val(gRoundStr(slDollar, "01.", 0))
    llTransNet = llTransNet - llNetDollar

    slDollar = gMulStr(slCommPct, slAmount)
    llCommDollar = Val(gRoundStr(slDollar, "01.", 0))

    If ilReverseFlag Then
        'tmGrf.lDollars(18) = tmGrf.lDollars(18) - llNetDollar
        tmGrf.lDollars(17) = tmGrf.lDollars(17) - llNetDollar
    Else
        'tmGrf.lDollars(18) = tmGrf.lDollars(18) + llNetDollar     'accum total years net
        tmGrf.lDollars(17) = tmGrf.lDollars(17) + llNetDollar     'accum total years net
    End If
    Exit Sub
End Sub

Public Function mParPayReverseSign(llTransNet As Long, llTransGross As Long) As Integer
    Dim ilReverseFlag As Integer
    Dim slAmount As String

    ilReverseFlag = False
    
    If llTransNet < 0 Then          'billed & booked adjustment, subtract it out
        
        ilReverseFlag = True        'reverse gross & net signs to always work with positive values
        llTransGross = -llTransGross
        slAmount = gLongToStrDec(llTransGross, 2)
        'gStrToPDN slAmount, 2, 6, tmRvf.sGross
        tmRvf.sGross = mParPayWhichAmtToPDN(slAmount, 2, 6)
    
        llTransNet = -llTransNet
        slAmount = gLongToStrDec(llTransNet, 2)
        'gStrToPDN slAmount, 2, 6, tmRvf.sNet
        tmRvf.sNet = mParPayWhichAmtToPDN(slAmount, 2, 6)
    End If
    
    mParPayReverseSign = ilReverseFlag
    Exit Function
End Function

'           Billed & Booked - determine whether to use gross or net from RVF (PHF),
'           or use the Acquisition cost
'           <input> slAmt - gross or net value from RVF/PHF
'           <return> gross or net value from RVF/PHF or the acquistion cost as Long
Public Function mParPayWhichAmtToLong(slAmt As String) As Long
    Dim llAmt As Long
    gPDNToLong slAmt, llAmt
    mParPayWhichAmtToLong = llAmt
End Function

'           Billed & Booked - determine whether to use gross or net from RVF (PHF),
'           or use the Acquisition cost
'           <input> slAmt - Gross or net from rvf/phf
'           <return> gross or net value from RVF/PHF or the acquistion cost as a String
Public Function mParPayWhichAmtToStr(slAmt As String, ilDecPlaces As Integer) As String
    Dim slAmount As String
        
    gPDNToStr slAmt, ilDecPlaces, slAmount
    mParPayWhichAmtToStr = slAmount
End Function

'           Billed & Booked - determine whether to use gross or net from RVF (PHF),
'           or use the Acquisition cost
'           <input> slAmt - gross or net value from RVF/PHF
'           <return> gross or net value from RVF/PHF or the acquistion cost as a String
Public Function mParPayWhichAmtToPDN(slAmt As String, ilDecPlaces As Integer, ilPositions As Integer) As String
    Dim slAmount As String
    gStrToPDN slAmt, ilDecPlaces, ilPositions, slAmount
    mParPayWhichAmtToPDN = slAmount
End Function

'               mGetParticipantSplits - build the table of participant splits
'               for the Participant Payables
'           <input> llstdStartDate - effective start date to gather from PIF
'                   blOwnerOnly - when building the table of participants, if by Vehicle Group Participants for Vehicle option,
'                   force to 100% to for owners share; optional parameters and default to false if not present
'           <output> tmPifKey() - array of vehicles and indices into tmPifPct array
'                   tmPifPct() - array of participant percentages by vehicle
Public Sub gParPayGetAllParticipantSplits(llStdStartDate As Long, Optional blOwnerOnly As Boolean = False)
    gCreatePIFForRpts llStdStartDate, tmPifKey(), tmPifPct(), RptSelParPay, blOwnerOnly
End Sub

'           Test if vehicle groups used;  If it is, test to see if the
'           item selected matches the vehicle processing
'           return - true to process record
Public Function mParPayTestVGItem() As Integer
    Dim ilIncludeVehicleGroup As Integer
    Dim ilLoopOnVG As Integer
    Dim slStr As String
    Dim ilRet As Integer
    Dim ilVGListBox As Integer

    ilIncludeVehicleGroup = True
    If RptSelParPay!lbcSelection(1).SelCount > 0 Then
        ilIncludeVehicleGroup = False
        For ilLoopOnVG = 0 To RptSelParPay!lbcSelection(1).ListCount - 1 Step 1
            If RptSelParPay!lbcSelection(1).Selected(ilLoopOnVG) Then
                slStr = tgMnfCodeCT(ilLoopOnVG).sKey
                ilRet = gParseItem(slStr, 2, "\", slStr)
                'Determine which vehicle set to test and whether major or minor sort
                If imMajorSet > 0 Then
                    If tmGrf.iPerGenl(3) = Val(slStr) Then
                        ilIncludeVehicleGroup = True
                        Exit For
                    End If
                
                End If
            End If
        Next ilLoopOnVG
    End If
    mParPayTestVGItem = ilIncludeVehicleGroup
End Function

'
'           gGenParticipantPayables - create a report generated from Receivables, History, & Business on the Books
'           to provide Participant Payables based on collections or billing
'
Public Sub gGenParticipantPayables()
    Dim llPacingDate As Long
    Dim ilTemp As Integer
    Dim blOwnerOnly As Boolean
    Dim llLastBilled As Long
    Dim ilLastBilledInx As Integer
    Dim ilRet As Integer
    Dim slStr As String
    Dim ilLoopOnPHFRVF As Integer
    Dim blNewContract As Boolean
    ReDim llStdStartDates(0 To 13) As Long  'Index zero ignored

    ilTemp = igMonthOrQtr               '7-3-08 convert to start month vs start qtr
    llPacingDate = 0
    Screen.MousePointer = vbHourglass
    If gParPayOpenFiles() = 0 Then
        gSetupBOBDates 2, llStdStartDates(), llLastBilled, ilLastBilledInx, llPacingDate, ilTemp  'build array of std start & end dates
        mUniqueInfoGet Format(llStdStartDates(1), "ddddd"), Format(llStdStartDates(13) - 1, "ddddd")      'gather the unique contract #s from RVF, PHF, & CHF, return sorted by contract # with flags if contract exists in PHF, RVF or CHF
        gParPayGetAllParticipantSplits llStdStartDates(1)
        mBuildCommonParameters      'build tables such as slsp, office, etc
        'loop on contracts in contract order  , process 1 contract at a time
        While Not UniqueInfo_rst.EOF
            lmSingleCntr = UniqueInfo_rst!ContractNo
            '3-20-18 trouble sending contract # 0 because the routine that gets rvf/phf by contract # assumes 0 means get ALL, not any selective contract
            'Change the routine gObtainPhfRvfbyCntr to test for -1 and use 0 as a matching contract test, vs retrieve ALL
            If lmSingleCntr = 0 Then
                lmSingleCntr = -1
            End If
            If (lmSingleUserCntr = 0) Or ((lmSingleUserCntr > 0) And (lmSingleUserCntr = lmSingleCntr)) Then
                blNewContract = True
                ReDim tmPayByContract(0 To 0) As PAYBY_BREAKOUT

                'bypass past if projection only if dates all in future
                If (llStdStartDates(1) > llLastBilled) Then
                    If UniqueInfo_rst!chfFileFlag Then
                        ilRet = mParPayBuildProj(llStdStartDates(), llLastBilled, igPeriods, blNewContract)
                        mParPayWriteCntToGRF TOTAL_BYCNT, tmPayByContract         'only future
                        
                    End If
                Else
                    'for past, process history and receivables in one pass
                    If (UniqueInfo_rst!phfFileFlag) Or (UniqueInfo_rst!rvfFileFlag) Then
                        gCRParPay_Past llStdStartDates(), llLastBilled, ilLastBilledInx, igPeriods, blNewContract
                    End If
                    If llLastBilled + 1 < llStdStartDates(igPeriods + 1) Then              'past only or past & projection
                        If UniqueInfo_rst!chfFileFlag Then      'process the contracts only if it was found to exist
                            If (UniqueInfo_rst!phfFileFlag) Or (UniqueInfo_rst!rvfFileFlag) Then        'did phf or rvf exist?
                                blNewContract = False                       'if so, not first time for this contract
                            Else
                                blNewContract = True
                            End If

                            ilRet = mParPayBuildProj(llStdStartDates(), llLastBilled, igPeriods, blNewContract)
                        End If
                    End If
                    mParPayWriteCntToGRF TOTAL_BYCNT, tmPayByContract   'past and/or future
                    
                End If
            End If          'lmsingleusercntr = 0
            UniqueInfo_rst.MoveNext
        Wend
        
        mParPayWriteToGRF TOTAL_BYVEHICLE, tmPayByVehicle
        mParPayWriteToGRF TOTAL_BYPARTNER, tmPayByPartner
        mParPayWriteToGRF TOTAL_BYFINAL, tmPayByFinal
        gParPayCloseFiles
        mUniqueInfoClose

        Erase llStdStartDates
        Erase tmPayByContract, tmPayByVehicle, tmPayByPartner
        Screen.MousePointer = vbDefault
        Dim llPop As Long
        Dim llTime As Long
        ReDim ilNowTime(0 To 1) As Integer
        slStr = Format$(gNow(), "h:mm:ssAM/PM")       'end time of run
        gPackTime slStr, ilNowTime(0), ilNowTime(1)
        gUnpackTimeLong ilNowTime(0), ilNowTime(1), False, llTime
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, llPop   'start time of run
        llPop = llPop - llTime              'time in seconds in runtime
    End If
End Sub

'           Clear out the arrays in case re-entrant
'
Private Sub mUniqueInfoClose()
    On Error Resume Next
    If Not UniqueInfo_rst Is Nothing Then
        If (UniqueInfo_rst.State And adStateOpen) <> 0 Then
            UniqueInfo_rst.Close
        End If
        Set UniqueInfo_rst = Nothing
    End If
    Exit Sub
End Sub

'
'           gather the unique contract numbers in RVF and set the flag that it exists in RVF
'
Private Function mUniqueInfoPut(llContrNo As Long, llContrCode As Long, blPhfFileFlag As Boolean, blRvfFileFlag As Boolean, blChfFileFlag As Boolean) As Integer
    mUniqueInfoPut = False
    
    UniqueInfo_rst.Filter = "ContractNo = " & llContrNo
    If UniqueInfo_rst.EOF Then
        'Contract # does not exist, add to table
        UniqueInfo_rst.AddNew Array("ContractNo", "ContractCode", "phfFileFlag", "rvfFileFlag", "chfFileFlag"), Array(llContrNo, llContrCode, blPhfFileFlag, blRvfFileFlag, blChfFileFlag)
        mUniqueInfoPut = True
        imTempCount = imTempCount + 1
    Else
        If blPhfFileFlag Then
            UniqueInfo_rst!phfFileFlag = True
            mUniqueInfoPut = True
        ElseIf blRvfFileFlag Then
            UniqueInfo_rst!rvfFileFlag = True
            mUniqueInfoPut = True
        Else
            UniqueInfo_rst!chfFileFlag = True
            mUniqueInfoPut = True
        End If
    End If
End Function

'
'           Define the record set fields for extraction and sort
'
Private Function mUniqueInfoInit() As ADODB.Recordset
    Dim rst As ADODB.Recordset
        
    Set rst = New ADODB.Recordset
    With rst.Fields
        .Append "ContractNo", adInteger
        .Append "ContractCode", adInteger
        .Append "phfFileFlag", adBoolean
        .Append "rvfFileFlag", adBoolean
        .Append "chfFileFlag", adBoolean
    End With
    rst.Open
    rst!ContractNo.Properties("optimize") = True
    rst.Sort = "ContractNo"
    Set mUniqueInfoInit = rst
End Function

'           read History, Receivables, and Contract files effective the earliest date entered,
'           Build ADO disconnected arry of the unique contracts
'           <input>  slEArliestDate = start date of requested report
Public Sub mUniqueInfoGet(slEarliestDate As String, slLatestDate As String)
    Dim rst_Temp As ADODB.Recordset
    Dim slQuery As String
    Dim slDate As String
    Dim llTranDate As Long
    Dim ilTranDate(0 To 1) As Integer
    Dim ilRet As Integer
    Dim blOk As Boolean
    Dim llCntrNo As Long
    Dim slStr As String
    Dim ilTemp As Integer
    Dim slSingleCntr As String

    imTempCount = 0
    mUniqueInfoClose        'ensure all tables cleared in case re-entrant
    Set UniqueInfo_rst = mUniqueInfoInit()         'define the fields to create
    'process the History file
    If lmSingleUserCntr = 0 Then
        slSingleCntr = ""
    Else
        slSingleCntr = " and phfCntrno = " & lmSingleUserCntr
    End If
    slQuery = "Select phfcntrno, phftrandate, phftrantype, phfagePeriod, phfAgeYear from PHF_Payment_History where phfTranDate >= '" & Format(slEarliestDate, sgSQLDateForm) & "'" & slSingleCntr
    Set rst_Temp = gSQLSelectCall(slQuery)
    While Not rst_Temp.EOF
        llCntrNo = rst_Temp!phfcntrno
        slDate = Format$(rst_Temp!phfTranDate, "mm/dd/yyyy")
        llTranDate = gDateValue(slDate)
        If (rst_Temp!phfTranType = "IN" Or rst_Temp!phfTranType = "AN" Or rst_Temp!phfTranType = "HI") And (llTranDate <= gDateValue(slLatestDate)) Then       'billing transactions within the requested dates
            ilRet = mUniqueInfoPut(llCntrNo, 0, True, False, False)
        Else
            If (rst_Temp!phfTranType = "PI" Or Left$(rst_Temp!phfTranType, 1) = "W") Then           'collections, test on ageing month/year
                slDate = str$(rst_Temp!phfAgePeriod)
                slDate = slDate & Trim$("/15/")
                slDate = slDate & str$(rst_Temp!phfAgeYear)
                'get end date of the std bdcst for this ageing month & year
                slStr = gObtainStartStd(Trim$(slDate))   '4-20-00 chged to use start of bdcst date instead of end of bdcst month
                If gDateValue(slStr) <= gDateValue(slLatestDate) Then
                    ilRet = mUniqueInfoPut(llCntrNo, 0, True, False, False)
                End If
            End If
            
        End If
        rst_Temp.MoveNext
    Wend
    rst_Temp.Close
    
    'process the Receivables file
     If lmSingleUserCntr = 0 Then
        slSingleCntr = ""
    Else
        slSingleCntr = " and rvfCntrno = " & lmSingleUserCntr
    End If
    slQuery = "Select rvfcntrno, rvftrandate, rvftrantype, rvfagePeriod, rvfageyear from RVF_Receivables where rvfTranDate >= '" & Format(slEarliestDate, sgSQLDateForm) & "'" & slSingleCntr
    Set rst_Temp = gSQLSelectCall(slQuery)
    While Not rst_Temp.EOF
        llCntrNo = rst_Temp!rvfCntrno
        slDate = Format$(rst_Temp!rvfTranDate, "mm/dd/yyyy")
        llTranDate = gDateValue(slDate)
        If (rst_Temp!rvfTranType = "IN" Or rst_Temp!rvfTranType = "AN") And (llTranDate <= gDateValue(slLatestDate)) Then       'billing transactions within the requested dates
            ilRet = mUniqueInfoPut(llCntrNo, 0, False, True, False)
        Else
            If (rst_Temp!rvfTranType = "PI" Or Left$(rst_Temp!rvfTranType, 1) = "W") Then           'collections, test on ageing month/year
                slDate = str$(rst_Temp!rvfAgePeriod)
                slDate = slDate & Trim$("/15/")
                slDate = slDate & str$(rst_Temp!rvfAgeYear)
              
                'get end date of the std bdcst for this ageing month & year
                slStr = gObtainStartStd(Trim$(slDate))   '4-20-00 chged to use start of bdcst date instead of end of bdcst month
                If gDateValue(slStr) <= gDateValue(slLatestDate) Then
                    ilRet = mUniqueInfoPut(llCntrNo, 0, False, True, False)
                End If
            End If
            
        End If
        rst_Temp.MoveNext
    Wend
    rst_Temp.Close
    
    slDate = gAdjYear(Format$(slEarliestDate, sgShowDateForm))      'convert earliest std start to sql format
    slDate = Format$(slDate, sgSQLDateForm)
    'filter contracts based on active HOGN, and dates active effective std start
    If lmSingleUserCntr = 0 Then
        slSingleCntr = ""
    Else
        slSingleCntr = " and chfCntrno = " & lmSingleUserCntr
    End If
    slQuery = "Select chfcntrno, chfstartdate, chfenddate, chftype,chfstatus ,chfadfcode, chfPctTrade from chf_contract_header where chfdelete = 'N' and chfStatus in ('H','O','G','N') and chfEndDate >= " & "'" & slDate & "'" & slSingleCntr
    Set rst_Temp = gSQLSelectCall(slQuery)
    While Not rst_Temp.EOF
        llCntrNo = rst_Temp!chfCntrno
        If lmSingleUserCntr = rst_Temp!chfCntrno Or lmSingleUserCntr = 0 Then       'test for single contract selection
            'test selectivity on contract types
             'C=Standard; V=Reservation; T=Remnant; R=Direct Response; Q=Per inQuiry; S=PSA; M=Promo
            If (rst_Temp!chftype = "C" And imStandard) Or (rst_Temp!chftype = "V" And imReserv) Or (rst_Temp!chftype = "T" And imRemnant) Or (rst_Temp!chftype = "R" And imDR) Or (rst_Temp!chftype = "Q" And imPI) Or (rst_Temp!chftype = "S" And imPSA) Or (rst_Temp!chftype = "M" And imPromo) Then
                blOk = True
                'valid contract type
                'test for political or not
                'if the record is valid for inclusion/exclusion of politicals and non-politicals
                ilTemp = rst_Temp!chfAdfCode
                If gIsItPolitical(ilTemp) Then          'its a political, include this contract?
                     If Not imInclPolit Then
                        blOk = False
                    End If
                Else                                                'not a political advt, include this contract?
                     If Not imInclNonPolit Then
                        blOk = False
                    End If
                End If
                If blOk Then
                    blOk = False
                    'if excluding trades, only exclude if 100%
                    If (rst_Temp!chfPctTrade = 100 And imTrade) Or (rst_Temp!chfPctTrade < 100 And imAirTime) Then
                        blOk = True
                    End If
                End If
                If blOk Then
                    'valid contract type, political/non poltical, and trade options
                    'create entry in array to process
                    llCntrNo = rst_Temp!chfCntrno
                    ilRet = mUniqueInfoPut(llCntrNo, 0, False, False, True)
                End If
            End If
        End If
        rst_Temp.MoveNext
    Wend
    rst_Temp.Close
    UniqueInfo_rst.Filter = adFilterNone            'Insure last filter is cleared
    Exit Sub
End Sub

Private Sub mBuildCommonParameters()
    Dim ilTemp As Integer
    Dim ilRet As Integer

    'build array of selling office codes and their sales sources.
    'need to get the sales source from contracts slsp in order to find the correct participant entry
    ilTemp = 0
    ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tmSofList(0 To ilTemp) As SOFLIST
        tmSofList(ilTemp).iSofCode = tmSof.iCode
        tmSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
        ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
    Loop
'        ilRet = gObtainSlf(RptSelParPay, hmSlf, tmSlfList())

    'build array of vehicles to include or exclude
    gObtainCodesForMultipleLists 0, tgVehicle(), imInclVefCodes, imUsevefcodes(), RptSelParPay
    gAddDormantVefToExclList imInclVefCodes, tgMVef(), imUsevefcodes()          '8-4-17 if excluding vehicles, make sure dormant ones exluded since
                                                                                'they wont be in original list
    igPeriods = 12

    tmTranTypes.iAdj = True              'look only for adjustments in the History & Rec files
    tmTranTypes.iInv = True
    tmTranTypes.iWriteOff = True
    tmTranTypes.iPymt = True
    tmTranTypes.iCash = True
    tmTranTypes.iTrade = False
    tmTranTypes.iMerch = False
    tmTranTypes.iPromo = False
    
    tmTranTypes.iNTR = False         '9-17-02
    If imNTR Or imHardCost Then                   '4-25-05 need to gather the NTR for adjustments too
        tmTranTypes.iNTR = True
    End If
    
    If imTrade Then
        tmTranTypes.iTrade = True
    End If
    
    tmSBFType.iNTR = True           'this applies to Billed & Booked only to retrieve future $
    tmSBFType.iInstallment = False
    tmSBFType.iImport = False
   
    igPeriods = 12                  'always 1 year of data
    imMajorSet = 1                  'default vehicle group set to Participants
    imMinorSet = 0
    
    ilRet = btrGetFirst(hmMnf, tmMnf, imMnfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    
    Do While ilRet = BTRV_ERR_NONE
        If tmMnf.sType = "M" Then
            ReDim Preserve tmMnfList(0 To ilTemp) As MNFLIST
            tmMnfList(ilTemp).iMnfCode = tmMnf.iCode
            tmMnfList(ilTemp).iBillMissMG = tmMnf.iGroupNo
            ilTemp = ilTemp + 1
        End If
        ilRet = btrGetNext(hmMnf, tmMnf, imMnfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    
    ReDim tmPayByContract(0 To 0) As PAYBY_BREAKOUT
    ReDim tmPayByVehicle(0 To 0) As PAYBY_BREAKOUT
    ReDim tmPayByPartner(0 To 0) As PAYBY_BREAKOUT
    ReDim tmPayByFinal(0 To 1) As PAYBY_BREAKOUT
End Sub

'
'           Accumulate 1 contracts gross, net collected, partner share, owed $
'           <input> ilCategory  - index into the bucket to accumulate
'                   ilvefcode - vehicle code
'                   ilMnfCode - participant code
Private Sub mParPayAccumPayByContract(ilCategory As Integer, ilVefCode As Integer, ilMnfCode As Integer, ilMnfPct As Integer)
    Dim ilLoopOnMonth As Integer
    Dim blVehicleExists As Boolean

     blVehicleExists = False
     For ilLoopOnMonth = LBound(tmPayByContract) To UBound(tmPayByContract) - 1
         If tmPayByContract(ilLoopOnMonth).iVefCode = ilVefCode And tmPayByContract(ilLoopOnMonth).iMnfPartCode = ilMnfCode And tmPayByContract(ilLoopOnMonth).iPartSharePct = ilMnfPct Then
             mParPayRollTotals ilCategory, TOTAL_BYCNT, ilMnfCode, tmGrf.lDollars(), ilLoopOnMonth
               blVehicleExists = True
               Exit For
         End If
     Next ilLoopOnMonth
     If Not blVehicleExists Then
         tmPayByContract(ilLoopOnMonth).iVefCode = ilVefCode
         tmPayByContract(ilLoopOnMonth).iMnfPartCode = ilMnfCode
         tmPayByContract(ilLoopOnMonth).iPartSharePct = ilMnfPct
         tmPayByContract(ilLoopOnMonth).lCntrNo = lmSingleCntr
         tmPayByContract(ilLoopOnMonth).lCode = tmChf.lCode
         tmPayByContract(ilLoopOnMonth).iAdfCode = tmChf.iAdfCode
         tmPayByContract(ilLoopOnMonth).sProductName = tmChf.sProduct
         mParPayRollTotals ilCategory, TOTAL_BYCNT, ilMnfCode, tmGrf.lDollars(), ilLoopOnMonth
         ReDim Preserve tmPayByContract(0 To UBound(tmPayByContract) + 1) As PAYBY_BREAKOUT
    End If
     'Roll over totals for vehicle,
     blVehicleExists = False
     For ilLoopOnMonth = LBound(tmPayByVehicle) To UBound(tmPayByVehicle) - 1
         If tmPayByVehicle(ilLoopOnMonth).iVefCode = ilVefCode And tmPayByVehicle(ilLoopOnMonth).iMnfPartCode = ilMnfCode And tmPayByVehicle(ilLoopOnMonth).iPartSharePct = ilMnfPct Then
             mParPayRollTotals ilCategory, TOTAL_BYVEHICLE, ilMnfCode, tmGrf.lDollars(), ilLoopOnMonth
             blVehicleExists = True
             Exit For
         End If
     Next ilLoopOnMonth
     If Not blVehicleExists Then
         tmPayByVehicle(ilLoopOnMonth).iVefCode = ilVefCode
         tmPayByVehicle(ilLoopOnMonth).iMnfPartCode = ilMnfCode
         tmPayByVehicle(ilLoopOnMonth).iPartSharePct = ilMnfPct
         tmPayByVehicle(ilLoopOnMonth).iAdfCode = 0
         tmPayByVehicle(ilLoopOnMonth).lCntrNo = 0
         tmPayByVehicle(ilLoopOnMonth).lCode = 0
         mParPayRollTotals ilCategory, TOTAL_BYVEHICLE, ilMnfCode, tmGrf.lDollars(), ilLoopOnMonth
         ReDim Preserve tmPayByVehicle(0 To UBound(tmPayByVehicle) + 1) As PAYBY_BREAKOUT
     End If

     'Roll over totals for participant, do not test the share % could be different across the year
     blVehicleExists = False
     For ilLoopOnMonth = LBound(tmPayByPartner) To UBound(tmPayByPartner) - 1
         If tmPayByPartner(ilLoopOnMonth).iMnfPartCode = ilMnfCode Then
             mParPayRollTotals ilCategory, TOTAL_BYPARTNER, ilMnfCode, tmGrf.lDollars(), ilLoopOnMonth
             blVehicleExists = True
             Exit For
         End If
     Next ilLoopOnMonth
     If Not blVehicleExists Then
         tmPayByPartner(ilLoopOnMonth).iVefCode = 0  '32767          'no vehicle with totals by participant
         tmPayByPartner(ilLoopOnMonth).iMnfPartCode = ilMnfCode
         tmPayByPartner(ilLoopOnMonth).iPartSharePct = 0
         tmPayByPartner(ilLoopOnMonth).lCntrNo = 0   '999999999
         mParPayRollTotals ilCategory, TOTAL_BYPARTNER, ilMnfCode, tmGrf.lDollars(), ilLoopOnMonth
         ReDim Preserve tmPayByPartner(0 To UBound(tmPayByPartner) + 1) As PAYBY_BREAKOUT
     End If
     
     'Roll over totals for final totals,
     tmPayByFinal(0).iVefCode = 0    '32767
     tmPayByFinal(0).iMnfPartCode = 0    '32767
     tmPayByFinal(0).iPartSharePct = 0
     tmPayByFinal(0).iMnfPartCode = 0    '32767
     tmPayByFinal(0).lCntrNo = 0 '999999999
     mParPayRollTotals ilCategory, TOTAL_BYFINAL, ilMnfCode, tmGrf.lDollars(), 0

     Exit Sub
End Sub

'
'           Roll totals into the designated subtotal buckets
'           mParPayRollTotals
'           <input> ilCategory - which buckets to accumulate (gross, net, etc)
'                   1 = gross, 2 = net, 3 = collected, 4 =Outstanding, 5 = Partner Net, 6 = Owed
'                   ilTotalType - which subtotals to roll $ into: 1 = contract, 2 = vehicle, 3 = participant vehicle group, 4 = final
'                   ilMnfCode -participant code
'                   llMonthSource - monthly buckets to roll over
'                   ilTotalIndex - index into the group of vehicles/participants
Private Sub mParPayRollTotals(ilCategory As Integer, ilTotalType As Integer, ilMnfCode As Integer, llMonthSource() As Long, ilTotalIndex As Integer)
    Dim ilLoopOnMonth As Integer
    Dim llPartnerCalc As Long
    Dim slSharePct As String
    Dim slAmount As String
    Dim slStr As String
    Dim llProcessPct As Long

    If ilTotalType = TOTAL_BYCNT Then         'contract totals
        If ilCategory = CAT_GROSS Then      'gross
            For ilLoopOnMonth = 0 To 12
                tmPayByContract(ilTotalIndex).lGross(ilLoopOnMonth) = tmPayByContract(ilTotalIndex).lGross(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_NET Then  'net
            For ilLoopOnMonth = 0 To 12
                tmPayByContract(ilTotalIndex).lNet(ilLoopOnMonth) = tmPayByContract(ilTotalIndex).lNet(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_COLLECTED Then  'collected
            For ilLoopOnMonth = 0 To 12
                tmPayByContract(ilTotalIndex).lCollected(ilLoopOnMonth) = tmPayByContract(ilTotalIndex).lCollected(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_OUTSTANDING Then  'outstanding
            For ilLoopOnMonth = 0 To 12
                tmPayByContract(ilTotalIndex).lOutstanding(ilLoopOnMonth) = tmPayByContract(ilTotalIndex).lOutstanding(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_PARTNER Then  'partner net
            For ilLoopOnMonth = 0 To 12
                tmPayByContract(ilTotalIndex).lPartnersWorth(ilLoopOnMonth) = tmPayByContract(ilTotalIndex).lPartnersWorth(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
                'partner net has been changed to Partner collected (collected * partner %)
                'first save the partners full share, which is needed to calc the partner outstanding (partner pct * partner net share)
                
'9-24-17                tmPayByContract(ilTotalIndex).lPartnersWorth(ilLoopOnMonth) = tmPayByContract(ilTotalIndex).lPartnersWorth(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
                'calc partner collected when writing out the prepass; may not have all the collections
                'no need to update these buckets
                'tmPayByContract(ilTotalIndex).lPartner(ilLoopOnMonth) = tmPayByContract(ilTotalIndex).lPartner(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)

            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_PARTNER_COLLECT Then            'participants share of collections
                For ilLoopOnMonth = 0 To 12
                    tmPayByContract(ilTotalIndex).lPartner(ilLoopOnMonth) = tmPayByContract(ilTotalIndex).lPartner(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
                Next ilLoopOnMonth
        ElseIf ilCategory = CAT_OWED Then  'owed
'                    For ilLoopOnMonth = 0 To 12
'                        'was Owed to partner (partner full share - partner net sales * partner pct), now partner outstanding (partner pct * net ) minus partners full net
'                         'tmPayByContract(ilTotalIndex).lOwed(ilLoopOnMonth) = tmPayByContract(ilTotalIndex).lOwed(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
'                        slAmount = gLongToStrDec(tmPayByContract(ilTotalIndex).lNet(ilLoopOnMonth), 2)
'                        slSharePct = gLongToStrDec(tmPayByContract(ilTotalIndex).iPartSharePct * 100, 4)
'                        slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
'                        slStr = gRoundStr(slStr, "1", 0)
'                        llPartnerCalc = Val(slStr)
'                        tmPayByContract(ilTotalIndex).lOwed(ilLoopOnMonth) = (tmPayByContract(ilTotalIndex).lOwed(ilLoopOnMonth)) + llPartnerCalc
'                   Next ilLoopOnMonth
        End If
    ElseIf ilTotalType = TOTAL_BYVEHICLE Then     'vehicle totals
        If ilCategory = CAT_GROSS Then      'gross
            For ilLoopOnMonth = 0 To 12
'                        tmPayByVehicle(ilTotalIndex).lGross(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lGross(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
                tmPayByVehicle(ilTotalIndex).lGross(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lGross(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_NET Then  'net
            For ilLoopOnMonth = 0 To 12
'                        tmPayByVehicle(ilTotalIndex).lNet(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lNet(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
                tmPayByVehicle(ilTotalIndex).lNet(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lNet(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_COLLECTED Then  'collected
            For ilLoopOnMonth = 0 To 12
'                        tmPayByVehicle(ilTotalIndex).lCollected(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lCollected(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
                tmPayByVehicle(ilTotalIndex).lCollected(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lCollected(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_OUTSTANDING Then  'outstanding
            For ilLoopOnMonth = 0 To 12
'                        tmPayByVehicle(ilTotalIndex).lOutstanding(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lOutstanding(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
                tmPayByVehicle(ilTotalIndex).lOutstanding(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lOutstanding(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_PARTNER Then  'partner net
            For ilLoopOnMonth = 0 To 12
'                        tmPayByVehicle(ilTotalIndex).lPartner(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lPartner(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
                tmPayByVehicle(ilTotalIndex).lPartnersWorth(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lPartnersWorth(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_PARTNER_COLLECT Then
            For ilLoopOnMonth = 0 To 12
                tmPayByVehicle(ilTotalIndex).lPartner(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lPartner(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_OWED Then  'owed
            For ilLoopOnMonth = 0 To 12
'                        tmPayByVehicle(ilTotalIndex).lOwed(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lOwed(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
                tmPayByVehicle(ilTotalIndex).lOwed(ilLoopOnMonth) = tmPayByVehicle(ilTotalIndex).lOwed(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        End If
    ElseIf ilTotalType = TOTAL_BYPARTNER Then     'participant totals
        If ilCategory = CAT_GROSS Then      'gross
            For ilLoopOnMonth = 0 To 12
'                        tmPayByPartner(ilTotalIndex).lGross(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lGross(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
                tmPayByPartner(ilTotalIndex).lGross(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lGross(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_NET Then  'net
            For ilLoopOnMonth = 0 To 12
'                        tmPayByPartner(ilTotalIndex).lNet(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lNet(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
                tmPayByPartner(ilTotalIndex).lNet(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lNet(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_COLLECTED Then  'collected
            For ilLoopOnMonth = 0 To 12
'                        tmPayByPartner(ilTotalIndex).lCollected(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lCollected(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
                tmPayByPartner(ilTotalIndex).lCollected(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lCollected(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_OUTSTANDING Then  'outstanding
            For ilLoopOnMonth = 0 To 12
'                        tmPayByPartner(ilTotalIndex).lOutstanding(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lOutstanding(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
                tmPayByPartner(ilTotalIndex).lOutstanding(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lOutstanding(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_PARTNER Then  'partner net
            For ilLoopOnMonth = 0 To 12
'                         tmPayByPartner(ilTotalIndex).lPartner(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lPartner(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
                tmPayByPartner(ilTotalIndex).lPartnersWorth(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lPartnersWorth(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
           Next ilLoopOnMonth
        ElseIf ilCategory = CAT_PARTNER_COLLECT Then
            For ilLoopOnMonth = 0 To 12
                 tmPayByPartner(ilTotalIndex).lPartner(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lPartner(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        ElseIf ilCategory = CAT_OWED Then  'owed
            For ilLoopOnMonth = 0 To 12
'                        tmPayByPartner(ilTotalIndex).lOwed(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lOwed(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
                tmPayByPartner(ilTotalIndex).lOwed(ilLoopOnMonth) = tmPayByPartner(ilTotalIndex).lOwed(ilLoopOnMonth) + llMonthSource(ilLoopOnMonth)
            Next ilLoopOnMonth
        End If
        
'            ElseIf ilTotalType = TOTAL_BYFINAL Then     'final
'                If ilCategory = CAT_GROSS Then      'gross
'                    For ilLoopOnMonth = 0 To 12
'                        tmPayByFinal(0).lGross(ilLoopOnMonth) = tmPayByFinal(0).lGross(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
'                    Next ilLoopOnMonth
'                ElseIf ilCategory = CAT_NET Then  'net
'                    For ilLoopOnMonth = 0 To 12
'                        tmPayByFinal(0).lNet(ilLoopOnMonth) = tmPayByFinal(0).lNet(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
'                    Next ilLoopOnMonth
'                ElseIf ilCategory = CAT_COLLECTED Then  'collected
'                    For ilLoopOnMonth = 0 To 12
'                        tmPayByFinal(0).lCollected(ilLoopOnMonth) = tmPayByFinal(0).lCollected(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
'                    Next ilLoopOnMonth
'                ElseIf ilCategory = CAT_OUTSTANDING Then  'outstanding
'                    For ilLoopOnMonth = 0 To 12
'                        tmPayByFinal(0).lOutstanding(ilLoopOnMonth) = tmPayByFinal(0).lOutstanding(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
'                    Next ilLoopOnMonth
'                ElseIf ilCategory = CAT_PARTNER Then  'partner net
'                    For ilLoopOnMonth = 0 To 12
'                        tmPayByFinal(0).lPartner(ilLoopOnMonth) = tmPayByFinal(0).lPartner(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
'                    Next ilLoopOnMonth
'                ElseIf ilCategory = CAT_OWED Then  'owed
'                    For ilLoopOnMonth = 0 To 12
'                        tmPayByFinal(0).lOwed(ilLoopOnMonth) = tmPayByFinal(0).lOwed(ilLoopOnMonth) + tmGrf.lDollars(ilLoopOnMonth)
'                    Next ilLoopOnMonth
'                End If
    End If
    Exit Sub
End Sub

Public Sub mParPayWriteToGRF(ilTotalType As Integer, tlTotalBy() As PAYBY_BREAKOUT)
    Dim ilLoopOnCnt As Integer
    Dim ilTemp As Integer
    Dim ilRet As Integer
    Dim slAmount As String
    Dim slStr As String
    Dim slSharePct As String
    Dim llPartnerCalc As Long
    Dim llSharePct As Long
    'write out the contracts detail 6 records per vehicle
    For ilLoopOnCnt = LBound(tlTotalBy) To UBound(tlTotalBy) - 1
        tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
        tmGrf.iGenDate(1) = igNowDate(1)
        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
        tmGrf.lGenTime = lgNowTime
        tmGrf.lChfCode = tlTotalBy(ilLoopOnCnt).lCode           'contr internal code
        tmGrf.lCode4 = tlTotalBy(ilLoopOnCnt).lCntrNo             'contract #
        tmGrf.iVefCode = tlTotalBy(ilLoopOnCnt).iVefCode
        tmGrf.iCode2 = tlTotalBy(ilLoopOnCnt).iMnfPartCode
        tmGrf.iSofCode = tlTotalBy(ilLoopOnCnt).iPartSharePct
        tmGrf.iAdfCode = tlTotalBy(ilLoopOnCnt).iAdfCode
        tmGrf.sGenDesc = Trim$(tlTotalBy(ilLoopOnCnt).sProductName)

        tmGrf.iPerGenl(0) = ilTotalType     'TOTAL_BYCNT
        'accumulating total year causes overflow into tmgrf.ldollars(12) - accum in crystal
        'participant (partner)
        If ilTotalType = TOTAL_BYPARTNER Then
            tmGrf.lCode4 = 0    '32766
            tmGrf.iVefCode = 0  '32766
        End If
        tmGrf.iPerGenl(1) = CAT_GROSS
        For ilTemp = 0 To 11
            tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lGross(ilTemp)
        Next ilTemp
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
    
        tmGrf.iPerGenl(1) = CAT_NET
        For ilTemp = 0 To 11
            tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lNet(ilTemp)
        Next ilTemp
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        
        tmGrf.iPerGenl(1) = CAT_COLLECTED
        For ilTemp = 0 To 11
            tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lCollected(ilTemp)
        Next ilTemp
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        
        tmGrf.iPerGenl(1) = CAT_OUTSTANDING
        For ilTemp = 0 To 11
            tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lNet(ilTemp) + tlTotalBy(ilLoopOnCnt).lCollected(ilTemp)
            tlTotalBy(ilLoopOnCnt).lOutstanding(ilTemp) = tlTotalBy(ilLoopOnCnt).lNet(ilTemp) + tlTotalBy(ilLoopOnCnt).lCollected(ilTemp)
        Next ilTemp
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        
        tmGrf.iPerGenl(1) = CAT_PARTNER
        For ilTemp = 0 To 11
            tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lPartnersWorth(ilTemp)
        Next ilTemp
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)

        tmGrf.iPerGenl(1) = CAT_PARTNER_COLLECT
        'Partner Collected (Collected * part %)
        'Nothing currently is in Partner buckets
        For ilTemp = 0 To 11
'            slAmount = gLongToStrDec(tlTotalBy(ilLoopOnCnt).lCollected(ilTemp), 2)
'            llSharePct = CDbl(tlTotalBy(ilLoopOnCnt).iPartSharePct) * 100
'            slSharePct = gLongToStrDec(llSharePct, 4)
'            slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
'            slStr = gRoundStr(slStr, "1", 0)
'            llPartnerCalc = Val(slStr)
'            tlTotalBy(ilLoopOnCnt).lPartner(ilTemp) = llPartnerCalc     'payment to partner based on collections
'
'            tmGrf.lDollars(ilTemp) = llPartnerCalc   'partner collected share
            tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lPartner(ilTemp)
        Next ilTemp
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        
        tmGrf.iPerGenl(1) = CAT_OWED
            'Owed is now Partner OUtstanding (Part pct * partner full worth) minus participant paid in collections so far
        For ilTemp = 0 To 11
'            slAmount = gLongToStrDec(tlTotalBy(ilLoopOnCnt).lPartnersWorth(ilTemp), 2)
'            llSharePct = CDbl(tlTotalBy(ilLoopOnCnt).iPartSharePct) * 100
'            slSharePct = gLongToStrDec(llSharePct, 4)
'            slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
'            slStr = gRoundStr(slStr, "1", 0)
'            llPartnerCalc = Val(slStr)
'
'            tmGrf.lDollars(ilTemp) = (tlTotalBy(ilLoopOnCnt).lOwed(ilTemp) + tlTotalBy(ilLoopOnCnt).lPartnersWorth(ilTemp)) + tlTotalBy(ilLoopOnCnt).lPartner(ilTemp)   'outstanding to partner (.lpartner is probably negative)
            tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lOwed(ilTemp)
        Next ilTemp
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
    Next ilLoopOnCnt
    Exit Sub
End Sub

Public Sub mParPayWriteCntToGRF(ilTotalType As Integer, tlTotalBy() As PAYBY_BREAKOUT)
    Dim ilLoopOnCnt As Integer
    Dim ilTemp As Integer
    Dim ilRet As Integer
    Dim slAmount As String
    Dim slStr As String
    Dim slSharePct As String
    Dim llPartnerCalc As Long
    Dim llSharePct As Long
    Dim llNetSales As Long
    Dim llPayments As Long
    Dim blGotNetSales As Boolean        '1-18-19
    Dim blGotPayments As Boolean        '1-18-19

    'write out the contracts detail 6 records per vehicle
    For ilLoopOnCnt = LBound(tlTotalBy) To UBound(tlTotalBy) - 1
        llNetSales = 0
        llPayments = 0
        'if payments and net sales are 0, do not output the contract/vehicle
        '1-18-19 cannot test for 0 because there are cases when adjustments offset the invoice $ and they are in different months, making the result $0
        blGotNetSales = False
        blGotPayments = False
        For ilTemp = 0 To 11
            If tlTotalBy(ilLoopOnCnt).lNet(ilTemp) <> 0 Then
                blGotNetSales = True
            End If
            If tlTotalBy(ilLoopOnCnt).lCollected(ilTemp) <> 0 Then
                blGotPayments = True
            End If
'            llNetSales = llNetSales + tlTotalBy(ilLoopOnCnt).lNet(ilTemp)
'            llPayments = llPayments + tlTotalBy(ilLoopOnCnt).lCollected(ilTemp)
        Next ilTemp
 '       If llNetSales <> 0 Or llPayments <> 0 Then          '1-18-19 adjustments can cause negative values.  change from > to <>
        If (blGotNetSales) Or (blGotPayments) Then
            tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for removal of records
            tmGrf.iGenDate(1) = igNowDate(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tmGrf.lGenTime = lgNowTime
            tmGrf.lChfCode = tlTotalBy(ilLoopOnCnt).lCode           'contr internal code
            tmGrf.lCode4 = tlTotalBy(ilLoopOnCnt).lCntrNo             'contract #
            tmGrf.iVefCode = tlTotalBy(ilLoopOnCnt).iVefCode
            tmGrf.iCode2 = tlTotalBy(ilLoopOnCnt).iMnfPartCode
            tmGrf.iSofCode = tlTotalBy(ilLoopOnCnt).iPartSharePct
            tmGrf.iAdfCode = tlTotalBy(ilLoopOnCnt).iAdfCode
            tmGrf.sGenDesc = Trim$(tlTotalBy(ilLoopOnCnt).sProductName)
    
            tmGrf.iPerGenl(0) = ilTotalType     'TOTAL_BYCNT
           
            tmGrf.iPerGenl(1) = CAT_GROSS
            For ilTemp = 0 To 11
                tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lGross(ilTemp)
            Next ilTemp
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        
            tmGrf.iPerGenl(1) = CAT_NET
            For ilTemp = 0 To 11
                tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lNet(ilTemp)
            Next ilTemp
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            
            tmGrf.iPerGenl(1) = CAT_COLLECTED
            For ilTemp = 0 To 11
                tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lCollected(ilTemp)
            Next ilTemp
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
         
'            'determine participant total share due by cnt
'            For ilTemp = 0 To 11
'                slAmount = gLongToStrDec(tlTotalBy(ilLoopOnCnt).lNet(ilTemp), 2)
'                llSharePct = CDbl(tlTotalBy(ilLoopOnCnt).iPartSharePct) * 100
'                slSharePct = gLongToStrDec(llSharePct, 4)
'                slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
'                slStr = gRoundStr(slStr, "1", 0)
'                tlTotalBy(ilLoopOnCnt).lPartnersWorth(ilTemp) = Val(slStr)   'amt partner should be due based on the net sales and partners participant %
'            Next ilTemp
            
            tmGrf.iPerGenl(1) = CAT_OUTSTANDING
            For ilTemp = 0 To 11
                tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lNet(ilTemp) + tlTotalBy(ilLoopOnCnt).lCollected(ilTemp)        'this is to output the prepass record
                tlTotalBy(ilLoopOnCnt).lOutstanding(ilTemp) = tlTotalBy(ilLoopOnCnt).lNet(ilTemp) + tlTotalBy(ilLoopOnCnt).lCollected(ilTemp)       'this updates the tables
            Next ilTemp
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            
            'Partner Collected (Collected * part %)
            'Nothing currently is in Partner buckets
            If RptSelParPay!rbcBillOrCollect(0).Value Then      'billing rendered
                tmGrf.iPerGenl(1) = CAT_PARTNER
                For ilTemp = 0 To 11
                    tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lPartnersWorth(ilTemp)
                Next ilTemp
            Else
                tmGrf.iPerGenl(1) = CAT_PARTNER
                For ilTemp = 0 To 11
                    tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lPartnersWorth(ilTemp)
                Next ilTemp
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
'                mParPayRollPartnerColAndOutstand CAT_PARTNER, tmGrf.iVefCode, tmGrf.iCode2, tmGrf.iSofCode
                
                tmGrf.iPerGenl(1) = CAT_PARTNER_COLLECT
                For ilTemp = 0 To 11
'                    slAmount = gLongToStrDec(tlTotalBy(ilLoopOnCnt).lCollected(ilTemp), 2)
'                    llSharePct = CDbl(tlTotalBy(ilLoopOnCnt).iPartSharePct) * 100
'                    slSharePct = gLongToStrDec(llSharePct, 4)
'                    slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
'                    slStr = gRoundStr(slStr, "1", 0)
'                    llPartnerCalc = Val(slStr)
'                    tlTotalBy(ilLoopOnCnt).lPartner(ilTemp) = llPartnerCalc     'payment to partner based on collections
'                tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lNet(ilTemp) + tlTotalBy(ilLoopOnCnt).lCollected(ilTemp)
                tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lPartner(ilTemp)
                'tlTotalBy(ilLoopOnCnt).lOutstanding(ilTemp) = tlTotalBy(ilLoopOnCnt).lPartnersWorth(ilTemp) + tlTotalBy(ilLoopOnCnt).lPartner(ilTemp)
                Next ilTemp
            End If
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            
'            mParPayRollPartnerColAndOutstand CAT_PARTNER_COLLECT, tmGrf.iVefCode, tmGrf.iCode2, tmGrf.iSofCode
          
            tmGrf.iPerGenl(1) = CAT_OWED
                'Owed is now Partner OUtstanding (Part pct * partner full worth) minus participant paid in collections so far
            For ilTemp = 0 To 11
'                slAmount = gLongToStrDec(tlTotalBy(ilLoopOnCnt).lPartnersWorth(ilTemp), 2)
'                llSharePct = CDbl(tlTotalBy(ilLoopOnCnt).iPartSharePct) * 100
'                slSharePct = gLongToStrDec(llSharePct, 4)
'                slStr = gMulStr(slSharePct, slAmount)                       ' gross portion of possible split
'                slStr = gRoundStr(slStr, "1", 0)
'                llPartnerCalc = Val(slStr)
    
                tmGrf.lDollars(ilTemp) = tlTotalBy(ilLoopOnCnt).lPartnersWorth(ilTemp) + tlTotalBy(ilLoopOnCnt).lPartner(ilTemp)    'outstanding to partner (partners full worth minus the partner share from collections, which is negative)
                tlTotalBy(ilLoopOnCnt).lOwed(ilTemp) = tmGrf.lDollars(ilTemp)                       'result back into table to rollover totals
            Next ilTemp
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
            mParPayRollPartnerColAndOutstand CAT_OWED, tmGrf.iVefCode, tmGrf.iCode2, tmGrf.iSofCode
        End If              'blGotNetSales or blGotPayments
    Next ilLoopOnCnt
    Exit Sub
End Sub

'
'           Accumulate 1 contracts gross, net collected, partner share, owed $
'           <input> ilCategory  - index into the bucket to accumulate
'                   ilvefcode - vehicle code
'                   ilMnfCode - participant code
Private Sub mParPayRollPartnerColAndOutstand(ilCategory As Integer, ilVefCode As Integer, ilMnfCode As Integer, ilMnfPct As Integer)
    Dim ilLoopOnMonth As Integer

    'Roll over totals for vehicle,
    For ilLoopOnMonth = LBound(tmPayByVehicle) To UBound(tmPayByVehicle) - 1
        If tmPayByVehicle(ilLoopOnMonth).iVefCode = ilVefCode And tmPayByVehicle(ilLoopOnMonth).iMnfPartCode = ilMnfCode And tmPayByVehicle(ilLoopOnMonth).iPartSharePct = ilMnfPct Then
            mParPayRollTotals ilCategory, TOTAL_BYVEHICLE, ilMnfCode, tmGrf.lDollars(), ilLoopOnMonth
            Exit For
        End If
    Next ilLoopOnMonth

    'Roll over totals for participant, do not test the share % could be different across the year
    For ilLoopOnMonth = LBound(tmPayByPartner) To UBound(tmPayByPartner) - 1
        If tmPayByPartner(ilLoopOnMonth).iMnfPartCode = ilMnfCode Then
            mParPayRollTotals ilCategory, TOTAL_BYPARTNER, ilMnfCode, tmGrf.lDollars(), ilLoopOnMonth
           Exit For
        End If
    Next ilLoopOnMonth
    
    Exit Sub
End Sub
