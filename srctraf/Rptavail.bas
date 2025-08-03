Attribute VB_Name = "RPTAVAIL"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptavail.bas on Wed 6/17/09 @ 12:56 P
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  BOOKKEY                                                                               *
'******************************************************************************************

Option Explicit
Option Compare Text
Type BOOKKEY                    'this array is for the Quarterly Booked 'VBC NR
    iRdfCode As Integer         'dp code 'VBC NR
    lChfCode As Long            'contr # 'VBC NR
    lFsfCode As Long            'feed code 'VBC NR
    lRate As Long               'spot rate 'VBC NR
    sSpotType As String * 1     'tmSdf.sSpotType 'VBC NR
    sPriceType As String * 1    'tmsdf.sPricetype 'VBC NR
    sDysTms As String * 40      '11-20-02 DP days & times 'VBC NR
    ivefSellCode As Integer     'vehicle code (selling vehicle) 'VBC NR
    iLen As Integer             'spot length 'VBC NR
    sAirMissed As String * 1     'A=Aired, M = Missed 'VBC NR
    iPkgFlag As Integer          'package line ID (else 0) 'VBC NR
End Type 'VBC NR

Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer        'VEF record length
Dim hmVsf As Integer            'Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
'Quarterly Avails
        'AVR record length
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer
Dim tmClfSrchKey As CLFKEY0     'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmSdf As Integer            'Spot detail file handle
Dim tmSdfSrchKey2 As SDFKEY2            'SDF record image (key 2)
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdf As SDF
Dim tmSdfSrchKey3 As LONGKEY0     'SDF record image (SDF code as keyfield)
Dim hmSsf As Integer
Dim tmSsf As SSF                'SSF record image
Dim tmSsfSrchKey As SSFKEY0      'SSF key record image
Dim tmSsfSrchKey2 As SSFKEY2      'SSF key record image
Dim imSsfRecLen As Integer
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim hmSmf As Integer            'MG file handle
Dim tmSmf As SMF                'SMF record image
Dim imSmfRecLen As Integer        'SMF record length

Dim hmFsf As Integer                    'Feed spot file handle
Dim tmFSFSrchKey As LONGKEY0            'FSF record image
Dim imFsfRecLen As Integer              'FsF record length
Dim tmFsf As FSF

Dim hmAnf As Integer                    'Named avail file handle
Dim tmAnfSrchKey As INTKEY0            'ANF record image
Dim imAnfRecLen As Integer              'ANF record length
Dim tmAnf As ANF

'Log Calendar
Dim hmLcf As Integer            'Log Calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim imLcfRecLen As Integer        'LCF record length


'**********************************************************************
'*
'*      Procedure Name:gGetSpotCounts
'*
'*             Created:12/29/97      By:D. Hosaka
'*
'*             Copy of gGetAvailsCounts to access spot
'*             and save spot data for Quarterly Booked
'*             report
'*
'*            Comments:Obtain the Avail counts and spot
'*            detail
'*            8/5/99 - make generalized routine so it can be
'*              incorporated for all avails reports
'*          7-29-04 option to include/exclude contract or feed spots
'***********************************************************************
'Sub gGetSpotCounts (ilVefcode As Integer, ilVpfIndex As Integer, ilFirstQ As Integer, llSDate As Long, llEDate As Long, llSAvails() As Long, llEAvails() As Long, tlAvRdf() As RDF, tlRif() As RIF, tlCntTypes As CNTTYPES, llIgnoreCodes() As Long)
Sub gGetSpotCounts(tlAvailInfo As AVAILCOUNT, llSAvails() As Long, llEAvails() As Long, tlAvRdf() As RDF, tlRif() As RIF, tlCntTypes As CNTTYPES, llIgnoreCodes() As Long, tlAvr() As AVR, slUserRequest As String)
'
'   Where:
'
'   hmSsf (I) - handle to SSF file
'   hmSdf (I) - handle to Sdf file
'   hmLcf (I) - handle to Lcf file
'   hmChf (I) - handle to Chf file
'   ilVefCode (I) - vehicle code to process
'   ilVpfIndex (I) - vehicle options pointer
'   ilFirstQ (I)
'   llSDate (I) - start date to begin searching Avails
'   llEDate (I) - end date to stop searching avails
'   llSAvails(I)- Array of bucket start dates
'   llEAvails(I)- Array of bucket end dates
'   tlAvRdf() (I) - array of Dayparts
'   tlAvr() (O) - array of AVR records built for avails
'   tlCntTypes (I) - contract and spot types to include in search
'   tlCntTypes.iHold(I)- True = include hold contracts
'   tlCntTypes.iOrder(I)- True= include complete order contracts
'   tlCntTypes.iMissed(I)- True=Include missed
'   tlCntTypes.iXtra(I)- True=Include Xtra bonus spots
'   tlCntTypes.iTrade(I)- True = include trade contracts
'   tlCntTypes.iNC(I)- True = include NC spots
'   tlCntTypes.iReserv(I) - True = include Reservations spots
'   tlCntTypes.iRemnant(I)- True=Include Remnant
'   tlCntTypes.iStandard(I)- true = include std contracts
'   tlCntTypes.iDR(I)- True=Include Direct Response
'   tlCntTypes.iPI(I)- True=Include per Inquiry
'   tlCntTypes.iPSA(I)- True=Include PSA
'   tlCntTypes.iPromo(I)- True=Include Promo
'   tlCntTypes.sAvailType(I) = S=sellout, A = avails, I = inventory, P = % sellout
'   tlCntTypes.iOrphan(I) = true if showing orphan missed spots on separate line
'   tlCntTypes.iDayOption(I) - 0 = Avails by DP, 1 = dp within days, 2 = days within dp
'   tlCntTypes.iBuildDay(I) - true if qtrly booked and need to build addl tables in memory by line
'   tlCntTypes.iShowResvLine(I) - true if reserved spots should be combined with sold
'
'    llIgnoreCodes() - array of contract codes to ignore when picking up spots.  These contracts
                'have been superceded by other C, I, N or G that will be processed from contracts
'               This is for Pressure report only.  All other reports calling this will send llIgnore(0 to 0)
'   slUserRequest : show output in Units (U) or 30/60 (B).  Unit counts should not exceed the # units defined
'   Note: Remnants; Direct Response; per Inquiry; PSA and Promos are not
'         saved with a miss status
'         For scheduled spots the rank is used to determine if it is one
'         of the above (Direct reponse=1010; Remnant=1020; per Inquiry= 1030;
'         PSA=1060; Promo=1050.
'
'   slBucketType(I): A=Avail; S=Sold; I=Inventory  , P = Percent sellout    'forced to "A" for avail
'   3-24-03 change way to test fills & extras.  SSF no longer indicates this flag, need to look at SDF & maybe advt


    Dim ilType As Integer
    Dim slType As String
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilEvt As Integer
    Dim ilRet As Integer
    Dim ilSpot As Integer
    Dim llTime As Long
    Dim ilRdf As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilRec As Integer
    Dim ilRecIndex As Integer
    Dim ilLen As Integer
    Dim ilUnits As Integer
    Dim ilDay As Integer
    Dim ilSaveDay As Integer
    Dim slDays As String
    Dim ilLtfCode As Integer
    Dim ilAvailOk As Integer
    Dim ilPass As Integer
    Dim ilDayIndex As Integer
    Dim ilLoopIndex As Integer
    Dim ilBucketIndex As Integer
    Dim ilSpotOK As Integer
    Dim llLoopDate As Long
    Dim ilWeekDay As Integer
    Dim llLatestDate As Long
    Dim ilIndex As Integer
    Dim slStr As String
    Dim ilAdjAdd As Integer
    Dim ilAdjSub As Integer
    Dim slBucketType As String
    Dim ilLo As Integer
    Dim ilHi As Integer
    Dim ilWkNo As Integer               'week index to rate card
    Dim ilInclCntrSpots As Integer
    Dim ilInclFeedSpots As Integer

    Dim ilOrphanMissedLoop As Integer
    Dim ilOrphanFound As Integer
    Dim ilOrphanMax As Integer
    Dim ilVefCode As Integer
    Dim ilVpfIndex As Integer
    Dim ilFirstQ As Integer
    Dim llSDate As Long
    Dim ilSDate(0 To 1) As Integer
    Dim llEDate As Long
    Dim ilEDate(0 To 1) As Integer
    ReDim ilSAvailsDates(0 To 1) As Integer
    ReDim ilEvtType(0 To 14) As Integer
    ReDim ilRdfCodes(0 To 1) As Integer
    ReDim tlAvr(0 To 0) As AVR
    Dim slChfType As String * 1
    Dim slChfStatus As String * 1
    Dim ilNo30 As Integer                   '3-15-05 reqd for gGatherInventory
    Dim ilNo60 As Integer                   '3-15-05 reqd for gGatherInventory
    Dim ilVefIndex As Integer
    'Dim tlBkKey As BOOKKEY
    'Dim tlCharCurr As KEYCHAR
    'Dim tlCharPrev As KEYCHAR
    ilVefCode = tlAvailInfo.iVefCode            'setup variables sent via structure
    ilVefIndex = gBinarySearchVef(ilVefCode)
    ilVpfIndex = tlAvailInfo.iVpfIndex
    ilFirstQ = tlAvailInfo.iFirstBkt
    llSDate = tlAvailInfo.lSDate
    gPackDateLong llSDate, ilSDate(0), ilSDate(1)
    gPackDateLong llEDate, ilEDate(0), ilEDate(1)
    llEDate = tlAvailInfo.lEDate
    hmLcf = tlAvailInfo.hLcf
    hmSdf = tlAvailInfo.hSdf
    hmSsf = tlAvailInfo.hSsf
    hmSmf = tlAvailInfo.hSmf
    hmCHF = tlAvailInfo.hChf
    hmClf = tlAvailInfo.hClf
    hmCff = tlAvailInfo.hCff
    hmVef = tlAvailInfo.hVef
    hmVsf = tlAvailInfo.hVsf
    hmFsf = tlAvailInfo.hFsf
    hmAnf = tlAvailInfo.hAnf
    imLcfRecLen = Len(tmLcf)
    imSdfRecLen = Len(tmSdf)
    imSsfRecLen = Len(tmSsf)
    imSmfRecLen = Len(tmSmf)
    imCHFRecLen = Len(tmChf)
    imClfRecLen = Len(tmClf)
    imCffRecLen = Len(tmCff)
    imVefRecLen = Len(tmVef)
    imVsfRecLen = Len(tmVsf)
    imFsfRecLen = Len(tmFsf)
    imAnfRecLen = Len(tmAnf)
    slBucketType = tlCntTypes.sAvailType      'avails, sellout, inventory or % sellout?
    ilOrphanMax = 1                           'ignore orphan missed spots whose sold dp is not shown
    If tlCntTypes.iOrphan Then                'show orphan missed spots whose sold dp not shown on separate line?
        ilOrphanMax = 2
    End If
    ilInclFeedSpots = tlCntTypes.iNetwork
    If tlCntTypes.iHold Or tlCntTypes.iOrder Then
        ilInclCntrSpots = True
    Else
        ilInclCntrSpots = False
    End If
    slDate = Format$(llSAvails(1), "m/d/yy")
    gPackDate slDate, ilSAvailsDates(0), ilSAvailsDates(1)

    'ReDim ilWksInMonth(1 To 3) As Integer
    ReDim ilWksInMonth(0 To 3) As Integer       'Index zero ignored
    slStr = slDate
    For ilLoop = 1 To 3 Step 1
        llDate = gDateValue(gObtainStartStd(slStr))
        llLoopDate = gDateValue(gObtainEndStd(slStr)) + 1
        ilWksInMonth(ilLoop) = ((llLoopDate - llDate) / 7)
        slStr = Format(llLoopDate, "m/d/yy")
    Next ilLoop
    'Currently 14 week quarters are not handled - drop 14th week
    If ilWksInMonth(1) + ilWksInMonth(2) + ilWksInMonth(3) > 13 Then
        ilWksInMonth(3) = ilWksInMonth(3) - 1
    End If
    slType = "O"
    ilType = 0
    llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
    'set the type of events to get fro the day (only Contract avails)
    For ilLoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilLoop) = False
    Next ilLoop
    ilEvtType(2) = True
    If tgVpf(ilVpfIndex).sSSellOut = "B" Then           'if units & seconds - add 2 to 30 sec unit and take away 1 fro 60
        ilAdjAdd = 2
        ilAdjSub = 1
    ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then       'if units only - take 1 away from 60 count and add 1 to 30 count
        ilAdjAdd = 1
        ilAdjSub = 1
    End If
    imSdfRecLen = Len(tmSdf)
    imCHFRecLen = Len(tmChf)
    For llLoopDate = llSDate To llEDate Step 1
        slDate = Format$(llLoopDate, "m/d/yy")
        gPackDate slDate, ilDate0, ilDate1
        gObtainWkNo 0, slDate, ilWkNo, ilLo        'obtain the week bucket number
        imSsfRecLen = Len(tmSsf) 'Max size of variable length record
        'tmSsfSrchKey.sType = slType
        If tgMVef(ilVefIndex).sType <> "G" Then
            tmSsfSrchKey.iType = ilType
            tmSsfSrchKey.iVefCode = ilVefCode
            tmSsfSrchKey.iDate(0) = ilDate0
            tmSsfSrchKey.iDate(1) = ilDate1
            tmSsfSrchKey.iStartTime(0) = 0
            tmSsfSrchKey.iStartTime(1) = 0
            ilRet = gSSFGetGreaterOrEqual(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        Else
            tmSsfSrchKey2.iVefCode = ilVefCode
            tmSsfSrchKey2.iDate(0) = ilDate0
            tmSsfSrchKey2.iDate(1) = ilDate1
            ilRet = gSSFGetGreaterOrEqualKey2(hmSsf, tmSsf, imSsfRecLen, tmSsfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)
            ilType = tmSsf.iType
        End If
        'If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.sType <> slType) Or (tmSsf.iVefcode <> ilVefCode Or (tmSsf.iDate(0) <> ilDate0) And (tmSsf.iDate(1) = ilDate1)) Then
        If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iType <> ilType) Or (tmSsf.iVefCode <> ilVefCode Or (tmSsf.iDate(0) <> ilDate0) And (tmSsf.iDate(1) = ilDate1)) Then
            If (llLoopDate > llLatestDate) Then
                ReDim tlLLC(0 To 0) As LLC  'Merged library names
                If tgMVef(ilVefIndex).sType <> "G" Then
                    ilWeekDay = gWeekDayStr(slDate) + 1 '1=Monday TFN; 2=Tues,...7=Sunday TFN
                    If ilWeekDay = 1 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNMO", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 2 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTU", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 3 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNWE", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 4 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNTH", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 5 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNFR", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 6 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSA", "12M", "12M", ilEvtType(), tlLLC())
                    ElseIf ilWeekDay = 7 Then
                         ilRet = gBuildEventDay(ilType, "C", ilVefCode, "TFNSU", "12M", "12M", ilEvtType(), tlLLC())
                    End If
                End If
                'tmSsf.sType = "O"
                tmSsf.iType = ilType
                tmSsf.iVefCode = ilVefCode
                tmSsf.iDate(0) = ilDate0
                tmSsf.iDate(1) = ilDate1
                gPackTime tlLLC(0).sStartTime, tmSsf.iStartTime(0), tmSsf.iStartTime(1)
                tmSsf.iCount = 0
                'tmSsf.iNextTime(0) = 1  'Time not defined
                'tmSsf.iNextTime(1) = 0

                For ilIndex = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1

                    tmAvail.iRecType = Val(tlLLC(ilIndex).sType)
                    gPackTime tlLLC(ilIndex).sStartTime, tmAvail.iTime(0), tmAvail.iTime(1)
                    tmAvail.iLtfCode = tlLLC(ilIndex).iLtfCode
                    tmAvail.iAvInfo = tlLLC(ilIndex).iAvailInfo Or tlLLC(ilIndex).iUnits
                    tmAvail.iLen = CInt(gLengthToCurrency(tlLLC(ilIndex).sLength))
                    tmAvail.ianfCode = Val(tlLLC(ilIndex).sName)
                    tmAvail.iNoSpotsThis = 0
                    tmAvail.iOrigUnit = 0
                    tmAvail.iOrigLen = 0
                    tmSsf.iCount = tmSsf.iCount + 1
                    tmSsf.tPas(ADJSSFPASBZ + tmSsf.iCount) = tmAvail
                Next ilIndex
                ilRet = BTRV_ERR_NONE
            End If
        End If

        'Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.sType = slType) And (tmSsf.iVefcode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
        Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
            gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llDate
            ilBucketIndex = -1
            For ilLoop = 1 To 13 Step 1
                If (llDate >= llSAvails(ilLoop)) And (llDate <= llEAvails(ilLoop)) Then
                    ilBucketIndex = ilLoop
                    Exit For
                End If
            Next ilLoop
            If ilBucketIndex > 0 Then
                ilDay = gWeekDayLong(llDate)
                ilEvt = 1
                Do While ilEvt <= tmSsf.iCount
                   LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    If tmProg.iRecType = 1 Then    'Program (not working for nested prog)
                        ilLtfCode = tmProg.iLtfCode
                    ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 2) Then 'Contract Avails only
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                        'Determine which rate card program this is associated with
                        For ilRdf = LBound(tlAvRdf) To UBound(tlAvRdf) - 1 Step 1

                            ilAvailOk = mRdfEntry(tlAvRdf(), ilRdf, tlCntTypes, ilLtfCode, llDate, slDays, ilLoopIndex)

                            If ilAvailOk Then
                                'Determine if Avr created
                                ilFound = False
                                ilSaveDay = ilDay
                                If tlCntTypes.iDayOption = 0 Then              'daypart option, place all values in same record
                                                                                    'to get better availability
                                    ilDay = 0                                       'force all data in same day of week
                                End If
                                For ilRec = 0 To UBound(tlAvr) - 1 Step 1
                                    If (ilRdfCodes(ilRec) = tlAvRdf(ilRdf).iCode) And (tlAvr(ilRec).iFirstBucket = ilFirstQ) And (tlAvr(ilRec).iDay = ilDay) Then
                                        ilFound = True
                                        ilRecIndex = ilRec
                                        Exit For
                                    End If
                                Next ilRec
                                If Not ilFound Then
                                    ilRecIndex = UBound(tlAvr)
                                    tlAvr(ilRecIndex).iGenDate(0) = igNowDate(0)
                                    tlAvr(ilRecIndex).iGenDate(1) = igNowDate(1)
                                    'tlAvr(ilRecIndex).iGenTime(0) = igNowTime(0)
                                    'tlAvr(ilRecIndex).iGenTime(1) = igNowTime(1)
                                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                    tlAvr(ilRecIndex).lGenTime = lgNowTime
                                    tlAvr(ilRecIndex).iVefCode = ilVefCode
                                    tlAvr(ilRecIndex).iDay = ilDay
                                    tlAvr(ilRecIndex).iQStartDate(0) = ilSAvailsDates(0)
                                    tlAvr(ilRecIndex).iQStartDate(1) = ilSAvailsDates(1)
                                    tlAvr(ilRecIndex).iFirstBucket = ilFirstQ
                                    tlAvr(ilRecIndex).sBucketType = slBucketType
                                    'if fields are switched so that the sort code goes into rdfsortcode and DP code goes into .irdfcode
                                    'then the avails reports must be fixed in each of the formulas  since its grouping is by .irdcode (not .irdfsortcode)

                                    ilRdfCodes(ilRecIndex) = tlAvRdf(ilRdf).iCode
                                    tlAvr(ilRecIndex).iRdfCode = tlRif(ilRdf).iSort   'DP Sort code  from RIF
                                    If tlRif(ilRdf).iSort = 0 Then                  '4-24-15 if the sort code is 0, default it to the rdf internal code, to keep each DP separate so
                                                                                    'overlapping dayparts without sort codes are not combined in the summary version
                                        tlAvr(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iCode
                                    End If
                                    
                                    tlAvr(ilRecIndex).iRdfSortCode = tlAvRdf(ilRdf).iCode   'DP code to retrieve DP name description

                                    tlAvr(ilRecIndex).sInOut = tlAvRdf(ilRdf).sInOut
                                    tlAvr(ilRecIndex).ianfCode = tlAvRdf(ilRdf).ianfCode
                                    tlAvr(ilRecIndex).iDPStartTime(0) = tlAvRdf(ilRdf).iStartTime(0, ilLoopIndex)
                                    tlAvr(ilRecIndex).iDPStartTime(1) = tlAvRdf(ilRdf).iStartTime(1, ilLoopIndex)
                                    tlAvr(ilRecIndex).iDPEndTime(0) = tlAvRdf(ilRdf).iEndTime(0, ilLoopIndex)
                                    tlAvr(ilRecIndex).iDPEndTime(1) = tlAvRdf(ilRdf).iEndTime(1, ilLoopIndex)
                                    tlAvr(ilRecIndex).sDPDays = slDays
                                    tlAvr(ilRecIndex).sNot30Or60 = "N"
                                    ReDim Preserve tlAvr(0 To ilRecIndex + 1) As AVR
                                    ReDim Preserve ilRdfCodes(0 To ilRecIndex + 1)
                                End If
                                tlAvr(ilRecIndex).lRate(ilBucketIndex - 1) = tlRif(ilRdf).lRate(ilWkNo)
                                ilDay = ilSaveDay
                                'Always gather inventory
                                ilLen = tmAvail.iLen
                                ilUnits = tmAvail.iAvInfo And &H1F
                                gGatherInventory tlAvr(), ilVpfIndex, slBucketType, ilRecIndex, ilBucketIndex, ilLen, ilUnits, ilNo30, ilNo60, slUserRequest      '3-15-05 return ilNo30 & ilNo60 added to inventory counts

                                If tlCntTypes.sAvailType <> "I" Then            'unless Inventory only, always calc with spots.  Bypass to speed up processing if Inventory only
                                    'Always calculate Avails
                                    For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)

                                        ilSpotOK = mTestSSFRank(tlCntTypes)

                                        If ilSpotOK Then                            'continue testing other filters
                                            tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilSpotOK = False                    'invalid sdf code
                                            Else
                                                 '3-24-03, check for exclusion of fills & extras
                                                If tmSdf.sSpotType = "X" And Not tlCntTypes.iXtra Then
                                                    ilSpotOK = False
                                                End If
                                            End If
                                            slChfType = ""          'contract types dont apply with feed spots
                                            slChfStatus = ""       'status types dont apply with feed spots

                                            If ilRet = BTRV_ERR_NONE And tmSdf.lChfCode = 0 And ilSpotOK Then        'feed spot
                                               'obtain the network information
                                               tmFSFSrchKey.lCode = tmSdf.lFsfCode
                                               ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                               If ilRet <> BTRV_ERR_NONE Or Not ilInclFeedSpots Then
                                                   ilSpotOK = False                    'invalid network code
                                               Else
                                                    mGatherSpotSold tlAvr(), tlCntTypes, ilVpfIndex, slBucketType, ilRecIndex, ilBucketIndex, slChfStatus, slChfType
                                                End If
                                            Else            'Test for contract spots
                                                If ilRet = BTRV_ERR_NONE And tmSdf.lChfCode > 0 And ilSpotOK Then
                                                    'obtain contract info
                                                    If tmSpot.lSdfCode = tmSdf.lCode And ilRet = BTRV_ERR_NONE And ilSpotOK = True Then
                                                        If tmSdf.lChfCode <> tmChf.lCode Then               'if already in mem, don't reread
                                                            tmChfSrchKey.lCode = tmSdf.lChfCode
                                                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                                            If ilRet <> BTRV_ERR_NONE Then
                                                                ilSpotOK = False
                                                            End If
                                                        End If      'tmSdf.lChfCode <> tmChf.lCode
                                                        slChfType = tmChf.sType
                                                        slChfStatus = tmChf.sStatus
                                                        'Determine if spot within avail is OK to include in report
                                                        mProcessCntrChange tlAvRdf(), llIgnoreCodes(), ilSpotOK
                                                        If ilSpotOK Then
                                                            mGatherSpotSold tlAvr(), tlCntTypes, ilVpfIndex, slBucketType, ilRecIndex, ilBucketIndex, slChfStatus, slChfType
                                                        End If
                                                    End If          'tmSpot.lSdfCode = tmSdf.lCode And ilRet = BTRV_ERR_NONE And ilSpotOK = True
                                                End If              'ilRet = BTRV_ERR_NONE And tmSdf.lChfCode > 0 And ilSpotOK
                                            End If                  'ilRet = BTRV_ERR_NONE And tmSdf.lChfCode = 0 And ilSpotOK Then        'feed spot
                                        End If
                                    Next ilSpot             'loop from ssf file for # spots in avail
                                End If                      'cntypes.savailtype <> "I"
                            End If                          'Avail OK
                        Next ilRdf                          'ilRdf = lBound(tlAvRdf)
                        ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                    End If
                    ilEvt = ilEvt + 1   'Increment to next event
                Loop                                                        'do while ilEvt <= tmSsf.iCount
            End If                                                              'ilBucketIndex > 0
            imSsfRecLen = Len(tmSsf) 'Max size of variable length record
            ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            If tgMVef(ilVefIndex).sType = "G" Then
                ilType = tmSsf.iType
            End If
        Loop
    Next llLoopDate

    'Get missed
    '3/30/99 For each missed status (missed, ready, & unscheduled) there are up to 2 passes
    'for each spot.  The 1st pass looks or a daypart that matches the shedule lines DP.
    'If found, the missed spot is placed in that DP (if that DP is to be shown on the report).
    'If no DP are found that match, the 2nd pass places it in the first DP that surrounds
    'the missed spots time.
    If (tlCntTypes.iMissed) Then
        'Key 2: VefCode; SchStatus; AdfCode; Date, Time
        For ilPass = 0 To 2 Step 1
            tmSdfSrchKey2.iVefCode = ilVefCode
            If ilPass = 0 Then
                slType = "M"
            ElseIf ilPass = 1 Then
                slType = "R"
            ElseIf ilPass = 2 Then
                slType = "U"
            End If
            tmSdfSrchKey2.sSchStatus = slType
            tmSdfSrchKey2.iAdfCode = 0
            tmSdfSrchKey2.iDate(0) = ilSAvailsDates(0)
            tmSdfSrchKey2.iDate(1) = ilSAvailsDates(1)
            tmSdfSrchKey2.iTime(0) = 0
            tmSdfSrchKey2.iTime(1) = 0
            ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
            'This code added as replacement for Ext operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.sSchStatus = slType)
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate
                If (llDate >= llSAvails(ilFirstQ)) And (llDate <= llEAvails(13)) Then
                    ilBucketIndex = -1
                    For ilLoop = 1 To 13 Step 1
                        If (llDate >= llSAvails(ilLoop)) And (llDate <= llEAvails(ilLoop)) Then
                            ilBucketIndex = ilLoop
                            Exit For
                        End If
                    Next ilLoop
                    If ilBucketIndex > 0 And ((ilInclCntrSpots And tmSdf.lChfCode > 0) Or (ilInclFeedSpots And tmSdf.lChfCode = 0)) Then
                        ilDay = gWeekDayLong(llDate)
                        gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                        For ilOrphanMissedLoop = 1 To ilOrphanMax
                            ilOrphanFound = False
                            For ilRdf = LBound(tlAvRdf) To UBound(tlAvRdf) - 1 Step 1

                                ilAvailOk = False
                                If (tlAvRdf(ilRdf).iLtfCode(0) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(1) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(2) <> 0) Then
                                    If (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(0)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Then
                                        ilAvailOk = False    'True- code later
                                    End If
                                Else
                                    For ilLoop = LBound(tlAvRdf(ilRdf).iStartTime, 2) To UBound(tlAvRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                        If (tlAvRdf(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tlAvRdf(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                                            gUnpackTimeLong tlAvRdf(ilRdf).iStartTime(0, ilLoop), tlAvRdf(ilRdf).iStartTime(1, ilLoop), False, llStartTime
                                            gUnpackTimeLong tlAvRdf(ilRdf).iEndTime(0, ilLoop), tlAvRdf(ilRdf).iEndTime(1, ilLoop), True, llEndTime
                                            If UBound(tlAvRdf) - 1 = LBound(tlAvRdf) Then   'could be a conv bumped spot sched in
                                                                                        'in conven veh.  The VV has DP times different than the
                                                                                        'conven veh.
                                                llStartTime = llTime
                                                llEndTime = llTime + 1              'actual time of spot
                                            End If
                                            'Don't include the end time i.e. 10a-3p is 10a thru 2:59:59p
                                            'If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDay + 1) = "Y") Then
                                            If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDay) = "Y") Then
                                                ilAvailOk = True
                                                ilLoopIndex = ilLoop
                                                slDays = ""
                                                For ilDayIndex = 1 To 7 Step 1
                                                    If (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "Y") Or (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "N") Then
                                                        slDays = slDays & tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1)
                                                    Else
                                                        slDays = slDays & "N"
                                                    End If
                                                Next ilDayIndex
                                                Exit For
                                            End If
                                        End If
                                    Next ilLoop
                                End If
                                If ilAvailOk Or ilOrphanMissedLoop = 2 Then
                                    ilSpotOK = True                'assume spot is OK
                                    If tmSdf.lChfCode = 0 Then
                                        'obtain the network information
                                        tmFSFSrchKey.lCode = tmSdf.lFsfCode
                                        ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                        If ilRet <> BTRV_ERR_NONE Or Not ilInclFeedSpots Then
                                            ilSpotOK = False                    'invalid network code
                                        End If

                                    Else
                                        ilRet = BTRV_ERR_NONE
                                        tmClfSrchKey.lChfCode = tmSdf.lChfCode
                                        tmClfSrchKey.iLine = tmSdf.iLineNo
                                        tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
                                        tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                                        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)

                                        If (tmClf.lChfCode <> tmSdf.lChfCode) Then     'got the matching line reference?
                                            ilSpotOK = False
                                        Else
                                            If ilOrphanMissedLoop = 1 And tlCntTypes.iOrphan Then   '1st pass to see if spot falls in a sold DP, and show on separate line?
                                                If tmClf.iRdfCode <> tlAvRdf(ilRdf).iCode Then
                                                    ilSpotOK = False
                                                End If
                                            End If
                                        End If
                                        ilRet = BTRV_ERR_NONE
                                        If tmSdf.lChfCode <> tmChf.lCode Then               'if already in mem, don't reread
                                            tmChfSrchKey.lCode = tmSdf.lChfCode
                                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilSpotOK = False
                                            End If
                                        End If

                                        slChfStatus = tmChf.sStatus         '4-24-15 if hold, need to put into correct category.  pass to mGatherSpotSold
                                        slChfType = tmChf.sType
                                        
                                        If tmChf.sType = "T" And Not tlCntTypes.iRemnant Then
                                            ilSpotOK = False
                                        End If
                                        If tmChf.sType = "Q" And Not tlCntTypes.iPI Then
                                            ilSpotOK = False
                                        End If
                                        If tmChf.iPctTrade = 100 And Not tlCntTypes.iTrade Then
                                            ilSpotOK = False
                                        End If
                                        If tmSdf.sSpotType = "X" And Not tlCntTypes.iXtra Then
                                            ilSpotOK = False
                                        End If
                                        If tmChf.sType = "M" And Not tlCntTypes.iPromo Then
                                            ilSpotOK = False
                                        End If
                                        If tmChf.sType = "S" And Not tlCntTypes.iPSA Then
                                            ilSpotOK = False
                                        End If


                                        'Determine if spot within avail is OK to include in report
                                        mProcessCntrChange tlAvRdf(), llIgnoreCodes(), ilSpotOK
                                    End If

                                    If ilSpotOK Then
                                        ilOrphanFound = True
                                        'Determine if Avr created
                                        ilFound = False
                                        ilSaveDay = ilDay
                                        If tlCntTypes.iDayOption = 0 Then              'daypart option, place all values in same record
                                                                                            'to get better availability
                                            ilDay = 0                                       'force all data in same day of week
                                        End If

                                        If ilOrphanMissedLoop = 2 Then          'orphans to show on separate line
                                            For ilRec = 0 To UBound(tlAvr) - 1 Step 1
                                            If (ilRdfCodes(ilRec) = -1) And (tlAvr(ilRec).iFirstBucket = ilFirstQ) And (tlAvr(ilRec).iDay = ilDay) Then
                                                ilFound = True
                                                ilRecIndex = ilRec
                                                Exit For
                                            End If
                                            Next ilRec
                                        Else                                    'not an orphan spot, find the DP entry in the table

                                            For ilRec = 0 To UBound(tlAvr) - 1 Step 1
                                                'If (tlAvr(ilRec).iRdfCode = tlAvRdf(ilRdf).iCode) And (tlAvr(ilRec).iFirstBucket = ilFirstQ) And (tlAvr(ilRec).iDay = ilDay) Then
                                                If (ilRdfCodes(ilRec) = tlAvRdf(ilRdf).iCode) And (tlAvr(ilRec).iFirstBucket = ilFirstQ) And (tlAvr(ilRec).iDay = ilDay) Then
                                                    ilFound = True
                                                    ilRecIndex = ilRec
                                                    Exit For
                                                End If
                                            Next ilRec
                                        End If
                                        If Not ilFound Then
                                            ilRecIndex = UBound(tlAvr)
                                            tlAvr(ilRecIndex).iGenDate(0) = igNowDate(0)
                                            tlAvr(ilRecIndex).iGenDate(1) = igNowDate(1)
                                            'tlAvr(ilRecIndex).iGenTime(0) = igNowTime(0)
                                            'tlAvr(ilRecIndex).iGenTime(1) = igNowTime(1)
                                            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                            tlAvr(ilRecIndex).lGenTime = lgNowTime
                                            tlAvr(ilRecIndex).iDay = ilDay
                                            tlAvr(ilRecIndex).iQStartDate(0) = ilSAvailsDates(0)
                                            tlAvr(ilRecIndex).iQStartDate(1) = ilSAvailsDates(1)
                                            tlAvr(ilRecIndex).iFirstBucket = ilFirstQ
                                            tlAvr(ilRecIndex).sBucketType = slBucketType
                                            tlAvr(ilRecIndex).iDPStartTime(0) = tlAvRdf(ilRdf).iStartTime(0, ilLoopIndex)
                                            tlAvr(ilRecIndex).iDPStartTime(1) = tlAvRdf(ilRdf).iStartTime(1, ilLoopIndex)
                                            tlAvr(ilRecIndex).iDPEndTime(0) = tlAvRdf(ilRdf).iEndTime(0, ilLoopIndex)
                                            tlAvr(ilRecIndex).iDPEndTime(1) = tlAvRdf(ilRdf).iEndTime(1, ilLoopIndex)
                                            tlAvr(ilRecIndex).sDPDays = slDays
                                            tlAvr(ilRecIndex).sNot30Or60 = "N"

                                            tlAvr(ilRecIndex).iVefCode = ilVefCode
                                            'tlAvr(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iCode           'DP code
                                            'tlAvr(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iSortCode
                                            tlAvr(ilRecIndex).iRdfCode = tlRif(ilRdf).iSort   'DP Sort code  from RIF
                                            tlAvr(ilRecIndex).iRdfSortCode = tlAvRdf(ilRdf).iCode   'DP code to retrieve DP name description
                                            ilRdfCodes(ilRecIndex) = tlAvRdf(ilRdf).iCode
                                            tlAvr(ilRecIndex).sInOut = tlAvRdf(ilRdf).sInOut
                                            tlAvr(ilRecIndex).ianfCode = tlAvRdf(ilRdf).ianfCode
                                            tlAvr(ilRecIndex).sDPDays = slDays

                                            If ilOrphanMissedLoop = 2 Then
                                                'override some of the codes if its in the orphan pass (where no shown DP equals the DP of the missed spot)
                                                ilRdfCodes(ilRecIndex) = -1         'phoney daypart for orphaned missed spots 'tlAvrdf(ilRdf).icode
                                                tlAvr(ilRecIndex).iRdfCode = 32000     'sort it last tmRifSorts(ilRdf).isort   'DP Sort code  from RIF
                                                tlAvr(ilRecIndex).iRdfSortCode = -1 'tlAvrdf(ilRdf).icode   'DP code to retrieve DP name description
                                            End If

                                            ReDim Preserve tlAvr(0 To ilRecIndex + 1) As AVR
                                            ReDim Preserve ilRdfCodes(0 To ilRecIndex + 1)
                                        End If
                                        tlAvr(ilRecIndex).lRate(ilBucketIndex - 1) = tlRif(ilRdf).lRate(ilWkNo)
                                        ilDay = ilSaveDay
                                        If ilSpotOK Then
                                            mGatherSpotSold tlAvr(), tlCntTypes, ilVpfIndex, slBucketType, ilRecIndex, ilBucketIndex, slChfStatus, slChfType
                                        End If
                                        If ilSpotOK Then
                                            Exit For                'force exit on this missed if found a matching daypart
                                        End If
                                    End If                      'ilSpotOK
                                End If                          'ilAvailOK
                                If ilOrphanMissedLoop = 2 Then
                                    Exit For
                                End If
                            Next ilRdf
                            If ilOrphanFound Then
                                Exit For
                            End If
                        Next ilOrphanMissedLoop
                    End If
                End If
                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next ilPass
    End If

    'Adjust counts
    'If (slBucketType = "A") And (tgVpf(ilVpfIndex).sSSellOut = "B" Or tgVpf(ilVpfIndex).sSSellOut = "U") Then
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
        Next ilRec
    End If
    'Adjust counts for qtrly detail availbilty
    'If (RptSelCt!rbcSelC4(1).Value) And (tgVpf(ilVpfIndex).sSSellOut = "B" Or tgVpf(ilVpfIndex).sSSellOut = "U") Then                  'qtrly detail?
    'If (tgVpf(ilVpfIndex).sSSellOut = "B" Or tgVpf(ilVpfIndex).sSSellOut = "U") Then                  'qtrly detail?
    If (tgVpf(ilVpfIndex).sSSellOut = "B") Then                   'qtrly detail?
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
    'Combines weeks into the proper months for monthly figures
    For ilRec = 0 To UBound(tlAvr) - 1 Step 1  'next daypart
        For ilLoop = 1 To 3 Step 1
            If ilLoop = 1 Then
                ilLo = 1
                ilHi = ilWksInMonth(1)
            Else
                ilLo = ilHi + 1
                ilHi = ilHi + ilWksInMonth(ilLoop)
            End If
            For ilIndex = ilLo To ilHi Step 1
                If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                    tlAvr(ilRec).lMonth(ilLoop - 1) = tlAvr(ilRec).lMonth(ilLoop - 1) + (((tlAvr(ilRec).i60Avail(ilIndex - 1) * 2) + tlAvr(ilRec).i30Avail(ilIndex - 1)) * tlAvr(ilRec).lRate(ilIndex - 1))
                ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Or tgVpf(ilVpfIndex).sSSellOut = "U" Then
                    tlAvr(ilRec).lMonth(ilLoop - 1) = tlAvr(ilRec).lMonth(ilLoop - 1) + ((tlAvr(ilRec).i60Avail(ilIndex - 1) + tlAvr(ilRec).i30Avail(ilIndex - 1)) * tlAvr(ilRec).lRate(ilIndex - 1))
                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                End If
            Next ilIndex
        Next ilLoop
    Next ilRec
    Erase ilSAvailsDates
    Erase ilEvtType
    Erase ilRdfCodes
    Erase tlLLC
End Sub
'
'
'           mRdfEntry - obtain the Daypart entry matching the time of avail
'
'           <input> tlAvRdf - array of dayparts for the selcted R/C
'                   ilRdf - index into the Daypart entry
'                   ilLtfCode - library code
'                   llDate - date of avail
'           <output> slDays - valid days to air for daypart
'                   ilLoopIndex - index to the matching DP entry
'           Return - ilAvailOK - true if within daypart time range, else false

Public Function mRdfEntry(tlAvRdf() As RDF, ilRdf As Integer, tlCntTypes As CNTTYPES, ilLtfCode As Integer, llDate As Long, slDays As String, ilLoopIndex As Integer) As Integer
Dim ilLoop As Integer
Dim llStartTime As Long
Dim llEndTime As Long
Dim ilDayIndex As Integer
Dim ilAvailOk  As Integer
Dim ilDay As Integer
Dim llTime As Long
Dim ilRet As Integer

    ilAvailOk = False             'assume the avail is NG
    ilDay = gWeekDayLong(llDate)
    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime

    If (tlAvRdf(ilRdf).iLtfCode(0) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(1) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(2) <> 0) Then
        If (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(0)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Then
            ilAvailOk = False    'True- code later
        End If
    Else
        For ilLoop = LBound(tlAvRdf(ilRdf).iStartTime, 2) To UBound(tlAvRdf(ilRdf).iStartTime, 2) Step 1 'Row
            If (tlAvRdf(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tlAvRdf(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                gUnpackTimeLong tlAvRdf(ilRdf).iStartTime(0, ilLoop), tlAvRdf(ilRdf).iStartTime(1, ilLoop), False, llStartTime
                gUnpackTimeLong tlAvRdf(ilRdf).iEndTime(0, ilLoop), tlAvRdf(ilRdf).iEndTime(1, ilLoop), True, llEndTime
                'If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDay + 1) = "Y") Then
                If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDay) = "Y") Then
                    ilAvailOk = True
                    ilLoopIndex = ilLoop
                    slDays = ""
                    For ilDayIndex = 1 To 7 Step 1
                        If (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "Y") Or (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1) = "N") Then
                            slDays = slDays & tlAvRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1)
                        Else
                            slDays = slDays & "N"
                        End If
                    Next ilDayIndex
                    Exit For
                End If
            End If
        Next ilLoop
    End If

    If ilAvailOk Then
        If tlAvRdf(ilRdf).sInOut = "I" Then   'Book into
            If tmAvail.ianfCode <> tlAvRdf(ilRdf).ianfCode Then
                ilAvailOk = False
            End If
        ElseIf tlAvRdf(ilRdf).sInOut = "O" Then   'Exclude
            If tmAvail.ianfCode = tlAvRdf(ilRdf).ianfCode Then
                ilAvailOk = False
            End If
        End If
        '7-19-04 the Named avail property must allow local spots to be included
        tmAnfSrchKey.iCode = tmAvail.ianfCode
        ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If (ilRet = BTRV_ERR_NONE) Then
            If ((Not tlCntTypes.iOrder) And (Not tlCntTypes.iHold)) And tmAnf.sBookLocalFeed = "L" Then      'Local avail requested to be excluded, exclude if avail type = "L"
                ilAvailOk = False
            End If
            If Not tlCntTypes.iNetwork And tmAnf.sBookLocalFeed = "F" Then      'Network avail requested to be excluded, exclude if avail type = "F"
                ilAvailOk = False
            End If
        End If
        'allow the avail to be gathered if the field doesnt have a value, indicating an original avail defined as Both
        'allow the avail to be gathered even if the named avail code isnt found
    End If

    mRdfEntry = ilAvailOk
End Function
'
'           'accumulate the avails inventory
'           <input> ilVpfIndex - index to vehicles options table
'                   slBucketType - avails, sellout, inventory or % sellout
'                   ilRecIndex - index to the daypart avails buckets
'                   ilBucketIndex - week # to accumulate within the ilrecindex
'                   ilLen - length of avail
'                   ilUnits - # units of avail
'           <output>
'                   ilNo30 - count of 30s added to inventory (used to force the spots
'                           sold if excluding locked avail.  need to show the locked
'                           avails as tho they are soldout)
'                   ilNo60 - count of 60s added to inventory (same as ilNo30, except for 60")

'           12-14-04 ignore avail if no units are defined
'           3-15-05 Pass an additional 2 parameters (ilNo30 & ilNo60) which
'                   indicates how much inventory was accumulated.  Required if
'                   excluding locked avails.
Public Sub gGatherInventory(tlAvr() As AVR, ilVpfIndex As Integer, slBucketType As String, ilRecIndex As Integer, ilBucketIndex As Integer, ilLen As Integer, ilUnits As Integer, ilNo30 As Integer, ilNo60 As Integer, slUserRequest As String)

'   slUserRequest : show output in Units (U) or 30/60 (B).  Unit counts should not exceed the # units defined
'                   6-14-19 C = Counts - Implemented for Quarterly Booked to strictly use # of units for inventory

Dim ilUnitCount As Integer
Dim ilTempUnits As Integer
Dim ilTempLen As Integer


        ilNo30 = 0
        ilNo60 = 0
        If ilUnits = 0 Then     'ignore the avail if no units are defined
            Exit Sub
        End If

        If slUserRequest = "C" Then                    'C = unit counts only, nothing to do with hold the vehicle is sold
            ilNo30 = ilUnits
        Else
            If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                If slUserRequest = "B" Then         'show with both 30s & 60s (2 columns)
                    'Convert inventory to number of 30's and 60's
                    Do While ilLen >= 60
                        ilNo60 = ilNo60 + 1
                        ilLen = ilLen - 60
                    Loop
                    Do While ilLen >= 30
        
                        ilNo30 = ilNo30 + 1
                        ilLen = ilLen - 30
                    Loop
                    If ilLen < 30 And ilLen > 0 Then    '7-6-00 assume anything under 30" is 1-30" unit availability
                        ilNo30 = ilNo30 + 1
                        ilLen = 0
                    End If
                Else                        'selling method is 30/60, but show by units
                                            'do not double the 60, cannot exceed # of units that actually are defined
                    'try to maximum the avail in terms of units
                    ilTempUnits = ilUnits
                    ilTempLen = ilLen
                    ilUnitCount = 0
                    Do While ilTempUnits > 1 And ilTempLen >= 60        'test for 2/60 or more
                        ilTempUnits = ilTempUnits - 2
                        ilTempLen = ilTempLen - 60
                        ilUnitCount = ilUnitCount + 2
                        ilNo60 = ilNo60 + 2
                    Loop
                    
                    Do While ilTempUnits > 0 And ilTempLen >= 60        'test for 1 unit avails of 60"; do not want to show that in terms of 30" units
                        ilTempUnits = ilTempUnits - 1
                        ilTempLen = ilTempLen - 60
                        ilUnitCount = ilUnitCount + 1
                        ilNo60 = ilNo60 + 1
                    Loop
                    
                    Do While ilTempUnits > 0 And (ilTempLen <= 30 And ilTempLen > 0)       'test for  avails of 30" or less; anything less/equal to 30 is considered 1 or more units (i.e. 2/30 could put in 2 15")
                        ilTempUnits = ilTempUnits - 1
                        ilTempLen = ilTempLen - ilLen
                        ilUnitCount = ilUnitCount + 1
                        ilNo30 = ilNo30 + 1
                    Loop
                    
                    'ilNo30 = ilUnitCount
                    'ilNo60 = 0                  'dont want the crystal side to double the 60s for 30" avails in case not enuf units defined
                End If
    
            ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                'Count 30 or 60 and set flag if neither
                If ilLen = 60 Then
                    ilNo60 = 1
                ElseIf ilLen <= 30 Then
                    ilNo30 = 1
                Else
                    tlAvr(ilRecIndex).sNot30Or60 = "Y"
                    If ilLen <= 30 Then
                        ilNo30 = 1
                    Else
                        ilNo60 = 1
                    End If
                End If
            ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                'Count 30 or 60 and set flag if neither
                If ilLen = 60 Then
                    ilNo60 = 1
                ElseIf ilLen <= 30 Then
                    ilNo30 = 1
                Else
                    tlAvr(ilRecIndex).sNot30Or60 = "Y"
                End If
            ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
            End If
        End If
        If slBucketType <> "S" And slBucketType <> "P" Then    'sellout by min or pcts, don't update these yet
            tlAvr(ilRecIndex).i30Count(ilBucketIndex - 1) = tlAvr(ilRecIndex).i30Count(ilBucketIndex - 1) + ilNo30
            tlAvr(ilRecIndex).i60Count(ilBucketIndex - 1) = tlAvr(ilRecIndex).i60Count(ilBucketIndex - 1) + ilNo60
        End If
        'always put total inventory into record and avail bucket (avail bucket for qtrly detail)
        tlAvr(ilRecIndex).i30InvCount(ilBucketIndex - 1) = tlAvr(ilRecIndex).i30InvCount(ilBucketIndex - 1) + ilNo30
        tlAvr(ilRecIndex).i60InvCount(ilBucketIndex - 1) = tlAvr(ilRecIndex).i60InvCount(ilBucketIndex - 1) + ilNo60
        tlAvr(ilRecIndex).i30Avail(ilBucketIndex - 1) = tlAvr(ilRecIndex).i30Avail(ilBucketIndex - 1) + ilNo30
        tlAvr(ilRecIndex).i60Avail(ilBucketIndex - 1) = tlAvr(ilRecIndex).i60Avail(ilBucketIndex - 1) + ilNo60

End Sub
'
'           mTestSSFRank - test the ranks in the SSF spot entry
'           to see if it is to be included/excluded
'           Return - true if include spot
'
Public Function mTestSSFRank(tlCntTypes As CNTTYPES) As Integer
Dim ilSpotOK As Integer
        ilSpotOK = True                             'assume spot is OK to include

        If ((tmSpot.iRank And RANKMASK) = REMNANTRANK) And (Not tlCntTypes.iRemnant) Then
            ilSpotOK = False
        End If
        If ((tmSpot.iRank And RANKMASK) = PERINQUIRYRANK) And (Not tlCntTypes.iPI) Then
            ilSpotOK = False
        End If
        If ((tmSpot.iRank And RANKMASK) = TRADERANK) And (Not tlCntTypes.iTrade) Then
            ilSpotOK = False
        End If
        '3-24-03 ranking no longer indicates fill vs extra
        'If tmSpot.iRank = 1045 And Not tlCntTypes.iXtra Then
        '    ilSpotOK = False
        'End If
        If ((tmSpot.iRank And RANKMASK) = PROMORANK) And (Not tlCntTypes.iPromo) Then
            ilSpotOK = False
        End If
        If ((tmSpot.iRank And RANKMASK) = PSARANK) And (Not tlCntTypes.iPSA) Then
            ilSpotOK = False
        End If

        mTestSSFRank = ilSpotOK
End Function
'
'           mGatherSpotSold - accumulate the spot time sold
'
'           <input> tlCnttypes - record containing the user options for inclusion/exclusion
'                   ilVpfIndex - index to vehicle option table
'                   slBucketType - type of avails the to be reported (sellout, avails, percent)
'                   ilRecIndex - index to the avail daypart infor
'                   ilBucketIndex - index to the week to accumulate spot info
'                   slChfStatus - status of contract (holds/orders), or blank if feed spot
'                   slChfType - type of contract (std/remnant/psa, etc) or blank if feed spot
'
Public Sub mGatherSpotSold(tlAvr() As AVR, tlCntTypes As CNTTYPES, ilVpfIndex As Integer, slBucketType As String, ilRecIndex As Integer, ilBucketIndex As Integer, slChfStatus As String, slChfType As String)
Dim ilLen As Integer
Dim ilNo30 As Integer
Dim ilNo60 As Integer
Dim ilLoop As Integer
Dim ilBucketIndexMinusOne As Integer

ilBucketIndexMinusOne = ilBucketIndex - 1

        ilLen = tmSdf.iLen
        ilNo30 = 0
        ilNo60 = 0
        If tgVpf(ilVpfIndex).sSSellOut = "B" Then                   'both units and seconds
            'Convert inventory to number of 30's and 60's
            Do While ilLen >= 60
                ilNo60 = ilNo60 + 1
                ilLen = ilLen - 60
            Loop
            Do While ilLen >= 30
                ilNo30 = ilNo30 + 1
                ilLen = ilLen - 30
            Loop
            If ilLen < 30 And ilLen > 0 Then    '7-6-00 assume anything under 30" is 1-30" unit availability
                ilNo30 = ilNo30 + 1
                ilLen = 0
            End If

            If (slBucketType = "S") Or (slBucketType = "P") Then    'sellout or %sellout, accum sold
                If tlCntTypes.iDetail Then                        'qtrly detail report (has detail for sch lines)
                    If slChfType = "V" Then                       'Type reserve
                        If tlCntTypes.iShowReservLine Then          'show reserves (vs hide)
                            tlAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) + ilNo30
                            tlAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) + ilNo60
                         Else
                            tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                            tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                         End If
                    ElseIf slChfStatus = "H" Then                         'staus "Hold" , always show on separate line
                        tlAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) + ilNo30
                        tlAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) + ilNo60
                    Else
                        tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                        tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                    End If
                Else
                    tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                    tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                End If
            Else
                tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo60
                tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) - ilNo30
            End If
            'adjust the available buckets (used for qtrly detail  report only)
            tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
            tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
        ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then               'units sold
            'Count 30 or 60 and set flag if neither
            If ilLen = 60 Then
                ilNo60 = 1
            ElseIf ilLen = 30 Then
                ilNo30 = 1
            Else
                tlAvr(ilRecIndex).sNot30Or60 = "Y"
                If ilLen <= 30 Then
                    ilNo30 = 1
                Else
                    ilNo60 = 1
                End If
            End If
            If (ilNo60 <> 0) Or (ilNo30 <> 0) Then
                If (slBucketType = "S") Or (slBucketType = "P") Then
                    If tlCntTypes.iDetail Then                        'qtrly detail spots option
                        If slChfType = "V" Then                       'Type reserve
                            If tlCntTypes.iShowReservLine Then
                                tlAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) + ilNo30
                                tlAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) + ilNo60
                            Else
                                tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                            End If
                        ElseIf slChfStatus = "H" Then                         'staus "Hold", always show on separate line
                            tlAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) + ilNo30
                            tlAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) + ilNo60
                        Else
                            tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                            tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                        End If
                    Else
                           tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                           tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                    End If
                Else
                    If ilNo60 > 0 Then                     'spot found a 60?
                        tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) - ilNo60
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
                    End If
                End If
            End If
            'adjust the available buckets  (used for detail version only)
            If ilNo60 > 0 Then      '60 can only take away from 60s bucket since it cant go into anything less
                tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
            Else        'must be 30 or less
                'see if there are 30s to subtract, may have to take from 60 since its unit based
                If tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) > 0 Then   'theres enuf 30s to take away from
                    tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                Else
                    'take away from the 60s bucket if any exist, otherwise oversell the 30 avail
                    If tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) > 0 Then
                        tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo30
                    Else
                        tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                    End If
                End If
            End If
        ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then               'matching units
            'Count 30 or 60 and set flag if neither
            If ilLen = 60 Then
                ilNo60 = 1
            ElseIf ilLen = 30 Then
                ilNo30 = 1
            Else
                tlAvr(ilRecIndex).sNot30Or60 = "Y"
            End If
            If (slBucketType = "S") Or (slBucketType = "P") Then        'if Sellout or % sellout, accum the seconds sold
            'Qtrly detail has been forced to "Sellout" for internal testing
                If tlCntTypes.iDetail Then                    'qtrly detail booked has more options
                    If slChfType = "V" Then                       'Type reserve
                        If tlCntTypes.iShowReservLine Then
                            'Show on separate line or bury in sold?
                            tlAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) + ilNo30
                            tlAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) + ilNo60
                        Else
                            tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                            tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                        End If
                    ElseIf slChfStatus = "H" Then                         'staus "Hold", always show on separate line
                        tlAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) + ilNo30
                        tlAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) + ilNo60
                    Else            'not held or reserved, put in sold
                        tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                        tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tlAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                    End If
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

End Sub
'
'               mProcessCntrChange - if any contracts have been modified,
'               make sure its processed incase the daypart isnt shown
'               <input> tlAvRdf() - Daypart (RDF) array
'                       llIgnoreCodes() - array of contracts that are modified
'               <input/output> - flag to indicate that spot is OK (true) to process, else false
'
Public Sub mProcessCntrChange(tlAvRdf() As RDF, llIgnoreCodes() As Long, ilSpotOK As Integer)
Dim ilLoop As Integer
Dim ilRdf As Integer

If ilSpotOK Then
    For ilLoop = 0 To UBound(llIgnoreCodes) - 1 Step 1
        If tmSdf.lChfCode = llIgnoreCodes(ilLoop) Then
            'the spot is from a contract that is modified.  If this DP isnt
            'to be shown on rept, go ahead and process it since it wont be
            'counted when it goes thru the proposals since its not to be
            'shown on the report (only DP to be shown on report are in table)
            For ilRdf = LBound(tlAvRdf) To UBound(tlAvRdf) - 1 Step 1
                If (tlAvRdf(ilRdf).iCode = tmClf.iRdfCode) Then  'only DP to be shown are in the table
                    ilSpotOK = False
                    Exit For
                End If
            Next ilRdf
        End If
        If Not ilSpotOK Then                'if already determine to ignore the spot, no need to test
                                            'any other matching contracts
            Exit For
        End If
    Next ilLoop
End If
End Sub
