Attribute VB_Name = "RptCRIR"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptcrir.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmChfSrchKey                  imSmfRecLen                   tmSmf                     *
'*                                                                                        *
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

Dim hmNif As Integer            'Network Inventory file handle
Dim tmNIf As NIF                'NIF options record image
Dim imNIFRecLen As Integer      'NIF  options record length
Dim tmNifArray() As NetInvByWeek         'array of NIF records
Dim tmNifSold() As NetInvByWeek         'array of NIF records, sold spots subtracted from counts
Dim tmNIFSrchKey1 As NIFKEY1     'NIF key 1 image

'
Dim hmGrf As Integer            'Generic report record
Dim tmGrf As GRF                'GRF record image
Dim imGrfRecLen As Integer      'GRF record length

Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey1 As CHFKEY1  'CHF key record image (contract #)
Dim imCHFRecLen As Integer        'CHF record length
Dim tmChf As CHF
Dim hmClf As Integer            'Contract line file handle
Dim tmClfSrchKey As CLFKEY0     'CLF record image
Dim imClfRecLen As Integer        'CLF record length
Dim tmClf As CLF
Dim hmCff As Integer            'Contract flight file handle
Dim imCffRecLen As Integer      'CFF record length
Dim tmCff As CFF
Dim hmSdf As Integer            'Spot detail file handle
Dim imSdfRecLen As Integer        'SDF record length
Dim tmSdf As SDF
Dim tmSdfSrchKey2 As SDFKEY2     'SDF record image (SDF code as keyfield)

Dim hmSmf As Integer            'Spot mg file handle

Dim hmAgf As Integer            'Agency file handle
Dim imAgfRecLen As Integer        'AGF record length
Dim tmAgf As AGF
Dim tmAgfSrchKey As INTKEY0

Dim imIncludeCodes As Integer            'true = include codes stored in ilusecode array,
                                            'false = exclude codes store din ilusecode array
Dim imUseCodes() As Integer       'valid  vehicles codes to process--
                                              'orvehicles codes not to process
Dim lmSingleContrCode As Long       'selective contract code
Dim lmStartOfYear As Long           'std broadcast year start date
'***************************************************************************
'*
'*      Procedure Name:gCRInvRevenue (Network/Station Spot Report)
'*
'*      7-14-05 Obtain all contracts on the books and determine how much
'*          network spots are booked vs station acquisition spots.  Acquisition
'           spots are determined by the line acquisition cost equal to non-zero.
'           Since the information is obtained from the ordered lines, go thru
'           the spots and find any cancelled spots for adjustments in billing
'           that may have been made (those spots that will not be made good).
'           Go thru new file to find out what the stations acquired inventory
'           counts should be for the year.  Some stations determine their
'           inventory by the week, some by the year.  If by the week, once
'           the week is over,  any unsold station inventory is lost.  If by
'           the month, all unused inv can be carrried over into future weeks/months.
'*
'****************************************************************************
Sub gCRInvRevenue()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLoop                        slDate                        llDate                    *
'*  ilTemp                                                                                *
'******************************************************************************************

'
    Dim ilRet As Integer
    Dim ilVehicle As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim llStart As Long                 'start date of requested period
    Dim llEnd As Long                   'end date of requested period
    Dim tlCntTypes As CNTTYPES          'values of contract types to include/exclude
    Dim slStart As String               'user start date requested
    Dim slEnd As String                 'user end date requested
    Dim slMonth As String
    Dim slYear As String
    Dim slDay As String
    Dim llSingleContr As Long           'single contract number user requested

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

    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmVef
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)

    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)

    hmAgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAgf, "", sgDBPath & "Agf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmAgf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imAgfRecLen = Len(tmAgf)

    hmNif = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmNif, "", sgDBPath & "Nif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmNif)
        ilRet = btrClose(hmAgf)
        ilRet = btrClose(hmSdf)
        ilRet = btrClose(hmVef)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        btrDestroy hmNif
        btrDestroy hmAgf
        btrDestroy hmSdf
        btrDestroy hmVef
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        Exit Sub
    End If
    imNIFRecLen = Len(tmNIf)

    tlCntTypes.iHold = gSetCheck(RptSelIR!ckcSelC1(0).Value)
    tlCntTypes.iOrder = gSetCheck(RptSelIR!ckcSelC1(1).Value)
    tlCntTypes.iStandard = gSetCheck(RptSelIR!ckcSelC1(2).Value)
    tlCntTypes.iReserv = gSetCheck(RptSelIR!ckcSelC1(3).Value)
    tlCntTypes.iRemnant = gSetCheck(RptSelIR!ckcSelC1(4).Value)
    tlCntTypes.iDR = gSetCheck(RptSelIR!ckcSelC1(5).Value)
    tlCntTypes.iPI = gSetCheck(RptSelIR!ckcSelC1(6).Value)
    tlCntTypes.iTrade = gSetCheck(RptSelIR!ckcSelC1(9).Value)

    tlCntTypes.iPSA = gSetCheck(RptSelIR!ckcSelC1(7).Value)
    tlCntTypes.iPromo = gSetCheck(RptSelIR!ckcSelC1(8).Value)
    'tlCntTypes.iCntrSpots = gSetCheck(RptSelIR!ckcCntrFeed(0).Value)
    If (tlCntTypes.iHold) Or (tlCntTypes.iOrder) Then
        tlCntTypes.iCntrSpots = True
    Else
        tlCntTypes.iCntrSpots = False
    End If
    tlCntTypes.iFeedSpots = gSetCheck(RptSelIR!ckcCntrFeed(1).Value)

'    slStart = RptSelIR!edcSelCFrom.Text
    slStart = RptSelIR!CSI_CalFrom.Text         '9-5-19 use csi calendar control vs edit box
    llStart = gDateValue(slStart)
    slStart = Format$(llStart, "m/d/yy")          'insure year appended

'    slEnd = RptSelIR!edcSelCFrom1.Text
    slEnd = RptSelIR!CSI_CalTo.Text
    llEnd = gDateValue(slEnd)
    slEnd = Format$(llEnd, "m/d/yy")          'insure year appended
    gObtainYearMonthDayStr slEnd, True, slYear, slMonth, slDay
    lmStartOfYear = gDateValue(gObtainStartStd("1/15/" & Trim$(slYear)))

    tmGrf.iGenDate(0) = igNowDate(0)        'todays date used for retrieval/removal of records
    tmGrf.iGenDate(1) = igNowDate(1)
    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
    tmGrf.lGenTime = lgNowTime

    'see if the user requested single contract exists
    llSingleContr = Val(RptSelIR!edcContract.Text)   'single contract entered
    lmSingleContrCode = -1
    If llSingleContr > 0 Then
        tmChfSrchKey1.lCntrNo = llSingleContr
        tmChfSrchKey1.iCntRevNo = 32000
        tmChfSrchKey1.iPropVer = 32000
        ilRet = btrGetGreaterOrEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
        If (ilRet = BTRV_ERR_NONE) And (tmChf.lCntrNo = llSingleContr) Then
            lmSingleContrCode = tmChf.lCode
        End If
    End If
    'build array of vehicles to include/exclude
    gObtainCodes RptSelIR!lbcSelection, tgCSVNameCode(), imIncludeCodes, imUseCodes(), RptSelIR

    mBuildNetInvbyVehicle Val(slYear)      'build original network inventory defined into memory


    mProcessInvCnts slStart, slEnd, tlCntTypes          'build spots sold & revenue by contract

    'loop thru the vehicles selected and find any cancelled spots for adjustments
    'since the spots sold & revenue is based on ordered
    'search SDF by Key1 (vehicle, date, time, sch status (only cancelled)
    For ilVehicle = 0 To RptSelIR!lbcSelection.ListCount - 1 Step 1
        If (RptSelIR!lbcSelection.Selected(ilVehicle)) Then
            slNameCode = tgCSVNameCode(ilVehicle).sKey
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            mBuildCancelled ilVefCode, slStart, slEnd
        End If
    Next ilVehicle                              'For ilvehicle = 0 To RptSelIR!lbcSelection(0).ListCount - 1


    'write out the records for the inventory
    mCreateInvAfterAdjusted llStart, llEnd


    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmAgf)
    ilRet = btrClose(hmNif)

    btrDestroy hmSdf
    btrDestroy hmVef
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmAgf
    btrDestroy hmNif

    Erase tmNifArray, tmNifSold, imUseCodes
    Exit Sub
End Sub
'
'
'           mBuildInvRevenue - gather all contracts by dates requested
'           Process each contract and accumulate $ and spots.
'
Private Sub mProcessInvCnts(slStart As String, slEnd As String, tlCntrTypes As CNTTYPES)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llNetRev                      slNet                                                   *
'******************************************************************************************

Dim tlChfAdvtExt() As CHFADVTEXT
Dim ilHOState As Integer
Dim llDate As Long
Dim llDate2 As Long
Dim slCntrStatus As String
Dim slCntrType As String
Dim ilRet As Integer
Dim ilCurrentRecd As Integer            'index of contract to process from tlChfadvtext array
Dim llContrCode As Long
Dim ilClf As Integer                    'line loop
'Dim llProject(1 To 54) As Long           'revenue $ for year by week
Dim llProject(0 To 54) As Long           'revenue $ for year by week. Index zero ignored
                                        'retains inventory by week)
'Dim llStartDates(1 To 2) As Long        'earliest start & latest dates to gather
Dim llStartDates(0 To 2) As Long        'earliest start & latest dates to gather. Index zero ignored
Dim ilFoundOption As Integer
'Dim llSpots(1 To 54) As Long            'spots booked by week
Dim llSpots(0 To 54) As Long            'spots booked by week. Index zero ignored
Dim ilTemp As Integer
Dim ilWhichRate As Integer      '0 = use actual line cost, 1 = use acquisition cost
Dim ilNet As Integer            'true if net , change to string, use slGrossOrNet
Dim slGrossOrNet As String * 1  'G = gross, N = net
Dim ilInclTrade As Integer          'true to include trades (any %)
Dim ilCorT As Integer           'loop count for cash trade processing
Dim ilStartCorT As Integer
Dim ilEndCorT As Integer
Dim slPctTrade As String        'pct of trade from header
Dim slCashAgyComm As String     'agy comm from agy
Dim llAdjustedNetRev As Long
Dim llAdjustedStationRev As Long
Dim slPortionRevenue As String
Dim slNetworkRev As String
Dim slStationRev As String
Dim llContrNetworkRev As Long
Dim llContrStationRev As Long
Dim ilVehicle As Integer

    llStartDates(1) = gDateValue(slStart)
    llStartDates(2) = gDateValue(slEnd) + 1

    'ilNet = True
    slGrossOrNet = "N"
    If RptSelIR!rbcSelC2(0).Value = True Then       'do gross
        'ilNet = False
        slGrossOrNet = "G"
    End If

    ilInclTrade = False                                 '
    If tlCntrTypes.iTrade = True Then     'include the trades (any %)
        ilInclTrade = True
    End If

    slCntrStatus = ""                 'statuses: hold, order, unsch hold, uns order
    If RptSelIR!ckcSelC1(0).Value = vbChecked Then     'incl holds and uns holds
        slCntrStatus = "HG"
    End If
    If RptSelIR!ckcSelC1(1).Value = vbChecked Then  'incl order and uns oeswe
        slCntrStatus = slCntrStatus & "ON"
    End If

    slCntrType = ""
    If RptSelIR!ckcSelC1(2).Value = vbChecked Then      'std
        slCntrType = "C"
    End If
    If RptSelIR!ckcSelC1(3).Value = vbChecked Then      'resv
        slCntrType = slCntrType & "V"
    End If
    If RptSelIR!ckcSelC1(4).Value = vbChecked Then      'remnant
        slCntrType = slCntrType & "T"
    End If
    If RptSelIR!ckcSelC1(5).Value = vbChecked Then      'DR
        slCntrType = slCntrType & "R"
    End If
    If RptSelIR!ckcSelC1(6).Value = vbChecked Then      'PI
        slCntrType = slCntrType & "Q"
    End If
    If RptSelIR!ckcSelC1(7).Value = vbChecked Then      'PSA
        slCntrType = slCntrType & "S"
    End If
    If RptSelIR!ckcSelC1(8).Value = vbChecked Then      'Promo
        slCntrType = slCntrType & "M"
    End If
    If slCntrType = "CVTRQSM" Then          'all types: PI, DR, etc.  except PSA(p) and Promo(m)
        slCntrType = ""                     'blank out string for "All"
    End If
    ilHOState = 2                       'get latest orders & revisions  (HOGN plus any revised orders WCI)

    ilRet = gObtainCntrForDate(RptSelIR, slStart, slEnd, slCntrStatus, slCntrType, ilHOState, tlChfAdvtExt())

    'All contracts have been retrieved for the requested period
    For ilCurrentRecd = LBound(tlChfAdvtExt) To UBound(tlChfAdvtExt) - 1 Step 1
        'Retrieve the contract, schedule lines and flights
        llContrCode = tlChfAdvtExt(ilCurrentRecd).lCode
        ilRet = gObtainCntr(hmCHF, hmClf, hmCff, llContrCode, False, tgChf, tgClf(), tgCff())

        If Not ilRet Then
            On Error GoTo mProcessInvCntErr
            gBtrvErrorMsg ilRet, "mProcessInvCntErr :" & "Chf.Btr", RptSelIR
            On Error GoTo 0
        End If

        'include 100% cash or partial trades
        If ((tgChf.iPctTrade = 0) Or (ilInclTrade And tgChf.iPctTrade > 0)) And (lmSingleContrCode < 0 Or lmSingleContrCode = tgChf.lCode) Then
            'determine if the contracts start & end dates fall within the requested period
            gUnpackDateLong tgChf.iEndDate(0), tgChf.iEndDate(1), llDate2      'hdr end date converted to long
            gUnpackDateLong tgChf.iStartDate(0), tgChf.iStartDate(1), llDate    'hdr start date converted to long
            If llDate < llStartDates(2) And llDate2 >= llStartDates(1) Then       'does requested dates span the contract?
                'tmGrf.lDollars(1) = total spots booked (net & station)
                'tmGrf.lDollars(2) = net spots booked
                'tmgrf.ldollars(3) = station spots booked
                'tmgrf.ldollars(4) = Network inventory from NIF
                'tmgrf.ldollars(6) = net revenue
                'tmgrf.ldollars(5) = station revenue
                'tmgrf.lDollars(7) = network spots booked in seconds per line
                'tmgrf.lDollars(8) = station spots booked in seconds per line
                'tmgrf.lDollars(9) = original inventory from NIF (no adjustments)
                'tmgrf.ivefcode = vehicle from line
                'tmgrf.ichfcode = contract code
                'tmgrf.sBktType = T = 100%trade, t = partial trade, blank = 100% cash
                'tmgrf.sDateType = Y = yearly inventory from NIF, W = weekly inventory from NIF
                'tmgrf.icode2 = 0 = all contract spots, 1 = inventory from NIF or adjusted
                For ilClf = LBound(tgClf) To UBound(tgClf) - 1 Step 1
                    tmClf = tgClf(ilClf).ClfRec
                    ilFoundOption = mTestIncludeVehicle(tmClf.iVefCode)

                    If ilFoundOption Then
                        'Project the  spots & revenue from the flights
                        If tmClf.sType = "S" Or tmClf.sType = "H" Then      'look at only std or hidden lines (vs package lines)
                            'accumulate spots & revenue
                            ilWhichRate = 0                     'assume net
                            If tmClf.lAcquisitionCost <> 0 Then
                                ilWhichRate = 1
                            End If

                            'init record for new contract
                            For ilTemp = 1 To 8
                                tmGrf.lDollars(ilTemp - 1) = 0
                            Next ilTemp
                            llContrNetworkRev = 0
                            llContrStationRev = 0

                            For ilTemp = 1 To 54
                                llSpots(ilTemp) = 0
                                llProject(ilTemp) = 0
                            Next ilTemp

                            gBuildFlightSpotsAndRevenue ilClf, llStartDates(), 1, 2, llProject(), llSpots(), 2, 2, tgClf(), tgCff(), slGrossOrNet

                            'gather the accumulated $ and spots for the record to be written
                            'calculate net if requestesd
                            If ilWhichRate = 0 Then             'use actual rate
                                For ilTemp = 1 To 53
                                    'tmGrf.lDollars(2) = tmGrf.lDollars(2) + llSpots(ilTemp)      'net spots
                                    tmGrf.lDollars(1) = tmGrf.lDollars(1) + llSpots(ilTemp)      'net spots
                                    'tmGrf.lDollars(7) = tmGrf.lDollars(7) + (llSpots(ilTemp) * tmClf.iLen)      'total seconds booked
                                    tmGrf.lDollars(6) = tmGrf.lDollars(6) + (llSpots(ilTemp) * tmClf.iLen)      'total seconds booked
                                    llContrNetworkRev = llContrNetworkRev + llProject(ilTemp)     'net rev
                                Next ilTemp
                            Else            'use acquisition cost
                                For ilTemp = 1 To 53
                                    'tmGrf.lDollars(3) = tmGrf.lDollars(3) + llSpots(ilTemp)      'station spots
                                    tmGrf.lDollars(2) = tmGrf.lDollars(2) + llSpots(ilTemp)      'station spots
                                    llContrStationRev = llContrStationRev + llProject(ilTemp)    'station rev
                                    'tmGrf.lDollars(8) = tmGrf.lDollars(8) + (llSpots(ilTemp) * tmClf.iLen)      'total seconds booked
                                    tmGrf.lDollars(7) = tmGrf.lDollars(7) + (llSpots(ilTemp) * tmClf.iLen)      'total seconds booked

                                Next ilTemp
                           End If

                           'lines completed building spot counts and $, create record to print
                            'tmGrf.lDollars(1) = tmGrf.lDollars(2) + tmGrf.lDollars(3)   'total net & station inv.
                            tmGrf.lDollars(0) = tmGrf.lDollars(1) + tmGrf.lDollars(2)   'total net & station inv.
                            'find the agency to determine if commissionable, or if gross no comm applicable
                            'If (tgChf.iAgfCode = 0) Or (ilNet = False) Then
                            If (tgChf.iAgfCode = 0) Or (slGrossOrNet = "G") Then  'direct or Gross requested
                                slCashAgyComm = ".00"
                            Else
                                If tgChf.iAgfCode > 0 Then
                                    tmAgfSrchKey.iCode = tgChf.iAgfCode
                                    ilRet = btrGetEqual(hmAgf, tmAgf, imAgfRecLen, tmAgfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'get matching agy recd
                                    If ilRet = BTRV_ERR_NONE Then
                                        slCashAgyComm = gIntToStrDec(tmAgf.iComm, 2)
                                    End If          'ilret = btrv_err_none
                                Else
                                    slCashAgyComm = ".00"
                                End If              'iagfcode > 0
                            End If

                            If tgChf.iPctTrade = 100 Then
                                ilStartCorT = 2
                                ilEndCorT = 2
                                tmGrf.sBktType = "T"        '100% trade
                            ElseIf tgChf.iPctTrade > 0 Then
                                ilStartCorT = 1
                                ilEndCorT = 2
                                tmGrf.sBktType = "t"        'part cash/trade
                            Else
                                ilStartCorT = 1
                                ilEndCorT = 1
                                tmGrf.sBktType = ""         '100% cash
                            End If

                            For ilCorT = ilStartCorT To ilEndCorT                'if part cash/trade, need to split the $ and determine if the trade is commissionable
                                If ilCorT = 1 Then
                                    slPctTrade = gSubStr("100.", str$(tgChf.iPctTrade))  'cash portion
                                Else
                                    slPctTrade = str$(tgChf.iPctTrade)         'trade portion
                                    'see if trades are commissionable
                                    If tgChf.sAgyCTrade = "N" Then
                                        slCashAgyComm = ".00"
                                    End If
                                End If

                                'network revenue, do string math
                                slNetworkRev = gLongToStrDec(llContrNetworkRev, 2)       'get the network share
                                slPortionRevenue = gDivStr(gMulStr(slNetworkRev, slPctTrade), "100")              'slsp gross
                                slPortionRevenue = gMulStr(slPortionRevenue, gSubStr("100.00", slCashAgyComm))
                                llAdjustedNetRev = gRoundStr(slPortionRevenue, "01.", 0)
                                'tmGrf.lDollars(6) = tmGrf.lDollars(6) + llAdjustedNetRev
                                tmGrf.lDollars(5) = tmGrf.lDollars(5) + llAdjustedNetRev

                                'station revenue, do string math
                                '1-14-09 Acquisition costs are already net amounts, do not net down again
                                slStationRev = gLongToStrDec(llContrStationRev, 2)       'get the station share
                                slPortionRevenue = gDivStr(gMulStr(slStationRev, slPctTrade), "100")              'slsp gross
                                'slPortionRevenue = gMulStr(slPortionRevenue, gSubStr("100.00", slCashAgyComm))
                                slPortionRevenue = gMulStr(slPortionRevenue, gSubStr("100.00", ".00"))
                                llAdjustedStationRev = Val(gRoundStr(slPortionRevenue, "01.", 2))
                                'tmGrf.lDollars(5) = tmGrf.lDollars(5) + llAdjustedStationRev
                                tmGrf.lDollars(4) = tmGrf.lDollars(4) + llAdjustedStationRev

                            Next ilCorT
                            'If tmGrf.lDollars(1) <> 0 Then
                            If tmGrf.lDollars(0) <> 0 Then
                                'subtract the spot length from the original inventory
                                For ilVehicle = LBound(tmNifSold) To UBound(tmNifSold) - 1
                                    If tmNifSold(ilVehicle).iVefCode = tmClf.iVefCode Then
                                        tmGrf.sDateType = tmNifArray(ilVehicle).sInvWkYear
                                        'tmGrf.lDollars(9) = tmNifArray(ilVehicle).lTotalYear

                                        For ilTemp = 1 To 53
                                            tmNifSold(ilVehicle).lInvCount(ilTemp - 1) = tmNifSold(ilVehicle).lInvCount(ilTemp - 1) + (tmClf.iLen * llSpots(ilTemp))
                                            tmNifSold(ilVehicle).lTotalYear = tmNifSold(ilVehicle).lTotalYear + (tmClf.iLen * llSpots(ilTemp))        'total time sold
                                        Next ilTemp
                                        Exit For
                                    End If
                                Next ilVehicle
                                tmGrf.lChfCode = tgChf.lCode
                                tmGrf.iVefCode = tmClf.iVefCode
                                tmGrf.iCode2 = 0                'sort order, spot info followed by inventory info
                                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)

                            End If

                        End If
                    End If
                Next ilClf                      'loop thru schedule lines


            End If                       'llDate < llStartDates(2) And llDate2 >= llStartDates(1)
        End If                           'include cash and trades
    Next ilCurrentRecd

    Erase tlChfAdvtExt, llProject, llSpots
    sgCntrForDateStamp = ""
    Exit Sub

mProcessInvCntErr:

End Sub
'           mBuildCancelled - search SDF (KEY1) for any cancelled spots
'           to adjust the revenue & spots booked.  All information is
'           obtained from the Contract file as ordered, and cancelled
'           spots must be subtracted out for more accurate results.
'
'           <input> ilVefCode - vehicle code to search
'                   slStart - start date to search for cancelled spots
'                   slEnd - end date to search for cancelled spots
'                   slGrossOrNet - G = Gross , N = Net (default to Net).  USed to acquisition costs computation if using Acq commissions
'       1-13-06 dh wrong key used for cancel spots search
Public Sub mBuildCancelled(ilVefCode As Integer, slStart As String, slEnd As String, Optional slGrossOrNet As String = "N")

Dim llEndDate As Long
Dim llStartDate As Long
Dim llSpotDate As Long
Dim ilRet As Integer
Dim ilVehicle As Integer
Dim ilTemp As Integer

Dim ilAcqCommPct As Integer
Dim ilAcqLoInx As Integer
Dim ilAcqHiInx As Integer
Dim llAcqNet As Long
Dim llAcqComm As Long
Dim blAcqOK As Boolean

    llEndDate = gDateValue(slEnd)
    llStartDate = gDateValue(slStart)

    tmSdfSrchKey2.iVefCode = ilVefCode
    'gPackDate slStart, tmSdfSrchKey2.iDate(0), tmSdfSrchKey2.iDate(1)
    tmSdfSrchKey2.iDate(0) = 0
    tmSdfSrchKey2.iDate(1) = 0
    tmSdfSrchKey2.iTime(0) = 0
    tmSdfSrchKey2.iTime(1) = 0
    tmSdfSrchKey2.sSchStatus = "C"   'Cancelled only
    tmSdfSrchKey2.iAdfCode = 0      'all advt
    imSdfRecLen = Len(tmSdf)

    ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
    If (ilRet <> BTRV_ERR_END_OF_FILE) Then
        gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llSpotDate
        Do While tmSdf.iVefCode = ilVefCode And tmSdf.sSchStatus = "C"
            If ((tmSdf.lChfCode = lmSingleContrCode Or lmSingleContrCode < 0) And (llSpotDate >= llStartDate And llSpotDate <= llEndDate)) Then
                'access the line to determine the cost
                If tmSdf.lChfCode <> tmClf.lChfCode Or tmSdf.iLineNo <> tmClf.iLine Then         'only read line when necessary
                    tmClfSrchKey.lChfCode = tmSdf.lChfCode
                    tmClfSrchKey.iLine = tmSdf.iLineNo
                    tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
                    tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
                    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
                If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And tmSdf.sSchStatus = "C" Then
                     'init record for new contract
                    For ilTemp = 1 To 8
                        tmGrf.lDollars(ilTemp - 1) = 0
                    Next ilTemp
                     If tmClf.lAcquisitionCost = 0 Then             'use actual rate
                        'need actual rate of spot
                        If gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tmCff) Then
                            'tmGrf.lDollars(2) = tmGrf.lDollars(2) - 1      'net spots
                            tmGrf.lDollars(1) = tmGrf.lDollars(1) - 1      'net spots
                            'tmGrf.lDollars(6) = tmGrf.lDollars(6) - tmCff.lActPrice  'net rev
                            tmGrf.lDollars(5) = tmGrf.lDollars(5) - tmCff.lActPrice  'net rev
                        End If
                     Else            'use acquisition cost
                        'tmGrf.lDollars(3) = tmGrf.lDollars(2) - 1      'station spots
                        tmGrf.lDollars(2) = tmGrf.lDollars(1) - 1      'station spots
                        If (Asc(tgSaf(0).sFeatures2) And ACQUISITIONCOMMISSIONABLE) = ACQUISITIONCOMMISSIONABLE Then
                            If slGrossOrNet = "N" Then
                                ilAcqCommPct = 0
                                blAcqOK = gGetAcqCommInfoByVehicle(tmClf.iVefCode, ilAcqLoInx, ilAcqHiInx)
                                ilAcqCommPct = gGetEffectiveAcqComm(llSpotDate, ilAcqLoInx, ilAcqHiInx)
                                gCalcAcqComm ilAcqCommPct, tmClf.lAcquisitionCost, llAcqNet, llAcqComm
                                'tmGrf.lDollars(5) = tmGrf.lDollars(6) - llAcqNet   'station rev with acq commission
                                tmGrf.lDollars(4) = tmGrf.lDollars(5) - llAcqNet   'station rev with acq commission
                            Else            'gross acquisitions
                                 'tmGrf.lDollars(5) = tmGrf.lDollars(6) - tmClf.lAcquisitionCost   'station rev
                                 tmGrf.lDollars(4) = tmGrf.lDollars(5) - tmClf.lAcquisitionCost   'station rev
                            End If
                        Else
                            'tmGrf.lDollars(5) = tmGrf.lDollars(6) - tmClf.lAcquisitionCost   'station rev
                            tmGrf.lDollars(4) = tmGrf.lDollars(5) - tmClf.lAcquisitionCost   'station rev
                        End If
                        ''tmGrf.lDollars(5) = tmGrf.lDollars(6) - tmClf.lAcquisitionCost   'station rev
                    End If
                    'subtract the spot length from the original inventory
                    For ilVehicle = LBound(tmNifSold) To UBound(tmNifSold) - 1
                        If tmNifSold(ilVehicle).iVefCode = tmSdf.iVefCode Then
                            tmGrf.sDateType = tmNifArray(ilVehicle).sInvWkYear
                            'tmGrf.lDollars(9) = tmNifArray(ilVehicle).lTotalYear            'inv by year or week, original value

                            For ilTemp = 1 To 53
                                tmNifSold(ilVehicle).lInvCount(ilTemp - 1) = tmNifSold(ilVehicle).lInvCount(ilTemp - 1) - tmSdf.iLen    'weekly count; give time back, it has been cancelled
                                tmNifSold(ilVehicle).lTotalYear = tmNifSold(ilVehicle).lTotalYear - tmSdf.iLen                      'yearly count
                            Next ilTemp
                            Exit For
                        End If
                    Next ilVehicle

                    tmGrf.lChfCode = tmSdf.lChfCode
                    tmGrf.iVefCode = tmSdf.iVefCode
                    'tmGrf.lDollars(1) = tmGrf.lDollars(2) + tmGrf.lDollars(3)
                    tmGrf.lDollars(0) = tmGrf.lDollars(1) + tmGrf.lDollars(2)
                    tmGrf.iCode2 = 0                'sort order, spot info followed by inventory info
                    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)

                End If
            End If
            ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    End If
    Exit Sub
End Sub
'
'                   mBuildNetInvbyVehicle - obtain all the vehicles inventory stored in
'                   NIF and build into memory into array tmNIFArray
'           tmNifArray is the original information defined with the vehicle
'           tmNifSold is the sold time gathered from lines (in seconds) for each vehicle.
'           tmNifSold is required if Inventory by week is defined and there is
'           no rollover of unused inventory in the past.
'
Private Sub mBuildNetInvbyVehicle(ilYear As Integer)
Dim ilVehicle As Integer
Dim slNameCode As String
Dim slName As String
Dim slCode As String
ReDim tmNifArray(0 To 0) As NetInvByWeek        'original inventory defined
ReDim tmNifSold(0 To 0) As NetInvByWeek         'sold counts
Dim ilUpper As Integer
Dim ilRet As Integer
Dim ilWeek As Integer

    If Not RptSelIR!rbcSortBy(2).Value = True Then           'sort by advt or cnt, inventory not applicable since its stored by vehicle
        Exit Sub
    End If
    ilUpper = 0
    For ilVehicle = 0 To RptSelIR!lbcSelection.ListCount - 1 Step 1
        If (RptSelIR!lbcSelection.Selected(ilVehicle)) Then
            slNameCode = tgCSVNameCode(ilVehicle).sKey
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)

            tmNIFSrchKey1.iVefCode = Val(slCode)
            tmNIFSrchKey1.iYear = ilYear
            ilRet = btrGetEqual(hmNif, tmNIf, imNIFRecLen, tmNIFSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
            If ilRet = BTRV_ERR_NONE Then
                tmNifArray(ilUpper).iVefCode = Val(slCode)
                tmNifArray(ilUpper).iYear = ilYear
                tmNifArray(ilUpper).sAllowRollover = tmNIf.sAllowRollover
                tmNifArray(ilUpper).sInvWkYear = tmNIf.sInvWkYear

                tmNifSold(ilUpper).iVefCode = Val(slCode)
                tmNifSold(ilUpper).iYear = ilYear
                tmNifSold(ilUpper).sAllowRollover = tmNIf.sAllowRollover
                tmNifSold(ilUpper).sInvWkYear = tmNIf.sInvWkYear

                For ilWeek = 1 To 53
                    'tmNifArray is for the original inventory defined
                    If tmNifArray(ilUpper).sInvWkYear = "Y" Then        'if yearly, the years totals have been
                                                                        'distributed across all weeks in the year
                        tmNifArray(ilUpper).lTotalYear = tmNifArray(ilUpper).lTotalYear + tmNIf.lInvCount(ilWeek - 1)
                    Else
                        'tmNifArray(ilUpper).lTotalYear = tmNifArray(ilUpper).lInvCount(1)     'minutes per week
                        tmNifArray(ilUpper).lTotalYear = tmNifArray(ilUpper).lInvCount(0)     'minutes per week
                    End If
                    tmNifArray(ilUpper).lInvCount(ilWeek - 1) = tmNIf.lInvCount(ilWeek - 1)
                    'tmNifSold is the array used to keep track of whats been sold
                    tmNifSold(ilUpper).lTotalYear = 0
                    tmNifSold(ilUpper).lInvCount(ilWeek - 1) = 0
                Next ilWeek
                ilUpper = ilUpper + 1
                ReDim Preserve tmNifArray(0 To ilUpper) As NetInvByWeek
                ReDim Preserve tmNifSold(0 To ilUpper) As NetInvByWeek
            End If
        End If
    Next ilVehicle                              'For ilvehicle = 0 To RptSelIR!lbcSelection(0).ListCount - 1
    Exit Sub
End Sub
'
'
'                   mTestIncludeVehicle - test the vehicle to
'                   see if it should be included based on user selection
'
'           <input> ilvehicle - vehicle code to test for inclusion/exclusion
'            return - true if include, else false to exclude
Private Function mTestIncludeVehicle(ilVehicle As Integer) As Integer
Dim ilTemp As Integer
Dim ilFoundOption As Integer

    If imIncludeCodes Then          'include the any of the codes in array?
        ilFoundOption = False
        For ilTemp = LBound(imUseCodes) To UBound(imUseCodes) - 1 Step 1
            If imUseCodes(ilTemp) = ilVehicle Then
                ilFoundOption = True                    'include the matching vehicle
                Exit For
            End If

        Next ilTemp
    Else                            'exclude any of the codes in array?
        ilFoundOption = True
        For ilTemp = LBound(imUseCodes) To UBound(imUseCodes) - 1 Step 1
            If imUseCodes(ilTemp) = ilVehicle Then
                ilFoundOption = False                  'exclude the matching vehicle
                Exit For
            End If
        Next ilTemp
    End If
    mTestIncludeVehicle = ilFoundOption
    Exit Function
End Function
'
'           mCreateInvAfterAdjusted - create the network inventory
'           data from the NIF (information stored with the vehicle).
'           Determine if inventory is stored by the year, all unused inventory
'           is available for the entire year.
'           Determine if Inventory is stored by week and unsold inventory can be
'           used during the year;
'           Determine if Inventory is stored by week and unsold inventory is lost
'           if in the past.
'           Create records in GRF to define the inventory defined to calculate %
'           of network inventory sold todate.
'           <input> llStart - start date requested
'                   llend - end date requested
Private Sub mCreateInvAfterAdjusted(llStart As Long, llEnd As Long)
Dim ilVehicle As Integer
Dim llTodayDate As Long         'todays date
Dim ilRet As Integer
Dim llLastLogDate As Long       'last log date for vehicle procesing

Dim ilStartWeek As Integer          'start week of requested period for the past
Dim ilEndWeek As Integer            'end week of requested period for the past
Dim llLatestDate As Long            'today date or last log date, whichever is greater
Dim ilTemp As Integer
Dim llTotalWeeksInv As Long         'total weeks inventory after the past has been adjusted to be soldout
Dim ilCurrentWeekInx As Integer

    gUnpackDateLong igNowDate(0), igNowDate(1), llTodayDate

    For ilVehicle = LBound(tmNifArray) To UBound(tmNifArray) - 1
        tmGrf.sDateType = tmNifArray(ilVehicle).sInvWkYear
        'tmGrf.lDollars(9) = tmNifArray(ilVehicle).lTotalYear            'inv by year or week, original value
        tmGrf.lDollars(8) = tmNifArray(ilVehicle).lTotalYear            'inv by year or week, original value

        For ilTemp = 1 To 8                         'init the count arrays
            tmGrf.lDollars(ilTemp - 1) = 0
        Next ilTemp
        'if the Inventory is by week, check to see if the unused avails in the past can be rolled over.  If not,
        'that week is assumed to be 100% sold out.


        'If tmNifArray(ilVehicle).sAllowRollover = "Y" Then          'does this vehicle allow rollover of unused inventory within the year?
            tmGrf.lChfCode = 0
            tmGrf.iVefCode = tmNifArray(ilVehicle).iVefCode
        '    tmGrf.lDollars(4) = tmNifArray(ilVehicle).lTotalYear        'original iventory
        '    ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        'Else

            'Determine first week as current- todays date or last log date, which ever is greater
            llLatestDate = llTodayDate
            ilRet = -1
            ilRet = gBinarySearchVpf(tmNifArray(ilVehicle).iVefCode)
            If ilRet <> -1 Then         'no vpf options table found, use todays date as latest
                gUnpackDateLong tgVpf(ilRet).iLLD(0), tgVpf(ilRet).iLLD(1), llLastLogDate
                If llLatestDate < llLastLogDate Then
                    llLatestDate = llLastLogDate
                End If
            End If
            'no rollover allowed, must be by week (year is mandatory rollover
            ilStartWeek = (llStart - lmStartOfYear) \ 7 + 1
            ilEndWeek = (llEnd - lmStartOfYear) \ 7 + 1

            ilCurrentWeekInx = (llLatestDate - lmStartOfYear) \ 7 + 1
            If gWeekDayLong(llLatestDate) = 6 Then     'if this date is a sunday, the week is in the past-nothing more to sell
                ilCurrentWeekInx = ilCurrentWeekInx + 1
            End If
            'loop thru the weeks and consider the past sold out if rollover disallowed, and
            'Make the original inventory the same as the sold, unless
            ' the sold exceed the original inv--so that it can show
            'over 100% sellout
            For ilTemp = 1 To 53
                If ilTemp > ilEndWeek Then          'for future weeks, use unsold avails
                    'past the requested date, zero the inventory which has been averaged across the year
                    tmNifArray(ilVehicle).lInvCount(ilTemp - 1) = 0
                Else            'make sold out if in the past
                    If ilTemp < ilStartWeek Then            'user did not start report at beginning of year
                        'for the period between start of bdcst year and user start date: zero the inventory
                        'since its not applicable
                        tmNifArray(ilVehicle).lInvCount(ilTemp - 1) = 0
                    Else      'for all weeks from user requsted start date thru the current week, assume sold out
                              'if inventory remaining, make it the same as whats already been sold;
                              'if inventory less than whats sold, it has been oversold; leave alone
                        If tmNifArray(ilVehicle).sAllowRollover = "Y" Then          'does this vehicle allow rollover of unused inventory within the year?
                            'do nothing
                        Else                'rollover disallowed, consider past as 100% sold out
                            'first check to see if it this week is in the past
                            If ilTemp < ilCurrentWeekInx Then       'in the past
                                If tmNifSold(ilVehicle).lInvCount(ilTemp - 1) < tmNifArray(ilVehicle).lInvCount(ilTemp - 1) Then
                                    tmNifArray(ilVehicle).lInvCount(ilTemp - 1) = tmNifSold(ilVehicle).lInvCount(ilTemp - 1)
                                End If
                            End If
                        End If
                    End If
                End If
            Next ilTemp

            'weeks in the past adjusted, now add up all the inventory so that % network used can be calculated
            llTotalWeeksInv = 0
            For ilTemp = 1 To 53
                llTotalWeeksInv = llTotalWeeksInv + tmNifArray(ilVehicle).lInvCount(ilTemp - 1)
            Next ilTemp
            tmGrf.lChfCode = 0
            tmGrf.iVefCode = tmNifArray(ilVehicle).iVefCode
            'tmGrf.lDollars(4) = llTotalWeeksInv         'adjusted inventory: past ununsed inv is lost
            tmGrf.lDollars(3) = llTotalWeeksInv         'adjusted inventory: past ununsed inv is lost
            tmGrf.iCode2 = 1                'sort order, spot info followed by inventory info
            ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
        'End If
     Next ilVehicle
     Exit Sub
End Sub
