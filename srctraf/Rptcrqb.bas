Attribute VB_Name = "RPTCRQB"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Rptcrqb.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmSofSrchKey                                                                          *
'******************************************************************************************

Option Explicit
Option Compare Text
Dim hmVef As Integer            'Vehicle file handle
Dim tmVef As VEF                'VEF record image
Dim imVefRecLen As Integer         'VEF record length
Dim hmVsf As Integer            'Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
'Quarterly Avails
Dim hmAvr As Integer            'Quarterly Avails file handle
Dim tmAvr() As AVR                'AVR record image
Dim imAvrRecLen As Integer        'AVR record length
Dim tmNamedAvails() As Integer           'list of named avail code
Type KEYCHAR                    'method to test compare structures
    sChar As String * 63        'length of BOOKKEY
End Type
Type BOOKKEY
    iRdfCode As Integer         'dp code
    lChfCode As Long          'contr #
    lFsfCode As Long            'feed code
    lRate As Long               'spot rate
    sSpotType As String * 1     'tmSdf.sSpotType
    sPriceType As String * 1    'tmsdf.sPricetype
    sDysTms As String * 40      'DP days & times
    ivefSellCode As Integer     'vehicle code (selling vehicle)
    iLen As Integer             'spot length
    sAirMissed As String * 1     'A=Aired, M = Missed
    iPkgFlag As Integer          'package line ID (else 0)
End Type
Type BOOKSPOTS
    BkKeyRec As BOOKKEY                  'key to each unique record
    'iSpotCounts(1 To 13) As Integer     '1 quarters spot counts by week
    iSpotCounts(0 To 13) As Integer     '1 quarters spot counts by week. Index zero ignored
End Type
Dim tmBooked() As BOOKSPOTS
'Dim lmSAvailsDates(1 To 13) As Long   'Start Dates of avail week
Dim lmSAvailsDates(0 To 13) As Long   'Start Dates of avail week. Index zero ignored
'Dim lmEAvailsDates(1 To 13) As Long   'End dates of avail week
Dim lmEAvailsDates(0 To 13) As Long   'End dates of avail week. Index zero ignored
Dim hmCHF As Integer            'Contract header file handle
Dim tmChfSrchKey As LONGKEY0            'CHF record image
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
'Log Calendar
Dim hmLcf As Integer            'Log Calendar file handle
Dim tmLcf As LCF                'LCF record image
Dim imLcfRecLen As Integer        'LCF record length
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim tmRcf As RCF

Dim tmFsf As FSF
Dim hmFsf As Integer            'Feed Spots  file handle
Dim imFsfRecLen As Integer      'FSF record length

Dim tmAnf As ANF
Dim hmAnf As Integer            'Feed Spots  file handle
Dim tmAnfSrchKey As INTKEY0     'ANF record image
Dim imAnfRecLen As Integer      'ANF record length

Dim tmSof As SOF
Dim hmSof As Integer            'Sales Office  file handle
Dim imSofRecLen As Integer      'SOF record length

Dim tmSlf As SLF
Dim hmSlf As Integer            'Slsp file handle
Dim tmSlfSrchKey As INTKEY0     'SLF record image
Dim imSlfRecLen As Integer      'SLF record length

Dim tmSofList() As SOFLIST

Dim imRcf As Integer            'selected rate card index
Dim sm30sOrUnits As String * 1   '6-14-19 3 = 30" units, U =unit count (spot count)
Dim imGrossNet As Integer       '6-14-19 0 = gross, 1= net
Dim imVefList() As Integer       'List of Vehicle Codes
Dim tmPLPcf() As PCFTYPESORT    'TTP 10729 - Quarterly Booked Spots report: add digital lines

'***************************************************************************
'*
'*      Procedure Name:gCRQtrlyBookSpots
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
'*      12/13/00 Implement option to use start qtr or start date entered
'*      7-22-04 Implement include/exclude contract/feed spots
'*      4-3-06 For Missed spots, test ordered DP against the DP that its
'*              trying to put the missed spot into.  Test Book Into field.
'****************************************************************************
Sub gCRQtrlyBookSpots()
'
    Dim illoop As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim ilVehicle As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
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
    Dim slEffDate As String                     'user date entered
    Dim llEffDate As Long
    Dim ilNoQtrs As Integer                     'user # qtrs requested
    Dim tlCntTypes As CNTTYPES
    Dim ilSaveSort As Integer                   'DP or RIF field:  sort code
    Dim slSaveReport As String                  'DP or RIF field:  Save on report
    Dim tlAvr As AVR
    Dim ilStdQtr As Integer                     'true if using start qtr, else use start date entered
    Dim ilTemp As Integer
    Dim ilWksEntered As Integer                 '4-27-12
    Dim ilHowManyWks As Integer
    Dim slStartDate As String
    Dim slEndDate As String
    Dim llTempStartDate As Long
    Dim llTempEndDate As Long
    Dim ilSAvailsDates(0 To 1) As Integer
    Dim tlTranTypes As TRANTYPES
    ReDim tlRvf(0 To 0) As RVF
    Dim llAmount As Long
    Dim ilHowManyDays As Integer
    Dim llMonthlyAmount As Long
    
    'TTP 10674 - Get Last Billed Date
    Dim llLastBilledStd  As Long
    Dim llLastBilledCal  As Long
    gUnpackDate tgSpf.iBLastStdMnth(0), tgSpf.iBLastStdMnth(1), slDate 'convert last bdcst billing date to string
    llLastBilledStd = gDateValue(slDate)            'convert last month billed to long
    gUnpackDate tgSpf.iBLastCalMnth(0), tgSpf.iBLastCalMnth(1), slDate 'convert last bdcst billing date to string
    llLastBilledCal = gDateValue(slDate)            'convert last month billed to long
    
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

    hmSof = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSof, "", sgDBPath & "Sof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAvr)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSof
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
    imSofRecLen = Len(tmSof)

    hmSlf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSlf, "", sgDBPath & "Slf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmSlf)
        ilRet = btrClose(hmSof)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmAvr)
        ilRet = btrClose(hmCHF)
        btrDestroy hmSlf
        btrDestroy hmSof
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
    imSlfRecLen = Len(tmSlf)

    tlCntTypes.iHold = gSetCheck(RptSelQB!ckcSelC1(0).Value)
    tlCntTypes.iOrder = gSetCheck(RptSelQB!ckcSelC1(1).Value)
    tlCntTypes.iStandard = gSetCheck(RptSelQB!ckcSelC1(2).Value)
    tlCntTypes.iReserv = gSetCheck(RptSelQB!ckcSelC1(3).Value)
    tlCntTypes.iRemnant = gSetCheck(RptSelQB!ckcSelC1(4).Value)
    tlCntTypes.iDR = gSetCheck(RptSelQB!ckcSelC1(5).Value)
    tlCntTypes.iPI = gSetCheck(RptSelQB!ckcSelC1(6).Value)
    tlCntTypes.iTrade = gSetCheck(RptSelQB!ckcSelC1(9).Value)
    tlCntTypes.iMissed = gSetCheck(RptSelQB!ckcSelC1(10).Value)
    tlCntTypes.iNC = gSetCheck(RptSelQB!ckcSelC1(11).Value)
    tlCntTypes.iXtra = gSetCheck(RptSelQB!ckcSelC1(12).Value)
    tlCntTypes.iPSA = gSetCheck(RptSelQB!ckcSelC1(7).Value)
    tlCntTypes.iPromo = gSetCheck(RptSelQB!ckcSelC1(8).Value)
    If (tlCntTypes.iHold) Or (tlCntTypes.iOrder) Then
        tlCntTypes.iCntrSpots = True
    Else
        tlCntTypes.iCntrSpots = False
    End If
    tlCntTypes.iFeedSpots = gSetCheck(RptSelQB!ckcCntrFeed(1).Value)

    If RptSelQB!rbc30sOrUnits(0).Value Then              '30s units
        sm30sOrUnits = "3"
    Else
        sm30sOrUnits = "U"                               'unit (spot) counts
    End If
    
    If RptSelQB!rbcGrossNet(0).Value Then           'gross
        imGrossNet = 0
    Else
        imGrossNet = 1
    End If
    
    ilRet = gObtainRcfRifRdf()          'get the rate cards and assoc dayparts


    ilStdQtr = True             '12-13-00
    If RptSelQB!rbcStart(1).Value Then          'use start date entered (dont do the standard qtr)
        ilStdQtr = False
    End If

    'build array of selling office codes and their sales sources.  Need to test selectivity of sales
    'sources against the spots
    ilTemp = 0
    ilRet = btrGetFirst(hmSof, tmSof, imSofRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
    Do While ilRet = BTRV_ERR_NONE
        ReDim Preserve tmSofList(0 To ilTemp) As SOFLIST
        tmSofList(ilTemp).iSofCode = tmSof.iCode
        tmSofList(ilTemp).iMnfSSCode = tmSof.iMnfSSCode
        ilRet = btrGetNext(hmSof, tmSof, imSofRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        ilTemp = ilTemp + 1
    Loop

    ReDim tmNamedAvails(0 To 0) As Integer
    For ilTemp = 0 To RptSelQB!lbcSelection(3).ListCount - 1
        If RptSelQB!lbcSelection(3).Selected(ilTemp) Then
            slNameCode = tgNamedAvail(ilTemp).sKey         'sales source code
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmNamedAvails(UBound(tmNamedAvails)) = Val(slCode)
            ReDim Preserve tmNamedAvails(0 To UBound(tmNamedAvails) + 1) As Integer
        End If
    Next ilTemp

    'get all the dates needed to work with
'    slDate = RptSelQB!edcSelCFrom.Text               'effective date entred
    slDate = RptSelQB!CSI_CalFrom.Text               '12-11-19 chg to use csi calendar control; effective date entred
    llEffDate = gDateValue(slDate)

    If ilStdQtr Then                    '12-13-00
        'get standard bdcst year from the start date entered
        slEffDate = Format$(llEffDate, "m/d/yy")           'insure the year is formatted from input
        slDate = gObtainYearStartDate(0, slDate)
        llDate = gDateValue(slDate)
        Do While (llEffDate < llDate) Or (llEffDate > llDate + 13 * 7 - 1)
            llDate = llDate + 13 * 7
        Loop
    Else
        llDate = gDateValue(slDate)         'use date entered
        'backup to Monday
        illoop = gWeekDayLong(llDate)
        Do While illoop <> 0
            llDate = llDate - 1
            illoop = gWeekDayLong(llDate)
        Loop
    End If
    
    slQSDate = Format$(llDate, "m/d/yy")
    tmVef.iCode = 0
   
    'ilWksEntered = Val(RptSelQB!edcSelCFrom1.Text)
    ilWksEntered = igMonthOrQtr         'adjusted # of weeks to print if using std qtr and startdate not on the std qtr
    '4-27-12 # quarters selectivity changed to # weeks.  Determine # of qtrs based on # weeks entered
    ilNoQtrs = ilWksEntered \ 13
    If (ilNoQtrs * 13) < ilWksEntered Then
        ilNoQtrs = ilNoQtrs + 1
    End If
    
    'TTP 10729 - Quarterly Booked Spots report: add digital lines
    ReDim imVefList(0 To 0) 'TTP 10787 - Quarterly Booked Spots report: subscript out of range error when following specific steps
    If RptSelQB!rbcVersion(1).Value = True Then
        ilNoQtrs = 1
        'ilWksEntered = 5 'The “# Weeks” field will not be available. This version of the report always runs for one broadcast month, which is 4 or 5 weeks.
        slQSDate = gObtainStartStd(RptSelQB!CSI_CalFrom.Text) 'The start date will always be the start of a standard broadcast month. If a start date is selected that is not the start of the broadcast month, it will get automatically set to the start of the standard broadcast month that the selected date is in.
        llEffDate = gDateValue(slQSDate)
        slStartDate = gObtainStartStd(RptSelQB!CSI_CalFrom.Text)
        slEndDate = gObtainEndStd(RptSelQB!CSI_CalFrom.Text)
        ilWksEntered = DateDiff("W", gObtainStartStd(RptSelQB!CSI_CalFrom.Text), gObtainEndStd(RptSelQB!CSI_CalFrom.Text)) + 1
        
        'Get all distinct digital vehicleIDs between start/end dates
        mObtainPCFVehicles slStartDate, slEndDate
        
        'Month Start Date
        gPackDate slQSDate, ilSAvailsDates(0), ilSAvailsDates(1)

        'Get Receivables for Podcast
        tlTranTypes.iAdj = True
        tlTranTypes.iAirTime = True
        tlTranTypes.iCash = True
        tlTranTypes.iInv = True
        tlTranTypes.iMerch = False
        tlTranTypes.iNTR = False
        tlTranTypes.iPromo = False
        tlTranTypes.iPymt = False
        tlTranTypes.iTrade = False
        tlTranTypes.iWriteOff = False
        
        ReDim tlRvf(0 To 0) As RVF
        
        'If running report for period AFTER the last billing date – Skip digital
        If DateValue(slStartDate) < DateValue(Format(llLastBilledStd, "ddddd")) Then
            ilRet = gObtainPhfRvf(RptSelQB, slQSDate, slEndDate, tlTranTypes, tlRvf(), 0)
            Debug.Print "gCRQtrlyBookSpots with digital lines: " & slQSDate & " to " & slEndDate
        Else
            'Skip Digital (Dont load Receivables)
            Debug.Print "gCRQtrlyBookSpots skip digital lines, report period:" & slStartDate & " is after Last billed date: " & Format(llLastBilledStd, "ddddd")
        End If
    End If

    For ilQNo = 1 To ilNoQtrs Step 1
        llDate = gDateValue(slQSDate)
        If ilWksEntered <= 13 Then           '1 qtr or less than 1 qtr to process?
            ilHowManyWks = ilWksEntered
        Else
            ilHowManyWks = 13
            ilWksEntered = ilWksEntered - 13
        End If
        ilFirstQ = 1
        For illoop = 1 To ilHowManyWks
            If (llEffDate >= llDate) And (llEffDate <= llDate + 6) Then
                ilFirstQ = illoop
            End If
            lmSAvailsDates(illoop) = llDate
            lmEAvailsDates(illoop) = llDate + 6
            llDate = llDate + 7
        Next illoop
        
        '-----------------------------------------------------
        'Loop through each Selected vehicle
        For ilVehicle = 0 To RptSelQB!lbcSelection(0).ListCount - 1 Step 1
            If (RptSelQB!lbcSelection(0).Selected(ilVehicle)) Then
                slNameCode = tgCSVNameCode(ilVehicle).sKey 'RptSelSP!lbcCSVNameCode.List(ilVehicle)
                ilRet = gParseItem(slNameCode, 1, "\", slName)
                ilRet = gParseItem(slName, 3, "|", slName)
                ilRet = gParseItem(slNameCode, 2, "\", slCode)
                ilVefCode = Val(slCode)
                ilVpfIndex = -1
                illoop = gBinarySearchVpf(ilVefCode)
                If illoop <> -1 Then
                    ilVpfIndex = illoop
                End If
                If ilVpfIndex >= 0 Then
                    imRcf = -1
                    For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1
                        tmRcf = tgMRcf(ilRcf)
                        ilDateOk = False
                        For illoop = 0 To RptSelQB!lbcSelection(1).ListCount - 1 Step 1
                            slNameCode = tgRateCardCode(illoop).sKey
                            ilRet = gParseItem(slNameCode, 3, "\", slCode)
                            If Val(slCode) = tgMRcf(ilRcf).iCode Then
                                If (RptSelQB!lbcSelection(1).Selected(illoop)) Then
                                    ilDateOk = True
                                    imRcf = ilRcf
                                End If
                                Exit For
                            End If
                        Next illoop
                        
                        If ilDateOk Then
                            ReDim tmBooked(0 To 0) As BOOKSPOTS
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
                                            For illoop = LBound(tmAvRdf) To ilUpper - 1 Step 1
                                                If tmAvRdf(illoop).iCode = tgMRdf(ilRdf).iCode Then
                                                    ilFound = True
                                                    Exit For
                                                End If
                                            Next illoop
                                            If Not ilFound Then
                                                tmAvRdf(ilUpper) = tgMRdf(ilRdf)
                                                tmRifRate(ilUpper) = tgMRif(llRif)
                                                tmRifRate(ilUpper).iSort = ilSaveSort
                                                ilUpper = ilUpper + 1
                                                ReDim Preserve tmAvRdf(0 To ilUpper) As RDF
                                                ReDim Preserve tmRifRate(0 To ilUpper) As RIF
                                            End If
                                        End If
                                    End If
                                End If
                            Next llRif
                            
                            '-----------------------------------------------------
                            'Get Spots
                            llStart = lmSAvailsDates(ilFirstQ)
                            'llEnd = lmEAvailsDates(ilWksEntered)
                            llEnd = lmEAvailsDates(ilHowManyWks)
                            mGetSpotCountsQB ilVefCode, ilVpfIndex, ilFirstQ, lmSAvailsDates(1), llEnd, lmSAvailsDates(), lmEAvailsDates(), ilHowManyWks, tmAvRdf(), tmRifRate(), tlCntTypes
                            'tmBooked contains list of unique spot; tmAvr contains avails, inventory, etc information
                            For ilRec = 0 To UBound(tmBooked) - 1 Step 1
                                For ilIndex = 0 To UBound(tmAvr) - 1 Step 1
                                    If tmAvr(ilIndex).iRdfCode = tmBooked(ilRec).BkKeyRec.iRdfCode Then
                                        tlAvr = tmAvr(ilIndex)
                                        tlAvr.iVefCode = ilVefCode                          'vehicle code
                                        tlAvr.iDay = tmBooked(ilRec).BkKeyRec.ivefSellCode  'selling vehicle (if spot moved)
                                        tlAvr.sInOut = tmBooked(ilRec).BkKeyRec.sSpotType   'Spot type (use line rates or its a fill)
                                        tlAvr.sBucketType = tmBooked(ilRec).BkKeyRec.sPriceType 'show $, or N/C
                                        tlAvr.ianfCode = tmBooked(ilRec).BkKeyRec.iLen      'spot length
                                        tlAvr.sDPDays = Trim$(tmBooked(ilRec).BkKeyRec.sDysTms)     'daypart days & times override
                                        tlAvr.sNot30Or60 = tmBooked(ilRec).BkKeyRec.sAirMissed      'A = aired, M = missed
                                        tlAvr.lMonth(0) = tmBooked(ilRec).BkKeyRec.lChfCode         'contract code
                                        tlAvr.lMonth(1) = tmBooked(ilRec).BkKeyRec.lRate            'spot rate
                                        tlAvr.lMonth(2) = tmBooked(ilRec).BkKeyRec.lFsfCode         'feed spot code
                                        tlAvr.iWksInQtr = tmBooked(ilRec).BkKeyRec.iPkgFlag         'if nonzero - a mg spot is from a package line
                                        For illoop = 1 To 13
                                            tlAvr.lRate(illoop - 1) = tmBooked(ilRec).iSpotCounts(illoop) 'spot counts for 13 weeks
                                        Next illoop
                                        tlAvr.iGenDate(0) = igNowDate(0)
                                        tlAvr.iGenDate(1) = igNowDate(1)
                                        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                        tlAvr.lGenTime = lgNowTime
                                        ilRet = btrInsert(hmAvr, tlAvr, imAvrRecLen, INDEXKEY0)
                                    End If
                                Next ilIndex
                            Next ilRec
                            
                            '-----------------------------------------------------
                            'TTP 10729 - Quarterly Booked Spots report: add digital lines
                            'Check if current vehicle has any Digital lines
                            ilFound = False
                            For ilIndex = 0 To UBound(imVefList)
                                If imVefList(ilIndex) = ilVefCode Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilIndex
                            
                            '---------------------
                            'This only works on Invoiced items
                            If UBound(tlRvf) = 0 Then ilFound = False
                            
                            '---------------------
                            'Found a Digital vehicle that may have PCF Lines
                            If ilFound = True Then
                                '---------------------
                                'Get PCF records for vehicle
                                ReDim tmPLPcf(0 To 0) As PCFTYPESORT
                                mObtainSelPcf 6, ilVefCode, slStartDate, slEndDate, tlCntTypes
                                '---------------------
                                'Loop through all the PCF records for the vehicle in the Date Range
                                ilFound = False
                                For ilRec = 0 To UBound(tmPLPcf) - 1 Step 1
                                    'Get Receivables for this pcf Record for this Month
                                     llMonthlyAmount = 0
                                    For illoop = LBound(tlRvf) To UBound(tlRvf) - 1
                                        If tlRvf(illoop).lCntrNo = tmPLPcf(ilRec).lCntrNo Then
                                            If gObtainPcfCPMID(tlRvf(illoop).lPcfCode) = tmPLPcf(ilRec).tPcf.iPodCPMID Then
                                                gPDNToLong tlRvf(illoop).sGross, llAmount
                                                llMonthlyAmount = llMonthlyAmount + llAmount
                                                ilFound = True
                                            End If 'Matching Line
                                        End If 'Matching Contract
                                    Next illoop
                                    
                                    If ilFound Then 'this contract has been invoiced
                                        'Net or Gross?
                                        If imGrossNet = 1 Then
                                            llMonthlyAmount = gGetGrossNetTNetFromPrice(imGrossNet, llMonthlyAmount, 0, tmPLPcf(ilRec).iAgfCode)
                                        End If

                                        tlAvr.iQStartDate(0) = ilSAvailsDates(0)
                                        tlAvr.iQStartDate(1) = ilSAvailsDates(1)
                                        tlAvr.iVefCode = ilVefCode                          'vehicle code
                                        tlAvr.sInOut = "D" 'to indicate to Crystal "Digital"
                                        If tmPLPcf(ilRec).tPcf.sPriceType = "C" Then tlAvr.sBucketType = "C" '"C"=CPM or "D"=Flat
                                        If tmPLPcf(ilRec).tPcf.sPriceType = "F" Then tlAvr.sBucketType = "D" '"C"=CPM or "D"=Flat
                                        tlAvr.ianfCode = 0      'spot length
                                        tlAvr.iRdfCode = tmPLPcf(ilRec).tPcf.iRdfCode
                                        tlAvr.sDPDays = ""     'daypart days & times override
                                        ilRdf = gBinarySearchRdf(tmPLPcf(ilRec).tPcf.iRdfCode)
                                        If ilRdf <> -1 Then
                                            tlAvr.sDPDays = tgMRdf(ilRdf).sName     'daypart days & times override
                                        End If
                                        tlAvr.lMonth(0) = tmPLPcf(ilRec).tPcf.lChfCode        'contract code
                                        gUnpackDateLong tmPLPcf(ilRec).tPcf.iStartDate(0), tmPLPcf(ilRec).tPcf.iStartDate(1), llTempStartDate
                                        gUnpackDateLong tmPLPcf(ilRec).tPcf.iEndDate(0), tmPLPcf(ilRec).tPcf.iEndDate(1), llTempEndDate
                                        
                                        ilHowManyDays = mGetNumberOfDaysRunning(slQSDate, slEndDate, Format(llTempStartDate, "ddddd"), Format(llTempEndDate, "ddddd"))
                                        'Daily rate
                                        If ilHowManyDays = 0 Then
                                            tlAvr.lMonth(1) = 0
                                        Else
                                            tlAvr.lMonth(1) = llMonthlyAmount / ilHowManyDays
                                        End If
                                        tlAvr.lMonth(2) = llMonthlyAmount 'to prevent rounding issues
                                        For illoop = 1 To ilWksEntered
                                            tlAvr.lRate(illoop - 1) = mGetNumberOfDaysRunning(Format(llTempStartDate, "ddddd"), Format(llTempEndDate, "ddddd"), Format(lmSAvailsDates(illoop), "ddddd"), Format(lmEAvailsDates(illoop), "ddddd")) 'number of days running for each week
                                        Next illoop
                                        tlAvr.iGenDate(0) = igNowDate(0)
                                        tlAvr.iGenDate(1) = igNowDate(1)
                                        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                        tlAvr.lGenTime = lgNowTime
                                        ilRet = btrInsert(hmAvr, tlAvr, imAvrRecLen, INDEXKEY0)
                                        
                                        tlAvr.lMonth(2) = 0
                                    End If
                                Next ilRec
                            End If
                        End If
                    Next ilRcf                      'next rate card (should only be 1)
                End If                              'vpfindex > 0
            End If                                  'vehicle selected
        Next ilVehicle                              'For ilvehicle = 0 To RptSelQb!lbcSelection(0).ListCount - 1
        llDate = gDateValue(slQSDate) + 13 * 7
        slQSDate = Format$(llDate, "m/d/yy")
        llEffDate = llDate
    Next ilQNo                                      'next quarter
    Erase tmAvRdf, tmAvr, tmBooked, tmSofList
    sgMNFCodeTagRpt = ""
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmFsf)
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
    btrDestroy hmAnf
    btrDestroy hmFsf
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

'**********************************************************************
'*                                                                    *
'*      Procedure Name:mGetSpotCountsQB                               *
'*                                                                    *
'*             Created:12/29/97      By:D. Hosaka                     *
'*                                                                    *
'*             Copy of gGetAvailsCounts to access spot                *
'*             and save spot data for Quarterly Booked                *
'*             report                                                 *
'*                                                                    *
'*            Comments:Obtain the Avail counts and spot               *
'*            detail                                                  *
'           6-12-00 If an orphan missed spot is processed, the        *
'               incorrect daypart information may be shown for the    *
'               spot                                                  *
'                                                                     *
'   4-27-12 implement varying # weeks vs forcing to a qtr at a time   *
'**********************************************************************
Sub mGetSpotCountsQB(ilVefCode As Integer, ilVpfIndex As Integer, ilFirstQ As Integer, llSDate As Long, llEDate As Long, llSAvails() As Long, llEAvails() As Long, ilHowManyWks As Integer, tlAvRdf() As RDF, tlRif() As RIF, tlCntTypes As CNTTYPES)
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
'   tmAvr() (O) - array of AVR records built for avails
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
'
'   Note: Remnants; Direct Response; per Inquiry; PSA and Promos are not
'         saved with a miss status
'         For scheduled spots the rank is used to determine if it is one
'         of the above (Direct reponse=1010; Remnant=1020; per Inquiry= 1030;
'         PSA=1060; Promo=1050.
'
'   slBucketType(I): A=Avail; S=Sold; I=Inventory  , P = Percent sellout    'forced to "A" for avail
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
    Dim illoop As Integer
    Dim ilFound As Integer
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilRec As Integer
    Dim ilRecIndex As Integer
    Dim ilLen As Integer
    Dim ilUnits As Integer
    Dim ilNo30 As Integer
    Dim ilNo60 As Integer
    Dim ilDay As Integer
    Dim ilSaveDay As Integer
    Dim slDays As String
    Dim ilLtfCode As Integer
    Dim ilAvailOk As Integer
    Dim ilPass As Integer
    Dim ilDayIndex As Integer
    Dim ilLoopIndex As Integer
    Dim ilBucketIndex As Integer
    Dim ilBucketIndexMinusOne As Integer
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
    Dim ilVefIndex As Integer
    Dim ilOrphanMissedLoop As Integer
    Dim ilOrphanFound As Integer
    ReDim ilSAvailsDates(0 To 1) As Integer
    ReDim ilEvtType(0 To 14) As Integer
    ReDim ilRdfCodes(0 To 1) As Integer
    ReDim tmAvr(0 To 0) As AVR
    Dim tlBkKey As BOOKKEY
    Dim tlOrderedRDF As RDF     'image for missed spot ordered DP info
    Dim llRif As Long
    Dim ilAvailLoop As Integer
    Dim ilFoundAvail As Integer
    Dim slUserRequest As String * 1

    slBucketType = "S"                          'force to Sellout, which calculations inventory, avails & sold
    slDate = Format$(llSAvails(1), "m/d/yy")
    gPackDate slDate, ilSAvailsDates(0), ilSAvailsDates(1)

    'ReDim ilWksInMonth(1 To 3) As Integer
    ReDim ilWksInMonth(0 To 3) As Integer   'Index zero ignored
    slStr = slDate
    For illoop = 1 To 3 Step 1
        llDate = gDateValue(gObtainStartStd(slStr))
        llLoopDate = gDateValue(gObtainEndStd(slStr)) + 1
        ilWksInMonth(illoop) = ((llLoopDate - llDate) / 7)
        slStr = Format(llLoopDate, "m/d/yy")
    Next illoop
    'Currently 14 week quarters are not handled - drop 14th week
    If ilWksInMonth(1) + ilWksInMonth(2) + ilWksInMonth(3) > 13 Then
        ilWksInMonth(3) = ilWksInMonth(3) - 1
    End If
    slType = "O"
    ilType = 0
    ilVefIndex = gBinarySearchVef(ilVefCode)
    llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
    'set the type of events to get fro the day (only Contract avails)
    For illoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(illoop) = False
    Next illoop
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
    'For llLoopDate = llSDate To llEDate Step 1
    '6-6-12
    For llLoopDate = llSAvails(ilFirstQ) To llEDate
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
                tmSsf.iCount = 0
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
            'For ilLoop = 1 To 13 Step 1
            For illoop = 1 To ilHowManyWks              '4-27-12
                If (llDate >= llSAvails(illoop)) And (llDate <= llEAvails(illoop)) Then
                    ilBucketIndex = illoop
                    Exit For
                End If
            Next illoop
            If ilBucketIndex > 0 Then
                ilBucketIndexMinusOne = ilBucketIndex - 1
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
                            ilAvailOk = False
                            If (tlAvRdf(ilRdf).iLtfCode(0) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(1) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(2) <> 0) Then
                                If (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(0)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Then
                                    ilAvailOk = False    'True- code later
                                End If
                            Else

                                For illoop = LBound(tlAvRdf(ilRdf).iStartTime, 2) To UBound(tlAvRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                    If (tlAvRdf(ilRdf).iStartTime(0, illoop) <> 1) Or (tlAvRdf(ilRdf).iStartTime(1, illoop) <> 0) Then
                                        gUnpackTimeLong tlAvRdf(ilRdf).iStartTime(0, illoop), tlAvRdf(ilRdf).iStartTime(1, illoop), False, llStartTime
                                        gUnpackTimeLong tlAvRdf(ilRdf).iEndTime(0, illoop), tlAvRdf(ilRdf).iEndTime(1, illoop), True, llEndTime
'                                        If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDay+1) = "Y") Then
                                        If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(illoop, ilDay) = "Y") Then
                                            ilAvailOk = True
                                            ilLoopIndex = illoop
                                            slDays = ""
                                            For ilDayIndex = 1 To 7 Step 1
                                                If (tlAvRdf(ilRdf).sWkDays(illoop, ilDayIndex - 1) = "Y") Or (tlAvRdf(ilRdf).sWkDays(illoop, ilDayIndex - 1) = "N") Then
                                                    slDays = slDays & tlAvRdf(ilRdf).sWkDays(illoop, ilDayIndex - 1)
                                                Else
                                                    slDays = slDays & "N"
                                                End If
                                            Next ilDayIndex
                                            Exit For
                                        End If
                                    End If
                                Next illoop
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
                            End If

                             '7-19-04 the Named avail property must allow local spots to be included
                            tmAnfSrchKey.iCode = tmAvail.ianfCode
                            ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                            If (ilRet = BTRV_ERR_NONE) Then
                                If Not tlCntTypes.iCntrSpots And tmAnf.sBookLocalFeed = "L" Then      'Local avail requested to be excluded, exclude if avail type = "L"
                                    ilAvailOk = False
                                End If
                                If Not tlCntTypes.iFeedSpots And tmAnf.sBookLocalFeed = "F" Then      'Network avail requested to be excluded, exclude if avail type = "F"
                                    ilAvailOk = False
                                End If
                            End If

                            If ilAvailOk Then
                                'Determine if Avr created
                                ilFound = False
                                ilSaveDay = ilDay
                                'If RptSelCt!rbcSelCInclude(0).Value Then              'daypart option, place all values in same record
                                                                                    'to get better availability
                                    ilDay = 0                                       'force all data in same day of week
                                'End If
                                For ilRec = 0 To UBound(tmAvr) - 1 Step 1
                                    'If (tmAvr(ilRec).iRdfCode = tlAvRdf(ilRdf).iCode) And (tmAvr(ilRec).iFirstBucket = ilFirstQ) And (tmAvr(ilRec).iDay = ilDay) Then
                                    If (ilRdfCodes(ilRec) = tlAvRdf(ilRdf).iCode) And (tmAvr(ilRec).iFirstBucket = ilFirstQ) And (tmAvr(ilRec).iDay = ilDay) Then
                                        ilFound = True
                                        ilRecIndex = ilRec
                                        Exit For
                                    End If
                                Next ilRec
                                If Not ilFound Then
                                    ilRecIndex = UBound(tmAvr)
                                    tmAvr(ilRecIndex).iGenDate(0) = igNowDate(0)
                                    tmAvr(ilRecIndex).iGenDate(1) = igNowDate(1)
                                    gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                    tmAvr(ilRecIndex).lGenTime = lgNowTime
                                    tmAvr(ilRecIndex).iVefCode = ilVefCode
                                    tmAvr(ilRecIndex).iDay = ilDay
                                    tmAvr(ilRecIndex).iQStartDate(0) = ilSAvailsDates(0)
                                    tmAvr(ilRecIndex).iQStartDate(1) = ilSAvailsDates(1)
                                    tmAvr(ilRecIndex).iFirstBucket = ilFirstQ
                                    tmAvr(ilRecIndex).sBucketType = slBucketType
                                    ilRdfCodes(ilRecIndex) = tlAvRdf(ilRdf).iCode
                                    tmAvr(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iCode
                                    tmAvr(ilRecIndex).iRdfSortCode = tlRif(ilRdf).iSort
                                    tmAvr(ilRecIndex).sInOut = tlAvRdf(ilRdf).sInOut
                                    tmAvr(ilRecIndex).ianfCode = tlAvRdf(ilRdf).ianfCode
                                    tmAvr(ilRecIndex).iDPStartTime(0) = tlAvRdf(ilRdf).iStartTime(0, ilLoopIndex)
                                    tmAvr(ilRecIndex).iDPStartTime(1) = tlAvRdf(ilRdf).iStartTime(1, ilLoopIndex)
                                    tmAvr(ilRecIndex).iDPEndTime(0) = tlAvRdf(ilRdf).iEndTime(0, ilLoopIndex)
                                    tmAvr(ilRecIndex).iDPEndTime(1) = tlAvRdf(ilRdf).iEndTime(1, ilLoopIndex)
                                    tmAvr(ilRecIndex).sDPDays = slDays
                                    tmAvr(ilRecIndex).sNot30Or60 = "N"
                                    ReDim Preserve tmAvr(0 To ilRecIndex + 1) As AVR
                                    ReDim Preserve ilRdfCodes(0 To ilRecIndex + 1)
                                End If
                                tmAvr(ilRecIndex).lRate(ilBucketIndexMinusOne) = tlRif(ilRdf).lRate(ilWkNo)
                                ilDay = ilSaveDay
                                'Always gather inventory
                                ilLen = tmAvail.iLen
                                ilUnits = tmAvail.iAvInfo And &H1F
                                slUserRequest = "B"             'default to 30" unit count
                                If sm30sOrUnits = "U" Then      'unit (spot counts(
                                    slUserRequest = "C"
                                End If
                                gGatherInventory tmAvr(), ilVpfIndex, slBucketType, ilRecIndex, ilBucketIndex, ilLen, ilUnits, ilNo30, ilNo60, slUserRequest      '8-13-10 handle output as how the

                                'Always calculate Avails
                                For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                   LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                                    ilSpotOK = True                             'assume spot is OK to include
                                    'these spot rankings are changed (decremented by 1) if no priority is used and its last week of order
                                    'continue to test with contract header type
                                    If ((tmSpot.iRank And RANKMASK) = REMNANTRANK) And (Not tlCntTypes.iRemnant) Then
                                        ilSpotOK = False
                                    End If
                                    If ((tmSpot.iRank And RANKMASK) = PERINQUIRYRANK) And (Not tlCntTypes.iPI) Then
                                        ilSpotOK = False
                                    End If
                                    If ((tmSpot.iRank And RANKMASK) = TRADERANK) And (Not tlCntTypes.iTrade) Then
                                        ilSpotOK = False
                                    End If
                                    If ((tmSpot.iRank And RANKMASK) = EXTRARANK) And (Not tlCntTypes.iXtra) Then
                                        ilSpotOK = False
                                    End If
                                    If ((tmSpot.iRank And RANKMASK) = PROMORANK) And (Not tlCntTypes.iPromo) Then
                                        ilSpotOK = False
                                    End If
                                    If ((tmSpot.iRank And RANKMASK) = PSARANK) And (Not tlCntTypes.iPSA) Then
                                        ilSpotOK = False
                                    End If

                                    If ilSpotOK Then                            'continue testing other filters
                                        tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                        ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)

                                        If tmSpot.lSdfCode = tmSdf.lCode And ilRet = BTRV_ERR_NONE Then

                                            If tmSdf.lChfCode = 0 Then             'feed spot
                                                If Not tlCntTypes.iFeedSpots Then
                                                    ilSpotOK = False
                                                End If
                                                tmChfSrchKey.lCode = tmSdf.lFsfCode
                                                ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                mBuildFeedDetail ilVefCode, ilBucketIndex, ilRdf, "A", tlCntTypes, tlBkKey, tlAvRdf(), ilSpotOK, 0
                                            Else
                                                If Not tlCntTypes.iCntrSpots Then
                                                    ilSpotOK = False
                                                End If
                                                If tmSdf.lChfCode <> tmChf.lCode Then               'if already in mem, don't reread
                                                    tmChfSrchKey.lCode = tmSdf.lChfCode
                                                    ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                                Else
                                                    ilRet = BTRV_ERR_NONE
                                                End If
                                                If ilRet <> BTRV_ERR_NONE Then
                                                    ilSpotOK = False
                                                End If

                                                If (tmChf.sType = "C") And (Not tlCntTypes.iStandard) Then      'include Standard types?
                                                    ilSpotOK = False
                                                End If
                                                If (tmChf.sType = "V") And (Not tlCntTypes.iReserv) Then      'include reservations ?
                                                    ilSpotOK = False
                                                End If
                                                If (tmChf.sType = "R") And (Not tlCntTypes.iDR) Then      'include DR?
                                                    ilSpotOK = False
                                                End If
                                                
                                                If (tmChf.sType = "T") And (Not tlCntTypes.iRemnant) Then
                                                    ilSpotOK = False
                                                End If
                                                If (tmChf.sType = "Q") And (Not tlCntTypes.iPI) Then
                                                    ilSpotOK = False
                                                End If
                                                                                                
                                                If (tmChf.sType = "M") And (Not tlCntTypes.iPromo) Then
                                                    ilSpotOK = False
                                                End If
                                                If (tmChf.sType = "S") And (Not tlCntTypes.iPSA) Then
                                                    ilSpotOK = False
                                                End If
                                                
                                                If (tmChf.iPctTrade = 100) And (Not tlCntTypes.iTrade) Then       'exclude only if 100% trade
                                                    ilSpotOK = False
                                                End If

                                                If ilSpotOK Then
                                                    'Determine if spot within avail is OK to include in report
                                                     '6-12-00 add orphan flag to parameeters
                                                     mBuildCntrDetail ilVefCode, ilBucketIndex, ilRdf, "A", tlCntTypes, tlBkKey, tlAvRdf(), ilSpotOK, 0
                                                End If
                                            End If
                                        End If

                                        If ilSpotOK Then
                                            ilLen = tmSdf.iLen
                                            ilNo30 = 0
                                            ilNo60 = 0
                                            If sm30sOrUnits = "3" Then                          '30" unit counts

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
                                                        If tmChf.sType = "V" Then                       'Type reserve
                                                            tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) + ilNo30
                                                            tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) + ilNo60
                                                        ElseIf tmChf.sStatus = "H" Then                         'staus "Hold" , always show on separate line
                                                            tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) + ilNo30
                                                            tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) + ilNo60
                                                        Else
                                                            tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                            tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                        End If
                                                    End If
                                                    'adjust the available buckets (used for qtrly detail  report only)
                                                    tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                                    tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then               'units sold
                                                    'Count 30 or 60 and set flag if neither
                                                    If ilLen = 60 Then
                                                        ilNo60 = 1
                                                    ElseIf ilLen = 30 Then
                                                        ilNo30 = 1
                                                    Else
                                                        tmAvr(ilRecIndex).sNot30Or60 = "Y"
                                                        If ilLen <= 30 Then
                                                            ilNo30 = 1
                                                        Else
                                                            ilNo60 = 1
                                                        End If
                                                    End If
                                                    If (ilNo60 <> 0) Or (ilNo30 <> 0) Then
                                                        If (slBucketType = "S") Or (slBucketType = "P") Then
                                                            If tmChf.sType = "V" Then                       'Type reserve
                                                                tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) + ilNo30
                                                                tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) + ilNo60
                                                            ElseIf tmChf.sStatus = "H" Then                         'staus "Hold", always show on separate line
                                                                tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) + ilNo30
                                                                tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) + ilNo60
                                                            Else
                                                                tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                                tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                            End If
                                                        End If
                                                    End If
                                                    'adjust the available buckets (used for qtrly detail report only)
                                                    tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                                    tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then               'matching units
                                                    'Count 30 or 60 and set flag if neither
                                                    If ilLen = 60 Then
                                                        ilNo60 = 1
                                                    ElseIf ilLen = 30 Then
                                                        ilNo30 = 1
                                                    Else
                                                        tmAvr(ilRecIndex).sNot30Or60 = "Y"
                                                    End If
                                                    If (slBucketType = "S") Or (slBucketType = "P") Then        'if Sellout or % sellout, accum the seconds sold
                                                        'Qtrly detail has been forced to "Sellout" for internal testing
                                                        If tmChf.sType = "V" Then                       'Type reserve
                                                            'Show on separate line or bury in sold?
                                                                tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) + ilNo30
                                                                tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) + ilNo60
                                                        ElseIf tmChf.sStatus = "H" Then                         'staus "Hold", always show on separate line
                                                            tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) + ilNo30
                                                            tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) + ilNo60
                                                        Else            'not held or reserved, put in sold
                                                            tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                            tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                        End If
                                                    End If
                                                    'adjust the available bucket (used for qrtrly detail report only)
                                                    tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                                    tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                                End If
                                            Else                            '6-14-19 unit (spot) counts
                                                'Count 30 or 60 and set flag if neither
                                                If ilLen = 60 Then
                                                    ilNo60 = 1
                                                ElseIf ilLen = 30 Then
                                                    ilNo30 = 1
                                                Else
                                                    tmAvr(ilRecIndex).sNot30Or60 = "Y"
                                                    If ilLen <= 30 Then
                                                        ilNo30 = 1
                                                    Else
                                                        ilNo60 = 1
                                                    End If
                                                End If
                                                If (ilNo60 <> 0) Or (ilNo30 <> 0) Then
                                                    If (slBucketType = "S") Or (slBucketType = "P") Then
                                                        If tmChf.sType = "V" Then                       'Type reserve
                                                            tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Reserve(ilBucketIndexMinusOne) + ilNo30
                                                            tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Reserve(ilBucketIndexMinusOne) + ilNo60
                                                        ElseIf tmChf.sStatus = "H" Then                         'staus "Hold", always show on separate line
                                                            tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Hold(ilBucketIndexMinusOne) + ilNo30
                                                            tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Hold(ilBucketIndexMinusOne) + ilNo60
                                                        Else
                                                            tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                            tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                        End If
                                                    End If
                                                End If
                                                'adjust the available buckets (used for qtrly detail report only)
                                                tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                                tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                            End If
                                        End If                              'ilspotOK
                                    End If
                                Next ilSpot                             'loop from ssf file for # spots in avail
                            End If                                          'Avail OK
                        Next ilRdf                                          'ilRdf = lBound(tlAvRdf)
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
                'If (llDate >= llSAvails(ilFirstQ)) And (llDate <= llEAvails(13)) Then
                 If (llDate >= llSAvails(ilFirstQ)) And (llDate <= llEAvails(ilHowManyWks)) Then            '4-27-12
                    ilBucketIndex = -1
                    'For ilLoop = 1 To 13 Step 1
                    For illoop = 1 To ilHowManyWks                  '4-27-12
                        If (llDate >= llSAvails(illoop)) And (llDate <= llEAvails(illoop)) Then
                            ilBucketIndex = illoop
                            Exit For
                        End If
                    Next illoop

                    '4-3-06 exclude the missed spot if ordered DP doesnt meet printed DP specs
                    If tmSdf.lChfCode = 0 Then
                        'no daypart
                        tmClf.iRdfCode = 0
                    Else
                         '6-4-04 Gather line daypart info.  If not using other missed spots category, need to check daypart to see if exclusions/exclusions
                        tmClfSrchKey.lChfCode = tmSdf.lChfCode
                        tmClfSrchKey.iLine = tmSdf.iLineNo
                        tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
                        tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                        ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                        If (tmClf.lChfCode <> tmSdf.lChfCode) Then
                            ilBucketIndex = -1          'error in reading line, bypass spot
                        Else
                            'need to get the schedule lines daypart info to see if spot can go/cant go into certain avails
                            For llRif = LBound(tgMRif) To UBound(tgMRif) - 1
                                If tgMRif(llRif).iRcfCode = tgMRcf(imRcf).iCode And tgMRif(llRif).iVefCode = ilVefCode And tgMRif(llRif).iRdfCode = tmClf.iRdfCode Then
                                    ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                                    If ilRdf <> -1 Then
                                        'got the daypart, save if there are includes/excludes and named avails
                                        tlOrderedRDF = tgMRdf(ilRdf)
                                        'if not all avails selected, make sure to process missed spot with matching named avail
                                        ilFoundAvail = True
                                        If RptSelQB!ckcAllAvails.Value = vbUnchecked Then      '11-20-08
                                            ilFoundAvail = False

                                            If tlOrderedRDF.sInOut <> "I" And tlOrderedRDF.sInOut <> "O" Then
                                                ilFoundAvail = True
                                            Else
                                                If tlOrderedRDF.sInOut = "I" Then           'only book into specific avail
                                                    For ilAvailLoop = 0 To UBound(tmNamedAvails) - 1
                                                        If tmNamedAvails(ilAvailLoop) = tlOrderedRDF.ianfCode Then        'ok to see this avail
                                                            ilFoundAvail = True
                                                            Exit For
                                                        End If
                                                    Next ilAvailLoop
                                                Else            'exclude specific avail
                                                    ilFoundAvail = True                             'ok to see spot if not one of them selected
                                                    For ilAvailLoop = 0 To UBound(tmNamedAvails) - 1        'if this is a matching avail to include from selectivity, the daypart of spot defined is to exclude it
                                                        If tmNamedAvails(ilAvailLoop) = tlOrderedRDF.ianfCode Then        'ok to see this avail
                                                            ilFoundAvail = False
                                                            Exit For
                                                        End If
                                                    Next ilAvailLoop
                                                End If
                                            End If
                                        End If
                                        If Not ilFoundAvail Then        'didnt find a matching named avail to use based on selectivity
                                            ilBucketIndex = -1
                                        End If
                                    Else
                                        ilBucketIndex = -1          'ignore this spot, cannot find ordered DP
                                    End If
                                End If
                            Next llRif
                        End If
                    End If

                    If ilBucketIndex > 0 And ((tlCntTypes.iCntrSpots = True And tmSdf.lChfCode > 0) Or (tlCntTypes.iFeedSpots = True And tmSdf.lChfCode = 0)) Then
                        ilBucketIndexMinusOne = ilBucketIndex - 1
                        ilDay = gWeekDayLong(llDate)
                        gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                        For ilOrphanMissedLoop = 1 To 3     '6-6-00
                            ilOrphanFound = False
                            For ilRdf = LBound(tlAvRdf) To UBound(tlAvRdf) - 1 Step 1
                                ilAvailOk = False
                                If (tlAvRdf(ilRdf).iLtfCode(0) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(1) <> 0) Or (tlAvRdf(ilRdf).iLtfCode(2) <> 0) Then
                                    If (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(0)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Or (ilLtfCode = tlAvRdf(ilRdf).iLtfCode(1)) Then
                                        ilAvailOk = False    'True- code later
                                    End If
                                Else
                                    For illoop = LBound(tlAvRdf(ilRdf).iStartTime, 2) To UBound(tlAvRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                        If (tlAvRdf(ilRdf).iStartTime(0, illoop) <> 1) Or (tlAvRdf(ilRdf).iStartTime(1, illoop) <> 0) Then
                                            gUnpackTimeLong tlAvRdf(ilRdf).iStartTime(0, illoop), tlAvRdf(ilRdf).iStartTime(1, illoop), False, llStartTime
                                            gUnpackTimeLong tlAvRdf(ilRdf).iEndTime(0, illoop), tlAvRdf(ilRdf).iEndTime(1, illoop), True, llEndTime
                                            If UBound(tlAvRdf) - 1 = LBound(tlAvRdf) Then   'could be a conv bumped spot sched in
                                                                                        'in conven veh.  The VV has DP times different than the
                                                                                        'conven veh.
                                                llStartTime = llTime
                                                llEndTime = llTime + 1              'actual time of spot
                                            End If
                                            'Don't include the end time i.e. 10a-3p is 10a thru 2:59:59p
                                            ilLoopIndex = 1     '11-11-99 day spotmissed isnt valid for DP to be shown
                                            'If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(ilLoop, ilDay + 1) = "Y") Then
                                            If (llTime >= llStartTime) And (llTime < llEndTime) And (tlAvRdf(ilRdf).sWkDays(illoop, ilDay) = "Y") Then
                                                ilAvailOk = True
                                                '3-26-09 no need to test the avail types since there is no avail with a missed spot
                                                'need only to test the ordered daypart (sch line) with the dayparts shown on the report

                                                '4-3-06 test book into ordered DP against the DP attemtping to put spot into
'                                                If tlAvRdf(ilRdf).sInOut = "I" Then     'this avail has an only certain ones allowed
'                                                    If tlAvRdf(ilRdf).ianfCode <> tlOrderedRDF.ianfCode Then    'is the named avails that is allowed in this daypart the same as the one tht is missed?  if not, dont put missed spot in it
'                                                        ilAvailOk = False
'                                                    End If
'                                                ElseIf tlAvRdf(ilRdf).sInOut = "O" Then 'this avail has certain avails that cant be scheduled into it
'                                                    If tlAvRdf(ilRdf).ianfCode = tlOrderedRDF.ianfCode Then 'is the named avail that is disallowed in this dypart the same as the one that is missed?  if it is, dont put missed pot in it
'                                                        ilAvailOk = False
'                                                    End If
'                                                End If

                                                ilLoopIndex = illoop
                                                slDays = ""
                                                For ilDayIndex = 1 To 7 Step 1
                                                    If (tlAvRdf(ilRdf).sWkDays(illoop, ilDayIndex - 1) = "Y") Or (tlAvRdf(ilRdf).sWkDays(illoop, ilDayIndex - 1) = "N") Then
                                                        slDays = slDays & tlAvRdf(ilRdf).sWkDays(illoop, ilDayIndex - 1)
                                                    Else
                                                        slDays = slDays & "N"
                                                    End If
                                                Next ilDayIndex
                                                Exit For
                                            End If
                                        End If
                                    Next illoop
                                End If
                                If ilAvailOk Or ilOrphanMissedLoop = 3 Then     '6-6-00
                                    ilSpotOK = True                'assume spot is OK
                                    ilRet = BTRV_ERR_NONE

                                    If tmSdf.lChfCode = 0 Then          'feed spot
                                        If Not tlCntTypes.iFeedSpots Then
                                            ilSpotOK = False
                                        End If
                                        tmChfSrchKey.lCode = tmSdf.lFsfCode
                                        ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                        mBuildFeedDetail ilVefCode, ilBucketIndex, ilRdf, "M", tlCntTypes, tlBkKey, tlAvRdf(), ilSpotOK, 0
                                    Else
                                        If tlCntTypes.iCntrSpots Then
                                            tmClfSrchKey.lChfCode = tmSdf.lChfCode
                                            tmClfSrchKey.iLine = tmSdf.iLineNo
                                            tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
                                            tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                                            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                            If (tmClf.lChfCode <> tmSdf.lChfCode) Then     'got the matching line reference?
                                                ilSpotOK = False
                                            Else
                                                If ilOrphanMissedLoop = 1 Then
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

                                            If tmChf.sType = "C" And Not tlCntTypes.iStandard Then
                                                ilSpotOK = False
                                            End If
                                            If tmChf.sType = "V" And Not tlCntTypes.iReserv Then      'include reservations ?
                                                ilSpotOK = False
                                            End If
                                            If tmChf.sType = "T" And Not tlCntTypes.iRemnant Then
                                                ilSpotOK = False
                                            End If
                                            If tmChf.sType = "R" And Not tlCntTypes.iDR Then
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
                                            If tmChf.iPctTrade = 100 And Not tlCntTypes.iTrade Then
                                                ilSpotOK = False
                                            End If

                                            'Determine if spot within avail is OK to include in report
                                            '6-12-00 add orphan flag to parameters
                                            mBuildCntrDetail ilVefCode, ilBucketIndex, ilRdf, "M", tlCntTypes, tlBkKey, tlAvRdf(), ilSpotOK, 2
                                        Else                'do not include contr spots
                                            ilSpotOK = False
                                        End If
                                    End If
                                    If ilSpotOK Then
                                        ilOrphanFound = True
                                        'Determine if Avr created
                                        ilFound = False
                                        ilSaveDay = ilDay
                                        'If RptSelCt!rbcSelCInclude(0).Value Then              'daypart option, place all values in same record
                                                                                            'to get better availability
                                            ilDay = 0                                       'force all data in same day of week
                                        'End If
                                        For ilRec = 0 To UBound(tmAvr) - 1 Step 1
                                            'If (tmAvr(ilRec).iRdfCode = tlAvRdf(ilRdf).iCode) And (tmAvr(ilRec).iFirstBucket = ilFirstQ) And (tmAvr(ilRec).iDay = ilDay) Then
                                            If (ilRdfCodes(ilRec) = tlAvRdf(ilRdf).iCode) And (tmAvr(ilRec).iFirstBucket = ilFirstQ) And (tmAvr(ilRec).iDay = ilDay) Then
                                                ilFound = True
                                                ilRecIndex = ilRec
                                                Exit For
                                            End If
                                        Next ilRec
                                        If Not ilFound Then
                                            ilRecIndex = UBound(tmAvr)
                                            tmAvr(ilRecIndex).iGenDate(0) = igNowDate(0)
                                            tmAvr(ilRecIndex).iGenDate(1) = igNowDate(1)
                                            'tmAvr(ilRecIndex).iGenTime(0) = igNowTime(0)
                                            'tmAvr(ilRecIndex).iGenTime(1) = igNowTime(1)
                                            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                            tmAvr(ilRecIndex).lGenTime = lgNowTime
                                            tmAvr(ilRecIndex).iDay = ilDay
                                            tmAvr(ilRecIndex).iQStartDate(0) = ilSAvailsDates(0)
                                            tmAvr(ilRecIndex).iQStartDate(1) = ilSAvailsDates(1)
                                            tmAvr(ilRecIndex).iFirstBucket = ilFirstQ
                                            tmAvr(ilRecIndex).sBucketType = slBucketType
                                            tmAvr(ilRecIndex).iDPStartTime(0) = tlAvRdf(ilRdf).iStartTime(0, ilLoopIndex)
                                            tmAvr(ilRecIndex).iDPStartTime(1) = tlAvRdf(ilRdf).iStartTime(1, ilLoopIndex)
                                            tmAvr(ilRecIndex).iDPEndTime(0) = tlAvRdf(ilRdf).iEndTime(0, ilLoopIndex)
                                            tmAvr(ilRecIndex).iDPEndTime(1) = tlAvRdf(ilRdf).iEndTime(1, ilLoopIndex)
                                            tmAvr(ilRecIndex).sDPDays = slDays
                                            tmAvr(ilRecIndex).sNot30Or60 = "N"

                                            tmAvr(ilRecIndex).iVefCode = ilVefCode
                                            tmAvr(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iCode           'DP code
                                            'tmAvr(ilRecIndex).iRdfCode = tlAvRdf(ilRdf).iSortCode
                                            ilRdfCodes(ilRecIndex) = tlAvRdf(ilRdf).iCode
                                            tmAvr(ilRecIndex).sInOut = tlAvRdf(ilRdf).sInOut
                                            tmAvr(ilRecIndex).ianfCode = tlAvRdf(ilRdf).ianfCode
                                            tmAvr(ilRecIndex).sDPDays = slDays

                                            ReDim Preserve tmAvr(0 To ilRecIndex + 1) As AVR
                                            ReDim Preserve ilRdfCodes(0 To ilRecIndex + 1)
                                        End If
                                        tmAvr(ilRecIndex).lRate(ilBucketIndexMinusOne) = tlRif(ilRdf).lRate(ilWkNo)
                                        ilDay = ilSaveDay
                                        ilNo30 = 0
                                        ilNo60 = 0
                                        ilLen = tmSdf.iLen
                                        If sm30sOrUnits = "3" Then                  '30" unit counts
                                        If tgVpf(ilVpfIndex).sSSellOut = "B" Then
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
                                                If (slBucketType = "S") Or (slBucketType = "P") Then
                                                    tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                    tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                End If
                                                'adjust the available bucket (used for qrtrly detail report only)
                                                tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                                tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "U" Then
                                                'Count 30 or 60 and set flag if neither
                                                If ilLen = 60 Then
                                                    ilNo60 = 1
                                                ElseIf ilLen = 30 Then
                                                    ilNo30 = 1
                                                Else
                                                    tmAvr(ilRecIndex).sNot30Or60 = "Y"
                                                    If ilLen <= 30 Then
                                                        ilNo30 = 1
                                                    Else
                                                        ilNo60 = 1
                                                    End If
                                                End If
                                                If (slBucketType = "S") Or (slBucketType = "P") Then
                                                    tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                    tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                End If
                                                'adjust the available bucket (used for qrtrly detail report only)
                                                If ilNo60 > 0 Then                      'spot found a 60?
                                                    tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                                Else
                                                    If tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) > 0 Then
                                                        tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                                    Else
                                                        If tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) > 0 Then
                                                            tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo30
                                                        Else                        'oversold units
                                                            tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                                        End If
                                                    End If
                                                End If
    
                                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Then
                                                'Count 30 or 60 and set flag if neither
                                                If ilLen = 60 Then
                                                    ilNo60 = 1
                                                ElseIf ilLen = 30 Then
                                                    ilNo30 = 1
                                                Else
                                                    tmAvr(ilRecIndex).sNot30Or60 = "Y"
                                                End If
                                                If (slBucketType = "S") Or (slBucketType = "P") Then
                                                    tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                    tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                                End If
                                                'adjust the available bucket (used for qrtrly detail report only)
                                                tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                                tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                            ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                                            End If
                                        Else                'spot counts
                                            'Count 30 or 60 and set flag if neither
                                            If ilLen = 60 Then
                                                ilNo60 = 1
                                            ElseIf ilLen = 30 Then
                                                ilNo30 = 1
                                            Else
                                                tmAvr(ilRecIndex).sNot30Or60 = "Y"
                                                If ilLen <= 30 Then
                                                    ilNo30 = 1
                                                Else
                                                    ilNo60 = 1
                                                End If
                                            End If
                                            If (slBucketType = "S") Or (slBucketType = "P") Then
                                                tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Count(ilBucketIndexMinusOne) + ilNo30
                                                tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Count(ilBucketIndexMinusOne) + ilNo60
                                            'Else
                                            '    tmAvr(ilRecIndex).i60Count(ilBucketIndex) = tmAvr(ilRecIndex).i60Count(ilBucketIndex) - ilNo60
                                            '    tmAvr(ilRecIndex).i30Count(ilBucketIndex) = tmAvr(ilRecIndex).i30Count(ilBucketIndex) - ilNo30
                                            End If
                                            'adjust the available bucket (used for qrtrly detail report only)
                                            If ilNo60 > 0 Then                      'spot found a 60?
                                                tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo60
                                            Else
                                                If tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) > 0 Then
                                                    tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                                Else
                                                    If tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) > 0 Then
                                                        tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i60Avail(ilBucketIndexMinusOne) - ilNo30
                                                    Else                        'oversold units
                                                        tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) = tmAvr(ilRecIndex).i30Avail(ilBucketIndexMinusOne) - ilNo30
                                                    End If
                                                End If
                                            End If
    
                                        End If
                                        If ilSpotOK Then
                                            Exit For                'force exit on this missed if found a matching daypart
                                        End If
                                    End If                      'ilSpotOK
                                End If                          'ilAvailOK
                                If ilOrphanMissedLoop = 3 Then  '6-6-00
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
    If sm30sOrUnits = "3" Then
        If (slBucketType = "A" And tgVpf(ilVpfIndex).sSSellOut = "B") Then
            For ilRec = 0 To UBound(tmAvr) - 1 Step 1
                'For ilLoop = 1 To 13 Step 1
                For illoop = LBound(tmAvr(ilRec).i30Count) To UBound(tmAvr(ilRec).i30Count) Step 1
                    If tmAvr(ilRec).i30Count(illoop) < 0 Then
                        Do While (tmAvr(ilRec).i60Count(illoop) > 0) And (tmAvr(ilRec).i30Count(illoop) < 0)
                            tmAvr(ilRec).i60Count(illoop) = tmAvr(ilRec).i60Count(illoop) - ilAdjSub    '1
                            tmAvr(ilRec).i30Count(illoop) = tmAvr(ilRec).i30Count(illoop) + ilAdjAdd    '2
                        Loop
                    ElseIf (tmAvr(ilRec).i60Count(illoop) < 0) Then
                    End If
                Next illoop
            Next ilRec
        End If
        'Adjust counts for qtrly detail availbilty
        'If (RptSelCt!rbcSelC4(1).Value) And (tgVpf(ilVpfIndex).sSSellOut = "B" Or tgVpf(ilVpfIndex).sSSellOut = "U") Then                  'qtrly detail?
        'If (tgVpf(ilVpfIndex).sSSellOut = "B" Or tgVpf(ilVpfIndex).sSSellOut = "U") Then                  'qtrly detail?
        If (tgVpf(ilVpfIndex).sSSellOut = "B") Then                   'qtrly detail?
            For ilRec = 0 To UBound(tmAvr) - 1 Step 1
                'For ilLoop = 1 To 13 Step 1
                For illoop = LBound(tmAvr(ilRec).i30Avail) To UBound(tmAvr(ilRec).i30Avail) Step 1
                    If tmAvr(ilRec).i30Avail(illoop) < 0 Then
                        Do While (tmAvr(ilRec).i60Avail(illoop) > 0) And (tmAvr(ilRec).i30Avail(illoop) < 0)
                            tmAvr(ilRec).i60Avail(illoop) = tmAvr(ilRec).i60Avail(illoop) - 1
                            tmAvr(ilRec).i30Avail(illoop) = tmAvr(ilRec).i30Avail(illoop) + 2
                        Loop
                    ElseIf (tmAvr(ilRec).i60Avail(illoop) < 0) Then
                    End If
                Next illoop
            Next ilRec
        End If
    End If
    'Combines weeks into the proper months for monthly figures
    For ilRec = 0 To UBound(tmAvr) - 1 Step 1  'next daypart
        For illoop = 1 To 3 Step 1
            If illoop = 1 Then
                ilLo = 1
                ilHi = ilWksInMonth(1)
            Else
                ilLo = ilHi + 1
                ilHi = ilHi + ilWksInMonth(illoop)
            End If
            For ilIndex = ilLo To ilHi Step 1
                If tgVpf(ilVpfIndex).sSSellOut = "B" Then
                    tmAvr(ilRec).lMonth(illoop - 1) = tmAvr(ilRec).lMonth(illoop - 1) + (((tmAvr(ilRec).i60Avail(ilIndex - 1) * 2) + tmAvr(ilRec).i30Avail(ilIndex - 1)) * tmAvr(ilRec).lRate(ilIndex - 1))
                ElseIf tgVpf(ilVpfIndex).sSSellOut = "M" Or tgVpf(ilVpfIndex).sSSellOut = "U" Then
                    tmAvr(ilRec).lMonth(illoop - 1) = tmAvr(ilRec).lMonth(illoop - 1) + ((tmAvr(ilRec).i60Avail(ilIndex - 1) + tmAvr(ilRec).i30Avail(ilIndex - 1)) * tmAvr(ilRec).lRate(ilIndex - 1))
                ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
                End If
            Next ilIndex
        Next illoop
    Next ilRec
    Erase ilSAvailsDates
    Erase ilEvtType
    Erase ilRdfCodes
    Erase tlLLC
End Sub

'---------------------------------------------------------------------------------------------------------
'               mBuildDetail - Test header and line exclusions for user request.
'                       Determine spot rates and build table of unique spots found,
'                       (one entry per advt, rate flag, spot length, cntr, daypart (with override days & times)).
'
'               <input> ilVefCode - airing vehicle
'                       ilbucketIndex - Week index (1 - 13) of week processing
'                       ilRdf - index into Daypart processing
'                       slAirMissedFlag - A = processing aired spots, M = processing Missed spots
'                       tlCntTypes - structure of inclusions/exclusions of contract types & status
'                       tlBkKey - Key fields to keep cntr/length,rate etc apart
'                       tlAvRdf() - daypart processing
'                       ilOrphanMissedLoop - if processing orphan missed spot, flag = 2, otherwise does not matter
'                                           (when flag = 2, the daypart of the schedule line is shown)
'              <output> ilSpotOk - true if spot is OK, else false to ignore spot
'
'               dh 8-22-00 Combine all fills for the same cnt on same line (previously showed on different output lines
'               because the fills came from different vehicles
'               7-22-04 Include/Exclude contract/feed spots

'
'
'           Build up the detail spot information.  Put out same cnts DP together
'
Sub mBuildCntrDetail(ilVefCode As Integer, ilBucketIndex As Integer, ilRdf As Integer, slAirMissedFlag As String, tlCntTypes As CNTTYPES, tlBkKey As BOOKKEY, tlAvRdf() As RDF, ilSpotOK As Integer, ilOrphanMissedLoop As Integer)
    Dim ilRet As Integer
    Dim llActPrice As Long
    Dim slPrice As String
    Dim slTempDays As String
    Dim slDysTms As String
    Dim slDays As String
    Dim illoop As Integer
    Dim ilShowOVTimes As Integer
    Dim ilShowOVDays As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilUniqueSpots As Integer
    Dim slStr As String
    Dim ilFound As Integer
    Dim tlCharCurr As KEYCHAR
    Dim tlCharPrev As KEYCHAR
    Dim ilRealRdf As Integer
    Dim ilTemp As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilMatchSSCode As Integer
    Dim ilSSOK As Integer
    Dim ilXMid As Integer
    Dim llrunningEndtime As Long
    Dim llrunningStartTime As Long
    Dim slOVTemp As String
    Dim ilLoop3 As Integer

    If ilSpotOK Then
        'Test header exclusions (types of contrcts and statuses)
        ilSSOK = True                                           'assume all Sales sources selected
        If Not RptSelQB!ckcAllSS.Value = vbChecked Then         'all sales selected, no need to do all the checkiing
            'Determine the sales source of the contract
            ilSSOK = False
            ilMatchSSCode = 0
            tmSlfSrchKey.iCode = tmChf.iSlfCode(0)
            ilRet = btrGetEqual(hmSlf, tmSlf, imSlfRecLen, tmSlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            For illoop = LBound(tmSofList) To UBound(tmSofList)
                If tmSofList(illoop).iSofCode = tmSlf.iSofCode Then     'get the matching office for the contracts slsp
                    ilMatchSSCode = tmSofList(illoop).iMnfSSCode          'Sales source
                    Exit For
                End If
            Next illoop
            For ilTemp = 0 To RptSelQB!lbcSelection(2).ListCount - 1
                If RptSelQB!lbcSelection(2).Selected(ilTemp) Then
                    slNameCode = tgMNFCodeRpt(ilTemp).sKey         'sales source code
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    If Val(slCode) = ilMatchSSCode Then      'sales source of this contract match one selected?
                        ilSSOK = True
                        Exit For
                    End If
                End If
            Next ilTemp
        End If

        If Not ilSSOK Then              'invalid sales source, dont continue
            ilSpotOK = False
            Exit Sub
        End If

        If tmChf.sStatus = "H" Then
            If Not tlCntTypes.iHold Then
                ilSpotOK = False
            End If
        ElseIf tmChf.sStatus = "O" Then
            If Not tlCntTypes.iOrder Then
                ilSpotOK = False
            End If
        Else
            ilSpotOK = False
        End If
        
        '3-16-10 wrong code was tested for standard (tested S not C)
        If tmChf.sType = "C" And Not tlCntTypes.iStandard Then      'include Standard types?
            ilSpotOK = False
        End If
        If tmChf.sType = "V" And Not tlCntTypes.iReserv Then      'include reservations ?
            ilSpotOK = False
        End If
        If tmChf.sType = "R" And Not tlCntTypes.iDR Then      'include DR?
            ilSpotOK = False
        End If

        If tmSdf.lChfCode <> tmChf.lChfCode Then
            tmClfSrchKey.lChfCode = tmSdf.lChfCode
            tmClfSrchKey.iLine = tmSdf.iLineNo
            tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
            tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
            ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
        End If
        If (tmClf.lChfCode <> tmSdf.lChfCode) Then     'got the matching line reference?
            ilSpotOK = False
        End If
        'Retrieve spot cost from flight ; flight not returned if spot type is Extra/Fill
        'otherwise flight returned in tgPriceCff
        ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slPrice)

        llActPrice = gStrDecToLong(slPrice, 2)
        llActPrice = gGetGrossNetTNetFromPrice(imGrossNet, llActPrice, tmClf.lAcquisitionCost, tmChf.iAgfCode)

        If llActPrice = 0 And Not tlCntTypes.iNC Then   'exclude NC
            ilSpotOK = False
        End If
        If ilSpotOK Then
            'read schedule line to build DP or override days & times
            'Build key to test for unique record

            tlBkKey.iRdfCode = tlAvRdf(ilRdf).iCode               'dp code
            tlBkKey.lChfCode = tmSdf.lChfCode                     'contr code
            tlBkKey.lFsfCode = 0                            'feed code
            tlBkKey.lRate = llActPrice                      'line spot rate
            tlBkKey.ivefSellCode = tmClf.iVefCode                'selling vehicle
            tlBkKey.iLen = tmSdf.iLen                       'spot length
            tlBkKey.sAirMissed = Trim$(slAirMissedFlag)                          'aired flag (vs missed)
            tlBkKey.sSpotType = tmSdf.sSpotType
            If tmClf.sType = "H" And tmSdf.sSchStatus = "G" Then                       'hidden line, indicate that this spot is from package
                tlBkKey.iPkgFlag = tmClf.iPkLineNo          'update with hidden line ID, (for now it will only indicate that its a pkg)
            Else
                tlBkKey.iPkgFlag = 0
            End If
            tlBkKey.sPriceType = tmSdf.sPriceType
            '2-13-03
            If tmSdf.sSpotType <> "X" Then              'not fill spot
                If tgPriceCff.sPriceType <> "T" Then               'dont use line price, its a different kind of rate (adu, spinoff, etc)
                    'flight returned in tgPriceCff
                    tlBkKey.sPriceType = tgPriceCff.sPriceType
                End If
            End If

            If tmSdf.sSpotType = "X" Then                   'extra or bonus
                tlBkKey.sDysTms = tlAvRdf(ilRdf).sName
                tlBkKey.ivefSellCode = ilVefCode        '8-22-00 Combine all fills for the same cnt on same line (previously showed on different output lines
                                                        'because the fills came from different vehicles
            Else
                If tgPriceCff.sDyWk = "W" Then
                    slTempDays = gDayNames(tgPriceCff.iDay(), tgPriceCff.sXDay(), 2, slStr)
                    slDysTms = ""
                    'Retrieve the days this flight is to air, strip out the commas & blanks from the text string
                    For illoop = 1 To Len(slTempDays) Step 1
                        slDays = Mid$(slTempDays, illoop, 1)
                        If slDays <> "" And slDays <> "," Then
                            slDysTms = Trim$(slDysTms) & Trim$(slDays)
                        End If
                    Next illoop
                Else                '11-19-02
                    'Setup # spots/day
                    slDysTms = ""
                    For illoop = 0 To 6
                        slDysTms = slDysTms + " " + Format$(str(tgPriceCff.iDay(illoop)), "0")
                    Next illoop
                End If

                ilRealRdf = gBinarySearchRdf(tmClf.iRdfCode)
                If ilRealRdf = -1 Then
                    tlBkKey.sDysTms = "Unknown DP"
                End If
               
                For illoop = 0 To 6 Step 1                  'see if there are override days from the flights compared to ordered DP
'                    If tgMRdf(ilRealRdf).sWkDays(7, ilLoop+1) = "Y" And tgPriceCff.iDay(ilLoop) <> 0 Then             'is DP a valid day
                    If tgMRdf(ilRealRdf).sWkDays(6, illoop) = "Y" And tgPriceCff.iDay(illoop) <> 0 Then             'is DP a valid day- chg for 0 based
                        ilShowOVDays = False
                    Else
                        ilShowOVDays = True
                        Exit For
                    End If
                Next illoop
                'see if override times compared to DP
                ilShowOVTimes = False
                If ((tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And ((tmClf.iEndTime(0) <> 1) Or (tmClf.iEndTime(1) <> 0)) Then
                    gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slStartTime
                    gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slEndTime
                    ilShowOVTimes = True
                    slOVTemp = Trim$(slStartTime) & "-" & Trim$(slEndTime)
                Else
                    ilXMid = False
                    'if there are multiple segments and it crosses midnight, show the earliest start time
                    'and xmidnight end time
                    slOVTemp = ""
                    For illoop = LBound(tgMRdf(ilRealRdf).iStartTime, 2) To UBound(tgMRdf(ilRealRdf).iStartTime, 2) Step 1
                        If (tgMRdf(ilRealRdf).iStartTime(0, illoop) <> 1) Or (tgMRdf(ilRealRdf).iStartTime(1, illoop) <> 0) Then
                            gUnpackTime tgMRdf(ilRealRdf).iStartTime(0, illoop), tgMRdf(ilRealRdf).iStartTime(1, illoop), "A", "1", slStartTime
                            gUnpackTime tgMRdf(ilRealRdf).iEndTime(0, illoop), tgMRdf(ilRealRdf).iEndTime(1, illoop), "A", "1", slEndTime
                            gUnpackTimeLong tgMRdf(ilRealRdf).iEndTime(0, illoop), tgMRdf(ilRealRdf).iEndTime(1, illoop), True, llrunningEndtime
                            'Exit For
                            If llrunningEndtime = 86400 Then    'its 12M end of day
                                ilXMid = True
                            End If
                            For ilLoop3 = illoop + 1 To UBound(tgMRdf(ilRealRdf).iStartTime, 2)
                                gUnpackTimeLong tgMRdf(ilRealRdf).iStartTime(0, ilLoop3), tgMRdf(ilRealRdf).iStartTime(1, ilLoop3), False, llrunningStartTime
                                 If llrunningStartTime = 0 And llrunningEndtime = 86400 Then
                                    If ilXMid Then
                                        gUnpackTime tgMRdf(ilRealRdf).iEndTime(0, ilLoop3), tgMRdf(ilRealRdf).iEndTime(1, ilLoop3), "A", "1", slEndTime
                                        Exit For
                                    End If
                                Else
                                    gUnpackTimeLong tgMRdf(ilRealRdf).iEndTime(0, ilLoop3), tgMRdf(ilRealRdf).iEndTime(1, ilLoop3), True, llrunningEndtime
                                End If
                            Next ilLoop3
                            
                            '4-12-10 show all times defined in the DP when there is an override
                            If slOVTemp = "" Then
                                slOVTemp = slStartTime & "-" & slEndTime
                            Else
                                slOVTemp = slOVTemp & "," & slStartTime & "-" & slEndTime
                            End If
                            If ilXMid Then          'cross midnight time
                                ilShowOVTimes = True
                                Exit For
                            End If
                        End If
                    Next illoop
                End If
                If (ilShowOVDays Or ilShowOVTimes) And ilOrphanMissedLoop <= 2 Then     '6-6-00 ilorphan = 0 or 1 its a sched spot , 2 missed spot with a DP it falls into,or 3 missed spot or orphan missed
                    'tlBkKey.sDysTms = slDysTms & " " & Trim$(slStartTime) & "-" & Trim$(slEndTime)
                    tlBkKey.sDysTms = slDysTms & " " & Trim$(slOVTemp)
                Else
                    'tlBkKey.sDysTms = Trim$(tlAvRdf(ilRdf).sName)
                    'Get the name of the DP defined from the line
                    tlBkKey.sDysTms = "Missing DP Code" & Trim$(str$(tmClf.iRdfCode))
                    'For ilRealRdf = LBound(tgMRdf) To UBound(tgMRdf) - 1 Step 1
                    '    If tgMRdf(ilRealRdf).iCode = tmClf.iRdfcode Then
                        ilRealRdf = gBinarySearchRdf(tmClf.iRdfCode)
                        If ilRealRdf <> -1 Then
                            tlBkKey.sDysTms = Trim$(tgMRdf(ilRealRdf).sName)
                    '        Exit For
                        End If
                    'Next ilRealRdf
                End If
            End If                              'SpotType = "X"

            ilFound = False
            LSet tlCharCurr = tlBkKey
            For ilUniqueSpots = 0 To UBound(tmBooked) - 1 Step 1
                LSet tlCharPrev = tmBooked(ilUniqueSpots).BkKeyRec
                If StrComp(tlCharCurr.sChar, tlCharPrev.sChar, 0) = 0 Then
                    ilFound = True
                    'Accum the spot counts in the proper week
                    tmBooked(ilUniqueSpots).iSpotCounts(ilBucketIndex) = tmBooked(ilUniqueSpots).iSpotCounts(ilBucketIndex) + 1
                    Exit For
                End If
            Next ilUniqueSpots
            If Not ilFound Then
                ilUniqueSpots = UBound(tmBooked)
                tmBooked(ilUniqueSpots).BkKeyRec = tlBkKey
                'accum # spots booked for the processing week
                tmBooked(ilUniqueSpots).iSpotCounts(ilBucketIndex) = 1
                ReDim Preserve tmBooked(0 To ilUniqueSpots + 1)
            End If
        End If                              'ilspotOK
    End If                                  'ilspotOK
End Sub

Public Sub mBuildFeedDetail(ilVefCode As Integer, ilBucketIndex As Integer, ilRdf As Integer, slAirMissedFlag As String, tlCntTypes As CNTTYPES, tlBkKey As BOOKKEY, tlAvRdf() As RDF, ilSpotOK As Integer, ilOrphanMissedLoop As Integer)
    Dim slTempDays As String
    Dim slDysTms As String
    Dim slDays As String
    Dim illoop As Integer
    Dim ilShowOVTimes As Integer
    Dim ilShowOVDays As Integer
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilUniqueSpots As Integer
    Dim slStr As String
    Dim ilFound As Integer
    Dim tlCharCurr As KEYCHAR
    Dim tlCharPrev As KEYCHAR
    Dim slWorkDys(0 To 6)  As String * 1        'temp for valid days of week for generalized date rtn
    Dim llrunningEndtime As Long
    Dim llrunningStartTime As Long
    Dim slOVTemp As String
    Dim ilLoop3 As Integer
    Dim ilRealRdf As Integer
        
    If ilSpotOK Then
        'read schedule line to build DP or override days & times
        'Build key to test for unique record

        tlBkKey.iRdfCode = tlAvRdf(ilRdf).iCode               'dp code
        tlBkKey.lFsfCode = tmSdf.lFsfCode                     'feed spot code
        tlBkKey.lChfCode = 0                                  'contr code
        tlBkKey.lRate = 0                                   'feed spots dont have a $
        tlBkKey.ivefSellCode = tmFsf.iVefCode                'selling vehicle
        tlBkKey.iLen = tmSdf.iLen                       'spot length
        tlBkKey.sAirMissed = Trim$(slAirMissedFlag)                          'aired flag (vs missed)
        tlBkKey.sSpotType = tmSdf.sSpotType
        tlBkKey.iPkgFlag = 0                            'hidden line falg
        tlBkKey.sPriceType = "F"                        'tmsdf.sPriceType :  force to Feed

        If tmFsf.sDyWk = "W" Then
            For illoop = 0 To 6
                slWorkDys(illoop) = ""
            Next illoop
            slTempDays = gDayNames(tmFsf.iDays(), slWorkDys(), 2, slStr)
            slDysTms = ""
            'Retrieve the days this flight is to air, strip out the commas & blanks from the text string
            For illoop = 1 To Len(slTempDays) Step 1
                slDays = Mid$(slTempDays, illoop, 1)
                If slDays <> "" And slDays <> "," Then
                    slDysTms = Trim$(slDysTms) & Trim$(slDays)
                End If
            Next illoop
        Else                '11-19-02
            'Setup # spots/day
            slDysTms = ""
            For illoop = 0 To 6
                slDysTms = slDysTms + " " + Format$(str(tmFsf.iDays(illoop)), "0")
            Next illoop
        End If

        ilRealRdf = gBinarySearchRdf(tmClf.iRdfCode)
        If ilRealRdf = -1 Then
            tlBkKey.sDysTms = "Unknown DP"
        End If
                
        For illoop = 0 To 6 Step 1                  'see if there are override days from the flights compared to DP
            'If tgMRdf(ilRdf).sWkDays(7, ilLoop + 1) = "Y" Then           'is DP a valid day
            If tgMRdf(ilRdf).sWkDays(6, illoop) = "Y" Then           'is DP a valid day
                If tmFsf.iDays(illoop) >= 0 Then
                    ilShowOVDays = True
                    Exit For
                Else
                    ilShowOVDays = False
                End If
            End If
        Next illoop
        'see if override times compared to DP
        ilShowOVTimes = False
        If ((tmFsf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And ((tmFsf.iEndTime(0) <> 1) Or (tmClf.iEndTime(1) <> 0)) Then
            gUnpackTime tmFsf.iStartTime(0), tmFsf.iStartTime(1), "A", "1", slStartTime
            gUnpackTime tmFsf.iEndTime(0), tmFsf.iEndTime(1), "A", "1", slEndTime
            ilShowOVTimes = True
        Else
            For illoop = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1
                If (tgMRdf(ilRdf).iStartTime(0, illoop) <> 1) Or (tgMRdf(ilRdf).iStartTime(1, illoop) <> 0) Then
                    gUnpackTime tgMRdf(ilRdf).iStartTime(0, illoop), tgMRdf(ilRdf).iStartTime(1, illoop), "A", "1", slStartTime
                    gUnpackTime tgMRdf(ilRdf).iEndTime(0, illoop), tgMRdf(ilRdf).iEndTime(1, illoop), "A", "1", slEndTime
                    Exit For
                End If
            Next illoop
        End If
        If (ilShowOVDays Or ilShowOVTimes) And ilOrphanMissedLoop <= 2 Then     '6-6-00 ilorphan = 0 or 1 its a sched spot , 2 missed spot with a DP it falls into,or 3 missed spot or orphan missed
            tlBkKey.sDysTms = slDysTms & " " & Trim$(slStartTime) & "-" & Trim$(slEndTime)
        Else
            'Booked dayparts dont exist for feed spots,
            tlBkKey.sDysTms = ""

            'Get the name of the DP defined from the line
            'tlBkKey.sDysTms = "Missing DP Code" & Trim$(Str$(tmClf.iRdfcode))
            'ilRealRdf = gBinarySearchRdf(tmClf.iRdfcode)
            'If ilRealRdf <> -1 Then
            '    tlBkKey.sDysTms = Trim$(tgMRdf(ilRealRdf).sName)
            'End If

        End If


        ilFound = False
        LSet tlCharCurr = tlBkKey
        For ilUniqueSpots = 0 To UBound(tmBooked) - 1 Step 1
            LSet tlCharPrev = tmBooked(ilUniqueSpots).BkKeyRec
            If StrComp(tlCharCurr.sChar, tlCharPrev.sChar, 0) = 0 Then
                ilFound = True
                'Accum the spot counts in the proper week
                tmBooked(ilUniqueSpots).iSpotCounts(ilBucketIndex) = tmBooked(ilUniqueSpots).iSpotCounts(ilBucketIndex) + 1
                Exit For
            End If
        Next ilUniqueSpots
        If Not ilFound Then
            ilUniqueSpots = UBound(tmBooked)
            tmBooked(ilUniqueSpots).BkKeyRec = tlBkKey
            'accum # spots booked for the processing week
            tmBooked(ilUniqueSpots).iSpotCounts(ilBucketIndex) = 1
            ReDim Preserve tmBooked(0 To ilUniqueSpots + 1)
        End If

    End If                                  'ilspotOK
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainPCFVehicles              *
'*          Digital Vehicle                            *
'*                                                     *
'*             Created:6/14/23       By:J. White       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the Digital VEFCodes     *
'*                     For PCF Lines between given     *
'*                     Start and End Dates             *
'*            populates imVefList with vehicle codes   *
'*******************************************************
'TTP 10729 - Quarterly Booked Spots report: add digital lines
Sub mObtainPCFVehicles(slStartDate As String, slEndDate As String)
'   where:
'       slStartDate     - Digital Line start date
'       slEndDate       - Digital Line end date
    Dim slSQLQuery As String
    Dim rst As ADODB.Recordset
    Dim illoop As Long
    Dim llUpper As Integer
    ReDim imVefList(0 To 0) As Integer

    slSQLQuery = ""
    slSQLQuery = slSQLQuery & "SELECT DISTINCT pcfVefCode from pcf_Pod_CPM_Cntr"
    slSQLQuery = slSQLQuery & " WHERE "
    slSQLQuery = slSQLQuery & "  pcfStartDate <= '" & Format(slEndDate, "yyyy-mm-dd") & "'"
    slSQLQuery = slSQLQuery & " AND "
    slSQLQuery = slSQLQuery & "  pcfEndDate >= '" & Format(slStartDate, "yyyy-mm-dd") & "'"
    
    Set rst = gSQLSelectCall(slSQLQuery)
    If Not rst.EOF Then
        Do While Not rst.EOF
            llUpper = UBound(imVefList)
            
            imVefList(llUpper) = rst.Fields("pcfVefCode").Value
            ReDim Preserve imVefList(0 To llUpper + 1) As Integer
            
            rst.MoveNext
        Loop
    End If
    rst.Close
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainSelPCF                   *
'*          Digital Lines by Vehicle                   *
'*                                                     *
'*             Created:7/3/23        By:J. White       *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Obtain the PCF records to be    *
'*                     reported                        *
'*                                                     *
'*******************************************************
'TTP 10729 - Quarterly Booked Spots report: add digital lines
Sub mObtainSelPcf(ilWhichKey, llItemCode As Integer, slStartDate As String, slEndDate As String, tlCntTypes As CNTTYPES)
'   where:
'       ilWhichKey      - 0=Contracts, 1=Agency, 5=Advertiser, 6=Vehicles
'       llItemCode      - the value of the Contr,Agency,Adv or Vehicle (Depending on ilWhichKey)
'       slStartDate     - Digital Line start date
'       slEndDate       - Digital Line end date
'       slCntrStartDate - Contract entered start date
'       slCntrEndDate   - Contract entered End date
'       ilSelType       - 0=Advertiser, 1=Agency; 2=Salesperson; 3=No selection
'       tmSelChf        - contains the selections
    Dim blValid As Boolean
    Dim slSQLQuery As String
    Dim rst As ADODB.Recordset
    Dim illoop As Long
    Dim llUpper As Long
    Dim blNeedComma As Boolean
    Dim ilTemp As Integer
    Dim ilSSOK As Integer
    Dim slNameCode As String
    Dim slCode As String
    Dim ilRet As Integer
    
    slSQLQuery = ""
    slSQLQuery = slSQLQuery & "SELECT "
    slSQLQuery = slSQLQuery & "  chf.chfCode,"
    slSQLQuery = slSQLQuery & "  chf.chfBillCycle,"
    slSQLQuery = slSQLQuery & "  chf.chfCntrNo,"
    slSQLQuery = slSQLQuery & "  chf.chfExtCntrNo,"
    slSQLQuery = slSQLQuery & "  chf.chfType,"
    slSQLQuery = slSQLQuery & "  chf.chfStartDate,"
    slSQLQuery = slSQLQuery & "  chf.chfEndDate,"
    slSQLQuery = slSQLQuery & "  chf.chfAgfCode,"
    slSQLQuery = slSQLQuery & "  chf.chfAdfCode,"
    slSQLQuery = slSQLQuery & "  chf.chfProduct,"
    slSQLQuery = slSQLQuery & "  chf.chfSlfCode1,"
    slSQLQuery = slSQLQuery & "  slf.slfsofCode,"
    slSQLQuery = slSQLQuery & "  SOF.sofmnfSSCode,"
    slSQLQuery = slSQLQuery & "  pcf.pcfCode,"
    slSQLQuery = slSQLQuery & "  pcf.pcfPodCPMID,"
    slSQLQuery = slSQLQuery & "  pcf.pcfVefCode,"
    slSQLQuery = slSQLQuery & "  pcf.pcfPriceType,"
    slSQLQuery = slSQLQuery & "  pcf.pcfStartDate,"
    slSQLQuery = slSQLQuery & "  pcf.pcfEndDate,"
    slSQLQuery = slSQLQuery & "  pcf.pcfPodCPM,"
    slSQLQuery = slSQLQuery & "  pcf.pcfRdfCode," 'Daypart
    slSQLQuery = slSQLQuery & "  pcf.pcfCxfCode," 'Line Comment
    
    slSQLQuery = slSQLQuery & "  pcf.pcfTotalCost"
    slSQLQuery = slSQLQuery & " FROM "
    slSQLQuery = slSQLQuery & "  CHF_Contract_Header chf"
    slSQLQuery = slSQLQuery & "  JOIN pcf_Pod_CPM_Cntr pcf on pcf.pcfChfCode = chf.chfCode"
    slSQLQuery = slSQLQuery & "  JOIN SLF_Salespeople slf on slf.SlfCode = chf.chfSlfCode1"
    slSQLQuery = slSQLQuery & "  JOIN SOF_Sales_Offices sof on sof.sofCode = slf.slfsofcode"
    slSQLQuery = slSQLQuery & " WHERE "
    
    'Contract Status
    slSQLQuery = slSQLQuery & "  chf.chfStatus in ("
    blNeedComma = False
    'Chf.sStatus = "H" tlCntTypes.iHold
    If tlCntTypes.iHold Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'H'" 'Holds
        blNeedComma = True
    End If
    'Chf.sStatus = "O" tlCntTypes.iOrder
    If tlCntTypes.iOrder Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'O'" 'Orders (Scheduled Order)
        blNeedComma = True
    End If
    'Chf.sType = "C" tlCntTypes.iStandard
    If tlCntTypes.iStandard Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'C'" 'include Standard Orders (Scheduled Order)
        blNeedComma = True
    End If
    'Chf.sType = "V" tlCntTypes.iReserv
    If tlCntTypes.iReserv Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'V'" 'include reservations ?
        blNeedComma = True
    End If
    'Chf.sType = "R" tlCntTypes.iDR
    If tlCntTypes.iDR Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'R'" 'include DR?
        blNeedComma = True
    End If
    'Chf.sType = "S"  tlCntTypes.iPSA
    If tlCntTypes.iPSA Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'S'" 'include PSA ?
        blNeedComma = True
    End If
    'Chf.sType = "M"  tlCntTypes.iPromo
    If tlCntTypes.iPromo Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'M'" 'include Promo?
        blNeedComma = True
    End If
    'Chf.sType = "T"  tlCntTypes.iRemnant
    If tlCntTypes.iRemnant Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'T'" 'include Remnant?
        blNeedComma = True
    End If
    'Chf.sType = "Q" tlCntTypes.iPI
    If tlCntTypes.iPI Then
        If blNeedComma Then slSQLQuery = slSQLQuery & ","
        slSQLQuery = slSQLQuery & "'Q'" 'include PI?
        blNeedComma = True
    End If
    slSQLQuery = slSQLQuery & ")"
    
    slSQLQuery = slSQLQuery & "  AND chf.chfDelete <> 'Y'"
    If ilWhichKey = 0 Then 'by Contract
        slSQLQuery = slSQLQuery & "  AND chf.chfCntrNo = " & llItemCode
    ElseIf ilWhichKey = 1 Then 'by Agency
        slSQLQuery = slSQLQuery & "  AND chf.chfAgfCode = " & llItemCode
    ElseIf ilWhichKey = 5 Then 'by Advertiser
        slSQLQuery = slSQLQuery & "  AND chf.chfAdfCode = " & llItemCode
    ElseIf ilWhichKey = 6 Then 'by Vehicle
        slSQLQuery = slSQLQuery & "  AND pcf.pcfVefCode = " & llItemCode
    End If
    
    If slEndDate <> "" Then
        slSQLQuery = slSQLQuery & "  AND pcf.pcfStartDate <= '" & Format(slEndDate, "yyyy-mm-dd") & "'"
    End If
    If slStartDate <> "" Then
        slSQLQuery = slSQLQuery & "  AND pcf.pcfEndDate >= '" & Format(slStartDate, "yyyy-mm-dd") & "'"
    End If
    slSQLQuery = slSQLQuery & "  AND pcf.pcfStartDate <= pcf.pcfEndDate " 'No CBS lines
    slSQLQuery = slSQLQuery & "  AND pcf.pcfDelete <> 'Y'"
    
    Set rst = gSQLSelectCall(slSQLQuery)
    If Not rst.EOF Then
        Do While Not rst.EOF
            '--------------------------------------------
            'filter contract to selected Sales Source(s)
            ilSSOK = False
            If RptSelQB!ckcAllSS.Value = vbChecked Then
                ilSSOK = True
            Else
                For ilTemp = 0 To RptSelQB!lbcSelection(2).ListCount - 1
                    If RptSelQB!lbcSelection(2).Selected(ilTemp) Then
                        slNameCode = tgMNFCodeRpt(ilTemp).sKey         'sales source code
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        If Val(slCode) = rst.Fields("sofmnfSSCode").Value Then      'sales source of this contract match one selected?
                            ilSSOK = True
                            Exit For
                        End If
                    End If
                Next ilTemp
            End If
            
            '--------------------------------------------
            'Add pcf record to tmPLPcf array
            If ilSSOK Then
                llUpper = UBound(tmPLPcf)
                'Contract Info
                tmPLPcf(llUpper).lCntrNo = rst.Fields("chfCntrNo").Value
                'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 5)
                tmPLPcf(llUpper).lExtCntrNo = rst.Fields("chfExtCntrNo").Value
                tmPLPcf(llUpper).sCalType = rst.Fields("chfBillCycle").Value
                tmPLPcf(llUpper).sCntrStartDate = rst.Fields("chfStartDate").Value
                tmPLPcf(llUpper).sCntrEndDate = rst.Fields("chfEndDate").Value
                Select Case rst.Fields("chfType").Value
                    Case "C": tmPLPcf(llUpper).sContractType = "Standard"
                    Case "V": tmPLPcf(llUpper).sContractType = "Reservation"
                    Case "T": tmPLPcf(llUpper).sContractType = "Remnant"
                    Case "R": tmPLPcf(llUpper).sContractType = "Direct Response"
                    Case "Q": tmPLPcf(llUpper).sContractType = "Per Inquiry" 'RE: Test(B2603):v8.1 Test Traffic & Affiliate, 6/14/23 (Issue 7)
                    Case "S": tmPLPcf(llUpper).sContractType = "PSA"
                    Case "M": tmPLPcf(llUpper).sContractType = "Promo"
                    Case Else: tmPLPcf(llUpper).sContractType = Trim(rst.Fields("chfType").Value)
                End Select
                tmPLPcf(llUpper).iAgfCode = rst.Fields("chfAgfCode").Value
                tmPLPcf(llUpper).iAdfCode = rst.Fields("chfAdfCode").Value
                tmPLPcf(llUpper).sProduct = rst.Fields("chfProduct").Value
                'Line Info
                tmPLPcf(llUpper).tPcf.lCode = rst.Fields("pcfCode").Value
                tmPLPcf(llUpper).tPcf.iVefCode = rst.Fields("pcfVefCode").Value
                tmPLPcf(llUpper).tPcf.lChfCode = rst.Fields("chfCode").Value
                tmPLPcf(llUpper).tPcf.lTotalCost = rst.Fields("pcfTotalCost").Value
                tmPLPcf(llUpper).tPcf.iPodCPMID = rst.Fields("pcfPodCPMID").Value
                gPackDate rst.Fields("pcfEndDate").Value, tmPLPcf(llUpper).tPcf.iEndDate(0), tmPLPcf(llUpper).tPcf.iEndDate(1)
                gPackDate rst.Fields("pcfStartDate").Value, tmPLPcf(llUpper).tPcf.iStartDate(0), tmPLPcf(llUpper).tPcf.iStartDate(1)
                tmPLPcf(llUpper).tPcf.sPriceType = rst.Fields("pcfPriceType").Value
                tmPLPcf(llUpper).tPcf.lCxfCode = rst.Fields("pcfCxfCode").Value
                tmPLPcf(llUpper).tPcf.iRdfCode = rst.Fields("pcfRdfCode").Value
                ReDim Preserve tmPLPcf(0 To llUpper + 1) As PCFTYPESORT
            End If

nextRec:
            rst.MoveNext
        Loop
    End If
    rst.Close
    Exit Sub
End Sub

'How many days does line run?
'Provide Start/End date of Line
'Provide Start/End date of period to limit results to
Function mGetNumberOfDaysRunning(slStartDate As String, slEndDate As String, Optional slStartDate2 As String, Optional slEndDate2 As String) As Integer
    Dim sltmpStartDate As String
    Dim sltmpEndDate As String
    sltmpStartDate = slStartDate
    If slStartDate2 <> "" Then
        If DateValue(slStartDate2) > DateValue(slStartDate) Then sltmpStartDate = slStartDate2
    End If
    sltmpEndDate = slEndDate
    If slEndDate2 <> "" Then
        If DateValue(slEndDate2) < DateValue(slEndDate) Then sltmpEndDate = slEndDate2
    End If
    
    If DateValue(sltmpStartDate) <= DateValue(sltmpEndDate) Then
        mGetNumberOfDaysRunning = DateDiff("d", DateValue(sltmpStartDate), DateValue(sltmpEndDate)) + 1
    End If
    If mGetNumberOfDaysRunning < 0 Then mGetNumberOfDaysRunning = 0
End Function
