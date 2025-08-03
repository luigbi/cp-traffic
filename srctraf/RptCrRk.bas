Attribute VB_Name = "RptCrRk"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptcrRk.bas on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Private Variables (Removed)                                                            *
'*  tmRdfSrchKey                                                                          *
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
Dim hmVsf As Integer            'Vehicle file handle
Dim tmVsf As VSF                'VSF record image
Dim imVsfRecLen As Integer        'VSF record length
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
Dim tmGrf As GRF
Dim hmGrf As Integer
Dim imGrfRecLen As Integer        'GPF record length
Dim hmMnf As Integer            'Multiname file handle
Dim imMnfRecLen As Integer      'MNF record length
Dim tmMnf As MNF
Dim tmRcf As RCF
Dim hmRdf As Integer            'Dayparts file handle
Dim imRdfRecLen As Integer      'RD record length
Dim tmRdf As RDF

Dim hmFsf As Integer            'Feed file handle
Dim tmFSFSrchKey As LONGKEY0    'FSF search key
Dim imFsfRecLen As Integer      'FSF record length
Dim tmFsf As FSF                'FSF record buffer
Dim hmAnf As Integer            'Named avail file handle
Dim tmAnfSrchKey As INTKEY0    'ANF record image
Dim imAnfRecLen As Integer      'ANF record length
Dim tmAnf As ANF

Dim hmGsf As Integer            'Game file handle
Dim tmGsfSrchKey3 As GSFKEY3    'vehicle, then game #
Dim imGsfRecLen As Integer      'GSF record length
Dim tmGsf As GSF                'GSF image

Dim hmGhf As Integer            'Game file handle
Dim tmGhfSrchKey As LONGKEY0    'key by code
Dim imGhfRecLen As Integer      'GHF record length
Dim tmGhf As GHF                'GHF image

Dim hmIsf As Integer            'Multimedia inventory file handle
Dim tmIsfSrchKey1 As ISFKEY1    'key by ghfcode, gameno
Dim imIsfRecLen As Integer      'ISF record length
Dim tmIsf As ISF                'ISF image

Dim hmMsf As Integer            'Multimedia sold file handle
Dim tmMsfSrchKey1 As MSFKEY1    'key by vehicle
Dim imMsfRecLen As Integer      'MSF record length
Dim tmMsf As MSF                'MSF image

Dim hmMgf As Integer            'Multimedia sold by game file handle
Dim tmMgfSrchKey1 As MGFKEY1    'key by msfcode, gameno
Dim imMgfRecLen As Integer      'MGF record length
Dim tmMgf As MGF                'MGF image


Dim imGameIndFlag As Integer       'for Multimedia info only;  flag to indicate to process the multimedia
                                'only once per vehicle otherwise inventory/sold will be overstated
                            
'User input options
Dim imGrossNetTNet As Integer    '0 = gross, 1 = net, 2 = tnet
Dim imTotsByDPVef As Integer     '0 = totals by DP, 1 = totals by vehicle
Dim imUnder30 As Boolean        'true to include everything, else false to exclude any any and/or spot under 30"
Dim imWhichSort As Integer      'column to sort top down
Dim lmContract As Long

Dim imNamedAvails() As Integer      'list of named avails to include
Dim tmStatsByDP() As STATSBYDP
Dim tmStatsByAdvt() As STATSBYADVT
Type STATSBYADVT                'structure for sold and revenue status for an advertiser within dp/vehicle
    sKey As String * 30         'vehicle code|dp code|sort top down by (% of inv, % of rev,# units sched, Avg 30" rate or Gross rev).  Fields are negated to sort descending
                                '00125|00200|9999999149     (9999999999-value of sort field column)
                                'all like DP together, followed by the advt sort field
    iType As Integer            '0 = DP, 1 = vehicle  (totals by dp or vehicle)
    lKeyCode As Long            'dp code (0 if by totals by vehicle)
    iDPInxToStatsByDP As Integer   'index into DP array
    iVefCode As Integer
    iAdfCode As Integer
    lLenSched As Long           'total seconds sched (used to calc #30 units)
    lRev As Long                'total $ scheduled
    lUnits As Long              'total units to calc avg rate
    iSortCode As Integer        'sort code from RIF if by dp, else 0 for games
    iPctUnitsSold As Integer         '% sold against original inv for this advt/dp
    iPctRevSold As Integer          '% rev sold against total rev for this advt /dp
    lAvgRate As Long                'avg 30" unit rate
    iGameNo As Integer
End Type

Type STATSBYDP                  'structure for overall inventory and revenue for a dp/vehicle
    iType As Integer            '0 = DP, 1 = vehicle  (totals by dp or vehicle)
    lKeyCode As Long            'dp code (0 if by totals by vehicle)
    iVefCode As Integer
    iSortCode As Integer        'sort code from RIF if by dp, else 0 for games
    lInv As Long                'Inventory of seconds for a dp/vehicle
    lUnits As Long              '# units of inventory
    lLenSched As Long           'total spot length scheduled for dp/vehicle (used to calc % of Inv for each advt in this DP)
    lRev As Long                'total revenue scheduled for dp/vehicle (used to calc % of rev for each advt in this DP)
    'iStartTime(0 To 1, 1 To 7) As Integer    'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iStartTime(0 To 1, 0 To 6) As Integer    'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
                        'If Hund Sec = 1 and all other times = 0, then time not defined
                        'In Basic the left most dimension expands first
    'iEndTime(0 To 1, 1 To 7) As Integer    'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iEndTime(0 To 1, 0 To 6) As Integer    'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    'sWkDays(1 To 7, 1 To 7)  As String * 1  'Dimension 1: Time; Dimension 2:Days;
    sWkDays(0 To 6, 0 To 6)  As String * 1  'Dimension 1: Time; Dimension 2:Days;
                                            'Dimension 2: Index 1= Monday, 2= tuesday,...
                                            'Weekday flag: Y=day allowed; N=Day disallowed
    ianfCode As Integer
    sInOut As String * 1
End Type



'
'       gFilterContractTypeFromSpot - determine if spot should be processed or ignored
'       based on user parameters set up in structure CNTTYPES.
'       The spot is from SSF and uses the iamge CSSPOTSS.
'       <input>  tlSpot -image of spot from SSF
'                tlCnttypes - contract types to include/exclude
'       Return - true to process spot, else ignore
'
'Private Function mFilterContractTypeFromSpot(tlSpot As CSPOTSS, tlCntTypes As CNTTYPES) As Integer
''******************************************************************************************
''* Note: VBC id'd the following unreferenced items and handled them as described:         *
''*                                                                                        *
''* Local Variables (Removed)                                                              *                                                                            *
''******************************************************************************************
'
'Dim ilSpotOK As Integer
'        ilSpotOK = True                 'assume the spot is ok to include
'        If (tlSpot.iRank And RANKMASK) = 1010 Then      'DR
'            If Not tlCntTypes.iDR Then
'                ilSpotOK = False
'            End If
'        ElseIf (tlSpot.iRank And RANKMASK) = 1020 Then
'            If Not tlCntTypes.iRemnant Then
'                ilSpotOK = False
'            End If
'
'        ElseIf (tlSpot.iRank And RANKMASK) = 1030 Then    'PI
'            If Not tlCntTypes.iPI Then
'                ilSpotOK = False
'            End If
'
''        ElseIf (tlSpot.iRank And RANKMASK) = 1040 Then  'trades            'test trade from header
''            If Not tlCntTypes.iTrade Then
''                ilSpotOK = False
''            End If
'
'        ElseIf (tlSpot.iRank And RANKMASK) = 1050 Then  'promo
'            If Not tlCntTypes.iPromo Then
'                ilSpotOK = False
'            End If
'
'        ElseIf (tlSpot.iRank And RANKMASK) = 1060 Then  'psa
'            If Not tlCntTypes.iPSA Then
'                ilSpotOK = False
'            End If
'        End If
'        mFilterContractTypeFromSpot = ilSpotOK
'        Exit Function
'End Function
'*******************************************************************
'*                                                                 *
'*      Procedure Name:gCreateSpotPriceRanking                          *
'*                                                                 *
'*             Created:7/30/12       By:D. Hosaka                  *
'*                                                                 *

Sub gCreateSpotPriceRanking()

    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim slDate As String
    Dim slStr As String
    Dim ilSaveMonth As Integer
    ReDim ilDate(0 To 1) As Integer
    Dim ilVehicle As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilIndex As Integer
    Dim llEndDate As Long
    Dim llDateEntered As Long
    Dim tlCntTypes As CNTTYPES
    Dim ilContinue As Integer
    Dim ilPeriodType As Integer     'type of period for monthly to gather (corp, std, cal; currently only std)
    'Dim llStartDates(1 To 13) As Long
    Dim llStartDates(0 To 13) As Long   'Index zero ignored
    Dim ilStartDate(0 To 1) As Integer
    Dim ilEndDate(0 To 1) As Integer
    Dim ilLastBilledInx As Integer      'used in subrtn to gather dates, return unused
    Dim llLastBilled As Long            'used in subrtn to gather dates, return unused
    Dim llDate As Long
    Dim ilHowManyPer As Integer
    Dim ilLoopOnListBox As Integer
    Dim llRank As Long
    Dim ilAtLeastOneDP As Integer       'if no DP procesed, output no date found message on report
    On Error GoTo gCreateSpotPricingRanking_Error
   
    hmCHF = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imCHFRecLen = Len(tmChf)

    hmGrf = CBtrvTable(TEMPHANDLE)  '(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGrf, "", sgDBPath & "Grf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imGrfRecLen = Len(tmGrf)

    hmClf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imClfRecLen = Len(tmClf)

    hmCff = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imCffRecLen = Len(tmCff)
    hmMnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imMnfRecLen = Len(tmMnf)
    hmVef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)
    hmSdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imSdfRecLen = Len(tmSdf)
    hmSsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    hmLcf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmLcf, "", sgDBPath & "Lcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imLcfRecLen = Len(tmLcf)
    hmSmf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imSmfRecLen = Len(tmSmf)
    hmVsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imVsfRecLen = Len(tmVsf)
    hmRdf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imRdfRecLen = Len(tmRdf)

    hmFsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imFsfRecLen = Len(tmFsf)

    hmAnf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmAnf, "", sgDBPath & "Anf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imAnfRecLen = Len(tmAnf)

    hmGsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
    ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        mCloseAll
        Exit Sub
    End If
    imGsfRecLen = Len(tmGsf)




    tlCntTypes.iHold = gSetCheck(RptSelRk!ckcCType(0).Value)
    tlCntTypes.iOrder = gSetCheck(RptSelRk!ckcCType(1).Value)
    tlCntTypes.iNetwork = gSetCheck(RptSelRk!ckcCType(2).Value)
    tlCntTypes.iStandard = gSetCheck(RptSelRk!ckcCType(3).Value)
    tlCntTypes.iReserv = gSetCheck(RptSelRk!ckcCType(4).Value)
    tlCntTypes.iRemnant = gSetCheck(RptSelRk!ckcCType(5).Value)
    tlCntTypes.iDR = gSetCheck(RptSelRk!ckcCType(6).Value)
    tlCntTypes.iPI = gSetCheck(RptSelRk!ckcCType(7).Value)
    tlCntTypes.iPSA = gSetCheck(RptSelRk!ckcCType(8).Value)
    tlCntTypes.iPromo = gSetCheck(RptSelRk!ckcCType(9).Value)
    tlCntTypes.iTrade = gSetCheck(RptSelRk!ckcCType(10).Value)
    tlCntTypes.iMissed = gSetCheck(RptSelRk!ckcSpots(0).Value)
    tlCntTypes.iCharge = gSetCheck(RptSelRk!ckcSpots(1).Value)
    tlCntTypes.iZero = gSetCheck(RptSelRk!ckcSpots(2).Value)
    tlCntTypes.iADU = gSetCheck(RptSelRk!ckcSpots(3).Value)
    tlCntTypes.iBonus = gSetCheck(RptSelRk!ckcSpots(4).Value)
    tlCntTypes.iXtra = gSetCheck(RptSelRk!ckcSpots(5).Value)
    tlCntTypes.iFill = gSetCheck(RptSelRk!ckcSpots(6).Value)
    tlCntTypes.iNC = gSetCheck(RptSelRk!ckcSpots(7).Value)
    tlCntTypes.iRecapturable = gSetCheck(RptSelRk!ckcSpots(8).Value)
    tlCntTypes.iSpinoff = gSetCheck(RptSelRk!ckcSpots(9).Value)
    tlCntTypes.iMG = gSetCheck(RptSelRk!ckcSpots(10).Value)
    tlCntTypes.iPolit = gSetCheck(RptSelRk!ckcSpots(11).Value)
    tlCntTypes.iNonPolit = gSetCheck(RptSelRk!ckcSpots(12).Value)
    tlCntTypes.iFeedSpots = False                                     'not dealing with feed spots

    If (tlCntTypes.iHold) Or (tlCntTypes.iOrder) Then        '1-26-05 set general cntr type for inclusion/exclusion if hold or ordered selected
        tlCntTypes.iCntrSpots = True
    Else
        tlCntTypes.iCntrSpots = False
    End If
    ilHowManyPer = Val(RptSelRk!edcSelCFrom1.Text)
    
     
    If RptSelRk!rbcRevType(1).Value Then             'net
        imGrossNetTNet = 1
    ElseIf RptSelRk!rbcRevType(2).Value Then         'tnet
        imGrossNetTNet = 2
    Else                                                'gross
        imGrossNetTNet = 0
    End If

    imTotsByDPVef = 0                       'totals by DP
    If RptSelRk!rbcTotalsBy(1).Value Then     'totals by vehicle
        imTotsByDPVef = 1
    End If
    
    imUnder30 = False
    If RptSelRk!ckcUnder30.Value = vbChecked Then        'include avails and/or spots under 30".  If so, will show fractions on report; otherwise no fractions
        imUnder30 = True
    End If
    
    imWhichSort = RptSelRk!cbcSort.ListIndex            'sort column for top down

    lmContract = 0
    If Trim$(RptSelRk!edcContract.Text) <> "" Then
        lmContract = Val(RptSelRk!edcContract.Text)     'selective contract
    End If
    
    If RptSelRk!rbcPeriodType(0).Value Then         'month (vs week)
         slStr = RptSelRk!edcSelCFrom.Text             'month in text form (jan..dec)
        gGetMonthNoFromString slStr, ilSaveMonth          'getmonth #
        If ilSaveMonth = 0 Then                           'input isn't text month name, try month #
            ilSaveMonth = Val(slStr)
        End If
       
        If RptSelRk!rbcMonthType(0).Value Then              'Calendar
            'Format the base date Month & year spans to send to Crystal
            slDate = Trim$(str$(ilSaveMonth)) & "/01/" & Trim$(RptSelRk!edcYear.Text)
            slDate = gObtainStartCal(slDate)
            llDateEntered = gDateValue(slDate)          'cal start
            ilPeriodType = 3
            'llLastBilled & ilLastBillInx not used in this report
            'create the array of start dates to get the end date of the requested periods
            'igYear contains Year that routine needs, igMonthOrQtr = Month to start getting dates
            gSetupBOBDates ilPeriodType, llStartDates(), llLastBilled, ilLastBilledInx, 0, igMonthOrQtr  'build array of start/end dates
            llEndDate = llStartDates(ilHowManyPer + 1) - 1
        ElseIf RptSelRk!rbcMonthType(1).Value Then          'std
            'Format the base date Month & year spans to send to Crystal
            slDate = Trim$(str$(ilSaveMonth)) & "/15/" & Trim$(RptSelRk!edcYear.Text)
            slDate = gObtainStartStd(slDate)
            llDateEntered = gDateValue(slDate)          'standard start
            ilPeriodType = 2
            'llLastBilled & ilLastBillInx not used in this report
            'create the array of start dates to get the end date of the requested periods
            'igYear contains Year that routine needs, igMonthOrQtr = Month to start getting dates
            gSetupBOBDates ilPeriodType, llStartDates(), llLastBilled, ilLastBilledInx, 0, igMonthOrQtr  'build array of start/end dates
            llEndDate = llStartDates(ilHowManyPer + 1) - 1
        Else                    'corp
             ilPeriodType = 1
            'llLastBilled & ilLastBillInx not used in this report
            'create the array of start dates to get the end date of the requested periods
            'igYear contains Year that routine needs, igMonthOrQtr = Month to start getting dates
            gSetupBOBDates ilPeriodType, llStartDates(), llLastBilled, ilLastBilledInx, 0, igMonthOrQtr  'build array of start/end dates
            llEndDate = llStartDates(ilHowManyPer + 1) - 1
        End If
     Else
        'slDate = RptSelRk!edcSelCFrom.Text                  'verify the date entered
        slDate = RptSelRk!CSI_CalWeek.Text
        llDateEntered = gDateValue(slDate)
        llEndDate = llDateEntered + (ilHowManyPer - 1) * 7 + 6              'get to sunday end date based on # periods
        ilPeriodType = 0                                    'weekly flag
    End If
   
    gPackDateLong llDateEntered, ilStartDate(0), ilStartDate(1)
    gPackDateLong llEndDate, ilEndDate(0), ilEndDate(1)
    
    ilRet = gObtainRcfRifRdf()          'get the rate cards and assoc dayparts.  Dayparts required for missed spots named avails

    ReDim tmStatsByAdvt(0 To 0) As STATSBYADVT
    ReDim imNamedAvails(0 To 0) As Integer
    'build array of the named avails to be used in the report
    For ilLoopOnListBox = 0 To RptSelRk!lbcSelection(2).ListCount - 1
        If RptSelRk!lbcSelection(2).Selected(ilLoopOnListBox) Then               'selected by user
            slNameCode = tgNamedAvail(ilLoopOnListBox).sKey
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            imNamedAvails(UBound(imNamedAvails)) = Val(slCode)
            ReDim Preserve imNamedAvails(0 To UBound(imNamedAvails) + 1) As Integer
        End If
    Next ilLoopOnListBox
    'Process and create prepass records for 1 selected vehicle at a time
    For ilVehicle = 0 To RptSelRk!lbcSelection(0).ListCount - 1 Step 1
        If (RptSelRk!lbcSelection(0).Selected(ilVehicle)) Then
            slNameCode = tgCSVNameCode(ilVehicle).sKey
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
            ReDim tmStatsByAdvt(0 To 0) As STATSBYADVT
            ReDim tmStatsByDP(0 To 0) As STATSBYDP
            
            mInitDPData ilVefCode                    'create array of the dp to process.  If overall totals for vehicle and not broken out by DP,

            'process all dayparts or named avails for a vehicle
            mGatherDPStatsByAdvt ilVefCode, llDateEntered, llEndDate, tlCntTypes
            
            'all advertiser have been gathered and accumulated by dp/vehicle.
            'Calculate the % inventory, % rev, avg rate, and sort for the top down
            mCalcDPOutput
            ilAtLeastOneDP = False                      'flag if at least one record to print, otherwise put out a phoney record for a header
            llRank = 1
            For ilIndex = LBound(tmStatsByAdvt) To UBound(tmStatsByAdvt) - 1 Step 1
            '      8/10/12 grfFields
            '      GenDate - Generation date
            '      GenTime - Generation Time
            '      VefCode - vehicle code
            '      rdfCode - DP
            '      adfCode - advertiser
            '      Code2 - dp or item sort code
            '      sofcode - 0 = weekly, 1 = corp, 2 = std, 3 = cal
            '      slfcode - 0 = gross, 1 = net, 2 = tnet
            '      DateGenl(1) - start date
            '      DateGenl(2) - end date
            '      BktType - No date for requested period "*"
            '      iPerGenl(1) - game #
            '      lDollars(1) - % units sold
            '      lDollars(2) - % rev sold
            '      lDollars(3) - # 30 units sold
            '      lDollars(4) - avg 30" rate
            '      lDollars(5) - gross/net/t-net rev
            '      lDollars(6) - rank
            '
                tmGrf.iGenDate(0) = igNowDate(0)
                tmGrf.iGenDate(1) = igNowDate(1)
                tmGrf.lGenTime = lgNowTime
                tmGrf.iVefCode = tmStatsByAdvt(ilIndex).iVefCode
                tmGrf.iRdfCode = tmStatsByAdvt(ilIndex).lKeyCode
                tmGrf.iAdfCode = tmStatsByAdvt(ilIndex).iAdfCode
                tmGrf.iCode2 = tmStatsByAdvt(ilIndex).iSortCode          'dp/item sort code
                'tmGrf.lDollars(1) = tmStatsByAdvt(ilIndex).iPctUnitsSold
                'tmGrf.lDollars(2) = tmStatsByAdvt(ilIndex).iPctRevSold
                'tmGrf.lDollars(3) = tmStatsByAdvt(ilIndex).lUnits
                'tmGrf.lDollars(4) = tmStatsByAdvt(ilIndex).lAvgRate
                'tmGrf.lDollars(5) = tmStatsByAdvt(ilIndex).lRev
                'tmGrf.lDollars(6) = llRank
                tmGrf.lDollars(0) = tmStatsByAdvt(ilIndex).iPctUnitsSold
                tmGrf.lDollars(1) = tmStatsByAdvt(ilIndex).iPctRevSold
                tmGrf.lDollars(2) = tmStatsByAdvt(ilIndex).lUnits
                tmGrf.lDollars(3) = tmStatsByAdvt(ilIndex).lAvgRate
                tmGrf.lDollars(4) = tmStatsByAdvt(ilIndex).lRev
                tmGrf.lDollars(5) = llRank
                tmGrf.iSofCode = ilPeriodType                   'for rpt header
                tmGrf.iSlfCode = imGrossNetTNet                 'for rpt header
                'tmGrf.iDateGenl(0, 1) = ilStartDate(0)                'user start date requested
                'tmGrf.iDateGenl(1, 1) = ilStartDate(1)
                'tmGrf.iDateGenl(0, 2) = ilEndDate(0)                'user end date requested
                'tmGrf.iDateGenl(1, 2) = ilEndDate(1)
                tmGrf.iDateGenl(0, 0) = ilStartDate(0)                'user start date requested
                tmGrf.iDateGenl(1, 0) = ilStartDate(1)
                tmGrf.iDateGenl(0, 1) = ilEndDate(0)                'user end date requested
                tmGrf.iDateGenl(1, 1) = ilEndDate(1)
                tmGrf.sBktType = ""                             'data exists flag for crystal
                'tmGrf.iPerGenl(1) = tmStatsByAdvt(ilIndex).iGameNo
                tmGrf.iPerGenl(0) = tmStatsByAdvt(ilIndex).iGameNo
                ilAtLeastOneDP = True
                ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                llRank = llRank + 1
            Next ilIndex

        End If                                  'vehicle selected
    Next ilVehicle                              'For ilvehicle = 0 To RptSelRk!lbcSelection(0).ListCount - 1
 
    If Not ilAtLeastOneDP Then
        tmGrf.iGenDate(0) = igNowDate(0)
        tmGrf.iGenDate(1) = igNowDate(1)
        tmGrf.lGenTime = lgNowTime
        tmGrf.iSofCode = ilPeriodType                   'for rpt header
        tmGrf.iSlfCode = imGrossNetTNet                 'for rpt header
        'tmGrf.iDateGenl(0, 1) = ilStartDate(0)                'user start date requested
        'tmGrf.iDateGenl(1, 1) = ilStartDate(1)
        'tmGrf.iDateGenl(0, 2) = ilEndDate(0)                'user end date requested
        'tmGrf.iDateGenl(1, 2) = ilEndDate(1)
        tmGrf.iDateGenl(0, 0) = ilStartDate(0)                'user start date requested
        tmGrf.iDateGenl(1, 0) = ilStartDate(1)
        tmGrf.iDateGenl(0, 1) = ilEndDate(0)                'user end date requested
        tmGrf.iDateGenl(1, 1) = ilEndDate(1)
        tmGrf.sBktType = "*"                                'no data exists message to show in crystal
        'tmGrf.iPerGenl(1) = 0
        tmGrf.iPerGenl(0) = 0
        ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
    End If
    
    Erase tmStatsByAdvt
    Erase tmStatsByDP
    Erase imNamedAvails
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hmSsf)
    ilRet = btrClose(hmSdf)
    ilRet = btrClose(hmVef)
    ilRet = btrClose(hmGrf)
    ilRet = btrClose(hmLcf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hmCHF)
    ilRet = btrClose(hmMnf)
    ilRet = btrClose(hmFsf)
    ilRet = btrClose(hmAnf)
    ilRet = btrClose(hmGsf)
    btrDestroy hmRdf
    btrDestroy hmSdf
    btrDestroy hmVef
    btrDestroy hmSsf
    btrDestroy hmLcf
    btrDestroy hmCff
    btrDestroy hmClf
    btrDestroy hmGrf
    btrDestroy hmCHF
    btrDestroy hmMnf
    btrDestroy hmFsf
    btrDestroy hmAnf
    btrDestroy hmGsf
    Exit Sub

   On Error GoTo 0
   Exit Sub

gCreateSpotPricingRanking_Error:
    gDbg_HandleError "RptCrRK: gCreateSpotPriceRanking"

    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure gCreateSpotPriceRanking of Module RptCrRk"
End Sub
'*****************************************************************
'*                                                               *
'*                                                               *
'*                                                               *
'*      Procedure Name:mGatherDPStatsByAdvt for Avails Combo report    *
'*                                                               *
'*      Created:9/27/00       By:D. Hosaka                       *
'*                                                               *
'*
'*      3-24-03 change way to test fill/extra spots.  Use SDF not SSF
'*****************************************************************
Sub mGatherDPStatsByAdvt(ilVefCode As Integer, llSDate As Long, llEDate As Long, tlCntTypes As CNTTYPES)
'
'   Where:
'   ilVefCode (I) - vehicle code to process
'   llSDate (I) - start date to begin searching Avails
'   llEDate (I) - end date to stop searching avails
'   tlCntTypes - types of contracts/spots to incl/excl

    Dim slType As String
    Dim ilType As Integer
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
    Dim ilRemLen As Integer     'time in seconds of avail, each spot length subtracted to get remaining seconds
    Dim ilRemUnits As Integer   '# of units of avail, each spot length subtracted to get remaining units
    ReDim ilEvtType(0 To 14) As Integer
    Dim slChfType As String
    Dim slChfStatus As String
    Dim ilVefIndex As Integer
    Dim llSpotAmount As Long
    Dim ilSpotsSoldInSec As Integer
    Dim ilSpotsSoldInUnits As Integer
    Dim ilUnitCount As Integer
    Dim ilAnfCode As Integer
    Dim ilSetDoNotShow As Integer
    Dim llPropPrice As Long     'proposal price of spot
    Dim ilOtherMissedLoop As Integer
    Dim ilFoundMissedDP As Integer
    ReDim tlorigllc(0 To 0) As LLC
    Dim ilLoopOnLLC As Integer
    Dim llOrigTime As Long
    Dim ilTemp As Integer
    Dim ilLoopOnDPStats As Integer
    Dim ilAvailEventInx As Integer
    Dim slAvailTime As String

    On Error GoTo mGatherDPStatsByAdvt_Error:
    slType = "O"
    ilType = 0
    llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
    'set the type of events to get fro the day (only Contract avails)
    For ilLoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(ilLoop) = False
    Next ilLoop
    ilEvtType(2) = True
    imSdfRecLen = Len(tmSdf)
    imCHFRecLen = Len(tmChf)
    ilVefIndex = gBinarySearchVef(ilVefCode)
    For llLoopDate = llSDate To llEDate Step 1
            slDate = Format$(llLoopDate, "m/d/yy")
            gPackDate slDate, ilDate0, ilDate1
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

            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilType) And (tmSsf.iVefCode = ilVefCode And (tmSsf.iDate(0) = ilDate0) And (tmSsf.iDate(1) = ilDate1))
                'Build the events that originally made up this day in case the day has been posted in the past
                'only build the avails to retreive the original inventory
                ReDim tlorigllc(0 To 0) As LLC
                ilRet = gBuildEventDay(ilType, "C", ilVefCode, slDate, "12M", "12M", ilEvtType(), tlorigllc())
                
                gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llDate
                ilBucketIndex = gWeekDayLong(llDate)        'day of week bucket index
                ilEvt = 1

                Do While ilEvt <= tmSsf.iCount
                   LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    If tmProg.iRecType = 1 Then    'Program (not working for nested prog)
                        ilLtfCode = tmProg.iLtfCode
                    ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 2) Then 'Contract Avails only
                        'ilAvailEventInx = ilEvt                         'save starting point of avail entry
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                        'Determine which rate card avail this is associated with
                        'Loop on Base dayparts associated with this vehicle.  If overlapping, data will go into each daypart that meets criteria
                        For ilRdf = LBound(tmStatsByDP) To UBound(tmStatsByDP) - 1 Step 1
                            ilAvailOk = False
                            'ilEvt = ilAvailEventInx                 'use the same starting point of avail.  Process avail and all spots within it for all DP before advancing to next avail
                            For ilLoop = LBound(tmStatsByDP(ilRdf).iStartTime, 2) To UBound(tmStatsByDP(ilRdf).iStartTime, 2) Step 1 'Row
                                If (tmStatsByDP(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tmStatsByDP(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                                    gUnpackTimeLong tmStatsByDP(ilRdf).iStartTime(0, ilLoop), tmStatsByDP(ilRdf).iStartTime(1, ilLoop), False, llStartTime
                                    gUnpackTimeLong tmStatsByDP(ilRdf).iEndTime(0, ilLoop), tmStatsByDP(ilRdf).iEndTime(1, ilLoop), True, llEndTime
                                    'If (llTime >= llStartTime) And (llTime < llEndTime) And (tmStatsByDP(ilRdf).sWkDays(ilLoop, ilBucketIndex + 1) = "Y") Then
                                    If (llTime >= llStartTime) And (llTime < llEndTime) And (tmStatsByDP(ilRdf).sWkDays(ilLoop, ilBucketIndex) = "Y") Then
                                        ilAvailOk = True
                                       'ilRdf is the index for the DP to accumulate inventory and gross $
                                        Exit For
                                    End If
                                End If
                            Next ilLoop
                            slAvailTime = gFormatTimeLong(llTime, "A", "1")
                            If ilAvailOk Then       'valid daypart or named avail, and avail time falls within the valid times
                            
                                If tmStatsByDP(ilRdf).sInOut = "I" Then   'Book into
                                    If tmAvail.ianfCode <> tmStatsByDP(ilRdf).ianfCode Then
                                        ilAvailOk = False
                                    End If
                                        
                                ElseIf tmStatsByDP(ilRdf).sInOut = "O" Then   'Exclude
                                    If tmAvail.ianfCode = tmStatsByDP(ilRdf).lKeyCode Then
                                        ilAvailOk = False
                                    End If
                                ElseIf tmStatsByDP(ilRdf).sInOut = "N" Then     'any - ok
                                    ilAvailOk = ilAvailOk
                                ElseIf tmStatsByDP(ilRdf).sInOut = "A" Then       'All avails; initialized when building the named avails table.  This is for the Vehicle summary option, where no vehicle DP is used
                                    ilAvailOk = ilAvailOk
                                End If
                                
                                mTestNamedAvail ilAvailOk, tmAvail.ianfCode
                                ilRemLen = tmAvail.iLen                 'total defined
                                ilRemUnits = tmAvail.iAvInfo And &H1F   'total units defined
                                If ilAvailOk And ((ilRemLen >= 30) Or (ilRemLen < 30 And imUnder30 = True)) Then        'valid avail if its over 30" or more, or if less than 30 it should be included

                                    'Count of Inventory in seconds
                                    tmStatsByDP(ilRdf).lInv = tmStatsByDP(ilRdf).lInv + (ilRemLen * 10)
                                    tmStatsByDP(ilRdf).lUnits = tmStatsByDP(ilRdf).lUnits + ilRemUnits
                                    
                                    For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                                        ilSpotOK = True
                                        'ilSpotOK = mFilterContractTypeFromSpot(tmSpot, tlCntTypes)      'filter out contr header types (remnant, DR, PI, etc)

                                        If (tmSpot.iPosLen >= 30 Or imUnder30) Then                             'include avails and/or spots under 30"?
                                            tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                            ilRet = btrGetEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)

                                            If ilRet <> BTRV_ERR_NONE Then
                                                ilSpotOK = False                    'invalid sdf code
                                            End If

                                            'Test for Feed Spot
                                            If ilRet = BTRV_ERR_NONE And tmSdf.lChfCode = 0 And ilSpotOK Then        'feed spot
                                                'obtain the network information
                                                tmFSFSrchKey.lCode = tmSdf.lFsfCode
                                                ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                                If ilRet <> BTRV_ERR_NONE Or Not tlCntTypes.iNetwork Then
                                                    ilSpotOK = False                    'invalid network code
                                                End If

                                                slChfType = ""          'contract types dont apply with feed spots
                                                slChfStatus = ""       'status types dont apply with feed spots
                                            Else            'Test for contract spots
                                                If ilRet = BTRV_ERR_NONE And tmSdf.lChfCode > 0 And ilSpotOK Then
                                                    'obtain contract info
                                                    If tmSpot.lSdfCode = tmSdf.lCode And ilRet = BTRV_ERR_NONE And ilSpotOK = True Then
                                                        If tmSdf.lChfCode <> tmChf.lCode Then                      'if already in mem, don't reread
                                                            tmChfSrchKey.lCode = tmSdf.lChfCode
                                                            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                                            'if error reading the spot recd or contrct line chfcode doesnt match the spot recd, dont process the spot
                                                            If ilRet <> BTRV_ERR_NONE Or tmChf.lCode <> tmSdf.lChfCode Then
                                                                ilSpotOK = False
                                                            End If
                                                        End If
                                                        slChfType = tmChf.sType
                                                        slChfStatus = tmChf.sStatus
                                                        mFilterSpot ilVefCode, tlCntTypes, ilSpotOK
                                                    End If
                                                End If
                                            End If


                                            If ilSpotOK Then
                                                llSpotAmount = mGetSpotPrice()   'get spot price
                                                mAccumStatsByAdvt ilRdf, llSpotAmount

                                            End If                              'ilspotOK
                                        End If
                                    Next ilSpot                                 'loop from ssf file for # spots in avail
                                End If                                          'Avail OK
                            End If                                          'avail ok
                        Next ilRdf                                          'ilRdf = lBound(tlComboType).  Process all DP for the avail before advancing to next avail
                        ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                    End If
                    ilEvt = ilEvt + 1   'Increment to next event
                Loop                                                        'do while ilEvt <= tmSsf.iCount

                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                If tgMVef(ilVefIndex).sType = "G" Then
                    ilType = tmSsf.iType
                End If

            Loop
      
    Next llLoopDate

    'Get missed; any spot not within a dp is excluded
    slDate = Format$(llSDate, "m/d/yy")
    gPackDate slDate, ilDate0, ilDate1

    If (tlCntTypes.iMissed) Then            'include missed
        'Key 2: VefCode; SchStatus; AdfCode; Date, Time
        For ilPass = 0 To 2 Step 1
            tmSdfSrchKey2.iVefCode = ilVefCode
            If ilPass = 0 Then
                slType = "M"
            ElseIf ilPass = 1 Then              'currently Ready to schedule & Unschedule makegoods not used
                slType = "R"
            ElseIf ilPass = 2 Then
                slType = "U"
            End If
            tmSdfSrchKey2.sSchStatus = slType
            tmSdfSrchKey2.iAdfCode = 0
            tmSdfSrchKey2.iDate(0) = ilDate0
            tmSdfSrchKey2.iDate(1) = ilDate1
            tmSdfSrchKey2.iTime(0) = 0
            tmSdfSrchKey2.iTime(1) = 0
            ilRet = btrGetGreaterOrEqual(hmSdf, tmSdf, imSdfRecLen, tmSdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get first record as starting point
            'This code added as replacement for Ext operation
            Do While (ilRet = BTRV_ERR_NONE) And (tmSdf.iVefCode = ilVefCode) And (tmSdf.sSchStatus = slType)
                gUnpackDateLong tmSdf.iDate(0), tmSdf.iDate(1), llDate

                'missed spot must be within selected date parameters, spot must be selected for inclusion (air time vs feed spot), and the day selectivity must be OK (currently no selectivity, all days always included)
                If (llDate >= llSDate And llDate <= llEDate) And ((tlCntTypes.iCntrSpots = True And tmSdf.lChfCode > 0) Or (tlCntTypes.iNetwork = True And tmSdf.lChfCode = 0)) Then        'Has this day of the week been selected?
                    ilBucketIndex = gWeekDayLong(llDate)
                    gUnpackTimeLong tmSdf.iTime(0), tmSdf.iTime(1), False, llTime
                    'obtain the daypart and its named avail
                    ilSpotOK = True
                    If tmSdf.lChfCode = 0 Then
                        'obtain the network information
                        tmFSFSrchKey.lCode = tmSdf.lFsfCode
                        ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFSFSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                        If ilRet <> BTRV_ERR_NONE Or Not tlCntTypes.iNetwork Then
                            ilSpotOK = False                    'invalid network code
                        End If
                        slChfType = ""          'contract types dont apply with feed spots
                        slChfStatus = ""       'status types dont apply with feed spots

                        'Show on separate line with Missed-Any Avails; no DP exists to determine what avail it should fall into
                        ilSpotOK = False
                    Else
                       ilRet = BTRV_ERR_NONE
                       tmClfSrchKey.lChfCode = tmSdf.lChfCode
                       tmClfSrchKey.iLine = tmSdf.iLineNo
                       tmClfSrchKey.iCntRevNo = 32000 ' 0 show latest version
                       tmClfSrchKey.iPropVer = 32000 ' 0 show latest version
                       ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                       If (tmClf.lChfCode <> tmSdf.lChfCode) Then     'got the matching line reference?
                           ilSpotOK = False
                       End If
                        ilAnfCode = gBinarySearchRdf(tmClf.iRdfCode)        'search for matching dp to get the named avail code
                        If ilAnfCode < 0 Then
                            ilSpotOK = False
                        End If

                    End If

                    If ilSpotOK Then
                        ilAvailOk = False
                        ilFoundMissedDP = False

                            'find the associated DP or named avail
                            For ilRdf = LBound(tmStatsByDP) To UBound(tmStatsByDP) - 1 Step 1

                                For ilLoop = LBound(tmStatsByDP(ilRdf).iStartTime, 2) To UBound(tmStatsByDP(ilRdf).iStartTime, 2) Step 1 'Row
                                    If (tmStatsByDP(ilRdf).iStartTime(0, ilLoop) <> 1) Or (tmStatsByDP(ilRdf).iStartTime(1, ilLoop) <> 0) Then
                                        gUnpackTimeLong tmStatsByDP(ilRdf).iStartTime(0, ilLoop), tmStatsByDP(ilRdf).iStartTime(1, ilLoop), False, llStartTime
                                        gUnpackTimeLong tmStatsByDP(ilRdf).iEndTime(0, ilLoop), tmStatsByDP(ilRdf).iEndTime(1, ilLoop), True, llEndTime
                                        'If (llTime >= llStartTime) And (llTime < llEndTime) And (tmStatsByDP(ilRdf).sWkDays(ilLoop, ilBucketIndex + 1) = "Y") Then
                                        If (llTime >= llStartTime) And (llTime < llEndTime) And (tmStatsByDP(ilRdf).sWkDays(ilLoop, ilBucketIndex) = "Y") Then
                                            ilAvailOk = True

                                            If (tmClf.iRdfCode <> tmStatsByDP(ilRdf).lKeyCode) Then     'match the sold dp against the dp being processed
                                           ' mTestNamedAvail ilAvailOk, tgMRdf(ilAnfCode).ianfCode       'match on named avails of the DP selected
                                                ilAvailOk = False
                                            Else
                                             If ilAvailOk Then
                                                ilFoundMissedDP = True      'missed spot daypart found
                                                Exit For
                                            End If
                                            End If
                                         
                                        End If
                                    End If
                                Next ilLoop
                                If (ilAvailOk) And ((tmSdf.iLen >= 30) Or (tmSdf.iLen < 30 And imUnder30 = True)) Then

                                    ilSpotOK = True                'assume spot is OK
                                    If tmSdf.lChfCode <> tmChf.lCode Then               'if already in mem, don't reread
                                       tmChfSrchKey.lCode = tmSdf.lChfCode
                                       ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                       If ilRet <> BTRV_ERR_NONE Then
                                           ilSpotOK = False
                                       End If
                                   End If

                                   mFilterSpot ilVefCode, tlCntTypes, ilSpotOK

                                    If ilSpotOK Then
                                        llSpotAmount = mGetSpotPrice()   'get spot price
                                        mAccumStatsByAdvt ilRdf, llSpotAmount
                                        
                                    End If                      'ilSpotOK
                                End If                          'ilavailok
                            Next ilRdf
                    End If                              'ilanfcode < 0 : missing dp from line
                End If              'dates within filter
                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next ilPass
    End If

    Erase ilEvtType
    Erase tlLLC
    Exit Sub
    
mGatherDPStatsByAdvt_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mGatherDPStatsByAdvt of Module RptCrRk"
End Sub
'
'
'
'               mFilterSpot - Test header and line exclusions for user request.
'
'               <input> ilVefCode - airing vehicle
'                       tlCntTypes - structure of inclusions/exclusions of contract types & status
'              <output> ilSpotOk - true if spot is OK, else false to ignore spot
'
Sub mFilterSpot(ilVefCode As Integer, tlCntTypes As CNTTYPES, ilSpotOK As Integer)
Dim ilRet As Integer
Dim slPrice As String
Dim ilIsItPolitical As Integer

    If ilSpotOK Then
        'Test header exclusions (types of contrcts and statuses)
        
        If lmContract > 0 And lmContract <> tmChf.lCntrNo Then
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
        End If

        If tmChf.sType = "C" Then           '3-16-10 wrong flag tested for standard, should be C, not S
            If Not tlCntTypes.iStandard Then       'include Standard types?
                ilSpotOK = False
            End If

        ElseIf tmChf.sType = "V" Then
            If Not tlCntTypes.iReserv Then      'include reservations ?
                ilSpotOK = False
            End If

        ElseIf tmChf.sType = "R" Then
            If Not tlCntTypes.iDR Then       'include DR?
                ilSpotOK = False
            End If
            
        ElseIf tmChf.sType = "T" Then
            If Not tlCntTypes.iRemnant Then       'include Remnant?
                ilSpotOK = False
        End If
        
        ElseIf tmChf.sType = "Q" Then
            If Not tlCntTypes.iPI Then       'include PI?
                ilSpotOK = False
            End If

        ElseIf tmChf.sType = "M" Then
            If Not tlCntTypes.iPromo Then       'include Promo?
                ilSpotOK = False
            End If
            
        ElseIf tmChf.sType = "S" Then
            If Not tlCntTypes.iPSA Then       'include PSA?
                ilSpotOK = False
            End If

        End If
        
        If (tmChf.iPctTrade = 100 And Not tlCntTypes.iTrade) Then
            ilSpotOK = False
        End If

        'Political
        ilIsItPolitical = gIsItPolitical(tmChf.iAdfCode)           'its a political, include this contract?
        If ilIsItPolitical Then
            If Not (tlCntTypes.iPolit) Then                   'its a political
                ilSpotOK = False
            End If
        Else
            If Not tlCntTypes.iNonPolit Then              'include non politicals?
                ilSpotOK = False           'no, exclude them
            End If
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
        'look for inclusion of spot types
        If (InStr(slPrice, ".") <> 0) Then    'found spot cost
            'is it a .00?
            If gCompNumberStr(slPrice, "0.00") = 0 Then     'its a .00 spot
                If Not tlCntTypes.iZero Then
                    ilSpotOK = False
                End If
            Else
                If Not tlCntTypes.iCharge Then           'exclude charged spots
                    ilSpotOK = False
                End If
            End If
        ElseIf Trim$(slPrice) = "ADU" Then
            If Not tlCntTypes.iADU Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "Bonus" Then
            If Not tlCntTypes.iBonus Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "+ Fill" Then       '3-24-03
            If Not tlCntTypes.iXtra Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "- Fill" Then        '3-24-03
            If Not tlCntTypes.iFill Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "N/C" Then
            If Not tlCntTypes.iNC Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "Recapturable" Then
            If Not tlCntTypes.iRecapturable Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "Spinoff" Then
            If Not tlCntTypes.iSpinoff Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "MG" Then               '10-28-10
            If Not tlCntTypes.iMG Then
                ilSpotOK = False
            End If
        End If
        
        '10-28-10  if excluding MG, that includes MG rate spot types & MG/outside spots
        'test for mg/outside (cant be a bonus spot) and if MG should be included
        If ((tmSdf.sSchStatus = "G" Or tmSdf.sSchStatus = "O") And tmSdf.sSpotType <> "X") And (Not tlCntTypes.iMG) Then
            ilSpotOK = False
        End If
    End If                                  'ilspotOK
End Sub
'
'           mInitDPData - create an array of dayparts to process
'
'           <input> ilvefCode : vehicle code
'
Public Sub mInitDPData(ilVefCode As Integer)
Dim ilRcf As Integer
Dim ilSelected As Integer
Dim ilLoop As Integer
Dim slNameCode As String
Dim ilRet As Integer
Dim slCode As String
Dim llRif As Long
Dim ilRdf As Integer
Dim ilSaveSort As Integer
Dim ilFound As Integer
Dim ilUpper As Integer
Dim ilAvailNameLoop As Integer
Dim ilDayIndex As Integer
Dim slName As String
Dim ilAnfCode As Integer
Dim ilMissed As Integer
Dim ilLooponNamedAvails As Integer
Dim ilInclude As Integer
Dim ilIndex As Integer

        On Error GoTo mInitDPData_Error
        ilUpper = 0
        ReDim tmStatsByDP(0 To 0) As STATSBYDP                  'intialized for each vehicle because each vehicle has different dayparts
            
        If RptSelRk!rbcTotalsBy(0).Value = True Then
            For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1     'non-sports avails, build key info for the selected r/c
                tmRcf = tgMRcf(ilRcf)
                ilSelected = False
                For ilLoop = 0 To RptSelRk!lbcSelection(1).ListCount - 1 Step 1
                    slNameCode = tgRateCardCode(ilLoop).sKey
                    ilRet = gParseItem(slNameCode, 3, "\", slCode)
                    If Val(slCode) = tgMRcf(ilRcf).iCode Then
                        If (RptSelRk!lbcSelection(1).Selected(ilLoop)) Then
                            ilSelected = True
                        End If
                        Exit For
                    End If
                Next ilLoop

                If ilSelected Then              'found the R/c
                    'Setup the DP with rates to process
                    For llRif = LBound(tgMRif) To UBound(tgMRif) - 1 Step 1
                        If tgMRif(llRif).iRcfCode = tgMRcf(ilRcf).iCode Then
                            ilRdf = gBinarySearchRdf(tgMRif(llRif).iRdfCode)
                            If ilRdf <> -1 Then
                                If tgMRif(llRif).iSort = 0 Then
                                    ilSaveSort = tgMRdf(ilRdf).iSortCode            'no item sort code, use dp sort code
                                Else
                                    ilSaveSort = tgMRif(llRif).iSort            'use sort from the item vs daypart
                                End If
                                If ilSaveSort = 0 Then
                                    ilSaveSort = 32000          'no DP sort Code in Items record, make it last
                                End If
                                'match items (of dp/rc) against valid rate card, must be base DP that is Active and match vehicle selected
                                If tgMRdf(ilRdf).iCode = tgMRif(llRif).iRdfCode And tgMRif(llRif).sBase = "Y" And tgMRdf(ilRdf).sState <> "D" And tgMRif(llRif).iVefCode = ilVefCode Then
                                    ilInclude = False
                                    'Build an entry for a valid DP to include in report.  this will contains the overall gross revenue and overall 30" units to calculate % of the advt sold
                                    'see if the named avails should be included
                                    For ilLooponNamedAvails = LBound(imNamedAvails) To UBound(imNamedAvails) - 1
                                        If tgMRdf(ilRdf).sInOut = "I" Then
                                            If imNamedAvails(ilLooponNamedAvails) = tgMRdf(ilRdf).ianfCode Then     'include this named avail?
                                                'ok to include it
                                                ilInclude = True
                                                Exit For
                                            End If
                                        ElseIf tgMRdf(ilRdf).sInOut = "O" Then
                                            If imNamedAvails(ilLooponNamedAvails) = tgMRdf(ilRdf).ianfCode Then
                                                ilInclude = False
                                                Exit For
                                            End If
                                        ElseIf tgMRdf(ilRdf).sInOut = "N" Then          'any avail ok
                                            ilInclude = True
                                            Exit For
                                        End If
                                    Next ilLooponNamedAvails
                                         
                                    If ilInclude Then
                                        tmStatsByDP(ilUpper).iType = 0         'dp info vs overall vehicle info, which doesnt have a internal dp code
                                        tmStatsByDP(ilUpper).lKeyCode = tgMRdf(ilRdf).iCode     'internal dp code
                                        tmStatsByDP(ilUpper).iSortCode = ilSaveSort
                                        ilIndex = LBound(tmStatsByDP(ilUpper).iStartTime, 2)
                                        For ilLoop = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                            tmStatsByDP(ilUpper).iStartTime(0, ilIndex) = tgMRdf(ilRdf).iStartTime(0, ilLoop)
                                            tmStatsByDP(ilUpper).iStartTime(1, ilIndex) = tgMRdf(ilRdf).iStartTime(1, ilLoop)
                                            tmStatsByDP(ilUpper).iEndTime(0, ilIndex) = tgMRdf(ilRdf).iEndTime(0, ilLoop)
                                            tmStatsByDP(ilUpper).iEndTime(1, ilIndex) = tgMRdf(ilRdf).iEndTime(1, ilLoop)
                                            For ilDayIndex = 1 To 7
                                                tmStatsByDP(ilUpper).sWkDays(ilIndex, ilDayIndex - 1) = tgMRdf(ilRdf).sWkDays(ilLoop, ilDayIndex - 1)
                                            Next ilDayIndex
                                            ilIndex = ilIndex + 1
                                        Next ilLoop
                                        tmStatsByDP(ilUpper).ianfCode = tgMRdf(ilRdf).ianfCode
                                        tmStatsByDP(ilUpper).sInOut = tgMRdf(ilRdf).sInOut
                                        tmStatsByDP(ilUpper).iVefCode = ilVefCode
                                       
                                        ilUpper = ilUpper + 1
                                        ReDim Preserve tmStatsByDP(0 To ilUpper) As STATSBYDP
                                    End If  'not ilfound
                                End If      'tgMRdf(ilRdf).iCode = tgMRif(llRif).iRdfcode And tgMRdf(ilRdf).sState <> "D" And tgMRif(llRif).iVefCode = ilVefCode
                            End If          'ilRdf <> -1
                        End If              'tgMRif(llRif).iRcfCode = tgMRcf(ilRcf).iCode
                    Next llRif
                End If                      'ilSelected
            Next ilRcf
        Else                                'sort by vehicle (no dayparts, gather entire vehicle
                                            'create a phoney entry for 12m-12m M-f to gather the entire vehicle
            tmStatsByDP(ilUpper).iType = 1         'vehicle info, which doesnt have a internal dp code
            tmStatsByDP(ilUpper).lKeyCode = 0    'internal dp code
            tmStatsByDP(ilUpper).iSortCode = 0
            'if only 1 time period defined in the dp, its in the last entry of the array of 7
            gPackTime "12M", tmStatsByDP(ilUpper).iStartTime(0, 6), tmStatsByDP(ilUpper).iStartTime(1, 6)
            gPackTime "11:59:59PM", tmStatsByDP(ilUpper).iEndTime(0, 6), tmStatsByDP(ilUpper).iEndTime(1, 6)
            For ilDayIndex = 1 To 7
                tmStatsByDP(ilUpper).sWkDays(6, ilDayIndex - 1) = "Y"
            Next ilDayIndex
             
            tmStatsByDP(ilUpper).ianfCode = 0
            tmStatsByDP(ilUpper).sInOut = "A"                   'flag for All avails
            tmStatsByDP(ilUpper).iVefCode = ilVefCode
            
            ilUpper = ilUpper + 1
            ReDim Preserve tmStatsByDP(0 To ilUpper) As STATSBYDP
        End If
        Exit Sub
mInitDPData_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mInitDPData of Module RptCrRk"
End Sub
'
'       mGetSpotPrice - get spot rate for Combo Avails report
'       <return> Spot rate
'
Private Function mGetSpotPrice() As Long


Dim ilRet As Integer
Dim llActualPrice As Long
        On Error GoTo mGetSpotPrice_Error
        llActualPrice = 0
        If tmSdf.sSpotType <> "X" And tmSdf.sSpotType <> "O" And tmSdf.sSpotType <> "C" Then                  'bonus?
            'need to get flight for proposal price and actual price
            ilRet = gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tmCff)

            llActualPrice = tmCff.lActPrice
            If (tmSdf.sPriceType = "N") Then
                llActualPrice = 0
            ElseIf (tmSdf.sPriceType = "P") Then
                llActualPrice = 0
            End If
        End If
        llActualPrice = gGetGrossNetTNetFromPrice(imGrossNetTNet, llActualPrice, tmClf.lAcquisitionCost, tmChf.iAgfCode)
        mGetSpotPrice = llActualPrice
        Exit Function
mGetSpotPrice_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mGetSpotPrice of Module RptCrRk"
End Function
'
'           mCloseAll - close all files (some may not have been opened)
'
Public Sub mCloseAll()
Dim ilRet As Integer

        ilRet = btrClose(hmGsf)
        ilRet = btrClose(hmAnf)
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmRdf)
        ilRet = btrClose(hmVsf)
        ilRet = btrClose(hmSmf)
        ilRet = btrClose(hmMnf)
        ilRet = btrClose(hmCff)
        ilRet = btrClose(hmClf)
        ilRet = btrClose(hmGrf)
        ilRet = btrClose(hmCHF)
        ilRet = btrClose(hmGhf)
        ilRet = btrClose(hmIsf)
        ilRet = btrClose(hmMsf)
        ilRet = btrClose(hmMgf)
        btrDestroy hmGsf
        btrDestroy hmAnf
        btrDestroy hmFsf
        btrDestroy hmRdf
        btrDestroy hmVsf
        btrDestroy hmMnf
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmClf
        btrDestroy hmGrf
        btrDestroy hmCHF
        btrDestroy hmGhf
        btrDestroy hmIsf
        btrDestroy hmMsf
        btrDestroy hmMgf

        Exit Sub
End Sub
'
'                   mAccumStatsByADvt - Update the advertiser by dp/vehicle for total revenue sched and spots sold
'                                       also update the DP/Vehicle for overall scheduled into this DP to calc % rev sold for advt
'                   <input> ilRdf - Index into array of summary DP entries to keep track of overall DP stats (inventory, sold, rev)
'                           llSpotPrice - spot price (gross, net or tnet value)
'                           ilUnitCount - # units sold for avg rate
'
Public Function mAccumStatsByAdvt(ilRdf As Integer, llSpotPrice As Long) As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilUnitCount As Integer
    Dim llTemp As Long
    Dim ilGame As Integer
    
        On Error GoTo mAccumStatsByAdvt_Error
    
        mAccumStatsByAdvt = -1              'if not found or no index into array for advt stats
        ilFound = False
         'accumulate 30" units.  Spots divisible by 30" = one unit each; 15" double a 30" rate, 60 is 1/2 of 30" rate
        ilUnitCount = (tmSdf.iLen * 10) / 30
        If imTotsByDPVef = 0 Then           'dp, game # could be applicable
            ilGame = tmSdf.iGameNo
        Else
            ilGame = 0
        End If
         
        For ilLoop = LBound(tmStatsByAdvt) To UBound(tmStatsByAdvt) - 1
            If tmStatsByAdvt(ilLoop).iAdfCode = tmSdf.iAdfCode And tmStatsByAdvt(ilLoop).lKeyCode = tmStatsByDP(ilRdf).lKeyCode And ilGame = tmStatsByAdvt(ilLoop).iGameNo Then      'match on advt and DP
                tmStatsByAdvt(ilLoop).lRev = tmStatsByAdvt(ilLoop).lRev + llSpotPrice                     'accumulate spot price by advt
                tmStatsByAdvt(ilLoop).lUnits = tmStatsByAdvt(ilLoop).lUnits + ilUnitCount                 'accumulate units sold for avg rate
                tmStatsByAdvt(UBound(tmStatsByAdvt)).lLenSched = tmSdf.iLen
                ilFound = True
                Exit For
            End If
        Next ilLoop
                
        If Not ilFound Then                 'no entry for this dp and advt yet
            tmStatsByAdvt(UBound(tmStatsByAdvt)).iAdfCode = tmSdf.iAdfCode
            tmStatsByAdvt(UBound(tmStatsByAdvt)).lKeyCode = tmStatsByDP(ilRdf).lKeyCode                   'dp code
            tmStatsByAdvt(UBound(tmStatsByAdvt)).iDPInxToStatsByDP = ilRdf              'index into the DP Stats array
            tmStatsByAdvt(UBound(tmStatsByAdvt)).iType = imTotsByDPVef              'totals by dp or vehicle
            tmStatsByAdvt(UBound(tmStatsByAdvt)).iVefCode = tmSdf.iVefCode
            tmStatsByAdvt(UBound(tmStatsByAdvt)).lLenSched = tmSdf.iLen
            tmStatsByAdvt(UBound(tmStatsByAdvt)).lUnits = ilUnitCount
            tmStatsByAdvt(UBound(tmStatsByAdvt)).iGameNo = ilGame
            tmStatsByAdvt(ilLoop).lRev = llSpotPrice                      'accumulate spot price by advt
            
            
            ReDim Preserve tmStatsByAdvt(0 To UBound(tmStatsByAdvt) + 1) As STATSBYADVT
        End If
        
        'accumulate overall revenue by DP
        tmStatsByDP(ilRdf).lRev = tmStatsByDP(ilRdf).lRev + llSpotPrice
        

        Exit Function
mAccumStatsByAdvt_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mAccumStatsByAdvt of Module RptCrRk"
        
End Function
'
'           Cycle thru the arrays created by advt/dp/vehicle to compute the % inv sold, % rev sold, avg 30" rate
'
Public Sub mCalcDPOutput()

    Dim ilRdf As Integer
    Dim llUnits As Long
    Dim llTemp As Long
    Dim ilIndex As Integer
    Dim slStr As String
    Dim slTemp As String
    
        On Error GoTo mCalcDPOutput_Error
        For ilIndex = LBound(tmStatsByAdvt) To UBound(tmStatsByAdvt) - 1
            '% inventory for this advt; maintain to xxx.x
            ilRdf = tmStatsByAdvt(ilIndex).iDPInxToStatsByDP
            llUnits = tmStatsByDP(ilRdf).lInv / 30                  'total seconds for entire dp/vehicle
            If llUnits <> 0 Then
                llTemp = ((tmStatsByAdvt(ilIndex).lUnits) / (llUnits)) * 1000  '% inv sold = # units sold for this dp/advt divided by total inv 30" units
                tmStatsByAdvt(ilIndex).iPctUnitsSold = llTemp               'xxx.x
            End If
            '% revenue for this advt, carry out to xxx.x
            If (tmStatsByDP(ilRdf).lRev) <> 0 Then
                llTemp = CSng(tmStatsByAdvt(ilIndex).lRev / (tmStatsByDP(ilRdf).lRev)) * 1000   '% rev sold = rev sold for this dp/advt divided by total rev for dp
                tmStatsByAdvt(ilIndex).iPctRevSold = llTemp     'xxx.x
            End If
            'average 30" rate
            If tmStatsByAdvt(ilIndex).lUnits <> 0 Then
                llTemp = CSng(tmStatsByAdvt(ilIndex).lRev) * 10 / tmStatsByAdvt(ilIndex).lUnits
                tmStatsByAdvt(ilIndex).lAvgRate = llTemp
            End If
            'Determine the top down sort
            If imWhichSort = 0 Then         '% inventory column
                llTemp = tmStatsByAdvt(ilIndex).iPctUnitsSold
            ElseIf imWhichSort = 1 Then     '% rev column
                llTemp = tmStatsByAdvt(ilIndex).iPctRevSold
            ElseIf imWhichSort = 2 Then     '# Units sch
                llTemp = tmStatsByAdvt(ilIndex).lUnits
            ElseIf imWhichSort = 3 Then     'Avg 30 rate
                llTemp = tmStatsByAdvt(ilIndex).lAvgRate
            Else                            'gros/net/tnet revenue
                llTemp = tmStatsByAdvt(ilIndex).lRev
            End If
                        
            'sort key is Vehicle/game/DP/ #/Value from sort column selected; keep all game numbers together for all dp
            slTemp = str$(9999999999# - llTemp)   'reverse the sort to descending
         
            Do While Len(slTemp) < 10
                slTemp = "0" & slTemp
            Loop
            tmStatsByAdvt(ilIndex).sKey = Trim$(slTemp)
            
            slTemp = Trim$(str(tmStatsByAdvt(ilIndex).lKeyCode))          'dp code to keep same dp together
            Do While Len(slTemp) < 5
                slTemp = "0" & slTemp
            Loop
            tmStatsByAdvt(ilIndex).sKey = Trim$(slTemp) & "|" & tmStatsByAdvt(ilIndex).sKey
            
            slTemp = Trim$(str(tmStatsByAdvt(ilIndex).iGameNo))
            Do While Len(slTemp) < 3
                slTemp = "0" & slTemp
            Loop
            tmStatsByAdvt(ilIndex).sKey = Trim$(slTemp) & "|" & tmStatsByAdvt(ilIndex).sKey
            
            slTemp = Trim$(str(tmStatsByAdvt(ilIndex).iVefCode))         'vehicle code
            Do While Len(slTemp) < 5
                slTemp = "0" & slTemp
            Loop
            
            tmStatsByAdvt(ilIndex).sKey = Trim$(slTemp) & "|" & tmStatsByAdvt(ilIndex).sKey
            
        Next ilIndex
        
        If UBound(tmStatsByAdvt) - 1 > 1 Then
            ArraySortTyp fnAV(tmStatsByAdvt(), 0), UBound(tmStatsByAdvt), 0, LenB(tmStatsByAdvt(0)), 0, LenB(tmStatsByAdvt(0).sKey), 0
        End If
        Exit Sub
        
mCalcDPOutput_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure mCalcDPOutput of Module RptCrRk"
End Sub
'
'                   mTestNamedAvail - test the named avail against the user selected ones
'                   <input> ilAnfCode - named avail code from avail entry
'                   <output> ilAvailOK - return false if the named avail is not one selected; otherwise do not change
Private Sub mTestNamedAvail(ilAvailOk As Integer, ilAnfCode As Integer)
    Dim ilLoopOnName As Integer
    Dim ilFound As Integer
        On Error GoTo mTestNamedAvailErr
               
        ilFound = False
        For ilLoopOnName = LBound(imNamedAvails) To UBound(imNamedAvails) - 1
            If ilAnfCode = imNamedAvails(ilLoopOnName) Then
                ilFound = True
                Exit For
            End If
        Next ilLoopOnName
        If Not ilFound Then
            ilAvailOk = False
        End If
        Exit Sub
mTestNamedAvailErr:
    On Error GoTo 0
    ilFound = ilFound
    Exit Sub
End Sub
