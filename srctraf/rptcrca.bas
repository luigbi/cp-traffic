Attribute VB_Name = "RptCrCA"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of rptcrca.bas on Wed 6/17/09 @ 12:56 PM
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
Dim tmGrf() As GRF
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

Dim imMajorSet As Integer       'for vehicle group selected
Dim imInclVGCodes As Integer
Dim imUseVGCodes() As Integer     'vehicle group items for selection (ie. items within participant, or market, or format, etc)
Dim imHowManyHL As Integer     '# of highlighted lengths entered
Dim imGameIndFlag As Integer       'for Multimedia info only;  flag to indicate to process the multimedia
                                'only once per vehicle otherwise inventory/sold will be overstated
Dim imUsingEqualize As Integer
'Dim fmLenRatio(1 To 5) As Single      '30/60 ratio for spot lengths to highlight  i.e. if equalizing 30s : 10" = 3, 15" = 2, 5" = 6, 60" = .5, 30" = 1
Dim fmLenRatio(0 To 5) As Single      '30/60 ratio for spot lengths to highlight  i.e. if equalizing 30s : 10" = 3, 15" = 2, 5" = 6, 60" = .5, 30" = 1
                                        '                                                if equalizing 60s:  10" = 6, 15" = 4, 5" = 12, 60" = 1, 30" = 2
                                        'Index zero ignored
Dim imEqualizeOption As Integer             '30 (equalize to 30s), 60 (equalize to 60s) or 0 (none)

Dim tmComboType() As COMBOTYPE
Dim tmComboData() As COMBODATA
Type COMBODATA                  'structure for avails data
    iType As Integer            '0= named avails, 1 = dp info
    lKeyCode As Long            'named avail code or DP code
    lDate As Long               'date of stats
    lInv As Long                'Inventory for day
    lSold(0 To 9) As Long        'spots sold per highlighted spot lengths, may not use all
    lAvails(0 To 9) As Long      'avails per highlighted spot lengths
    lTotalRev As Long           'total $ sold
    lTotalSpots As Long         'total spots sold
    iVefCode As Integer         'vehicle code
    iSortCode As Integer        'sort code from RIF if by dp, else 0 for games
End Type

Type COMBOTYPE
    iType As Integer            '0= named avails, 1 = dp info
    lKeyCode As Long            'named avail code or DP code
    iSortCode As Integer        'sort code from RIF if by dp, else 0 for games
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

Type MMINFO                 'multimedia info, inventory and sold
    iIhfCode As Integer        'inventory header code
    iGameNo As Integer      '0 = game independent
    lRate As Long           'total $
    lNoUnits As Long        '# units inventory
    lRateSold As Long       'total $ sold
    lNoUnitsSold As Long    'total units sold
End Type

'TTP 10434 - Event and Sports export (WWO)
Dim tmEventSportAvailsExportData() As EVENTSPORTSAVAILSEXPORTDATA
Type EVENTSPORTSAVAILSEXPORTDATA
    iVehicle As Integer                     'the name of the selected sports vehicle
    sVehicle As String
    iBroadcastWeek(0 To 1) As Integer       'the Monday date of the event air date. Format: mm/dd/yyyy
    iEventAirDate(0 To 1) As Integer        'the scheduled air date of the event
    iEventTime(0 To 1) As Integer           'the start time of the event from the programming schedule 'Time (Byte 0:Hund sec; Byte 1: sec.; Byte 2: min.; Byte 3:hour)
    iEventNo As Integer                     'the number of the event as defined in programming
    iHomeTeam As Integer                    'the name of the home team for the event
    sHomeTeam As String
    iVisitingTeam As Integer                'the name of the visiting team for the event
    sVisitingTeam As String
    sEventStatus As String * 1              'the four options are T = Tentative, F = Firm, C = Cancelled, P = Postponed
    
    'lStandardInventoryNormalized As Long    'Inventory Normalized to :30s: the total inventory for the standard avails for this event divided by 30. Example: if the inventory total is 120 seconds, the normalized inventory count would be 4.
    lStandardInventorySeconds As Long
    
    lStandardSoldNormalized As Long         'for this event, for each spot in the standard avail name inventory, for each spot length, multiply it by the number of units for that spot length, then add them together, then divide by 30 to get the normalized sold count. Example: there are two 30s and one 60s, which is 120 seconds. Divide 120 by 30 to get 4, which would be the normalized 30s sold count.
    lStandardSoldSeconds As Long
    
    lStandardRemainingNormalized As Long    'for this event, for the standard avail name inventory, add up the remaining available inventory and divide it by 30. For example if there is 60 seconds of remaining available inventory, the normalized remaining inventory amount would be 2.
    iStandardPercentSold As Integer         'for this event, for the standard avail name inventory, the total sold length in seconds divided by inventory length in seconds expressed as a percentage. A background color will be used for the percentages
    lStandardAverage30UnitRateGross As Long 'gross dollars divided by normalized 30 second sold unit count.
    lStandardGross As Long                  'total gross dollars for each spot booked in this event for the standard avail name inventory (excludes Drop-in, Billboard, and Extra inventory).
    lStandardNet As Long                    'total net dollars for each spot booked in this event and for the standard avail name inventory, using standard 15% agency commission for non-direct advertisers, and 0% for direct advertisers.
    lStandardAQHDemo As Long                'for the demo category selected on the report selectivity screen, using the default vehicle research book, display the AQH, by getting it from the research book. The selected demo category will be displayed in the column header. Header Example: "AQH A25-54".
    lStandardGrimpsRemainingAvails As Long  'AQH multiplied by normalized 30s remaining avail count. Example: if the AQH is 100, and there is 60s (normalized 30 count of 2) remaining available, the Grimps value will be 200. Note: the selected demo category will be shown in the column header.
    iHowManyParts As Integer                'When Calculating AQHDemo, the Demo is Added.  Once done, it needs to be divided by this #... to get Avg
    
    iDropInInventorySeconds As Integer      'for the Drop-in avail category, the inventory count in seconds for this event.
    iDropInSoldUnits As Integer             'for the Drop-in avail category, the sold unit count.
    iDropInRemainingUnits As Integer        'for the Drop-in avail category, the remaining avail count in units.
    
    iBillboardInventorySeconds As Integer   'for the Billboard avail category, the inventory count in seconds for this event.
    iBillboardSoldUnits As Integer          'for the Billboard avail category, the sold unit count.
    iBillboardRemainingUnits As Integer     'for the Billboard avail category, the remaining avail count in units.
    
    iExtraInventorySeconds As Integer       'for the Extra avail category, the inventory count in seconds for this event.
    iExtraSoldUnits As Integer              'for the Extra avail category, the sold unit count.
    iExtraRemainingUnits As Integer         'for the Extra avail category, the remaining avail count in units.
End Type
Dim tlGsf As GSF
Dim omBook As Object
Dim omSheet As Object
Dim imExcelRow As Integer

Dim hmDrf As Integer            'DRF_Demo_Rsrch_Data
Dim imDrfRecLen As Integer      'DRF record length
Dim tmDrf As DRF                'DRF image

Dim hmDpf As Integer            'DPF_Demo_Plus_Data
Dim imDpfRecLen As Integer      'DPF record length
Dim tmDpf As DPF                'DPF image

Dim hmDef As Integer            'DEF_Demo_Estimates
Dim imDefRecLen As Integer      'DEF record length
Dim tmDef As DEF                'DEF image

Dim hmRaf As Integer            'RAF_Region_Area
Dim imRafRecLen As Integer      'RAF record length
Dim tmRaf As RAF                'RAF image

'
'       gFilterContractTypeFromSpot - determine if spot should be processed or ignored
'       based on user parameters set up in structure CNTTYPES.
'       The spot is from SSF and uses the iamge CSSPOTSS.
'       <input>  tlSpot -image of spot from SSF
'                tlCnttypes - contract types to include/exclude
'       Return - true to process spot, else ignore
'
Private Function mFilterContractTypeFromSpot(tlSpot As CSPOTSS, tlCntTypes As CNTTYPES) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilCTypes                                                                              *
'******************************************************************************************

Dim ilSpotOK As Integer
        ilSpotOK = True                 'assume the spot is ok to include
        If (tlSpot.iRank And RANKMASK) = DIRECTRESPONSERANK Then      'DR
            If Not tlCntTypes.iDR Then
                ilSpotOK = False
            End If
        ElseIf (tlSpot.iRank And RANKMASK) = REMNANTRANK Then
            If Not tlCntTypes.iRemnant Then
                ilSpotOK = False
            End If

        ElseIf (tlSpot.iRank And RANKMASK) = PERINQUIRYRANK Then    'PI
            If Not tlCntTypes.iPI Then
                ilSpotOK = False
            End If

        ElseIf (tlSpot.iRank And RANKMASK) = TRADERANK Then  'trades
            If Not tlCntTypes.iTrade Then
                ilSpotOK = False
            End If

        ElseIf (tlSpot.iRank And RANKMASK) = PROMORANK Then  'promo
            If Not tlCntTypes.iPromo Then
                ilSpotOK = False
            End If

        ElseIf (tlSpot.iRank And RANKMASK) = PSARANK Then  'psa
            If Not tlCntTypes.iPSA Then
                ilSpotOK = False
            End If
        End If
        mFilterContractTypeFromSpot = ilSpotOK
        Exit Function
End Function
'*******************************************************************
'*                                                                 *
'*      Procedure Name:gCreateComboAvails                          *
'*                                                                 *
'*             Created:12/24/07      By:D. Hosaka                  *
'*            Modified:              By:                           *
'*                                                                 *
'*  Comments: Generate Avails Combo for Sports                     *
'*            Report name changed to Game AVails  (sports avails)  *
'       This avails report combines data from 6 other reports:
'       Inventory, Sold Units, Sold Avails, % Sold, % Available,
'       and Avg Spot Rate
'       There are 2 versions of this report:  one for sports and
'       one for non-sports.  The sports version is by game and the
'       the other is for non-sports.  NOTE: Non-sports has some
'       hooks toimplement but is not fully implemented/tested.
'       Crystal report has not been created for non-sports version.
'
'*                                                                 *
'*******************************************************************
'---------------------------------------------------------------------------------------
' Procedure : gCreateComboAvails
' DateTime  : 12/17/08 16:05
' Author    : Darlene
' Purpose   :
'---------------------------------------------------------------------------------------
'
Sub gCreateComboAvails()
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilUpper                       ilDateOK                      llRif                     *
'*  ilRdf                         ilRcf                         ilFound                   *
'*  llDateEntered                 ilNoWks                       ilSaveSort                *
'*  ilWeek                        ilLoopOnComboType             tlComboType               *
'*                                                                                        *
'******************************************************************************************

'
    Dim illoop As Integer
    Dim ilRet As Integer
    Dim slDate As String
    On Error GoTo gCreateComboAvails_Error

    ReDim ilDate(0 To 1) As Integer
    Dim ilVehicle As Integer
    Dim slNameCode As String
    Dim slName As String
    Dim slCode As String
    Dim ilVefCode As Integer
    Dim ilDay As Integer
    Dim ilIndex As Integer
    Dim llEffDate As Long
    Dim llEndDate As Long
    Dim tlCntTypes As CNTTYPES
    Dim ilDone As Integer
    Dim ilMajorSet As Integer                      'vehicle sort group
    Dim ilMinorSet As Integer                   'minor vehicle group (not used)
    Dim ilMnfMajorCode As Integer               'vehicle group mnf code
    Dim ilmnfMinorCode As Integer               'minor MNF code (not used)
    Dim llDateOK As Long
    Dim llEnteredDate As Long
    Dim ilListIndex As Integer
    Dim ilContinue As Integer
    Dim tlGrf As GRF
    Dim ilMaxLengths As Integer                 '4 or 5 lengths for user defined (if using equalize from site, 4 allowed)
    Dim ilTempEqualize As Integer
    Dim ilVGSort As Integer
    Dim ilVGItemCode As Integer
    Dim blFoundOne As Boolean
    'TTP 10434 - Event and Sports export (WWO) -- Excel!
    Dim ilAnf As Integer
    Dim ilAvailsExportIndex As Integer
    Dim ilDnfCode As Integer 'DNF_Demo_Rsrch_Names (Research Book)
    Dim ilRdfCode As Integer 'RDF_Standard_Daypart
    Dim ilMnfDemo As Integer 'MNF_Multi_Names: Demo
    Dim llRafCode As Long
    Dim llDate As Long
    Dim llTime As Long
    ReDim ilInputDays(0 To 6) As Integer
    Dim llAvgAud As Long
    Dim llPopEst As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim slRecord As String
    Dim ilRow As Integer
    Dim ilColumn As Integer
    Dim slDelimiter As String
    slDelimiter = Chr$(30) 'ASCII 30 is defined as a "Record Separator" - https://www.asciitable.com/
    Dim ilAvailGroup As Integer
    Dim blSkip As Boolean
    Dim blCancelled As Boolean 'TTP 10585 - Event and Sports Avails export: if team name or event is set to "cancelled", show row in red font color
    
    'TTP 10925 move this up so it can be used earlier
    ilListIndex = RptSelCA!lbcRptType.ListIndex     'combo avails for sports or non-sports
    
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
    imGrfRecLen = Len(tmGrf(0))

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

    'TTP 10434 - Event and Sports export (WWO) - get the Demo Code from the Selected Demo in cbcDemo using the tgDemoCode array
    'TTP 10925 - Avails Combo by Day/Week Report - Error 9 (Subscript out of Range)
    If ilListIndex = AVAILSCOMBO_SPORTS Then
        If RptSelCA.cbcDemo.ListIndex <> -1 Then
            slNameCode = tgDemoCode(RptSelCA.cbcDemo.ListIndex).sKey
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilMnfDemo = CInt(slCode)
        End If
    End If
    
    If RptSelCA!ckcMultimedia.Value = vbChecked And RptSelCA.rbcOutput(3).Value = False Then        'if including multimedia, open addl files
        hmGhf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mCloseAll
            Exit Sub
        End If
        imGhfRecLen = Len(tmGhf)

        hmIsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmIsf, "", sgDBPath & "Isf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mCloseAll
            Exit Sub
        End If
        imIsfRecLen = Len(tmIsf)

        hmMsf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmMsf, "", sgDBPath & "Msf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mCloseAll
            Exit Sub
        End If
        imMsfRecLen = Len(tmMsf)

        hmMgf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmMgf, "", sgDBPath & "Mgf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mCloseAll
            Exit Sub
        End If
        imMgfRecLen = Len(tmMgf)
    End If
    
    'TTP 10434 - Event and Sports export (WWO)
    If RptSelCA.rbcOutput(3).Value = True Then 'export
        hmDrf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mCloseAll
            Exit Sub
        End If
        imDrfRecLen = Len(tmDrf)

        hmDpf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mCloseAll
            Exit Sub
        End If
        imDpfRecLen = Len(tmDpf)

        hmDef = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mCloseAll
            Exit Sub
        End If
        imDefRecLen = Len(tmDef)

        hmRaf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            mCloseAll
            Exit Sub
        End If
        imRafRecLen = Len(tmRaf)
    End If
    

    tlCntTypes.iHold = gSetCheck(RptSelCA!ckcCType(0).Value)
    tlCntTypes.iOrder = gSetCheck(RptSelCA!ckcCType(1).Value)
    tlCntTypes.iNetwork = gSetCheck(RptSelCA!ckcCType(2).Value)
    tlCntTypes.iStandard = gSetCheck(RptSelCA!ckcCType(3).Value)
    tlCntTypes.iReserv = gSetCheck(RptSelCA!ckcCType(4).Value)
    tlCntTypes.iRemnant = gSetCheck(RptSelCA!ckcCType(5).Value)
    tlCntTypes.iDR = gSetCheck(RptSelCA!ckcCType(6).Value)
    tlCntTypes.iPI = gSetCheck(RptSelCA!ckcCType(7).Value)
    tlCntTypes.iPSA = gSetCheck(RptSelCA!ckcCType(8).Value)
    tlCntTypes.iPromo = gSetCheck(RptSelCA!ckcCType(9).Value)
    tlCntTypes.iTrade = gSetCheck(RptSelCA!ckcCType(10).Value)
    tlCntTypes.iMissed = gSetCheck(RptSelCA!ckcSpots(0).Value)
    tlCntTypes.iCharge = gSetCheck(RptSelCA!ckcSpots(1).Value)
    tlCntTypes.iZero = gSetCheck(RptSelCA!ckcSpots(2).Value)
    tlCntTypes.iADU = gSetCheck(RptSelCA!ckcSpots(3).Value)
    tlCntTypes.iBonus = gSetCheck(RptSelCA!ckcSpots(4).Value)
    tlCntTypes.iXtra = gSetCheck(RptSelCA!ckcSpots(5).Value)
    tlCntTypes.iFill = gSetCheck(RptSelCA!ckcSpots(6).Value)
    tlCntTypes.iNC = gSetCheck(RptSelCA!ckcSpots(7).Value)
    tlCntTypes.iRecapturable = gSetCheck(RptSelCA!ckcSpots(8).Value)
    tlCntTypes.iSpinoff = gSetCheck(RptSelCA!ckcSpots(9).Value)
    tlCntTypes.iMG = gSetCheck(RptSelCA!ckcSpots(10).Value)            '10-28-10
    
    tlCntTypes.iPostpone = gSetCheck(RptSelCA!ckcGameType(1).Value)
    tlCntTypes.iCancelled = gSetCheck(RptSelCA!ckcGameType(0).Value)

    If (tlCntTypes.iHold) Or (tlCntTypes.iOrder) Then        '1-26-05 set general cntr type for inclusion/exclusion if hold or ordered selected
        tlCntTypes.iCntrSpots = True
    Else
        tlCntTypes.iCntrSpots = False
    End If

    'no day selectivity; set all days valid
    For illoop = 0 To 6
        tlCntTypes.iValidDays(illoop) = True
    Next illoop

    imUsingEqualize = False
    If ilListIndex = AVAILSCOMBO_SPORTS Then                'only change the Combo avails by day/week for now;leave sports avails as it
        ilMaxLengths = 5
        imEqualizeOption = 0
    Else
        If tgSpf.sAvailEqualize = "3" Then
            imEqualizeOption = 30
            ilMaxLengths = 4
            imUsingEqualize = True
        ElseIf tgSpf.sAvailEqualize = "6" Then
            ilMaxLengths = 4
            imEqualizeOption = 60
            imUsingEqualize = True
        Else
            ilMaxLengths = 5
            imEqualizeOption = 0
        End If
    End If
    
    'TTP 10434 - Event and Sports export (WWO)
    If RptSelCA.rbcOutput(3).Value Then 'Export, setup length values; We need first Spot Length 30, and the Rest 0
        tlCntTypes.iLenHL(0) = 30
        For illoop = 1 To ilMaxLengths - 1
            tlCntTypes.iLenHL(illoop) = 0
        Next illoop
    Else
        'build the spot lengths high to lo order
        For illoop = 0 To ilMaxLengths - 1
            tlCntTypes.iLenHL(illoop) = Val(RptSelCA!edcLength(illoop))
        Next illoop
    End If
    ilDone = False
    Do While Not ilDone
        ilDone = True
        'place in spot length order
        For illoop = 1 To ilMaxLengths - 1
            If tlCntTypes.iLenHL(illoop - 1) < tlCntTypes.iLenHL(illoop) Then
                'swap the two
                ilRet = tlCntTypes.iLenHL(illoop - 1)
                tlCntTypes.iLenHL(illoop - 1) = tlCntTypes.iLenHL(illoop)
                tlCntTypes.iLenHL(illoop) = ilRet
                ilDone = False
            End If
        Next illoop
    Loop
    imHowManyHL = 0
    For illoop = 0 To ilMaxLengths - 1
        If tlCntTypes.iLenHL(illoop) > 0 Then
            fmLenRatio(illoop + 1) = imEqualizeOption / tlCntTypes.iLenHL(illoop)
            imHowManyHL = imHowManyHL + 1
        End If
    Next illoop

    'Get the vehicle group selected for sorting
    ilRet = RptSelCA!cbcGroup.ListIndex
    ilMajorSet = gFindVehGroupInx(ilRet, tgVehicleSets1())
    ilRet = gObtainRcfRifRdf()          'get the rate cards and assoc dayparts.  Dayparts required for missed spots named avails -> tgMRdf

    If ilListIndex = AVAILSCOMBO_SPORTS Then
        'get all the dates needed to work with
'        slDate = RptSelCA!edcSelCFrom.Text               'start date entred
        slDate = RptSelCA!CSI_CalFrom.Text               'start date entred 9-4-19 use csi calendar control vs edit box
        llEnteredDate = gDateValue(slDate)
'        slDate = RptSelCA!edcSelCFrom1.Text             'end date
        slDate = RptSelCA!CSI_CalTo.Text             'end date
        llEndDate = gDateValue(slDate)
    Else                                                'availscombo for non-sports not coded yet
        'get all the dates needed to work with
'        slDate = RptSelCA!edcSelCFrom.Text               'effective date entred
        slDate = RptSelCA!CSI_CalFrom.Text               'effective date entred
        llEffDate = gDateValue(slDate)
        'backup to Monday
        ilDay = gWeekDayLong(llEffDate)
        Do While ilDay <> 0
            llEffDate = llEffDate - 1
            ilDay = gWeekDayLong(llEffDate)
        Loop
        llEnteredDate = llEffDate               'save orig date to calculate week index
        llEndDate = ((Val(RptSelCA!edcSelCFrom1.Text) * 7) - 1) + llEnteredDate
        '6-28-30 add vehicle group items selection
        imMajorSet = 0
        ilVGSort = RptSelCA!cbcGroup.ListIndex
        
        If ilVGSort >= 0 And (RptSelCA!cbcGroup.ListIndex > 0) Then
            imMajorSet = gFindVehGroupInx(ilVGSort, tgVehicleSets1())
            gObtainCodesForMultipleLists 3, tgSOCode(), imInclVGCodes, imUseVGCodes(), RptSelCA
        Else
            imInclVGCodes = 0
            ReDim imUseVGCodes(0 To 0) As Integer
        End If
    End If

    ReDim tmComboData(0 To 0) As COMBODATA
    'TTP 10434 - Event and Sports export (WWO)
    ReDim tmEventSportAvailsExportData(0 To 0) As EVENTSPORTSAVAILSEXPORTDATA
    'Process and create prepass records for 1 selected vehicle at a time
    For ilVehicle = 0 To RptSelCA!lbcSelection(0).ListCount - 1 Step 1
        If (RptSelCA!lbcSelection(0).Selected(ilVehicle)) Then
            slNameCode = tgCSVNameCode(ilVehicle).sKey
            ilRet = gParseItem(slNameCode, 1, "\", slName)
            ilRet = gParseItem(slName, 3, "|", slName)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            ilVefCode = Val(slCode)
Debug.Print "Vehicle:" & slName & " (" & ilVefCode & ")"

            If RptSelCA.rbcOutput(3).Value Then
                RptSelCA.lblExportStatus.Caption = "Processing " & slName
                RptSelCA.lblExportStatus.Refresh
            End If
            If ilListIndex = AVAILSCOMBO_SPORTS Then
                blFoundOne = True
            Else
                gGetVehGrpSets ilVefCode, 0, imMajorSet, ilmnfMinorCode, ilVGItemCode
                'check selectivity of vehicle groups
                blFoundOne = True
                If (imMajorSet > 0) Then
                    blFoundOne = False
                    If gFilterLists(ilVGItemCode, imInclVGCodes, imUseVGCodes()) Then
                        blFoundOne = True
                    End If
                Else
                    blFoundOne = True
                End If
            End If
                
            If blFoundOne Then
                imGameIndFlag = False                           'for Multimedia processing only; flag to indicate to gather multimedia once per vehicle
                mInitComboData ilVefCode, tlCntTypes            'determine which named avails or dayparts to process for vehicle
                
                'process all dayparts or named avails for a vehicle
                mGetSpotCountsCA ilListIndex, ilVefCode, llEnteredDate, llEndDate, tmComboType(), tlCntTypes
                'Calculate week # for sorting
                gPackDateLong llDateOK, ilDate(0), ilDate(1)
    
                For ilIndex = 0 To UBound(tmGrf) - 1 Step 1
                    '      9/26/00 grfFields
                    '      GenDate - Generation date
                    '      GenTime - Generation Time
                    '      VefCode - vehicle code
                    '      rdfCode - DP or ANF code
                    '      adfCode - DP/ANF Sort code from RIF for vehicle or ANF
                    '      slfCode = Game # if applicable (if avails combo by game), or DP sort code from RIF (for avails combo for non-sorts)
                    '      sofCode - Week Index #
                    '      StartDate - Date of game data (1 record/day)
                    '      Date      - start date of week
                    '      Year - Day of Week (0-6, M-Su)
                    '      Code2 - mnf code for vehicle group sort
                    '      Code4 - Multimedia avails:  ghfcode (if game avails) or date as long (avails for non-sports)
                    '      ChfCode - multimedia avails:  gsfcode
                    '      DateType - flag indicating if avail name is oversold (non blank)
                    '      BktType - blank = Game Avails,D = game dependent Multimedia, I = game independent multimedia
                    '      Long - 0 if OK to print, else -1
        
                    '      lDollars(1) - Orig Inventory (seconds) by avail or multimedia
                    '      lDollars(2) -  Total seconds sold per time period (named avail or DP)
                    '                     used to calc % sellout because a spot length not highlited may
                    '                     result in oversold condition since counts would not be totally accurate
                    '                     ie.  HL 60 30 15 10 .  Avail is 2/60
                    '                     spots are 1 45 & 1 15.  Count would be 1 60 and 1 30 totally 90" sold but only 60" inventory.
                    '      lDollars(3) - Missed
                    '                    % unsold calculated in Crystal
                    '      lDollars(4) - Seconds remaining per named avail
                    '      lDollars(5) - Revenue $ per named avail or $ per multimedia
                    '      lDollars(6) - Units remaining per named avail
                    '      lDollars(7) -  Units sold per named avail or multimedia
                    '      lDollars(8) -  Counts remaining for HL #1 spot length when oversold
                    '      lDollars(9) -  Counts remaining for HL #2 spot length when oversold
                    '      lDollars(10) - Counts remaining for HL #3 spot length when oversold
                    '      lDollars(11) - Counts remaining for HL #4 spot length when oversold
                    '      lDollars(12) - Counts remaining for HL #5 spot length when oversold
                    '      lDollars(13) - Counts remaining for OTHER spot length when oversold
                    '      lDollars(14) - proposal price
                    '      lDollars(15)
                    '      lDollars(16)
                    '      lDollars(17)
                    '      lDollars(18)
                    '      PerGenl(1)   - Length #1 to highlight spot count
                    '      PerGenl(2)     Length #2    ""
                    '      PerGenl(3)     Length #3    ""
                    '      PerGenl(4)     Length #4    ""
                    '      PerGenl(5)     length #5    ""       1/24/12 if using equalize 30/60 from site, 4 max lengths allowed.  5th is used for equalize value
                    '      PerGenl(6)     All Others   ""
                    '      PerGenl(7)   - Length #1 to highlight for report hders
                    '      PerGenl(8)     Length #2    ""
                    '      PerGenl(9)     Length #3    ""
                    '      PerGenl(10)     Length #4    ""
                    '      PerGenl(11)     Length #5     ""     1/24/12 if using equalize 30/60 from site, 4 max lengths allowed.  5th is used for equalize value
                    '      Pergenl(12)     Multimedia Avails:  IHF code   (inventory header which points to inv type and item)
                    '      PerGenl(13)     length #1 to HL avail count
                    '      PerGenl(14)    length #2 to HL avail count
                    '      PerGenl(15)    length #3 to HL avail count
                    '      PerGenl(16)    length #4 to HL avail count
                    '      PerGenl(17)    length #5 to HL avail count     1/24/12 if using equalize 30/60 from site, 4 max lengths allowed.  5th is used for equalize value
                    '      PerGenl(18)    All Others to avail count
                    '      GenDesc        blank for avails detail, "Missed-Any Avail" for missed stats, "Game Info"
                    '      Long - flag to indicate to bypass the entry : user has requested to show unsold only.  If at least 1 unsold avail, show all avails

                    ilContinue = True
                    If RptSelCA!ckcUnsoldOnly(0).Value = vbChecked And RptSelCA.rbcOutput(3).Value = False Then      'show unsold only?
                        If tmGrf(ilIndex).sGenDesc = "Game Info" Then       'if this is a game info recd, there are no unsold avails left to show, only show game header
                            tlGrf.iGenDate(0) = tmGrf(ilIndex).iGenDate(0)
                            tlGrf.iGenDate(1) = tmGrf(0).iGenDate(1)
                            tlGrf.lGenTime = tmGrf(ilIndex).lGenTime
                            tlGrf.iVefCode = tmGrf(ilIndex).iVefCode
                            tlGrf.iSlfCode = tmGrf(ilIndex).iSlfCode
                            tlGrf.iSofCode = tmGrf(ilIndex).iSofCode
                            tlGrf.iStartDate(0) = tmGrf(ilIndex).iStartDate(0)
                            tlGrf.iStartDate(1) = tmGrf(ilIndex).iStartDate(1)
                            tlGrf.iDate(0) = tmGrf(ilIndex).iDate(0)
                            tlGrf.iDate(1) = tmGrf(ilIndex).iDate(1)
                            tlGrf.iCode2 = tmGrf(ilIndex).iCode2
                            tlGrf.lCode4 = tmGrf(ilIndex).lCode4
                            tlGrf.lChfCode = tmGrf(ilIndex).lChfCode
                            tlGrf.sDateType = tmGrf(ilIndex).sDateType
                            tlGrf.sBktType = tmGrf(ilIndex).sBktType
                            tlGrf.sGenDesc = "Game Info"
                            ilRet = btrInsert(hmGrf, tlGrf, imGrfRecLen, INDEXKEY0)
                            ilContinue = False
                        Else                                            'show unsold only, ignore flags with -1, create entries with flag = 0
                            If tmGrf(ilIndex).lLong = 0 Then            'ok to create, -1 is ignore entry
                                ilContinue = True
                            Else
                                ilContinue = False
                            End If
                        End If
                    End If
    
                    If ilContinue Then
                        'Setup the major sort factor
                        'If (tmGrf(ilIndex).lDollars(1) <> 0 And tmGrf(ilIndex).iRdfcode > 0) Or tmGrf(ilIndex).iRdfcode = -1 Then     'if no inventory and not the hard-coded Missed entry, dont create the record
    
                            gGetVehGrpSets ilVefCode, ilMinorSet, ilMajorSet, ilmnfMinorCode, ilMnfMajorCode
                            tmGrf(ilIndex).iCode2 = ilMnfMajorCode  'vehicle group mnf code
    
                            'Pass spot lengths to highlight for headers (grfPerGenl(7-11)
                            For illoop = 0 To imHowManyHL - 1
                                'tmGrf(ilIndex).iPerGenl(7 + ilLoop) = tlCntTypes.iLenHL(ilLoop)
                                tmGrf(ilIndex).iPerGenl(6 + illoop) = tlCntTypes.iLenHL(illoop)
                            Next illoop
                            
                            'TTP 10407 - Avails Combo: Equalize by 30 and equalize by 60 column header and calculation is wrong (Send as a Parameter instead in function Function gCmcGenCA)
                            'if equalizing 30/60, show the value in header
                            'If imUsingEqualize Then
                            '    If tgSpf.sAvailEqualize = "3" Then
                            '        'tmGrf(ilIndex).iPerGenl(11) = "30"
                            '        tmGrf(ilIndex).iPerGenl(10) = "30"
                            '    ElseIf tgSpf.sAvailEqualize = "6" Then
                            '        'tmGrf(ilIndex).iPerGenl(11) = "60"
                            '        tmGrf(ilIndex).iPerGenl(10) = "60"
                            '    End If
                            'End If
                            
                            mGetAvailsRemaining tlCntTypes, ilIndex
                            
                            'may need to swap Other fields with the Equalized field if applicable to avoid having to change all the total lines for the Other and equalized columns
                            'normal report has 5 lengths to highlight, with 6 being OTHER.  This is for both Sold and AVailable columns.
                            'if equalization 30/60s, there will be 4 lengths to highlight, along with Other (spots that dont fall into the highlighted lengths).
                            'The last(6th) column is to show the equalized number of 30s or 60s.
                            'swap the fields (5th highlight length is the equalized value which is shown in 6th column; Oth column will be shown as the 5th column.  This is for both Sold and AVails.
                            'grf.Pergenl(1-5) spot lengths to highlight .  If equalizing, grfPergenl(5) will contain the equalized value, to be printed after the OTH column.
                            If imUsingEqualize Then                                         'swp equalize and OTH columns that show in columns 5 & 6 for Sold portion of report
                                'ilTempEqualize = tmGrf(ilIndex).iPerGenl(5)                 'equalize Sold value
                                'tmGrf(ilIndex).iPerGenl(5) = tmGrf(ilIndex).iPerGenl(6)     'put OTH into equalize column
                                'tmGrf(ilIndex).iPerGenl(6) = ilTempEqualize                 'put equalize value into OTH column
                                ilTempEqualize = tmGrf(ilIndex).iPerGenl(4)                 'equalize Sold value
                                tmGrf(ilIndex).iPerGenl(4) = tmGrf(ilIndex).iPerGenl(5)     'put OTH into equalize column
                                tmGrf(ilIndex).iPerGenl(5) = ilTempEqualize                 'put equalize value into OTH column
                                
                                'ilTempEqualize = tmGrf(ilIndex).iPerGenl(17)                 'equalize Avail  value
                                'tmGrf(ilIndex).iPerGenl(17) = tmGrf(ilIndex).iPerGenl(18)     'put OTH into equalize column
                                'tmGrf(ilIndex).iPerGenl(18) = ilTempEqualize                 'put equalize value into OTH column
                                ilTempEqualize = tmGrf(ilIndex).iPerGenl(16)                 'equalize Avail  value
                                tmGrf(ilIndex).iPerGenl(16) = tmGrf(ilIndex).iPerGenl(17)     'put OTH into equalize column
                                tmGrf(ilIndex).iPerGenl(17) = ilTempEqualize                 'put equalize value into OTH column
                                
                                ''need to adjust the avg 30/60 Rate based on the number of equalized spots
                                ''tmGrf(ilIndex).lDollars(7) = tmGrf(ilIndex).iPerGenl(6)     'adjust field to make it the equalized value instead  computation of avg equalized rate
                                'tmGrf(ilIndex).lDollars(6) = tmGrf(ilIndex).iPerGenl(6)     'adjust field to make it the equalized value instead  computation of avg equalized rate
                                tmGrf(ilIndex).lDollars(6) = tmGrf(ilIndex).iPerGenl(5)     'adjust field to make it the equalized value instead  computation of avg equalized rate
                            End If
                            
'Debug.Print "Vef:" & tmGrf(ilIndex).iVefCode & ", Event#:" & tmGrf(ilIndex).iSlfCode & ", GHFCode:" & tmGrf(ilIndex).lCode4 & ", AnfCode:" & tmGrf(ilIndex).iRdfCode & ", gross:" & Format(tmGrf(ilIndex).lDollars(4) / 100, "#,###.00");

                            'TTP 10434 - Event and Sports export (WWO)
                            If RptSelCA.rbcOutput(3).Value Then 'Export, store values in tmEventSportAvailsExportData, send to Excel later
                                blSkip = False
                                '----------------------------
                                'Find which tmEventSportAvailsExportData to place this Data (Add to existing or create a New Row?)
                                blFoundOne = False
                                For illoop = 0 To UBound(tmEventSportAvailsExportData)
                                    'Match on Event Date and Event Number
                                    If tmEventSportAvailsExportData(illoop).iEventAirDate(0) = tmGrf(ilIndex).iStartDate(0) And _
                                        tmEventSportAvailsExportData(illoop).iEventAirDate(1) = tmGrf(ilIndex).iStartDate(1) And _
                                        tmEventSportAvailsExportData(illoop).iVehicle = tmGrf(ilIndex).iVefCode And _
                                        tmEventSportAvailsExportData(illoop).iEventNo = tmGrf(ilIndex).iSlfCode Then
                                        blFoundOne = True
                                        ilAvailsExportIndex = illoop
                                        Exit For
                                    End If
                                Next illoop
                                If blFoundOne = False Then
                                    ilAvailsExportIndex = UBound(tmEventSportAvailsExportData)
                                    '----------------------------
                                    ''Lookup GSF from using tmGrf(ilIndex).lCode4 = GhfCode, tmGrf(ilIndex).iVefCode = VehicleCode, and tmGrf(ilIndex).iSlfCode = GameNo
                                    tlGsf = LookupGSF(tmGrf(ilIndex).lCode4, tmGrf(ilIndex).iVefCode, tmGrf(ilIndex).iSlfCode)
                                    If tlGsf.lCode = -1 Then
'Debug.Print "Error Locating GameCode Ghf:" & tmGrf(ilIndex).lCode4 & " for Veh:" & tmGrf(ilIndex).iVefCode & ", Event#:" & tmGrf(ilIndex).iSlfCode
                                        'Exit Sub
                                        blSkip = True
                                    End If
                                End If

                                If blSkip = False Then
                                    '9/21/22 - Fix v81 - event and sports testing 9/21/22: Issue 3 - WWO database, Billboard, Drop-In, Extra not showing
                                    '----------------------------
                                    ''Lookup ANF_Avail_Names using anfCode = tmGrf(ilIndex).iRdfCode
                                    If tmGrf(ilIndex).iRdfCode <> 0 Then
                                        ilAnf = gBinarySearchAnf(tmGrf(ilIndex).iRdfCode, tgAvailAnf())
                                        If ilAnf = -1 Then
                                            MsgBox "Error Locating AvailCode:" & tmGrf(ilIndex).iRdfCode, vbCritical + vbOKOnly, "Error Exporting"
                                            Exit Sub
                                        End If
                                    Else
                                        ilAnf = -1
                                    End If
'Debug.Print ", InvSec:" & tmGrf(ilIndex).lDollars(0) & ", Sold:" & tmGrf(ilIndex).lDollars(6) & ", SoldSec:" & tmGrf(ilIndex).lDollars(1)
                                    
                                    '----------------------------
                                    'Vehicle
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).iVehicle = tmGrf(ilIndex).iVefCode
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).sVehicle = slName
                                    
                                    'Broadcast Week
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).iBroadcastWeek(0) = tmGrf(ilIndex).iDateGenl(0, 0) 'start date of week
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).iBroadcastWeek(1) = tmGrf(ilIndex).iDateGenl(1, 0) 'start date of week
                                    
                                    'Event Air Date
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).iEventAirDate(0) = tmGrf(ilIndex).iStartDate(0)    'grfStartDate = GameDate
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).iEventAirDate(1) = tmGrf(ilIndex).iStartDate(1)    'grfStartDate = GameDate
                                    
                                    'Event Time
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).iEventTime(0) = tlGsf.iAirTime(0)                  'Event Time
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).iEventTime(1) = tlGsf.iAirTime(1)                  'Event Time
                                    
                                    'Game # if applicable (grf.slfCode)
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).iEventNo = tmGrf(ilIndex).iSlfCode                 'Event/Game Number
                                    
                                    'Home Team
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).iHomeTeam = tlGsf.iHomeMnfCode                     'Home Team
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).sHomeTeam = mGetMnfName(tlGsf.iHomeMnfCode)        'Home Team
                                    
                                    'Visiting Team
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).iVisitingTeam = tlGsf.iVisitMnfCode                'Visiting Team
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).sVisitingTeam = mGetMnfName(tlGsf.iVisitMnfCode)   'Visiting Team
                                    
                                    'Event Status
                                    tmEventSportAvailsExportData(ilAvailsExportIndex).sEventStatus = tlGsf.sGameStatus                   'GameStatus - T = Tentative, F = Firm, C = Cancelled, P = Postponed
                                    
                                    '----------------------------
                                    If ilAnf = -1 Then
                                        ilAvailGroup = 0
                                    Else
                                        ilAvailGroup = tgAvailAnf(ilAnf).iEventAvailsGroup
                                    End If

                                    Select Case ilAvailGroup
                                        Case 0 'Standard
                                            'Standard Inventory Normalized to :30s
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardInventorySeconds = tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardInventorySeconds + tmGrf(ilIndex).lDollars(0)
                                            
                                            'Sold Normalized to :30s
                                            '9/21/22 - Fix v81 - event and sports testing 9/21/22: Issue 1 - "Sold Normalized to :30s"
                                            'tmEventSportAvailsExportData(ilAvailsExportIndex).iStandardSoldNormalized = tmEventSportAvailsExportData(ilAvailsExportIndex).iStandardSoldNormalized + tmGrf(ilIndex).lDollars(6)
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardSoldNormalized = tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardSoldNormalized + (tmGrf(ilIndex).lDollars(1) / 30) * 100
                                            
                                            'Sold in Seconds
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardSoldSeconds = tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardSoldSeconds + tmGrf(ilIndex).lDollars(1)
                                            
                                            'Gross
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardGross = tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardGross + tmGrf(ilIndex).lDollars(4)
                                            
                                            'Net
                                            If tmGrf(ilIndex).sDateType = "D" Then
                                                'is Direct, use 0%
                                                tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardNet = tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardNet + tmGrf(ilIndex).lDollars(4)
                                            Else
                                                'has Advertiser, use 15%
                                                tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardNet = tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardNet + tmGrf(ilIndex).lDollars(4) - (tmGrf(ilIndex).lDollars(4) * 0.15)
                                            End If
                                                                            
                                            'Lookup "Default" Research Book for this vehicle
                                            ilDnfCode = gGetDefaultBook(tmGrf(ilIndex).iVefCode)
                                            
                                            'Dummy llRafCode for gGetDemoAvgAud?
                                            llRafCode = 0
                                            
                                            'RDF DayPart
                                            If ilAnf > 0 Then
                                                ilRdfCode = mGetRDFFromAnfCode(tgAvailAnf(ilAnf).iCode)
                                            Else
                                                ilRdfCode = 0
                                            End If
                                            'Event Start Date
                                            gUnpackDateLong tmGrf(ilIndex).iStartDate(0), tmGrf(ilIndex).iStartDate(1), llDate
                                            
                                            'Event Air Time
                                            gUnpackTimeLong tlGsf.iAirTime(0), tlGsf.iAirTime(1), 1, llTime
                                            
                                            'AQH AvgAud
                                            ReDim ilInputDays(0 To 6) As Integer
                                            For illoop = 0 To 6
                                                ilInputDays(illoop) = True
                                            Next illoop
                                            ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, tmGrf(ilIndex).iVefCode, 0, _
                                                                   ilMnfDemo, llDate, llDate, ilRdfCode, 0, 0, ilInputDays(), _
                                                                   "S", llRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode)
                                            
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardAQHDemo = tmEventSportAvailsExportData(ilAvailsExportIndex).lStandardAQHDemo + llAvgAud
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).iHowManyParts = tmEventSportAvailsExportData(ilAvailsExportIndex).iHowManyParts + 1
                                            
                                        Case 1 'BillBoard
                                            'BillBoard Inventory (seconds)
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).iBillboardInventorySeconds = tmEventSportAvailsExportData(ilAvailsExportIndex).iBillboardInventorySeconds + tmGrf(ilIndex).lDollars(0)
                                            'BillBoard Sold (units)
                                            'fix 10434 - per Jason email: v81 - event and sports testing 9/21/22: Issue 3
                                            'tmEventSportAvailsExportData(ilAvailsExportIndex).iBillboardSoldUnits = tmEventSportAvailsExportData(ilAvailsExportIndex).iBillboardSoldUnits + tmGrf(ilIndex).iPerGenl(0) * 30
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).iBillboardSoldUnits = tmEventSportAvailsExportData(ilAvailsExportIndex).iBillboardSoldUnits + tmGrf(ilIndex).iPerGenl(5) 'Sold Units
                                            'JW 10/4/22: This shouldnt work against the Avg AQH
                                            'tmEventSportAvailsExportData(ilAvailsExportIndex).iHowManyParts = tmEventSportAvailsExportData(ilAvailsExportIndex).iHowManyParts + 1
                                            'fix 10434 - RE: v81 - event and sports testing 9/21/22: Issue 3 (Remaining units column)
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).iBillboardRemainingUnits = tmEventSportAvailsExportData(ilAvailsExportIndex).iBillboardRemainingUnits + tmGrf(ilIndex).lDollars(5)
                                            
                                        Case 2 'Drop-In
                                            'Drop-In Inventory (seconds)
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).iDropInInventorySeconds = tmEventSportAvailsExportData(ilAvailsExportIndex).iDropInInventorySeconds + tmGrf(ilIndex).lDollars(0)
                                            'Drop-In Sold (units)
                                            'fix per Jason email: v81 - event and sports testing 9/21/22: Issue 3
                                            'tmEventSportAvailsExportData(ilAvailsExportIndex).iDropInSoldUnits = tmEventSportAvailsExportData(ilAvailsExportIndex).iDropInSoldUnits + tmGrf(ilIndex).iPerGenl(0) * 30
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).iDropInSoldUnits = tmEventSportAvailsExportData(ilAvailsExportIndex).iDropInSoldUnits + tmGrf(ilIndex).iPerGenl(5) 'Sold Units
                                            'JW 10/4/22: This shouldnt work against the Avg AQH
                                            'tmEventSportAvailsExportData(ilAvailsExportIndex).iHowManyParts = tmEventSportAvailsExportData(ilAvailsExportIndex).iHowManyParts + 1
                                            'fix 10434 - RE: v81 - event and sports testing 9/21/22: Issue 3 (Remaining units column)
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).iDropInRemainingUnits = tmEventSportAvailsExportData(ilAvailsExportIndex).iDropInRemainingUnits + tmGrf(ilIndex).lDollars(5)
                                            
                                        Case 3 'Extra
                                            'Extra Inventory (seconds)
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).iExtraInventorySeconds = tmEventSportAvailsExportData(ilAvailsExportIndex).iExtraInventorySeconds + tmGrf(ilIndex).lDollars(0)
                                            'Extra Sold (units)
                                            'fix per Jason email: v81 - event and sports testing 9/21/22: Issue 3
                                            'tmEventSportAvailsExportData(ilAvailsExportIndex).iExtraSoldUnits = tmEventSportAvailsExportData(ilAvailsExportIndex).iExtraSoldUnits + tmGrf(ilIndex).iPerGenl(0) * 30
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).iExtraSoldUnits = tmEventSportAvailsExportData(ilAvailsExportIndex).iExtraSoldUnits + tmGrf(ilIndex).iPerGenl(5) 'Sold Units
                                            'JW 10/4/22: This shouldnt work against the Avg AQH
                                            'tmEventSportAvailsExportData(ilAvailsExportIndex).iHowManyParts = tmEventSportAvailsExportData(ilAvailsExportIndex).iHowManyParts + 1
                                            'fix 10434 - RE: v81 - event and sports testing 9/21/22: Issue 3 (Remaining units column)
                                            tmEventSportAvailsExportData(ilAvailsExportIndex).iExtraRemainingUnits = tmEventSportAvailsExportData(ilAvailsExportIndex).iExtraRemainingUnits + tmGrf(ilIndex).lDollars(5)

                                    End Select
                                    
                                    '----------------------------
                                    'Create a New Row if we're on a Different Event
                                    If blFoundOne = False Then
                                        ReDim Preserve tmEventSportAvailsExportData(0 To UBound(tmEventSportAvailsExportData) + 1) As EVENTSPORTSAVAILSEXPORTDATA
                                    End If
                                End If
                            Else
                                ilRet = btrInsert(hmGrf, tmGrf(ilIndex), imGrfRecLen, INDEXKEY0)
                            End If
                            
                       ' End If
                    End If
                Next ilIndex
            End If                              'blfoundone
        End If                                  'vehicle selected
    Next ilVehicle                              'For ilvehicle = 0 To RptSelCA!lbcSelection(0).ListCount - 1
    
    'TTP 10434 - Event and Sports export (WWO)
    If RptSelCA.rbcOutput(3).Value Then 'Export, send the data to Excel
        RptSelCA.lblExportStatus.Caption = "Generating Worksheet..."
        RptSelCA.lblExportStatus.Refresh
        Screen.MousePointer = vbHourglass
        'Create Excel
        bgExcelCreated = False
        ilRet = gExcelOutputGeneration("C")
        
        'Open (Book and Sheet): (Parameters: slAction, olBook, olSheet, ilSheetNo)
        ilRet = gExcelOutputGeneration("O", omBook, omSheet, 1)
        
        'Header 1
        ilRow = 1
        ilColumn = 1
        'slRecord = "Demo: " & RptSelCA.cbcDemo.Text & slDelimiter   'A
        slRecord = "" & slDelimiter                                 'A
        slRecord = slRecord & "" & slDelimiter                      'B
        slRecord = slRecord & "" & slDelimiter                      'C
        slRecord = slRecord & "" & slDelimiter                      'D
        slRecord = slRecord & "" & slDelimiter                      'E
        slRecord = slRecord & "" & slDelimiter                      'F
        slRecord = slRecord & "" & slDelimiter                      'G
        slRecord = slRecord & "" & slDelimiter                      'H
        slRecord = slRecord & "Standard Avail Names" & slDelimiter  'I
        slRecord = slRecord & "" & slDelimiter                      'J
        slRecord = slRecord & "" & slDelimiter                      'K
        slRecord = slRecord & "" & slDelimiter                      'L
        slRecord = slRecord & "" & slDelimiter                      'M
        slRecord = slRecord & "" & slDelimiter                      'N
        slRecord = slRecord & "" & slDelimiter                      'O
        slRecord = slRecord & "" & slDelimiter                      'P
        slRecord = slRecord & "" & slDelimiter                      'Q
        If RptSelCA.ckcIncludeAvailGroupSections.Value = vbChecked Then
            slRecord = slRecord & "Billboard" & slDelimiter             'R
            slRecord = slRecord & "" & slDelimiter                      'S
            slRecord = slRecord & "" & slDelimiter                      'T
            slRecord = slRecord & "Drop-in" & slDelimiter               'U
            slRecord = slRecord & "" & slDelimiter                      'V
            slRecord = slRecord & "" & slDelimiter                      'W
            slRecord = slRecord & "Extra" & slDelimiter                 'X
            slRecord = slRecord & "" & slDelimiter                      'Y
            slRecord = slRecord & "" & slDelimiter                      'Z
        End If
        'Write Header to Excel
        ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilRow, ilColumn, slDelimiter)
        
        'Header 2
        ilRow = 2
        ilColumn = 1
        slRecord = "Vehicle" & slDelimiter                                                              'A "Vehicle"
        slRecord = slRecord & "Broadcast Week" & slDelimiter                                            'B "Broadcast Week"
        slRecord = slRecord & "Event Air Date" & slDelimiter                                            'C "Event Air Date"
        slRecord = slRecord & "Event Time" & slDelimiter                                                'D "Event Time"
        slRecord = slRecord & "Event #" & slDelimiter                                                   'E "Event #"
        slRecord = slRecord & "Home Team" & slDelimiter                                                 'F "Home Team"
        slRecord = slRecord & "Visiting Team" & slDelimiter                                             'G "Visiting Team"
        slRecord = slRecord & "Event Status" & slDelimiter                                              'H "Event Status"
        slRecord = slRecord & "Inventory Normalized to :30s" & slDelimiter                              'I "Inventory Normalized to :30s"
        slRecord = slRecord & "Sold Normalized to :30s" & slDelimiter                                   'J "Sold Normalized to :30s"
        slRecord = slRecord & "Remaining Normalized to :30s" & slDelimiter                              'K "Remaining Normalized to :30s"
        slRecord = slRecord & "% Sold" & slDelimiter                                                    'L "% Sold"
        slRecord = slRecord & "Average 30 Unit Rate (gross)" & slDelimiter                              'M "Average 30 Unit Rate (gross)"
        slRecord = slRecord & "Gross" & slDelimiter                                                     'N "Gross"
        slRecord = slRecord & "Net" & slDelimiter                                                       'O "Net"
        slRecord = slRecord & "AQH " & RptSelCA.cbcDemo.Text & "" & slDelimiter                         'P "AQH Demo"
        slRecord = slRecord & "Grimps Remaining Avails " & RptSelCA.cbcDemo.Text & "" & slDelimiter     'Q "Grimps Remaining Avails"
        If RptSelCA.ckcIncludeAvailGroupSections.Value = vbChecked Then
            slRecord = slRecord & "Inventory (seconds)" & slDelimiter                                   'R Billboard "Inventory (seconds)"
            slRecord = slRecord & "Sold (units)" & slDelimiter                                          'S Billboard "Sold (units)"
            slRecord = slRecord & "Remaining Avails (units)" & slDelimiter                              'T Billboard "Remaining Avails (units)"
            slRecord = slRecord & "Inventory (seconds)" & slDelimiter                                   'U Drop-in "Inventory (seconds)"
            slRecord = slRecord & "Sold (units)" & slDelimiter                                          'V Drop-in "Sold (units)"
            slRecord = slRecord & "Remaining Avails (units)" & slDelimiter                              'W Drop-in "Remaining Avails (units)"
            slRecord = slRecord & "Inventory (seconds)" & slDelimiter                                   'X Extra "Inventory (seconds)"
            slRecord = slRecord & "Sold (units)" & slDelimiter                                          'Y Extra "Sold (units)"
            slRecord = slRecord & "Remaining Avails (units)" & slDelimiter                              'Z Extra "Remaining Avails (units)"
        End If
        'Write Header to Excel
        ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilRow, ilColumn, slDelimiter)
        
        'Wrap Header Text
        ilRow = 2
        For ilColumn = 1 To omSheet.UsedRange.Columns.Count Step 1
            ilRet = gExcelOutputGeneration("WT", omBook, omSheet, , True, ilRow, ilColumn)
        Next ilColumn
        
        'Export Rows of Data
        For illoop = 0 To UBound(tmEventSportAvailsExportData) - 1
            blCancelled = False 'TTP 10585 - Event and Sports Avails export: if team name or event is set to "cancelled", show row in red font color
            RptSelCA.lblExportStatus.Caption = "Generating Worksheet: Record " & illoop + 1
            RptSelCA.lblExportStatus.Refresh
            
            ilRow = illoop + 3: 'Row #, starting on Row 3 (Header is Row 1 and 2)
            
            'Calc L "% Sold" (from Inventory Normalized to :30s and Remaining Normalized to :30s (9999 = 99.99))
            'fix per Jason email: v81 - event and sports testing 9/21/22: Issue 2
            If tmEventSportAvailsExportData(illoop).lStandardInventorySeconds <> 0 Then
                tmEventSportAvailsExportData(illoop).iStandardPercentSold = ((tmEventSportAvailsExportData(illoop).lStandardSoldSeconds / (tmEventSportAvailsExportData(illoop).lStandardInventorySeconds)) * 100) * 100
            Else
                tmEventSportAvailsExportData(illoop).iStandardPercentSold = 0
            End If
            
            'calc K "Remaining Normalized to :30s"
            If (tmEventSportAvailsExportData(illoop).lStandardInventorySeconds) - (tmEventSportAvailsExportData(illoop).lStandardSoldSeconds) > 0 Then
                tmEventSportAvailsExportData(illoop).lStandardRemainingNormalized = _
                (((tmEventSportAvailsExportData(illoop).lStandardInventorySeconds) - (tmEventSportAvailsExportData(illoop).lStandardSoldSeconds)) / 30) * 100     'K "Remaining Normalized to :30s"
            Else
                tmEventSportAvailsExportData(illoop).lStandardRemainingNormalized = 0
            End If
            
            'Calc M "Average 30 Unit Rate (gross)
            'fix per Jason email: v81 TTP 10434 - Thu 9/15/22 4:26 PM
            If (tmEventSportAvailsExportData(illoop).lStandardSoldNormalized / 100) <> 0 Then
                tmEventSportAvailsExportData(illoop).lStandardAverage30UnitRateGross = tmEventSportAvailsExportData(illoop).lStandardGross / (tmEventSportAvailsExportData(illoop).lStandardSoldNormalized / 100)
            End If
            
            'Calc P "AQH Demo"
            'fix per Jason email: v81 - event and sports testing 9/21/22: Issue 2
            If tmEventSportAvailsExportData(illoop).iHowManyParts <> 0 Then
                'JW - 10/13/22 Fix per Jason Email: RE: v81 Event and Sports testing 10-5 - Thu 10/13/22 11:11 AM
                If tgSpf.sSAudData = "H" Then     'hundreds
                    tmEventSportAvailsExportData(illoop).lStandardAQHDemo = Format(tmEventSportAvailsExportData(illoop).lStandardAQHDemo / tmEventSportAvailsExportData(illoop).iHowManyParts, "#.00") * 10
                ElseIf tgSpf.sSAudData = "N" Then   'tens
                    tmEventSportAvailsExportData(illoop).lStandardAQHDemo = Format(tmEventSportAvailsExportData(illoop).lStandardAQHDemo / tmEventSportAvailsExportData(illoop).iHowManyParts, "#.00") * 1
                ElseIf tgSpf.sSAudData = "U" Then   'units
                    tmEventSportAvailsExportData(illoop).lStandardAQHDemo = Format((tmEventSportAvailsExportData(illoop).lStandardAQHDemo / 10) / tmEventSportAvailsExportData(illoop).iHowManyParts, "#.00")
                Else        'tgspf.sSAudData = "T"   'thousands
                    tmEventSportAvailsExportData(illoop).lStandardAQHDemo = Format(tmEventSportAvailsExportData(illoop).lStandardAQHDemo / tmEventSportAvailsExportData(illoop).iHowManyParts, "#.00") * 100
                End If
            Else
                tmEventSportAvailsExportData(illoop).lStandardAQHDemo = 0
            End If
            
            'Calc Q "Grimps Remaining Avails
            'Fix issue RE: v81 - event and sports testing 9/21/22 - Item #4
            'tmEventSportAvailsExportData(illoop).lStandardGrimpsRemainingAvails = (tmEventSportAvailsExportData(illoop).lStandardRemainingNormalized) * (tmEventSportAvailsExportData(illoop).lStandardAQHDemo / 100)
            tmEventSportAvailsExportData(illoop).lStandardGrimpsRemainingAvails = (tmEventSportAvailsExportData(illoop).lStandardRemainingNormalized / 100) * (tmEventSportAvailsExportData(illoop).lStandardAQHDemo / 100)
            
            'fix 10434 - RE: v81 - event and sports testing 9/21/22: Issue 3 (Remaining units column)
            'Calc T Billboard "Remaining Avails (units)"
            'tmEventSportAvailsExportData(illoop).iBillboardRemainingUnits = tmEventSportAvailsExportData(illoop).iBillboardInventorySeconds - tmEventSportAvailsExportData(illoop).iBillboardSoldUnits
            'Calc W Drop-in "Remaining Avails (units)"
            'tmEventSportAvailsExportData(illoop).iDropInRemainingUnits = tmEventSportAvailsExportData(illoop).iDropInInventorySeconds - tmEventSportAvailsExportData(illoop).iDropInSoldUnits
            'Calc Z Extra "Remaining Avails (units)"
            'tmEventSportAvailsExportData(illoop).iExtraRemainingUnits = tmEventSportAvailsExportData(illoop).iExtraInventorySeconds - tmEventSportAvailsExportData(illoop).iExtraSoldUnits
            
            'generate slValue row values String
            ilColumn = 1
            slRecord = tmEventSportAvailsExportData(illoop).sVehicle & slDelimiter                                              'A  "Vehicle"
            gUnpackDate tmEventSportAvailsExportData(illoop).iBroadcastWeek(0), tmEventSportAvailsExportData(illoop).iBroadcastWeek(1), slDate
            slRecord = slRecord & slDate & slDelimiter                                                                          'B "Broadcast Week"
            gUnpackDate tmEventSportAvailsExportData(illoop).iEventAirDate(0), tmEventSportAvailsExportData(illoop).iEventAirDate(1), slDate
            slRecord = slRecord & slDate & slDelimiter                                                                          'C "Event Air Date"
            gUnpackTime tmEventSportAvailsExportData(illoop).iEventTime(0), tmEventSportAvailsExportData(illoop).iEventTime(1), "A", "4", slDate
            slRecord = slRecord & slDate & slDelimiter                                                                          'D "Event Time"
            slRecord = slRecord & tmEventSportAvailsExportData(illoop).iEventNo & slDelimiter                                   'E "Event #"
            'TTP 10585 - Event and Sports Avails export: if team name or event is set to "cancelled", show row in red font color
            If InStr(1, LCase(tmEventSportAvailsExportData(illoop).sHomeTeam), "cancelled") > 0 Then
                blCancelled = True
            End If
            If InStr(1, LCase(tmEventSportAvailsExportData(illoop).sVisitingTeam), "cancelled") > 0 Then
                blCancelled = True
            End If
            slRecord = slRecord & tmEventSportAvailsExportData(illoop).sHomeTeam & slDelimiter                                  'F "Home Team"
            slRecord = slRecord & tmEventSportAvailsExportData(illoop).sVisitingTeam & slDelimiter                              'G "Visiting Team"
            
            Select Case tmEventSportAvailsExportData(illoop).sEventStatus                                                       'H "Event Status" - T = Tentative, F = Firm, C = Cancelled, P = Postponed
                Case "T"
                    slRecord = slRecord & "Tentative" & slDelimiter
                Case "F"
                    slRecord = slRecord & "Firm" & slDelimiter
                Case "C"
                    slRecord = slRecord & "Cancelled" & slDelimiter
                    blCancelled = True
                Case "P"
                    slRecord = slRecord & "Postponed" & slDelimiter
                Case Else
                    slRecord = slRecord & "" & slDelimiter
            End Select
            'Standard
            slRecord = slRecord & Format(tmEventSportAvailsExportData(illoop).lStandardInventorySeconds / 30, "#.00") & slDelimiter        'I "Inventory Normalized to :30s"
            
            '9/21/22 - Fix v81 - event and sports testing 9/21/22: Issue 1 - "Sold Normalized to :30s"
            'slRecord = slRecord & tmEventSportAvailsExportData(illoop).iStandardSoldNormalized & slDelimiter             'J "Sold Normalized to :30s"
            'slRecord = slRecord & Format(tmEventSportAvailsExportData(illoop).lStandardSoldSeconds / 30, "#.00") & slDelimiter              'J "Sold Normalized to :30s"
            slRecord = slRecord & Format(tmEventSportAvailsExportData(illoop).lStandardSoldNormalized / 100, "#.00") & slDelimiter            'J "Sold Normalized to :30s"
            
            slRecord = slRecord & Format(tmEventSportAvailsExportData(illoop).lStandardRemainingNormalized / 100, "#.00") & slDelimiter              'K "Remaining Normalized to :30s"
            slRecord = slRecord & Format(tmEventSportAvailsExportData(illoop).iStandardPercentSold / 100 / 100, "#.00%") & slDelimiter 'L "% Sold"
            slRecord = slRecord & Format(tmEventSportAvailsExportData(illoop).lStandardAverage30UnitRateGross / 100, "#.00") & slDelimiter 'M "Average 30 Unit Rate (gross)"
            slRecord = slRecord & tmEventSportAvailsExportData(illoop).lStandardGross / 100 & slDelimiter                       'N "Gross"
            slRecord = slRecord & Format(tmEventSportAvailsExportData(illoop).lStandardNet / 100, "#.00") & slDelimiter         'O "Net"
            slRecord = slRecord & Format(tmEventSportAvailsExportData(illoop).lStandardAQHDemo / 100, "#.00") & slDelimiter     'P "AQH Demo"
            slRecord = slRecord & tmEventSportAvailsExportData(illoop).lStandardGrimpsRemainingAvails & slDelimiter             'Q "Grimps Remaining Avails
            If RptSelCA.ckcIncludeAvailGroupSections.Value = vbChecked Then
                'Billboard
                slRecord = slRecord & tmEventSportAvailsExportData(illoop).iBillboardInventorySeconds & slDelimiter             'R Billboard "Inventory (seconds)"
                slRecord = slRecord & tmEventSportAvailsExportData(illoop).iBillboardSoldUnits & slDelimiter                    'S Billboard "Sold (units)"
                slRecord = slRecord & tmEventSportAvailsExportData(illoop).iBillboardRemainingUnits & slDelimiter               'T Billboard "Remaining Avails (units)"
                'Drop-in
                slRecord = slRecord & tmEventSportAvailsExportData(illoop).iDropInInventorySeconds & slDelimiter                'U Drop-in "Inventory (seconds)"
                slRecord = slRecord & tmEventSportAvailsExportData(illoop).iDropInSoldUnits & slDelimiter                       'V Drop-in "Sold (units)"
                slRecord = slRecord & tmEventSportAvailsExportData(illoop).iDropInRemainingUnits & slDelimiter                  'W Drop-in "Remaining Avails (units)"
                'Extra
                slRecord = slRecord & tmEventSportAvailsExportData(illoop).iExtraInventorySeconds & slDelimiter                 'X Extra "Inventory (seconds)"
                slRecord = slRecord & tmEventSportAvailsExportData(illoop).iExtraSoldUnits & slDelimiter                        'Y Extra "Sold (units)"
                slRecord = slRecord & tmEventSportAvailsExportData(illoop).iExtraRemainingUnits & slDelimiter                   'Z Extra "Remaining Avails (units)"
            End If
            
            'Write into row(ilRow), column(ilColumn): (Parameters: slAction, , olSheet, ilSheetNo, slValue, ilRow, ilColumn, slSplitDelimiter)
            ilRet = gExcelOutputGeneration("W", omBook, omSheet, , slRecord, ilRow, ilColumn, slDelimiter)
            
            'Background Color (Percent sold:) - L "% Sold"
            ilColumn = 12
            Select Case tmEventSportAvailsExportData(illoop).iStandardPercentSold / 100
                'Red: 100% or greater (one of the main uses for this export is to easily see when an event is oversold, therefore oversold events are shown in red)
                Case 100 To 30999
                    ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(Red), ilRow, ilColumn)
                'Pink: 95-99%
                Case 95 To 99
                    ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(PINK), ilRow, ilColumn)
                'Orange: 85-94%
                Case 85 To 94
                    ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(Orange), ilRow, ilColumn)
                'Yellow: 75-84%
                Case 75 To 84
                    ilRet = gExcelOutputGeneration("BC", omBook, omSheet, , str(Yellow), ilRow, ilColumn)
                'No background color if it's 74% or less
            End Select
            
            'Number format
            ilColumn = 13 'Average 30 Unit Rate (gross)
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "$#,##0.00", -1, ilColumn)
            ilColumn = 14 'Gross
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "$#,##0.00", -1, ilColumn)
            ilColumn = 15 'Net
            ilRet = gExcelOutputGeneration("NF", omBook, omSheet, , "$#,##0.00", -1, ilColumn)
            
            'Cancel color
            If blCancelled = True Then
                For ilColumn = 1 To omSheet.UsedRange.Columns.Count
                    ilRet = gExcelOutputGeneration("FC", omBook, omSheet, , str(DARKRED), ilRow, ilColumn)
                    ilRet = gExcelOutputGeneration("FX", omBook, omSheet, , "", ilRow, ilColumn)
                Next ilColumn
            End If
        Next illoop
        
        'Column Width
        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 25, , 1) 'Vehicle
        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 10, , 2) 'Broadcast Week
        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 10, , 3) 'Event Air Date
        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 10, , 4) 'Event Time
        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 8, , 5) 'Event #
        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 20, , 6) 'Home Team
        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 20, , 7) 'Visiting Team
        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 12, , 8) 'Event Status
        For ilColumn = 9 To omSheet.UsedRange.Columns.Count Step 1
            ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 12, , ilColumn)
        Next ilColumn
        ilRet = gExcelOutputGeneration("CW", omBook, omSheet, , 13, , 17) 'Grimps
        
        'Font Size
        For ilColumn = 1 To omSheet.UsedRange.Columns.Count Step 1
            'slAction, , olSheet, , , ,ilColumn
            ilRet = gExcelOutputGeneration("FS", omBook, omSheet, , "12", , ilColumn)
        Next ilColumn
        
        RptSelCA.lblExportStatus.Caption = ""
        Screen.MousePointer = vbDefault
        '"V" View generated Excel spread sheet: (Parameters: slAction)
        ilRet = MsgBox("This report will be sent to Excel for you to review and save", vbOKOnly + vbApplicationModal + vbInformation, "Event and Sports Combo")
        ilRet = gExcelOutputGeneration("V")
        
        'Close / Destroy Excel
        Set omSheet = Nothing
        Set omBook = Nothing
        'ilRet = gExcelOutputGeneration("Q")
        Set ogExcel = Nothing
    End If
    
    RptSelCA.lblExportStatus.Caption = ""
    RptSelCA.lblExportStatus.Refresh

    Erase tmGrf, tmComboData, tmComboType, tmEventSportAvailsExportData
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
    Screen.MousePointer = vbDefault
    Exit Sub

   On Error GoTo 0
   Exit Sub

gCreateComboAvails_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure gCreateComboAvails of Module RptCrCA"
End Sub
'*****************************************************************
'*                                                               *
'*                                                               *
'*                                                               *
'*      Procedure Name:mGetSpotCounts for Avails Combo report    *
'*                                                               *
'*      Created:9/27/00       By:D. Hosaka                       *
'*                                                               *
'*
'*      3-24-03 change way to test fill/extra spots.  Use SDF not SSF
'*****************************************************************
Sub mGetSpotCountsCA(ilListIndex As Integer, ilVefCode As Integer, llSDate As Long, llEDate As Long, tlComboType() As COMBOTYPE, tlCntTypes As CNTTYPES)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  llRif                         slSpotAmount                  ilKeyCode                 *
'*  ilWeekInx                                                                             *
'******************************************************************************************

'
'   Where:
'   ilListIndex - Report type
'   ilVefCode (I) - vehicle code to process
'   llSDate (I) - start date to begin searching Avails
'   llEDate (I) - end date to stop searching avails
'   tlComboType() (I) - array of Dayparts
'   tlCntTypes (I) - contract and spot types to include in search
'
'   Note: Remnants; Direct Response; per Inquiry; PSA and Promos are not
'         saved with a miss status
'         For scheduled spots the rank is used to determine if it is one
'         of the above (Direct reponse=1010; Remnant=1020; per Inquiry= 1030;
'         PSA=1060; Promo=1050.
'

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
    Dim ilRifDPSort As Integer
    Dim illoop As Integer
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
    Dim ilCTypes As Integer       'bit map of cnt types to include starting lo order bit
                                  '0 = unused, 1= unused, 2 = network, 3 = std, 4 = Reserved, 5 = remanant, 6 = DR
                                  '7 = PI, 8 = psa, 9 = promo, 10 = trade
                                  'bit 0 and 1 previously hold & order; but it conflicts with other contract types
    Dim ilSpotTypes As Integer    'bit map of spot types to include starting lo order bit:
                                  '0 = missed, 1 = charge, 2 = 0.00, 3 = adu, 4 = bonus, 5 = extra
                                  '6 = fill, 7 = n/c 8 = recapturable, 9 = spinoff, 10 - mg
    Dim ilFilterDay As Integer
    ReDim ilEvtType(0 To 14) As Integer
    Dim slChfType As String
    Dim slChfStatus As String
    Dim ilVefIndex As Integer
    Dim ilGameNo As Integer
    Dim llSpotAmount As Long
    Dim ilSpotsSoldInSec As Integer
    Dim ilSpotsSoldInUnits As Integer
    Dim ilUnitCount As Integer
    Dim ilAnfCode As Integer
    Dim ilLoInx As Integer
    Dim ilHiInx As Integer
    Dim ilSetDoNotShow As Integer
    Dim llPropPrice As Long     'proposal price of spot
    Dim ilOtherMissedLoop As Integer
    Dim ilFoundMissedDP As Integer
    ReDim tlorigllc(0 To 0) As LLC
    Dim ilLoopOnLLC As Integer
    Dim llOrigTime As Long
    Dim ilTemp As Integer
    Dim ilGsfFound As Integer
    Dim blGrossNet As Boolean                   'Date: 1/15/2020 added Gross/Net options
    
    'Date: 1/17/2020 added Net option
    blGrossNet = True                           'default to Gross
    If RptSelCA!rbcGrossNet(1).Value = True Then blGrossNet = False

    slType = "O"
    ilType = 0
    llLatestDate = gGetLatestLCFDate(hmLcf, "C", ilVefCode)
    'set the type of events to get fro the day (only Contract avails)
    For illoop = LBound(ilEvtType) To UBound(ilEvtType) Step 1
        ilEvtType(illoop) = False
    Next illoop
    ilEvtType(2) = True
    imSdfRecLen = Len(tmSdf)
    imCHFRecLen = Len(tmChf)
    ilVefIndex = gBinarySearchVef(ilVefCode)
    'maintain running lowest/highest index by the game.  For each game, if user wants to see only unsold avails,
    'need to cycle thru the table to determine if anything to print.  If not, set flag in record to ignore it.
    ilLoInx = 0
    ilHiInx = 0
    For llLoopDate = llSDate To llEDate Step 1
        ilFilterDay = gWeekDayLong(llLoopDate)
        If tlCntTypes.iValidDays(ilFilterDay) Then      'Has this day of the week been selected?
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
                gUnpackDateLong tmSsf.iDate(0), tmSsf.iDate(1), llDate
                ilBucketIndex = gWeekDayLong(llDate)        'day of week bucket index
                ilEvt = 1
                ilGameNo = tmSsf.iType                  'game number, if applicable

                If ilListIndex = AVAILSCOMBO_SPORTS Then
                    ilGsfFound = False
                    tmGsfSrchKey3.iGameNo = ilGameNo
                    tmGsfSrchKey3.iVefCode = ilVefCode
                    ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                    
                    Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = ilVefCode) And (tmGsf.iGameNo = tmSsf.iType)
                        If (tmGsf.iAirDate(0) = tmSsf.iDate(0)) And (tmGsf.iAirDate(1) = tmSsf.iDate(1)) Then
                            ilGsfFound = True
                            Exit Do
                        End If
                        ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                    Loop
                    
                    'If ilRet <> BTRV_ERR_NONE Then          'no game applicable, ignore this entry
                    If Not ilGsfFound Then          'no game applicable, ignore this entry
                        ilEvt = tmSsf.iCount + 1            'force exit
                    Else
                        'determine if postponed or cancelled
                        If (tmGsf.sGameStatus = "C" And Not tlCntTypes.iCancelled) Or (tmGsf.sGameStatus = "P" And Not tlCntTypes.iPostpone) Then
                            ilEvt = tmSsf.iCount + 1        'force exit
                        End If
                    End If
                End If

                'Build the events that originally made up this day in case the day has been posted in the past
                'only build the avails to retreive the original inventory
                ReDim tlorigllc(0 To 0) As LLC
                ilRet = gBuildEventDay(ilType, "C", ilVefCode, slDate, "12M", "12M", ilEvtType(), tlorigllc())


                Do While ilEvt <= tmSsf.iCount
                   LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    If tmProg.iRecType = 1 Then    'Program (not working for nested prog)
                        ilLtfCode = tmProg.iLtfCode
                    ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 2) Then 'Contract Avails only
                       LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llTime
                        'Determine which rate card program this is associated with
                        'If Game AVails, theres only 1 daypart entry, defined as 12m-12m, Mo-su
                        For ilRdf = LBound(tlComboType) To UBound(tlComboType) - 1 Step 1
                            
                            If (tmAvail.ianfCode = tlComboType(ilRdf).lKeyCode And ilListIndex = AVAILSCOMBO_SPORTS) Or ilListIndex <> AVAILSCOMBO_SPORTS Then
                                ilAvailOk = False
                                ilRifDPSort = 32000         'no DP sort Code in Items record, make it last

                                For illoop = LBound(tlComboType(ilRdf).iStartTime, 2) To UBound(tlComboType(ilRdf).iStartTime, 2) Step 1 'Row
                                    If (tlComboType(ilRdf).iStartTime(0, illoop) <> 1) Or (tlComboType(ilRdf).iStartTime(1, illoop) <> 0) Then
                                        gUnpackTimeLong tlComboType(ilRdf).iStartTime(0, illoop), tlComboType(ilRdf).iStartTime(1, illoop), False, llStartTime
                                        gUnpackTimeLong tlComboType(ilRdf).iEndTime(0, illoop), tlComboType(ilRdf).iEndTime(1, illoop), True, llEndTime
                                        'If (llTime >= llStartTime) And (llTime < llEndTime) And (tlComboType(ilRdf).sWkDays(ilLoop, ilBucketIndex + 1) = "Y") Then
                                        If (llTime >= llStartTime) And (llTime < llEndTime) And (tlComboType(ilRdf).sWkDays(illoop, ilBucketIndex) = "Y") Then
                                            ilAvailOk = True
                                            ilLoopIndex = ilRdf
                                            Exit For
                                        End If
                                    End If
                                Next illoop

                                If ilAvailOk Then       'valid daypart or named avail, and avail time falls within the valid times
                                    ilSpotsSoldInSec = 0      'running total of sec found to be scheduled and highlited
                                    ilSpotsSoldInUnits = 0    'running total of units found to be scheduled and highlighted
                                    If tlComboType(ilLoopIndex).sInOut = "I" Then   'Book into
                                        If ilListIndex = AVAILSCOMBO_SPORTS Then
                                            If tmAvail.ianfCode <> tlComboType(ilLoopIndex).ianfCode Then
                                                ilAvailOk = False
                                            End If
                                        Else
                                            If tmAvail.ianfCode <> tlComboType(ilLoopIndex).ianfCode Then
                                                ilAvailOk = False
                                            End If
                                        End If
                                    ElseIf tlComboType(ilLoopIndex).sInOut = "O" Then   'Exclude
                                        If ilListIndex = AVAILSCOMBO_SPORTS Then
                                            If tmAvail.ianfCode = tlComboType(ilLoopIndex).lKeyCode Then
                                                ilAvailOk = False
                                            End If
                                        Else
                                            If tmAvail.ianfCode = tlComboType(ilLoopIndex).ianfCode Then
                                                ilAvailOk = False
                                            End If
                                        End If
                                    ElseIf tlComboType(ilLoopIndex).sInOut = "A" Then       'All avails; initialized when building the named avails table
                                        'include all avails, no testing
                                    End If

                                     '7-19-04 the Named avail property must allow local spots to be included
                                    tmAnfSrchKey.iCode = tmAvail.ianfCode
                                    ilRet = btrGetEqual(hmAnf, tmAnf, imAnfRecLen, tmAnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                    If (ilRet = BTRV_ERR_NONE) Then
                                        If Not (tlCntTypes.iHold Or tlCntTypes.iOrder) And tmAnf.sBookLocalFeed = "L" Then      'Local avail requested to be excluded, exclude if avail type = "L"
                                            ilAvailOk = False
                                        End If
                                        If Not tlCntTypes.iNetwork And tmAnf.sBookLocalFeed = "F" Then      'Network avail requested to be excluded, exclude if avail type = "F"
                                            ilAvailOk = False
                                        End If
                                    End If
                                    'allow the avail to be gathered if the field doesnt have a value, indicating an original avail defined as Both
                                    'allow the avail to be gathered even if the named avail code isnt found
                                End If
                                If ilAvailOk Then
                                    ilRemLen = tmAvail.iLen
                                    ilRemUnits = tmAvail.iAvInfo And &H1F
                                    'see if the orig inventory is less than whats found; if in past inventory may have been due to adding spots and
                                    'overbooking.  Want to see that the avail is oversold.
                                    For ilLoopOnLLC = LBound(tlorigllc) To UBound(tlorigllc) - 1
                                        'llTime = time of avail in SSF
                                        llOrigTime = gTimeToLong(tlorigllc(ilLoopOnLLC).sStartTime, True)
                                        If llTime = llOrigTime Then
                                            If ilRemUnits > tlorigllc(ilLoopOnLLC).iUnits Then
                                                ilRemUnits = tlorigllc(ilLoopOnLLC).iUnits
                                            End If
                                            ilTemp = CInt(gLengthToCurrency(tlorigllc(ilLoopOnLLC).sLength))

                                            If ilRemLen > ilTemp Then
                                                ilRemLen = ilTemp
                                            End If
                                            Exit For
                                        End If
                                    Next ilLoopOnLLC

                                    ilSpotsSoldInSec = 0
                                    ilSpotsSoldInUnits = 0
                                    'Determine if Grf created - create 1 record per day per daypart
                                    ilFound = False
                                    For ilRec = 0 To UBound(tmGrf) - 1 Step 1
                                        If ilListIndex = AVAILSCOMBO_SPORTS Then
                                            'match on daypart code or named avail code  and day of week
                                            'grf.irdfcode is common field for named avail code or daypart code
                                            If (tmGrf(ilRec).iRdfCode = tlComboType(ilLoopIndex).lKeyCode) And (tmGrf(ilRec).iSlfCode = ilGameNo) And (tmGrf(ilRec).lCode4 = tmGsf.lghfcode) Then '(tmGrf(ilRec).iYear = ilBucketIndex) And (tmGrf(ilRec).iSlfCode = ilGame) Then
                                                ilFound = True
                                                ilRecIndex = ilRec
                                                Exit For
                                            End If
                                        Else
                                            'match on daypart  and day of week
                                            If (tmGrf(ilRec).iRdfCode = tlComboType(ilLoopIndex).lKeyCode) And (tmGrf(ilRec).iStartDate(0) = ilDate0 And tmGrf(ilRec).iStartDate(1) = ilDate1) Then  '(tmGrf(ilRec).iYear = ilBucketIndex) And (tmGrf(ilRec).iSlfCode = tmSdf.iGameNo) Then
                                                ilFound = True
                                                ilRecIndex = ilRec
                                                Exit For
                                            End If
                                        End If
                                    Next ilRec

                                    If Not ilFound Then
                                        ilRecIndex = UBound(tmGrf)
                                        tmGrf(ilRecIndex).iGenDate(0) = igNowDate(0)
                                        tmGrf(ilRecIndex).iGenDate(1) = igNowDate(1)
                                        gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                        tmGrf(ilRecIndex).lGenTime = lgNowTime
                                        tmGrf(ilRecIndex).iVefCode = ilVefCode
                                        tmGrf(ilRecIndex).iYear = ilBucketIndex        'day of week
                                        tmGrf(ilRecIndex).iStartDate(0) = ilDate0       'date of avails
                                        tmGrf(ilRecIndex).iStartDate(1) = ilDate1
                                        tmGrf(ilRecIndex).iSofCode = (llLoopDate - llSDate) \ 7 + 1          'get relative week index for sorting in Crystal
                                        tmGrf(ilRecIndex).lCode4 = llLoopDate                               'date (as a number) for sorting
                                        'get start of week to show for week total line
                                        ilDayIndex = gWeekDayLong(llLoopDate)
                                        llDate = llLoopDate
                                        Do While ilDayIndex <> 0
                                            llDate = llDate - 1
                                            ilDayIndex = gWeekDayLong(llDate)
                                        Loop
                                        'gPackDateLong llDate, tmGrf(ilRecIndex).iDateGenl(0, 1), tmGrf(ilRecIndex).iDateGenl(1, 1)   'start date ofweek for week total caption
                                        gPackDateLong llDate, tmGrf(ilRecIndex).iDateGenl(0, 0), tmGrf(ilRecIndex).iDateGenl(1, 0)   'start date ofweek for week total caption

                                        tmGrf(ilRecIndex).iAdfCode = tlComboType(ilLoopIndex).iSortCode      '
                                        tmGrf(ilRecIndex).iRdfCode = tlComboType(ilLoopIndex).lKeyCode
                                        tmGrf(ilRecIndex).iSlfCode = tlComboType(ilLoopIndex).lKeyCode
                                        tmGrf(ilRecIndex).sDateType = ""                'assume not an oversold avail name
                                        tmGrf(ilRecIndex).sBktType = ""                 'game avail (vs multimedia)
                                        tmGrf(ilRecIndex).lLong = 0                     'flag for later to determine if only unsold avails to be shown
                                        If ilListIndex = AVAILSCOMBO_SPORTS Then
                                            tmGrf(ilRecIndex).lChfCode = tmGsf.lCode
                                            tmGrf(ilRecIndex).lCode4 = tmGsf.lghfcode
                                            tmGrf(ilRecIndex).sGenDesc = ""             'description indicating if game header, missed spots or just avails detail
                                            tmGrf(ilRecIndex).iSlfCode = ilGameNo          'game number, if applicable
                                        End If
                                        ReDim Preserve tmGrf(0 To ilRecIndex + 1) As GRF
                                    End If

                                    'Count of Inventory in seconds
                                    'tmGrf(ilRecIndex).lDollars(1) = tmGrf(ilRecIndex).lDollars(1) + ilRemLen
                                    tmGrf(ilRecIndex).lDollars(0) = tmGrf(ilRecIndex).lDollars(0) + ilRemLen

                                    For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt + ilSpot)
                                        'ilSpotOK = True                             'assume spot is OK to include
                                        ilSpotTypes = 0
                                        ilCTypes = 0

                                        ilSpotOK = mFilterContractTypeFromSpot(tmSpot, tlCntTypes)

                                        If ilSpotOK Then                            'continue testing other filters
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

                                                ilCTypes = &H4          'flag as Feed spot
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
                                                        mFilterSpot ilVefCode, tlCntTypes, ilSpotOK, ilCTypes, ilSpotTypes
                                                    End If
                                                End If
                                            End If


                                            If ilSpotOK Then
                                                llSpotAmount = mGetSpotPrice(llPropPrice)   'get spot price and proposal price
                                                'Date: 1/21/2020 added Net option
                                                If blGrossNet = False Then 'Net
                                                    llSpotAmount = gGetGrossOrNetFromRate(llSpotAmount, "N", tmChf.iAgfCode)
                                                    llPropPrice = gGetGrossOrNetFromRate(llPropPrice, "N", tmChf.iAgfCode)
                                                End If
                                                'accumulate gross $ sold
                                                'tmGrf(ilRecIndex).lDollars(5) = tmGrf(ilRecIndex).lDollars(5) + llSpotAmount
                                                tmGrf(ilRecIndex).lDollars(4) = tmGrf(ilRecIndex).lDollars(4) + llSpotAmount
                                                'tmGrf(ilRecIndex).lDollars(14) = tmGrf(ilRecIndex).lDollars(14) + llPropPrice
                                                tmGrf(ilRecIndex).lDollars(13) = tmGrf(ilRecIndex).lDollars(13) + llPropPrice
                                                
                                                'accumulate 30" units.  Spots divisible by 30" = one unit each.  Anything less/equal to 30 is a unit
                                                ilUnitCount = tmSdf.iLen \ 30
                                                If (tmSdf.iLen Mod 30) > 0 Then
                                                    ilUnitCount = ilUnitCount + 1
                                                End If
                                                'tmGrf(ilRecIndex).lDollars(7) = tmGrf(ilRecIndex).lDollars(7) + ilUnitCount
                                                tmGrf(ilRecIndex).lDollars(6) = tmGrf(ilRecIndex).lDollars(6) + ilUnitCount

                                                'Count of spot sold in seconds to calc % sellout
                                                'tmGrf(ilRecIndex).lDollars(2) = tmGrf(ilRecIndex).lDollars(2) + tmSdf.iLen
                                                tmGrf(ilRecIndex).lDollars(1) = tmGrf(ilRecIndex).lDollars(1) + tmSdf.iLen
                                                ilSpotsSoldInSec = ilSpotsSoldInSec + tmSdf.iLen
                                                ilSpotsSoldInUnits = ilSpotsSoldInUnits + 1
                                                mAccumSpotsSoldByLen tlCntTypes, ilRecIndex           'accum sold by highlighted lengths
                                                
                                                'TTP 10434 - Event and Sports export (WWO)
                                                If RptSelCA.rbcOutput(3).Value = True Then  'Export
                                                    tmGrf(ilRecIndex).iPerGenl(1) = tmGrf(ilRecIndex).iPerGenl(1) + tmSdf.iLen 'Inventory Seconds
                                                    tmGrf(ilRecIndex).iPerGenl(2) = tmGrf(ilRecIndex).iPerGenl(2) + ilSpotsSoldInUnits 'Sold Units - fix Event & Sports Export per Jason email: v81 - event and sports testing 9/21/22: Issue 3
                                                    tmGrf(ilRecIndex).iPerGenl(2) = tmGrf(ilRecIndex).iPerGenl(3) + ilUnitCount
                                                    If tmSdf.iAdfCode <> 0 Then
                                                        tmGrf(ilRecIndex).sDateType = "A" 'Has Advertiser (For calculating NET)
                                                    Else
                                                        tmGrf(ilRecIndex).sDateType = "D" 'Is Direct (For calculating NET)
                                                    End If
                                                End If
                                            End If                              'ilspotOK
                                        End If
                                    Next ilSpot                                 'loop from ssf file for # spots in avail
                                    'tmGrf(ilRecIndex).lDollars(4) = tmGrf(ilRecIndex).lDollars(4) + (ilRemLen - ilSpotsSoldInSec)      'accum total second unsold
                                    tmGrf(ilRecIndex).lDollars(3) = tmGrf(ilRecIndex).lDollars(3) + (ilRemLen - ilSpotsSoldInSec)      'accum total second unsold
                                    'tmGrf(ilRecIndex).lDollars(6) = tmGrf(ilRecIndex).lDollars(6) + (ilRemUnits - ilSpotsSoldInUnits)   'accum total units unsold
                                    tmGrf(ilRecIndex).lDollars(5) = tmGrf(ilRecIndex).lDollars(5) + (ilRemUnits - ilSpotsSoldInUnits)   'accum total units unsold
                                End If                                          'Avail OK
                            End If                                          'anfcode = tlComboType(ilRdf).lkeycode
                        Next ilRdf                                          'ilRdf = lBound(tlComboType)
                        ilEvt = ilEvt + tmAvail.iNoSpotsThis                'bypass spots
                    End If
                    ilEvt = ilEvt + 1   'Increment to next event
                Loop                                                        'do while ilEvt <= tmSsf.iCount

                'see if user requested to print unsold only.  if theres at least 1 avail unsold, show them all.
                'if everything is unsold, dont show anything but header
                If RptSelCA!ckcUnsoldOnly(0).Value = vbChecked And RptSelCA.rbcOutput(3).Value = False Then
                    ilSetDoNotShow = True
                    ilHiInx = UBound(tmGrf) - 1

                    For illoop = ilLoInx To ilHiInx
                        'If tmGrf(ilLoop).iPerGenl(13) + tmGrf(ilLoop).iPerGenl(14) + tmGrf(ilLoop).iPerGenl(15) + tmGrf(ilLoop).iPerGenl(16) + tmGrf(ilLoop).iPerGenl(17) + tmGrf(ilLoop).iPerGenl(18) > 0 Then
                        If tmGrf(illoop).iPerGenl(12) + tmGrf(illoop).iPerGenl(14) + tmGrf(illoop).iPerGenl(14) + tmGrf(illoop).iPerGenl(15) + tmGrf(illoop).iPerGenl(16) + tmGrf(illoop).iPerGenl(17) > 0 Then
                            ilSetDoNotShow = False
                            Exit For
                        End If
                    Next illoop
                    If ilSetDoNotShow Then
                        For illoop = ilLoInx To ilHiInx
                            If illoop = ilLoInx Then
                                tmGrf(illoop).sGenDesc = "Game Info"    'need to create this in crystal for the game info to print
                            Else
                                tmGrf(illoop).lLong = -1            'flag to bypass this avail; its sold out
                            End If
                        Next illoop
                    End If
                End If

                'finished with game or day, if game avails, do the multimedia avails
                If RptSelCA!ckcMultimedia.Value = vbChecked And RptSelCA.rbcOutput(3).Value = False Then
                    mGenMMAvailsForCombo ilVefCode
                End If
                imSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilRet = gSSFGetNext(hmSsf, tmSsf, imSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                If tgMVef(ilVefIndex).sType = "G" Then
                    ilType = tmSsf.iType
                End If
                ilLoInx = UBound(tmGrf)

            Loop
        End If
    Next llLoopDate

    'Get missed
    '3/30/99 For each missed status (missed, ready, & unscheduled) there are up to 2 passes
    'for each spot.  The 1st pass looks or a daypart that matches the shedule lines DP.
    'If found, the missed spot is placed in that DP (if that DP is to be shown on the report).
    'If no DP are found that match, the 2nd pass places it in the first DP that surrounds
    'the missed spots time.
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
                ilFilterDay = gWeekDayLong(llDate)

                'missed spot must be within selected date parameters, spot must be selected for inclusion (air time vs feed spot), and the day selectivity must be OK (currently no selectivity, all days always included)
                If (llDate >= llSDate And llDate <= llEDate) And ((tlCntTypes.iCntrSpots = True And tmSdf.lChfCode > 0) Or (tlCntTypes.iNetwork = True And tmSdf.lChfCode = 0)) And (tlCntTypes.iValidDays(ilFilterDay)) Then        'Has this day of the week been selected?
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
                        For ilOtherMissedLoop = 1 To 2

                            'find the associated DP or named avail
                            For ilRdf = LBound(tlComboType) To UBound(tlComboType) - 1 Step 1
                                'If (tgMRdf(ilAnfCode).ianfCode = tlComboType(ilRdf).lKeyCode And ilListIndex = AVAILSCOMBO_SPORTS) Or (tmClf.iRdfcode = tlComboType(ilRdf).lKeyCode And ilListIndex = AVAILSCOMBO_NONSPORTS) Or ilOtherMissedLoop = 2 Then

                                For illoop = LBound(tlComboType(ilRdf).iStartTime, 2) To UBound(tlComboType(ilRdf).iStartTime, 2) Step 1 'Row
                                    If (tlComboType(ilRdf).iStartTime(0, illoop) <> 1) Or (tlComboType(ilRdf).iStartTime(1, illoop) <> 0) Then
                                        gUnpackTimeLong tlComboType(ilRdf).iStartTime(0, illoop), tlComboType(ilRdf).iStartTime(1, illoop), False, llStartTime
                                        gUnpackTimeLong tlComboType(ilRdf).iEndTime(0, illoop), tlComboType(ilRdf).iEndTime(1, illoop), True, llEndTime
                                        'If (llTime >= llStartTime) And (llTime < llEndTime) And (tlComboType(ilRdf).sWkDays(ilLoop, ilBucketIndex + 1) = "Y") Then
                                        If (llTime >= llStartTime) And (llTime < llEndTime) And (tlComboType(ilRdf).sWkDays(illoop, ilBucketIndex) = "Y") Then
                                            ilAvailOk = True

'                                                If tlComboType(ilLoopIndex).sInOut = "I" Then   'Book into
'                                                    If ilListIndex = AVAILSCOMBO_SPORTS Then
'                                                        If tmAvail.ianfCode <> tlComboType(ilLoopIndex).ianfCode Then
'                                                            ilAvailOk = False
'                                                        End If
'                                                    Else
'                                                        If tmAvail.ianfCode <> tlComboType(ilLoopIndex).ianfCode Then
'                                                            ilAvailOk = False
'                                                        End If
'                                                    End If
'                                                ElseIf tlComboType(ilLoopIndex).sInOut = "O" Then   'Exclude
'                                                    If ilListIndex = AVAILSCOMBO_SPORTS Then
'                                                        If tmAvail.ianfCode = tlComboType(ilLoopIndex).lKeyCode Then
'                                                            ilAvailOk = False
'                                                        End If
'                                                    Else
'                                                        If tmAvail.ianfCode = tlComboType(ilLoopIndex).ianfCode Then
'                                                            ilAvailOk = False
'                                                        End If
'                                                    End If
'                                                ElseIf tlComboType(ilLoopIndex).sInOut = "A" Then       'All avails; initialized when building the named avails table
'                                                    'include all avails, no testing
'                                                End If
                                            If ilOtherMissedLoop = 1 Then           'pass one must find matching daypart with sch line
                                                If (tgMRdf(ilAnfCode).ianfCode <> tlComboType(ilRdf).lKeyCode And ilListIndex = AVAILSCOMBO_SPORTS) Or (tmClf.iRdfCode <> tlComboType(ilRdf).lKeyCode And ilListIndex = AVAILSCOMBO_NONSPORTS) Then
                                                    ilAvailOk = False
                                                Else
                                                    ilLoopIndex = ilRdf
                                                    ilFoundMissedDP = True      'missed spot daypart found
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                Next illoop
                                If (ilAvailOk) Or ((ilOtherMissedLoop = 2) And (Not ilFoundMissedDP)) Then

                                    ilSpotOK = True                'assume spot is OK
                                    If tmSdf.lChfCode <> tmChf.lCode Then               'if already in mem, don't reread
                                       tmChfSrchKey.lCode = tmSdf.lChfCode
                                       ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)  'Get first record as starting point of extend operation
                                       If ilRet <> BTRV_ERR_NONE Then
                                           ilSpotOK = False
                                       End If
                                   End If
                                   If tmChf.sType = "T" Then
                                       ilSpotTypes = &H10
                                       If Not tlCntTypes.iRemnant Then
                                           ilSpotOK = False
                                       End If

                                   ElseIf tmChf.sType = "Q" Then
                                       ilSpotTypes = &H800
                                       If Not tlCntTypes.iPI Then
                                           ilSpotOK = False
                                       End If

                                   ElseIf tmChf.iPctTrade = 100 Then
                                       ilSpotTypes = &H400
                                       If Not tlCntTypes.iTrade Then
                                           ilSpotOK = False
                                       End If

                                   ElseIf tmSdf.sSpotType = "X" Then
                                       ilSpotTypes = &H20
                                       If Not tlCntTypes.iXtra Then
                                           ilSpotOK = False
                                       End If

                                   ElseIf tmChf.sType = "M" Then
                                       ilSpotTypes = &H20
                                       If Not tlCntTypes.iPromo Then
                                           ilSpotOK = False
                                       End If

                                   ElseIf tmChf.sType = "S" Then
                                       ilSpotTypes = &H100
                                       If Not tlCntTypes.iPSA Then
                                           ilSpotOK = False
                                       End If
                                   End If

                                   mFilterSpot ilVefCode, tlCntTypes, ilSpotOK, ilCTypes, ilSpotTypes
                                'End If

                                    If ilSpotOK Then

                                        ilFound = False
                                        For ilRec = 0 To UBound(tmGrf) - 1 Step 1
                                            'if the missed spot didnt have an associated DP; put in category by itself
                                            If (Not ilFoundMissedDP) And (ilOtherMissedLoop = 2) Then   'processing othermissed spots
                                                If tmGrf(ilRec).iRdfCode = -1 And tmGrf(ilRec).iStartDate(0) = tmSdf.iDate(0) And tmGrf(ilRec).iStartDate(1) = tmSdf.iDate(1) Then
                                                    ilFound = True
                                                    ilRecIndex = ilRec
                                                    ilFoundMissedDP = True
                                                    Exit For
                                                End If
                                            Else
                                                If ilListIndex = AVAILSCOMBO_SPORTS Then
                                                    '4-25-19 get the correct game reference (ttp 9314)
                                                    ilGsfFound = False
                                                    tmGsfSrchKey3.iGameNo = tmSdf.iGameNo
                                                    tmGsfSrchKey3.iVefCode = tmSdf.iVefCode
                                                    ilRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                    
                                                    Do While (ilRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = tmSdf.iVefCode) And (tmGsf.iGameNo = tmSdf.iGameNo)
                                                        If (tmGsf.iAirDate(0) = tmSdf.iDate(0)) And (tmGsf.iAirDate(1) = tmSdf.iDate(1)) Then
                                                            If (tmGsf.sGameStatus = "C" And Not tlCntTypes.iCancelled) Or (tmGsf.sGameStatus = "P" And Not tlCntTypes.iPostpone) Then
                                                                ilGsfFound = False
                                                            Else
                                                                ilGsfFound = True
                                                                Exit Do
                                                            End If
                                                        End If
                                                        ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                    Loop
                                                    
                                                    'If ilRet <> BTRV_ERR_NONE Then          'no game applicable, ignore this entry
                                                    If Not ilGsfFound Then          'no game applicable, ignore this entry
                                                        ilEvt = tmSsf.iCount + 1            'force exit
                                                    Else
                                                        'determine if postponed or cancelled
                                                        If (tmGsf.sGameStatus = "C" And Not tlCntTypes.iCancelled) Or (tmGsf.sGameStatus = "P" And Not tlCntTypes.iPostpone) Then
                                                            ilEvt = tmSsf.iCount + 1        'force exit
                                                        End If
                                                    End If
                                                    'match on daypart (or named avail code) and day of week
                                                    If (tmGrf(ilRec).iRdfCode = tlComboType(ilLoopIndex).lKeyCode) And (tmGrf(ilRec).iSlfCode = tmSdf.iGameNo) Then  '(tmGrf(ilRec).iYear = ilBucketIndex) And (tmGrf(ilRec).iSlfCode = tmSdf.iGameNo) Then
                                                        ilFound = True
                                                        ilRecIndex = ilRec
                                                        Exit For
                                                    End If
                                                Else
                                                    'match on daypart  and day of week
                                                    If (tmGrf(ilRec).iRdfCode = tlComboType(ilLoopIndex).lKeyCode) And (tmGrf(ilRec).iStartDate(0) = tmSdf.iDate(0) And tmGrf(ilRec).iStartDate(1) = tmSdf.iDate(1)) Then  '(tmGrf(ilRec).iYear = ilBucketIndex) And (tmGrf(ilRec).iSlfCode = tmSdf.iGameNo) Then
                                                        ilFound = True
                                                        ilRecIndex = ilRec
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        Next ilRec
                                        If Not ilFound Then
                                            ilRecIndex = UBound(tmGrf)
                                            tmGrf(ilRecIndex).iGenDate(0) = igNowDate(0)
                                            tmGrf(ilRecIndex).iGenDate(1) = igNowDate(1)
                                            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
                                            tmGrf(ilRecIndex).lGenTime = lgNowTime
                                            tmGrf(ilRecIndex).iVefCode = ilVefCode
                                            tmGrf(ilRecIndex).iYear = ilBucketIndex        'day of week
                                            tmGrf(ilRecIndex).iStartDate(0) = tmSdf.iDate(0)       'date of avails
                                            tmGrf(ilRecIndex).iStartDate(1) = tmSdf.iDate(1)
                                            tmGrf(ilRecIndex).iSofCode = (llDate - llSDate) / 7 + 1          'get relative week index for sorting in Crystal
                                            tmGrf(ilRecIndex).lCode4 = llDate                               'date (as a number) for sorting
                                            'get start of week to show for week total line
                                            ilDayIndex = gWeekDayLong(llDate)
                                            Do While ilDayIndex <> 0
                                                llDate = llDate - 1
                                                ilDayIndex = gWeekDayLong(llDate)
                                            Loop
                                            'gPackDateLong llDate, tmGrf(ilRecIndex).iDateGenl(0, 1), tmGrf(ilRecIndex).iDateGenl(1, 1)   'start date ofweek for week total caption
                                            gPackDateLong llDate, tmGrf(ilRecIndex).iDateGenl(0, 0), tmGrf(ilRecIndex).iDateGenl(1, 0)   'start date ofweek for week total caption
                                            tmGrf(ilRecIndex).iAdfCode = tlComboType(ilLoopIndex).iSortCode

                                            tmGrf(ilRecIndex).iRdfCode = tlComboType(ilLoopIndex).lKeyCode          'internal code for named avail
                                            tmGrf(ilRecIndex).sDateType = ""                    'assume not an oversold avail name yet
                                            tmGrf(ilRecIndex).sGenDesc = "Missed"
                                            If ilListIndex = AVAILSCOMBO_SPORTS Then
                                                tmGrf(ilRecIndex).lChfCode = tmGsf.lCode
                                                tmGrf(ilRecIndex).lCode4 = tmGsf.lghfcode
                                                tmGrf(ilRecIndex).iSlfCode = tmSdf.iGameNo          'game number, if applicable
                                            End If
                                            If ilOtherMissedLoop = 2 Then
                                                'override some of the codes if its in the orphan pass (where no shown DP equals the DP of the missed spot)
                                                tmGrf(ilRecIndex).iRdfCode = -1     'sort it last tmRifSorts(ilRdf).isort   'DP Sort code  from RIF
                                                tmGrf(ilRecIndex).iSlfCode = tmSdf.iGameNo      '9-30-10
                                                tmGrf(ilRecIndex).iAdfCode = 32000   'DP code to retrieve DP name description
                                            End If
                                            ReDim Preserve tmGrf(0 To ilRecIndex + 1) As GRF
                                        End If

                                        If ilSpotOK Then
                                            llSpotAmount = mGetSpotPrice(llPropPrice)   'obtain actual spot price and proposal price
                                            'Date: 1/21/2020 added Net option
                                            If blGrossNet = False Then 'Net
                                                llSpotAmount = gGetGrossOrNetFromRate(llSpotAmount, "N", tmChf.iAgfCode)
                                                llPropPrice = gGetGrossOrNetFromRate(llPropPrice, "N", tmChf.iAgfCode)
                                            End If
                                            'tmGrf(ilRecIndex).lDollars(5) = tmGrf(ilRecIndex).lDollars(5) + llSpotAmount      'missed $
                                            'tmGrf(ilRecIndex).lDollars(14) = tmGrf(ilRecIndex).lDollars(14) + llPropPrice
                                            ''Count of spot sold in seconds to calc % sellout
                                            'tmGrf(ilRecIndex).lDollars(2) = tmGrf(ilRecIndex).lDollars(2) + tmSdf.iLen
                                            tmGrf(ilRecIndex).lDollars(4) = tmGrf(ilRecIndex).lDollars(4) + llSpotAmount      'missed $
                                            tmGrf(ilRecIndex).lDollars(13) = tmGrf(ilRecIndex).lDollars(13) + llPropPrice
                                            
                                            'Count of spot sold in seconds to calc % sellout
                                            tmGrf(ilRecIndex).lDollars(1) = tmGrf(ilRecIndex).lDollars(1) + tmSdf.iLen

                                            'accumulate 30" units.  Spots divisible by 30" = one unit each.  Anything less/equal to 30 is a unit
                                            ilUnitCount = tmSdf.iLen \ 30
                                            If (tmSdf.iLen Mod 30) > 0 Then
                                                ilUnitCount = ilUnitCount + 1
                                            End If
                                            'tmGrf(ilRecIndex).lDollars(7) = tmGrf(ilRecIndex).lDollars(7) + ilUnitCount

                                            ''decrement amount of seconds remaining
                                            'tmGrf(ilRecIndex).lDollars(4) = tmGrf(ilRecIndex).lDollars(4) - tmSdf.iLen
                                            ''decrement # of units remaining
                                            'tmGrf(ilRecIndex).lDollars(6) = tmGrf(ilRecIndex).lDollars(6) - 1
                                            tmGrf(ilRecIndex).lDollars(6) = tmGrf(ilRecIndex).lDollars(6) + ilUnitCount

                                            'decrement amount of seconds remaining
                                            tmGrf(ilRecIndex).lDollars(3) = tmGrf(ilRecIndex).lDollars(3) - tmSdf.iLen
                                            'decrement # of units remaining
                                            tmGrf(ilRecIndex).lDollars(5) = tmGrf(ilRecIndex).lDollars(5) - 1

                                            'if oversold, adjust the counts based on oversold spot lengths
                                            'tmGrf.lDollars(8)-(12) = spot counts of oversold, tmGrf.lDollars(13) = Other oversold count
                                            'If tmGrf(ilRecIndex).lDollars(4) < 0 Or tmGrf(ilRecIndex).lDollars(6) < 0 Then
                                            If tmGrf(ilRecIndex).lDollars(3) < 0 Or tmGrf(ilRecIndex).lDollars(5) < 0 Then
                                                tmGrf(ilRecIndex).sDateType = "O"           'flag as oversold avail name
                                                ilFound = False
                                                For illoop = 0 To imHowManyHL - 1
                                                    If tmSdf.iLen = tlCntTypes.iLenHL(illoop) Then
                                                        tmGrf(ilRecIndex).lDollars(illoop + 8 - 1) = tmGrf(ilRecIndex).lDollars(illoop + 8 - 1) - 1
                                                        ilFound = True
                                                        Exit For
                                                    End If
                                                Next illoop
                                                If Not ilFound Then             'no match found, put into the All Others count
                                                    'tmGrf(ilRecIndex).lDollars(13) = tmGrf(ilRecIndex).lDollars(13) + 1
                                                    tmGrf(ilRecIndex).lDollars(12) = tmGrf(ilRecIndex).lDollars(12) + 1
                                                    tmGrf(ilRecIndex).sGenDesc = "Missed"
                                                    tmGrf(ilRecIndex).lLong = 0                     'flag for later to determine if only unsold avails to be shown
                                                End If
                                            End If

                                            mAccumSpotsSoldByLen tlCntTypes, ilRecIndex           'accum sold by highlighted lengths
                                            Exit For                'force exit on this missed if found a matching daypart
                                        End If
                                    End If                      'ilSpotOK
                            'End If                          'ilAvailOK
                                End If
                                If ilFoundMissedDP = True And ilOtherMissedLoop = 2 Then        'put the missed spot in Other Missed category
                                    Exit For
                                End If
                            Next ilRdf
                            If ilFoundMissedDP = True And ilOtherMissedLoop = 1 Then        'if found matching DP; no need to look to place in other missed spots
                                Exit For
                            End If

                        Next ilOtherMissedLoop
                    End If                              'ilanfcode < 0 : missing dp from line
                End If              'dates within filter
                ilRet = btrGetNext(hmSdf, tmSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next ilPass
    End If

    Erase ilEvtType
    Erase tlLLC
End Sub
'
'
'
'               mFilterSpot - Test header and line exclusions for user request.
'
'               <input> ilVefCode - airing vehicle
'                       tlCntTypes - structure of inclusions/exclusions of contract types & status
'              <output> ilSpotOk - true if spot is OK, else false to ignore spot
'                       ilCTypes - set bit to 1 for the matching type
'                       ilSpotTypes - set bit to 1 for matching type
'
Sub mFilterSpot(ilVefCode As Integer, tlCntTypes As CNTTYPES, ilSpotOK As Integer, ilCTypes As Integer, ilSpotTypes As Integer)
Dim ilRet As Integer
Dim slPrice As String

    If ilSpotOK Then
        'Test header exclusions (types of contrcts and statuses)
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
            ilCTypes = &H8
            If Not tlCntTypes.iStandard Then       'include Standard types?
                ilSpotOK = False
            End If

        ElseIf tmChf.sType = "V" Then
            ilCTypes = &H10
            If Not tlCntTypes.iReserv Then      'include reservations ?
                ilSpotOK = False
            End If

        ElseIf tmChf.sType = "R" Then
            ilCTypes = &H20
            If Not tlCntTypes.iDR Then       'include DR?
                ilSpotOK = False
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
                ilSpotTypes = &H4
                If Not tlCntTypes.iZero Then
                    ilSpotOK = False
                End If
            Else
                ilSpotTypes = &H2
                If Not tlCntTypes.iCharge Then           'exclude charged spots
                    ilSpotOK = False
                End If
            End If
        ElseIf Trim$(slPrice) = "ADU" Then
            ilSpotTypes = &H8
            If Not tlCntTypes.iADU Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "Bonus" Then
            ilSpotTypes = &H10
            If Not tlCntTypes.iBonus Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "+ Fill" Then       '3-24-03
            ilSpotTypes = &H20
            If Not tlCntTypes.iXtra Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "- Fill" Then        '3-24-03
            ilSpotTypes = &H40
            If Not tlCntTypes.iFill Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "N/C" Then
            ilSpotTypes = &H80
            If Not tlCntTypes.iNC Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "Recapturable" Then
            ilSpotTypes = &H100
            If Not tlCntTypes.iRecapturable Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "Spinoff" Then
            ilSpotTypes = &H200
            If Not tlCntTypes.iSpinoff Then
                ilSpotOK = False
            End If
        ElseIf Trim$(slPrice) = "MG" Then               '10-28-10
            ilSpotTypes = &H400
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
'           mInitComboData - create an array of either dayparts to gather (if by non-sports)
'           or avail names to gather (if sports) which contains counts for inventory, sold,
'           avail, $, spots by date
'           <input> ilvefCode : vehicle code
'
Public Sub mInitComboData(ilVefCode As Integer, tlCntTypes As CNTTYPES)
Dim ilRcf As Integer
Dim ilSelected As Integer
Dim illoop As Integer
Dim slNameCode As String
Dim ilRet As Integer
Dim slCode As String
Dim llRif As Long
Dim ilRdf As Integer
Dim slSaveReport As String
Dim ilSaveSort As Integer
Dim ilFound As Integer
Dim ilUpper As Integer
Dim ilAvailNameLoop As Integer
Dim ilDayIndex As Integer
ReDim tmComboType(0 To 0) As COMBOTYPE
Dim ilLoopOnListBox As Integer
Dim slName As String
Dim ilAnfCode As Integer
Dim ilMissed As Integer

        ReDim tmGrf(0 To 0) As GRF
        ilUpper = 0
        If RptSelCA!lbcRptType.ListIndex = AVAILSCOMBO_SPORTS Then      'sports avails, build key info for all named avails
            For ilLoopOnListBox = 0 To RptSelCA!lbcSelection(2).ListCount - 1
                'TTP 10434 - Event and Sports export (WWO)
                'If RptSelCA!lbcSelection(2).Selected(ilLoopOnListBox) Then                'selected by user
                If RptSelCA!lbcSelection(2).Selected(ilLoopOnListBox) Or RptSelCA.rbcOutput(3).Value = True Then              'selected by user
                    If UBound(tmComboData) = 0 Then                 'if upper limit of table is 0, it hasn't been built yet
                        ilMissed = False                            'not creating the missed entry for a named avail
                        slNameCode = tgNamedAvail(ilLoopOnListBox).sKey
                        ilRet = gParseItem(slNameCode, 1, "\", slName)
                        ilRet = gParseItem(slName, 3, "|", slName)
                        ilRet = gParseItem(slNameCode, 2, "\", slCode)
                        ilAnfCode = Val(slCode)
                        For ilAvailNameLoop = LBound(tgAvailAnf) To UBound(tgAvailAnf) - 1
                            If tgAvailAnf(ilAvailNameLoop).iCode = ilAnfCode Then
                                mSetNamedAvailValuesInComboData ilUpper, ilAvailNameLoop, ilMissed
                            End If
                        Next ilAvailNameLoop
                    End If                              'tmcombodata = 0
                End If
            Next ilLoopOnListBox
            '9-30-10 extra entry was created in wrong place, creating too many
            'completed building all named avails; create an extra one without any named avail reference for Missed spots bucket
            If tlCntTypes.iMissed Then          'if including missed, create the extra hard-coded Missed entry for dayparts
                                                'that dont have any named avail reference
                ilMissed = True                 'set up the extra entry for Missed spots that dont have named avails reference
                mSetNamedAvailValuesInComboData ilUpper, 0, ilMissed
            End If
        Else
            ReDim tmComboType(0 To 0) As COMBOTYPE                  'intialized for each vehicle because each vehicle has different dayparts
            For ilRcf = LBound(tgMRcf) To UBound(tgMRcf) - 1 Step 1     'non-sports avails, build key info for the selected r/c
                tmRcf = tgMRcf(ilRcf)
                ilSelected = False
                For illoop = 0 To RptSelCA!lbcSelection(1).ListCount - 1 Step 1
                    slNameCode = tgRateCardCode(illoop).sKey
                    ilRet = gParseItem(slNameCode, 3, "\", slCode)
                    If Val(slCode) = tgMRcf(ilRcf).iCode Then
                        If (RptSelCA!lbcSelection(1).Selected(illoop)) Then
                            ilSelected = True
                        End If
                        Exit For
                    End If
                Next illoop

                If ilSelected Then              'found the R/c

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
                                    For illoop = LBound(tmComboType) To ilUpper - 1 Step 1
                                        If tmComboType(illoop).lKeyCode = tgMRdf(ilRdf).iCode Then
                                            ilFound = True
                                            Exit For
                                        End If
                                    Next illoop
                                    If Not ilFound Then
                                        tmComboType(ilUpper).iType = 1          'dp
                                        tmComboType(ilUpper).lKeyCode = tgMRdf(ilRdf).iCode     'internal dp code
                                        tmComboType(ilUpper).iSortCode = ilSaveSort
                                        For illoop = LBound(tgMRdf(ilRdf).iStartTime, 2) To UBound(tgMRdf(ilRdf).iStartTime, 2) Step 1 'Row
                                            tmComboType(ilUpper).iStartTime(0, illoop) = tgMRdf(ilRdf).iStartTime(0, illoop)
                                            tmComboType(ilUpper).iStartTime(1, illoop) = tgMRdf(ilRdf).iStartTime(1, illoop)
                                            tmComboType(ilUpper).iEndTime(0, illoop) = tgMRdf(ilRdf).iEndTime(0, illoop)
                                            tmComboType(ilUpper).iEndTime(1, illoop) = tgMRdf(ilRdf).iEndTime(1, illoop)
                                            For ilDayIndex = 1 To 7
                                                'tmComboType(ilUpper).sWkDays(ilLoop, ilDayIndex) = tgMRdf(ilRdf).sWkDays(ilLoop, ilDayIndex)
                                                tmComboType(ilUpper).sWkDays(illoop, ilDayIndex - 1) = tgMRdf(ilRdf).sWkDays(illoop, ilDayIndex - 1)
                                            Next ilDayIndex
                                        Next illoop
                                        tmComboType(ilUpper).ianfCode = tgMRdf(ilRdf).ianfCode
                                        tmComboType(ilUpper).sInOut = tgMRdf(ilRdf).sInOut

                                        ilUpper = ilUpper + 1
                                        ReDim Preserve tmComboType(0 To ilUpper) As COMBOTYPE
                                    End If  'not ilfound
                                End If      'tgMRdf(ilRdf).iCode = tgMRif(llRif).iRdfcode And slSaveReport = "Y" And tgMRdf(ilRdf).sState <> "D" And tgMRif(llRif).iVefCode = ilVefCode
                            End If          'ilRdf <> -1
                        End If              'tgMRif(llRif).iRcfCode = tgMRcf(ilRcf).iCode
                    Next llRif
                End If                      'ilSelected
            Next ilRcf
        End If
        Exit Sub
End Sub
'
'       mGetSpotPrice - get spot rate for Combo Avails report
'       <output>  llPropPrice - flight proposal price ( from cffpropprice)
'       <return> Spot rate
'
Private Function mGetSpotPrice(llPropPrice As Long) As Long
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  slSpotAmount                  llSpotAmount                                            *
'******************************************************************************************

Dim ilRet As Integer
Dim llActualPrice As Long

        'slSpotAmount = ".00"
        llActualPrice = 0
        llPropPrice = 0
        If tmSdf.sSpotType <> "X" And tmSdf.sSpotType <> "O" And tmSdf.sSpotType <> "C" Then                  'bonus?
            'need to get flight for proposal price and actual price
            ilRet = gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tmCff)

            llActualPrice = tmCff.lActPrice
            llPropPrice = tmCff.lPropPrice * 100    'price is in whole $
            'ilRet = gGetSpotPrice(tmSdf, tmClf, hmCff, hmSmf, hmVef, hmVsf, slSpotAmount)
'            If InStr(slSpotAmount, ".") = 0 Then        'didnot find decimal point, its NC spino
'                slSpotAmount = ".00"
'            End If
            If (tmSdf.sPriceType = "N") Then
                'slSpotAmount = ".00"
                llActualPrice = 0
            ElseIf (tmSdf.sPriceType = "P") Then
                'slSpotAmount = ".00"
                llActualPrice = 0
            End If
        End If
        'llSpotAmount = gStrDecToLong(slSpotAmount, 2)
        'mGetSpotPrice = llSpotAmount
        mGetSpotPrice = llActualPrice
End Function

'
'           mAccumSpotsSoldByLen
'           <input>  tlCntTypes - record of selected parameters
'                    ilRecIndex - index into the array of info to report on
Private Sub mAccumSpotsSoldByLen(tlCntTypes As CNTTYPES, ilRecIndex As Integer)
    Dim illoop As Integer
    Dim ilFound As Integer

        'accumulate sold by highlighted spot lengths
        'if spot length not found, it is placed into All Others bucket
        'GrfPerGenl(1-5) = user highlighted lengths sold counts
        'GrfPergenl(6) = All others when no match found

        ilFound = False
        For illoop = 0 To imHowManyHL - 1
            If tmSdf.iLen = tlCntTypes.iLenHL(illoop) Then
                'tmGrf(ilRecIndex).iPerGenl(ilLoop + 1) = tmGrf(ilRecIndex).iPerGenl(ilLoop + 1) + 1
                tmGrf(ilRecIndex).iPerGenl(illoop) = tmGrf(ilRecIndex).iPerGenl(illoop) + 1
                ilFound = True
                Exit For
            End If
        Next illoop
        If Not ilFound Then             'no match found, put into the All Others count
            'tmGrf(ilRecIndex).iPerGenl(6) = tmGrf(ilRecIndex).iPerGenl(6) + 1
            tmGrf(ilRecIndex).iPerGenl(5) = tmGrf(ilRecIndex).iPerGenl(5) + 1
        End If
        Exit Sub
End Sub

'
'               mGetAvailsRemaining
'               <input> tlCntTypes - record of selected parameters
'                       ilRecIndex index into the grf array of entry to process
'
'               1-24-12 adjust the sold and avail columns if equalizing 30/60 counts
'
Private Sub mGetAvailsRemaining(tlCntTypes As CNTTYPES, ilRecIndex As Integer)
    Dim illoop As Integer
    Dim ilFound As Integer
    Dim ilRemLen As Integer
    Dim ilRemUnits As Integer
    Dim ilOrigRemLen As Integer
    Dim ilOrigRemUnits As Integer
    Dim flTruncate As Single
    Dim flTempAvail As Single
    Dim flTempSold As Single
    Dim llTemp As Long

    'determine how much of each highlighted spot length remains by named avail
    'decide based on the highest spot length to the lowest
    'GrfPerGenl(13-17) Avail counts for spotlengths to highlight
    
    '1-24-12 if using site AvailsEqualize 30 or 60, grfPerGenl(17) will be used to accumulate the equalize value
    'GrfPerGenl(18) = avail counts for All Others when no matching length found
    ilRemLen = tmGrf(ilRecIndex).lDollars(3)        'get remaining seconds for the named avail
    ilRemUnits = tmGrf(ilRecIndex).lDollars(5)      'get remaining units for the named avail
    ilFound = False

    'check if already oversold, adjust counts to show minus for avail counts
    If ilRemLen < 0 Or ilRemUnits < 0 Then              '1-30-15 changed from AND to OR
        If tmGrf(ilRecIndex).sDateType = "O" Then       'avail name already found to be oversold?
            For illoop = 0 To imHowManyHL - 1
                tmGrf(ilRecIndex).iPerGenl(illoop + 12) = tmGrf(ilRecIndex).lDollars(7 + illoop)
            Next illoop
        Else
            '8-18-10 avoid looping when a spot length isnt found to highlight
            ilOrigRemLen = ilRemLen
            ilOrigRemUnits = ilRemUnits
            Do While ilRemLen < 0 And ilRemUnits < 0            'units and sec must remain to show any availability
                For illoop = 0 To imHowManyHL - 1
                    Do While ilRemLen < 0 And ilRemUnits < 0
                        If -(ilRemLen) >= tlCntTypes.iLenHL(illoop) Then
                              ilRemLen = ilRemLen + tlCntTypes.iLenHL(illoop)
                              ilRemUnits = ilRemUnits + 1
                              'accum count of avails for the highlighted spot length
                              'tmGrf(ilRecIndex).iPerGenl(ilLoop + 13) = tmGrf(ilRecIndex).iPerGenl(ilLoop + 13) - 1
                              tmGrf(ilRecIndex).iPerGenl(illoop + 12) = tmGrf(ilRecIndex).iPerGenl(illoop + 12) - 1
                              ilFound = True
                              'Exit For
                        Else
                            Exit Do
                        End If
                    Loop
                Next illoop
                If ilFound Then
                    Exit Do
                Else
                    '8-18-10 avoid looping when no spot lengths found
                    If ilOrigRemLen = ilRemLen And ilOrigRemUnits = ilRemUnits Then
                        Exit Do
                    End If
                End If
            Loop

        End If
    Else
        Do While ilRemLen > 0 And ilRemUnits > 0            'units and sec must remain to show any availability
            For illoop = 0 To imHowManyHL - 1
                Do While ilRemLen > 0 And ilRemUnits > 0
                    If ilRemLen >= tlCntTypes.iLenHL(illoop) Then
                          ilRemLen = ilRemLen - tlCntTypes.iLenHL(illoop)
                          ilRemUnits = ilRemUnits - 1
                          'accum count of avails for the highlighted spot length
                          'tmGrf(ilRecIndex).iPerGenl(ilLoop + 13) = tmGrf(ilRecIndex).iPerGenl(ilLoop + 13) + 1
                          tmGrf(ilRecIndex).iPerGenl(illoop + 12) = tmGrf(ilRecIndex).iPerGenl(illoop + 12) + 1
                          ilFound = True
                          'Exit For
                    Else
                        ilFound = True
                        Exit Do
                    End If
                Loop
            Next illoop
            If ilFound Then
                Exit Do
            Else                    'not found, are there any more lengths to test?
                If illoop <= imHowManyHL Then
                    Exit Do
                End If
            End If
        Loop
        If ilRemLen > 0 And ilRemUnits > 0 Then
            'accum another unit into All Others.  User should be highlighting valid selling spot lengths so
            'nothing falls into All Others
            'tmGrf(ilRecIndex).iPerGenl(18) = tmGrf(ilRecIndex).iPerGenl(18) + 1
            tmGrf(ilRecIndex).iPerGenl(17) = tmGrf(ilRecIndex).iPerGenl(17) + 1
        End If
    End If
    
    
    '1-24-12 equalize avails remaining to 30/60 if applicable.
    'Orig method was to take the count of each spot length and determine how the equalized counts (based on a ratio).
    'Change that method to add up all the spot length counts and divide by 30 or 60.  Ignore OTH counts
    If imUsingEqualize Then
        flTempAvail = 0
        flTempSold = 0
        For illoop = 0 To imHowManyHL - 1
            'equalize avail (grfPerGenl(13-17) & sold counts (grfPerGenl(5-9), grfPerGenl(7-11) = spot lengths to highlight
            'flTempAvail = flTempAvail + (tmGrf(ilRecIndex).iPerGenl(ilLoop + 13) * tmGrf(ilRecIndex).iPerGenl(ilLoop + 7))
            'flTempSold = flTempSold + (tmGrf(ilRecIndex).iPerGenl(ilLoop + 1) * tmGrf(ilRecIndex).iPerGenl(ilLoop + 7))
            flTempAvail = flTempAvail + (tmGrf(ilRecIndex).iPerGenl(illoop + 12) * tmGrf(ilRecIndex).iPerGenl(illoop + 6))
            flTempSold = flTempSold + (tmGrf(ilRecIndex).iPerGenl(illoop) * tmGrf(ilRecIndex).iPerGenl(illoop + 6))
'                flTruncate = tmGrf(ilRecIndex).iPerGenl(ilLoop + 13) / fmLenRatio(ilLoop + 1)
'                llTemp = flTruncate * 10        'remove decimal
'                tmGrf(ilRecIndex).iPerGenl(17) = tmGrf(ilRecIndex).iPerGenl(17) + (llTemp \ 10) 'CInt(flTruncate)
'
'                'equalize the counts sold
'                flTruncate = tmGrf(ilRecIndex).iPerGenl(ilLoop + 1) / fmLenRatio(ilLoop + 1)        'when round, vb rounds on odd number, and truncates on even number.  i.e. 1.5 rounds to 2, 2.5 truncates to 2
'                llTemp = flTruncate * 10                'remove decimal
'                tmGrf(ilRecIndex).iPerGenl(5) = tmGrf(ilRecIndex).iPerGenl(5) + (llTemp \ 10) 'CInt(flTruncate)
        Next illoop
        llTemp = (flTempAvail * 10) / imEqualizeOption          'remove decimal
        'tmGrf(ilRecIndex).iPerGenl(17) = llTemp '(llTemp \ 10)
        tmGrf(ilRecIndex).iPerGenl(16) = llTemp '(llTemp \ 10)
        llTemp = flTempSold * 10 / imEqualizeOption          'remove decimal
        'tmGrf(ilRecIndex).iPerGenl(5) = llTemp  '(llTemp \ 10)
        tmGrf(ilRecIndex).iPerGenl(4) = llTemp  '(llTemp \ 10)
       
    End If

    Exit Sub
End Sub
'
'           setup values for the selected Name avails as well as an
'           mSetNamedAvailValuesInComboData
'               <input> ilUpper - index into entry to setup
'                       ilNamedAvailLoop - entry into the Named Avail table to process
'                       ilMissed - Setup missed entry (true/false)
'           extra one for Missed spots.  When processing missed spots,
'           the named avail is obtained from the schedule lines daypart.
'           The daypart may be "All Avails" in which case there is no
'           named avails defined
'
Public Sub mSetNamedAvailValuesInComboData(ilUpper As Integer, ilAvailNameLoop As Integer, ilMissed As Integer)
Dim illoop As Integer
Dim ilDayIndex As Integer

        tmComboType(ilUpper).iType = 0                      'named avails array flag
        If ilMissed Then                                    'named avails entry for Missed
            tmComboType(ilUpper).lKeyCode = 0          'internal code will still be 0 for matching missed spots without named avails references
            tmComboType(ilUpper).iSortCode = 32000                                      'make these missed spots last in sort
            tmComboType(ilUpper).ianfCode = 0
        Else
            tmComboType(ilUpper).lKeyCode = tgAvailAnf(ilAvailNameLoop).iCode
            tmComboType(ilUpper).iSortCode = tgAvailAnf(ilAvailNameLoop).iSortCode
            tmComboType(ilUpper).ianfCode = tgAvailAnf(ilAvailNameLoop).iCode
        End If
        tmComboType(ilUpper).sInOut = "A"                   'flag to test all avails
        For illoop = 1 To 7
            If illoop = 7 Then          'default start/end times to 12m-12m
                tmComboType(ilUpper).iStartTime(0, 6) = 0
                tmComboType(ilUpper).iStartTime(1, 6) = 0
                tmComboType(ilUpper).iEndTime(0, 6) = 0
                tmComboType(ilUpper).iEndTime(1, 6) = 0
            Else                                'first 6 times do not apply
                tmComboType(ilUpper).iStartTime(0, illoop - 1) = 1
                tmComboType(ilUpper).iStartTime(1, illoop - 1) = 0
                tmComboType(ilUpper).iEndTime(0, illoop - 1) = 1
                tmComboType(ilUpper).iEndTime(1, illoop - 1) = 0
            End If
            For ilDayIndex = 1 To 7
                'tmComboType(ilUpper).sWkDays(ilLoop, ilDayIndex) = "Y"          'default all days to Y
                tmComboType(ilUpper).sWkDays(illoop - 1, ilDayIndex - 1) = "Y"      'default all days to Y
            Next ilDayIndex
        Next illoop
        ilUpper = ilUpper + 1
        ReDim Preserve tmComboType(0 To ilUpper) As COMBOTYPE
    Exit Sub
End Sub
'
'       mGenMMAvails - generate Multimedia avails for  sports for the
'                      Game Avails report
'                      Get the inventory belonging to the game and the
'                      multimedia sold for the game
'                      Create a grf record with unique type code to be used
'                      in a subreport called from GameAvail.rpt
'               <input> ilvefcode - vehicle code
'
'               Return in global: imGameIndFlag - flag to indicate if the game independent info has been
'                                       obtained for the vehicle.  this should only be done
'                                       once per vehicle
'       Assumptions:  Gsf (game schedule ) in memory
Public Sub mGenMMAvailsForCombo(ilVefCode As Integer)
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilVefInx                                                                              *
'******************************************************************************************

Dim ilLoopOnMM As Integer
Dim ilLoopOnSoldMM As Integer
Dim ilRet As Integer
Dim ilGameNo As Integer
Dim ilFound As Integer
Dim ilLoopOnMMInfo As Integer
Dim ilUpperMMInfo As Integer
ReDim tlMMInfo(0 To 0) As MMINFO
Dim ilGameInd As Integer
Dim tlGrf As GRF

        tmGhfSrchKey.lCode = tmGsf.lghfcode
        ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet <> BTRV_ERR_NONE Then          'no game applicable, ignore this entry
            Exit Sub
        End If

        ilUpperMMInfo = 0
        'Build Inventory
        'if game independent multimedia has already been obtained for this vehicle, ignore it.
        'only do game independent once per vehicle to prevent overstating inventory/sold
        If imGameIndFlag = True Then            'only retrieved the game independent inventory/sold
            ilGameInd = 1
        Else
            ilGameInd = 2
        End If
        For ilLoopOnMM = 1 To ilGameInd         'dependent vs independent
            If ilLoopOnMM = 1 Then      'dependent Multimedia
                ilGameNo = tmGsf.iGameNo
            Else
                ilGameNo = 0            'game independent multimedia doesnt have game # association
                imGameIndFlag = True
            End If

            tmIsfSrchKey1.iGameNo = ilGameNo
            tmIsfSrchKey1.lghfcode = tmGhf.lCode    'game header
            ilRet = btrGetEqual(hmIsf, tmIsf, imIsfRecLen, tmIsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
            Do While ilGameNo = tmIsf.iGameNo And tmGhf.lCode = tmIsf.lghfcode And ilRet = BTRV_ERR_NONE
                If tmIsf.iNoUnits <> 0 Then
                    ilFound = False
                    For ilLoopOnMMInfo = 0 To UBound(tlMMInfo) - 1
                        If tlMMInfo(ilUpperMMInfo).iIhfCode = tmIsf.iIhfCode And tlMMInfo(ilUpperMMInfo).iGameNo = ilGameNo Then
                            'Inventory: accumulate the $ and units
                            tlMMInfo(ilLoopOnMMInfo).lRate = tlMMInfo(ilLoopOnMMInfo).lRate + tmIsf.lRate
                            tlMMInfo(ilLoopOnMMInfo).lNoUnits = tlMMInfo(ilLoopOnMMInfo).lNoUnits + tmIsf.iNoUnits
                            ilFound = True
                            Exit For
                        End If
                    Next ilLoopOnMMInfo

                    If Not ilFound Then
                        'new inventory
                        tlMMInfo(ilUpperMMInfo).iIhfCode = tmIsf.iIhfCode       'inventory header code
                        tlMMInfo(ilLoopOnMMInfo).lRate = tmIsf.lRate
                        tlMMInfo(ilLoopOnMMInfo).lNoUnits = tmIsf.iNoUnits
                        tlMMInfo(ilLoopOnMMInfo).iGameNo = ilGameNo         'game #, 0 = game independent
                        ilUpperMMInfo = ilUpperMMInfo + 1
                        ReDim Preserve tlMMInfo(0 To ilUpperMMInfo) As MMINFO
                    End If
                End If
                ilRet = btrGetNext(hmIsf, tmIsf, imIsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            Loop
        Next ilLoopOnMM

        'Build Sold; retrieve all multimedia sold based on the game header reference
        tmMsfSrchKey1.iVefCode = ilVefCode
        ilRet = btrGetEqual(hmMsf, tmMsf, imMsfRecLen, tmMsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        'get all the sold inventory for the matching vehicle
        Do While ilVefCode = tmMsf.iVefCode And ilRet = BTRV_ERR_NONE
            'obtain the associated contract to see if this is a history sold record
            tmChfSrchKey.lCode = tmMsf.lChfCode
            ilRet = btrGetEqual(hmCHF, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            If tmMsf.lghfcode = tmGhf.lCode And ilRet = BTRV_ERR_NONE And tmChf.sDelete = "N" Then        'need to match on the game header code
                For ilLoopOnSoldMM = 1 To ilGameInd                 '1 pass for game dependent multimedia sold, 1 pass for game independent
                    If ilLoopOnSoldMM = 1 Then
                        ilGameNo = tmGsf.iGameNo
                    Else
                        ilGameNo = 0                    'game independent
                    End If

                    tmMgfSrchKey1.iGameNo = ilGameNo
                    tmMgfSrchKey1.lMsfCode = tmMsf.lCode
                    ilRet = btrGetEqual(hmMgf, tmMgf, imMgfRecLen, tmMgfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
                    Do While tmMgf.iGameNo = ilGameNo And tmMgf.lMsfCode = tmMsf.lCode And ilRet = BTRV_ERR_NONE
                        For ilLoopOnMM = 0 To UBound(tlMMInfo) - 1
                            If tlMMInfo(ilLoopOnMM).iIhfCode = tmMsf.iIhfCode Then
                                'accumulate the sold for this piece of inventory
                                tlMMInfo(ilLoopOnMM).lRateSold = tlMMInfo(ilLoopOnMM).lRateSold + tmMgf.lRate
                                tlMMInfo(ilLoopOnMM).lNoUnitsSold = tlMMInfo(ilLoopOnMM).lNoUnitsSold + tmMgf.iNoUnits
                                Exit For
                            End If
                        Next ilLoopOnMM
                    ilRet = btrGetNext(hmMgf, tmMgf, imMgfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                Next ilLoopOnSoldMM
            End If
            ilRet = btrGetNext(hmMsf, tmMsf, imMsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop


        'Create the prepass records for Crystal
        For ilLoopOnMM = 0 To UBound(tlMMInfo) - 1
            tlGrf.iGenDate(0) = igNowDate(0)
            tlGrf.iGenDate(1) = igNowDate(1)
            gUnpackTimeLong igNowTime(0), igNowTime(1), False, lgNowTime
            tlGrf.lGenTime = lgNowTime
            tlGrf.iVefCode = ilVefCode
            tlGrf.iSlfCode = tlMMInfo(ilLoopOnMM).iGameNo          'game number, if applicable
            If tlMMInfo(ilLoopOnMM).iGameNo = 0 Then                'game independent
                tlGrf.sBktType = "I"
            Else
                tlGrf.sBktType = "D"            'game dependent
            End If
            tlGrf.lChfCode = tmGsf.lCode        'Game schedule
            tlGrf.lCode4 = tmGsf.lghfcode       'game header
            'tlGrf.iPerGenl(12) = tlMMInfo(ilLoopOnMM).iIhfCode  'Inventory header code
            tlGrf.iPerGenl(11) = tlMMInfo(ilLoopOnMM).iIhfCode  'Inventory header code
            'tlGrf.lDollars(1) = tlMMInfo(ilLoopOnMM).lNoUnits   'total inventory
            'tlGrf.lDollars(7) = tlMMInfo(ilLoopOnMM).lNoUnitsSold    'total inventory sold
            'tlGrf.lDollars(5) = tlMMInfo(ilLoopOnMM).lRateSold       'sold $
            tlGrf.lDollars(0) = tlMMInfo(ilLoopOnMM).lNoUnits   'total inventory
            tlGrf.lDollars(6) = tlMMInfo(ilLoopOnMM).lNoUnitsSold    'total inventory sold
            tlGrf.lDollars(4) = tlMMInfo(ilLoopOnMM).lRateSold       'sold $
            If RptSelCA!ckcUnsoldOnly(1).Value = vbChecked And RptSelCA.rbcOutput(3).Value = False Then          'show unsold only
                'test for anything available
                If tlMMInfo(ilLoopOnMM).lNoUnits - tlMMInfo(ilLoopOnMM).lNoUnitsSold > 0 Then
                    ilRet = btrInsert(hmGrf, tlGrf, imGrfRecLen, INDEXKEY0)
                End If
            Else                                                    'show everything
                ilRet = btrInsert(hmGrf, tlGrf, imGrfRecLen, INDEXKEY0)
            End If
        Next ilLoopOnMM
    Erase tlMMInfo
    Exit Sub
End Sub
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

Private Function LookupGSF(lghfcode As Long, iVehicleCode As Integer, iGameNo As Integer) As GSF
    Dim slSQLQuery As String
    Dim rst_Temp As ADODB.Recordset
    On Error GoTo LookupGSFError
    
    slSQLQuery = "SELECT * FROM GSF_Game_Schd "
    slSQLQuery = slSQLQuery & " WHERE gsfGhfCode = " & lghfcode
    slSQLQuery = slSQLQuery & " AND gsfVefCode = " & iVehicleCode
    slSQLQuery = slSQLQuery & " AND gsfGameNo = " & iGameNo
    
    LookupGSF.lCode = -1
    
    Set rst_Temp = gSQLSelectCall(slSQLQuery)
    If Not rst_Temp.EOF Then
        gPackDate rst_Temp!gsfAirDate, LookupGSF.iAirDate(0), LookupGSF.iAirDate(1)
        gPackTime Trim$(rst_Temp!gsfAirTime), LookupGSF.iAirTime(0), LookupGSF.iAirTime(1)
        LookupGSF.iAirVefCode = rst_Temp!gsfAirVefCode
        LookupGSF.iGameNo = rst_Temp!gsfGameNo
        LookupGSF.iHomeMnfCode = rst_Temp!gsfHomeMnfCode
        LookupGSF.iLangMnfCode = rst_Temp!gsfLangMnfCode
        LookupGSF.iSubtotal1MnfCode = rst_Temp!gsfSubtotal1MnfCode
        LookupGSF.iSubtotal2MnfCode = rst_Temp!gsfSubtotal2MnfCode
        LookupGSF.iVefCode = rst_Temp!gsfVefCode
        LookupGSF.iVisitMnfCode = rst_Temp!gsfVisitMnfCode
        LookupGSF.lCode = rst_Temp!gsfCode
        LookupGSF.lghfcode = rst_Temp!gsfGhfCode
        LookupGSF.lLvfCode = rst_Temp!gsfLvfCode
        LookupGSF.sBus = Trim$(rst_Temp!gsfBus)
        LookupGSF.sFeedSource = Trim$(rst_Temp!gsfFeedSource)
        LookupGSF.sGameStatus = Trim$(rst_Temp!gsfGameStatus)
        LookupGSF.sLiveLogMerge = Trim$(rst_Temp!gsfLiveLogMerge)
        LookupGSF.sXDSProgCodeID = Trim$(rst_Temp!gsfXDSProgCodeID)
    End If
    rst_Temp.Close
    Exit Function
    
LookupGSFError:
    rst_Temp.Close
    LookupGSF.lCode = -1
End Function

Private Function mGetMnfName(ilMnfCode As Integer) As String
    Dim slSQLQuery As String
    Dim tmp_rst As ADODB.Recordset
    slSQLQuery = "Select mnfName from MNF_Multi_Names where mnfCode = " & ilMnfCode
    Set tmp_rst = gSQLSelectCall(slSQLQuery)
    If Not tmp_rst.EOF Then
        mGetMnfName = Trim$(tmp_rst!mnfName)
    Else
        mGetMnfName = ""
    End If
End Function

Private Function mGetRDFFromAnfCode(ilAnf As Integer) As Integer
    mGetRDFFromAnfCode = -1
    
    Dim slSQLQuery As String
    Dim tmp_rst As ADODB.Recordset
    slSQLQuery = "Select rdfCode from RDF_Standard_Daypart where rdfanfCode = " & ilAnf
    Set tmp_rst = gSQLSelectCall(slSQLQuery)
    If Not tmp_rst.EOF Then
        mGetRDFFromAnfCode = Trim$(tmp_rst!rdfCode)
    End If
End Function

Private Function gGetDefaultBook(ilVef As Integer) As Integer
    gGetDefaultBook = -1
    
    Dim slSQLQuery As String
    Dim tmp_rst As ADODB.Recordset
    slSQLQuery = "Select vefdnfCode from VEF_Vehicles where vefCode = " & ilVef
    Set tmp_rst = gSQLSelectCall(slSQLQuery)
    If Not tmp_rst.EOF Then
        gGetDefaultBook = Trim$(tmp_rst!vefdnfCode)
    End If
End Function
