Attribute VB_Name = "modRptSubs"

'
'       gOutputMethod -Interrogate Output method selected and setup
'       export if required
'       <Input>  Form
'                CrystalReport Name
'       <output> None
'

Option Explicit
Private cprst As ADODB.Recordset
Private Advrst As ADODB.Recordset           '11-30-12
Private tmAltForAST As ADODB.Recordset       'alt
Private tmCPDat() As DAT
Private tmAstInfo() As ASTINFO
Private tmASTForMGRepl As ADODB.Recordset       '7-12-13
Private tmAdfRst As ADODB.Recordset

Private tmAstInfoSort() As ASTSORTKEY       '2-28-13
Private rst_rht As ADODB.Recordset
Private rst_ret As ADODB.Recordset
Private rst_Mnf As ADODB.Recordset         '5-31-16

Public igRptSource As Integer
Public igRptIndex As Integer
Public sgGenDate As String          '7-10-13 make the prepass gen date and time global for all reports to use
Public sgGenTime As String

Public lgRptTtlTime1 As Long
Public lgRptSTime1 As Long
Public lgRptETime1 As Long

Public lgCPCount As Long
Public lgSpotCount As Long
Public lgSpotCount2 As Long

Type VEHICLETEXTINFO
    iVefCode As Long
    lCode As Long
    sType As String * 1         'Header or footer
End Type

Public tgVtfInfo() As VEHICLETEXTINFO

Type SEASONINFO
    lStartDate As Long
    lEndDate As Long
    sName As String * 20
    lGhfCode As Long
End Type
Public tgSeasonInfo() As SEASONINFO

Public tgRadarHdrInfo() As RADAR_HDRINFO
Public tgRadarDetailInfo() As RADAR_DETAILINFO

'Report name equivalences
Dim hmAfr As Integer
Dim tmAfr As AFR
Dim imAfrRecLen As Integer
Dim tmAfrKey As AFRKEY0
Dim tmAssignInfo As ASSIGNINFO
Private tmCrfRst As ADODB.Recordset
Private tmTzfRst As ADODB.Recordset
Private tmCatRst As ADODB.Recordset

Type RADAR_HDRINFO
    iVefCode As Integer     'vehicle code
    lrhtCode As Long        'RHT internal code
    sNetworkCode As String * 2  'network code to show on report
    sVehicleCode As String * 3  'more than one vehicle & network code can exist, vehicle code keeps them apart (this is a radar vehicle code)
    lStartInx As Long           'start index into Radar detail info array for associated vehicle/network code
    lEndInx As Long             'end index into radar detail info array for associated vehicle/network code
End Type

'RADAR_HDRINFO array points to RADAR_DETAILINFO array
Type RADAR_DETAILINFO
    lrhtCode As Long
    lStartTime As Long          'start time of event
    lEndTime As Long            'end time of event
    sDayType As String * 2      'valid airing days
End Type

Type STATUSOPTIONS
'    iInclUnresolveMissed As Boolean
    iInclResolveMissed As Boolean             'is this needed to determine whether to include the missed side of a mg or replacment spot
'    iInclReplacement As Boolean
'    iInclBonus As Boolean
    'All Status codes 0-13
    iInclLive0 As Boolean
    iInclDelay1 As Boolean
    iInclMissed2 As Boolean
    iInclMissed3 As Boolean
    iInclMissed4 As Boolean
    iInclMissed5 As Boolean
    iInclAirOutPledge6 As Boolean
    iInclAiredNotPledge7 As Boolean
    iInclNotCarry8 As Boolean
    iInclDelayCmmlOnly9 As Boolean
    iInclAirCmmlOnly10 As Boolean
    iInclMG11 As Boolean
    iInclBonus12 As Boolean
    iInclRepl13 As Boolean
    iNotReported As Boolean
    bStatusDiscrep As Boolean           '12-10-13 option to see status discrepancy only; implemented question in Fed vs Aired with new .rpt created
    iCompBarter As Integer
    iCompPayStation As Integer
    iCompPayNetwork As Integer
    iInclMissedMGBypass14 As Boolean        '4-12-17
    iInclCopyChanges As Boolean         'Date: 2020/3/23 flag to include/exclude copy changes
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
    lCrfCsfCode           As Long            ' regional copy script or comment
    lAttCode              As Long            ' Agreement code
    sCompliant            As String * 1      ' Spot compliant (Y or N)
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
    sCallLetters          As String * 10     ' Original Call Letters
    iPledgeDate(0 To 1)   As Integer         ' Pledge date
    iPledgeStartTime(0 To 1) As Integer      ' Pledge start time
    iPledgeEndTime(0 To 1) As Integer        ' Pledge end time
    iPledgeStatus         As Integer         ' Pledge status
    sSplitNet              As String * 1     ' Split Network (blank = no , p = primary, s = secondary, test for P or S
    lID                    As Long           ' line #
    iLen                   As Integer        ' spot length
    sSpotType              As String * 1     ' spot id (0 = sch; 1-mg;2-filled,3-outside, 4=?, 5 = added; 6 = open bb, 7=closebb
    iAirPlayNo             As Integer        ' air play
    sUnused               As String * 10      ' Unused


End Type


Type AFRKEY0
    iGenDate(0 To 1)      As Integer
    lGenTime              As Long
End Type

Type ASSIGNINFO
        iGenDate(0 To 1) As Integer  'generation date for filter
        lGenTime As Long             ' generation time for filter
        sGenDate As String
        sGenTime As String
        iDate(0 To 1) As Integer     ' date of spot
        sAirDate As String
        iTime(0 To 1) As Integer     ' time of spot
        sAirTime As String
        iAdfCode As Integer          'advertiser code
        iCode2 As Integer            ' spot length
        lCrfCode As Long             ' CRF (rotation header) internal code
        lChfCode As Long             ' internal contract code
        iLine As Integer             ' line #
        iRot As Integer              ' rot #
        iSeq As Integer              ' Seq # (unused for now)
        lSdfCode As Long          ' Internal spot code (used to keep same spot info together when time zone copy used)
        iVefCode As Integer          ' vehicle name that spot is in
        sBktType As String * 1       ' G = generic copy (vs REgional)
        lCifCode As Long              ' CIF (inventory internal code)
        iCrfVefCode As Integer       ' rotation vehicle
        iShttCode As Integer
        sSpotType As String * 1      ' M = Missed
        iRegionFlag As Integer         ' 0= no regional exists, non zero = at least 1 regional.  Only required because bug in Crystal where it wont
                                     'suppress the section if none exists.  Need to format the section with this flag.
        iZoneIndex As Integer
        lRegionCifCode As Long       'regional copy cif code if applicable
        lAttCode As Long            'agreement code to get the load factor
        lRRsfCode As Long           '10-30-10 region RSF code for assigned copy
        lAstCode As Long            '3-27-12 ast spot code (debug only)
        iastStatus As Integer       '3-27-12 ast spot status (debug only)
End Type

'2-28-13 ASTInfo Array that is sorted
Type ASTSORTKEY
    sKey As String * 25         'vehicle|station|contract#   (6,6,9)
    lCode As Long
    lIndex As Long
    iAdfCode As Integer             'advt internal code
End Type

Type SPOT_RPT_OPTIONS
    sStartDate As String        'start and end dates of filter
    sEndDate As String
    bUseAirDAte As Boolean     'use air date or feed dates for filtering
    iAdvtOption As Integer     'selective advt
    iCreateAstInfo As Integer  'true to create AST Info in tmAstInfo
    iShowExact As Integer      'True indicates to show only those spots fed to station (Not Carried ignored)
'                               False gets everything regardless if carried or not
    iIncludeNonRegionSpots As Integer       'true to include non-region defined spots (generic)
    iFilterCatBy As Integer         'for Adv Fulfillment report to filter out selected categories (state, format, markets, etc)
'                                   -1 indicates no testing, 0 = dma market name, 1 = dma market rank, 2 = format, 3 = msa market name
'                                   4 = msa market rank, 5 = state, 6 = unused (can be station), 7 = time zone, 8 = unused (can be vehicle)
    bFilterAvailNames As Boolean    'true if avail names selectivity
    bIncludePledgeInfo As Boolean   'true to include pledge info (set to false if really not needed tohelp speed up processing)
    'blDiscrepOnly As Boolean        'include discrepancies only (whether it be by network or by station)
    bNetworkDiscrep As Boolean         'if discrepancy only, include  network (vs station)
    bStationDiscrep As Boolean     'if discrepan only, include station (vs network_
    lContractNumber As Long         '6-4-18
End Type

Type AMR
    lCode              As Long               ' Affiliate Measurement report
                                             ' auto-increment code
    iGenDate(0 To 1)   As Integer       ' Generation Date
    lGenTime           As Long               ' Generation Time
    lSmtCode           As Long               ' Station Measurement internal
                                             ' reference code
    lAudience          As Long               ' P12+ audience
    iWeekInfoDate(0 To 1)   As Integer        ' Week of measurement data gathered
    iRunDate(0 To 1)   As Integer            ' Date measurment run(updated)
    sVehicleName       As String * 40        ' Vehicle Name
    sCallLetters       As String * 40        ' Station call letters
    sMarket            As String * 60        ' Market name
    iRank              As Integer            ' Market rank(O-xxxxx)
    sFormat            As String * 60        ' Format description (Country,
                                             ' news,...)
    sOwner             As String * 60        ' Owner Name
    sSalesRep          As String * 20        ' Station sales rep
    sServRep           As String * 20        ' Station service rep
    iWeeksAired        As Integer            ' # weeks aired in last 52
    iWeeksMissing      As Integer            ' # weeks missing
    lSpotsPosted       As Long               ' # spots posted
    lSpotsPostedSNC    As Long               ' # spots posted station
                                             ' non-compliant
    lSpotsPostedNNC    As Long               ' # spots posted network
                                             ' non-compliant
    iDaysSubmitted     As Integer            ' # of unique dates station
                                             ' submitted posting
    iShttCode          As Integer            ' Internal station code reference
    iVefCode           As Integer            ' Internal vehicle code reference
    sString1           As String * 40        ' Extra string field
    sUnused            As String * 40
End Type

Global Const OVERDUE_RPT = 2              'overdue affidavits
Global Const PLEDGEVSAIR_RPT = 13      'pledged vs aired report
Global Const FEDVSAIR_RPT = 14         'fed vs aired report
Global Const VERIFY_RPT = 15           'Feed Verification
Global Const EXPJOURNAL_RPT = 16       'Export Journal (affiliate to web)
Global Const NCR_RPT = 19               'Non-Compliant
Global Const USEROPT_RPT = 20
Global Const SITEOPT_RPT = 21
Global Const REGIONASSIGN_RPT = 22      '1-19-10 Regionaly copy assignment (external)
Global Const REGIONASSIGNTRACE_RPT = 23 'regional copy assignment tracing report (internal)
Global Const AFFSMISSINGWKS_RPT = 24    'Affiliates Missing Weeks
Global Const GROUP_RPT = 25             'Group dumps (markets, format, vehicle, state, time zone)
Global Const ADVFULFILL_RPT = 26                'Advertiser Fulfillment, copy splits by advt
Global Const CONTACTCOMMENTS_RPT = 27       'Contact Comments report
Global Const WEBLOGIMPORT_RPT = 28          '12-15-11 web log import, generated from .txt file
Global Const LOGDELIVERY_RPT = 29              '2-21-12 Log Type report (which vehicles are web, cumulus, mkt,univ, etc)
Global Const SPOTMGMT_Rpt = 30              '3-9-12 Affiliate Spot Management
Global Const EXPHISTORY_Rpt = 31            '6-7-12 Affiliate Export History
Global Const SPORTDECLARE_Rpt = 32          '10-9-12 Sports Declaration contract
Global Const SPORTCLEARANCE_Rpt = 33        '10-16-12 Sports Clearance
Global Const RENEWALSTATUS_Rpt = 34         '11-8-12 Agreement Renewal Status
Global Const ADVCOMPLY_Rpt = 35             '2-25-13 Advertiser Compliance
Global Const RADARCLEAR_Rpt = 36            '9-24-13 Radar spot Clearance
Global Const ADVPLACEMENT_Rpt = 37          '2-10-14 advertiser placement
Global Const MEASUREMENT_Rpt = 38           '12-17-14 Affiliate Measurement
Global Const VEHICLE_VISUAL_RPT = 39        '5-23-16 Vehicle Visual report
Global Const WEB_VENDOR_RPT = 40            '2-9-17 Web VEndor Export/Import report
Global Const DELIVERY_DETAIL_RPT = 41       '5-22-18 Delivery report detail
Global Const STATION_PERSONNEL_RPT = 42     '8-22-18 Station Personnel report detail    FYM
Global Const AGREE_CLUSTER_RPT = 43         '3-27-20 Affiliate agreement cluster report

'2-20-20 Define sort options (Advt Fulfillment, Advt Placement)
Global Const SORT_DMA_NAME = 0
Global Const SORT_DMA_RANK = 1
Global Const SORT_FORMAT = 2
Global Const SORT_ISCI_STN = 3
Global Const SORT_ISCI_VEH = 4
Global Const SORT_MSA_NAME = 5
Global Const SORT_MSA_RANK = 6
Global Const SORT_STATE = 7
Global Const SORT_STN = 8
Global Const SORT_TZ = 9
Global Const SORT_VEH_DMA = 10
Global Const SORT_VEH_RANK = 11

Public Enum CsiReportCall
    StartReports
    Normal
    FinishReports
End Enum

Private tgContracts() As Long                'array of selected contracts from a listbox     Date:8/2/2019   FYM
Private Function mCheckSelectedContracts(ByVal llContract As Long) As Boolean
    Dim ilCounter As Integer
    
    mCheckSelectedContracts = False
    For ilCounter = 0 To UBound(tgContracts) - 1
        If llContract = tgContracts(ilCounter) Then
            mCheckSelectedContracts = True
            Exit For
        End If
    Next ilCounter
    
End Function

'
'           mTestCategory - this tests category selections from Advertiser Fulfillment report.
'           Categories include State, Format, MSA Market/Rank, DMA Market/Rank and Time Zones
'           <input>  ilShttCode -Station code to retrieve categories
'                     ilFilterCatBy: -1 no testing, process spot, 0 & 1 = market code & rank,
'                     the codes changed 2-20-20 (see constants below)
'                     2 = format, 3 & 4 = msa mkt & rank, 5 = state, 7 - time zone
'                     Any other code processes the spot
'SORT_DMA_NAME = 0
'SORT_DMA_RANK = 1
'SORT_FORMAT = 2
'SORT_ISCI_STN = 3
'SORT_ISCI_VEH = 4
'SORT_MSA_NAME = 5
'SORT_MSA_RANK = 6
'SORT_STATE = 7
'SORT_STN = 8
'SORT_TZ = 9
'SORT_VEH_DMA = 10
'SORT_VEH_RANK = 11
'                     tlListBox - list of selected categories
'           return - true to process, false to ignore
Public Function mTestCategory(ilShttCode As Integer, ilFilterCatBy As Integer, tlListBox As control) As Integer
Dim ilFoundCat As Integer
Dim ilLoopCat As Integer
Dim llLoopCat As Long   'was using ilLoopCat (integer)
Dim ilFieldCode As Integer
Dim llFieldCode As Long 'was using ilFieldCode (integer)
Dim slState As String * 2
Dim slTemp As String * 2
    On Error GoTo ErrHand:
        ilFoundCat = False
        'from the station, get dma/msa markets, state, time zone, formats
        SQLQuery = "SELECT shttmktcode, shttmetcode, shttfmtcode, shtttztcode, shttState "
        SQLQuery = SQLQuery & " From shtt where shttcode = " & ilShttCode

        Set tmCatRst = gSQLSelectCall(SQLQuery)
        While Not tmCatRst.EOF
'            If ilFilterCatBy = 0 Or ilFilterCatBy = 1 Then         'market name/rank
            If ilFilterCatBy = SORT_DMA_NAME Or ilFilterCatBy = SORT_DMA_RANK Then         'market name/rank
                llFieldCode = tmCatRst!shttMktCode
'            ElseIf ilFilterCatBy = 2 Then     'format
            ElseIf ilFilterCatBy = SORT_FORMAT Then     'format
                llFieldCode = tmCatRst!shttFmtCode
'            ElseIf ilFilterCatBy = 3 Or ilFilterCatBy = 4 Then     'dma market name/rank
            ElseIf ilFilterCatBy = SORT_MSA_NAME Or ilFilterCatBy = SORT_MSA_RANK Then     'dma market name/rank
                llFieldCode = tmCatRst!shttMetCode
'            ElseIf ilFilterCatBy = 5 Then     'state
            ElseIf ilFilterCatBy = SORT_STATE Then     'state
                slState = tmCatRst!shttState        'get only 2 char (postalname)
'            ElseIf ilFilterCatBy = 7 Then     'time zone
            ElseIf ilFilterCatBy = SORT_TZ Then     'time zone
                llFieldCode = tmCatRst!shttTztCode
            End If
            
'            If ilFilterCatBy = 5 Then                       'State
            If ilFilterCatBy = SORT_STATE Then                       'State
                For ilLoopCat = 0 To tlListBox.ListCount - 1
                    If tlListBox.Selected(ilLoopCat) Then
                        slTemp = Mid(tlListBox.List(ilLoopCat), 1, 2)
                        If slTemp = slState Then
                            ilFoundCat = True
                            Exit For
                        End If
                    End If
                Next ilLoopCat
                
'            ElseIf (ilFilterCatBy >= 0 And ilFilterCatBy <= 4) Or ilFilterCatBy = 7 Then
            ElseIf (ilFilterCatBy = SORT_DMA_NAME Or ilFilterCatBy = SORT_DMA_RANK Or ilFilterCatBy = SORT_FORMAT Or ilFilterCatBy = SORT_MSA_NAME Or ilFilterCatBy = SORT_MSA_RANK Or ilFilterCatBy = SORT_TZ) Then
                For llLoopCat = 0 To tlListBox.ListCount - 1
                    If tlListBox.Selected(llLoopCat) Then
                        If tlListBox.ItemData(llLoopCat) = llFieldCode Then
                            ilFoundCat = True
                            Exit For
                        End If
                    End If
                Next llLoopCat
                ilFoundCat = ilFoundCat
            Else
                ilFoundCat = True
            End If
            
             
            tmCatRst.MoveNext
        Wend
        mTestCategory = ilFoundCat
        Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modRptSubs-mTestCategory"
    Exit Function
End Function
'
'           mFindCopyandREgions - obtain the generic copy and see if any regions exits
'           The regions are printed via a subreport in Crystal.
'
'       GRF variables:
'       grfgendate - generation date for filter
'       grfgentime - generation time for filter
'       grfDate     - date of spot
'       grfTime     - time of spot
'       grfCode2    - spot length
'       grfLong     - CRF (rotation header) internal code
'       grfChfCode  - internal contract code
'       grfPerGenl(1)- line #
'       grfPerGenl(2)- rot #
'       grfPerGenl(3)- Seq # (unused for now)
'       grfPerGenl(4)- Time zone flag (0 = All zones, 1 = EST, 2 = PST, 3 = CST, 4 = MST)
'       grfDollars(1)- Internal spot code (used to keep same spot info together when time zone copy used)
'       grfDollars(2) - Regional copy CIF code
'       grfDollars(3) - Agreement (ATT) code for load factor info
'       grfvefCode  - vehicle name that spot is in
'       grfBktType  - G = generic copy (vs REgional)
'       grfCode4    - CIF (inventory internal code)
'       grfSofCode  - rotation vehicle
'       grfDateType - M = Missed, G = MG, O = OUtside, + = show on inv fill, - is do not show on inv. fill
'       grfrdfcode  - 0= no regional exists, non zero = at least 1 regional.  Only required because bug in Crystal where it wont
'                     suppress the section if none exists.  Need to format the section with this flag.
Sub mFindCopyAndRegions(tlAstInfo As ASTINFO, sGenDate As String, sGenTime As String)
Dim SQLQuery As String
Dim slType As String
Dim ilRet As Integer
Dim ilZone As Integer
Dim slTimeZones As String * 12
Dim ilPos As Integer
Dim ilLoop As Integer
'Dim llTzfCifZone(1 To 6) As Long       'Code number for cif or ccf (if negative)
Dim llTzfCifZone(0 To 5) As Long       'Code number for cif or ccf (if negative)
'Dim ilTzfRotNo(1 To 6) As Integer      'Rotation number
Dim ilTzfRotNo(0 To 5) As Integer      'Rotation number
'Dim slTzfZone(1 To 6) As String * 3    'Zone (ALL for all others) (if assign for EST only then EST for index 1 and ALL for PST, CST, MST as Index 2)
Dim slTzfZone(0 To 5) As String * 3    'Zone (ALL for all others) (if assign for EST only then EST for index 1 and ALL for PST, CST, MST as Index 2)

On Error GoTo ErrHand:

        slTimeZones = "ESTPSTCSTMST"
        tmAssignInfo.sGenDate = sGenDate
        tmAssignInfo.sGenTime = sGenTime
        tmAssignInfo.lSdfCode = 0
        tmAssignInfo.iastStatus = tlAstInfo.iStatus     'ast spot status (debugging)
        tmAssignInfo.lAstCode = tlAstInfo.lCode         'ast spot code (debugging)
'        SQLQuery = "SELECT * From SDF_Spot_Detail Where sdfcode = " & tlAstInfo.lSdfCode      '8-5-19 retrieve the SDF record from AST info
        SQLQuery = "SELECT * From lst left outer join sdf_spot_detail on lstsdfcode = sdfcode  Where lstcode =" & tlAstInfo.lLstCode      '8-5-19 retrieve the SDF record from lst info (astsdfcode may not exist)
        Set rst = gSQLSelectCall(SQLQuery)
        While Not rst.EOF
            tmAssignInfo.iVefCode = tlAstInfo.iVefCode
            'pack date
            gPackDate tlAstInfo.sAirDate, tmAssignInfo.iDate(0), tmAssignInfo.iDate(1)        'sched date
            gPackTime tlAstInfo.sAirTime, tmAssignInfo.iTime(0), tmAssignInfo.iTime(1)        'sched time
            tmAssignInfo.sAirDate = Trim$(tlAstInfo.sAirDate)
            tmAssignInfo.sAirTime = Trim$(tlAstInfo.sAirTime)
            tmAssignInfo.iLine = rst!sdfLineNo        'sch line #
            tmAssignInfo.lChfCode = rst!sdfChfCode     'internal contract code
            tmAssignInfo.iCode2 = tlAstInfo.iLen      'spot length
            tmAssignInfo.iAdfCode = tlAstInfo.iAdfCode
            tmAssignInfo.iShttCode = tlAstInfo.iShttCode
            tmAssignInfo.lRRsfCode = tlAstInfo.lRRsfCode        '3-30-10 test against spot to see if this is the region copy assigned
            'test for Missed here

            tmAssignInfo.lSdfCode = rst!sdfcode 'Internal spot code

            tmAssignInfo.sBktType = "G"            'flag as generic copy vs regional
            tmAssignInfo.iSeq = 0           'seq # for generic/timezone copy
            tmAssignInfo.lCifCode = 0           'initalize copy inventory code
            tmAssignInfo.iRot = 0               'init rotation #
            tmAssignInfo.lCrfCode = 0           'init CRF internal code
            tmAssignInfo.iRegionFlag = 0
            tmAssignInfo.iZoneIndex = 0         'zone index (0 = all zones, 1 = EST, 2 = pst, 3 = cst, 4 = mst)

            If rst!sdfCopy > 0 Then
                tmAssignInfo.lCifCode = rst!sdfCopy     'generic copy
                'copy exists, see if any regional exists for this spot
                'tmRsfSrchKey1.lCode = tmSdf.lCode
                'ilRet = btrGetEqual(hmRsf, tmRsf, imRsfRecLen, tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                'If ilRet = BTRV_ERR_NONE Then
                '8-22-12 garbage date in rsfdateadded; only retrieve rsfcode since we only need to know if rsf exists
                '4/11/13: Bypass Airing and blackout type regions because the lst contents the assigned spot
                'SQLQuery = "SELECT rsfcode From Rsf_Region_Schd_Copy Where rsfsdfcode =" & tlAstInfo.lSdfCode        'retrieve the SDF record from AST info
                SQLQuery = "SELECT rsfcode From Rsf_Region_Schd_Copy "
'                SQLQuery = SQLQuery & " Where (rsfSdfCode = " & tlAstInfo.lSdfCode     '8-5-19
                SQLQuery = SQLQuery & " Where (rsfSdfCode = " & rst!sdfcode             '8-5-19
                SQLQuery = SQLQuery & " AND rsfType <> 'B'"     'Blackout
                SQLQuery = SQLQuery & " AND rsfType <> 'A'" & ")"    'Airing vehicle copy
                Set rst2 = gSQLSelectCall(SQLQuery)
                While Not rst2.EOF
                    'found at least 1 matching regional copy.  Rquired to test in Crystal to suppress regional copy subreport if none exists
                    tmAssignInfo.iRegionFlag = 1
                    rst2.MoveNext
                Wend
                'End If
                If rst!sdfPointer = "1" Then             'regular copy inventory
                    tmAssignInfo.iSeq = 1               'seq #
                    tmAssignInfo.iRot = rst!sdfRotNo    'Rotation #
            
                    If rst!sdfSpotType = "O" Then       'open bb
                        slType = "O"
                    ElseIf rst!sdfSpotType = "C" Then    'closed bb
                        slType = "C"
                    Else
                        slType = "A"
                    End If
                    SQLQuery = "SELECT * From Crf_Copy_Rot_Header Where crfRotType = " & "'" & slType & "'" & " and crfAdfCode = " & rst!sdfAdfCode & " and crfChfCode = " & rst!sdfChfCode & " and crfRotNo = " & rst!sdfRotNo
                    
                    Set tmCrfRst = gSQLSelectCall(SQLQuery)
                    While Not tmCrfRst.EOF
                        If tmCrfRst!crfRotNo = rst!sdfRotNo And (tmCrfRst!crfRotType = slType) And (tmCrfRst!crfAdfCode = rst!sdfAdfCode) Then
                            'found the rotation pattern
                            tmAssignInfo.lCifCode = rst!sdfCopy
                            'retrieve rotation from rotation header
                            'tmAssignInfo.iCrfVefCode = tmCrfRst!CrfVefCode  'rotation vehicle,
                            tmAssignInfo.lCrfCode = tmCrfRst!crfCode               'rotation internal code
                            'ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                            mInsertGrfForCopy
                        End If
                        tmCrfRst.MoveNext
                    Wend
                ElseIf rst!sdfPointer = "3" Then             'time zone copy
                    'tmTzfSrchKey0.lCode = tmSdf.lCopyCode
                    tmAssignInfo.iSeq = 1               'seq #
    
                    'ilRet = btrGetEqual(hmTzf, tmTzf, imTzfRecLen, tmTzfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                    SQLQuery = "SELECT * from TZF_Time_Zone_Copy where tzfCode = " & rst!sdfCopy
                    Set tmTzfRst = gSQLSelectCall(SQLQuery)
                    While Not tmTzfRst.EOF
                        llTzfCifZone(0) = tmTzfRst!tzfCifZone1
                        llTzfCifZone(1) = tmTzfRst!tzfCifZone2
                        llTzfCifZone(2) = tmTzfRst!tzfCifZone3
                        llTzfCifZone(3) = tmTzfRst!tzfCifZone4
                        llTzfCifZone(4) = tmTzfRst!tzfCifZone5
                        llTzfCifZone(5) = tmTzfRst!tzfCifZone6
                        ilTzfRotNo(0) = tmTzfRst!TzfRotNo1
                        ilTzfRotNo(1) = tmTzfRst!TzfRotNo2
                        ilTzfRotNo(2) = tmTzfRst!TzfRotNo3
                        ilTzfRotNo(3) = tmTzfRst!TzfRotNo4
                        ilTzfRotNo(4) = tmTzfRst!TzfRotNo5
                        ilTzfRotNo(5) = tmTzfRst!TzfRotNo6
                        slTzfZone(0) = Trim$(tmTzfRst!tzfZone1)
                        slTzfZone(1) = Trim$(tmTzfRst!tzfZone2)
                        slTzfZone(2) = Trim$(tmTzfRst!tzfZone3)
                        slTzfZone(3) = Trim$(tmTzfRst!tzfZone4)
                        slTzfZone(4) = Trim$(tmTzfRst!tzfZone5)
                        slTzfZone(5) = Trim$(tmTzfRst!tzfZone6)
                        'For ilZone = 1 To 6 Step 1
                        For ilZone = 0 To 5 Step 1
                            tmAssignInfo.lCrfCode = 0             'init each zones crf internal code
                            If (Trim$(slTzfZone(ilZone)) <> "") And (llTzfCifZone(ilZone) <> 0) Then
                                If rst!sdfSpotType = "O" Then       'open bb
                                    slType = "O"
                                ElseIf rst!sdfSpotType = "C" Then       'closed bb
                                    slType = "C"
                                Else
                                    slType = "A"
                                End If
                                'ilRet = btrGetGreaterOrEqual(hmCrf, tmCrf, imCrfRecLen, tmCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                                SQLQuery = "SELECT * From Crf_Copy_Rot_Header Where crfRotType = " & "'" & slType & "'" & " and crfAdfCode = " & rst!sdfAdfCode & " and crfChfCode = " & rst!sdfChfCode & " and crfRotNo = " & ilTzfRotNo(ilZone)
                                Set tmCrfRst = gSQLSelectCall(SQLQuery)
                                While Not tmCrfRst.EOF
                                    If tmCrfRst!crfRotNo = ilTzfRotNo(ilZone) Then
                                        'found the rotation pattern
                                        tmAssignInfo.lCifCode = llTzfCifZone(ilZone)
                                        tmAssignInfo.iCrfVefCode = tmCrfRst!CrfVefCode 'rotation vehicle
                                        tmAssignInfo.iRot = ilTzfRotNo(ilZone)        'Rotation #
                                        'tmAssignInfo.iSeq = tmAssignInfo.iSeq + ilZone      'seq # for each time zone within generic copy
                                        tmAssignInfo.iSeq = tmAssignInfo.iSeq + ilZone + 1    'seq # for each time zone within generic copy
                                        tmAssignInfo.lCrfCode = tmCrfRst!crfCode       'rotation internal code
                                        'determine which time zone for flag to send to crystal
                                        ilPos = InStr(1, slTimeZones, slTzfZone(ilZone))
                                        If ilPos > 0 Then
                                            tmAssignInfo.iZoneIndex = (ilPos \ 3) + 1       'determine zone index
                                            mInsertGrfForCopy
                                            'ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                                            'ilRet = BTRV_ERR_END_OF_FILE            'force exit of dowhile
                                        End If
                                    End If
                                    tmCrfRst.MoveNext
                                Wend
                            Else
                                Exit For            'no more time zones in this record
                            End If
                        Next ilZone
                        tmTzfRst.MoveNext
                    Wend
                End If                              'sdf.sPtType =
            Else                                    'no copy exists
                tmAssignInfo.iRot = 0              'no rotation vehicle exists
                tmAssignInfo.lCifCode = 0          'no copy exists
                'ilRet = btrInsert(hmGrf, tmGrf, imGrfRecLen, INDEXKEY0)
                mInsertGrfForCopy
            End If
            rst.MoveNext
        Wend
        Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "modRptSubs-mFindCopyandRegions"
    Exit Sub
End Sub

'       gGenRegionalCopyRept - create station spots in AST and return in tmAstInfo array.
'       Build one vehicle at a time for all stations
'       <input> hmAst - AST handle
'               StartDate - start date of requested report
'               enddate - end date of requested report
'               tlVehAff - list box of vehicles
'               tlAdvtList - list of advertisers only if advt option; otherwise just a listbox control that wont be tested
'               GenDate - Afr generation date for crystal rtrieval & removal
'               GenTime - Afr generation time for crystal retrieval and removal
'               ilCreateAstoInfo - true to create AST Info in tmAstInfo
'               ilIncludeMissed - true/false
'               llSelectedContracts() as long - array of contracts entered
'        <output> GRF - station spots
'       Gather the valid spots based on the cptt (whether reported or not)
'       and create the spots in AST if necessary.
'       Return the spots in tmASTInfo array, which is created in prepass file
'       grf.  Grf contains all generic copy info; A subreport will contain all the
'       regions defined
'
'
'       1-13-10 This subroutine was extracted and copied from gBuildAstStnClr.  Modified to
'       remove creation of AST records, but instead create GRF records containing
'       generic copy assigned to a spot.  A separate subreport will be written for
'       the crystal side to generate the Regions if any defined.
'       There is an associated report on Traffic is called also called Regional Copy Assignment
'       2-23-10 Add option to exclude spots lacking regional copy
Public Sub gGenRegionalCopyRept(hmAst As Integer, sStartDate As String, sEndDate As String, tlVehAff As control, tlLbcStations As control, tlAdvtList As control, sGenDate As String, sGenTime As String, ilCreateASTInfo As Integer, ilShowExact As Integer, ilIncludeMissed As Integer, llSelectedContracts() As Long)
    
    Dim sSDate As String
    Dim iNoWeeks As Integer
    Dim sRptOption As String
    Dim iLoop As Integer
    Dim iVef As Integer
    Dim iRet As Integer
    Dim iAdfCode As Integer
    Dim llLoopAST As Long
    Dim llCpttCount As Long
    'ReDim ilusecodes(1 To 1) As Integer
    ReDim ilusecodes(0 To 0) As Integer
    Dim ilIncludeCodes As Integer
    'ReDim ilAdvtCodes(1 To 1) As Integer
    ReDim ilAdvtCodes(0 To 0) As Integer
    Dim ilInclExclAdvtCodes As Integer
    Dim ilFoundStation As Integer
    Dim ilfoundAdvt As Integer
    Dim ilTemp As Integer
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim llSpotDate As Long
    Dim ilFoundCntr As Integer
    Dim ilFoundValidSpot As Integer
    Dim ilLoop As Integer

    '9-2-06 next 5 fields for regional copy
    Dim sRCart As String
    Dim sRProd As String
    Dim sRISCI As String
    Dim sRCreative As String
    Dim lRCrfCsfCode As Long
    Dim lRCrfCode As Long
    Dim llTempDate As Long
    Dim llActualStart As Long
    Dim SQLQuery As String
    Dim SQLAltQuery As String
    Dim blUseAirDAte As Boolean
    Dim blGetPledgeInfo As Boolean
    Dim blGetServiceAttSpots As Boolean
    
    bgTaskBlocked = False
    sgTaskBlockedName = sgReportListName
    
    ReDim tmAstInfo(0 To 0) As ASTINFO      'init in case tmastinfo should not be created
    iAdfCode = -1   'ilAdvt  get all advt, filter later
    llStartDate = gDateValue(sStartDate)
    llEndDate = gDateValue(sEndDate)
    
    sSDate = gObtainPrevMonday(sStartDate)      'cp returned status is based on weeks
    iNoWeeks = (DateValue(gAdjYear(sEndDate)) - DateValue(gAdjYear(sSDate))) \ 7 + 1
    
    'llActualStart = DateValue(slActualStartDate)
    gObtainCodes tlLbcStations, ilIncludeCodes, ilusecodes()        'build array of which codes to incl/excl
    gObtainCodes tlAdvtList, ilInclExclAdvtCodes, ilAdvtCodes()        'build array of which advt codes to incl/excl
    
    blUseAirDAte = True         'use air dates vs feed dates
    blGetPledgeInfo = False     'no need for pledge info
    blGetServiceAttSpots = True        'get service agreements spots
    llCpttCount = 0
    For iLoop = 1 To iNoWeeks Step 1
        For iVef = 0 To tlVehAff.ListCount - 1 Step 1       'always loop on vehicle
            If tlVehAff.Selected(iVef) Then
                ''Get CPTT so that Stations requiring CP can be obtained
                SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttvefCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP"
                SQLQuery = SQLQuery + " FROM shtt, cptt, att"
                SQLQuery = SQLQuery + " WHERE (ShttCode = cpttShfCode"
                SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
                'always select the required cptts for the vehicle.  No advertiser testing
                'since the filter will be in SQL statement going to Crystal
                sRptOption = " AND cpttVefCode = " & tlVehAff.ItemData(iVef)
                
                SQLQuery = SQLQuery & sRptOption
                SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sSDate, sgSQLDateForm) & "')"
                Set cprst = gSQLSelectCall(SQLQuery)
                While Not cprst.EOF
                    ReDim tgCPPosting(0 To 1) As CPPOSTING

                    tgCPPosting(0).lCpttCode = cprst!cpttCode
                    tgCPPosting(0).iStatus = cprst!cpttStatus
                    tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                    tgCPPosting(0).lAttCode = cprst!cpttatfCode
                    tgCPPosting(0).iAttTimeType = cprst!attTimeType
                    tgCPPosting(0).iVefCode = cprst!cpttvefcode
                    tgCPPosting(0).iShttCode = cprst!shttCode
                    tgCPPosting(0).sZone = cprst!shttTimeZone
                    tgCPPosting(0).sDate = Format$(sSDate, sgShowDateForm)
                    tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                    
                    ilFoundStation = False
                    If ilIncludeCodes Then
                        For ilTemp = LBound(ilusecodes) To UBound(ilusecodes) - 1 Step 1
                            If ilusecodes(ilTemp) = cprst!shttCode Then
                                ilFoundStation = True
                                Exit For
                            End If
                        Next ilTemp
                    Else
                        ilFoundStation = True        '8/23/99 when more than half selected, selection fixed
                        For ilTemp = LBound(ilusecodes) To UBound(ilusecodes) - 1 Step 1
                            If ilusecodes(ilTemp) = cprst!shttCode Then
                                ilFoundStation = False
                                Exit For
                            End If
                        Next ilTemp
                    End If
                    
                    If (ilFoundStation) Then
                        llCpttCount = llCpttCount + 1       'debugging only
                        'Create AST records
                        igTimes = 1 'By Week
                        iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), iAdfCode, True, True, ilCreateASTInfo, , , , blUseAirDAte, blGetPledgeInfo, blGetServiceAttSpots)
                        
                        'loop thru the AST spots, filter advertisers, contract selectivity and
                        'write valid spot to the prepass file
                        For llLoopAST = LBound(tmAstInfo) To UBound(tmAstInfo) - 1
                            ilfoundAdvt = False
                            If ilInclExclAdvtCodes Then                 'include any of the codes in the list
                                'For ilTemp = 1 To UBound(ilAdvtCodes) - 1 Step 1
                                For ilTemp = LBound(ilAdvtCodes) To UBound(ilAdvtCodes) - 1 Step 1
                                    If ilAdvtCodes(ilTemp) = tmAstInfo(llLoopAST).iAdfCode Then
                                        ilfoundAdvt = True
                                        Exit For
                                    End If
                                Next ilTemp
                            Else                                'exclude any of the codes in the list
                                ilfoundAdvt = True        '8/23/99 when more than half selected, selection fixed
                                For ilTemp = LBound(ilAdvtCodes) To UBound(ilAdvtCodes) - 1 Step 1
                                    If ilAdvtCodes(ilTemp) = tmAstInfo(llLoopAST).iAdfCode Then
                                        ilfoundAdvt = False
                                        Exit For
                                    End If
                                Next ilTemp
                            End If
                            
'                                If ilShowExact Then ' The user selected show exact
'                                'Debug.Print tgStatusTypes(tmAstInfo(llLoopAST).iPledgeStatus).iPledged
'                                    'If (tgStatusTypes(tmAstInfo(llLoopAST).iPledgeStatus).iPledged = 2) Then
'                                    'tet for pledging not to feed; if so, see if the spot was posted as not carried too
'                                    If (tmAstInfo(llLoopAST).iPledgeStatus = 4 Or tmAstInfo(llLoopAST).iPledgeStatus = 8) And (tmAstInfo(llLoopAST).iStatus = 8) Then
'
'                                        ilfoundAdvt = False
'                                    End If
'                                End If
                            If ilfoundAdvt Then
                            
                                llSpotDate = gDateValue(tmAstInfo(llLoopAST).sAirDate)
                                'see if spot falls with requested user dates to  continue
                                ilFoundValidSpot = True
                                If llSpotDate < llStartDate Or llSpotDate > llEndDate Then
                                    ilFoundValidSpot = False
                                End If
                                tmAssignInfo.sSpotType = ""         'default to not a missed spot
                                'test for different types of not aired
                                If (gGetAirStatus(tmAstInfo(llLoopAST).iStatus) >= 2 And gGetAirStatus(tmAstInfo(llLoopAST).iStatus) <= 5) Or (gGetAirStatus(tmAstInfo(llLoopAST).iStatus) = ASTAIR_MISSED_MG_BYPASS) Then    '14 = 4-12-17 status is a not aired spot, or missed mg bypassed
                                    
'                                    '3-27-12 see if this spot is a resolved missed
'                                    SQLAltQuery = "SELECT altLinkToAstCode, altastCode, altMnfMissed,altAiredISCI,  astAirDate, astAirTime, astcode, astStatus, astlsfcode From alt left outer JOIN ast ON altastcode = astcode "
'                                    SQLAltQuery = SQLAltQuery & "Where altlinktoastcode = " & tmAstInfo(llLoopAST).lCode & " or altastcode = " & tmAstInfo(llLoopAST).lCode
'
'                                    Set tmAltForAST = gSQLSelectCall(SQLAltQuery)         'read the associated ALT (associations) for the spot
'                                    While Not tmAltForAST.EOF
'                                        'is there an associated mg or replacement for the missed spot just found?  It has to have a link to the ast spot that it references.  If so, should the associated mg or replacement spot be included in the report?
'                                        If (tmAltForAST!altLinkToAstCode > 0) Then   'this is a resolved missed, dont print it
'                                            ilFoundValidSpot = False
'                                        End If
'                                        tmAltForAST.MoveNext
'                                    Wend
'
                                    If tmAstInfo(llLoopAST).lLkAstCode > 0 Then    'already resolved missed spot, ignore the missed portion
                                        ilFoundValidSpot = False
                                    End If
                                    'its a Not aired spot, see if it is to be included
                                    tmAssignInfo.sSpotType = "M"
                                    If gGetAirStatus(tmAstInfo(llLoopAST).iStatus) = ASTAIR_MISSED_MG_BYPASS Then
                                        tmAssignInfo.sSpotType = "B"                '4-13-17 missed-by mg
                                    End If
                                    If Not ilIncludeMissed Then       'exclude not aired spots
                                        ilFoundValidSpot = False
                                    End If
                                End If

                                ilFoundCntr = False
                                '3-27-12 need to determine how to get the spots if no sdf code
                                'If tmAstInfo(llLoopAST).lCntrNo = 0 Then            'alwaysinclude the contract number if 0, could be bonus spot
                                '    ilFoundValidSpot = True
                                'Else
                                    If LBound(llSelectedContracts) <> UBound(llSelectedContracts) Then     'nothing in list, defaults to ALL
                                        For ilLoop = LBound(llSelectedContracts) To UBound(llSelectedContracts) - 1
                                            If llSelectedContracts(ilLoop) = tmAstInfo(llLoopAST).lCntrNo Then
                                                ilFoundCntr = True
                                                Exit For
                                            End If
                                        Next ilLoop
                                        If Not ilFoundCntr Then
                                            ilFoundValidSpot = False
                                        End If
                                    End If
                                'End If
                            
         
                                If ilFoundValidSpot Then
                                    tmAssignInfo.lRegionCifCode = tmAstInfo(llLoopAST).lRCifCode     'regional CIF code
                                    tmAssignInfo.lAttCode = tmAstInfo(llLoopAST).lAttCode           'agreement code to get load factor
                                    mFindCopyAndRegions tmAstInfo(llLoopAST), sGenDate, sGenTime
                                End If
                            End If
                        Next llLoopAST
                    End If
                    
                    cprst.MoveNext
                Wend
                If (tlLbcStations.ListCount = 0) Or (tlLbcStations.ListCount = tlLbcStations.SelCount) Then
                    gClearASTInfo True
                Else
                    gClearASTInfo False
                End If
            End If
        Next iVef
        sSDate = DateAdd("d", 7, sSDate)
    Next iLoop
    gCloseRegionSQLRst
 
     If bgTaskBlocked And igReportSource <> 2 Then
         gMsgBox "Some spots were blocked during the Report generation." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
    End If
    
    bgTaskBlocked = False
    sgTaskBlockedName = ""
 
    Erase tmAstInfo
    Erase tmCPDat
    Erase ilusecodes, ilAdvtCodes
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "modRptSubs-gBuildAstStnClr"
End Sub
'
'       find all the stations associated with a vehicle that has an agreement
'       Do for a single vehicle only
'
'       <input> ilvefcode - vehicle to find matching stations
'       return - list box filled with stations
Public Sub gFillStations(ilVefCode As Integer, lbcStation As control)
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    lbcStation.Clear
    SQLQuery = "SELECT DISTINCT shttCallLetters, shttCode"
    SQLQuery = SQLQuery & " FROM shtt, att"
    SQLQuery = SQLQuery & " WHERE (attVefCode = " & ilVefCode
    SQLQuery = SQLQuery & " AND shttCode = attShfCode)"
    SQLQuery = SQLQuery & " ORDER BY shttCallLetters"
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        lbcStation.AddItem Trim$(rst!shttCallLetters)
        lbcStation.ItemData(lbcStation.NewIndex) = rst!shttCode
        rst.MoveNext
    Wend
    'chkAllStation.Value = vbChecked
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modRptSubs-gFillStations"
End Sub
'
'       gObtainCodes - determine how many items are selected vs not selected.
'       build array of whichever is less, so that looping and testing of the
'       array is faster

Sub gObtainCodes(tlListBox As control, ilIncludeCodes, ilusecodes() As Integer)
Dim ilHowManyDefined As Integer
Dim ilHowMany As Integer
Dim slNameCode As String
Dim ilLoop As Integer
Dim slCode As String
Dim ilRet As Integer
    ilHowManyDefined = tlListBox.ListCount
    ilHowMany = tlListBox.SelCount
    If ilHowMany > ilHowManyDefined / 2 Then    'more than half selected
        ilIncludeCodes = False
    Else
        ilIncludeCodes = True
    End If
    For ilLoop = 0 To tlListBox.ListCount - 1 Step 1
        If tlListBox.Selected(ilLoop) And ilIncludeCodes Then               'selected ?
            ilusecodes(UBound(ilusecodes)) = tlListBox.ItemData(ilLoop)
            ReDim Preserve ilusecodes(LBound(ilusecodes) To UBound(ilusecodes) + 1)
        Else        'exclude these
            If (Not tlListBox.Selected(ilLoop)) And (Not ilIncludeCodes) Then
                ilusecodes(UBound(ilusecodes)) = tlListBox.ItemData(ilLoop)
                ReDim Preserve ilusecodes(LBound(ilusecodes) To UBound(ilusecodes) + 1)
            End If
        End If
    Next ilLoop
End Sub
'       gBuildAstSpotsByStatus - create station spots in AST and return in tmAstInfo array.
'       Build one vehicle at a time for all stations
'       <input> hmAst - AST handle
'               StartDate - start date of requested report
'               enddate - end date of requested report
'               iRptType - 12-4-12 placed with blUseAirDAtes :  true to use air dates, else false to use feed dates
'               blUseAirDate  - true to use air dates, else false to use feed dates
'               tlVehAff - list box of vehicles
'               ilAdvt -  -1 if all advt, else the advt code to retrieve (12-5-12 removed)
'               ilAdvtOption : true if need to test for one or more advt (if all Advt, set to false)
'               tlAdvtList - list of advertisers only if advt option; otherwise just a listbox control that wont be tested
'               GenDate - Afr generation date for crystal rtrieval & removal
'               GenTime - Afr generation time for crystal retrieval and removal
'               ilCreateAstoInfo - true to create AST Info in tmAstInfo
'               ilShowExact - True indicates to show only those spots fed to station (Not Carried ignored)
'                             False gets everything regardless if carried or not
'               ilIncludeNonRegionSpots - true to include spots with/without regional copy, false to exclude non-region spots
'               ilFilterCategory - for Adv Fulfillment and others to filter out selected categories (state, format, markets, etc)
'                                   -1 indicates no testing, 0 = dma market name, 1 = dma market rank, 2 = format, 3 = msa market name
'                                   4 = msa market rank, 5 = state, 6 = unused (can be station), 7 = time zone, 8 = unused (can be vehicle)
'               lbcCategory - list box containing info selected
'               tlStatusOptions - array of spot statuses to include
'               blFilterAvailNames - false if no testing of avail name selectivity,
'               lbcAvailNames - list box to test if testing avail name selectivity
'               blTestAsSold - optional:  true to test the contract schedule line, else false
'               tlLbcContracts - optional: list of Contracts   Date: 8/1/2019 FYM
'        <output> AFR - station spots
'       Gather the valid spots based on the cptt (whether reported or not)
'       and create the spots in AST if necessary.
'       Return the spots in tmASTInfo array, which is created in prepass file
'       afr.  AFR only contains a pointer to the AST file for faster access.
'       More information such as regional copy and MG/replacement info is placed into AFR
'       to avoid all the aliases for retrieval in Crystal reports.
'       Created:  3-20-12 Modeled from gBuildAstStnClr
'
'Public Sub gBuildAstSpotsByStatus(hmAst As Integer, sStartDate As String, sEndDate As String, blUseAirDate As Boolean, tlVehAff As control, tlLbcStations As control, ilAdvtOption As Integer, tlAdvtList As control, sGenDate As String, sGenTime As String, ilCreateAstInfo As Integer, ilShowExact As Integer, ilIncludeNonRegionSpots As Integer, ilFilterCatBy As Integer, lbcCategory As control, tlStatusOptions As STATUSOPTIONS, Optional blTestAsSold As Boolean = False, Optional blDiscrepOnly As Boolean = False)
'   7-10-13 remove sGenDAte & sGenTime as parameters.  Use global variables
'Public Sub gBuildAstSpotsByStatus(hmAst As Integer, sStartDate As String, sEndDate As String, blUseAirDAte As Boolean, tlVehAff As control, tlLbcStations As control, ilAdvtOption As Integer, tlAdvtList As control, ilCreateAstInfo As Integer, ilShowExact As Integer, ilIncludeNonRegionSpots As Integer, ilFilterCatBy As Integer, lbcCategory As control, tlStatusOptions As STATUSOPTIONS, blFilterAvailNames As Boolean, lbcAvailNames As control, Optional blTestAsSold As Boolean = False, Optional blDiscrepOnly As Boolean = False)
'Public Sub gBuildAstSpotsByStatus(hmAst As Integer, tlSpotRptOptions As SPOT_RPT_OPTIONS, tlStatusOptions As STATUSOPTIONS, tlVehAff As control, tlLbcStations As control, tlAdvtList As control, lbcCategory As control, lbcAvailNames As control, Optional blTestAsSold As Boolean = False, Optional blDiscrepOnly As Boolean = False)
'8-4-14 remove the discrepancy (non-compliant options). option for Non-compliant is in tlSpotRptOptions array.  Noncompliant Flag is always returned from ggetastinfo.
'Public Sub gBuildAstSpotsByStatus(hmAst As Integer, tlSpotRptOptions As SPOT_RPT_OPTIONS, tlStatusOptions As STATUSOPTIONS, tlVehAff As control, tlLbcStations As control, tlAdvtList As control, lbcCategory As control, lbcAvailNames As control, Optional blTestAsSold As Boolean = False)
'5-30-18 change to use an integer array of selected vehicles vs the list box.  Based on speeding up spot clr by contract #
Public Sub gBuildAstSpotsByStatus(hmAst As Integer, tlSpotRptOptions As SPOT_RPT_OPTIONS, tlStatusOptions As STATUSOPTIONS, ilSelectedVehicles() As Integer, tlLbcStations As control, tlAdvtList As control, lbcCategory As control, lbcAvailNames As control, Optional blTestAsSold As Boolean = False, Optional tLbcContracts As control = Nothing)
    Dim sSDate As String
    Dim sEDate As String                   '11-30-12
    Dim iNoWeeks As Integer
    Dim sRptOption As String
    Dim iLoop As Integer
    Dim iVef As Integer
    Dim iRet As Integer
    Dim iAdfCode As Integer
    Dim llLoopAST As Long
    Dim llCpttCount As Long
    'ReDim ilusecodes(1 To 1) As Integer
    ReDim ilusecodes(0 To 0) As Integer
    Dim ilIncludeCodes As Integer
    'ReDim ilAdvtCodes(1 To 1) As Integer
    ReDim ilAdvtCodes(0 To 0) As Integer
    Dim ilInclExclAdvtCodes As Integer
    'ReDim ilAvailCodes(1 To 1) As Integer            '7-11-13 avail names selectivity option
    ReDim ilAvailCodes(0 To 0) As Integer            '7-11-13 avail names selectivity option
    Dim ilInclExclAvailCodes As Integer

    Dim ilFoundStation As Integer
    Dim ilfoundAdvt As Integer
    Dim ilTemp As Integer
    Dim ilFoundFilter As Integer
    '9-2-06 next 5 fields for regional copy
    Dim sRCart As String
    Dim sRProd As String
    Dim sRISCI As String
    Dim sRCreative As String
    Dim lRCrfCsfCode As Long
    Dim lRCrfCode As Long
    Dim llTempDate As Long
    Dim llActualStart As Long
    Dim llPriorDate As Long
    Dim ilStatus As Integer
    Dim ilStatusOK As Boolean
    Dim slMGReplAdfName As String
    Dim slMGReplProdName As String
    Dim llMGReplAirTime As Long
    Dim slMGReplAirDate As String
    Dim ilMissedMnfCode As Integer
    Dim slLinkStatus As String * 1
    Dim slMGReplISCI As String
    Dim slAdvtCodesSelected As String               '11-30-12
    Dim blProcessWeek As Boolean                    '11-30-12
    Dim blAdvtExist As Boolean
    '2-25-13 parameters returned if testing Line parameters for compliancy
    Dim slAllowedStartDate As String
    Dim slAllowedEndDate  As String
    Dim slAllowedStartTime As String
    Dim slAllowedEndTime As String
    Dim ilDays As Integer
    Dim ilAllowedDays(0 To 6) As Integer
    Dim slDays(0 To 6) As String * 1
    Dim slLineDays As String
    Dim ilCompliant As Integer
    '
    Dim ilGetLineParameters As Integer
    Dim llUpper As Long
    Dim llStartASTInfoSort As Long
    Dim llEndASTInfoSort As Long
    Dim llAdfExistsIndex As Long
    Dim slStr As String
    Dim blOnly1AdvtSelected As Boolean
    Dim ilAdvtCountSelected As Integer
    Dim slCallLetters As String
    Dim ilShttRet As Integer
    Dim blRadarOK As Boolean
    ReDim tlRadarHdrInfo(0 To 0) As RADAR_HDRINFO
    Dim ilVefCode As Integer
    Dim ilVefInx As Integer
    Dim slAttPledgeType As String
    Dim slNC As String
    Dim ilHowManyCPTT As Integer
    Dim slNowStart As String
    Dim slNowEnd As String
    
    Dim sStartDate As String
    Dim sEndDate As String
    Dim blUseAirDAte As Boolean
    Dim ilAdvtOption As Integer
    Dim ilCreateASTInfo As Integer
    Dim ilShowExact As Integer
    Dim ilIncludeNonRegionSpots As Integer
    Dim ilFilterCatBy As Integer
    Dim blFilterAvailNames As Boolean
    Dim blIncludePledgeInfo As Boolean
    Dim blRegionExists As Boolean
    Dim blNetworkDiscrep As Boolean
    Dim blStationDiscrep As Boolean
    Dim slNetworkDiscrep As String * 1
    Dim slStationDiscrep As String * 1
    Dim blInclDiscrepAndNonDiscrep As Boolean
    Dim slDaysText As String * 14
    Dim llSingleContract As Long                     '6-4-18
    Dim bSelectedContract As Boolean                'check for selected contracts   Date:8/2/2019   FYM
    Dim blIncludeCopyChanges As Boolean             'flag to include copy changes    Date:2/23/2020
    On Error GoTo ErrHand
    
    bgTaskBlocked = False
    sgTaskBlockedName = sgReportListName
    
    slNowStart = gNow()
    llPriorDate = 0
    
    hmAfr = CBtrvTable(ONEHANDLE)
    iRet = btrOpen(hmAfr, "", sgDBPath & "AFR.mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If iRet <> BTRV_ERR_NONE Then
        gMsgBox "btrOpen Failed on AFR.mkd"
        iRet = btrClose(hmAfr)
        btrDestroy hmAfr
        Exit Sub
    End If
    imAfrRecLen = Len(tmAfr)
    
    lgCPCount = 0
    
    'retrieve the options sent into local variables
    sStartDate = tlSpotRptOptions.sStartDate
    sEndDate = tlSpotRptOptions.sEndDate
    blUseAirDAte = tlSpotRptOptions.bUseAirDAte
    ilAdvtOption = tlSpotRptOptions.iAdvtOption
    ilCreateASTInfo = tlSpotRptOptions.iCreateAstInfo
    ilShowExact = tlSpotRptOptions.iShowExact
    ilIncludeNonRegionSpots = tlSpotRptOptions.iIncludeNonRegionSpots
    ilFilterCatBy = tlSpotRptOptions.iFilterCatBy
    blFilterAvailNames = tlSpotRptOptions.bFilterAvailNames
    blIncludePledgeInfo = tlSpotRptOptions.bIncludePledgeInfo
    blNetworkDiscrep = tlSpotRptOptions.bNetworkDiscrep
    blStationDiscrep = tlSpotRptOptions.bStationDiscrep
    blInclDiscrepAndNonDiscrep = True                   'include all spots, regardless of non-compliant or not
    If (blNetworkDiscrep) Or (blStationDiscrep) Then
        blInclDiscrepAndNonDiscrep = False
    End If
    llSingleContract = tlSpotRptOptions.lContractNumber         '6-4-18, 2-20-20 -1 indicates filter by isci, not contract #
    
    slDaysText = "MOTUWETHFRSASU"
    
    iAdfCode = -1           '12-5-12 assume to retrieve all advertisers in gGetAstInfo
    ReDim tmAstInfo(0 To 0) As ASTINFO      'init in case tmastinfo should not be created
    'iAdfCode = ilAdvt              '12-5-12
    'sSDate = gObtainPrevMonday(sStartDate)      '3-4-13 the date has already been backed up to previous week from the calling rtn

    sSDate = sStartDate                         '3-4-13 the date has already been backed up to previous week from the calling rtn, use the start date that came into rtn
    iNoWeeks = (DateValue(gAdjYear(sEndDate)) - DateValue(gAdjYear(sSDate))) \ 7 + 1        '12-11-09 (doesnt get all the weeks)
    
    'llActualStart = DateValue(slActualStartDate)
    slAdvtCodesSelected = ""                                '11-30-12 setup the sql call to see if advt within week/vehicle.  avoid going thru getting spots for the vehicle/week if by advt and it doesnt exist in LST
    gObtainCodes tlLbcStations, ilIncludeCodes, ilusecodes()        'build array of which codes to incl/excl
    blOnly1AdvtSelected = False
    ilAdvtCountSelected = 0
    If (ilAdvtOption) Then                            'advt option with at least 1 advt to select
        gObtainCodes tlAdvtList, ilInclExclAdvtCodes, ilAdvtCodes()        'build array of which advt codes to incl/excl
        For iLoop = LBound(ilAdvtCodes) To UBound(ilAdvtCodes) - 1
            If Trim$(slAdvtCodesSelected) = "" Then
                If ilInclExclAdvtCodes = True Then                          'include the list
                    slAdvtCodesSelected = " IN (" & Str(ilAdvtCodes(iLoop))
                    ilAdvtCountSelected = ilAdvtCountSelected + 1
                    iAdfCode = ilAdvtCodes(iLoop)
                Else                                                        'exclude the list
                    'if more than half has been excluded, blOnly1AdvtSelected flag remains false so it doesnt go thru single testing; otherwise nothing will be found
                    slAdvtCodesSelected = " Not IN (" & Str(ilAdvtCodes(iLoop))
                End If
            Else
                'has at least one entry
                slAdvtCodesSelected = slAdvtCodesSelected & "," & Str(ilAdvtCodes(iLoop))
                ilAdvtCountSelected = ilAdvtCountSelected + 1
            End If
        Next iLoop
        If Trim$(slAdvtCodesSelected) <> "" Then
            slAdvtCodesSelected = " and lstAdfCode " & slAdvtCodesSelected & ")"
        End If
    End If
     
    If blFilterAvailNames Then              'has avail names been selected?
        gObtainCodes lbcAvailNames, ilInclExclAvailCodes, ilAvailCodes()        'build array of which avail codes to incl/excl
    End If
    
    If ilAdvtCountSelected = 1 Then         'do the single advt testing for speed-up
        blOnly1AdvtSelected = True
    Else
        iAdfCode = -1
    End If
    
    'LOOP on # weeks:
    '    if using feed dates, the calling routine  backs up the start date by 1 week.  This
    '    is required because some pledges may cross into the next week (i.e. pledge to air on Sun, but airing on Mon
    '    in the next week).
    'LOOP on Vehicle:
    '   If Advt option, check to see if advt exists for the vehicle
    '   if advt option and advt exists, or not advt option:
    '       read cptts for the week
    '       LOOP on CPTTs for the week:
    '          if first week and OK to process week and station is a selected one, read pledges to
    '          see if any cross the week, or see if any spots exists in the week with a date fed from
    '          previous week where user has moved the spot.
    '          if these tests find a pledge or spot for the week, or not first week and selected go thru gGetASTInfo.
    '               LOOP on AST spots
    '
    '
    
'debug:  timing parameters
'lgTtlTime1 = 0
'lgTtlTime2 = 0
'lgTtlTime3 = 0
'lgTtlTime5 = 0
'lgTtlTime6 = 0
'lgTtlTime7 = 0
'
'lgTtlTime8 = 0
'lgTtlTime9 = 0
'lgTtlTime10 = 0
'lgTtlTime11 = 0
'lgTtlTime12 = 0
'lgTtlTime14 = 0

    lgRptTtlTime1 = 0

    If igRptIndex = RADARCLEAR_Rpt Then
'        blRadarOK = gBuildRadarInfo(tlVehAff)   'pass list of vehicle list box
        blRadarOK = gBuildRadarInfo(ilSelectedVehicles)         '5-30-18
        If Not blRadarOK Then
            gMsg = "An error building Radar table has occured: gBuildAstSpotsByStatus: "
            'gMsgBox gMsg & gErrSQL.Description & "; Error #" & gErrSQL.NativeError & "; Line #" & Erl, vbCritical
            Exit Sub
        End If
    End If
    llCpttCount = 0
    
'TEST only to see what the optimal time is to get ast spots with just a sql call
'    For iLoop = 1 To iNoWeeks Step 1
'        sEDate = DateAdd("d", 6, sSDate)
'        SQLQuery = "Select * from ast where astairdate >= '" & Format$(sSDate, sgSQLDateForm) & "' and astairdate <= '" & Format$(sEDate, sgSQLDateForm) & "'"
'        Set cprst = gSQLSelectCall(SQLQuery)
'        While Not cprst.EOF
'            lgCpttCount = lgCpttCount + 1
'            cprst.MoveNext
'        Wend
'
'        sSDate = DateAdd("d", 7, sSDate)
'    Next iLoop
'    cprst.Close
'    Exit Sub

    'create array for selected Contracts    Date:8/1/2019   FYM
    ReDim tgContracts(0)
    If Not tLbcContracts Is Nothing Then
        For iLoop = 0 To tLbcContracts.ListCount - 1 Step 1
            If tLbcContracts.Selected(iLoop) Then               'selected ?
'                tLbcContracts.ListIndex = iLoop                '2-20-20  this is deselecting the All contracts checkbox
                If llSingleContract < 0 Then                    '2-20-20 isci filtering
                    tgContracts(UBound(tgContracts)) = tLbcContracts.ItemData(iLoop)
                Else                                            'contract filtering
                    tgContracts(UBound(tgContracts)) = tLbcContracts.List(iLoop)
                End If
                ReDim Preserve tgContracts(LBound(tgContracts) To UBound(tgContracts) + 1)
            End If
        Next iLoop
    End If
    
    For iLoop = 1 To iNoWeeks Step 1
        '11-30-12 determine week processing
        'sSDate = Monday date
        sEDate = DateAdd("d", 6, sSDate)
        
'        For iVef = 0 To tlVehAff.ListCount - 1 Step 1       'always loop on vehicle
        For iVef = 0 To UBound(ilSelectedVehicles) - 1 Step 1       '5-30-18 always loop on vehicle from common integer array
'            If tlVehAff.Selected(iVef) Then
'                ilVefCode = tlVehAff.ItemData(iVef)
                ilVefCode = ilSelectedVehicles(iVef)            '5-30-18
                'if Radar information required, extract the header info to be passed to  routine to obtain the network code
                If igRptIndex = RADARCLEAR_Rpt Then
                    ReDim tlRadarHdrInfo(0 To 0) As RADAR_HDRINFO
                    iRet = 0
                    ilVefInx = gBinarySearchVef(CLng(ilVefCode))
                    For ilTemp = LBound(tgRadarHdrInfo) To UBound(tgRadarHdrInfo) - 1
                        If tgRadarHdrInfo(ilTemp).iVefCode = ilVefCode Then
                            iRet = -1
                            tlRadarHdrInfo(UBound(tlRadarHdrInfo)) = tgRadarHdrInfo(ilTemp)
                            ReDim Preserve tlRadarHdrInfo(0 To UBound(tlRadarHdrInfo) + 1) As RADAR_HDRINFO
                        End If
                        
                    Next ilTemp
                End If
                
                ''Get CPTT so that Stations requiring CP can be obtained
                'SQLQuery = "SELECT shttCallLetters, shttMarket, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttvefCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP"
                'SQLQuery = SQLQuery + " FROM shtt, cptt, att"
                'SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttvefCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, mktName"
                'SQLQuery = SQLQuery + " FROM shtt LEFT OUTER JOIN mkt on shttMktCode = mktCode, cptt, att"

                blAdvtExist = False
                If ilAdvtOption Then                 ' advt option
                    'see if the advt exists for this vehicle; if not, bypass it
'                    SQLQuery = "Select count(*) as AdvtSpotCount from lst where lstLogVefCode = " & tlVehAff.ItemData(iVef) & " and lstLogDate >= '" & Format$(sStartDate, sgSQLDateForm) & "' and lstLogDate <= '" & Format$(sEndDate, sgSQLDateForm) & "'"
                    'backup one week for those spots fed in one week and aired in the following
                    SQLQuery = "Select count(*) as AdvtSpotCount from lst where lstLogVefCode = " & ilVefCode & " and lstLogDate >= '" & Format$((DateAdd("d", -7, sStartDate)), sgSQLDateForm) & "' and lstLogDate <= '" & Format$(sEndDate, sgSQLDateForm) & "'"
                    'SQLQuery = SQLQuery & "lstAdfCode " & slAdvtCodesSelected
                    SQLQuery = SQLQuery & slAdvtCodesSelected
                    Set Advrst = gSQLSelectCall(SQLQuery)
                    'if week 1, its the week that was adjusted because of finding pledged spots in a different week
                    'ie. spot is fed on Sun, but is pledged to air on Mon the following day (which is different week), we need to find those spots
                    If Not Advrst.EOF Then
                        If Advrst.Fields("AdvtSpotCount").Value > 0 Then
                            blAdvtExist = True
                            Advrst.Close
                        End If
                    End If
                Else
                    blAdvtExist = True                'always process if not advt
                End If
                
                If blAdvtExist Then     'gather the cptts for the week.  Flag always true if not advertiser option
                    'debugging to see how many cptts were found for the vehicle/week
'                    SQLQuery = "Select count(*) as CPTTCount from shtt,cptt,att where shttcode = cpttshfcode and attcode = cpttatfcode and cpttvefcode = " & tlVehAff.ItemData(iVef) & " AND cpttStartDate = '" & Format$(sSDate, sgSQLDateForm) & "'"
'                    Set cprst = gSQLSelectCall(SQLQuery)
'                    If Not cprst.EOF Then
'                        If cprst.Fields("CPTTCount").Value > 0 Then
'                            ilHowManyCPTT = cprst.Fields("CPTTCount").Value
'                            cprst.Close
'                        End If
'                    End If
                    
                    SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttvefCode, cpttStatus, cpttPostingStatus, cpttAstStatus,attPledgeType, attPrintCP, attTimeType, attGenCP"
                    SQLQuery = SQLQuery + " FROM shtt, cptt, att"
                    SQLQuery = SQLQuery + " WHERE (ShttCode = cpttShfCode"
                    SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"
    
                    'If iRptType = 0 Then           'station option, unused for this report
                    '    sRptOption = " AND cpttshfCode = " & tlVehAff.ItemData(iVef)
                    'ElseIf iRptType = 1 Then                          'vehicle option
                        
                        'always select the required cptts for the vehicle.  No advertiser testing
                        'since the filter will be in SQL statement going to Crystal
'                        sRptOption = " AND cpttVefCode = " & tlVehAff.ItemData(iVef)
                        sRptOption = " AND cpttVefCode = " & ilVefCode
                    
                    'Else
                    '    iAdfCode = tlVehAff.ItemData(iVef)
                    '    sRptOption = ""     'no selection on station or vehicle
                    'End If
                    
                    SQLQuery = SQLQuery & sRptOption
                    SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sSDate, sgSQLDateForm) & "')"
                    Set cprst = gSQLSelectCall(SQLQuery)
                    While Not cprst.EOF
                        ReDim tgCPPosting(0 To 1) As CPPOSTING
                        tgCPPosting(0).lCpttCode = cprst!cpttCode
                        tgCPPosting(0).iStatus = cprst!cpttStatus
                        tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                        tgCPPosting(0).lAttCode = cprst!cpttatfCode
                        tgCPPosting(0).iAttTimeType = cprst!attTimeType
                        tgCPPosting(0).iVefCode = cprst!cpttvefcode
                        tgCPPosting(0).iShttCode = cprst!shttCode
                        tgCPPosting(0).sZone = cprst!shttTimeZone
                        tgCPPosting(0).sDate = Format$(sSDate, sgShowDateForm)
                        tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                        slAttPledgeType = cprst!attPledgeType
                        'determine valid station for inclusion
'                        ilFoundStation = False
'                        If ilIncludeCodes Then
'                            For ilTemp = 1 To UBound(ilUseCodes) - 1 Step 1
'                                If ilUseCodes(ilTemp) = cprst!shttCode Then
'                                    ilFoundStation = True
'                                    Exit For
'                                End If
'                            Next ilTemp
'                        Else
'                            ilFoundStation = True        '8/23/99 when more than half selected, selection fixed
'                            For ilTemp = 1 To UBound(ilUseCodes) - 1 Step 1
'                                If ilUseCodes(ilTemp) = cprst!shttCode Then
'                                    ilFoundStation = False
'                                    Exit For
'                                End If
'                            Next ilTemp
'                        End If
                        
                        ilFoundStation = gTestIncludeExclude(cprst!shttCode, ilIncludeCodes, ilusecodes())
                        blProcessWeek = True                'default case
'2-5-14 No need to test for week 1 when backing up the date due to using Air Date filtering
'                        If (iLoop = 1) And (blAdvtExist) And (blUseAirDAte) And (ilFoundStation) Then         'if week 1, determine if theres any pledges that indicate spot will run in the following week (pledged After) for a station that has been selected
'                            'advt option or not, see if there are any pledges defined to air following week (AFter)
'                            blProcessWeek = False
'                            'look at DAT (pledge times to see if any pledges not in the same week
'                            SQLQuery = "Select count(*) as PledgeCountAfter from dat where datatfCode = " & cprst!cpttatfCode & " and datPdDayFed = " & "'A'"
'                            Set Advrst = gSQLSelectCall(SQLQuery)
'                            If Not Advrst.EOF Then
'                                If Advrst.Fields("PledgeCountAfter").Value > 0 Then
'                                    blProcessWeek = True
'                                    Advrst.Close
'                                End If
'                            End If
'                            If Not blProcessWeek Then           'no pledges for "A" (pledge after), see if theres any spots that user posted in different week than fed
'                                SQLQuery = "Select count(*) as SpotCountPostedNextWeek from ast where astatfcode = " & cprst!cpttatfCode & "  and astAirDate >= '" & Format$(sNextWkSDate, sgSQLDateForm) & "' and astAirDate <= '" & Format$(sNextWkEDate, sgSQLDateForm) & "' and "
'                                SQLQuery = SQLQuery & " astfeeddate < '" & Format$(sNextWkSDate, sgSQLDateForm) & "'"
'                                Set Advrst = gSQLSelectCall(SQLQuery)
'                                If Not Advrst.EOF Then
'                                    If Advrst.Fields("SpotCountPostedNextWeek").Value > 0 Then
'                                        blProcessWeek = True
'                                        Advrst.Close
'                                    End If
'                                End If
'                            End If
'                        End If
                        lgRptSTime1 = timeGetTime
                        
                        'found a valid station with a matching advt (if not by advt, its always set to process), and its a valid week
                        If (ilFoundStation) And (blProcessWeek) Then
                            llCpttCount = llCpttCount + 1       'debugging only
                            'Create AST records
                            igTimes = 1 'By Week
                            'iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), iAdfCode, True, True, ilCreateAstInfo)
                            '2-5-14 send flag to use air dates to make use of new key, along with whether to obtain pledge info

                            lgCPCount = lgCPCount + 1
                            iRet = gGetAstInfo(hmAst, tmCPDat(), tmAstInfo(), iAdfCode, True, True, ilCreateASTInfo, , , , blUseAirDAte, blIncludePledgeInfo)
                            'Create the prepass file for the Clearance report.  Create all
                            'the spots for the vehicle and filtering will be done when the
                            'crystal report is called.
                            
                            'sort the spots returned for this vehicle/station
                            'create the sort key from tmAstInfo array
                            llUpper = UBound(tmAstInfo)
                            ReDim tmAstInfoSort(0 To llUpper) As ASTSORTKEY
                           
                            For llLoopAST = 0 To UBound(tmAstInfo) - 1
                                slStr = Trim$(Str$(tmAstInfo(llLoopAST).iAdfCode))
                                Do While Len(slStr) < 6
                                    slStr = "0" & slStr
                                Loop
                                tmAstInfoSort(llLoopAST).sKey = Trim$(slStr)

                                slStr = Trim$(Str$(tmAstInfo(llLoopAST).lCntrNo))
                                Do While Len(slStr) < 8
                                    slStr = "0" & slStr
                                Loop
                                tmAstInfoSort(llLoopAST).sKey = Trim$(tmAstInfoSort(llLoopAST).sKey) & "|" & Trim$(slStr)

                                slStr = Trim$(Str$(tmAstInfo(llLoopAST).iLineNo))
                                Do While Len(slStr) < 4
                                    slStr = "0" & slStr
                                Loop
                                tmAstInfoSort(llLoopAST).sKey = Trim$(tmAstInfoSort(llLoopAST).sKey) & "|" & Trim$(slStr)

                                tmAstInfoSort(llLoopAST).lCode = tmAstInfo(llLoopAST).lCode     'ast spot id
                                tmAstInfoSort(llLoopAST).lIndex = llLoopAST
                                tmAstInfoSort(llLoopAST).iAdfCode = tmAstInfo(llLoopAST).iAdfCode
                            Next llLoopAST
                            
                            If llUpper > 0 Then
                                ArraySortTyp fnAV(tmAstInfoSort(), 0), llUpper, 0, LenB(tmAstInfoSort(0)), 0, LenB(tmAstInfoSort(0).sKey), 0
                            End If

                            'dfault to none exists
                            llStartASTInfoSort = 0
                            llEndASTInfoSort = 0
                            If ilAdvtOption Then
                                'can only clamp down the min and max indices if only 1 advt selected
                                If blOnly1AdvtSelected Then
                                    'llAdfExistsIndex = mBinarySearchAdfInASTInfo(ilAdvtCodes(1))
                                    llAdfExistsIndex = mBinarySearchAdfInASTInfo(ilAdvtCodes(LBound(ilusecodes)))
                                    If llAdfExistsIndex >= 0 Then
                                        llStartASTInfoSort = llAdfExistsIndex
                                        llEndASTInfoSort = llAdfExistsIndex
                                        'step backwards from the index found, and find the 1st occurence of the matching advt code
                                        For llLoopAST = llAdfExistsIndex - 1 To LBound(tmAstInfoSort) Step -1
                                            'If tmAstInfoSort(llLoopAST).iAdfCode = ilAdvtCodes(1) Then
                                            If tmAstInfoSort(llLoopAST).iAdfCode = ilAdvtCodes(LBound(ilusecodes)) Then
                                                llStartASTInfoSort = llLoopAST
                                            Else
                                                Exit For
                                            End If
                                        Next llLoopAST
                                        
                                        'now go forward from the index found, and find the last occurence of the matching advt code
                                        For llLoopAST = llAdfExistsIndex + 1 To UBound(tmAstInfoSort) - 1 Step 1
                                            'If tmAstInfoSort(llLoopAST).iAdfCode = ilAdvtCodes(1) Then
                                            If tmAstInfoSort(llLoopAST).iAdfCode = ilAdvtCodes(LBound(ilusecodes)) Then
                                                llEndASTInfoSort = llLoopAST
                                            Else
                                                Exit For
                                            End If
                                        Next llLoopAST
                                        llEndASTInfoSort = llEndASTInfoSort + 1
                                    Else
                                        llEndASTInfoSort = llEndASTInfoSort     'for debugging
                                    End If
                                Else
                                    llStartASTInfoSort = LBound(tmAstInfoSort)
                                    llEndASTInfoSort = UBound(tmAstInfoSort)
                                End If
                            Else                'process all the spots returned from ggetastinfo
                                llStartASTInfoSort = LBound(tmAstInfoSort)
                                llEndASTInfoSort = UBound(tmAstInfoSort)
                            End If
                                
                            'For llLoopAST = LBound(tmAstInfo) To UBound(tmAstInfo) - 1
                             For llUpper = llStartASTInfoSort To llEndASTInfoSort - 1
                                llLoopAST = tmAstInfoSort(llUpper).lIndex
                                slMGReplAdfName = ""
                                slMGReplProdName = ""
                                llMGReplAirTime = 0
                                slMGReplAirDate = ""
                                ilMissedMnfCode = 0
                                slLinkStatus = ""
                                slMGReplISCI = ""
                                slCallLetters = ""
                                ilStatus = gGetAirStatus(tmAstInfo(llLoopAST).iStatus)

                                'if excluding unresolved missed, need to include those that are associated with mg and/or replacements if they are also included
                                ilStatusOK = False
                                
                                'Date: 3/30/2020 modified to include/exclude copy changes only for Spot Management
                                If ((tmAstInfo(llLoopAST).iStatus >= ASTEXTENDED_ISCICHGD) And (tlStatusOptions.iInclCopyChanges)) Then
                                    ilStatusOK = True
                                End If
                                
                                If tmAstInfo(llLoopAST).iCPStatus = 0 Then      'not reported?
                                    If tlStatusOptions.iNotReported Then        'include not reported
                                        ilStatusOK = True
                                    End If
'                                ElseIf tmAstInfo(llLoopAST).iStatus >= ASTEXTENDED_ISCICHGD Then    'Date: 3/24/2020 include copy changes
'                                    If tlStatusOptions.iInclCopyChanges Then
'                                        ilStatusOK = True
'                                    End If
                                ElseIf ilStatus = 0 Then                        'live
                                    If tlStatusOptions.iInclLive0 Then
                                        ilStatusOK = True
                                    End If
                                ElseIf ilStatus = 1 Then                        'delayed aired
                                    If tlStatusOptions.iInclDelay1 Then
                                        ilStatusOK = True
                                    End If
                                ElseIf ilStatus = 6 Then
                                    If tlStatusOptions.iInclAirOutPledge6 Then      'aired outside pledge
                                        ilStatusOK = True
                                    End If
                                ElseIf ilStatus = 7 Then
                                    If tlStatusOptions.iInclAiredNotPledge7 Then      'aired, not pledge
                                        ilStatusOK = True
                                    End If
                                ElseIf ilStatus = 8 Then
                                    If tlStatusOptions.iInclNotCarry8 Then      'not carried
                                        ilStatusOK = True
                                    End If
                                ElseIf ilStatus = 9 Then
                                    If tlStatusOptions.iInclDelayCmmlOnly9 Then      'delay, air comml only
                                        ilStatusOK = True
                                    End If
                                ElseIf ilStatus = 10 Then
                                    If tlStatusOptions.iInclAirCmmlOnly10 Then      'live, air comml only
                                        ilStatusOK = True
                                    End If
    
                                ElseIf ((ilStatus = ASTEXTENDED_MG) And (tlStatusOptions.iInclMG11)) Or ((ilStatus = ASTEXTENDED_REPLACEMENT) And (tlStatusOptions.iInclRepl13)) Then
                                     ilStatusOK = True
                                     'find the associated missed spot that this mg is for
                                     '12-24-13 altairedisci no longer required for copy in mg/repl; always pull from ast
'                                        SQLQuery = "SELECT altastcode, altlinktoastcode, astcode, astshfCode, astAirDate, astAirTime, astlsfcode, adfName,adfcode, lstProd, lstcode, lstadfcode   From alt left outer JOIN ast ON altastcode = astcode "
'                                        SQLQuery = SQLQuery & " INNER JOIN lst on ast.astlsfcode = lstcode "
'                                        SQLQuery = SQLQuery & " INNER JOIN adf_advertisers on lstadfcode = adfcode "
'                                        SQLQuery = SQLQuery & "Where altlinktoastcode = " & tmAstInfo(llLoopAST).lCode & " or altastcode = " & tmAstInfo(llLoopAST).lCode
                                         
                                     SQLQuery = "SELECT  astAirDate, astAirTime, astcode, astStatus, astlkastcode, astshfcode, adfname  From ast "
                                     SQLQuery = SQLQuery & " inner join adf_Advertisers on astadfcode = adfcode "
                                     'SQLQuery = SQLQuery & "Where astlkAstCode  = " & tmAstInfo(llLoopAST).lCode
                                     SQLQuery = SQLQuery & " Where astcode = " & tmAstInfo(llLoopAST).lLkAstCode
                                     Set tmAltForAST = gSQLSelectCall(SQLQuery)         'read the associated ALT (associations) for the spot
                                     ilStatusOK = True
                                     While Not tmAltForAST.EOF
                                         'find the associated missed for this replacement or makegood spot.  If it's a reference it should always be shown.
                                         If tmAltForAST!astLkAstCode = tmAstInfo(llLoopAST).lCode Then         'missed side, get the mnf missed reference
'                                                slMGReplISCI = ""               '12-24-13  tmAltForAST!altAiredISCI
'                                            Else
                                             slMGReplAdfName = Trim$(tmAltForAST!adfName)
                                             'slMGReplProdName = Trim$(tmAltForAST!lstProd)
                                             slMGReplAirDate = Format$(tmAltForAST!astAirDate, sgShowDateForm)
                                             llMGReplAirTime = gTimeToLong(tmAltForAST!astAirTime, False)
                                             ilShttRet = gBinarySearchStationInfoByCode(tmAltForAST!astShfCode)
                                                 If ilShttRet <> -1 Then
                                                     slCallLetters = Trim$(tgStationInfoByCode(ilShttRet).sCallLetters)
                                                 End If
                                             If ilStatus = ASTEXTENDED_MG Then
                                                 slLinkStatus = "R"          'replacement
                                             Else
                                                 slLinkStatus = "M"          'makegood
                                             End If
                                         End If
                                         ilStatus = ilStatus
                                    
                                         tmAltForAST.MoveNext
                                     Wend
                                     
                                ElseIf (ilStatus >= 2 And ilStatus <= 5) Or (ilStatus = ASTAIR_MISSED_MG_BYPASS) Then             '4-12-17 Not aired status or missed mg bypassed
                                    'test for inclusion of Not Aired status
                                    '2-26-14 Remove ALT for mg/replacements due to redesign.  Missed and Mg/repl now have forward/backward pointers
                                    '4-12-17 test for new status Missed MG-bypassed
                                    If (ilStatus = 2 And tlStatusOptions.iInclMissed2 = True) Or (ilStatus = 3 And tlStatusOptions.iInclMissed3 = True) Or (ilStatus = 4 And tlStatusOptions.iInclMissed4 = True) Or (ilStatus = 5 And tlStatusOptions.iInclMissed5 = True) Or (ilStatus = ASTAIR_MISSED_MG_BYPASS And tlStatusOptions.iInclMissedMGBypass14 = True) Then
                                        'if excluding mg and/or replacements, exclude the resolved missed portion, or by option exclude the resolved missed
                                        'SQLQuery = "SELECT altLinkToAstCode, altastCode, altMnfMissed,altAiredISCI, adfName, astAirDate, astAirTime, astcode, astStatus, astlsfcode, lstcode, lstadfcode, lstProd, adfcode  From alt left outer JOIN ast ON altastcode = astcode "
                                        '12-24-13 altairedISCI no longer stored in alt, always retrieve directly from ast
'                                        SQLQuery = "SELECT altLinkToAstCode, altastCode, altMnfMissed, adfName, astAirDate, astAirTime, astcode, astStatus, astlsfcode, lstcode, lstadfcode, lstProd, adfcode  From alt left outer JOIN ast ON altastcode = astcode "
'                                        SQLQuery = SQLQuery & " INNER JOIN lst on ast.astlsfcode = lstcode "
'                                        SQLQuery = SQLQuery & " INNER JOIN adf_advertisers on lstadfcode = adfcode "
'                                        SQLQuery = SQLQuery & "Where altlinktoastcode = " & tmAstInfo(llLoopAST).lCode & " or altastcode = " & tmAstInfo(llLoopAST).lCode
'
'                                        Set tmAltForAST = gSQLSelectCall(SQLQuery)         'read the associated ALT (associations) for the spot
'                                        ilStatusOK = True
'                                        While Not tmAltForAST.EOF
'                                            'is there an associated mg or replacement for the missed spot just found?  It has to have a link to the ast spot that it references.  If so, should the associated mg or replacement spot be included in the report?
'                                            If (tmAltForAST!altLinkToAstCode > 0) And ((Not tlStatusOptions.iInclRepl13 And tmAltForAST!astStatus Mod 100 = ASTEXTENDED_REPLACEMENT) Or (Not tlStatusOptions.iInclMG11 And tmAltForAST!astStatus Mod 100 = ASTEXTENDED_MG)) Then
'                                                ilStatusOK = False
'                                            Else            'reference should exist
'                                                If tmAltForAST!altastcode = tmAstInfo(llLoopAST).lCode Then         'missed side, get the mnf missed reference
'                                                    'there is a mg or replacement
'                                                    ilMissedMnfCode = tmAltForAST!altMnfMissed
'                                                    'get the associated mg/replacement to see what it is
'                                                    SQLQuery = "Select astCode, astshfcode, astStatus, astAirDate, astAirTime from alt  "
'                                                    SQLQuery = SQLQuery & " LEFT OUTER JOIN ast on altlinktoastcode = astcode where astcode = " & tmAltForAST!altLinkToAstCode
'                                                    Set tmASTForMGRepl = gSQLSelectCall(SQLQuery)
'                                                    While Not tmASTForMGRepl.EOF
'                                                        'makegood and/or replacement date & time for the missed spot
'                                                        slMGReplAirDate = Format$(tmASTForMGRepl!astAirDate, sgShowDateForm)
'                                                        llMGReplAirTime = gTimeToLong(tmASTForMGRepl!astAirTime, False)
'                                                        ilShttRet = gBinarySearchStationInfoByCode(tmASTForMGRepl!astShfCode)
'                                                        If ilShttRet <> -1 Then
'                                                            slCallLetters = Trim$(tgStationInfoByCode(ilShttRet).sCallLetters)
'                                                        End If
'
'                                                        If tmASTForMGRepl!astStatus = ASTEXTENDED_MG Then
'                                                            slLinkStatus = "M"
''                                                            If tmASTForMGRepl!astshfcode <> tmAstInfo(lLoopAst).iShttCode Then
''                                                                slLinkStatus = "G"
''                                                            End If
'                                                        ElseIf tmASTForMGRepl!astStatus = ASTEXTENDED_REPLACEMENT Then
'                                                            slLinkStatus = "R"
''                                                            If tmASTForMGRepl!astshfcode <> tmAstInfo(lLoopAst).iShttCode Then
''                                                                slLinkStatus = "P"
''                                                            End If
'                                                        End If
'
'                                                        tmASTForMGRepl.MoveNext
'                                                    Wend
'                                                Else
'                                                    If Not tlStatusOptions.iInclResolveMissed Then      'include the resolved misses
'                                                        ilStatusOK = False
'                                                    Else
'                                                        slMGReplAdfName = Trim$(tmAltForAST!adfName)
'                                                        slMGReplProdName = Trim$(tmAltForAST!lstProd)
'                                                        slMGReplAirDate = Format$(tmAltForAST!astAirDate, sgShowDateForm)
'                                                        llMGReplAirTime = gTimeToLong(tmAltForAST!astAirTime, False)
'                                                        'slMGReplISCI = tmAltForAST!altAiredISCI                         'mg or replacement ISCI aired
'                                                        If tmAltForAST!astStatus = ASTEXTENDED_REPLACEMENT Then
'                                                            slLinkStatus = "R"
'                                                        Else
'                                                            slLinkStatus = "M"
'                                                        End If
'                                                    End If
'
'                                                    ilStatus = ilStatus
'                                                End If
'                                            End If
'                                            tmAltForAST.MoveNext
'                                        Wend
'
                                        '2-26-14 ALT has been removed entirely
                                        ilStatusOK = True
                                        ilMissedMnfCode = tmAstInfo(llLoopAST).iMissedMnfCode           'missed reason
                                        '4-5-19 if makegood, should the resolved missed be included?
                                        '4-9-19 tested inclusion of unresolved spots in wrong place, creating error in excluding missed spots
                                        If tmAstInfo(llLoopAST).lLkAstCode > 0 Then        'include the resolved misses
                                            'is there an associated mg or replacement for the missed spot just found?  It has to have a link to the ast spot that it references.  If so, should the associated mg or replacement spot be included in the report?
                                            If (tmAstInfo(llLoopAST).lLkAstCode > 0 And tlStatusOptions.iInclResolveMissed) Then
                                                'there is a mg or replacement
                                                'get the associated mg/replacement to see what it is
                                                SQLQuery = "Select astCode, astshfcode, astStatus, astAirDate, astAirTime, adfcode, adfName, cpfcode, cpfName  from ast  "
                                                SQLQuery = SQLQuery & " inner join adf_Advertisers on astadfcode = adfcode "
                                                SQLQuery = SQLQuery & " left outer join CPF_Copy_Prodct_ISCI on astcpfcode = cpfcode "
                                                SQLQuery = SQLQuery & " where astcode = " & tmAstInfo(llLoopAST).lLkAstCode
                                                Set tmASTForMGRepl = gSQLSelectCall(SQLQuery)
                                                While Not tmASTForMGRepl.EOF
                                                    If ((tlStatusOptions.iInclRepl13 And tmASTForMGRepl!astStatus Mod 100 = ASTEXTENDED_REPLACEMENT) Or (tlStatusOptions.iInclMG11 And tmASTForMGRepl!astStatus Mod 100 = ASTEXTENDED_MG)) Then
                                                        'makegood and/or replacement date & time for the missed spot
                                                        'slMGReplAdfName = Trim$(tmASTForMGRepl!adfName)
    '                                                   slMGReplProdName = Trim$(tmAltForAST!lstProd)
                                                        slMGReplAirDate = Format$(tmASTForMGRepl!astAirDate, sgShowDateForm)
                                                        llMGReplAirTime = gTimeToLong(tmASTForMGRepl!astAirTime, False)
                                                        ilShttRet = gBinarySearchStationInfoByCode(tmASTForMGRepl!astShfCode)
                                                        If ilShttRet <> -1 Then
                                                            slCallLetters = Trim$(tgStationInfoByCode(ilShttRet).sCallLetters)
                                                        End If
                                                        If tmASTForMGRepl!astStatus Mod 100 = ASTEXTENDED_MG Then       '4-13-17 adjust for extended status code
                                                            slLinkStatus = "M"
                                                        ElseIf tmASTForMGRepl!astStatus Mod 100 = ASTEXTENDED_REPLACEMENT Then          '4-13-17 adjust for extended status code
                                                            slLinkStatus = "R"
                                                            slMGReplAdfName = Trim$(tmASTForMGRepl!adfName)
                                                            If Not IsNull(tmASTForMGRepl!cpfName) Then
                                                                slMGReplProdName = Trim$(tmASTForMGRepl!cpfName)
                                                            End If
                                                        End If
                                                    End If
                                                    tmASTForMGRepl.MoveNext
                                                Wend
                                            Else
                                                ilStatusOK = False
                                            End If
                                        End If
                                        
                                    Else                    'missed not selected for inclusion, but still need to include it if it has a mg or replacement that is to be shown
                                        If tmAstInfo(llLoopAST).lLkAstCode > 0 And Not tlStatusOptions.iInclResolveMissed Then       'include the resolved misses
                                            ilStatusOK = False
                                        Else          'a mg or replacement is defined
                                            'read the mg/replace spot then see if they should be included.  if so, need to bring in that associated missed spot
                                            
                                            SQLQuery = "SELECT  adfName, astAirDate, astAirTime, astcode, astStatus  From ast "
                                            SQLQuery = SQLQuery & " INNER JOIN adf_advertisers on astadfcode = adfcode "
                                            'SQLQuery = SQLQuery & "Where astlkAstCode  = " & tmAstInfo(llLoopAST).lCode
                                            SQLQuery = SQLQuery & " Where astcode = " & tmAstInfo(llLoopAST).lLkAstCode
                                            Set tmASTForMGRepl = gSQLSelectCall(SQLQuery)         'read the associated ALT (associations) for the spot
                                            ilStatusOK = False      'True
                                            While Not tmASTForMGRepl.EOF
                                                'is there an associated mg or replacement for the missed spot just found?  It has to have a link to the ast spot that it references.  If so, should the associated mg or replacement spot be included in the report?
                                                If ((tlStatusOptions.iInclRepl13 And tmASTForMGRepl!astStatus Mod 100 = ASTEXTENDED_REPLACEMENT) Or (tlStatusOptions.iInclMG11 And tmASTForMGRepl!astStatus Mod 100 = ASTEXTENDED_MG)) Then
                                                    ilStatusOK = True
                                                    slMGReplAdfName = Trim$(tmASTForMGRepl!adfName)
                                                    'slMGReplProdName = Trim$(tmAltForAST!lstProd)
                                                    slMGReplAirDate = Format$(tmASTForMGRepl!astAirDate, sgShowDateForm)
                                                    llMGReplAirTime = gTimeToLong(tmASTForMGRepl!astAirTime, False)
                                                    If tmASTForMGRepl!astStatus Mod 100 = ASTEXTENDED_REPLACEMENT Then      '4-13-17 adjust for extended status code
                                                        slLinkStatus = "R"
                                                    Else
                                                        slLinkStatus = "M"
                                                    End If
                                                End If
    
                                                tmASTForMGRepl.MoveNext
                                            Wend
                                        End If
                                    End If
    
                                ElseIf ilStatus = ASTEXTENDED_BONUS Then
                                    'If tlStatusOptions.iInclBonus = True Then
                                    If tlStatusOptions.iInclBonus12 Then
                                        ilStatusOK = True
                                    End If
                                    
                                Else            'all other status N/G
                                    ilStatusOK = False
                                End If
                               
                                If tlStatusOptions.bStatusDiscrep = True Then               '12-11-13 is this a status discrepancy report (option in fed vs aired)
                                    If ilStatusOK Then                                      'continue only if OK since it will be filtered out if its not OK
                                        If tmAstInfo(llLoopAST).iPledgeStatus = 8 And gGetAirStatus(tmAstInfo(llLoopAST).iStatus) = 8 Then
                                            ilStatusOK = False                          'pledge status matches spot status of Not Carried, not inconsistent
                                        End If
                                        If (tmAstInfo(llLoopAST).iPledgeStatus = 8 And gGetAirStatus(tmAstInfo(llLoopAST).iStatus) <> 8) Or (tmAstInfo(llLoopAST).iPledgeStatus <> 8 And gGetAirStatus(tmAstInfo(llLoopAST).iStatus) = 8) Then    'discrepant if one or the other is a Not carried
                                            ilStatusOK = ilStatusOK
                                        Else
                                            ilStatusOK = False
                                        End If
                                        'if the status was True before, continue to show it as discrepant
                                    End If
                                End If
                                
                                If (ilAdvtOption) Then                   'is option by advt and status is good
'                                    ilfoundAdvt = False
'                                    If ilInclExclAdvtCodes Then                 'include any of the codes in the list
'                                        For ilTemp = 1 To UBound(ilAdvtCodes) - 1 Step 1
'                                            If ilAdvtCodes(ilTemp) = tmAstInfo(llLoopAST).iAdfCode Then
'                                                ilfoundAdvt = True
'                                                Exit For
'                                            End If
'                                        Next ilTemp
'                                    Else                                'exclude any of the codes in the list
'                                        ilfoundAdvt = True        '8/23/99 when more than half selected, selection fixed
'                                        For ilTemp = 1 To UBound(ilAdvtCodes) - 1 Step 1
'                                            If ilAdvtCodes(ilTemp) = tmAstInfo(llLoopAST).iAdfCode Then
'                                                ilfoundAdvt = False
'                                                Exit For
'                                            End If
'                                        Next ilTemp
'                                    End If
                                    ilfoundAdvt = gTestIncludeExclude(tmAstInfo(llLoopAST).iAdfCode, ilInclExclAdvtCodes, ilAdvtCodes())
                                Else                            'not advt option, no need to filter on advt
                                    ilfoundAdvt = True
                                End If
                                
                                If ilShowExact Then ' The user selected show exact
                                'Debug.Print tgStatusTypes(tmAstInfo(llLoopAST).iPledgeStatus).iPledged
                                    'If (tgStatusTypes(tmAstInfo(llLoopAST).iPledgeStatus).iPledged = 2) Then
                                    'tet for pledging not to feed; if so, see if the spot was posted as not carried too
                                    If (tmAstInfo(llLoopAST).iPledgeStatus = 4 Or tmAstInfo(llLoopAST).iPledgeStatus = 8) And (gGetAirStatus(tmAstInfo(llLoopAST).iStatus) = 8) Then
                                    
                                        ilfoundAdvt = False
                                    End If
                                End If
                                
                                '8-7-14 default fields to pledged when not using As sold (for network compliancy)
                                slAllowedStartDate = tmAstInfo(llLoopAST).sPledgeDate
                                slAllowedEndDate = ""
                                
                                For ilDays = 0 To 6
                                    ilAllowedDays(ilDays) = False
                                    slDays(ilDays) = ""
                                Next ilDays
                            
                                slLineDays = Trim$(tmAstInfo(llLoopAST).sPdDays)
                                ilGetLineParameters = 0
                                ilCompliant = True
                               
'                                tmAfr.sAdfName = Trim$(slMGReplAdfName)
'                                tmAfr.sProdName = Trim$(slMGReplProdName)
                                'Added check for selected multiple contracts    Date: 8/2/2019  FYM
                                bSelectedContract = False
                                If (UBound(tgContracts) > 0) Then
                                    If llSingleContract < 0 Then            '2-20-20 isci filtering
                                        bSelectedContract = mCheckSelectedContracts(tmAstInfo(llLoopAST).lCpfCode)      '3-13-20 if regional, its is returning the original code; if no match
                                                                                                                        'see if the regional matches selection
                                        If Not bSelectedContract Then           'not found, try regional
                                            bSelectedContract = mCheckSelectedContracts(tmAstInfo(llLoopAST).lRCpfCode)
                                        Else    'cpf code selected, but check if regional because tmastinfocpfcode is returning the orig code; we dont want to show the spot if the orig copy code matches but the regional doesnt
                                            If tmAstInfo(llLoopAST).lRRsfCode > 0 Then   'its regional, do not include this since the orig copy code shouldnt be used
                                                bSelectedContract = False
                                            End If
                                        End If
                                    Else
                                        bSelectedContract = mCheckSelectedContracts(tmAstInfo(llLoopAST).lCntrNo)
                                    End If
                                End If
                                
                                If (ilfoundAdvt) And (ilStatusOK) And ((llSingleContract = 0 And (UBound(tgContracts) = 0)) Or (llSingleContract = tmAstInfo(llLoopAST).lCntrNo) Or (bSelectedContract = True) Or (llSingleContract = -1 And UBound(tgContracts) = 0)) Then '6-4-18
                                    If blTestAsSold Then                '2-25-13 determine if one of the spot reports to test contract paramters
                                        ilAllowedDays(0) = tmAstInfo(llLoopAST).iLstMon
                                        ilAllowedDays(1) = tmAstInfo(llLoopAST).iLstTue
                                        ilAllowedDays(2) = tmAstInfo(llLoopAST).iLstWed
                                        ilAllowedDays(3) = tmAstInfo(llLoopAST).iLstThu
                                        ilAllowedDays(4) = tmAstInfo(llLoopAST).iLstFri
                                        ilAllowedDays(5) = tmAstInfo(llLoopAST).iLstSat
                                        ilAllowedDays(6) = tmAstInfo(llLoopAST).iLstSun
                                        slLineDays = gDayNames(ilAllowedDays(), slDays(), 2, slLineDays)
                                   
                                       ' tmAfr.sProduct = Trim$(slLineDays) & " " & Trim$(tmAstInfo(llLoopAST).sLstStartDate) & "-" & Trim$(tmAstInfo(llLoopAST).sLstEndDate)
                                        tmAfr.sProdName = Trim$(slLineDays) & " " & Trim$(tmAstInfo(llLoopAST).sLstStartDate) & "-" & Trim$(tmAstInfo(llLoopAST).sLstEndDate)
                                        'sold times
                                        gPackTime tmAstInfo(llLoopAST).sLstLnStartTime, tmAfr.iPledgeStartTime(0), tmAfr.iPledgeStartTime(1)
                                        gPackTime tmAstInfo(llLoopAST).sLstLnEndTime, tmAfr.iPledgeEndTime(0), tmAfr.iPledgeEndTime(1)

                                    Else            'pledged
                                        'tmAfr.sProduct = Trim$(slLineDays) & " " & Trim$(slAllowedStartDate)
                                        tmAfr.sProdName = Trim$(slLineDays) & " " & Trim$(slAllowedStartDate)
                                        If Trim$(slAllowedEndDate) <> "" Then
                                           ' tmAfr.sProduct = tmAfr.sProduct & "-" & Trim$(slAllowedEndDate)
                                            tmAfr.sProdName = tmAfr.sProdName & "-" & Trim$(slAllowedEndDate)
                                        End If
                                        'pledge times
                                        gPackTime tmAstInfo(llLoopAST).sPledgeStartTime, tmAfr.iPledgeStartTime(0), tmAfr.iPledgeStartTime(1)
                                        If Trim$(tmAstInfo(llLoopAST).sPledgeEndTime) = "" Then     'blank end time, make it same as pledge start time
                                            gPackTime tmAstInfo(llLoopAST).sPledgeStartTime, tmAfr.iPledgeEndTime(0), tmAfr.iPledgeEndTime(1)
                                        Else
                                            gPackTime tmAstInfo(llLoopAST).sPledgeEndTime, tmAfr.iPledgeEndTime(0), tmAfr.iPledgeEndTime(1)
                                        End If
                                    End If
'                                        'ilGetLineParameters = gGetLineParameters(True, tmAstInfo(llLoopAST), slAllowedStartDate, slAllowedEndDate, slAllowedStartTime, slAllowedEndTime, ilAllowedDays(), ilCompliant)
'                                        ilGetLineParameters = gGetAgyCompliant(tmAstInfo(llLoopAST), slAllowedStartDate, slAllowedEndDate, slAllowedStartTime, slAllowedEndTime, ilAllowedDays(), ilCompliant)
'                                        'Return parameter:
'                                        '      0=Ok
'                                        '      1=Unable to read AST, returning Pledge date/time as Allowed Date/Time, Compliant = False
'                                        '      2=Unable to read LST, returning Pledge date/time as Allowed Date/Time, Compliant = False
'                                        '      3=Unable to read SDF, returning Pledge date/time as Allowed Date/Time/Day, Compliant determined from Allowed Date/Time
'                                        '      4=Unable to read CLF, returning Pledge date/time as Allowed Date/Time/Day, Compliant determined from Allowed Date/Time
'                                        '      5=Unable to read RDF, returning Pledge date/time as Allowed Date/Time/Day, Compliant determined from Allowed Date/Time
'                                        '      6=Blackout, returning Pledge date/time as Allowed Date/Time/Day, Compliant = False
'                                        '      7=Line and Booked vehicles don't match, returning Pledge date/time as Allowed Date/Time/Day, Compliant determined from Allowed Date/Time
'                                        '      8=SQL Error Logged to file
'
'                                        slLineDays = gDayNames(ilAllowedDays(), slDays(), 2, slLineDays)
'                                        '2 product field strings used for days & dates and times
'                                        tmAfr.sProduct = slLineDays & " " & slAllowedStartDate & "-" & slAllowedEndDate
'                                        tmAfr.sProdName = slAllowedStartTime & "-" & slAllowedEndTime
'
                                    'End If
                                 'filter out any categories if applicable
                                    ilFoundFilter = True
                                    If ilFilterCatBy >= 0 Then        'something to filter (-1 indicates no filter testing)
                                        
                                        ilFoundFilter = mTestCategory(tmAstInfo(llLoopAST).iShttCode, ilFilterCatBy, lbcCategory)
                                    End If
                                    If ilFoundFilter And blFilterAvailNames Then        'if passed all other filters so far, continue for avail names filter testing
                                        iRet = gTestIncludeExclude(tmAstInfo(llLoopAST).iAnfCode, ilInclExclAvailCodes, ilAvailCodes())
                                        If iRet = False Then
                                            ilFoundFilter = False           'no match for user selected avail names
                                        End If
                                    End If
                                   
                                    If ilFoundFilter Then
                                        sRCart = ""
                                        sRProd = ""
                                        sRISCI = ""
                                        sRCreative = ""
                                        lRCrfCsfCode = 0
                                        
                                        'llTempDate = DateValue(Format$(tmAstInfo(llLoopAST).sAirDate, sgShowDateForm))
                                   
                                        '9-2-06 Check if any region copy defined for the spots
                                        'iRet = gGetRegionCopy(tmAstInfo(llLoopAST).iShttCode, tmAstInfo(llLoopAST).lSdfCode, tmAstInfo(llLoopAST).iVefCode, sRCart, sRProd, sRISCI, sRCreative, lRCrfCsfCode, lRCrfCode)
                                        'iRet = gGetRegionCopy(tmAstInfo(llLoopAST), sRCart, sRProd, sRISCI, sRCreative, lRCrfCsfCode, lRCrfCode)
                                        'comment out code to do sql inserts, replace with Pervasive API inserts for speed
        '                                     SQLQuery = "INSERT INTO " & "afr "
        '                                     SQLQuery = SQLQuery & " (afrAstCode, afrAttCode, afrCart, afrProduct, afrISCI , afrCreative, AfrCrfCsfCode, AfrRegionCopyExists, afrGenDate, afrGenTime)  "
        '
        '                                     SQLQuery = SQLQuery & " VALUES (" & tmAstInfo(llLoopAST).lCode & ", " & tmAstInfo(llLoopAST).lAttCode & ", '" & sRCart & "', '" & sRProd & "', '" & sRISCI & "', '" & sRCreative & "', " & lRCrfCsfCode & ", " & iRet & ", "
        '                                     SQLQuery = SQLQuery & "'" & Format$(sGenDate, sgSQLDateForm) & "', '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sGenTime, False))))) & "')"   '", "
        '                                     cnn.BeginTrans
        '                                     'cnn.Execute SQLQuery, rdExecDirect
        '                                    If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
        '                                        GoSub ErrHand:
        '                                    End If
        '                                    cnn.CommitTrans
                                        
                                        tmAfr.lAstCode = tmAstInfo(llLoopAST).lCode
                                        tmAfr.lAttCode = tmAstInfo(llLoopAST).lAttCode
                                        If tmAstInfo(llLoopAST).iRegionType > 0 Then
'                                            sRCart = Trim$(tmAstInfo(llLoopAST).sRCart)
'                                            sRProd = Trim$(tmAstInfo(llLoopAST).sRProduct)
'                                            sRISCI = Trim$(tmAstInfo(llLoopAST).sRISCI)
'                                            sRCreative = Trim$(tmAstInfo(llLoopAST).sRCreativeTitle)
'                                            lRCrfCsfCode = Trim$(tmAstInfo(llLoopAST).lRCrfCsfCode)
                                            tmAfr.sCart = sRCart
                                            
                                            'TTP 10871 - Fed vs Aired and Pledged vs Aired report: product not being shown for spots with regional copy.   It should get the inventory product
                                            'tmAfr.sProduct = sRProd
                                            tmAfr.sProduct = Trim$(tmAstInfo(llLoopAST).sRProduct)
                                            If sgReportListName = "Fed vs Aired Clearance" Or sgReportListName = "Pledged vs Aired Clearance" Then
                                                If Trim(tmAfr.sProduct) = "" Then
                                                    tmAfr.sProduct = Trim(tmAstInfo(llLoopAST).sProd)
                                                End If
                                            End If
                                            tmAfr.sISCI = sRISCI
                                            tmAfr.sCreative = sRCreative
                                            tmAfr.lCrfCsfCode = lRCrfCsfCode
                                            tmAfr.iRegionCopyExists = True      '11-21-14 set the regional flag to show on report
                                            blRegionExists = True
                                        Else
                                            blRegionExists = False
                                            tmAfr.iRegionCopyExists = False     '11-21-14 set the regional flag to show on report
                                            tmAfr.sProduct = Trim(tmAstInfo(llLoopAST).sProd)
                                        End If
                                        
                                        'regional copy info
'                                        tmAfr.sCart = sRCart
'                                        tmAfr.sProduct = sRProd
'                                        tmAfr.sISCI = sRISCI
'                                        tmAfr.sCreative = sRCreative
'                                        tmAfr.lCrfCsfCode = lRCrfCsfCode
'                                        tmAfr.iRegionCopyExists = iRet
                                        'tmAfr.sProdName = Trim$(tmAstInfo(llLoopAST).sProd)
    
                                        'fields for associated miss/mg/replacement adv, dates & times
                                        gPackTimeLong llMGReplAirTime, tmAfr.iMissReplTime(0), tmAfr.iMissReplTime(1)
                                        gPackDate slMGReplAirDate, tmAfr.iMissReplDate(0), tmAfr.iMissReplDate(1)
                                        tmAfr.sAdfName = Trim$(slMGReplAdfName)
                                        'tmAfr.sProdName = Trim$(slMGReplProdName)
                                        tmAfr.iMissedMnfCode = ilMissedMnfCode
                                        tmAfr.sLinkStatus = slLinkStatus
                                        'if mg or replacement, the ISCI aired from ALT supercedes any regional that may have been found for this spot
                                        If Trim$(slMGReplISCI) <> "" Then
                                            tmAfr.sISCI = slMGReplISCI
                                        End If
                                        
                                        '2-25-13 Fields for Compliance
                                        If tmAstInfo(llLoopAST).iCPStatus = 0 Then
                                            tmAfr.sCompliant = "R"          'not reported
                                            ilCompliant = False             'compliant flag for Not reported wasnt set to false
                                        ElseIf ilCompliant = True Then
                                            tmAfr.sCompliant = "Y"
                                        Else
                                            tmAfr.sCompliant = "N"
                                        End If
                                        tmAfr.iSeqNo = ilGetLineParameters      'Error code from gGetLineParameters (for compliance option only).  ie: cannot read cff, rdf, lst, etc.
                                        slNetworkDiscrep = ""                  'assume aired as sold, not discrepant
                                        slStationDiscrep = ""                  'assume aired as pledged, not discrepant
                                        If tmAstInfo(llLoopAST).sAgencyCompliant = "N" Then           'A = aired within pledge, blank not set so assume OK
                                            slNetworkDiscrep = "Y"               'set Network non-compliant
                                        End If
                                        If tmAstInfo(llLoopAST).sStationCompliant = "N" Then        'A = aired within pledge, blank not set so assume OK
                                            slStationDiscrep = "Y"              'set Station non-compliant
                                        End If
                                        tmAfr.lID = tmAstInfo(llLoopAST).iLineNo
                                        tmAfr.sSpotType = Chr((tmAstInfo(llLoopAST).iSpotType))     'needed to indicate fill
                                        
                                        tmAfr.sCallLetters = slCallLetters
                                        gPackDate sgGenDate, tmAfr.iGenDate(0), tmAfr.iGenDate(1)
                                        tmAfr.lGenTime = gTimeToLong(sgGenTime, False)
                                        
                                        If igRptIndex = RADARCLEAR_Rpt Then
                                            slNC = gObtainRadarNetworkCode(tmAstInfo(llLoopAST), ilVefInx, slAttPledgeType, tlRadarHdrInfo())
                                            tmAfr.sCart = slNC      'Radar report doesnt use the Cart field, override since no field defined to store it
                                        End If
                                        '12/13/13
                                        tmAfr.iPledgeStatus = tmAstInfo(llLoopAST).iPledgeStatus
                                        gPackDate tmAstInfo(llLoopAST).sPledgeDate, tmAfr.iPledgeDate(0), tmAfr.iPledgeDate(1)
                                        tmAfr.iAirPlayNo = tmAstInfo(llLoopAST).iAirPlay         '1-7-15
                                        'moved to above
'                                        gPackTime tmAstInfo(llLoopAST).sPledgeStartTime, tmAfr.iPledgeStartTime(0), tmAfr.iPledgeStartTime(1)
'                                        gPackTime tmAstInfo(llLoopAST).sPledgeEndTime, tmAfr.iPledgeStartTime(0), tmAfr.iPledgeEndTime(1)
                                    
                                        '8-4-14 include if:
                                        '   region exists include or if not regional and include non-regional spots
                                        '   including compliant and non-compliant spots (have not selected either network non-compliant or station non-compliant:  show all)
                                        '   include non-compliant only for Network non-compliant or Station non-compliant
                                        If ((blRegionExists = True) Or ((blRegionExists = False) And (ilIncludeNonRegionSpots = True))) And ((blInclDiscrepAndNonDiscrep = True) Or ((blInclDiscrepAndNonDiscrep = False) And ((blNetworkDiscrep = True And slNetworkDiscrep = "Y") Or (blStationDiscrep = True And slStationDiscrep = "Y")))) Then
                                            iRet = btrInsert(hmAfr, tmAfr, imAfrRecLen, INDEXKEY0)
                                        End If
                                    End If          'ilFoundFilter
                                End If              'ilFoundAdvt
                            'Next llLoopAST
                            Next llUpper
                        End If
                        cprst.MoveNext
                        lgRptETime1 = timeGetTime
                        lgRptTtlTime1 = lgRptTtlTime1 + (lgRptETime1 - lgRptSTime1)
                    Wend
                    
                    If (tlLbcStations.ListCount = 0) Or (tlLbcStations.ListCount = tlLbcStations.SelCount) Then
                        gClearASTInfo True
                    Else
                        gClearASTInfo False
                    End If
                End If
'            End If                              '5-30-18 tlVehAff.Selected(iVef)
        Next iVef
        sSDate = DateAdd("d", 7, sSDate)
    Next iLoop
        
    gCloseRegionSQLRst

    If bgTaskBlocked And igReportSource <> 2 Then
         gMsgBox "Some spots were blocked during the Report generation." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
    End If
    
    bgTaskBlocked = False
    sgTaskBlockedName = ""

    On Error Resume Next
    
    cprst.Close
    tmAltForAST.Close
    tmASTForMGRepl.Close
    
    iRet = btrClose(hmAfr)
    btrDestroy hmAfr
    Erase tmAstInfo
    Erase tmAstInfoSort
    Erase tmCPDat
    Erase ilusecodes, ilAdvtCodes
    slNowEnd = gNow()
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "modRptSubs-gBuildAstSpotsByStatus"
End Sub

'       Build the list of all valid export types form Crystal
'       3-11-04
'       Dan M change for cr2008 7-29-09
Public Sub gPopExportTypes(cboFileType As control, Optional blShowFirst As Boolean)
    'blshowfirst...from display, show 'pdf' when from loaded.
      '    0= pdf
      '    1= Excel-all headers
      '    2= excel-column headers
      '    3= excel-no headers
      '    4= Word
      '    5= Text
      '    6= csv
      '    7= rtf
  '  cboFileType.AddItem " "
    With cboFileType
        .AddItem "Adobe Acrobat(PDF)"
        .AddItem "Excel(XLS)-All headers"
        .AddItem "Excel(XLS)-Column headers"
        .AddItem "Excel(XLS)-No headers"
        .AddItem "Word(DOC)"
        .AddItem "Text(TXT)"
        .AddItem "Comma Separated Values(CSV)"
        .AddItem "Rich Text File(RTF)"
        If blShowFirst Then
            .ListIndex = 0
        End If
    End With

'      '    1= pdf
'      '    2= Excel
'      '    3= Word
'      '    4= Text
'      '    5= csv
'      '    6= rtf
'    cboFileType.AddItem " "
'    cboFileType.AddItem "Acrobat PDF"
'    cboFileType.AddItem "Excel"
'    'dan added 4/22/11-
'    cboFileType.AddItem "Excel with report header(XLS)"
'    cboFileType.AddItem "Word for Windows"
'    cboFileType.AddItem "Text"
'    cboFileType.AddItem "Comma separated value"
'    cboFileType.AddItem "Rich text file"
' '   cboFileType.ListIndex = 0
''    cboFileType.AddItem "Acrobat PDF"
''    cboFileType.AddItem "Comma separated value"
''    cboFileType.AddItem "Data Interchange"
''    cboFileType.AddItem "Excel 7"
''    cboFileType.AddItem "Excel 8"
''    cboFileType.AddItem "Text"
''    cboFileType.AddItem "Rich Text"
''    cboFileType.AddItem "Tab separated text"
''    cboFileType.AddItem "Paginated Text"
''    cboFileType.AddItem "Word for Windows"
''    cboFileType.AddItem "Crystal Reports"
'    'cboFileType.ListIndex = 0
End Sub

'*******************************************************
'*                                                     *
'*      Procedure Name:mOpenMsgFile                    *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Open error message file         *
'*                                                     *
'*******************************************************
Function OpenMsgFile(hlMail As Integer, slToFile As String) As Integer
    Dim slDateTime As String
    Dim slFileDate As String
    Dim slNowDate As String
    Dim ilRet As Integer
    Dim sLetter As String

    'On Error GoTo OpenMsgFileErr:
    sLetter = "A"
    Do
        ilRet = 0
        slToFile = sgExportDirectory & "A" & Format$(gNow(), "mm") & Format$(gNow(), "dd") & Format$(gNow(), "yy") & sLetter & ".csv"
        'slDateTime = FileDateTime(slToFile)
        ilRet = gFileExist(slToFile)
        If ilRet = 0 Then
            sLetter = Chr$(Asc(sLetter) + 1)
        End If
    Loop While ilRet = 0
    On Error GoTo 0
    ilRet = 0
    'On Error GoTo OpenMsgFileErr:
    'hlMail = FreeFile
    'Open slToFile For Output As hlMail
    ilRet = gFileOpen(slToFile, "Output", hlMail)
    If ilRet <> 0 Then
        Close hlMail
        hlMail = -1
        gMsgBox "Open File " & slToFile & " error#" & Str$(Err.Number), vbOKOnly
        OpenMsgFile = False
        Exit Function
    End If
    On Error GoTo 0
    OpenMsgFile = True
    Exit Function
'OpenMsgFileErr:
'    ilRet = 1
'    Resume Next
End Function


'Sub gOutputMethod(RptForm As Form, sReptName As String, sOutput As String)
'Dim iType As Integer
'Dim iStartPos As Integer
'Dim sNameNoExt As String
'
''********************
'    Exit Sub
'
'
'
'    RptForm!CRpt1.Destination = crptToFile
'    If RptForm!cboFileType.ListIndex < 0 Then
'        iType = 0
'    Else
'        iType = RptForm!cboFileType.ItemData(RptForm!cboFileType.ListIndex)
'    End If
'    iStartPos = InStr(sReptName, ".rpt")
'    If iStartPos = 0 Then
'        sNameNoExt = sReptName
'    Else
'        sNameNoExt = Left$(sReptName, iStartPos - 1)
'    End If
'    Select Case iType
'        Case 0, 1, 2
'            RptForm!CRpt1.PrintFileType = iType    'crptText
'            RptForm!CRpt1.PrintFileName = sgExportDirectory + sNameNoExt + ".txt"
'        Case 3
'            RptForm!CRpt1.PrintFileType = iType    'crptText
'            RptForm!CRpt1.PrintFileName = sgExportDirectory + sNameNoExt + ".DIF"
'        Case 4
'            RptForm!CRpt1.PrintFileType = iType    'crptText
'            RptForm!CRpt1.PrintFileName = sgExportDirectory + sNameNoExt + ".CSV"
'        Case 7
'            RptForm!CRpt1.PrintFileType = iType    'crptText
'            RptForm!CRpt1.PrintFileName = sgExportDirectory + sNameNoExt + ".RPT"
'        Case 10
'            RptForm!CRpt1.PrintFileType = iType    'crptText
'            RptForm!CRpt1.PrintFileName = sgExportDirectory + sNameNoExt + ".xls"
'        Case 13
'            RptForm!CRpt1.PrintFileType = iType    'crptText
'            RptForm!CRpt1.PrintFileName = sgExportDirectory + sNameNoExt + ".wks"
'        Case 15
'            RptForm!CRpt1.PrintFileType = iType    'crptText
'            RptForm!CRpt1.PrintFileName = sgExportDirectory + sNameNoExt + ".RTF"
'        Case 17
'            RptForm!CRpt1.PrintFileType = iType    'crptText
'            RptForm!CRpt1.PrintFileName = sgExportDirectory + sNameNoExt + ".Doc"
'    End Select
'    sOutput = RptForm!CRpt1.PrintFileName
'    'If RptForm!optRptDest(2).Value = True Then
'        'gMsgBox "Output Send To: " & sOutput, vbInformation
'    'End If
'
'End Sub

'
'
'           gDateDescription - take 2 dates and format the start and end dates
'           I.e. From 1/1/2004 ; All Dates; thru 12/31/2004; 1/1/2004 - 12/31/2004
'           <input> slStartDate - 1st date  (earliest date);
'                   slEndDate - 2nd date (latest date);
'           <return) string containing the concatenated date description
'
'           8-12-04
Public Function gDateDescription(slStartDate As String, slEndDate As String) As String
Dim dFWeek As Date
Dim slStr1 As String
Dim slStr2 As String
Dim slDesc As String

    dFWeek = CDate(slStartDate)
    If slStartDate = "1/1/70" Or slStartDate = "1/1/1970" Then
        slStr1 = ""
    Else
        slStr1 = dFWeek
End If
    
    dFWeek = CDate(slEndDate)
    If slEndDate = "12/31/69" Or slEndDate = "12/31/2069" Then
        slStr2 = ""
    Else
        slStr2 = dFWeek
    End If
    
    If slStr1 = "" And slStr2 = "" Then
        slDesc = "All Dates"
    ElseIf slStr1 = "" Then
        slDesc = "thru " & slStr2
    ElseIf slStr2 = "" Then
        slDesc = " from " & slStr1
    Else
        slDesc = slStr1 & " - " & slStr2
    End If
    
    gDateDescription = slDesc
    
End Function
'
'       determine what spot statuses the user has selected.
'       Send the string to show on the report on which inclusions/exclusions
'       were requested
'
'       <input> lbcStatus - list box containing all the status codes for inclusion/exclusion
'       <output> sStatus - SQL call for the statuses selected
'                slSelection - selection string for crystal
Public Sub gGetSQLStatus(lbcStatus As control, sStatus As String, slSelection As String, ilIncludeNotCarried As Integer)
Dim i As Integer
Dim ilSelected As Integer
Dim ilNotSelected As Integer
Dim slStatusSelected As String
Dim slStatusNotSelected As String

    sStatus = ""
    slStatusSelected = ""
    slStatusNotSelected = ""
    ilSelected = 0
    ilNotSelected = 0
    ilIncludeNotCarried = True
    For i = 0 To lbcStatus.ListCount - 1 Step 1
        If lbcStatus.Selected(i) Then
            ilSelected = ilSelected + 1
            If Len(sStatus) = 0 Then
                sStatus = "and ((Mod(astStatus, 100) = " & lbcStatus.ItemData(i) & ")"
                slStatusSelected = "Included:" & lbcStatus.List(i)
            Else
                sStatus = sStatus & " OR (Mod(astStatus, 100) = " & lbcStatus.ItemData(i) & ")"
                slStatusSelected = slStatusSelected & ", " & lbcStatus.List(i)
            End If
        Else
            ilNotSelected = ilNotSelected + 1
            If lbcStatus.List(i) = "9-Not Carried" Then  'if Not carried not selected, set flag to exclude
                ilIncludeNotCarried = False
            End If
                
            If Len(slStatusNotSelected) = 0 Then
                slStatusNotSelected = "Excluded:" & lbcStatus.List(i)
            Else
                slStatusNotSelected = slStatusNotSelected & "," & lbcStatus.List(i)
            End If
        End If
    Next i
    sStatus = sStatus & ")"
    
    If lbcStatus.ListCount <> ilSelected Then    'not all statuses selected
        'determine if inclusion or exclusion
        If ilSelected >= ilNotSelected Then
            slSelection = Trim$(slStatusNotSelected)       'less exclusions, show them
        Else
            slSelection = Trim(slStatusSelected)           'less inclusions, show them
        End If
    Else        'everything included
        slSelection = "Included:  All Statuses"
    End If
End Sub
'   Dan 7/20/11
' same as above, but crystal needs 'mod' done differently
'       determine what spot statuses the user has selected.
'       Send the string to show on the report on which inclusions/exclusions
'       were requested
'
'       <input> lbcStatus - list box containing all the status codes for inclusion/exclusion
'               ilIncludeNotCarried - true to include the Not Carried Spots
'               ilIncludeNotReported - true to include Not Reported Spots
'       <output> sStatus - SQL call for the statuses selected
'                slSelection - selection string for crystal
Public Sub gGetSQLStatusForCrystal(lbcStatus As control, sStatus As String, slSelection As String, ilIncludeNotCarried As Integer, ilIncludeNotReported As Integer)
'9-19-11 this iis a copy of gGetSqlStatus.
' Dan created this originally when using cr2008; rolling back to cr11 meant it wasn't needed.  Decided to leave and just change code.
Dim i As Integer
Dim ilSelected As Integer
Dim ilNotSelected As Integer
Dim slStatusSelected As String
Dim slStatusNotSelected As String

    sStatus = ""
    slStatusSelected = ""
    slStatusNotSelected = ""
    ilSelected = 0
    ilNotSelected = 0
    ilIncludeNotCarried = True
    For i = 0 To lbcStatus.ListCount - 1 Step 1
        If lbcStatus.Selected(i) Then
            ilSelected = ilSelected + 1
            If Len(sStatus) = 0 Then
                sStatus = "and ((Mod(astStatus, 100) = " & lbcStatus.ItemData(i) & ")"
                slStatusSelected = "Included:" & lbcStatus.List(i)
            Else
                sStatus = sStatus & " OR (Mod(astStatus, 100) = " & lbcStatus.ItemData(i) & ")"
                slStatusSelected = slStatusSelected & ", " & lbcStatus.List(i)
            End If
        Else
            ilNotSelected = ilNotSelected + 1
            If lbcStatus.List(i) = "9-Not Carried" Then  'if Not carried not selected, set flag to exclude
                ilIncludeNotCarried = False
            End If
                
            If Len(slStatusNotSelected) = 0 Then
                slStatusNotSelected = "Excluded:" & lbcStatus.List(i)
            Else
                slStatusNotSelected = slStatusNotSelected & "," & lbcStatus.List(i)
            End If
        End If
    Next i
    sStatus = sStatus & ")"
    
    If lbcStatus.ListCount <> ilSelected Then    'not all statuses selected
        'determine if inclusion or exclusion
        If ilSelected >= ilNotSelected Then
            slSelection = Trim$(slStatusNotSelected)       'less exclusions, show them
            If Not ilIncludeNotReported Then
                slSelection = slSelection & ", Not Reported"
            End If
        Else
            slSelection = Trim(slStatusSelected)           'less inclusions, show them
            If ilIncludeNotReported Then
                slSelection = slSelection & ", Not Reported"
            End If
        End If
    Else        'everything included
        slSelection = "Included:  All Statuses"
    End If

'Dim i As Integer
'Dim ilSelected As Integer
'Dim ilNotSelected As Integer
'Dim slStatusSelected As String
'Dim slStatusNotSelected As String
'
'    sStatus = ""
'    slStatusSelected = ""
'    slStatusNotSelected = ""
'    ilSelected = 0
'    ilNotSelected = 0
'    ilIncludeNotCarried = True
'    For i = 0 To lbcStatus.ListCount - 1 Step 1
'        If lbcStatus.Selected(i) Then
'            ilSelected = ilSelected + 1
'            If Len(sStatus) = 0 Then
'                sStatus = "and (((astStatus mod 100) = " & lbcStatus.ItemData(i) & ")"
'                slStatusSelected = "Included:" & lbcStatus.List(i)
'            Else
'                sStatus = sStatus & " OR ((astStatus mod 100) = " & lbcStatus.ItemData(i) & ")"
'                slStatusSelected = slStatusSelected & ", " & lbcStatus.List(i)
'            End If
'        Else
'            ilNotSelected = ilNotSelected + 1
'            If lbcStatus.List(i) = "9-Not Carried" Then  'if Not carried not selected, set flag to exclude
'                ilIncludeNotCarried = False
'            End If
'
'            If Len(slStatusNotSelected) = 0 Then
'                slStatusNotSelected = "Excluded:" & lbcStatus.List(i)
'            Else
'                slStatusNotSelected = slStatusNotSelected & "," & lbcStatus.List(i)
'            End If
'        End If
'    Next i
'    sStatus = sStatus & ")"
'
'    If lbcStatus.ListCount <> ilSelected Then    'not all statuses selected
'        'determine if inclusion or exclusion
'        If ilSelected >= ilNotSelected Then
'            slSelection = Trim$(slStatusNotSelected)       'less exclusions, show them
'        Else
'            slSelection = Trim(slStatusSelected)           'less inclusions, show them
'        End If
'    Else        'everything included
'        slSelection = "Included:  All Statuses"
'    End If
End Sub
'
'           mInsertGrfForCopy - Update GRF table for Regional Copy Assignment report
'           Create 1 record per AST, indicating the generic copy assigned
'           All spots will be created, regardless if copy is assigned
'           <input> tmAssignInfo as AssignInfo
'
'           See GRF definitions in mFindCopyAndRegions
'
Public Sub mInsertGrfForCopy()
        'option to exclude spot if lacking regions:  test flag (tmassigninfo.iregionflag) if > 0 then at least 1 region exists
        'ckcExclSpotLackReg = if checked, exclude spots without regional copy; if unchecked, ok to see All
        If (tmAssignInfo.iRegionFlag > 0 And frmRgAssignRpt!ckcExclSpotsLackReg.Value = vbChecked) Or (frmRgAssignRpt!ckcExclSpotsLackReg.Value = vbUnchecked) Then
            SQLQuery = "INSERT INTO " & "GRF_Generic_Report"
            SQLQuery = SQLQuery & " (grfgenDate, grfGenTime, "          'gen date & time
            SQLQuery = SQLQuery & " grfDate, grfTime, "                 'sched date & time
            SQLQuery = SQLQuery & " grfBktType, grfDateType, "          'generic flag vs Regional, spot status (missed)
            SQLQuery = SQLQuery & " grfCode2, grfPer1Genl, grfPer2Genl, grfrdfCode, "     'spot length, line #, rot, regionalflag (1 = atleast 1 exists),advertiser code
            SQLQuery = SQLQuery & " grfChfCode,grfLong, grfPer1, "               'chf internal code, crf internal code, sdf internal code
            SQLQuery = SQLQuery & " grfPer3Genl, grfPer4Genl, grfPer3, "               'seq #, time zone flag (for multiple time zone copy), attcode
            SQLQuery = SQLQuery & " grfPer4, grfYear ,"                                 'internal ast spot code; ast spot status
            SQLQuery = SQLQuery & " grfVefCode, grfSofCode , grfCode4, grfPer2) "           'vehicle name that spot is in, Rotation vehicle code, Copy Inventory code
            
            SQLQuery = SQLQuery & " VALUES ('" & Format$(tmAssignInfo.sGenDate, sgSQLDateForm) & "', " & "'" & Round(Trim$(Str$(CLng(gTimeToCurrency(tmAssignInfo.sGenTime, False))))) & "', "
            SQLQuery = SQLQuery & "'" & Format$(tmAssignInfo.sAirDate, sgSQLDateForm) & "', " & "'" & Format$(tmAssignInfo.sAirTime, sgSQLTimeForm) & "', "
            SQLQuery = SQLQuery & "'" & tmAssignInfo.sBktType & "', " & "'" & tmAssignInfo.sSpotType & "', "
            SQLQuery = SQLQuery & "'" & tmAssignInfo.iCode2 & "', " & "'" & tmAssignInfo.iLine & "', " & "'" & tmAssignInfo.iRot & "', " & "'" & tmAssignInfo.iRegionFlag & "', "
            SQLQuery = SQLQuery & "'" & tmAssignInfo.lChfCode & "', " & "'" & tmAssignInfo.lCrfCode & "', " & "'" & tmAssignInfo.lSdfCode & "', "
            SQLQuery = SQLQuery & "'" & tmAssignInfo.iSeq & "', " & "'" & tmAssignInfo.iZoneIndex & "', " & "'" & tmAssignInfo.lAttCode & "', "
            '3-27-12 added ast internal spot code and ast spot status for debugging
            SQLQuery = SQLQuery & "'" & tmAssignInfo.lAstCode & "', " & "'" & tmAssignInfo.iastStatus & "', "
            '3-30-10 change to update from region code, not cifcode to match to see which region copy has been assigned
            SQLQuery = SQLQuery & "'" & tmAssignInfo.iVefCode & "', " & "'" & tmAssignInfo.iShttCode & "', " & "'" & tmAssignInfo.lCifCode & "', " & "'" & tmAssignInfo.lRRsfCode & "') "
            If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand:
                Screen.MousePointer = vbDefault
                gHandleError "AffErrorLog.txt", "modRptSubs-mInsertGrfForCopy"
                Exit Sub
            End If
        End If
        Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modRptSubs-mInsertGrfForCopy"
End Sub
'Public Function gCallNetReporter(ilChoice As CsiReportCall, Optional slReportKeys As String) As Boolean
'    Dim slCommandLine As String
'    Dim blRet As Boolean
'
'    blRet = True
'    slCommandLine = gBuildNetReportCommandLine(ilChoice)
'    On Error GoTo errornoexe
'    Select Case ilChoice
'        Case CsiReportCall.StartReports
'            If Not bgReportModuleRunning Then
'                Shell slCommandLine
'                bgReportModuleRunning = True
'            End If
'        Case CsiReportCall.Normal
'            If Len(slReportKeys) > 0 Then
'                slCommandLine = slCommandLine & " " & slReportKeys
'                Shell slCommandLine
'                bgReportModuleRunning = True
'            End If
'        Case CsiReportCall.FinishReports
'            If bgReportModuleRunning Then
'                Shell slCommandLine
'                bgReportModuleRunning = False
'            End If
'        Case Else
'            blRet = False
'    End Select
'    gCallNetReporter = blRet
'    Exit Function
'errornoexe:
'        gCallNetReporter = False
'End Function
'Private Function gBuildNetReportCommandLine(ilChoice As CsiReportCall) As String
'    Dim slCommandLine As String
'    'dan 6/24/11 changed to 1.1 for excel header export
'    Const MYINTERFACEVERSION As String = "/Version1.1"
'    'dan 3/23/10 added space
'    Const DEBUGMODE = " /D"
'    Const QUITMODE = " /Q"
'    'Dan M 12/16/09  run csiNetReporterAlternate IF 'test' is in the folder name.
'    slCommandLine = gBuildAlternateAsNeeded
'    Select Case ilChoice
'        Case CsiReportCall.StartReports, CsiReportCall.Normal
'            If (Len(sgSpecialPassword) = 4) Then
'                slCommandLine = slCommandLine & DEBUGMODE & " "
'            End If
'            slCommandLine = slCommandLine & MYINTERFACEVERSION & " "
'            If LenB(sgStartupDirectory) > 0 Then
'                slCommandLine = slCommandLine & " """ & sgStartupDirectory & """ "
'            End If
'            If ilChoice = StartReports Then
'                slCommandLine = slCommandLine & " /PreRun"
'            End If
'        Case CsiReportCall.FinishReports
'            slCommandLine = slCommandLine & QUITMODE
'    End Select
'    gBuildNetReportCommandLine = slCommandLine
'End Function
'
'Private Function gBuildAlternateAsNeeded() As String
'    Dim ilRet As Integer
'    Dim slStr As String
'    Dim slTempPath As String
'
'    If InStr(1, sgExeDirectory, "Test", vbTextCompare) > 0 Then
'        slTempPath = sgExeDirectory & "csinetreporteralternate.exe "
'    Else
'        slTempPath = sgExeDirectory & "csinetreporter.exe "
'    End If
'    ilRet = 0
'    On Error GoTo FileErr
'    slStr = FileDateTime(slTempPath)
'    On Error GoTo 0
'    If ilRet = 1 Then
'        If InStr(1, sgExeDirectory, "Test", vbTextCompare) > 0 Then
'            slTempPath = "csinetreporteralternate.exe "
'        Else
'            slTempPath = "csinetreporter.exe "
'        End If
'    End If
'gBuildAlternateAsNeeded = slTempPath
'    Exit Function
'FileErr:
'    ilRet = 1
'    Resume Next
'End Function
Public Function gPopUst() As Integer
    
    Dim ilUpper As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    gPopUst = False
    ilUpper = 0
    ReDim tgUstInfo(0 To 0) As USTINFO
    SQLQuery = "SELECT * FROM UST where ustState = 0 "
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        tgUstInfo(ilUpper).iCode = rst!ustCode
        tgUstInfo(ilUpper).sName = rst!ustname
        tgUstInfo(ilUpper).iDntCode = rst!ustDntCode
        ilUpper = ilUpper + 1
        ReDim Preserve tgUstInfo(0 To ilUpper) As USTINFO
        rst.MoveNext
    Wend
    
    'Now sort them by the mktCode
    If UBound(tgUstInfo) > 1 Then
        ArraySortTyp fnAV(tgUstInfo(), 0), UBound(tgUstInfo), 0, LenB(tgUstInfo(0)), 0, -1, 0
    End If
    
    gPopUst = True
    rst.Close
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "modRptSubs-gPopUst"
    gPopUst = False
    Exit Function
End Function

Public Function gPopDept() As Integer
    
    Dim ilUpper As Integer
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    gPopDept = False
    ilUpper = 0
    ReDim tgDeptInfo(0 To 0) As DEPTINFO
    SQLQuery = "SELECT * FROM DNT "
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        tgDeptInfo(ilUpper).iCode = rst!dntCode
        tgDeptInfo(ilUpper).sName = rst!dntName
        tgDeptInfo(ilUpper).lColor = rst!dntColor
        tgDeptInfo(ilUpper).sType = rst!dntType
        ilUpper = ilUpper + 1
        ReDim Preserve tgDeptInfo(0 To ilUpper) As DEPTINFO
        rst.MoveNext
    Wend
    
    'Now sort them by the mktCode
    If UBound(tgDeptInfo) > 1 Then
        ArraySortTyp fnAV(tgDeptInfo(), 0), UBound(tgDeptInfo), 0, LenB(tgDeptInfo(0)), 0, -1, 0
    End If
    
    gPopDept = True
    rst.Close
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "modRptSubs-gPopDept"
    gPopDept = False
    Exit Function
End Function
'
'           Set all Status Selections to false
'           for report
'
Public Sub gInitStatusSelections(tlStatusOptions As STATUSOPTIONS)
        tlStatusOptions.iInclLive0 = False
        tlStatusOptions.iInclDelay1 = False
        tlStatusOptions.iInclMissed2 = False
        tlStatusOptions.iInclMissed3 = False
        tlStatusOptions.iInclMissed4 = False
        tlStatusOptions.iInclMissed5 = False
        tlStatusOptions.iInclAirOutPledge6 = False
        tlStatusOptions.iInclAiredNotPledge7 = False
        tlStatusOptions.iInclNotCarry8 = False
        tlStatusOptions.iInclDelayCmmlOnly9 = False
        tlStatusOptions.iInclAirCmmlOnly10 = False
        tlStatusOptions.iInclMG11 = False
        tlStatusOptions.iInclBonus12 = False
        tlStatusOptions.iInclRepl13 = False
        tlStatusOptions.iNotReported = False
        tlStatusOptions.iInclResolveMissed = False
        tlStatusOptions.iInclMissedMGBypass14 = False           '4-12-17 new status Missed spot mg is bypassed
        tlStatusOptions.iInclCopyChanges = False                'Date: 2020/3/23 flag to include/exclude copy changed
End Sub
'
'               set the STatus Options requested (missed, aired, mg, etc)
'               gSetStatusOptons
'               <input> lbcStatus - list box of predefined statuses
'               <output> array of status with include/exclude set
Public Sub gSetStatusOptions(lbcStatus As control, ilNotReported As Integer, tlStatusOptions As STATUSOPTIONS)
Dim ilIndex As Integer
Dim ilStatus As Integer

    For ilIndex = 0 To lbcStatus.ListCount - 1
        If lbcStatus.Selected(ilIndex) Then
            ilStatus = lbcStatus.ItemData(ilIndex)
            Select Case ilStatus
            Case 0
                tlStatusOptions.iInclLive0 = True
            Case 1
                tlStatusOptions.iInclDelay1 = True
            Case 2
                tlStatusOptions.iInclMissed2 = True
            Case 3
                tlStatusOptions.iInclMissed3 = True
            Case 4
                tlStatusOptions.iInclMissed4 = True
            Case 5
                tlStatusOptions.iInclMissed5 = True
            Case 6
                tlStatusOptions.iInclAirOutPledge6 = True
            Case 7
                tlStatusOptions.iInclAiredNotPledge7 = True
            Case 8
                tlStatusOptions.iInclNotCarry8 = True
            Case 9
                tlStatusOptions.iInclDelayCmmlOnly9 = True
            Case 10
                tlStatusOptions.iInclAirCmmlOnly10 = True
            Case 11
                tlStatusOptions.iInclMG11 = True
            Case 12
                tlStatusOptions.iInclBonus12 = True
            Case 13
                tlStatusOptions.iInclRepl13 = True
            Case 14
                tlStatusOptions.iInclMissedMGBypass14 = True        '4-12-17 New status as Missed-mg bypassed
            End Select
            If ilNotReported Then
                tlStatusOptions.iNotReported = True
            End If
        End If
    Next ilIndex
    Exit Sub
End Sub
'
'               gPopVtf - build Vehicle text information in tgVtfInfo array for Sports vehicles only
'
Public Function gPopVtf() As Integer
    Dim vtf_rst As ADODB.Recordset
    Dim ilUpper As Integer
    Dim llMax As Long
    Dim rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    SQLQuery = "Select MAX(vtfCode) from VTF_Vehicle_Text"
    Set rst = gSQLSelectCall(SQLQuery)
    If IsNull(rst(0).Value) Then
        ReDim tgVtfInfo(0 To 0) As VEHICLETEXTINFO
        gPopVtf = True
        Exit Function
    End If
    
    llMax = rst(0).Value
    ReDim tgVtfInfo(0 To llMax) As VEHICLETEXTINFO
    
    SQLQuery = "Select * from VTF_Vehicle_Text LEFT OUTER JOIN Vef_VEhicles on vtfVefcode = vefCode where vefType = 'G'"
    Set vtf_rst = gSQLSelectCall(SQLQuery)
    ilUpper = 0
    While Not vtf_rst.EOF
        If (vtf_rst!vtfType = "H") Or (vtf_rst!vtfType = "F") Then
            tgVtfInfo(ilUpper).lCode = vtf_rst!vtfCode
            tgVtfInfo(ilUpper).sType = vtf_rst!vtfType
            tgVtfInfo(ilUpper).iVefCode = vtf_rst!vtfvefCode    'sort field
            ilUpper = ilUpper + 1
        End If
        ilUpper = ilUpper + 1
        vtf_rst.MoveNext
    Wend

    ReDim Preserve tgVtfInfo(0 To ilUpper) As VEHICLETEXTINFO
    'Now sort them by the vefCode
    If UBound(tgVtfInfo) > 1 Then
        ArraySortTyp fnAV(tgVtfInfo(), 0), UBound(tgVtfInfo), 0, LenB(tgVtfInfo(1)), 0, -1, 0
    End If
    
    gPopVtf = True
    vtf_rst.Close
    rst.Close
    Exit Function

ErrHand:
    gHandleError "AffErrorLog.txt", "modRptSubs-gPopvtf"
    gPopVtf = False
    Exit Function
End Function


Public Function gPopSeasons(ilVefCode As Integer) As Integer
    Dim ilUpper As Integer
    Dim rst As ADODB.Recordset
    Dim slDate As String
    Dim llStartDate As Long
    Dim llEndDate As Long
    
    On Error GoTo ErrHand
    gPopSeasons = False
    ilUpper = 0
    ReDim tgSeasonInfo(0 To 0) As SEASONINFO
    SQLQuery = "SELECT * FROM ghf_Game_Header where ghfVefCode =  " & ilVefCode
    Set rst = gSQLSelectCall(SQLQuery)
    While Not rst.EOF
        tgSeasonInfo(ilUpper).lGhfCode = rst!ghfCode
        tgSeasonInfo(ilUpper).sName = Trim$(rst!ghfSeasonName)
        slDate = Format$(rst!ghfSeasonStartDate, "m/d/yyyy")
        llStartDate = gDateValue(slDate)
        tgSeasonInfo(ilUpper).lStartDate = llStartDate
        slDate = Format$(rst!ghfSeasonEndDate, "m/d/yyyy")
        llEndDate = gDateValue(slDate)
        tgSeasonInfo(ilUpper).lEndDate = llEndDate
        ilUpper = ilUpper + 1
        ReDim Preserve tgSeasonInfo(0 To ilUpper) As SEASONINFO
        rst.MoveNext
    Wend
    
    'Now sort them by the start date, descending
    If UBound(tgSeasonInfo) > 1 Then
        ArraySortTyp fnAV(tgSeasonInfo(), 0), UBound(tgSeasonInfo), 1, LenB(tgSeasonInfo(0)), 0, -2, 0
    End If
    
    gPopSeasons = True
    rst.Close
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "modRptSubs-gPopSeasons"
    gPopSeasons = False
    Exit Function
End Function
'
'           mBinarySearchAdfInAstInfo - determine if the advertiser being processed in
'           in the array of spots returned from gGetASTInfo
'
Public Function mBinarySearchAdfInASTInfo(ilAdfCode As Integer) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim llDecrem As Long
    
    llMin = LBound(tmAstInfoSort)
    llMax = UBound(tmAstInfoSort) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If ilAdfCode = tmAstInfoSort(llMiddle).iAdfCode Then
            'found the match
            mBinarySearchAdfInASTInfo = llMiddle
'            'since there could be multiple spots with same advt, backup to first one of the set
'            If llMiddle > LBound(tmAstInfoSort) Then
'                For llDecrem = llMiddle - 1 To LBound(tmAstInfoSort) Step -1
'                    If ilAdfCode <> tmAstInfoSort(llDecrem).iAdfCode Then
'                    mBinarySearchAdfInfoASTInfo = llDecrem
'                        Exit For
'                    Else
'                        mBinarySearchAdfInASTInfo = llDecrem
'                    End If
'                Next llDecrem
                Exit Function
'            End If
        ElseIf ilAdfCode < tmAstInfoSort(llMiddle).iAdfCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchAdfInASTInfo = -1
    Exit Function

ErrHand:
    gMsg = ""
    If (Err.Number <> 0) And (gMsg = "") Then
        gMsg = "A general error has occured in mBinarySearchAdfInASTInfo: "
        gLogMsg "Error: " & gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, "AffErrorLog.Txt", False
        gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    End If
    mBinarySearchAdfInASTInfo = -1
    Exit Function
End Function
'
'               populate the Avail names in global array
'               Sort them in alphabetical order
'
Public Sub gPopAndSortAvailNames(blIncludeAllAvailNames As Boolean, lbcListBox As control)
Dim iLoop As Integer
Dim ilRet As Integer

    ilRet = gPopAvailNames
    'sort the avail names; cant use the list box option if [All] to be included as it should be the first one
    ReDim tlAvailNamesInfo(0 To 0) As AVAILNAMESINFO
    For iLoop = LBound(tgAvailNamesInfo) To UBound(tgAvailNamesInfo) - 1
        tlAvailNamesInfo(iLoop) = tgAvailNamesInfo(iLoop)
        ReDim Preserve tlAvailNamesInfo(0 To UBound(tlAvailNamesInfo) + 1) As AVAILNAMESINFO
    Next iLoop
    
    ArraySortTyp fnAV(tlAvailNamesInfo(), 0), UBound(tlAvailNamesInfo) - 1, 0, LenB(tlAvailNamesInfo(0)), 2, LenB(tlAvailNamesInfo(0).sName), 0

    lbcListBox.Clear
    For iLoop = 0 To UBound(tlAvailNamesInfo) - 1 Step 1
        If iLoop = 0 And blIncludeAllAvailNames Then
            lbcListBox.AddItem "[All Avail Names]"
            lbcListBox.ItemData(lbcListBox.NewIndex) = 0
            lbcListBox.ListIndex = 0                        'default All avails selected
            lbcListBox.AddItem Trim$(tlAvailNamesInfo(iLoop).sName)
            lbcListBox.ItemData(lbcListBox.NewIndex) = tlAvailNamesInfo(iLoop).iCode
        Else
            lbcListBox.AddItem Trim$(tlAvailNamesInfo(iLoop).sName)
            lbcListBox.ItemData(lbcListBox.NewIndex) = tlAvailNamesInfo(iLoop).iCode
        End If
    Next iLoop
    Exit Sub
End Sub

'
'           gTestIncludeExclude - test array of codes (stored in ilUseCodes) to determine if the value
'           should be included or excluded (determined by flag ilIncludeCodes)
'           <input>  ilValue = value to test for inclusion/exclusion
'                    ilIncludeCodes - true = include code, false = exclude code
'                    ilUseCodes() - array of codes to include or exclude
'           <return> true to include, false to exclude
'
Public Function gTestIncludeExclude(ilValue As Integer, ilIncludeCodes As Integer, ilusecodes() As Integer) As Integer
Dim ilTemp As Integer
Dim ilFoundOption As Integer

    If ilIncludeCodes Then          'include the any of the codes in array?
        ilFoundOption = False
        For ilTemp = LBound(ilusecodes) To UBound(ilusecodes) - 1 Step 1
            If ilusecodes(ilTemp) = ilValue Then
                ilFoundOption = True                    'include the matching vehicle
                Exit For
            End If

        Next ilTemp
    Else                            'exclude any of the codes in array?
        ilFoundOption = True
        For ilTemp = LBound(ilusecodes) To UBound(ilusecodes) - 1 Step 1
            If ilusecodes(ilTemp) = ilValue Then
                ilFoundOption = False                  'exclude the matching vehicle
                Exit For
            End If
        Next ilTemp
    End If
    gTestIncludeExclude = ilFoundOption
    Exit Function
End Function
'              gBuildRadarInfo - build 2 array created from the RADAR header (RHT) and
'              Radar Detail (RET).  The RADAR Header info array will contain the vehicle
'              codes for radar vehicles, along with the network code (i.e. DM, DL).  Each
'              entry will also contain a starting and ending pointer into the detail array.
'              The detail array will contain the start and end times of each event maintained
'              for the vehicle.  Affiliate spot times are matched up to the correct event
'              time within the table, and the associated network code is shown with the spot.
'              if an associated time isnt found, nothing is shown and its probably non-
'              compliant
'Public Function gBuildRadarInfo(lbcVehList As control) As Boolean
Public Function gBuildRadarInfo(ilSelectedVehicles() As Integer) As Boolean     '5-30-18 chg to use array vs list box due to speed up of some reports by cnt #

Dim ilLoop As Integer
Dim ilVef As Integer
Dim ilVefCode As Integer
Dim slNC As String * 2
Dim llIndex As Long
Dim llUpperHdr As Long
Dim llUpperDetail As Long
Dim blFirstDetail As Boolean

    gBuildRadarInfo = True
    ReDim tgRadarHdrInfo(0 To 0) As RADAR_HDRINFO
    ReDim tgRadarDetailInfo(0 To 0) As RADAR_DETAILINFO
    
    Set rst_ret = Nothing   'Date: 3/25/2020 initialize recorset to fix error
    Set rst_rht = Nothing   'Date: 3/25/2020 initialize recorset to fix error
    
    On Error GoTo ErrHand
    
'    For ilVef = 0 To lbcVehList.ListCount - 1
    For ilVef = 0 To UBound(ilSelectedVehicles) - 1

'        If lbcVehList.Selected(ilVef) Then
'            ilVefCode = lbcVehList.ItemData(ilVef)
            ilVefCode = ilSelectedVehicles(ilVef)           '5-30-18

            llUpperHdr = UBound(tgRadarHdrInfo)
            tgRadarHdrInfo(llUpperHdr).iVefCode = ilVefCode
            tgRadarHdrInfo(llUpperHdr).lrhtCode = 0
            tgRadarHdrInfo(llUpperHdr).sNetworkCode = ""
            tgRadarHdrInfo(llUpperHdr).lStartInx = -1
            tgRadarHdrInfo(llUpperHdr).lStartInx = -1
            
'            SQLQuery = "SELECT * FROM rht WHERE (rhtVefCode = " & lbcVehList.ItemData(ilVef) & ")"
            SQLQuery = "SELECT * FROM rht WHERE (rhtVefCode = " & ilVefCode & ")"

            Set rst_rht = gSQLSelectCall(SQLQuery)
            Do While Not rst_rht.EOF
                llUpperHdr = UBound(tgRadarHdrInfo)
                tgRadarHdrInfo(llUpperHdr).iVefCode = ilVefCode
                tgRadarHdrInfo(llUpperHdr).lrhtCode = rst_rht!rhtCode           'RHT internal code
                tgRadarHdrInfo(llUpperHdr).sNetworkCode = rst_rht!rhtRadarNetCode 'RHT radar network code, used to show on report
                tgRadarHdrInfo(llUpperHdr).sVehicleCode = rst_rht!rhtRadarVehCode   'RHT radar vehicle code, used to keep unique for same vehicle & network code
                ReDim Preserve tgRadarHdrInfo(0 To llUpperHdr + 1) As RADAR_HDRINFO
                
                blFirstDetail = True            'first time for detail entries for the vehicle/network code
                SQLQuery = "SELECT * FROM ret WHERE (retRhtCode = " & rst_rht!rhtCode & ")"
                Set rst_ret = gSQLSelectCall(SQLQuery)
                Do While Not rst_ret.EOF
                    llUpperDetail = UBound(tgRadarDetailInfo)
                    If blFirstDetail Then
                        tgRadarHdrInfo(llUpperHdr).lStartInx = llUpperDetail
                        blFirstDetail = False
                    End If
                    tgRadarDetailInfo(llUpperDetail).lrhtCode = rst_rht!rhtCode     'internal code for the vehicle/network hdr
                    tgRadarDetailInfo(llUpperDetail).lStartTime = gTimeToLong(Format$(CStr(rst_ret!retStartTime), sgShowTimeWSecForm), False)
                    tgRadarDetailInfo(llUpperDetail).lEndTime = gTimeToLong(Format$(CStr(rst_ret!retEndTime), sgShowTimeWSecForm), True)
                    tgRadarDetailInfo(llUpperDetail).sDayType = rst_ret!retDayType
                    tgRadarHdrInfo(llUpperHdr).lEndInx = UBound(tgRadarDetailInfo)
                    ReDim Preserve tgRadarDetailInfo(0 To UBound(tgRadarDetailInfo) + 1) As RADAR_DETAILINFO
                    rst_ret.MoveNext    'next time  entry detail record
                Loop
            rst_rht.MoveNext            'next vehicle/network code
            Loop
'        End If
    Next ilVef
    If Not rst_ret Is Nothing Then rst_ret.Close    'Date: 3/25/2020 added "Is Nothing" check to fix error
    If Not rst_rht Is Nothing Then rst_rht.Close    'Date: 3/25/2020 added "Is Nothing" check to fix error
    On Error GoTo 0
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "modRptSubs-gBuildRadarInfo"
    Exit Function
End Function
'
'               gObtainRadarNetworkCode - Based on an ast spot, find the associated Network
'               Code that this spot falls within.  The Radar codes are obtained from RHT and RET
'               which contain internal vehicle codes, network codes, and events with times/days.
'
'               <input> ASTINFO - spot from ASTINFO array
'                       ilVefInx - index into Vehicle table array to pull off zone information
'                       slAttPledgeType - agreement Pledege type:  A = avail, D = dp, C = CD
'                       tlRadar_HdrInfo - array of current vehicles radar info, which points to the
'                       array of detail time information (tgRadar_DetailInfo)
'               <Return> 2 char network code from radar tables, or blank if non applicable
'
'       tgRadar_HdrInfo = array of Radar vehicles containing their network codes
'       tgRadar_DetailInfo - array of radar program (avail) times and valid days
'
Public Function gObtainRadarNetworkCode(tlAstInfo As ASTINFO, ilVefInx As Integer, slAttPledgeType As String, tlRadar_HdrInfo() As RADAR_HDRINFO) As String
Dim ilDACode As Integer
Dim ilIncludeSpot As Integer
Dim llDate As Long
Dim llTime As Long
Dim slZone As String
Dim ilZoneFound As Integer
Dim ilNumberAsterisk As Integer
Dim ilZone As Integer
Dim ilLocalAdj As Integer
Dim llSpotTime As Long
Dim slDate As String
Dim ilWeekDay As Integer
Dim slNC  As String
Dim slProgCode As String
Dim blFoundNC As Boolean
Dim ilLoopOnRadarHdr As Integer
Dim llStartDetailInx As Long
Dim llEndDetailInx As Long
Dim llTest As Long
Dim ilStatus As Integer

        slNC = ""
        If (((tgCPPosting(0).iStatus = 1) Or (tgCPPosting(0).iStatus = 2)) And (tgCPPosting(0).iPostingStatus = 2)) Then
             On Error GoTo ErrHand
             ilStatus = gGetAirStatus(tlAstInfo.iStatus)     'get the ast status code (live, delayed, not aired, etc
             If (ilStatus < 2 Or ilStatus > 5) And (ilStatus <> ASTAIR_MISSED_MG_BYPASS) Then                '4-21-17  If missed, dont show a code
                 If slAttPledgeType = "D" Then
                     ilDACode = 0
                 ElseIf slAttPledgeType = "A" Then
                     ilDACode = 1
                 ElseIf slAttPledgeType = "C" Then
                     ilDACode = 2
                 Else
                     ilDACode = -1
                 End If
    
                ilIncludeSpot = True
                llDate = DateValue(tlAstInfo.sFeedDate)
                llTime = gTimeToLong(tlAstInfo.sFeedTime, False)
                'Translate time based on zone
                'Select Case UCase$(Trim$(cprst!shttTimeZone))
                '    Case "EST"
                '        llSpotTime = llTime
                '    Case "CST"
                '        llSpotTime = llTime + 3600
                '    Case "MST"
                '        llSpotTime = llTime + 2 * 3600
                '    Case "PST"
                '        llSpotTime = llTime + 3 * 3600
                '    Case Else
                '        llSpotTime = llTime
                'End Select
                'If (llSpotTime >= 24 * CLng(3600)) Then
                '    'Adjust date
                '    llDate = llDate + 1
                '    llSpotTime = llSpotTime - 24 * CLng(3600)
                'End If
                slZone = UCase$(Trim$(tgCPPosting(0).sZone))
                ilLocalAdj = 0
                ilZoneFound = False
                ilNumberAsterisk = 0
                ' Adjust time zone properly.
                If Len(slZone) <> 0 Then
                    'Get zone
                    DoEvents
                    For ilZone = LBound(tgVehicleInfo(ilVefInx).sZone) To UBound(tgVehicleInfo(ilVefInx).sZone) Step 1
                        If Trim$(tgVehicleInfo(ilVefInx).sZone(ilZone)) = slZone Then
                            If tgVehicleInfo(ilVefInx).sFed(ilZone) <> "*" Then
                                slZone = tgVehicleInfo(ilVefInx).sZone(tgVehicleInfo(ilVefInx).iBaseZone(ilZone))
                                ilLocalAdj = tgVehicleInfo(ilVefInx).iLocalAdj(ilZone)
                                ilZoneFound = True
                            End If
                            Exit For
                        End If
                    Next ilZone
                    For ilZone = LBound(tgVehicleInfo(ilVefInx).sZone) To UBound(tgVehicleInfo(ilVefInx).sZone) Step 1
                        If tgVehicleInfo(ilVefInx).sFed(ilZone) = "*" Then
                            ilNumberAsterisk = ilNumberAsterisk + 1
                        End If
                    Next ilZone
                End If
                If (Not ilZoneFound) And (ilNumberAsterisk <= 1) Then
                    slZone = ""
                End If
                ilLocalAdj = -1 * ilLocalAdj
                llSpotTime = llTime + 3600 * ilLocalAdj
                If llSpotTime < 0 Then
                    llSpotTime = llSpotTime + 86400
                    llDate = llDate - 1
                ElseIf llSpotTime > 86400 Then
                    llSpotTime = llSpotTime - 86400
                    llDate = llDate + 1
                End If
                If ilDACode = 2 Then  'Tape/CD
                    llDate = DateValue(tlAstInfo.sFeedDate)
                    llSpotTime = gTimeToLong(tlAstInfo.sFeedTime, False)
                End If
                'Test if within Program Schedule
                slDate = Format$(llDate, "m/d/yy")
                ilWeekDay = Weekday(slDate)
                If tgStatusTypes(gGetAirStatus(tlAstInfo.iPledgeStatus)).iPledged <> 2 Then
                    blFoundNC = False
                    For ilLoopOnRadarHdr = LBound(tlRadar_HdrInfo) To UBound(tlRadar_HdrInfo) - 1
                        llStartDetailInx = tlRadar_HdrInfo(ilLoopOnRadarHdr).lStartInx
                        llEndDetailInx = tlRadar_HdrInfo(ilLoopOnRadarHdr).lEndInx
                        If llStartDetailInx >= 0 Then       'Date: 8/5/2019 added to fix subscript out of range issue   FYM
                            For llTest = llStartDetailInx To llEndDetailInx Step 1
                                DoEvents
                                ilIncludeSpot = True
                                If (llSpotTime < tgRadarDetailInfo(llTest).lStartTime) Or (llSpotTime > tgRadarDetailInfo(llTest).lEndTime) Then
                                    ilIncludeSpot = False
                                End If
                                If ilIncludeSpot Then
                                    Select Case tgRadarDetailInfo(llTest).sDayType
                                        Case "MF"
                                            If (ilWeekDay = vbSaturday) Or (ilWeekDay = vbSunday) Then
                                                ilIncludeSpot = False
                                            End If
                                        Case "Mo"
                                            If ilWeekDay <> vbMonday Then
                                                ilIncludeSpot = False
                                            End If
                                        Case "Tu"
                                            If ilWeekDay <> vbTuesday Then
                                                ilIncludeSpot = False
                                            End If
                                        Case "We"
                                            If ilWeekDay <> vbWednesday Then
                                                ilIncludeSpot = False
                                            End If
                                        Case "Th"
                                            If ilWeekDay <> vbThursday Then
                                                ilIncludeSpot = False
                                            End If
                                        Case "Fr"
                                            If ilWeekDay <> vbFriday Then
                                                ilIncludeSpot = False
                                            End If
                                        Case "Sa"
                                            If ilWeekDay <> vbSaturday Then
                                                ilIncludeSpot = False
                                            End If
                                        Case "Su"
                                            If ilWeekDay <> vbSunday Then
                                                ilIncludeSpot = False
                                            End If
                                    End Select
                                End If
                                If ilIncludeSpot Then
                                    'slNC = tgRadarHdrInfo(ilLoopOnRadarHdr).sNetworkCode
                                    slNC = tlRadar_HdrInfo(ilLoopOnRadarHdr).sNetworkCode       '5-27-14 was using wrong array
                                    blFoundNC = True
                                    Exit For
                                End If
                            Next llTest
                        End If          'llStartDetailInx >= 0
                        If blFoundNC Then
                            Exit For
                        End If
                    Next ilLoopOnRadarHdr
                Else
                    ilIncludeSpot = False
                End If
            End If
        End If
        On Error GoTo 0
        gObtainRadarNetworkCode = slNC          'return the Network code
        Exit Function
        
ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "modrptSubs-gObtainRadarNetworkCode"
        gObtainRadarNetworkCode = slNC
        Exit Function
End Function
'
'               Clear AFR Temporary prepass table used in various spot reports
'               Return:  true if good removal
'
Public Function gClearAFR(Form As Form) As Boolean
Dim temp_rst As ADODB.Recordset
Dim llLoopOnCount As Long

     On Error GoTo gClearAFRErr:
     gClearAFR = True
     lgRptSTime1 = timeGetTime          'for debug
    
     SQLQuery = "Select Count(afrastcode) FROM afr "
     SQLQuery = SQLQuery & "WHERE (afrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
     Set temp_rst = gSQLSelectCall(SQLQuery)
     
     If Not temp_rst.EOF Then
         If temp_rst(0).Value > 0 Then
             llLoopOnCount = temp_rst(0).Value
         Else
             llLoopOnCount = 0
         End If
     Else
         llLoopOnCount = 0
     End If
     
     Do While llLoopOnCount > 0
'         SQLQuery = "delete from afr where datepart(dayofyear, afrGenDate)+afrGenTime+afrAstCode In (Select Top 10000 datepart(dayofyear,afrGenDate)+afrGenTime+afrAstCode From afr "
'         SQLQuery = SQLQuery & "Where afrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
          SQLQuery = "delete from afr where afrAstCode In (Select Top 10000 afrAstCode From afr "
          SQLQuery = SQLQuery & "Where afrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
          SQLQuery = SQLQuery & " and (afrGenDate = '" & Format$(sgGenDate, sgSQLDateForm) & "' " & "and afrGenTime = '" & Round(Trim$(Str$(CLng(gTimeToCurrency(sgGenTime, False))))) & "')"
        If gSQLWaitNoMsgBox(SQLQuery, False) <> 0 Then
            '6/13/16: Replaced GoSub
            'GoSub gClearAFRErr:
            Screen.MousePointer = vbDefault
            gHandleError "AffErrorLog.txt", "modRptSubs-gClearAFR"
            gClearAFR = False
            On Error Resume Next
            temp_rst.Close
            Exit Function
         End If
         llLoopOnCount = llLoopOnCount - 10000
     Loop
     
     'for debugging
     lgRptETime1 = timeGetTime
     lgRptTtlTime1 = (lgRptETime1 - lgRptSTime1)
     
     temp_rst.Close
     Exit Function
     
gClearAFRErr:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modRptSubs-gClearAFR"
    gClearAFR = False
    Exit Function
End Function
'
'                   Basic copy  gBuildASTSpotsByStatus, but do not create a prepass and go straight to an export csv file
'                   Selection by Vehicle and Stations, No advertiser selection
'                   gBuildASTForStationComp:
'                   <input>  hlAST - AST file handle
'                            hlExportFile - CSV file handle
'                            sStartDate - requested start date for export AST data
'                            sEndDate - requested end date for export AST Data
'                            blUseAirDate - true to use air date vs feed date
'                            tlVehAff - vehicle list box with selected vehicles
'                            tlLbcStation - station list box with selected stations
'                            tlStatusOptions - record containing the spot status selectivity
'                            slExportLogName - error logging file
'                   <return> blExportOK = true or false
'                   Created is csv file named "Station Comp mmddyy-mmddyy"
'
Public Function gBuildAstForStationComp(hlAst As Integer, hlExportFile As Integer, sStartDate As String, sEndDate As String, blUseAirDAte As Boolean, tlVehAff As control, tlLbcStations As control, tlStatusOptions As STATUSOPTIONS, Optional slExportLogName As String = "ExportStationComp.txt") As Boolean
    Dim sSDate As String
    Dim sEDate As String                   '11-30-12
    Dim sNextWkSDate As String
    Dim sNextWkEDate As String
    Dim iNoWeeks As Integer
    Dim iLoop As Integer
    Dim iVef As Integer
    Dim iRet As Integer
    Dim iAdfCode As Integer
    Dim llLoopAST As Long
    Dim llCpttCount As Long
    'ReDim ilusecodes(1 To 1) As Integer
    ReDim ilusecodes(0 To 0) As Integer
    Dim ilIncludeCodes As Integer
    Dim sOption As String
    Dim slRecord As String

    Dim ilFoundStation As Integer
    Dim ilTemp As Integer
    '9-2-06 next 5 fields for regional copy
    Dim sRCart As String
    Dim sRProd As String
    Dim sRISCI As String
    Dim sRCreative As String
    Dim lRCrfCsfCode As Long
    Dim lRCrfCode As Long
    Dim ilStatus As Integer
    Dim ilStatusOK As Boolean
    Dim slMGReplAdfName As String
    Dim slMGReplProdName As String
    Dim llMGReplAirTime As Long
    Dim slMGReplAirTime As String
    Dim slMGReplAirDate As String
    Dim llMGReplContract As Long
    Dim ilMissedMnfCode As Integer
    Dim slLinkStatus As String * 1
    Dim slMGReplISCI As String
    Dim blProcessWeek As Boolean                    '11-30-12
    Dim llUpper As Long
    Dim llStartASTInfoSort As Long
    Dim llEndASTInfoSort As Long
    Dim llExistIndex As Long
    Dim ilExistIndex As Integer
    Dim slStr As String
    Dim ilShttRet As Integer
    Dim ilVefCode As Integer
    Dim ilVefInx As Integer
    Dim slAttPledgeType As String
    Dim slNC As String
    Dim ilHowManyCPTT As Integer
    Dim ilCreateASTInfo As Integer
    Dim llTemp As Long
    Dim ilPreviousAdfCode As Integer
    Dim slMGReplCallLetters As String
    Dim slDelimiter As String * 4
    Dim slTemp As String
    Dim blAnyErrorFound As Boolean
    Dim llCPCountThisVehicle As Long
    Dim slCompSelected As String
    
   '****** fields to create spot record
    Dim slVehicleID As String
    Dim slVehicleName As String
    Dim slPermStationID As String
    Dim slCallLetters As String
    Dim slAttCode As String
    Dim slCompFlag As String * 1
    Dim slPledgeDay As String * 3
    Dim slPledgeDate As String
    Dim slPledgeStartTime As String
    Dim slPledgeEndTime As String
    Dim slDayAired As String * 3
    Dim slDateAired As String
    Dim slTimeAired As String
    Dim slAdfName As String
    Dim slProd As String
    Dim slISCI As String
    Dim slContract As String
    Dim slLen As String
    Dim slStatus As String * 1
    Dim slRepAdvtProd As String
    Dim slRepContract As String
    Dim slRepISCI As String
    Dim slMissedDate As String
    Dim slMissedTime As String
    Dim slDaysOfWk As String * 21
    Dim blExportOK As Boolean
    Dim llTotalNoRecs As Long
    Dim llProcessedNoRecs As Long
    Dim llPercent As Long
    '
    '***** end of spot data
    '
    If (sgSQLTrace = "Y") And (hgSQLTrace >= 0) Then
        gLogMsgWODT "W", hgSQLTrace, "Export Start Time: " & gNow()
    End If
    
    bgTaskBlocked = False
    sgTaskBlockedName = sgReportListName
       
    ReDim tmAstInfo(0 To 0) As ASTINFO      'init in case tmastinfo should not be created
    slDaysOfWk = "MonTueWedThuFriSatSun"
    slDelimiter = """"                          'quote
    blExportOK = True
    blAnyErrorFound = False
    sSDate = sStartDate                         '3-4-13 the date has already been backed up to previous week from the calling rtn, use the start date that came into rtn
    iNoWeeks = (DateValue(gAdjYear(sEndDate)) - DateValue(gAdjYear(sSDate))) \ 7 + 1        '12-11-09 (doesnt get all the weeks)
    
    gObtainCodes tlLbcStations, ilIncludeCodes, ilusecodes()        'build array of which codes to incl/excl
    
    slCompSelected = ""
    If tlStatusOptions.iCompBarter = True And tlStatusOptions.iCompPayStation = True And tlStatusOptions.iCompPayNetwork = True Then
        'no need to do any testing of this selectivity
        slCompSelected = slCompSelected
    Else
        If tlStatusOptions.iCompBarter = True Then
            slCompSelected = "0"
        End If
        If tlStatusOptions.iCompPayStation = True Then
            If Trim$(slCompSelected) = "" Then
                slCompSelected = "1"
            Else
                slCompSelected = slCompSelected & ",1"
            End If
        End If
        If tlStatusOptions.iCompPayNetwork = True Then
            If Trim$(slCompSelected) = "" Then
                slCompSelected = "2"
            Else
                slCompSelected = slCompSelected & ",2"
            End If
        End If
        slCompSelected = " And attComp IN (" & slCompSelected & ")"
    End If
    
'debug:  timing parameters
'lgTtlTime1 = 0
'lgTtlTime2 = 0
'lgTtlTime3 = 0
'lgTtlTime5 = 0
'lgTtlTime6 = 0
'lgTtlTime7 = 0

    iAdfCode = -1                            'all advertisers
    ilCreateASTInfo = True
    llCpttCount = 0
    llTotalNoRecs = tlVehAff.ListCount * iNoWeeks
    llProcessedNoRecs = 0
    frmExportStationComp!plcGauge.Visible = True
    frmExportStationComp!lacProgress.Visible = True
    For iLoop = 1 To iNoWeeks Step 1
        sEDate = DateAdd("d", 6, sSDate)
        sNextWkSDate = DateAdd("d", 7, sSDate)      'get the start/enddates of following week
        sNextWkEDate = DateAdd("d", 6, sNextWkSDate)
        
        For iVef = 0 To tlVehAff.ListCount - 1 Step 1       'always loop on vehicle
            llProcessedNoRecs = llProcessedNoRecs + 1
            llPercent = (llProcessedNoRecs * CSng(100)) / llTotalNoRecs
            If llPercent >= 100 Then
                llPercent = 100
            End If
            frmExportStationComp!plcGauge.Value = llPercent
            DoEvents
            frmExportStationComp!lacProgress.Caption = "Processing Vehicle " & llProcessedNoRecs & " of " & llTotalNoRecs

            If tlVehAff.Selected(iVef) Then
                ilVefCode = tlVehAff.ItemData(iVef)
            
                On Error GoTo gBuildASTForStationCompCPTTErr:
                
                SQLQuery = "SELECT count(*) as CPCount "
                SQLQuery = SQLQuery + " FROM shtt, cptt, att"
                SQLQuery = SQLQuery + " WHERE (ShttCode = cpttShfCode"
                SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"

                'always select the required cptts for the vehicle.  No advertiser testing
                'since the filter will be in SQL statement going to Crystal
                sOption = " AND cpttVefCode = " & tlVehAff.ItemData(iVef)
                
                SQLQuery = SQLQuery & sOption
                SQLQuery = SQLQuery & slCompSelected

                SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sSDate, sgSQLDateForm) & "')"
                
                'debugging, to see how many cptts to process
                Set cprst = gSQLSelectCall(SQLQuery)
                If Not cprst.EOF Then
                    If cprst.Fields("CPCount").Value > 0 Then
                        llCPCountThisVehicle = cprst.Fields("CPCount").Value
                        gLogMsg tlVehAff.List(iVef) & "  " & Trim$(Str(llCPCountThisVehicle)) & " agreements", slExportLogName, False
                    Else
                        gLogMsg tlVehAff.List(iVef) & "  " & Trim$(Str(0)) & " agreements", slExportLogName, False
                    End If
                End If
                
                
                SQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttCode, cpttvefCode, cpttStatus, cpttPostingStatus, cpttAstStatus,attPledgeType, attComp, attTimeType"
                SQLQuery = SQLQuery + " FROM shtt, cptt, att"
                SQLQuery = SQLQuery + " WHERE (ShttCode = cpttShfCode"
                SQLQuery = SQLQuery + " AND attCode = cpttAtfCode"

                'always select the required cptts for the vehicle.  No advertiser testing
                'since the filter will be in SQL statement going to Crystal
                sOption = " AND cpttVefCode = " & tlVehAff.ItemData(iVef)
                
                SQLQuery = SQLQuery & sOption
                SQLQuery = SQLQuery & slCompSelected

                SQLQuery = SQLQuery & " AND cpttStartDate = '" & Format$(sSDate, sgSQLDateForm) & "')"
                Set cprst = gSQLSelectCall(SQLQuery)
                If blAnyErrorFound Then
                    gCloseRegionSQLRst
                    bgTaskBlocked = False
                    sgTaskBlockedName = ""
                    Erase tmAstInfo
                    Erase tmAstInfoSort
                    Erase tmCPDat
                    Erase ilusecodes
                    gBuildAstForStationComp = blExportOK
                    Exit Function
                End If
                On Error GoTo 0
                While Not cprst.EOF
                    ReDim tgCPPosting(0 To 1) As CPPOSTING

                    tgCPPosting(0).lCpttCode = cprst!cpttCode
                    tgCPPosting(0).iStatus = cprst!cpttStatus
                    tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
                    tgCPPosting(0).lAttCode = cprst!cpttatfCode
                    tgCPPosting(0).iAttTimeType = cprst!attTimeType
                    tgCPPosting(0).iVefCode = cprst!cpttvefcode
                    tgCPPosting(0).iShttCode = cprst!shttCode
                    tgCPPosting(0).sZone = cprst!shttTimeZone
                    tgCPPosting(0).sDate = Format$(sSDate, sgShowDateForm)
                    tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
                    slAttPledgeType = cprst!attPledgeType
                    'determine valid station for inclusion
                    ilFoundStation = gTestIncludeExclude(cprst!shttCode, ilIncludeCodes, ilusecodes())

                    blProcessWeek = True                'default case
'10-8-14; no need to test spot in different week, we now have key by air date if using trueair dates (vs feed dates)
'                    If (iLoop = 1) And (blUseAirDAte) And (ilFoundStation) Then          'if week 1, determine if theres any pledges that indicate spot will run in the following week (pledged After) for a station that has been selected
'                        'advt option or not, see if there are any pledges defined to air following week (AFter)
'                        blProcessWeek = False
'                        'look at DAT (pledge times to see if any pledges not in the same week
'                        SQLQuery = "Select count(*) as PledgeCountAfter from dat where datatfCode = " & cprst!cpttatfCode & " and datPdDayFed = " & "'A'"
'                        Set Advrst = gSQLSelectCall(SQLQuery)
'                        If Not Advrst.EOF Then
'                            If Advrst.Fields("PledgeCountAfter").Value > 0 Then
'                                blProcessWeek = True
'                                Advrst.Close
'                            End If
'                        End If
'                        If Not blProcessWeek Then           'no pledges for "A" (pledge after), see if theres any spots that user posted in different week than fed
'                            SQLQuery = "Select count(*) as SpotCountPostedNextWeek from ast where astatfcode = " & cprst!cpttatfCode & "  and astAirDate >= '" & Format$(sNextWkSDate, sgSQLDateForm) & "' and astAirDate <= '" & Format$(sNextWkEDate, sgSQLDateForm) & "' and "
'                            SQLQuery = SQLQuery & " astfeeddate < '" & Format$(sNextWkSDate, sgSQLDateForm) & "'"
'                            Set Advrst = gSQLSelectCall(SQLQuery)
'                            If Not Advrst.EOF Then
'                                If Advrst.Fields("SpotCountPostedNextWeek").Value > 0 Then
'                                    blProcessWeek = True
'                                    Advrst.Close
'                                End If
'                            End If
'                        End If
'                    End If
                    
                    'found a valid station with a matching advt (if not by advt, its always set to process), and its a valid week
                    If (ilFoundStation) And (blProcessWeek) Then
                        llCpttCount = llCpttCount + 1       'debugging only
                        'Create AST records
                        igTimes = 1 'By Week
                        iRet = gGetAstInfo(hlAst, tmCPDat(), tmAstInfo(), iAdfCode, True, True, ilCreateASTInfo, , , , blUseAirDAte)    'remaining parameters are optional and defaulted per routine
                                                    
                        'sort the spots returned for this vehicle/station
                        'create the sort key from tmAstInfo array
                        llUpper = UBound(tmAstInfo)
                        ReDim tmAstInfoSort(0 To llUpper) As ASTSORTKEY
                       
                        For llLoopAST = 0 To UBound(tmAstInfo) - 1
                            slStr = Trim$(Str$(tmAstInfo(llLoopAST).iAdfCode))
                            Do While Len(slStr) < 6
                                slStr = "0" & slStr
                            Loop
                            tmAstInfoSort(llLoopAST).sKey = Trim$(slStr)

                            slStr = Trim$(Str$(tmAstInfo(llLoopAST).lCntrNo))
                            Do While Len(slStr) < 9
                                slStr = "0" & slStr
                            Loop
                            tmAstInfoSort(llLoopAST).sKey = Trim$(tmAstInfoSort(llLoopAST).sKey) & "|" & Trim$(slStr)

                            tmAstInfoSort(llLoopAST).lCode = tmAstInfo(llLoopAST).lCode     'ast spot id
                            tmAstInfoSort(llLoopAST).lIndex = llLoopAST
                            tmAstInfoSort(llLoopAST).iAdfCode = tmAstInfo(llLoopAST).iAdfCode

                        Next llLoopAST
                        
                        If llUpper > 0 Then
                            ArraySortTyp fnAV(tmAstInfoSort(), 0), llUpper, 0, LenB(tmAstInfoSort(0)), 0, LenB(tmAstInfoSort(0).sKey), 0
                        End If
                                        
                        llStartASTInfoSort = LBound(tmAstInfoSort)
                        llEndASTInfoSort = UBound(tmAstInfoSort)
                
                        For llUpper = llStartASTInfoSort To llEndASTInfoSort - 1
                            llLoopAST = tmAstInfoSort(llUpper).lIndex
                            If llUpper = llStartASTInfoSort Then            '1st time thru this agreement
                                'get and convert all infomation that is unique to this agreement for all spots
                                llExistIndex = gBinarySearchAdf(CLng(tmAstInfo(llLoopAST).iAdfCode))
                                slAdfName = ""
                                If llExistIndex <> -1 Then
                                    slAdfName = Trim$(tgAdvtInfo(llExistIndex).sAdvtName)
                                End If
                                
                                'agreement, convert to string
                                slAttCode = Trim$(Str$(tmAstInfo(llLoopAST).lAttCode))
                               
                                'vehicle name
                                ilExistIndex = gBinarySearchVef(CLng(tmAstInfo(llLoopAST).iVefCode))
                                slVehicleName = Trim$(tgVehicleInfo(ilExistIndex).sVehicle)
                                'vehicle ID
                                slVehicleID = Trim$(Str$(tgVehicleInfo(ilExistIndex).iCode))
                                
                                'station call letters
                                llExistIndex = gBinarySearchStationInfoByCode(tmAstInfo(llLoopAST).iShttCode)
                                slCallLetters = Trim$(tgStationInfoByCode(CSng(llExistIndex)).sCallLetters)
                                'station Permanent ID
                                slPermStationID = Trim$(Str$(tgStationInfoByCode(CSng(llExistIndex)).lPermStationID))
                                
                                'agreement compensation flag
                                slCompFlag = "B"                    'barter
                                If tmAstInfo(llLoopAST).iComp = 1 Then      'pay station
                                    slCompFlag = "S"
                                ElseIf tmAstInfo(llLoopAST).iComp = 2 Then      'pay network
                                    slCompFlag = "N"
                                End If
                              
                                ilPreviousAdfCode = tmAstInfo(llLoopAST).iAdfCode       'save to compare against each record, retrieve name when it changes
                           End If
                           
                           If ilPreviousAdfCode <> tmAstInfo(llLoopAST).iAdfCode Then
                                llExistIndex = gBinarySearchAdf(CLng(tmAstInfo(llLoopAST).iAdfCode))
                                slAdfName = ""
                                If llExistIndex <> -1 Then
                                    slAdfName = Trim$(tgAdvtInfo(llExistIndex).sAdvtName)
                                End If
                                ilPreviousAdfCode = tmAstInfo(llLoopAST).iAdfCode
                           End If
                            
                            slMGReplAdfName = ""
                            slMGReplProdName = ""
                            llMGReplAirTime = 0
                            slMGReplAirDate = ""
                            slMGReplAirTime = ""
                            llMGReplContract = 0
                            ilMissedMnfCode = 0
                            slLinkStatus = ""
                            slMGReplISCI = ""
                            ilStatus = gGetAirStatus(tmAstInfo(llLoopAST).iStatus)
                            'if excluding unresolved missed, need to include those that are associated with mg and/or replacements if they are also included
                            ilStatusOK = False
                            
                            slStatus = "A"                                   'spot status to update in output
                            If tmAstInfo(llLoopAST).iCPStatus = 0 Then      'not reported?
                                If tlStatusOptions.iNotReported Then        'include not reported
                                    ilStatusOK = True
                                End If
                            ElseIf ilStatus = 0 Then                        'live
                                If tlStatusOptions.iInclLive0 Then
                                    ilStatusOK = True
                                End If
                            ElseIf ilStatus = 1 Then                        'delayed aired
                                If tlStatusOptions.iInclDelay1 Then
                                    ilStatusOK = True
                                End If
                            ElseIf ilStatus = 6 Then
                                If tlStatusOptions.iInclAirOutPledge6 Then      'aired outside pledge
                                    ilStatusOK = True
                                End If
                            ElseIf ilStatus = 7 Then
                                If tlStatusOptions.iInclAiredNotPledge7 Then      'aired, not pledge
                                    ilStatusOK = True
                                End If
                            ElseIf ilStatus = 8 Then
                                If tlStatusOptions.iInclNotCarry8 Then      'not carried
                                    ilStatusOK = True
                                End If
                            ElseIf ilStatus = 9 Then
                                If tlStatusOptions.iInclDelayCmmlOnly9 Then      'delay, air comml only
                                    ilStatusOK = True
                                End If
                            ElseIf ilStatus = 10 Then
                                If tlStatusOptions.iInclAirCmmlOnly10 Then      'live, air comml only
                                    ilStatusOK = True
                                End If

                            ElseIf ((ilStatus = ASTEXTENDED_MG) And (tlStatusOptions.iInclMG11)) Or ((ilStatus = ASTEXTENDED_REPLACEMENT) And (tlStatusOptions.iInclRepl13)) Then
                                 ilStatusOK = True
                                 SQLQuery = "SELECT  astAirDate, astAirTime, astcode, astStatus, astlkastcode, astshfcode, astcntrno, astcpfcode, cpfISCI, adfname  From ast "
'                                 SQLQuery = SQLQuery & " inner join adf_Advertisers on " & tmAstInfo(llLoopAST).iAdfCode & " = adfcode  left outer join cpf_Copy_Prodct_ISCI on " & tmAstInfo(llLoopAST).lCpfCode & " = cpfcode "
                                 SQLQuery = SQLQuery & " inner join adf_Advertisers on astAdfCode  = adfcode  left outer join cpf_Copy_Prodct_ISCI on astCpfCode = cpfcode "
'                                 SQLQuery = SQLQuery & "Where " & tmAstInfo(llLoopAST).lLkAstCode & " = " & tmAstInfo(llLoopAST).lCode
                                 SQLQuery = SQLQuery & "Where astcode = " & tmAstInfo(llLoopAST).lLkAstCode
                                 
                                 '4/22/18
                                 'Set tmAltForAST = cnn.Execute(SQLQuery)         'read the associated ALT (associations) for the spot
                                 Set tmAltForAST = gSQLSelectCall(SQLQuery)
                                 ilStatusOK = True
                                 While Not tmAltForAST.EOF
                                     'find the associated missed for this replacement or makegood spot.  If it's a reference it should always be shown.
                                     If tmAltForAST!astLkAstCode = tmAstInfo(llLoopAST).lCode Then         'missed side, get the mnf missed reference
'                                                slMGReplISCI = ""               '12-24-13  tmAltForAST!altAiredISCI
'                                            Else
                                         slMGReplAirDate = Format$(tmAltForAST!astAirDate, sgShowDateForm)
                                         llMGReplAirTime = gTimeToLong(tmAltForAST!astAirTime, False)
                                         slMGReplAirTime = Format$(tmAltForAST!astAirTime, sgShowTimeWSecForm)
                                         If ilStatus = ASTEXTENDED_REPLACEMENT Then
                                             slMGReplAdfName = Trim$(tmAltForAST!adfName)
                                            
                                             llMGReplContract = tmAltForAST!astCntrNo
                                             slMGReplISCI = tmAltForAST!cpfISCI
                                             
                                             slMGReplAirDate = ""
                                             slMGReplAirTime = ""
                                        End If
                                         ilShttRet = gBinarySearchStationInfoByCode(tmAltForAST!astShfCode)
                                             If ilShttRet <> -1 Then
                                                 slCallLetters = Trim$(tgStationInfoByCode(ilShttRet).sCallLetters)
                                             End If
                                         If ilStatus = ASTEXTENDED_MG Then
                                            '2-10-15 slMGReplAirDate = missed date; slMGReplAirTime : missed time
                                             slLinkStatus = "M"          'mg
                                             slStatus = "M"
                                             slMGReplAdfName = ""       'no adv name required in replacement adv name field
                                         Else
                                             slLinkStatus = "R"          'replacement
                                             slStatus = "R"
                                         End If
                                     End If
                                     ilStatus = ilStatus
                                
                                     tmAltForAST.MoveNext
                                 Wend

                            ElseIf (ilStatus >= 2 And ilStatus <= 5) Or (ilStatus = ASTAIR_MISSED_MG_BYPASS) Then              '4-12-17 Not aired spot or missed mg bypassed spot
                                slStatus = "N"              'not aired
                                If (ilStatus = 2 And tlStatusOptions.iInclMissed2 = True) Or (ilStatus = 3 And tlStatusOptions.iInclMissed3 = True) Or (ilStatus = 4 And tlStatusOptions.iInclMissed4 = True) Or (ilStatus = 5 And tlStatusOptions.iInclMissed5 = True) Or (ilStatus = 14 And tlStatusOptions.iInclMissedMGBypass14 = True) Then
                                    ilStatusOK = True
                                    ilMissedMnfCode = tmAstInfo(llLoopAST).iMissedMnfCode           'missed reason
                                    'is there an associated mg or replacement for the missed spot just found?  It has to have a link to the ast spot that it references.  If so, should the associated mg or replacement spot be included in the report?
                                    If (tmAstInfo(llLoopAST).lLkAstCode > 0) Then       'if there is a pointer, this is a resolved missed, ignore it
                                        ilStatusOK = False
                                        'there is a mg or replacement
                                        'get the associated mg/replacement to see what it is
'                                        SQLQuery = "Select astCode, astshfcode, astStatus, astAirDate, astAirTime, adfcode, adfName, cpfcode, cpfName  from ast  "
'                                        SQLQuery = SQLQuery & " inner join adf_Advertisers on astadfcode = adfcode "
'                                        SQLQuery = SQLQuery & " left outer join CPF_Copy_Prodct_ISCI on astcpfcode = cpfcode "
'                                        SQLQuery = SQLQuery & " where astcode = " & tmAstInfo(llLoopAST).lLkAstCode
'                                        Set tmASTForMGRepl = gSQLSelectCall(SQLQuery)
'                                        While Not tmASTForMGRepl.EOF
'                                            If ((tlStatusOptions.iInclRepl13 And tmASTForMGRepl!astStatus Mod 100 = ASTEXTENDED_REPLACEMENT) Or (tlStatusOptions.iInclMG11 And tmASTForMGRepl!astStatus Mod 100 = ASTEXTENDED_MG)) Then
'                                                'makegood and/or replacement date & time for the missed spot
'                                                'slMGReplAdfName = Trim$(tmASTForMGRepl!adfName)
''                                                   slMGReplProdName = Trim$(tmAltForAST!lstProd)
'                                                slMGReplAirDate = Format$(tmASTForMGRepl!astAirDate, sgShowDateForm)
'                                                llMGReplAirTime = gTimeToLong(tmASTForMGRepl!astAirTime, False)
'                                                ilShttRet = gBinarySearchStationInfoByCode(tmASTForMGRepl!astShfCode)
'                                                If ilShttRet <> -1 Then
'                                                    slCallLetters = Trim$(tgStationInfoByCode(ilShttRet).sCallLetters)
'                                                End If
'                                                If tmASTForMGRepl!astStatus = ASTEXTENDED_MG Then
'                                                    slLinkStatus = "M"
'                                                ElseIf tmASTForMGRepl!astStatus = ASTEXTENDED_REPLACEMENT Then
'                                                    slLinkStatus = "R"
'                                                    slMGReplAdfName = Trim$(tmASTForMGRepl!adfName)
'                                                    If Not IsNull(tmASTForMGRepl!cpfName) Then
'                                                        slMGReplProdName = Trim$(tmASTForMGRepl!cpfName)
'                                                    End If
'                                                End If
'                                            End If
'                                            tmASTForMGRepl.MoveNext
'                                        Wend
                                    End If
                                'there is no option to exclude Missed spots, always include them
                                
                                Else                    'missed not selected for inclusion, but still need to include it if it has a mg or replacement that is to be shown
                                    ilStatusOK = False
'                                    If tmAstInfo(llLoopAST).lLkAstCode > 0 And Not tlStatusOptions.iInclResolveMissed Then       'include the resolved misses
'                                        ilStatusOK = False
'                                    Else          'a mg or replacement is defined
'                                        'read the mg/replace spot then see if they should be included.  if so, need to bring in that associated missed spot
'
'                                        SQLQuery = "SELECT  adfName, astAirDate, astAirTime, astcode, astStatus  From ast "
'                                        SQLQuery = SQLQuery & " INNER JOIN adf_advertisers on astadfcode = adfcode "
'                                        SQLQuery = SQLQuery & "Where astlkAstCode  = " & tmAstInfo(llLoopAST).lCode
'
'                                        Set tmASTForMGRepl = gSQLSelectCall(SQLQuery)         'read the associated ALT (associations) for the spot
'                                        ilStatusOK = False      'True
'                                        While Not tmASTForMGRepl.EOF
'                                            'is there an associated mg or replacement for the missed spot just found?  It has to have a link to the ast spot that it references.  If so, should the associated mg or replacement spot be included in the report?
'                                            If ((tlStatusOptions.iInclRepl13 And tmASTForMGRepl!astStatus Mod 100 = ASTEXTENDED_REPLACEMENT) Or (tlStatusOptions.iInclMG11 And tmASTForMGRepl!astStatus Mod 100 = ASTEXTENDED_MG)) Then
'                                                ilStatusOK = True
'                                                slMGReplAdfName = Trim$(tmASTForMGRepl!adfName)
'                                                'slMGReplProdName = Trim$(tmAltForAST!lstProd)
'                                                slMGReplAirDate = Format$(tmASTForMGRepl!astAirDate, sgShowDateForm)
'                                                llMGReplAirTime = gTimeToLong(tmASTForMGRepl!astAirTime, False)
'                                                If tmASTForMGRepl!astStatus = ASTEXTENDED_REPLACEMENT Then
'                                                    slLinkStatus = "R"
'                                                Else
'                                                    slLinkStatus = "M"
'                                                End If
'                                            End If
'
'                                            tmASTForMGRepl.MoveNext
'                                        Wend
'                                    End If
                                End If
   
    
                            ElseIf ilStatus = ASTEXTENDED_BONUS Then
                                slStatus = "B"
                                If tlStatusOptions.iInclBonus12 Then
                                    ilStatusOK = True
                                End If
                            Else            'all other status N/G
                                ilStatusOK = False
                            End If
                            
                            'test for pledging not to feed; if so, see if the spot was posted as not carried too
                            If (tmAstInfo(llLoopAST).iPledgeStatus = 4 Or tmAstInfo(llLoopAST).iPledgeStatus = 8) And (gGetAirStatus(tmAstInfo(llLoopAST).iStatus) = 8) Then
                                ilStatusOK = False
                            End If
                            If ilStatusOK Then
                                sRCart = ""
                                sRProd = ""
                                sRISCI = ""
                                sRCreative = ""
                                lRCrfCsfCode = 0
                                
                                slISCI = Trim$(tmAstInfo(llLoopAST).sISCI)
                                slProd = Trim$((tmAstInfo(llLoopAST).sProd))
                                
                                If tmAstInfo(llLoopAST).iRegionType > 0 Then
                                    sRCart = Trim$(tmAstInfo(llLoopAST).sRCart)
                                    If Trim$(tmAstInfo(llLoopAST).sRProduct) <> "" Then
                                        slProd = Trim$(tmAstInfo(llLoopAST).sRProduct)
                                    End If
                                    'sRISCI = Trim$(tmAstInfo(llLoopAST).sRISCI)
                                    If Trim$(tmAstInfo(llLoopAST).sRISCI) <> "" Then        'use regional if it exists
                                        slISCI = Trim$(tmAstInfo(llLoopAST).sRISCI)
                                    End If
                                    sRCreative = Trim$(tmAstInfo(llLoopAST).sRCreativeTitle)
                                    lRCrfCsfCode = Trim$(tmAstInfo(llLoopAST).lRCrfCsfCode)
                                End If
                                
                                'Vehicle ID (Internal code)
                                slRecord = Trim$(slVehicleID) & ","
                                
                                'Vehicle name
                                slRecord = slRecord & Trim$(slDelimiter) & slVehicleName & Trim$(slDelimiter) & ","
                                
                                'Station ID
                                slRecord = slRecord & Trim$(slPermStationID) & ","
                                
                                'Station Call Letters
                                slRecord = slRecord & Trim$(slDelimiter) & slCallLetters & Trim$(slDelimiter) & ","
                                
                                'Internal agreement ID
                                slRecord = slRecord & slAttCode & ","
                                
                                'compensation flag
                                slRecord = slRecord & Trim$(slDelimiter) & slCompFlag & Trim$(slDelimiter) & ","
                                
                                'Pledge Day
                                llTemp = gDateValue(tmAstInfo(llLoopAST).sPledgeDate)
                                ilTemp = gWeekDayLong(llTemp)
                                slPledgeDay = Mid$(slDaysOfWk, (ilTemp * 3) + 1, 3)
                                slRecord = slRecord & Trim$(slDelimiter) & slPledgeDay & Trim$(slDelimiter) & ","

                                'Pledge Date
                                slRecord = slRecord & Trim$(tmAstInfo(llLoopAST).sPledgeDate) & ","
                                
                                'pledge start time
                                slRecord = slRecord & Trim$(tmAstInfo(llLoopAST).sPledgeStartTime) & ","
                                
                                'pledge end time, make start and end times the same if no end time
                                If Trim$(tmAstInfo(llLoopAST).sPledgeEndTime) = "" Then
                                    slRecord = slRecord & Trim$(tmAstInfo(llLoopAST).sPledgeStartTime) & ","
                                Else
                                    slRecord = slRecord & Trim$(tmAstInfo(llLoopAST).sPledgeEndTime) & ","
                                End If
                                
                                'Day aired
                                llTemp = gDateValue((tmAstInfo(llLoopAST).sAirDate))
                                ilTemp = gWeekDayLong(llTemp)
                                slDayAired = Mid$(slDaysOfWk, (ilTemp * 3) + 1, 3)
                                slRecord = slRecord & Trim$(slDelimiter) & slDayAired & Trim$(slDelimiter) & ","
                                
                                slRecord = slRecord & Trim$(tmAstInfo(llLoopAST).sAirDate) & ","
                                
                                'time aired
                                slRecord = slRecord & Trim$(tmAstInfo(llLoopAST).sAirTime) & ","
                                
                                'advt / product
                                slTemp = slAdfName
                                If Trim$(tmAstInfo(llLoopAST).sProd) <> "" Then
                                    slTemp = Trim$(slAdfName) & "/" & Trim$((tmAstInfo(llLoopAST).sProd))
                                End If
                                
                                slRecord = slRecord & Trim$(slDelimiter) & slTemp & Trim$(slDelimiter) & ","
                                
                                'ISCI, could be regional or generic
                                slRecord = slRecord & Trim$(slDelimiter) & Trim$(slISCI) & Trim$(slDelimiter) & ","
                                
                                'contract #
                                slRecord = slRecord & Trim$(Str(tmAstInfo(llLoopAST).lCntrNo)) & ","
                                
                                'spot Len
                                slRecord = slRecord & Trim$(Str(tmAstInfo(llLoopAST).iLen)) & ","
                                
                                'spot status
                                slRecord = slRecord & Trim$(slDelimiter) & slStatus & Trim$(slDelimiter) & ","
                                
                                'replaced advt prod
                                slTemp = Trim$(slMGReplAdfName)
                                If Trim$(slMGReplProdName) <> "" Then
                                    slTemp = slTemp & "/" & Trim$(slMGReplProdName)
                                End If
                                slRecord = slRecord & Trim$(slDelimiter) & slTemp & Trim$(slDelimiter) & ","
                               
                                
                                'replaced contract
                                slRecord = slRecord & Trim$(Str(llMGReplContract)) & ","
                                
                                'replaced ISCI
                                slRecord = slRecord & Trim$(slDelimiter) & Trim$(slMGReplISCI) & Trim$(slDelimiter) & ","
                                
                                'missed date
                                slRecord = slRecord & Trim$(slMGReplAirDate) & ","
                                
                                'missed time
                                slRecord = slRecord & Trim$(slMGReplAirTime) & ","
                                On Error GoTo gBuildAstForStationCompWriteErr:
                                Print #hlExportFile, slRecord
                                On Error GoTo 0
                                If blAnyErrorFound Then
                                    gCloseRegionSQLRst
                                    bgTaskBlocked = False
                                    sgTaskBlockedName = ""
                                    Erase tmAstInfo
                                    Erase tmAstInfoSort
                                    Erase tmCPDat
                                    Erase ilusecodes
                                    gBuildAstForStationComp = blExportOK
                                    Exit Function
                                End If
                            End If
  
                        Next llUpper
                    End If
                    
                    cprst.MoveNext
                Wend
                If (tlLbcStations.ListCount = 0) Or (tlLbcStations.ListCount = tlLbcStations.SelCount) Then
                    gClearASTInfo True
                Else
                    gClearASTInfo False
                End If
            End If

        Next iVef                   'loop on vehicle
        sSDate = DateAdd("d", 7, sSDate)
    Next iLoop                      'loop on weeks
    
    gCloseRegionSQLRst
    
    If (sgSQLTrace = "Y") And (hgSQLTrace >= 0) Then
        gLogMsgWODT "W", hgSQLTrace, "SQL Overall Time: " & gTimeString(lgTtlTimeSQL / 1000, True) & " cpttCount: " & llCpttCount
        gLogMsgWODT "W", hgSQLTrace, "Export Completed Time: " & gNow()
    End If
    
    If bgTaskBlocked And igReportSource <> 2 Then
         gMsgBox "Some spots were blocked during the Report generation." & vbCrLf & "Please refer to the Messages folder for file: TaskBlocked_" & sgTaskBlockedDate & ".txt for further information.", vbCritical
    End If
    
    bgTaskBlocked = False
    sgTaskBlockedName = ""
    
    Erase tmAstInfo
    Erase tmAstInfoSort
    Erase tmCPDat
    Erase ilusecodes
    gBuildAstForStationComp = blExportOK
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "modRptSubs-gBuildASTForStationComp"
    blAnyErrorFound = True
    blExportOK = False
    Resume Next
gBuildAstForStationCompWriteErr:
    gMsg = "Error writing export record"
    gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    gLogMsg gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, slExportLogName, False

    blAnyErrorFound = True
    blExportOK = False
    Resume Next
gBuildASTForStationCompCPTTErr:
    gMsg = "A SQL Error has occurred reading CPTT in gBuildAstSpotsByStatus" & Err.Description
    gMsgBox gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, vbCritical
    gLogMsg gMsg & Err.Description & "; Error #" & Err.Number & "; Line #" & Erl, slExportLogName, False
    blAnyErrorFound = True
    blExportOK = False
    Resume Next
End Function
'
'           Format Include/Exclude selection query based on "Not In" or "IN" command
'
'           gFormInclExclQuery
'           <input> = slFieldName as string
'                     ilInclExclCodes = true to include codes, false to exclude codes
'                     ilCodes() - array of codes to include/exclude
'           return = Sqlquery command
'
Public Function gFormInclExclQuery(slFieldName As String, ilInclExclCodes As Integer, ilCodes() As Integer) As String
Dim slStr As String
Dim iLoop As Integer
Dim ilCount As Integer


        slStr = ""
        For iLoop = LBound(ilCodes) To UBound(ilCodes) - 1
            If Trim$(slStr) = "" Then
                If ilInclExclCodes = True Then                          'include the list
                    slStr = " IN (" & Str(ilCodes(iLoop))
                Else                                                        'exclude the list
                    'if more than half has been excluded, blOnly1AdvtSelected flag remains false so it doesnt go thru single testing; otherwise nothing will be found
                    slStr = " Not IN (" & Str(ilCodes(iLoop))
                End If
            Else
                'has at least one entry
                slStr = slStr & "," & Str(ilCodes(iLoop))
            End If
        Next iLoop
        If Trim$(slStr) <> "" Then
            'slStr = " and " & Trim$(slFieldName) & slStr & ")"
             slStr = Trim$(slFieldName) & slStr & ")"
       End If
    
    gFormInclExclQuery = slStr
End Function

Public Function gFilterValue(ilField As Integer, ilIncludeCodes As Integer, ilusecodes() As Integer) As Boolean
Dim blFound As Boolean
Dim ilTemp As Integer
        blFound = False
        If ilIncludeCodes Then
            For ilTemp = LBound(ilusecodes) To UBound(ilusecodes) - 1 Step 1
                If ilusecodes(ilTemp) = ilField Then
                    blFound = True
                    Exit For
                End If
            Next ilTemp
        Else
            blFound = True
            For ilTemp = LBound(ilusecodes) To UBound(ilusecodes) - 1 Step 1
                If ilusecodes(ilTemp) = ilField Then
                    blFound = False
                    Exit For
                End If
            Next ilTemp
        End If
        gFilterValue = blFound
End Function
Sub gObtainCodesLong(tlListBox As control, ilIncludeCodes, llUseCodes() As Long)
Dim ilHowManyDefined As Integer
Dim ilHowMany As Integer
Dim slNameCode As String
Dim ilLoop As Integer
Dim slCode As String
Dim ilRet As Integer
    ilHowManyDefined = tlListBox.ListCount
    ilHowMany = tlListBox.SelCount
    If ilHowMany > ilHowManyDefined / 2 Then    'more than half selected
        ilIncludeCodes = False
    Else
        ilIncludeCodes = True
    End If
    For ilLoop = 0 To tlListBox.ListCount - 1 Step 1
        If tlListBox.Selected(ilLoop) And ilIncludeCodes Then               'selected ?
            llUseCodes(UBound(llUseCodes)) = tlListBox.ItemData(ilLoop)
            'ReDim Preserve llUseCodes(1 To UBound(llUseCodes) + 1)
            ReDim Preserve llUseCodes(LBound(llUseCodes) To UBound(llUseCodes) + 1)
        Else        'exclude these
            If (Not tlListBox.Selected(ilLoop)) And (Not ilIncludeCodes) Then
                llUseCodes(UBound(llUseCodes)) = tlListBox.ItemData(ilLoop)
                'ReDim Preserve llUseCodes(1 To UBound(llUseCodes) + 1)
                ReDim Preserve llUseCodes(LBound(llUseCodes) To UBound(llUseCodes) + 1)
            End If
        End If
    Next ilLoop
End Sub
'
'           gTestIncludeExclude - test array of codes (stored in llUseCodes) to determine if the value
'           should be included or excluded (determined by flag ilIncludeCodes)
'           <input>  ilValue = value to test for inclusion/exclusion
'                    ilIncludeCodes - true = include code, false = exclude code
'                    llUseCodes() - array of codes to include or exclude
'           <return> true to include, false to exclude
'
Public Function gTestIncludeExcludeLong(llValue As Long, ilIncludeCodes As Integer, llUseCodes() As Long) As Integer
Dim ilTemp As Integer
Dim blFound As Integer

    If ilIncludeCodes Then          'include the any of the codes in array?
        blFound = False
        For ilTemp = LBound(llUseCodes) To UBound(llUseCodes) - 1 Step 1
            If llUseCodes(ilTemp) = llValue Then
                blFound = True                    'include the matching vehicle
                Exit For
            End If

        Next ilTemp
    Else                            'exclude any of the codes in array?
        blFound = True
        For ilTemp = LBound(llUseCodes) To UBound(llUseCodes) - 1 Step 1
            If llUseCodes(ilTemp) = llValue Then
                blFound = False                  'exclude the matching vehicle
                Exit For
            End If
        Next ilTemp
    End If
    gTestIncludeExcludeLong = blFound
    Exit Function
End Function
'
'
'                   Populate the unique groups for selection
'                   <input>  cbcSet1 - control name to populate
'                            ilUseNone  - add Item to List Box with "None" (true/false)
'

Sub gPopVehicleGroups(cbcSet As control, ilUseNone As Integer)
Dim ilRet As Integer
Dim ilLoop As Integer
Dim ilLoop2 As Integer
Dim ilFound As Integer
Dim ilIndex As Integer

    'cbcSet1.AddItem "None"
    'ReDim ilVehGroup1(1 To 1) As Integer
    ReDim ilVehGroup1(0 To 0) As Integer
    'ReDim slVehGroup1(1 To 1) As String * 1
    ReDim slVehGroup1(0 To 0) As String * 1
    
    SQLQuery = "Select * from Mnf_Multi_Names where mnfType = 'H'"
    Set rst_Mnf = gSQLSelectCall(SQLQuery)
    While Not rst_Mnf.EOF
    '
        ilFound = False
        'For ilIndex = 1 To UBound(slVehGroup1) - 1 Step 1
        For ilIndex = LBound(slVehGroup1) To UBound(slVehGroup1) - 1 Step 1
            If Trim$(rst_Mnf!mnfUnitType) = slVehGroup1(ilIndex) Then          'look for the vehicle set # built in array
                ilFound = True
                Exit For
            End If
        Next ilIndex
        If Not ilFound Then
            slVehGroup1(ilIndex) = Trim$(rst_Mnf!mnfUnitType)          'vehicle set #
            ilVehGroup1(ilIndex) = Val(slVehGroup1(ilIndex))
            'ReDim Preserve slVehGroup1(1 To UBound(slVehGroup1) + 1)
            ReDim Preserve slVehGroup1(0 To UBound(slVehGroup1) + 1)
            'ReDim Preserve ilVehGroup1(1 To UBound(ilVehGroup1) + 1)
            ReDim Preserve ilVehGroup1(0 To UBound(ilVehGroup1) + 1)
        End If
        rst_Mnf.MoveNext
    Wend
        
    'sort the unique set #s
    'For ilLoop = 1 To UBound(ilVehGroup1) - 1
    For ilLoop = LBound(ilVehGroup1) To UBound(ilVehGroup1) - 1
        For ilLoop2 = ilLoop + 1 To UBound(ilVehGroup1) - 1
            If ilVehGroup1(ilLoop) > ilVehGroup1(ilLoop2) Then
                'swap the two
                ilFound = ilVehGroup1(ilLoop)
                ilVehGroup1(ilLoop) = ilVehGroup1(ilLoop2)
                ilVehGroup1(ilLoop2) = ilFound
            End If
        Next ilLoop2
    Next ilLoop
        cbcSet.Clear

        If ilUseNone Then
            cbcSet.AddItem "None"
            cbcSet.ItemData(cbcSet.NewIndex) = 0
        End If
        'For ilLoop = 1 To UBound(ilVehGroup1) - 1 Step 1
        For ilLoop = LBound(ilVehGroup1) To UBound(ilVehGroup1) - 1 Step 1
            'ReDim Preserve tgVehicleSets(0 To UBound(tgVehicleSets) + 1) As POPICODENAME
            If ilVehGroup1(ilLoop) = 1 Then
                cbcSet.AddItem "Participants"
                cbcSet.ItemData(cbcSet.NewIndex) = 1
            ElseIf ilVehGroup1(ilLoop) = 2 Then
                cbcSet.AddItem "Sub-Totals"
                cbcSet.ItemData(cbcSet.NewIndex) = 2
            ElseIf ilVehGroup1(ilLoop) = 3 Then
                cbcSet.AddItem "Market"
                cbcSet.ItemData(cbcSet.NewIndex) = 3
           ElseIf ilVehGroup1(ilLoop) = 4 Then
                cbcSet.AddItem "Format"
                cbcSet.ItemData(cbcSet.NewIndex) = 4
            ElseIf ilVehGroup1(ilLoop) = 5 Then
                cbcSet.AddItem "Research"
                cbcSet.ItemData(cbcSet.NewIndex) = 5
            ElseIf ilVehGroup1(ilLoop) = 6 Then
                cbcSet.AddItem "Sub-Company"
                cbcSet.ItemData(cbcSet.NewIndex) = 6
            End If
        Next ilLoop
        cbcSet.ListIndex = 0
    Exit Sub
End Sub

'
'
'           gGetVehGrpSets - Given the vehicle code, find the vehicle
'               record in tgMVef, and return the major and minor vehicle
'               group set.
'           <input> ilvefcode = Vehicle code
'                   ilMinorSet = set # that determines the minor sort
'                   ilVGSet = set # that determines the major sort
'           <output> ilMnfVGCode - mnf code for the minor sort field
'                   ilmnfVGCode - mnf code for the major sort field
'
'           Created 6/10/98 D. Hosaka
'
Sub gGetVehGrpSets(ilVefCode As Integer, ilVGSet As Integer, ilmnfVGCode As Integer)
Dim ilLoop As Integer
        ilmnfVGCode = 0
        ilLoop = gBinarySearchVef(CLng(ilVefCode))
        If ilLoop <> -1 Then
            If ilVGSet = 1 Then
                ilmnfVGCode = tgVehicleInfo(ilLoop).iOwnerMnfCode       'owner/participant
            ElseIf ilVGSet = 2 Then
                ilmnfVGCode = tgVehicleInfo(ilLoop).iMnfVehGp2              'sub-totals
            ElseIf ilVGSet = 3 Then
                ilmnfVGCode = tgVehicleInfo(ilLoop).iMnfVehGp3Mkt           'markets
            ElseIf ilVGSet = 4 Then
                ilmnfVGCode = tgVehicleInfo(ilLoop).iMnfVehGp4Fmt           'formats
            ElseIf ilVGSet = 5 Then
                ilmnfVGCode = tgVehicleInfo(ilLoop).iMnfVehGp5Rsch           'reserach
            ElseIf ilVGSet = 6 Then
                ilmnfVGCode = tgVehicleInfo(ilLoop).iMnfVehGp6Sub           'Sub-Company
            End If
        End If
    
    Exit Sub
End Sub
'
'           gCopySelectedVehicles - Create an array of selected vehicle codes
'           from a vehicle list box
'           <input> VehicleListBox
'           <output> array of integers indicating selected vehicle codes
'           5-30-18
Public Sub gCopySelectedVehicles(tlVehAff As control, ilSelectedVehicles() As Integer)
Dim iVef As Integer
Dim ilUpper As Integer

            ilUpper = 0
            ReDim ilSelectedVehicles(0 To 0) As Integer
            For iVef = 0 To tlVehAff.ListCount - 1 Step 1       'always loop on vehicle
                If tlVehAff.Selected(iVef) Then
                    ilSelectedVehicles(ilUpper) = tlVehAff.ItemData(iVef)
                    ilUpper = ilUpper + 1
                    ReDim Preserve ilSelectedVehicles(0 To ilUpper) As Integer
                End If
            Next iVef
            Exit Sub
End Sub
'
'       gSelectiveContract - if for selective contract, go to lst and find all matching lst records within date span
'       <input> vehiclelistbox
'       <output> aray of integers indicating selectegd vehicles with matching contract #
'          5-30-18
Public Sub gSelectiveContract(tlVehAff As control, ilSelectedVehicles() As Integer, slContract As String, dfStart As Date, dfEnd As Date)
Dim ilUpper As Integer
Dim ilLoop As Integer
Dim ContractRst As ADODB.Recordset
Dim dfTempStart As Date
Dim dfTempEnd As Date

            dfTempStart = DateAdd("D", -7, dfStart)         'backup earliest date for airing in previous week
            ReDim ilSelectedVehicles(0 To 0) As Integer
            ilUpper = 0
            SQLQuery = "Select distinct lstlogvefcode from lst where lstcntrno = " & Trim$(slContract) & " and lstlogdate >= '" & Format$(dfTempStart, sgSQLDateForm) & "' and lstlogdate <= '" & Format$(dfEnd, sgSQLDateForm) & "'"
            Set ContractRst = gSQLSelectCall(SQLQuery)
            While Not ContractRst.EOF
                For ilLoop = 0 To tlVehAff.ListCount - 1
                    If tlVehAff.ItemData(ilLoop) = ContractRst!lstLogVefCode Then
                        If tlVehAff.Selected(ilLoop) Then
                            ilSelectedVehicles(ilUpper) = ContractRst!lstLogVefCode
                            ilUpper = ilUpper + 1
                            ReDim Preserve ilSelectedVehicles(0 To ilUpper) As Integer
                            Exit For
                        End If
                    End If
                Next ilLoop
                ContractRst.MoveNext
            Wend
            Exit Sub
End Sub

'TTP 9943 - Add ability to import stations for report selectivity
'       gSelectiveStationsFromImport - Selects Stations from a Txt or CSV file
'       <input> Station Listbox
'       <input> SelectAll Checkbox
'       <input> Full filename of Txt or CSV file
Public Sub gSelectiveStationsFromImport(lbcStationList As control, CkcAll As control, slListSelectionFilename As String)
    Dim ilLoop As Integer
    Dim ilFn As Integer
    Dim slTemp As String
    Dim svValues As Variant
    Dim blFound As Boolean
    'Dim ilSelectedCount As Integer
    On Error GoTo ImportError
    
    'UnSelect ALL
    'ilSelectedCount = 0
    CkcAll.Value = vbUnchecked
    For ilLoop = 0 To lbcStationList.ListCount - 1
        If lbcStationList.Selected(ilLoop) = True Then
            lbcStationList.Selected(ilLoop) = False
        End If
    Next ilLoop
    ReDim tgStaNameAndCode(0)
    
    'Read File
    ilFn = FreeFile(0)
    Open slListSelectionFilename For Input As #ilFn
    Do
        Line Input #ilFn, slTemp
        slTemp = Trim(slTemp)
        If slTemp <> "" Then
            'Is this Comma Separated?
            If InStr(1, slTemp, ",") > 0 Then
                'Get 1st value from a Comma Separated File
                svValues = Split(slTemp, ",")
                slTemp = Trim(svValues(0))
            End If
            If slTemp <> "" Then
                slTemp = UCase(slTemp)
                'Select the item in the List
                blFound = False
                For ilLoop = 0 To lbcStationList.ListCount - 1
                    'check if Listbox of stations contains Call Letters Comma Description
                    If InStr(1, lbcStationList.List(ilLoop), ",") > 0 Then
                        'Match Call Letter Up to the Comma
                        If Left(lbcStationList.List(ilLoop), Len(slTemp) + 1) = slTemp & "," Then
                            lbcStationList.Selected(ilLoop) = True
                            blFound = True
                            Exit For
                        End If
                    Else
                        'Match entire Call Letter (No Comma)
                        If Trim(lbcStationList.List(ilLoop)) = slTemp Then
                            lbcStationList.Selected(ilLoop) = True
                            blFound = True
                            Exit For
                        End If
                    End If
                    
                Next ilLoop
                
                If blFound = False Then
                    'tgStaNameAndCode(UBound(tgStaNameAndCode)).iStationCode = slTemp
                    tgStaNameAndCode(UBound(tgStaNameAndCode)).sInfo = "Call Letters were not found - Station not selected."
                    tgStaNameAndCode(UBound(tgStaNameAndCode)).sStationName = slTemp
                    ReDim Preserve tgStaNameAndCode(UBound(tgStaNameAndCode) + 1)
                End If
            End If
        End If
    Loop While Not EOF(ilFn)
    Close #ilFn
    
    'Show Warning
    If UBound(tgStaNameAndCode) > 0 Then
        frmFastAddWarning.Caption = "Select Stations from import File"
        frmFastAddWarning.lbcFastAddWarning = ""
        frmFastAddWarning.cmdCancel.Enabled = False
        frmFastAddWarning.cmdCancel.Visible = False
        frmFastAddWarning.cmdContinue.Left = (frmFastAddWarning.Width - frmFastAddWarning.cmdContinue.Width) / 2
        frmFastAddWarning.lblFastAddWarning = "The following problems were found while selecting stations from File."
        frmFastAddWarning.lblAdvise = "These stations will not be selected."
        frmFastAddWarning.cmdContinue.Default = True
        frmFastAddWarning.cmdContinue.Cancel = True
        frmFastAddWarning.Show vbModal
    End If
    Exit Sub
    
ImportError:
    Close #ilFn
    MsgBox "Error Importing Station selection List" + vbCrLf & Err & " - " & Error(Err), vbCritical + vbOKOnly, "Select Stations from File"
    
End Sub

