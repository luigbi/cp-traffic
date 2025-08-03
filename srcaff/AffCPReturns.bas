Attribute VB_Name = "modCPReturns"
'******************************************************
'*  modmodCPReturn - various global declarations
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

'   D.S. 9/11/02
'
'   CPStatus - Was designed for CP Receipt and CP Spot only.  This is what is tested
'              against to populate the CP Posting grid
'
'   CPStatus - 0 = Outstanding
'              1 = Received - complete or maybe only partially posted  or   all aired
'                             ---------------------------------------       ---------
'              pertains to posting by:      spots by date,                  receipt only,
'                                           spots by advertiser             spots by count
'              2 = Not Aired
'
'   CPPostingStatus - Was designed for Spots by Date and Spots by Advertiser.  This
'                     indicates how complete the task of posting is or isn't.
'
'   CPPostingStatus 0=Not Posted; 1=Partially Posted; 2=Posting Completed
'
'   Legitimate Combinations - Any other combinations don't make sense
'
'   CPPostingStatus = 0 Not Posted    CPStatus = 0 Outstanding
'   CPPostingStatus = 0               CPStatus = 1 Receipt Only and Received
'   CPPostingStatus = 0               CPStatus = 2 NOT Aired
'   CPPostingStatus = 1 Partial       CPStatus = 1 Received
'   CPPostingStatus = 2 Complete      CPStatus = 1 Received
'   CPPostingStatus = 2 Complete      CPStatus = 2 Not Aired
'
'   CP Receipt Only  ---- Did we get the CP back or not and did all of the spots run or not
'
'   CP by Spot Count ---- How many out of how many spots ran.  We don't care what the times were.
'                         (currently not used by any of our clients)
'
'   Spots by Date ------- When posting from CPs for the week using the true times that the spots
'                         ran or not.
'
'   Spots by Advertiser - When posting from invoices rather than CPs and it's done for the
'                         entire month.

Public igTimes As Integer '0=Advertiser; 1=Dates
Public igCPStatus As Integer
Public igCPPostingStatus As Integer
Public igUpdateDTGrid As Integer
Public lgPreviousAttCode As Long
Public igChangedNewErased As Integer
Public lgMaxAstCode As Long
Public bgAdvtBlout As Boolean
Public sgAdvtBkoutKey As String

Public tgDelAst() As ASTINFO        'Used in the exports to Web for deleting old records

Type CPVEHICLE
    lRow As Long
    lCPDateIndex As Long
    sName As String * 42    'Vehicle or Call Letters
    sMarket As String * 20
    sZone As String * 3
    iCode As Integer
End Type

Type CPDATE
    iCol As Integer
    iStatus As Integer
    lCpttCode As Long
    sDate As String * 10
    iAttPostingType As Integer  '0=CP Only; 1=Spot Count; 2=by Time; 3=by Advertiser
    lAttCode As Long
    iAttTimeType As Integer
    iPostingStatus As Integer
    iNoSpotsGen As Integer
    iNoSpotsAired As Integer
    sAstStatus As String * 1
End Type

Type CPPOSTING
    lCpttCode As Long
    lAttCode As Long
    iVefCode As Integer
    iShttCode As Integer
    iAttTimeType As Integer
    sDate As String * 10
    sZone As String * 3
    iStatus As Integer
    iPostingStatus As Integer
    sAstStatus As String * 1
    iNumberDays As Integer
End Type

Public tgCPPosting() As CPPOSTING
'Sync copy between master and child
Private tmCPPosting As CPPOSTING
Private tmSvCPPosting As CPPOSTING
Private lmMATTCode As Long
Private smStatiomType As String  'N=Not master or child; M=Master; C=Child

Type ATTCrossDates
    lAttCode As Long
    lStartDate As Long
    lEndDate As Long
    iLoadFactor As Integer
    iDACode As Integer
    sForbidSplitLive As String * 1
    iComp As Integer
    sServiceAgreement As String * 1
    sExcludeFillSpot As String * 1
    sExcludeCntrTypeQ As String * 1 'Per Inquiry
    sExcludeCntrTypeR As String * 1 'Direct Response
    sExcludeCntrTypeT As String * 1 'Remnant
    sExcludeCntrTypeM As String * 1 'Promo
    sExcludeCntrTypeS As String * 1 'PSA
    sExcludeCntrTypeV As String * 1 'Reservation
End Type
Private tmATTCrossDates() As ATTCrossDates
Public bgAnyAttExclusions As Boolean

Type ASTINFO
    sKey As String * 20 'Date, Time, Break# and Position
    lCode As Long
    lAttCode As Long
    iShttCode As Integer
    iVefCode As Integer
    lSdfCode As Long
    lLstCode As Long
    lDatCode As Long
    iStatus As Integer
    sAirDate As String * 10
    sAirTime As String * 11
    sFeedDate As String * 10
    sFeedTime As String * 11
    sPledgeDate As String * 10
    sPledgeStartTime As String * 11
    sPledgeEndTime As String * 11
    iPledgeStatus As Integer
    iAdfCode As Integer
    iAnfCode As Integer
    sProd As String * 35
    sCart As String * 7
    sISCI As String * 20
    lCifCode As Long
    lCrfCsfCode As Long
    lCpfCode As Long
    sPdDays As String * 14
    iCPStatus As Integer
    sLstZone As String * 3
    lCntrNo As Long
    iLen As Integer
    lgsfCode As Long
    iRegionType As Integer  '0=None; 1=Split Copy without blackout; 2=Blackout split copy
    sRCart As String * 15    'McfName = 6, Number = 5, 2 for cut
    sRProduct As String * 35
    sRISCI As String * 20
    sRCreativeTitle As String * 30
    lRCrfCsfCode As Long
    lRCrfCode As Long
    lRCifCode As Long
    lRRsfCode As Long
    lIrtCode As Long
    sReplacementCue As String * 30
    'Dan 4/19/13
    lRCpfCode As Long
    sTruePledgeEndTime As String * 11   'Pledge end time. sPledgeEndTime is blank or equal to PledgeStartTime when Pledge status is live
    sTruePledgeDays As String * 7       'Pledge days Test for Y or N.
    sPdTimeExceedsFdTime As String * 1  'Pledge Time Length > Feed Time Length; i.e. Consider pledge information by daypart
    sPdDayFed As String * 1     'Pledge day: A=After feed day; B=Before feed day. Test for B
    '11/18/11: Used to retain blackout lst when replaced with original traffic lst
    lPrevBkoutLstCode As Long   'Previous LstCode assigned to ast, used in search of tmBkoutLst record
    '2/27/13: Added following fields to speed-up gGetLineParameters
    iLstLnVefCode As Integer
    lLstBkoutLstCode As Long
    '10/9/14: Retain the Generic copy within the blackout astInfo
    sGProd As String * 35
    sGCart As String * 7
    sGISCI As String * 20
    sLstStartDate As String * 10
    sLstEndDate As String * 10
    iLstSpotsWk As Integer
    iLstMon As Integer
    iLstTue As Integer
    iLstWed As Integer
    iLstThu As Integer
    iLstFri As Integer
    iLstSat As Integer
    iLstSun As Integer
    iLineNo As Integer
    iSpotType As Integer
    lLkAstCode As Long
    iMissedMnfCode As Integer
    sSplitNet As String * 1
    iAirPlay As Integer
    iAgfCode As Integer
    iComp As Integer
    sLstLnStartTime As String * 11   'Line start time.
    sLstLnEndTime As String * 11   'Line end time.
    sEmbeddedOrROS As String * 1
    sStationCompliant As String * 1
    sAgencyCompliant As String * 1
    sAffidavitSource As String * 2
    lEvtIDCefCode As Long
End Type

Type AETINFO
    lCode As Long
    lAtfCode As Long
    iShfCode As Integer
    iVefCode As Integer
    lSdfCode As Long
    sFeedDate As String * 10
    sFeedTime As String * 11
    sPledgeStartDate As String * 10
    sPledgeEndDate As String * 10
    sPledgeStartTime As String * 11
    sPledgeEndTime As String * 11
    sAdvt As String * 30
    sProd As String * 35
    sCart As String * 12
    sISCI As String * 20
    sCreative As String * 30
    lAstCode As Long
    iLen As Integer
    lCntrNo As Long
    sStatus As String * 1
    sPledgeDays As String * 7
    iProcessed As Integer
End Type

Type ASTTIMERANGE
    lDate As Long
    iGameNo As Integer
    lStartTime As Long
    lEndTime As Long
End Type

Type STARGUIDEAST
    sKey As String * 20
    tAstInfo As ASTINFO
    iBreakLen As Integer
    iVefCode As Integer
End Type

Type STARGUIDEXREF
    sKey As String * 150    'Short Title|Cart|ISCI|Creative Title|Vehicle
    iVefCode As Integer
End Type

Private tmLst As LST
Private tmSvLst As LST
Private tmLastLst As LST  'Used to restore last LST
Private tmModelLST As LST
Type BKOUTLST
    tLST As LST
    '10/12/14: set used flag
    bMatched As Boolean
    iDelete As Integer
End Type
Private tmBkoutLst() As BKOUTLST
Private lst_rst As ADODB.Recordset
Private sdf_rst As ADODB.Recordset

Private lmAdjSdfCode As Long
Private imAdjVefCode As Long
Private lmAdjLstLogTime As Long
Private imPriorAdjAdfCode As Integer
Private imNextAdjAdfCode As Integer


Private tmAstInfo() As ASTINFO
Private lmAttCode As Long
Private imVefCode As Integer
Private imShttCode As Integer
Private lmSdfCode As Long
Private lmAstCode As Long
Private lmLstCode As Long

Private smDefaultEmbeddedOrROS As String

'6/29/06: change gGetAstInfo to use API call
Private imAstRecLen As Integer
Private tmAst As AST
Private tmAstSrchKey As LONGKEY0
'6/29/06:  End of change

'8/31/06: Added CPTT Check Module
'CPTT Check
Type CPTTCHECK
    sKey As String * 70 'Vehicle Name | Station | Agreement Start Date
    iNoPrior As Integer
    lAgreementStart As Long
    lAgreementEnd As Long
    lAgreementMoStart As Long
    iNoAfter As Integer
    iNoAdd As Integer
    iNoDelete As Integer
    iVefCode As Integer
    iShfCode As Integer
    lAttCode As Long
    iAirWeekAdj As Integer  '0=this week, 7=next week
End Type

Type REGIONBREAKSPOTS
    sSource As String * 1   'C=Split Copy; N=Split Network
    lLstCode As Long
    lLogDate As Long
    lLogTime As Long
    lBreakNo As Long
    iPositionNo As Integer
    lSdfCode As Long
    sISCI As String * 20
    iFirstSplitNetRegion As Integer
End Type

'Each Unique Region is formed by ANDing one Include Format Category with one Include Other Category and with ALL Exclude Categories
'           Region examples Two Include Format Categories; Three Include Other Categories and two Exclude Categories
'           6 unique regions will be formed
'           Format1^Other1^Exclude1^Exclude2 or Format1^Other2^Exclude1^Exclude2 or Format1^Other3^Exclude1^Exclude2
'           Format2^Other1^Exclude1^Exclude2 or Format2^Other2^Exclude1^Exclude2 or Format2^Other3^Exclude1^Exclude2
Type REGIONDEFINITION
    lRotNo As Long
    lRafCode As Long
    sRegionName As String * 80  'Region Name
    sCategory As String * 1
    sInclExcl As String * 1
    lFormatFirst As Long   'Reference SplitInclude for Format or SplitInclude for All Categories except Format
    lOtherFirst As Long    'Reference Split Includes for all Categories except Format
                                'If the first category is not format, then no link is required with other INCLUDES
                                '   Region:  Each INCLUDE is AND with each EXCLUDE
                                '           Other1^Exclude1^Exclude2 or Other2^Exclude1^Exclude2 or Other3^Exclude1^Exclude2
    lExcludeFirst As Long    'References Excludes that are to be AND with INCLUDES
    sPtType As String * 1
    lCopyCode As Long
    lCrfCode As Long
    lRsfCode As Long
    '3/3/18: Wegener Export only
    iStationCount As Integer    '<> 0 means split by station only: lOtherFirst references original split category. Export Wegener only
    lStationOtherFirst As Long
    iPoolNextFinal As Integer
    iPoolAdfCode As Integer
    lPoolCrfCode As Long
    bPoolUpdated As Boolean
End Type

Type SPLITCATEGORYINFO
    sCategory As String * 1
    sName As String * 40
    iIntCode As Integer
    lLongCode As Long
    lNext As Long
End Type

Type REGIONBREAKSPOTINFO
    lStartIndex As Long
    lEndIndex As Long
    lSdfCode As Long
    sISCI As String * 20
    iNoRegions As Integer
    iPositionNo As Integer
End Type

Type SPLITNETREGION
    lLstCode As Long
    iNext As Integer
End Type

Type CUSTOMGROUPNAMES
    sName As String * 10
    sCategoryGroup As String * 58
End Type

Type OLACUSTOMGROUPNAMES
    sName As String * 14
    iFirst As Integer
End Type

Type OLAUNIQUEGROUPNAMES
    sName As String * 60
    sGroupName As String * 14
    sGroupType As String * 10
End Type

Type OLACATEGORYNAME
    sCategory As String * 1
    sName As String * 10
    iNext As Integer
End Type

Type WEGENERIMPORT
    sCallLetters As String * 20 'KXXX-AM/KBBB-FM plus extra
    iShttCode As Integer
    iMktCode As Integer
    iMSAMktCode As Integer
    iFormatCode As Integer
    iTztCode As Integer
    sPostalName As String * 2
    sSerialNo1 As String * 10
    sPort As String * 1
    lVefCodeFirst As Long
    iRecGroupFd As Integer  'True=Call Letters found in RecGroup Import file
End Type

Type WEGENERVEHINFO
    sGroup As String * 30
    sPort As String * 1
    iVefCode As Integer
    lVefCodeNext As Long
End Type

Type WEGENERFORMATSORT
    iFmtCode As Integer
    iFirst As Integer
End Type

Type WEGENERTIMEZONESORT
    iTztCode As Integer
    iFirst As Integer
End Type

Type WEGENERMARKETSORT
    iMktCode As Integer
    iFirst As Integer
End Type

Type WEGENERPOSTALSORT
    sPostalName As String * 2
    iFirst As Integer
End Type

Type WEGENERINDEX
    iIndex As Integer
    iNext As Integer
End Type


Dim tmRegionDefinition() As REGIONDEFINITION
Dim tmSplitCategoryInfo() As SPLITCATEGORYINFO

Type XDFDINFO
    sKey As String * 130 'Vehicle Code (5) + ProgCodeID + ISCI + StationID
    iVefCode As Integer
    lStationID As Long
    sISCI As String * 80            'ISCI or Advt Abbr,Product(ISCI)
    sCreativeTitle As String * 30
    lRotStartDate As Long
    lRotEndDate As Long
    sShortTitle As String * 15
    sProgCodeID As String * 8
End Type

Type IDCGENERIC
    lCifCode As Long
    lCrfCsfCode As Long
    sISCI As String * 15
    'lFeedDate As Long
    lFirstSplit As Long
    lTriggerId As Long
End Type

Type IDCSPLIT
    lRCrfCode As Long
    'lRCifCode As Long
    'lRRsfCode As Long
    lRafCode As Long
    'sRISCI As String * 15
    lNextSplit As Long
    lFirstReceiver As Long
    sAdvName As String * 30
    'sUnfilteredISCI As String * 20
    lFirstRISCI As Long
    sCrfStartDate As String * 10
    sCrfEndDate As String * 10
    '5882
    sStartTime As String * 10
    sEndTime As String * 10
    iRotation As Integer
End Type

Type IDCRECEIVER
    sReceiverID As String * 5
    sCallLetters As String * 10
    '6419
    lAttCode As Long
    lNextReceiver As Long
End Type

Dim smGetAstKey As String
Type LSTPLUSDT
    tLST As LST
    lDate As Long
    lTime As Long
End Type
Dim tmEZLSTAst() As LSTPLUSDT
Dim tmCZLSTAst() As LSTPLUSDT
Dim tmMZLSTAst() As LSTPLUSDT
Dim tmPZLSTAst() As LSTPLUSDT
Dim tmNZLSTAst() As LSTPLUSDT 'No zone
Dim tmLSTAst() As LSTPLUSDT
Dim lmRCSdfCode() As Long
Dim bmImportSpot As Boolean
Private smGetDlfKey As String

Type BUILDLST
    tLST As LST
    iBreakLength As Integer
    bIgnore As Boolean
End Type

Private smBuildLstKey As String
Private tmBuildLst() As BUILDLST

Private lmAttForDat As Long
Private tmDatRst() As DATRST


Type REGIONASSIGNMENTINFO
    sKey As String * 20    'sdfCode and Seq number
    lSdfCode As Long
    lRDIndex As Long
    lSCIStartIndex As Long
    lSCIEndIndex As Long
End Type
Dim tmRegionAssignmentInfo() As REGIONASSIGNMENTINFO

Dim tmRegionDefinitionForSpots() As REGIONDEFINITION
Dim tmSplitCategoryInfoForSpots() As SPLITCATEGORYINFO

Type DATPLEDGEINFO
    lDatCode As Long
    lAttCode As Long
    iVefCode As Integer
    sFeedDate As String * 10
    sFeedTime As String * 11
    sPledgeDate As String * 10
    sPledgeStartTime As String * 11
    sPledgeEndTime As String * 11
    iPledgeStatus As Integer
End Type

Private lmCopyDate As Long

Dim lmMktronAttCode As Long
Dim smAttExportToMarketron As String
Dim att_rst As ADODB.Recordset

Dim rsf_rst As ADODB.Recordset
Dim sef_rst As ADODB.Recordset
Dim cif_rst As ADODB.Recordset
Dim crf_rst As ADODB.Recordset
Dim cvf_rst As ADODB.Recordset
Dim pvf_rst As ADODB.Recordset
Dim vef_rst As ADODB.Recordset
Dim cnf_rst As ADODB.Recordset
Dim cpf_rst As ADODB.Recordset
Private rst_lcf As ADODB.Recordset
Private rst_lvf As ADODB.Recordset
Private rst_Pet As ADODB.Recordset
Private dat_rst As ADODB.Recordset
Private DatPledge_rst As ADODB.Recordset
Private cptt_rst As ADODB.Recordset
Private rst_Rlf As ADODB.Recordset
Private rst_Ust As ADODB.Recordset
Private rst_Genl As ADODB.Recordset
Private rst_irt As ADODB.Recordset
Private ast_rst As ADODB.Recordset
Private rst_abf As ADODB.Recordset
Private chf_rst As ADODB.Recordset
Private rst_adf As ADODB.Recordset
Private rst_Shtt As ADODB.Recordset

Dim sgLastProgramDefinition As String
Dim tgProgramDefinition() As VEF



Public Function gDayMap(sInDays As String) As String
    Dim sDays As String
    
    sDays = Trim$(sInDays)
    If sDays = "MoTuWeThFrSaSu" Then
        sDays = "M-Su"
    ElseIf sDays = "MoTuWeThFrSa" Then
        sDays = "M-Sa"
    ElseIf sDays = "MoTuWeThFr" Then
        sDays = "M-F"
    ElseIf sDays = "MoTuWeTh" Then
        sDays = "M-Th"
    ElseIf sDays = "MoTuWe" Then
        sDays = "M-W"
    ElseIf sDays = "MoTu" Then
        sDays = "M-Tu"
    ElseIf sDays = "TuWeThFrSaSu" Then
        sDays = "Tu-Su"
    ElseIf sDays = "TuWeThFrSa" Then
        sDays = "Tu-Sa"
    ElseIf sDays = "TuWeThFr" Then
        sDays = "Tu-F"
    ElseIf sDays = "TuWeTh" Then
        sDays = "Tu-Th"
    ElseIf sDays = "TuWe" Then
        sDays = "Tu-W"
    ElseIf sDays = "WeThFrSaSu" Then
        sDays = "W-Su"
    ElseIf sDays = "WeThFrSa" Then
        sDays = "W-Sa"
    ElseIf sDays = "WeThFr" Then
        sDays = "W-F"
    ElseIf sDays = "WeTh" Then
        sDays = "W-Th"
    ElseIf sDays = "ThFrSaSu" Then
        sDays = "Th-Su"
    ElseIf sDays = "ThFrSa" Then
        sDays = "Th-Sa"
    ElseIf sDays = "ThFr" Then
        sDays = "Th-F"
    ElseIf sDays = "FrSaSu" Then
        sDays = "F-Su"
    ElseIf sDays = "FrSa" Then
        sDays = "F-Sa"
    ElseIf sDays = "SaSu" Then
        sDays = "S-S"
    End If
    gDayMap = sDays
End Function

Public Sub gUnMapDays(sInDays As String, iDays() As Integer)
    Dim sDays As String
    Dim iLoop As Integer
    Dim iStart As Integer
    Dim iEnd As Integer
    
    sDays = UCase$(Trim$(sInDays))
    For iLoop = 0 To 6 Step 1
        iDays(iLoop) = False
    Next iLoop
    If (sDays = "MOTUWETHFRSASU") Or (sDays = "M-SU") Then
        iStart = 0
        iEnd = 6
    ElseIf (sDays = "MOTUWETHFRSA") Or (sDays = "M-SA") Then
        iStart = 0
        iEnd = 5
    ElseIf (sDays = "MOTUWETHFR") Or (sDays = "M-F") Then
        iStart = 0
        iEnd = 4
    ElseIf (sDays = "MOTUWETH") Or (sDays = "M-TH") Then
        iStart = 0
        iEnd = 3
    ElseIf (sDays = "MOTUWE") Or (sDays = "M-W") Then
        iStart = 0
        iEnd = 2
    ElseIf (sDays = "MOTU") Or (sDays = "M-TU") Then
        iStart = 0
        iEnd = 1
    ElseIf (sDays = "TUWETHFRSASU") Or (sDays = "TU-SU") Then
        iStart = 1
        iEnd = 6
    ElseIf (sDays = "TUWETHFRSA") Or (sDays = "TU-SA") Then
        iStart = 1
        iEnd = 5
    ElseIf (sDays = "TUWETHFR") Or (sDays = "TU-F") Then
        iStart = 1
        iEnd = 4
    ElseIf (sDays = "TUWETH") Or (sDays = "TU-TH") Then
        iStart = 1
        iEnd = 3
    ElseIf (sDays = "TUWE") Or (sDays = "TU-W") Then
        iStart = 1
        iEnd = 2
    ElseIf (sDays = "WETHFRSASU") Or (sDays = "W-SU") Then
        iStart = 2
        iEnd = 6
    ElseIf (sDays = "WETHFRSA") Or (sDays = "W-SA") Then
        iStart = 2
        iEnd = 5
    ElseIf (sDays = "WETHFR") Or (sDays = "W-F") Then
        iStart = 2
        iEnd = 4
    ElseIf (sDays = "WETH") Or (sDays = "W-TH") Then
        iStart = 2
        iEnd = 3
    ElseIf (sDays = "THFRSASU") Or (sDays = "TH-SU") Then
        iStart = 3
        iEnd = 6
    ElseIf (sDays = "THFRSA") Or (sDays = "TH-SA") Then
        iStart = 3
        iEnd = 5
    ElseIf (sDays = "THFR") Or (sDays = "TH-F") Then
        iStart = 3
        iEnd = 4
    ElseIf (sDays = "FRSASU") Or (sDays = "F-SU") Then
        iStart = 4
        iEnd = 6
    ElseIf (sDays = "FRSA") Or (sDays = "F-SA") Then
        iStart = 4
        iEnd = 5
    ElseIf (sDays = "SASU") Or (sDays = "S-S") Then
        iStart = 4
        iEnd = 6
    ElseIf (sDays = "MO") Then
        iStart = 0
        iEnd = 0
    ElseIf (sDays = "TU") Then
        iStart = 1
        iEnd = 1
    ElseIf (sDays = "WE") Then
        iStart = 2
        iEnd = 2
    ElseIf (sDays = "TH") Then
        iStart = 3
        iEnd = 3
    ElseIf (sDays = "FR") Then
        iStart = 4
        iEnd = 4
    ElseIf (sDays = "SA") Then
        iStart = 5
        iEnd = 5
    ElseIf (sDays = "SU") Then
        iStart = 6
        iEnd = 6
    End If
    For iLoop = iStart To iEnd Step 1
        iDays(iLoop) = True
    Next iLoop
End Sub



Private Function mGetAstInfo(hlAst As Integer, tlCPDat() As DAT, tlAstInfo() As ASTINFO, ilInAdfCode As Integer, iAddAst As Integer, iUpdateCpttStatus As Integer, ilBuildAstInfo As Integer, Optional blInGetRegionCopy As Boolean = True, Optional llSelGsfCode As Long = -1, Optional blFeedAdjOnReturn As Boolean = False, Optional blFilterByAirDates As Boolean = False, Optional blIncludePledgeInfo As Boolean = True, Optional blCreateServiceATTSpots As Boolean = False) As Boolean
'   igTime: 0=By Month; 1=By Week sort by air date and air time; 2=By Week sort by feed date and feed time;
'           3=By Date Range sort by air date and air time; 4=By Date Range sort by feed date and feed time
    Dim iAdfCode As Integer
    Dim iLoop As Integer
    Dim iVef As Integer
    Dim iUpper As Integer
    Dim sFWkDate As String
    Dim sLWkDate As String
    Dim lFWkDate As Long
    Dim lLWkDate As Long
    Dim sAdvertiser As String
    Dim iDat As Integer
    Dim iFound As Integer
    Dim lSTime As Long
    Dim lETime As Long
    Dim lTime As Long
    Dim iIndex As Integer
    Dim sFdDate As String
    Dim sFdTime As String
    Dim sPdDate As String
    Dim sPdDays As String
    Dim sPdSTime As String
    Dim sPdETime As String
    Dim sTPdETime As String    'True pledge end time.  This is required because if live sPdETime is set to blank
    Dim iFdDay As Integer
    Dim iPdDay As Integer
    Dim sTPdDays As String
    Dim sStr As String
    Dim iDay As Integer
    Dim sAirDate As String
    Dim sAirTime As String
    Dim iStatus As Integer
    Dim sZone As String
    Dim iZone As Integer
    Dim iLocalAdj As Integer
    Dim lLogTime As Long
    Dim sLogTime As String
    Dim lLogDate As Long
    Dim sLogDate As String
    Dim iAst As Integer
    Dim iCPStatus As Integer
    Dim lAstCode As Long
    Dim l2LogDate As Long
    Dim l2LogTime As Long
    Dim iZoneFound As Integer
    Dim lTmpAstCode As Long
    Dim ilLoadFactor As Integer
    Dim ilLoadIdx As Integer
    Dim iPledged As Integer
    Dim llCount As Long
    Dim att_rst As ADODB.Recordset
    Dim llTemp As Long
    Dim ilPass As Integer
    Dim llAtt As Long
    Dim slTemp As String
    Dim slTemp2 As String
    Dim llAttValid As Long
    'Dim llAttOnAir As Long
    'Dim llAttOffAir As Long
    Dim ilAtt As Integer
    Dim ilShttCode As Integer
    Dim ilVefCode As Integer
    Dim ilSDFMatchOk As Integer
    Dim iAst1 As Integer
    Dim ilDatStart As Integer
    Dim ilPledgeStatus As Integer
    Dim llPdSTime As Long
    Dim llPdETime As Long
    Dim llMaxCode As Long
    Dim ilFillLen As Integer
    Dim blResetFillLen As Boolean
    Dim ilInsertError As Integer
    '6/29/06: change mGetAstInfo to use API call
    Dim tlAstSrchKey As LONGKEY0
    Dim tlAst As AST
    Dim ilRet As Integer
    Dim ilIncludeLST As Integer
    Dim ilIssueMoveNext As Integer
    Dim slSplitNetwork As String
    Dim ilChkLen As Integer
    Dim ilLen As Integer
    Dim llChkLenRafCode As Long
    Dim ilChkLenPosition As Integer
    Dim ilFindFill As Integer
    Dim ilEof As Integer
    Dim ilNumberAsterisk As Integer
    Dim ilPledgeMatch As Integer
    Dim ilPos As Integer
    Dim llCompareLST1 As Long
    Dim llCompareLST2 As Long
    Dim llEZUpperLst As Long
    Dim llCZUpperLst As Long
    Dim llMZUpperLst As Long
    Dim llPZUpperLst As Long
    Dim llNZUpperLst As Long
    Dim llLstLoop As Long
    Dim ilFound As Integer
    Dim llSdfCode As Long
    Dim slForbidSplitLive As String
    Dim ilDACode As Integer
    Dim llFdTimeLen As Long
    Dim llPdTimeLen As Long
    Dim slDayPart As String
    Dim llVpf As Long
    'Dim slPostedAirTime As String
    Dim llPostedAirDate As Long
    Dim llPostedAirTime As Long
    Dim ilAdjDay As Integer
    Dim blErrorMsgLogged As Boolean
    Dim ilLimit As Integer
    Dim blLstOk As Boolean
    Dim ilGsfLoop As Integer
    Dim llDATCode As Long
    Dim ilType2Count As Integer
    Dim llLogTestDate As Long
    Dim ilAvailLength As Integer
    Dim ilAvail As Integer
    Dim llDateTest As Long
    Dim llTimeTest As Long
    Dim tlDatPledgeInfo As DATPLEDGEINFO
    Dim blGetCopy As Boolean
    Dim ilAirPlay As Integer
    Dim ilAttComp As Integer
    Dim llLockRec As Long
    Dim slUser As String
    Dim llSTime As Long
    Dim llETime As Long
    Dim blProgressShown As Boolean
    Dim ilShtt As Integer
    Dim llVef As Long
    Dim blBlockWarningShown As Boolean
    '10/9/15: Retain Generic copy from LST and after black handled update astInfo
    Dim slGISCI As String
    Dim slGCart As String
    Dim slGProd As String
    Dim slServiceAgreement As String
    '11/20/14: Add Delivery links
    Dim llDlfTime As Long
    Dim blExtendExist As Boolean
    Dim slSortDate As String
    Dim slSortTime As String
    Dim slSortPosition As String
    '3/5/15: Rebuild spots if gBuildAstInfoFromAst failed
    Dim blRebuildAst As Boolean
    '3/8/16
    Dim blMGSpot As Boolean
    Dim slMGMissedFeedDate As String
    Dim slMGMissedFeedTime As String
    '4/7/16: Handle case where ast re-created and spot has been posted
    Dim slPostedAirDate As String
    Dim slPostedAirTime As String
    '8/31/16: Check if comment suppressed
    Dim ilVff As Integer
    Dim slHideCommOnWeb As String
    Dim slSQLQuery As String
    Dim blGetRegionCopy As Boolean
    Dim blAdvtBkout As Boolean
    Dim ilAstCPStatus As Integer
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilPledgeCompliantCount As Integer
    Dim ilAgyCompliantCount As Integer
    
    blErrorMsgLogged = False
    
    iAdfCode = ilInAdfCode
    imAstRecLen = Len(tmAst)
    gGetPoolAdf
    iLoop = 0
    '2/6/19: if readding ast, must reassign regional copy as it would not to assigned to the ast record
    blRebuildAst = False
    blGetRegionCopy = blInGetRegionCopy
    If (iAddAst = True) And (blInGetRegionCopy = False) Then
        blGetRegionCopy = True
    End If
    '2/24/17: Handle case where Split Fill changed
    If ((Asc(sgSpfUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS) Then
        If gObtainReplacments() = 2 Then
            tgCPPosting(iLoop).sAstStatus = "N"
        End If
    End If
    If (tgCPPosting(iLoop).sAstStatus = "C") Then
        If gBuildAstInfoFromAst(hlAst, tlCPDat(), tlAstInfo(), iAdfCode, blGetRegionCopy, llSelGsfCode, blFeedAdjOnReturn, blFilterByAirDates, blIncludePledgeInfo, blCreateServiceATTSpots) Then
            mGetAstInfo = True
            Exit Function
        End If
    End If
    iAdfCode = 0
    '6/29/06: end of change
    lgSTime8 = timeGetTime
        
    '11/15/11: Build copy if not alread built
    If blGetRegionCopy Then
        blGetCopy = False
        If lmCopyDate <= 0 Then
            blGetCopy = True
        Else
            If ((igTimes = 3) Or (igTimes = 4)) And (blFeedAdjOnReturn) Then
                If gDateValue(DateAdd("d", -1, gObtainPrevMonday(tgCPPosting(0).sDate))) < lmCopyDate Then
                    blGetCopy = True
                End If
            Else
                If gDateValue(gObtainPrevMonday(tgCPPosting(0).sDate)) < lmCopyDate Then
                    blGetCopy = True
                End If
            End If
        End If
        ilRet = 0
        'On Error GoTo CheckForCopyErr:
        'ilLimit = LBound(tgCifCpfInfo1)
        If PeekArray(tgCifCpfInfo1).Ptr <> 0 Then
            ilLimit = LBound(tgCifCpfInfo1)
            If (UBound(tgCifCpfInfo1) <= LBound(tgCifCpfInfo1)) Then
                ilRet = 1
            End If
        Else
            ilRet = 1
            ilLimit = 0
        End If
        
        'If (ilRet = 1) Or (UBound(tgCifCpfInfo1) <= LBound(tgCifCpfInfo1)) Then
        If (ilRet = 1) Then
            If ((igTimes = 3) Or (igTimes = 4)) And (blFeedAdjOnReturn) Then
                lmCopyDate = gDateValue(DateAdd("d", -1, gObtainPrevMonday(tgCPPosting(0).sDate)))
                ilRet = gPopCopy(Format$(DateAdd("d", -1, gObtainPrevMonday(tgCPPosting(0).sDate)), sgShowDateForm), "mGetAstInfo")
            Else
                lmCopyDate = gDateValue(gObtainPrevMonday(tgCPPosting(0).sDate))
                ilRet = gPopCopy(Format$(gObtainPrevMonday(tgCPPosting(0).sDate), sgShowDateForm), "mGetAstInfo")
            End If
        End If
    End If
    On Error GoTo ErrHand
    llLockRec = -1
    For iLoop = 0 To UBound(tgCPPosting) - 1 Step 1 ' This loop is always 0 to 0.
        'If igExportSource = 2 Then
            DoEvents
        'End If
        
        If llLockRec > 0 Then
            ilRet = gDeleteLockRec_ByRlfCode(llLockRec)
        End If
        llSTime = timeGetTime
        blBlockWarningShown = False
        
        Do
            llLockRec = gCreateLockRec("A", "G", tgCPPosting(iLoop).lCpttCode, True, slUser)
            If llLockRec > 0 Then
                Exit Do
            End If
            Sleep 200
            llETime = timeGetTime
            ''Retry upto 5 minutes twice for a total of 10 minutes (timeGetTime is in milliseconds)
            '2/8/18: Retry upto 2 1/2 minutes (150 seconds) twice for a total of 5 minutes (timeGetTime is in milliseconds)
            'If ((llETime - llSTime) > CLng(300000)) Then
            If ((llETime - llSTime) > CLng(150000)) Then
                ilShtt = gBinarySearchStationInfoByCode(tgCPPosting(iLoop).iShttCode)
                If ilShtt <> -1 Then
                    sStr = Trim$(tgStationInfoByCode(ilShtt).sCallLetters)
                Else
                    sStr = ""
                End If
                llVef = gBinarySearchVef(CLng(tgCPPosting(iLoop).iVefCode))
                If llVef <> -1 Then
                    sStr = Trim$(sStr & " " & tgVehicleInfo(llVef).sVehicle)
                End If
                If igExportSource = 2 Then
                    If Not blBlockWarningShown Then
                        'gLogMsg "User " & slUser & " is Blocking the Gathering of Station Spots for " & sStr, "AffErrorLog.txt", False
                        blBlockWarningShown = True
                        llSTime = timeGetTime
                    Else
                        bgTaskBlocked = True
                        sgTaskBlockedDate = Format(Now, "mmddyyyy")
                        If sgTaskBlockedName <> "" Then
                            gLogMsg "Redo Task for " & sStr & ", " & slUser & " Blocked Task, Running " & sgTaskBlockedName, "TaskBlocked_" & sgTaskBlockedDate & ".txt", False
                        Else
                            gLogMsg "Redo Task for " & sStr & ", " & slUser & " Blocked Task", "TaskBlocked_" & sgTaskBlockedDate & ".txt", False
                        End If
                        ReDim tlAstInfo(0 To 0) As ASTINFO
                        ReDim tgDelAst(0 To 0) As ASTINFO
                        mGetAstInfo = False
                        Exit Function
                    End If
                Else
                    '3/5/15: exit if failed twice
                    'MsgBox "User " & slUser & " is Blocking the Gathering of Station Spots for " & sStr & " Use View Blocks to examine the block", vbCritical + vbOKOnly, "Gather Spots Blocked"
                    'llSTime = timeGetTime
                    If Not blBlockWarningShown Then
                        'MsgBox "User " & slUser & " is Blocking the Gathering of Station Spots for " & sStr & " will continue trying for another 5 minutes", vbCritical + vbOKOnly, "Gather Spots Blocked"
                        blBlockWarningShown = True
                        llSTime = timeGetTime
                    Else
                        bgTaskBlocked = True
                        sgTaskBlockedDate = Format(Now, "mmddyyyy")
                        'MsgBox "User " & slUser & " is Blocking the Gathering of Station Spots for " & sStr & " Use View Blocks to examine the block. Try again later", vbCritical + vbOKOnly, "Gather Spots Blocked"
                        If sgTaskBlockedName <> "" Then
                            gLogMsg "Redo Task for " & sStr & ", " & slUser & " Blocked Task, Running " & sgTaskBlockedName, "TaskBlocked_" & sgTaskBlockedDate & ".txt", False
                        Else
                            gLogMsg "Redo Task for " & sStr & ", " & slUser & " Blocked Task", "TaskBlocked_" & sgTaskBlockedDate & ".txt", False
                        End If
                        ReDim tlAstInfo(0 To 0) As ASTINFO
                        ReDim tgDelAst(0 To 0) As ASTINFO
                        mGetAstInfo = False
                        Exit Function
                    End If
                End If
                DoEvents
            End If
        Loop While llLockRec <= 0
        
        ilShttCode = tgCPPosting(iLoop).iShttCode
        ilVefCode = tgCPPosting(iLoop).iVefCode
        'If (igTimes = 3) Or (igTimes = 4) Then
        '    If blFeedAdjOnReturn Then
        '        sFWkDate = Format$(DateAdd("d", -1, tgCPPosting(iLoop).sDate), sgShowDateForm)
        '    Else
        '        sFWkDate = Format$(tgCPPosting(iLoop).sDate, sgShowDateForm)
        '    End If
        'Else
        '    sFWkDate = Format$(gObtainPrevMonday(tgCPPosting(iLoop).sDate), sgShowDateForm)
        'End If
        'If igTimes = 0 Then
        '    sLWkDate = Format$(gObtainEndStd(tgCPPosting(iLoop).sDate), sgShowDateForm)
        'ElseIf (igTimes = 3) Or (igTimes = 4) Then
        '    If blFeedAdjOnReturn Then
        '        sLWkDate = Format$(DateAdd("d", tgCPPosting(iLoop).iNumberDays, tgCPPosting(iLoop).sDate), sgShowDateForm)
        '    Else
        '        sLWkDate = Format$(DateAdd("d", tgCPPosting(iLoop).iNumberDays - 1, tgCPPosting(iLoop).sDate), sgShowDateForm)
        '    End If
        'Else
        '    sLWkDate = Format$(gObtainNextSunday(tgCPPosting(iLoop).sDate), sgShowDateForm)
        'End If
        mGetAstDateRange tgCPPosting(iLoop), blFeedAdjOnReturn, sFWkDate, sLWkDate
        lFWkDate = DateValue(gAdjYear(sFWkDate))
        lLWkDate = DateValue(gAdjYear(sLWkDate))
        blAdvtBkout = mAdvtBkout(ilInAdfCode, ilVefCode, lFWkDate, lLWkDate, blFeedAdjOnReturn)
        If Not blAdvtBkout Then
            iAdfCode = ilInAdfCode
        End If
        
        llVpf = gBinarySearchVpf(CLng(tgCPPosting(iLoop).iVefCode))
        If llVpf <> -1 Then
            smDefaultEmbeddedOrROS = tgVpfOptions(llVpf).sEmbeddedOrROS
        End If
        If Trim$(smDefaultEmbeddedOrROS) = "" Then
            smDefaultEmbeddedOrROS = "R"
        End If
        '8/31/16: Check if comment suppressed
        slHideCommOnWeb = "N"
        ilVff = gBinarySearchVff(CLng(tgCPPosting(iLoop).iVefCode))
        If ilVff <> -1 Then
            slHideCommOnWeb = tgVffInfo(ilVff).sHideCommOnWeb
        End If
        ' C = created and valid. Non C(N, ' ', R ) = Currupt.
        If (tgCPPosting(iLoop).sAstStatus <> "C") Or ((tgCPPosting(iLoop).sAstStatus = "C") And (ilBuildAstInfo)) Or (blRebuildAst) Then
            sZone = tgCPPosting(iLoop).sZone
            iLocalAdj = 0
            iZoneFound = False
            ilNumberAsterisk = 0
            iVef = gBinarySearchVef(CLng(ilVefCode))
            
            '12/20/14
            ''11/20/14: Add Delivery links
            'If sgDelNet = "Y" Then
            '    sStr = ilVefCode & sFWkDate & sLWkDate & sZone
            '    If StrComp(smGetDlfKey, sStr, vbTextCompare) <> 0 Then
            '        smGetDlfKey = sStr
            '        'Set bgDlfExist
            '        gDlfExist ilVefCode, sFWkDate, sLWkDate, sZone
            '    End If
            'Else
            '    bgDlfExist = False
            'End If
            
            '12/12/14:
            '' Adjust time zone properly.
            'If (Len(Trim$(tgCPPosting(iLoop).sZone)) <> 0) And (Not bgDlfExist) Then
            If (Len(Trim$(tgCPPosting(iLoop).sZone)) <> 0) Then
                'Get zone
                If iVef <> -1 Then
                    For iZone = LBound(tgVehicleInfo(iVef).sZone) To UBound(tgVehicleInfo(iVef).sZone) Step 1
                        'If igExportSource = 2 Then
                            DoEvents
                        'End If
                        If Trim$(tgVehicleInfo(iVef).sZone(iZone)) = Trim$(tgCPPosting(iLoop).sZone) Then
                            If (tgVehicleInfo(iVef).sFed(iZone) <> "*") And (Trim$(tgVehicleInfo(iVef).sFed(iZone)) <> "") And (tgVehicleInfo(iVef).iBaseZone(iZone) <> -1) Then
                                sZone = tgVehicleInfo(iVef).sZone(tgVehicleInfo(iVef).iBaseZone(iZone))
                                iLocalAdj = tgVehicleInfo(iVef).iLocalAdj(iZone)
                                iZoneFound = True
                            End If
                            Exit For
                        End If
                    Next iZone
                    For iZone = LBound(tgVehicleInfo(iVef).sZone) To UBound(tgVehicleInfo(iVef).sZone) Step 1
                        If igExportSource = 2 Then
                            DoEvents
                        End If
                        If tgVehicleInfo(iVef).sFed(iZone) = "*" Then
                            ilNumberAsterisk = ilNumberAsterisk + 1
                        End If
                    Next iZone
                End If
            End If
            If (Not iZoneFound) And (ilNumberAsterisk <= 1) Then
                sZone = ""
            End If
            ' so we can set the AST posting status appropriate
            If tgCPPosting(iLoop).iPostingStatus = 0 Then
                iCPStatus = 0 'Not Posted
            Else
                If tgCPPosting(iLoop).iStatus = 2 Then
                    iCPStatus = 2 'Cptt posted as Not Aired
                Else
                    'iCPStatus = 1 'Posted as Complete or Partially Complete
                    If tgCPPosting(iLoop).iPostingStatus = 1 Then
                        iCPStatus = -1   'Partially Posted
                    Else
                        iCPStatus = 1 'Posted as Complete
                    End If
                End If
            End If
            
            '11/24/11: Build array of games allowed so that Sunday spots airing on Monday can be included
            ReDim llGsfCode(0 To 0) As Long
            If iVef <> -1 Then
                If tgVehicleInfo(iVef).sVehType = "G" Then
                    slSQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfVefCode = " & ilVefCode & " AND gsfAirDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "'" & " AND gsfAirDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "'" & ")"
                    Set rst_Genl = gSQLSelectCall(slSQLQuery)
                    Do While Not rst_Genl.EOF
                        llGsfCode(UBound(llGsfCode)) = rst_Genl!gsfCode
                        ReDim Preserve llGsfCode(0 To UBound(llGsfCode) + 1) As Long
                        rst_Genl.MoveNext
                    Loop
                
                End If
            End If
            If igExportSource = 2 Then
                DoEvents
            End If
'1/9/13:  Obtain all Agreement that are active for specified dates
'            'D.S. 10/25/04 Get the load factor from the agreement
'            slSQLQuery = "SELECT attLoad, attShfCode, attVefCode, attOnAir, attOffAir, attDropDate, attForbidSplitLive, attPledgeType"
'            slSQLQuery = slSQLQuery + " FROM att"
'            slSQLQuery = slSQLQuery + " WHERE (attCode = " & tgCPPosting(iLoop).lAttCode & ")"
'            Set rst = gSQLSelectCall(slSQLQuery)
'            ilLoadFactor = rst!attLoad
'            If ilLoadFactor < 1 Then
'                ilLoadFactor = 1
'            End If
'            slForbidSplitLive = rst!attForbidSplitLive
'            ilDACode = 1
'            If rst!attPledgeType = "D" Then
'                ilDACode = 0
'            ElseIf rst!attPledgeType = "A" Then
'                ilDACode = 1
'            ElseIf rst!attPledgeType = "C" Then
'                ilDACode = 2
'            End If
'            llAttValid = True
'            If DateValue(gAdjYear(Trim$(rst!attOffAir))) <= DateValue(gAdjYear(Trim$(rst!attDropDate))) Then
'                slTemp = Trim$(rst!attOffAir)
'            Else
'                slTemp = Trim$(rst!attDropDate)
'            End If
'
'            llAttOnAir = DateValue(gAdjYear(rst!attOnAir))
'            llAttOffAir = DateValue(gAdjYear(slTemp))
'
'            If DateValue(gAdjYear(slTemp)) < DateValue(gAdjYear(sFWkDate)) Then
'                llAttValid = False
'            End If
'
'            'now get adjacent attcodes from the adjacent agreement.  save the one that the latest
'            'off air or drop date that's prior to the current agreement.  Save the agreement number
'            'in a global var. lgPreviousAttCode.  If not found set it to -1
'            'D.S. 10/25/04 new code down to the 'Pledge info is in the dat table comment
'            slSQLQuery = "SELECT *"
'            slSQLQuery = slSQLQuery + " FROM att"
'            slSQLQuery = slSQLQuery + " WHERE (attShfCode= " & Trim$(rst!attshfCode) & " And attVefCode = " & Trim$(rst!attvefCode) & ")"
'            Set att_rst = gSQLSelectCall(slSQLQuery)
'
'            lgPreviousAttCode = -1
'            llCount = 32000
'            While Not att_rst.EOF
'                If igExportSource = 2 Then
'                    DoEvents
'                End If
'                If att_rst!attCode <> tgCPPosting(iLoop).lAttCode Then
'                    If DateValue(gAdjYear(Trim$(att_rst!attOffAir))) <= DateValue(gAdjYear(Trim$(att_rst!attDropDate))) Then
'                        slTemp2 = Trim$(att_rst!attOffAir)
'                    Else
'                        slTemp2 = Trim$(att_rst!attDropDate)
'                    End If
'                    llTemp = DateDiff("D", Trim$(slTemp2), Trim$(rst!attOnAir))
'                    If (llTemp < llCount) And llTemp > 0 Then
'                        llCount = llTemp
'                        lgPreviousAttCode = att_rst!attCode
'                    End If
'                End If
'                att_rst.MoveNext
'            Wend
            
            ReDim tmATTCrossDates(0 To 0) As ATTCrossDates
            bgAnyAttExclusions = False
            iUpper = 0
            slSQLQuery = "SELECT *"
            slSQLQuery = slSQLQuery + " FROM att"
            slSQLQuery = slSQLQuery + " WHERE (attShfCode= " & Trim$(Str(ilShttCode)) & " And attVefCode = " & Trim$(Str(ilVefCode)) & ")"
            slSQLQuery = slSQLQuery + " Order by attOnAir"
            Set att_rst = gSQLSelectCall(slSQLQuery)
            While Not att_rst.EOF
                If igExportSource = 2 Then
                    DoEvents
                End If
                If ((blCreateServiceATTSpots) Or (att_rst!attServiceAgreement <> "Y")) Then
                    iUpper = UBound(tmATTCrossDates)
                    If DateValue(gAdjYear(Trim$(att_rst!attOffAir))) <= DateValue(gAdjYear(Trim$(att_rst!attDropDate))) Then
                        slTemp2 = Trim$(att_rst!attOffAir)
                    Else
                        slTemp2 = Trim$(att_rst!attDropDate)
                    End If
                    tmATTCrossDates(iUpper).lAttCode = att_rst!attCode
                    tmATTCrossDates(iUpper).lStartDate = DateValue(gAdjYear(att_rst!attOnAir))
                    tmATTCrossDates(iUpper).lEndDate = DateValue(gAdjYear(slTemp2))
                    tmATTCrossDates(iUpper).iLoadFactor = att_rst!attLoad
                    If tmATTCrossDates(iUpper).iLoadFactor < 1 Then
                        tmATTCrossDates(iUpper).iLoadFactor = 1
                    End If
                    tmATTCrossDates(iUpper).sForbidSplitLive = att_rst!attForbidSplitLive
                    tmATTCrossDates(iUpper).iDACode = 1
                    If att_rst!attPledgeType = "D" Then
                        tmATTCrossDates(iUpper).iDACode = 0
                    ElseIf att_rst!attPledgeType = "A" Then
                        tmATTCrossDates(iUpper).iDACode = 1
                    ElseIf att_rst!attPledgeType = "C" Then
                        tmATTCrossDates(iUpper).iDACode = 2
                    End If
                    tmATTCrossDates(iUpper).iComp = att_rst!attComp
                    tmATTCrossDates(iUpper).sServiceAgreement = att_rst!attServiceAgreement
                    '4/3/19
                    tmATTCrossDates(iUpper).sExcludeFillSpot = att_rst!attExcludeFillSpot
                    tmATTCrossDates(iUpper).sExcludeCntrTypeQ = att_rst!attExcludeCntrTypeQ
                    tmATTCrossDates(iUpper).sExcludeCntrTypeR = att_rst!attExcludeCntrTypeR
                    tmATTCrossDates(iUpper).sExcludeCntrTypeT = att_rst!attExcludeCntrTypeT
                    tmATTCrossDates(iUpper).sExcludeCntrTypeM = att_rst!attExcludeCntrTypeM
                    tmATTCrossDates(iUpper).sExcludeCntrTypeS = att_rst!attExcludeCntrTypeS
                    tmATTCrossDates(iUpper).sExcludeCntrTypeV = att_rst!attExcludeCntrTypeV
                    If (igTimes = 3) Or (igTimes = 4) Then
                        If (tmATTCrossDates(iUpper).lEndDate >= lFWkDate) And (tmATTCrossDates(iUpper).lStartDate <= lLWkDate) Then
                            mSetAnyAttExclusions tmATTCrossDates(iUpper)
                            ReDim Preserve tmATTCrossDates(0 To iUpper + 1) As ATTCrossDates
                        End If
                    Else
                        If (tmATTCrossDates(iUpper).lEndDate >= lFWkDate - 1) And (tmATTCrossDates(iUpper).lStartDate <= lLWkDate + 1) Then
                            mSetAnyAttExclusions tmATTCrossDates(iUpper)
                            ReDim Preserve tmATTCrossDates(0 To iUpper + 1) As ATTCrossDates
                        End If
                    End If
                End If
                att_rst.MoveNext
            Wend
            
            att_rst.Close
            Set att_rst = Nothing
            ' Pledge info is in the dat table.
            '1/9/13: Replace following call with call to routine so that Dat can be obtained when attCode changes
            ReDim tlCPDat(0 To 0) As DAT
            tlCPDat(0).lAtfCode = 0
            
            'iUpper = 0
            'slSQLQuery = "SELECT * "
            'slSQLQuery = slSQLQuery + " FROM dat"
            'slSQLQuery = slSQLQuery + " WHERE (datatfCode= " & tgCPPosting(iLoop).lAttCode & ")"
            'Set rst = gSQLSelectCall(slSQLQuery)
            'While Not rst.EOF
            '    If igExportSource = 2 Then
            '        DoEvents
            '    End If
            '    tlCPDat(iUpper).iStatus = 1         'Used
            '    tlCPDat(iUpper).lCode = rst!datCode    '(0).Value
            '    tlCPDat(iUpper).lAtfCode = rst!datAtfCode  '(1).Value
            '    tlCPDat(iUpper).iShfCode = rst!datShfCode  '(2).Value
            '    tlCPDat(iUpper).iVefCode = rst!datVefCode  '(3).Value
            '    'tlCPDat(iUpper).iDACode = rst!datDACode    '(4).Value
            '    tlCPDat(iUpper).iFdDay(0) = rst!datFdMon   '(5).Value
            '    tlCPDat(iUpper).iFdDay(1) = rst!datFdTue   '(6).Value
            '    tlCPDat(iUpper).iFdDay(2) = rst!datFdWed   '(7).Value
            '    tlCPDat(iUpper).iFdDay(3) = rst!datFdThu   '(8).Value
            '    tlCPDat(iUpper).iFdDay(4) = rst!datFdFri   '(9).Value
            '    tlCPDat(iUpper).iFdDay(5) = rst!datFdSat   '(10).Value
            '    tlCPDat(iUpper).iFdDay(6) = rst!datFdSun   '(11).Value
            '    If Second(rst!datFdStTime) <> 0 Then
            '        tlCPDat(iUpper).sFdSTime = Format$(CStr(rst!datFdStTime), sgShowTimeWSecForm)
            '    Else
            '        tlCPDat(iUpper).sFdSTime = Format$(CStr(rst!datFdStTime), sgShowTimeWOSecForm)
            '    End If
            '    If Second(rst!datFdEdTime) <> 0 Then
            '        tlCPDat(iUpper).sFdETime = Format$(CStr(rst!datFdEdTime), sgShowTimeWSecForm)
            '    Else
            '        tlCPDat(iUpper).sFdETime = Format$(CStr(rst!datFdEdTime), sgShowTimeWOSecForm)
            '    End If
            '    tlCPDat(iUpper).iFdStatus = rst!datFdStatus    '(14).Value
            '    tlCPDat(iUpper).iPdDay(0) = rst!datPdMon   '(15).Value
            '    tlCPDat(iUpper).iPdDay(1) = rst!datPdTue   '(16).Value
            '    tlCPDat(iUpper).iPdDay(2) = rst!datPdWed   '(17).Value
            '    tlCPDat(iUpper).iPdDay(3) = rst!datPdThu   '(18).Value
            '    tlCPDat(iUpper).iPdDay(4) = rst!datPdFri   '(19).Value
            '    tlCPDat(iUpper).iPdDay(5) = rst!datPdSat   '(20).Value
            '    tlCPDat(iUpper).iPdDay(6) = rst!datPdSun   '(21).Value
            '    tlCPDat(iUpper).sPdDayFed = rst!datPdDayFed
            '    If tgStatusTypes(tlCPDat(iUpper).iFdStatus).iPledged <= 1 Then
            '        tlCPDat(iUpper).sPdSTime = Format$(CStr(rst!datPdStTime), sgShowTimeWSecForm)
            '        tlCPDat(iUpper).sPdETime = Format$(CStr(rst!datPdEdTime), sgShowTimeWSecForm)
            '    Else
            '        tlCPDat(iUpper).sPdSTime = ""
            '        tlCPDat(iUpper).sPdETime = ""
            '    End If
            '    iUpper = iUpper + 1
            '    ReDim Preserve tlCPDat(0 To iUpper) As DAT
            '    rst.MoveNext
            'Wend
            'mGetDat tgCPPosting(iLoop).lAttCode, tlCPDat()
            
            ' The next set of code is looking in the AST for previously created records.
            ReDim tlAstInfo(0 To 0) As ASTINFO
            ReDim tmAstInfo(0 To 0) As ASTINFO
            ReDim tgDelAst(0 To 0) As ASTINFO
            iUpper = 0
            '4/29/19: test moved to gathering of the att. Instead test if any agreenebt found
            'If ((Not blCreateServiceATTSpots) And (tmATTCrossDates(0).sServiceAgreement = "Y")) Then
            If UBound(tmATTCrossDates) <= LBound(tmATTCrossDates) Then
                If llLockRec > 0 Then
                    llLockRec = gDeleteLockRec_ByRlfCode(llLockRec)
                End If
                mFilterByAdvt tlAstInfo, ilInAdfCode
                mGetAstInfo = True
                Exit Function
            End If
            slServiceAgreement = tmATTCrossDates(0).sServiceAgreement
            'D.S. 10/25 Add FOR loop to code below to get the previous agreement if it exists
            'For ilPass = 0 To 1 Step 1
            For ilPass = 0 To UBound(tmATTCrossDates) - 1 Step 1
                If igExportSource = 2 Then
                    DoEvents
                End If
                'If ilPass = 0 Then
                '    llAtt = tgCPPosting(iLoop).lAttCode
                'Else
                '    llAtt = lgPreviousAttCode
                'End If
                llAtt = tmATTCrossDates(ilPass).lAttCode
                If igTimes = 0 Then
                    ' Retrieve an month for an advertiser.
                    '12/13/13: Obtain Pledge information from DAT
                    ''slSQLQuery = "SELECT astAtfCode, astShfCode, astVefCode, astlsfCode, astAirDate, astAirTime, astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, astPledgeStartTime, astPledgeEndTime, astPledgeStatus, astSdfCode, astCode "
                    'slSQLQuery = "SELECT astAtfCode, astShfCode, astVefCode, astlsfCode, astAirDate, astAirTime, astStatus, astCPStatus, astFeedDate, astFeedTime, astSdfCode, astDatCode, astAdfCode, astCpfCode, astRsfCode, astCntrNo, astCode"
                    slSQLQuery = "SELECT *"
                    slSQLQuery = slSQLQuery + " FROM  ast, lst, ADF_Advertisers, "
                    slSQLQuery = slSQLQuery & "VEF_Vehicles"
                    slSQLQuery = slSQLQuery + " WHERE (astatfCode = " & llAtt
                    ''slSQLQuery = slSQLQuery + " AND Mod(astStatus, 100) <> 22"    '22=Missed part of MG (Status 20)
                    'slSQLQuery = slSQLQuery + " AND Mod(astStatus, 100) < " & ASTEXTENDED_MG   'bypasss MG, Replacement and bonus spots so they're retained
                    slSQLQuery = slSQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')"
                    slSQLQuery = slSQLQuery + " And lstCode = astlsfCode"
                    slSQLQuery = slSQLQuery + " And adfCode = lstAdfCode"
                    If (iAdfCode > 0) Then
                        '6/1/18
                        slSQLQuery = slSQLQuery & " AND astAdfCode = " & iAdfCode  'lstAdfCode = " & iAdfCode
                    End If
                    If llSelGsfCode > 0 Then
                        slSQLQuery = slSQLQuery & " AND lstGsfCode = " & llSelGsfCode
                    End If
                    slSQLQuery = slSQLQuery + " AND vefCode = lstLogVefCode" & ")"
                    slSQLQuery = slSQLQuery + " ORDER BY vefName, adfName, astAirDate, astAirTime"
                Else
                    ' Retrieve a weeks worth.
                    If llSelGsfCode <= 0 Then
                        '12/13/13: Obtain Pledge information from DAT
                        ''slSQLQuery = "SELECT astAtfCode, astShfCode, astVefCode, astlsfCode, astAirDate, astAirTime, astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, astPledgeStartTime, astPledgeEndTime, astPledgeStatus, astSdfCode, astCode FROM ast"
                        'slSQLQuery = "SELECT astAtfCode, astShfCode, astVefCode, astlsfCode, astAirDate, astAirTime, astStatus, astCPStatus, astFeedDate, astFeedTime, astSdfCode, astDatCode, astAdfCode, astCpfCode, astRsfCode, astCode FROM ast"
                        slSQLQuery = "SELECT * FROM ast"
                        slSQLQuery = slSQLQuery + " WHERE (astatfCode= " & llAtt
                        ''slSQLQuery = slSQLQuery + " AND Mod(astStatus, 100) <> 22"    '22=Missed part of MG (Status 20)
                        'slSQLQuery = slSQLQuery + " AND Mod(astStatus, 100) < " & ASTEXTENDED_MG   'bypasss MG, Replacement and bonus spots so they're retained
                        '6/1/18
                        If (iAdfCode > 0) Then
                            slSQLQuery = slSQLQuery & " AND astAdfCode = " & iAdfCode
                        End If
                        slSQLQuery = slSQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
                        If igTimes <> 2 Then
                            slSQLQuery = slSQLQuery + " ORDER BY astAirDate, astAirTime, astCode"
                        Else
                            slSQLQuery = slSQLQuery + " ORDER BY astFeedDate, astFeedTime, astCode"
                        End If
                    Else
                        '12/13/13: Obtain Pledge information from DAT
                        ''slSQLQuery = "SELECT astAtfCode, astShfCode, astVefCode, astlsfCode, astAirDate, astAirTime, astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, astPledgeStartTime, astPledgeEndTime, astPledgeStatus, astSdfCode, astCode"
                        'slSQLQuery = "SELECT astAtfCode, astShfCode, astVefCode, astlsfCode, astAirDate, astAirTime, astStatus, astCPStatus, astFeedDate, astFeedTime, astSdfCode, astDatCode, astAdfCode, astCpfCode, astRsfCode, astCode"
                        slSQLQuery = "SELECT *"
                        slSQLQuery = slSQLQuery + " FROM ast, lst"
                        slSQLQuery = slSQLQuery + " WHERE (astatfCode= " & llAtt
                        ''slSQLQuery = slSQLQuery + " AND Mod(astStatus, 100) <> 22"    '22=Missed part of MG (Status 20)
                        'slSQLQuery = slSQLQuery + " AND Mod(astStatus, 100) < " & ASTEXTENDED_MG   'bypasss MG, Replacement and bonus spots so they're retained
                        slSQLQuery = slSQLQuery + " And lstCode = astlsfCode"
                        slSQLQuery = slSQLQuery & " AND lstGsfCode = " & llSelGsfCode
                        '6/1/18
                        If (iAdfCode > 0) Then
                            slSQLQuery = slSQLQuery & " AND astAdfCode = " & iAdfCode
                        End If
                        
                        slSQLQuery = slSQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
                        If (igTimes <> 2) And (igTimes <> 4) Then
                            slSQLQuery = slSQLQuery + " ORDER BY astAirDate, astAirTime, astCode"
                        Else
                            slSQLQuery = slSQLQuery + " ORDER BY astFeedDate, astFeedTime, astCode"
                        End If
                    End If
                End If
            'Next ilPass
                Set rst_Genl = gSQLSelectCall(slSQLQuery)
                'If (Not rst.EOF) Or (lgPreviousAttCode <= 0) Then
                '    If ilPass = 0 Then
                '        lgPreviousAttCode = -1
                '    End If
                '    Exit For
                'End If
            'Next ilPass
            
                'D.S. tmASTInfo is all of the AST records previously created for an agreement for the specified
                'date and time.
                While Not rst_Genl.EOF
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    
                    '12/13/13: Obtain Pledge information from Dat
                    tlDatPledgeInfo.lAttCode = rst_Genl!astAtfCode
                    tlDatPledgeInfo.lDatCode = rst_Genl!astDatCode
                    tlDatPledgeInfo.iVefCode = rst_Genl!astVefCode
                    tlDatPledgeInfo.sFeedDate = Format(rst_Genl!astFeedDate, "m/d/yy")
                    tlDatPledgeInfo.sFeedTime = Format(rst_Genl!astFeedTime, "hh:mm:ssam/pm")
                    lgSTime7 = timeGetTime
                    ilRet = gGetPledgeGivenAstInfo(tlDatPledgeInfo)
                    lgETime7 = timeGetTime
                    lgTtlTime7 = lgTtlTime7 + (lgETime7 - lgSTime7)
                    'If rst_Genl!astStatus < ASTEXTENDED_MG Then
                    If (gGetAirStatus(rst_Genl!astStatus) < ASTEXTENDED_MG) Or ((sgMissedMGBypass = "Y") And (gGetAirStatus(rst_Genl!astStatus) = ASTAIR_MISSED_MG_BYPASS)) Then
                        tmAstInfo(iUpper).lCode = rst_Genl!astCode
                    Else
                        tmAstInfo(iUpper).lCode = -rst_Genl!astCode
                        tmAstInfo(iUpper).iAdfCode = rst_Genl!astAdfCode
                    End If
                    tmAstInfo(iUpper).lLstCode = rst_Genl!astLsfCode
                    tmAstInfo(iUpper).iStatus = rst_Genl!astStatus
                    tmAstInfo(iUpper).iCPStatus = rst_Genl!astCPStatus
                    tmAstInfo(iUpper).sAirDate = Format$(rst_Genl!astAirDate, sgShowDateForm)
                    If Second(rst_Genl!astAirTime) <> 0 Then
                        tmAstInfo(iUpper).sAirTime = Format$(rst_Genl!astAirTime, sgShowTimeWSecForm)
                    Else
                        tmAstInfo(iUpper).sAirTime = Format$(rst_Genl!astAirTime, sgShowTimeWOSecForm)
                    End If
                    tmAstInfo(iUpper).sFeedDate = Format$(rst_Genl!astFeedDate, sgShowDateForm)
                    If Second(rst_Genl!astFeedTime) <> 0 Then
                        tmAstInfo(iUpper).sFeedTime = Format$(rst_Genl!astFeedTime, sgShowTimeWSecForm)
                    Else
                        tmAstInfo(iUpper).sFeedTime = Format$(rst_Genl!astFeedTime, sgShowTimeWOSecForm)
                    End If
                    '12/13/13: Obtain Pledge information from Dat
                    'tmAstInfo(iUpper).sPledgeDate = Format$(rst!astPledgeDate, sgShowDateForm)
                    'If Second(rst!astPledgeStartTime) <> 0 Then
                    '    tmAstInfo(iUpper).sPledgeStartTime = Format$(rst!astPledgeStartTime, sgShowTimeWSecForm)
                    'Else
                    '    tmAstInfo(iUpper).sPledgeStartTime = Format$(rst!astPledgeStartTime, sgShowTimeWOSecForm)
                    'End If
                    'If Not IsNull(rst!astPledgeEndTime) Then
                    '    If Second(rst!astPledgeEndTime) <> 0 Then
                    '        tmAstInfo(iUpper).sPledgeEndTime = Format$(rst!astPledgeEndTime, sgShowTimeWSecForm)
                    '    Else
                    '        tmAstInfo(iUpper).sPledgeEndTime = Format$(rst!astPledgeEndTime, sgShowTimeWOSecForm)
                    '    End If
                    'Else
                    '    tmAstInfo(iUpper).sPledgeEndTime = ""
                    'End If
                    'tmAstInfo(iUpper).iPledgeStatus = rst!astPledgeStatus
                    
                    tmAstInfo(iUpper).sPledgeDate = Format$(tlDatPledgeInfo.sPledgeDate, sgShowDateForm)
                    
                    '3/8/16: Get MG Pledge date from Missed
                    blMGSpot = gGetMissedPledgeForMG(tmAstInfo(iUpper).iStatus, tmAstInfo(iUpper).sFeedDate, rst_Genl!astLkAstCode, slMGMissedFeedDate, slMGMissedFeedTime)
                    If blMGSpot Then
                        tmAstInfo(iUpper).sPledgeDate = slMGMissedFeedDate
                        If mGetPledgeByEvent(tlDatPledgeInfo.iVefCode) <> "Y" Then
                            If tlDatPledgeInfo.lDatCode <= 0 Then
                                tmAstInfo(iUpper).sPledgeStartTime = slMGMissedFeedTime
                                tmAstInfo(iUpper).sPledgeEndTime = slMGMissedFeedTime
                            Else
                                tmAstInfo(iUpper).sPledgeDate = Format(DateValue(slMGMissedFeedDate) + (DateValue(tlDatPledgeInfo.sPledgeDate) - DateValue(tlDatPledgeInfo.sFeedDate)), "m/d/yy")
                            End If
                        Else
                            tmAstInfo(iUpper).sPledgeStartTime = slMGMissedFeedTime
                            tmAstInfo(iUpper).sPledgeEndTime = slMGMissedFeedTime
                        End If
                    End If
                    
                    If Second(tlDatPledgeInfo.sPledgeStartTime) <> 0 Then
                        tmAstInfo(iUpper).sPledgeStartTime = Format$(tlDatPledgeInfo.sPledgeStartTime, sgShowTimeWSecForm)
                    Else
                        tmAstInfo(iUpper).sPledgeStartTime = Format$(tlDatPledgeInfo.sPledgeStartTime, sgShowTimeWOSecForm)
                    End If
                    If (Not IsNull(tlDatPledgeInfo.sPledgeEndTime)) And (Trim$(tlDatPledgeInfo.sPledgeEndTime) <> "") And (Asc(tlDatPledgeInfo.sPledgeEndTime) <> 0) Then
                        If Second(tlDatPledgeInfo.sPledgeEndTime) <> 0 Then
                            tmAstInfo(iUpper).sPledgeEndTime = Format$(tlDatPledgeInfo.sPledgeEndTime, sgShowTimeWSecForm)
                        Else
                            tmAstInfo(iUpper).sPledgeEndTime = Format$(tlDatPledgeInfo.sPledgeEndTime, sgShowTimeWOSecForm)
                        End If
                    Else
                        tmAstInfo(iUpper).sPledgeEndTime = ""
                    End If
                    tmAstInfo(iUpper).iPledgeStatus = tlDatPledgeInfo.iPledgeStatus
                    
                    tmAstInfo(iUpper).lAttCode = rst_Genl!astAtfCode
                    tmAstInfo(iUpper).iShttCode = rst_Genl!astShfCode
                    tmAstInfo(iUpper).iVefCode = rst_Genl!astVefCode
                    tmAstInfo(iUpper).lSdfCode = rst_Genl!astSdfCode
                    tmAstInfo(iUpper).lDatCode = rst_Genl!astDatCode
                    '11/18/11: Set to blackout that was previous defined as a replacement.
                    '          astLstCode could be referecing the Blackout currently assigned
                    '          each ast will reference a unique blackout lst.  Those lst will reference the original lst from traffic
                    '          the field is set later if required
                    tmAstInfo(iUpper).lPrevBkoutLstCode = 0
                    tmAstInfo(iUpper).lCntrNo = rst_Genl!astCntrNo
                    tmAstInfo(iUpper).iLen = rst_Genl!astLen
                    tmAstInfo(iUpper).lLkAstCode = rst_Genl!astLkAstCode
                    tmAstInfo(iUpper).iMissedMnfCode = rst_Genl!astMissedMnfCode
                    tmAstInfo(iUpper).iComp = tmATTCrossDates(ilPass).iComp
                    tmAstInfo(iUpper).sStationCompliant = rst_Genl!astStationCompliant
                    tmAstInfo(iUpper).sAgencyCompliant = rst_Genl!astAgencyCompliant
                    tmAstInfo(iUpper).sAffidavitSource = gRemoveIllegalChars(rst_Genl!astAffidavitSource)
                    '7/8/20: adding missing lst items
                    
                    iUpper = iUpper + 1
                    ReDim Preserve tmAstInfo(0 To iUpper) As ASTINFO
                    rst_Genl.MoveNext
                Wend
            Next ilPass
            For iAst = 0 To UBound(tmAstInfo) Step 1
                If (tmAstInfo(iAst).iStatus = 20) Then
                    tmAstInfo(iAst).iStatus = ASTEXTENDED_MG
                End If
                If (tmAstInfo(iAst).iStatus = 21) Then
                    tmAstInfo(iAst).iStatus = ASTEXTENDED_BONUS
                End If
            Next iAst
            'Move found images to delete array
            If (igTimes > 0) And (igTimes < 3) Then
                ReDim tgDelAst(0 To UBound(tmAstInfo)) As ASTINFO
                For iAst = 0 To UBound(tmAstInfo) Step 1
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    tgDelAst(iAst) = tmAstInfo(iAst)
                    tgDelAst(iAst).iAdfCode = 0
                    tgDelAst(iAst).iAnfCode = 0
                    tgDelAst(iAst).sProd = ""
                    tgDelAst(iAst).sCart = ""
                    tgDelAst(iAst).sISCI = ""
                    tgDelAst(iAst).lCifCode = 0
                    tgDelAst(iAst).lCrfCsfCode = 0
                    tgDelAst(iAst).lCpfCode = 0
                Next iAst
            End If
            
            'D.S. The LST contains the airing spots that should have aired for the dates specified
            If igTimes = 0 Then
                'slSQLQuery = "SELECT lstProd, lstLogDate, lstLogTime, lstSdfCode, lstStatus, lstAdfCode, lstLogVefCode, lstType, lstZone, lstLen, lstCntrNo, lstCode, lstSplitNetwork, lstRafCode"
                slSQLQuery = "SELECT *"
                slSQLQuery = slSQLQuery + " FROM lst, ADF_Advertisers, "
                slSQLQuery = slSQLQuery & "VEF_Vehicles"
                slSQLQuery = slSQLQuery + " WHERE (adfCode = lstAdfCode"
                If (iAdfCode > 0) Then
                    slSQLQuery = slSQLQuery & " AND lstAdfCode = " & iAdfCode
                End If
                If llSelGsfCode > 0 Then
                    slSQLQuery = slSQLQuery & " AND lstGsfCode = " & llSelGsfCode
                End If
                slSQLQuery = slSQLQuery + " AND vefCode = lstLogVefCode"
                slSQLQuery = slSQLQuery + " AND lstLogVefCode = " & tgCPPosting(iLoop).iVefCode
                slSQLQuery = slSQLQuery + " AND lstBkoutLstCode = 0"
                slSQLQuery = slSQLQuery & " AND lstType <> 1"
                'If Trim$(sZone) <> "" Then
                '    slSQLQuery = slSQLQuery + " AND lstZone = '" & sZone & "'"
                'End If
                slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(lFWkDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(lLWkDate + 1, sgSQLDateForm) & "')" & ")"
                slSQLQuery = slSQLQuery + " ORDER BY vefName, adfName, lstLogDate, lstLogTime"
            Else
                'slSQLQuery = "SELECT lstProd, lstLogDate, lstLogTime, lstSdfCode, lstStatus, lstAdfCode, lstLogVefCode, lstType, lstZone, lstLen, lstCntrNo, lstCode, lstSplitNetwork, lstRafCode FROM lst"
                slSQLQuery = "SELECT * FROM lst"
                slSQLQuery = slSQLQuery + " WHERE (lstLogVefCode = " & tgCPPosting(iLoop).iVefCode
                If (iAdfCode > 0) Then
                    slSQLQuery = slSQLQuery & " AND lstAdfCode = " & iAdfCode
                End If
                If llSelGsfCode > 0 Then
                    slSQLQuery = slSQLQuery & " AND lstGsfCode = " & llSelGsfCode
                End If
                slSQLQuery = slSQLQuery + " AND lstBkoutLstCode = 0"
                slSQLQuery = slSQLQuery & " AND lstType <> 1"
                'If Trim$(sZone) <> "" Then
                '    slSQLQuery = slSQLQuery + " AND lstZone = '" & sZone & "'"
                'End If
                If ((igTimes = 3) Or (igTimes = 4)) And (blFeedAdjOnReturn) Then
                    'lFWkDate and lLWkDate has been adjusted because sFWkDate and sLWkDate have been adjusted by one day
                    slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(lFWkDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(lLWkDate, sgSQLDateForm) & "')" & ")"
                Else
                    slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(lFWkDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(lLWkDate + 1, sgSQLDateForm) & "')" & ")"
                End If
                slSQLQuery = slSQLQuery + " ORDER BY lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
            End If
            If igExportSource = 2 Then
                DoEvents
            End If
            If StrComp(smGetAstKey, slSQLQuery, vbTextCompare) <> 0 Or bgAnyAttExclusions Then
                Set rst_Genl = gSQLSelectCall(slSQLQuery)
                ''D.S. 10/20/08
                ''cover the case where the LST does not have time zone defined, but the vehicle in Traffic does
                'If rst.EOF Then
                '    If igTimes = 0 Then
                '        slSQLQuery = "SELECT *"
                '        slSQLQuery = slSQLQuery + " FROM lst, ADF_Advertisers, "
                '        slSQLQuery = slSQLQuery & "VEF_Vehicles"
                '        slSQLQuery = slSQLQuery + " WHERE (adfCode = lstAdfCode"
                '        If (iAdfCode > 0) Then
                '            slSQLQuery = slSQLQuery & " AND lstAdfCode = " & iAdfCode
                '        End If
                '        slSQLQuery = slSQLQuery + " AND vefCode = lstLogVefCode"
                '        slSQLQuery = slSQLQuery + " AND lstLogVefCode = " & tgCPPosting(iLoop).iVefCode
                '        slSQLQuery = slSQLQuery + " AND lstBkoutLstCode = 0"
                '        slSQLQuery = slSQLQuery & " AND lstType <> 1"
                '        If Trim$(sZone) <> "" Then
                '            slSQLQuery = slSQLQuery + " AND lstZone = ''"
                '        End If
                '        slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(lFWkDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(lLWkDate + 1, sgSQLDateForm) & "')" & ")"
                '        slSQLQuery = slSQLQuery + " ORDER BY vefName, adfName, lstLogDate, lstLogTime"
                '    Else
                '        slSQLQuery = "SELECT * FROM lst"
                '        slSQLQuery = slSQLQuery + " WHERE (lstLogVefCode = " & tgCPPosting(iLoop).iVefCode
                '        If (iAdfCode > 0) Then
                '            slSQLQuery = slSQLQuery & " AND lstAdfCode = " & iAdfCode
                '        End If
                '        slSQLQuery = slSQLQuery + " AND lstBkoutLstCode = 0"
                '        slSQLQuery = slSQLQuery & " AND lstType <> 1"
                '        If Trim$(sZone) <> "" Then
                '            slSQLQuery = slSQLQuery + " AND lstZone = ''"
                '        End If
                '        slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(lFWkDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(lLWkDate + 1, sgSQLDateForm) & "')" & ")"
                '        slSQLQuery = slSQLQuery + " ORDER BY lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
                '    End If
                '    If StrComp(smGetAstKey, slSQLQuery, vbTextCompare) <> 0 Then
                '        Set rst = gSQLSelectCall(slSQLQuery)
                '    End If
                'End If
                'smGetAstKey = slSQLQuery
                'ReDim tmLSTAst(0 To 100) As LST
                'ilUpperLst = 0
                'Do While Not rst.EOF
                '    gCreateUDTforLSTPlusDT rst, tmLSTAst(ilUpperLst)
                '    ilUpperLst = ilUpperLst + 1
                '    If ilUpperLst >= UBound(tmLSTAst) Then
                '        ReDim Preserve tmLSTAst(0 To UBound(tmLSTAst) + 100) As LST
                '    End If
                '    rst.MoveNext
                'Loop
                'ReDim Preserve tmLSTAst(0 To ilUpperLst) As LST
                
                '5/3/19: If using extlusions, get lst each time
                'smGetAstKey = slSQLQuery
                If bgAnyAttExclusions Then
                    smGetAstKey = ""
                Else
                    smGetAstKey = slSQLQuery
                End If
                bmImportSpot = False
                ReDim lmRCSdfCode(0 To 0) As Long
                ReDim tmEZLSTAst(0 To 100) As LSTPLUSDT
                ReDim tmCZLSTAst(0 To 100) As LSTPLUSDT
                ReDim tmMZLSTAst(0 To 100) As LSTPLUSDT
                ReDim tmPZLSTAst(0 To 100) As LSTPLUSDT
                ReDim tmNZLSTAst(0 To 100) As LSTPLUSDT
                llEZUpperLst = 0
                llCZUpperLst = 0
                llMZUpperLst = 0
                llPZUpperLst = 0
                llNZUpperLst = 0
                ilType2Count = 0
                Do While Not rst_Genl.EOF
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    
                    'test if lst should be included
                    If mIncludeLst_ExclusionCheck(rst_Genl, iLocalAdj) Then
                    
                        Select Case Left(rst_Genl!lstZone, 1)
                            Case "E"
                                gCreateUDTforLSTPlusDT rst_Genl, tmEZLSTAst(llEZUpperLst)
                                llEZUpperLst = llEZUpperLst + 1
                                If llEZUpperLst >= UBound(tmEZLSTAst) Then
                                    ReDim Preserve tmEZLSTAst(0 To UBound(tmEZLSTAst) + 100) As LSTPLUSDT
                                End If
                            Case "C"
                                gCreateUDTforLSTPlusDT rst_Genl, tmCZLSTAst(llCZUpperLst)
                                llCZUpperLst = llCZUpperLst + 1
                                If llCZUpperLst >= UBound(tmCZLSTAst) Then
                                    ReDim Preserve tmCZLSTAst(0 To UBound(tmCZLSTAst) + 100) As LSTPLUSDT
                                End If
                            Case "M"
                                gCreateUDTforLSTPlusDT rst_Genl, tmMZLSTAst(llMZUpperLst)
                                llMZUpperLst = llMZUpperLst + 1
                                If llMZUpperLst >= UBound(tmMZLSTAst) Then
                                    ReDim Preserve tmMZLSTAst(0 To UBound(tmMZLSTAst) + 100) As LSTPLUSDT
                                End If
                            Case "P"
                                gCreateUDTforLSTPlusDT rst_Genl, tmPZLSTAst(llPZUpperLst)
                                llPZUpperLst = llPZUpperLst + 1
                                If llPZUpperLst >= UBound(tmPZLSTAst) Then
                                    ReDim Preserve tmPZLSTAst(0 To UBound(tmPZLSTAst) + 100) As LSTPLUSDT
                                End If
                            Case Else
                                gCreateUDTforLSTPlusDT rst_Genl, tmNZLSTAst(llNZUpperLst)
                                llNZUpperLst = llNZUpperLst + 1
                                If llNZUpperLst >= UBound(tmNZLSTAst) Then
                                    ReDim Preserve tmNZLSTAst(0 To UBound(tmNZLSTAst) + 100) As LSTPLUSDT
                                End If
                                If tmNZLSTAst(llNZUpperLst - 1).tLST.iType = 2 Then
                                    ilType2Count = ilType2Count + 1
                                    tmEZLSTAst(llEZUpperLst) = tmNZLSTAst(llNZUpperLst - 1)
                                    llEZUpperLst = llEZUpperLst + 1
                                    If llEZUpperLst >= UBound(tmEZLSTAst) Then
                                        ReDim Preserve tmEZLSTAst(0 To UBound(tmEZLSTAst) + 100) As LSTPLUSDT
                                    End If
                                    tmCZLSTAst(llCZUpperLst) = tmNZLSTAst(llNZUpperLst - 1)
                                    llCZUpperLst = llCZUpperLst + 1
                                    If llCZUpperLst >= UBound(tmCZLSTAst) Then
                                        ReDim Preserve tmCZLSTAst(0 To UBound(tmCZLSTAst) + 100) As LSTPLUSDT
                                    End If
                                    tmMZLSTAst(llMZUpperLst) = tmNZLSTAst(llNZUpperLst - 1)
                                    llMZUpperLst = llMZUpperLst + 1
                                    If llMZUpperLst >= UBound(tmMZLSTAst) Then
                                        ReDim Preserve tmMZLSTAst(0 To UBound(tmMZLSTAst) + 100) As LSTPLUSDT
                                    End If
                                    tmPZLSTAst(llPZUpperLst) = tmNZLSTAst(llNZUpperLst - 1)
                                    llPZUpperLst = llPZUpperLst + 1
                                    If llPZUpperLst >= UBound(tmPZLSTAst) Then
                                        ReDim Preserve tmPZLSTAst(0 To UBound(tmPZLSTAst) + 100) As LSTPLUSDT
                                    End If
                                End If
                        End Select
                        ilFound = False
                        llSdfCode = rst_Genl!lstSdfCode
                        For llLstLoop = 0 To UBound(lmRCSdfCode) - 1 Step 1
                            If igExportSource = 2 Then
                                DoEvents
                            End If
                            If lmRCSdfCode(llLstLoop) = llSdfCode Then
                                ilFound = True
                                Exit For
                            End If
                        Next llLstLoop
                        If Not ilFound Then
                            lmRCSdfCode(UBound(lmRCSdfCode)) = llSdfCode
                            ReDim Preserve lmRCSdfCode(0 To UBound(lmRCSdfCode) + 1) As Long
                        End If
                        If rst_Genl!lstSpotType = 5 And llSdfCode = 0 Then
                            bmImportSpot = True
                        End If
                    End If
                    rst_Genl.MoveNext
                Loop
                If ilType2Count = llEZUpperLst Then
                    llEZUpperLst = 0
                End If
                ReDim Preserve tmEZLSTAst(0 To llEZUpperLst) As LSTPLUSDT
                If ilType2Count = llCZUpperLst Then
                    llCZUpperLst = 0
                End If
                ReDim Preserve tmCZLSTAst(0 To llCZUpperLst) As LSTPLUSDT
                If ilType2Count = llMZUpperLst Then
                    llMZUpperLst = 0
                End If
                ReDim Preserve tmMZLSTAst(0 To llMZUpperLst) As LSTPLUSDT
                If ilType2Count = llPZUpperLst Then
                    llPZUpperLst = 0
                End If
                ReDim Preserve tmPZLSTAst(0 To llPZUpperLst) As LSTPLUSDT
                If (llEZUpperLst <> 0) Or (llCZUpperLst <> 0) Or (llMZUpperLst <> 0) Or (llPZUpperLst <> 0) Then
                    llNZUpperLst = 0
                End If
                ReDim Preserve tmNZLSTAst(0 To llNZUpperLst) As LSTPLUSDT
                'Look for Blackouts
                ReDim tmBkoutLst(0 To 0) As BKOUTLST
                ilPos = InStr(1, slSQLQuery, "lstBkoutLstCode = 0", vbTextCompare)
                If ilPos > 0 Then
                    slSQLQuery = Left(slSQLQuery, ilPos - 1) & "lstBkoutLstCode <> 0" & Mid(slSQLQuery, ilPos + Len("lstBkoutLstCode = 0"))
                    Set lst_rst = gSQLSelectCall(slSQLQuery)
                    Do While Not lst_rst.EOF
                        If igExportSource = 2 Then
                            DoEvents
                        End If
                        gCreateUDTforLST lst_rst, tmBkoutLst(UBound(tmBkoutLst)).tLST
                        tmBkoutLst(UBound(tmBkoutLst)).iDelete = False
                        '10/12/14
                        tmBkoutLst(UBound(tmBkoutLst)).bMatched = False
                        ReDim Preserve tmBkoutLst(0 To UBound(tmBkoutLst) + 1) As BKOUTLST
                        lst_rst.MoveNext
                    Loop
                End If
                If Not blGetRegionCopy Then
                    ReDim tmRegionAssignmentInfo(0 To 0) As REGIONASSIGNMENTINFO
                    ReDim tmRegionDefinitionForSpots(0 To 0) As REGIONDEFINITION
                    ReDim tmSplitCategoryInfoForSpots(0 To 0) As SPLITCATEGORYINFO
                Else
                    mBuildRegionForSpots tgCPPosting(iLoop).iVefCode
                End If
            '1/12/15: Clear the blackout flag (Match)
            Else
                For llLstLoop = 0 To UBound(tmBkoutLst) - 1 Step 1
                    tmBkoutLst(llLstLoop).bMatched = False
                Next llLstLoop
            End If
            
            ReDim tmLSTAst(0 To 0) As LSTPLUSDT
            Select Case Left(sZone, 1)
                Case "E"
                    ReDim tmLSTAst(0 To UBound(tmEZLSTAst)) As LSTPLUSDT
                    For llLstLoop = 0 To UBound(tmEZLSTAst) Step 1
                        tmLSTAst(llLstLoop) = tmEZLSTAst(llLstLoop)
                    Next llLstLoop
                Case "C"
                    ReDim tmLSTAst(0 To UBound(tmCZLSTAst)) As LSTPLUSDT
                    For llLstLoop = 0 To UBound(tmCZLSTAst) Step 1
                        tmLSTAst(llLstLoop) = tmCZLSTAst(llLstLoop)
                    Next llLstLoop
                Case "M"
                    ReDim tmLSTAst(0 To UBound(tmMZLSTAst)) As LSTPLUSDT
                    For llLstLoop = 0 To UBound(tmMZLSTAst) Step 1
                        tmLSTAst(llLstLoop) = tmMZLSTAst(llLstLoop)
                    Next llLstLoop
                Case "P"
                    ReDim tmLSTAst(0 To UBound(tmPZLSTAst)) As LSTPLUSDT
                    For llLstLoop = 0 To UBound(tmPZLSTAst) Step 1
                        tmLSTAst(llLstLoop) = tmPZLSTAst(llLstLoop)
                    Next llLstLoop
            End Select
            If UBound(tmLSTAst) <= 0 Then
                ReDim tmLSTAst(0 To UBound(tmNZLSTAst)) As LSTPLUSDT
                For llLstLoop = 0 To UBound(tmNZLSTAst) Step 1
                    tmLSTAst(llLstLoop) = tmNZLSTAst(llLstLoop)
                Next llLstLoop
            End If
            If (Trim$(sZone) = "") And (UBound(tmLSTAst) <= 0) Then
                If UBound(tmEZLSTAst) > 0 Then
                    ReDim tmLSTAst(0 To UBound(tmEZLSTAst)) As LSTPLUSDT
                    For llLstLoop = 0 To UBound(tmEZLSTAst) Step 1
                        If igExportSource = 2 Then
                            DoEvents
                        End If
                        tmLSTAst(llLstLoop) = tmEZLSTAst(llLstLoop)
                    Next llLstLoop
                ElseIf UBound(tmCZLSTAst) > 0 Then
                    ReDim tmLSTAst(0 To UBound(tmCZLSTAst)) As LSTPLUSDT
                    For llLstLoop = 0 To UBound(tmCZLSTAst) Step 1
                        If igExportSource = 2 Then
                            DoEvents
                        End If
                        tmLSTAst(llLstLoop) = tmCZLSTAst(llLstLoop)
                    Next llLstLoop
                ElseIf UBound(tmMZLSTAst) > 0 Then
                    ReDim tmLSTAst(0 To UBound(tmMZLSTAst)) As LSTPLUSDT
                    For llLstLoop = 0 To UBound(tmMZLSTAst) Step 1
                        If igExportSource = 2 Then
                            DoEvents
                        End If
                        tmLSTAst(llLstLoop) = tmMZLSTAst(llLstLoop)
                    Next llLstLoop
                ElseIf UBound(tmPZLSTAst) > 0 Then
                    ReDim tmLSTAst(0 To UBound(tmPZLSTAst)) As LSTPLUSDT
                    For llLstLoop = 0 To UBound(tmPZLSTAst) Step 1
                        If igExportSource = 2 Then
                            DoEvents
                        End If
                        tmLSTAst(llLstLoop) = tmPZLSTAst(llLstLoop)
                    Next llLstLoop
                End If
            End If
            'ilEof = rst.EOF
            'If Not ilEof Then
            '    gCreateUDTforLSTPlusDT rst, tmLst
            'End If
            ilChkLen = False
            llChkLenRafCode = 0
            ilChkLenPosition = 0
            ilFindFill = False
            blResetFillLen = False
            'While Not ilEof
            ilEof = False
            llLstLoop = LBound(tmLSTAst)
            If llLstLoop < UBound(tmLSTAst) Then
                tmLst = tmLSTAst(llLstLoop).tLST
            Else
                ilEof = True
            End If
            Do While Not ilEof
                '10/9/14: Save as LST replaced with blackout in the compare logic
                slGISCI = tmLst.sISCI
                slGCart = tmLst.sCart
                slGProd = tmLst.sProd
                If igExportSource = 2 Then
                    DoEvents
                End If
                If Not ilChkLen Then
                    ilIssueMoveNext = 0
                    slSplitNetwork = tmLst.sSplitNetwork
                    If (slSplitNetwork = "P") Then
                        'Two kinds of fills:
                        '     Region spot for station not found- Create fill for max length of the regions.  Might need multi-fills to match max length
                        '     Region spot found but need to create fill to pad length to max region spot length
                        '     Multi-Fills:  A fill length of 45 sec might need to be filled with a 30 fill plus a 15 sec fill
                        ilFillLen = tmLst.iLen
                        llChkLenRafCode = tmLst.lRafCode
                        ilChkLenPosition = tmLst.iPositionNo
                        ilIncludeLST = mIncludeSplitNetwork(tmLst.sSplitNetwork, tgCPPosting(iLoop).iShttCode, tmLst.lRafCode, tmLst.iLogVefCode)
                        If Not ilIncludeLST Then
                            'Check if secondard is included
                            tmModelLST = tmLst
                            Do
                                If igExportSource = 2 Then
                                    DoEvents
                                End If
                                'rst.MoveNext
                                'If Not rst.EOF Then
                                llLstLoop = llLstLoop + 1
                                If llLstLoop < UBound(tmLSTAst) Then
                                    'gCreateUDTforLSTPlusDT rst, tmLst
                                    tmLst = tmLSTAst(llLstLoop).tLST
                                    slSplitNetwork = tmLst.sSplitNetwork
                                    If slSplitNetwork = "S" Then
                                        If ilFillLen <> tmLst.iLen Then
                                            ilChkLen = True
                                        End If
                                        If ilFillLen < tmLst.iLen Then
                                            ilFillLen = tmLst.iLen
                                        End If
                                        ilIncludeLST = mIncludeSplitNetwork(tmLst.sSplitNetwork, tgCPPosting(iLoop).iShttCode, tmLst.lRafCode, tmLst.iLogVefCode)
                                        tmModelLST = tmLst
                                        llChkLenRafCode = tmLst.lRafCode
                                        ilChkLenPosition = tmLst.iPositionNo
                                    Else
                                        'No Match found, fill are given Position between spot clusters
                                        'Break Spot  Region   Postion
                                        '  1    1      1      1
                                        '  1    2      2      2
                                        '  1    3      3      3
                                        '
                                        '  1    4      -      5
                                        'Position number added space for fills in LogSub.
                                        '
                                        If slSplitNetwork = "F" Then
                                            If tmLst.lRafCode = 0 Then
                                                llChkLenRafCode = 0
                                                ilIncludeLST = True
                                                If ilFillLen <> tmLst.iLen Then
                                                    ilChkLen = True
                                                End If
                                                ilFindFill = True
                                                'Force value so that other fills will be found
                                                ilFillLen = ilFillLen + tmLst.iLen
                                                blResetFillLen = True
                                            End If
                                        Else
                                            ilIncludeLST = True
                                            ilIssueMoveNext = 1
                                            tmLastLst = tmLst
                                            Exit Do
                                        End If
                                    End If
                                Else
                                    ilIncludeLST = True
                                    ilIssueMoveNext = 2
                                    Exit Do
                                End If
                            Loop While Not ilIncludeLST
                            If ilIssueMoveNext <> 0 Then    'Include not found
                                'Create LST, tmModelLST.iPosition references last spot within set that are combined
                                llChkLenRafCode = 0
                                If Not gCreateSplitFill(ilFillLen, llChkLenRafCode, tmModelLST, tmLst) Then
                                    tmLst = tmModelLST
                                    ilFillLen = ilFillLen + tmLst.iLen
                                    blResetFillLen = True
                                    ilIncludeLST = False
                                    ilChkLen = True
                                Else
                                    ilFillLen = 0
                                    ilChkLen = False
                                End If
                            End If
                        End If
                        If ilIncludeLST And (ilIssueMoveNext = 0) And (slSplitNetwork <> "F") Then
                            Do
                                If igExportSource = 2 Then
                                    DoEvents
                                End If
                                'rst.MoveNext
                                'If Not rst.EOF Then
                                llLstLoop = llLstLoop + 1
                                If llLstLoop < UBound(tmLSTAst) Then
                                    'If rst!lstsplitnetwork = "S" Then
                                    If tmLSTAst(llLstLoop).tLST.sSplitNetwork = "S" Then
                                        If ilFillLen <> tmLst.iLen Then
                                            ilChkLen = True
                                        End If
                                        'If ilFillLen < rst!lstLen Then
                                        If ilFillLen < tmLSTAst(llLstLoop).tLST.iLen Then
                                            'ilFillLen = rst!lstLen
                                            ilFillLen = tmLSTAst(llLstLoop).tLST.iLen
                                        End If
                                        'ilChkLenPosition = rst!lstPositionNo
                                        ilChkLenPosition = tmLSTAst(llLstLoop).tLST.iPositionNo
                                    Else
                                        'If (rst!lstsplitnetwork = "F") And (llChkLenRafCode = rst!lstRafCode) Then
                                        If (tmLSTAst(llLstLoop).tLST.sSplitNetwork = "F") And (llChkLenRafCode = tmLSTAst(llLstLoop).tLST.lRafCode) Then
                                            ilFindFill = True
                                        End If
                                        ilIssueMoveNext = 1
                                        'gCreateUDTforLSTPlusDT rst, tmLastLst
                                        tmLastLst = tmLSTAst(llLstLoop).tLST
                                        Exit Do
                                    End If
                                Else
                                    ilIssueMoveNext = 2
                                    Exit Do
                                End If
                            'Loop While rst!lstsplitnetwork = "S"
                            Loop While tmLSTAst(llLstLoop).tLST.sSplitNetwork = "S"
                        End If
                        'If ilChkLen Then
                        If blResetFillLen Then
                            ilFillLen = ilFillLen - tmLst.iLen
                            blResetFillLen = False
                        End If
                    ElseIf (slSplitNetwork = "S") Then  'Secondary split spot
                        ilIncludeLST = False
                    ElseIf (slSplitNetwork = "F") Then  'Split Fill
                        ilIncludeLST = False
                    Else
                        ilIncludeLST = True
                    End If
                End If
                llCount = llCount + 1
                'If ((iAdfCode = -1) Or (tmLst.iAdfCode = iAdfCode)) And (ilIncludeLST) Then
                If ((iAdfCode <= 0) Or (tmLst.iAdfCode = iAdfCode)) And (ilIncludeLST) Then
                    'Find pledged
                    ilDatStart = 0
                    Do
                        If igExportSource = 2 Then
                            DoEvents
                        End If
                        iFound = False
                        lLogDate = DateValue(gAdjYear(tmLst.sLogDate))
                        lLogTime = gTimeToLong(tmLst.sLogTime, False)
                        
                        l2LogDate = lLogDate
                        l2LogTime = lLogTime
                        
                        If tmLst.iType = 0 Then
                            lLogTime = lLogTime + 3600 * iLocalAdj
                            If lLogTime < 0 Then
                                lLogTime = lLogTime + 86400
                                lLogDate = lLogDate - 1
                            ElseIf lLogTime > 86400 Then
                                lLogTime = lLogTime - 86400
                                lLogDate = lLogDate + 1
                            End If
                        End If
                        sLogTime = Format$(gLongToTime(lLogTime), "hh:mm:ss")
                        sLogDate = Format$(lLogDate, "m/d/yyyy")
                        
                        '1/9/13: Check that correct DAT has been loaded
                        llAttValid = False
                        llAtt = -1
                        '4/29/19: Get service agreement here instead from above
                        slServiceAgreement = "N"
                        For ilAtt = 0 To UBound(tmATTCrossDates) - 1 Step 1
                            'Obtain next agreement if requirted (Sunday into Monday)
                            If ((igTimes = 3) Or (igTimes = 4)) And blFeedAdjOnReturn Then
                                If (l2LogDate = lLogDate + 1) And (Weekday(sLogDate, vbMonday) - 1 = 6) Then
                                    llLogTestDate = l2LogDate
                                Else
                                    llLogTestDate = lLogDate
                                End If
                            Else
                                llLogTestDate = lLogDate
                            End If
                            'If (lLogDate >= tmATTCrossDates(ilAtt).lStartDate) And (lLogDate <= tmATTCrossDates(ilAtt).lEndDate) Then
                            If (llLogTestDate >= tmATTCrossDates(ilAtt).lStartDate) And (llLogTestDate <= tmATTCrossDates(ilAtt).lEndDate) Then
                                llAtt = tmATTCrossDates(ilAtt).lAttCode
                                llAttValid = True
                                If tmATTCrossDates(ilAtt).lEndDate < DateValue(gAdjYear(sFWkDate)) Then
                                    llAttValid = False
                                End If
                                ilLoadFactor = tmATTCrossDates(ilAtt).iLoadFactor
                                slForbidSplitLive = tmATTCrossDates(ilAtt).sForbidSplitLive
                                ilDACode = tmATTCrossDates(ilAtt).iDACode
                                If tlCPDat(0).lAtfCode <> tmATTCrossDates(ilAtt).lAttCode Then
                                    mGetDat tmATTCrossDates(ilAtt).lAttCode, tlCPDat()
                                End If
                                ilAttComp = tmATTCrossDates(ilAtt).iComp
                                slServiceAgreement = tmATTCrossDates(ilAtt).sServiceAgreement
                                Exit For
                            End If
                        Next ilAtt
                        
                        
                        
                        
                        If UBound(tlCPDat) > 0 Then
                            'If tlCPDat(0).iDACode = 2 Then
                            If ilDACode = 2 Then
                                lLogDate = l2LogDate
                                lLogTime = l2LogTime
                                sLogTime = Format$(gLongToTime(lLogTime), "hh:mm:ss")
                                sLogDate = Format$(lLogDate, "m/d/yyyy")
                            End If
                        End If
                        '11/24/11:
                        blLstOk = False
                        If UBound(llGsfCode) <= LBound(llGsfCode) Then
                            If (lLogDate >= lFWkDate) And (lLogDate <= lLWkDate) Then
                                blLstOk = True
                            End If
                        Else
                            For ilGsfLoop = 0 To UBound(llGsfCode) - 1 Step 1
                                If tmLst.lgsfCode = llGsfCode(ilGsfLoop) Then
                                    blLstOk = True
                                    Exit For
                                End If
                            Next ilGsfLoop
                        End If
                        'If (lLogDate >= lFWkDate) And (lLogDate <= lLWkDate) Then
                        If blLstOk Then
                            
                            'correlate the lst with the dat - do we find matching records
                            'lst is the true spot that was on the log and dat is what was pledged
                            iIndex = -1
                            slDayPart = "N"
                            For iDat = ilDatStart To UBound(tlCPDat) - 1 Step 1
                                If igExportSource = 2 Then
                                    DoEvents
                                End If
                                'If tlCPDat(iDat).iDACode = 2 Then
                                '    lLogDate = l2LogDate
                                '    lLogTime = l2LogTime
                                '    sLogTime = Format$(gLongToTime(lLogTime), "hh:mm:ss")
                                '    sLogDate = Format$(lLogDate, "m/d/yyyy")
                                'End If
                                If tlCPDat(iDat).iFdDay(Weekday(sLogDate, vbMonday) - 1) <> 0 Then
                                    lSTime = gTimeToLong(tlCPDat(iDat).sFdSTime, False)
                                    
                                    ''If tlCPDat(iDat).iDACode = 0 Or tlCPDat(iDat).iDACode = 2 Then   'Live Dayparts
                                    'Test end time added for CD Tape (DACode = 2) on 6/22/04
                                    'If tlCPDat(iDat).iDACode = 0 Or tlCPDat(iDat).iDACode = 1 Then   'Live Dayparts
                                        lETime = gTimeToLong(tlCPDat(iDat).sFdETime, True)
                                        If lETime = lSTime Then
                                            lETime = lETime + 1
                                        End If
                                    'Else
                                    '    lETime = lSTime + 1
                                    'End If
                                    
                                    If (lLogTime >= lSTime) And (lLogTime < lETime) Then
                                        'If tgStatusTypes(tlCPDat(iDat).iFdStatus).iPledged <> 2 Then  '2=Not Carried
                                            iFound = True
                                            iIndex = iDat
                                            ilDatStart = iDat + 1
                                            llFdTimeLen = lETime - lSTime
                                            llPdTimeLen = gTimeToLong(tlCPDat(iDat).sPdETime, True) - gTimeToLong(tlCPDat(iDat).sPdSTime, False)
                                            '6/3/15: Add 5 minute rule
                                            'If llPdTimeLen > llFdTimeLen Then
                                            If (llPdTimeLen > llFdTimeLen) Or (((llPdTimeLen = llFdTimeLen)) And (llPdTimeLen > 300)) Then
                                                slDayPart = "Y"
                                            End If
                                        'End If
                                        Exit For
                                    End If
                                End If
                            Next iDat
                            If Not iFound And ilDatStart <> 0 Then
                                Exit Do
                            End If
                            'D.S. Set the pledge date and times
                            'If ((Not iFound) And (UBound(tlCPDat) <= LBound(tlCPDat))) Or (iFound) Then   'Treat as if live broadcast
                                'Need zone adjustment
                                sFdDate = Format$(sLogDate, sgShowDateForm)
                                
                                If Second(sLogTime) <> 0 Then
                                    sFdTime = Format$(sLogTime, sgShowTimeWSecForm)
                                Else
                                    sFdTime = Format$(sLogTime, sgShowTimeWOSecForm)
                                End If
                                sPdDate = ""
                                sPdDays = ""
                                sPdSTime = ""
                                sTPdDays = String(7, "N")
                                ilAdjDay = 0
                                ilAirPlay = 1
                                If iFound Then
                                    llDATCode = tlCPDat(iIndex).lCode
                                    ilAirPlay = tlCPDat(iIndex).iAirPlayNo
                                    iPledged = tgStatusTypes(tlCPDat(iIndex).iFdStatus).iPledged
                                    sStr = ""
                                    If iPledged = 0 Then
                                        Select Case Weekday(sFdDate, vbMonday) - 1
                                            Case 0
                                                sStr = "Mo"
                                            Case 1
                                                sStr = "Tu"
                                            Case 2
                                                sStr = "We"
                                            Case 3
                                                sStr = "Th"
                                            Case 4
                                                sStr = "Fr"
                                            Case 5
                                                sStr = "Sa"
                                            Case 6
                                                sStr = "Su"
                                        End Select
                                    ElseIf iPledged = 1 Then   'Delayed
                                        'If tlCPDat(iIndex).iDACode = 2 Then
                                        If ilDACode = 2 Then
                                             sStr = ""
                                             'The string date will be set later below.
                                        Else
                                            Select Case Weekday(sFdDate, vbMonday) - 1
                                                Case 0
                                                    sStr = "Mo"
                                                Case 1
                                                    sStr = "Tu"
                                                Case 2
                                                    sStr = "We"
                                                Case 3
                                                    sStr = "Th"
                                                Case 4
                                                    sStr = "Fr"
                                                Case 5
                                                    sStr = "Sa"
                                                Case 6
                                                    sStr = "Su"
                                            End Select
                                        End If
                                    ElseIf iPledged = 3 Then   'Delayed
                                        sStr = ""
                                    End If
                                    sPdDays = gDayMap(sStr)
                                    Select Case iPledged
                                        Case 0  'Carried- Live
                                            sPdDate = sFdDate
                                            'sPdSTime = Format$(gLongToTime(gTimeToLong(Format$(sLogTime, "h:mm:ssam/pm"), False) + gTimeToLong(tlCPDat(iIndex).sPdETime, False) - gTimeToLong(tlCPDat(iIndex).sPdSTime, False)), sgShowTimeWSecForm)
                                            sPdSTime = Format(sLogTime, "h:mm:ssa/p")
                                            sTPdETime = Format$(gLongToTime(gTimeToLong(Format$(sLogTime, "h:mm:ssam/pm"), False) + gTimeToLong(tlCPDat(iIndex).sFdETime, False) - gTimeToLong(tlCPDat(iIndex).sFdSTime, False)), sgShowTimeWSecForm)
                                            slDayPart = "N"
                                            '1/13/15: Set Pledge End time equal to Pledge End time specified with the Agreement
                                            ''D.S. 11/01/01 uncommented line below
                                            'sPdETime = ""
                                            sPdETime = sTPdETime
                                            sAirDate = sPdDate  'Format$(rst!lstLogDate, sgShowDateForm)
                                            sAirTime = sPdSTime
                                            Mid(sTPdDays, Weekday(sPdDate, vbMonday), 1) = "Y"
                                        Case 1  'Delayed
                                            'If tlCPDat(iIndex).iDACode = 2 Then
                                            'Jim 3/18/07:  Add Daypart as part of the day test
                                            'If (tlCPDat(iIndex).iDACode = 0) Or (tlCPDat(iIndex).iDACode = 2) Then
                                                For iDay = 0 To 6 Step 1
                                                    If tlCPDat(iIndex).iFdDay(iDay) <> 0 Then
                                                        iFdDay = iDay
                                                        Exit For
                                                    End If
                                                Next iDay
                                                For iDay = 0 To 6 Step 1
                                                    If tlCPDat(iIndex).iPdDay(iDay) <> 0 Then
                                                        iPdDay = iDay
                                                        Exit For
                                                    End If
                                                Next iDay
                                                For iDay = 0 To 6 Step 1
                                                    If tlCPDat(iIndex).iPdDay(iDay) <> 0 Then
                                                        Mid(sTPdDays, iDay + 1, 1) = "Y"
                                                    End If
                                                Next iDay
                                                If (iPdDay >= iFdDay) Then
                                                    ilAdjDay = iPdDay - iFdDay
                                                    If tlCPDat(iIndex).sPdDayFed = "B" Then
                                                        ilAdjDay = ilAdjDay - 7
                                                    End If
                                                    'sAirDate = Format$(DateValue(gAdjYear(sFdDate)) + iPdDay - iFdDay, sgShowDateForm)
                                                Else
                                                    If tlCPDat(iIndex).sPdDayFed = "B" Then
                                                        ilAdjDay = iPdDay - iFdDay
                                                        '11/26/14: Handle Monday to previous Sunday
                                                        If ilAdjDay > 0 Then
                                                            ilAdjDay = ilAdjDay - 1
                                                        End If
                                                        'sAirDate = Format$(DateValue(gAdjYear(sFdDate)) + iPdDay - iFdDay, sgShowDateForm)
                                                    Else
                                                        ilAdjDay = 7 + iPdDay - iFdDay
                                                        'sAirDate = Format$(DateValue(gAdjYear(sFdDate)) + 7 + iPdDay - iFdDay, sgShowDateForm)
                                                    End If
                                                End If
                                                sAirDate = Format$(DateValue(gAdjYear(sFdDate)) + ilAdjDay, sgShowDateForm)
                                                'D.S. 11/17/2003 Assume that tapes will air on one day per feed day
                                                'A one to one relationship on feed days to pledge days
                                                'We are not handling one feed to many pledge days
                                                sPdDate = sAirDate
                                                Select Case Weekday(sPdDate)
                                                    Case vbMonday
                                                        sPdDays = "Mo"
                                                    Case vbTuesday
                                                        sPdDays = "Tu"
                                                    Case vbWednesday
                                                        sPdDays = "We"
                                                    Case vbThursday
                                                        sPdDays = "Th"
                                                    Case vbFriday
                                                        sPdDays = "Fr"
                                                    Case vbSaturday
                                                        sPdDays = "Sa"
                                                    Case vbSunday
                                                        sPdDays = "Su"
                                                End Select
                                            'Else
                                            '    sPdDate = sFdDate
                                            '    sAirDate = sPdDate  'Format$(sLogDate, sgShowDateForm)
                                            'End If
                                            sPdSTime = Format$(tlCPDat(iIndex).sPdSTime, sgShowTimeWSecForm) 'Format$(gTimeToLong(sLogTime, False) + gTimeToLong(tlCPDat(iIndex).sPdSTime, False) - gTimeToLong(tlCPDat(iIndex).sFdSTime, False), sgShowTimeWSecForm)
                                            sPdETime = Format$(tlCPDat(iIndex).sPdETime, sgShowTimeWSecForm) 'Format$(gTimeToLong(sLogTime, False) + gTimeToLong(tlCPDat(iIndex).sPdETime, False) - gTimeToLong(tlCPDat(iIndex).sFdSTime, False), sgShowTimeWSecForm)
                                            sTPdETime = sPdETime
                                            sAirTime = Format$(gLongToTime(gTimeToLong(Format$(sLogTime, "h:mm:ssam/pm"), False) + gTimeToLong(tlCPDat(iIndex).sPdSTime, False) - gTimeToLong(tlCPDat(iIndex).sFdSTime, False)), sgShowTimeWSecForm)
                                        Case 2  'Not Carried
                                            'To avoid sql error set pledge date and pledge time
                                            sPdDate = sFdDate   '""
                                            sPdSTime = sFdTime  '""
                                            sPdETime = sFdTime  '""
                                            sTPdETime = sPdETime
                                            slDayPart = "N"
                                            sAirDate = Format$(sLogDate, sgShowDateForm)
                                            sAirTime = Format$(sLogTime, sgShowTimeWSecForm)
                                        Case 3  'Carried but no pledged dates/times
                                            'To avoid sql error set pledge date and pledge time
                                            sPdDate = sFdDate   '""
                                            sPdSTime = sFdTime  '""
                                            sPdETime = sFdTime  '""
                                            sTPdETime = sPdETime
                                            sTPdDays = String(7, "Y")
                                            sAirDate = sFdDate  'Format$(rst!lstLogDate, sgShowDateForm)
                                            sAirTime = sFdTime
                                            slDayPart = "N"
                                    End Select
                                    'D.S. 09/11/02 Added If statement to support None Aired
                                    '12/8/13: Separate processing CPTT set as Not Aired and Pledge set as Not Carried
                                    'If (iCPStatus = 2) Or (iPledged = 2) Then 'Not Aired
                                    If (iCPStatus = 2) Then  'Not Aired or not carried (if Game and by events)
                                        '12/6/13: If Not a Game and astCPStatus = 2, then set to 4 not 8.
                                        '         If game, see if by event, if not set to 4.  if game and by event, check pet to see status.
                                        '         If not carry then set to 8 otherwise 4.
                                        'iStatus = 8
                                        iStatus = mGetNotAirStatus(tmLst.lgsfCode, ilVefCode, ilShttCode, iPledged)
                                    ElseIf (iPledged = 2) Then
                                        iStatus = 8
                                    Else
                                        'Use lst status unless completed posted or partially posted, then use ast status
                                        If tmLst.iStatus = 0 Then
                                            iStatus = tlCPDat(iIndex).iFdStatus
                                        Else
                                            iStatus = tmLst.iStatus
                                        End If
                                    End If
                                Else
                                    iPledged = 0
                                    llDATCode = 0   'Set so that PledgeStatus will be set to Live
                                    If UBound(tlCPDat) > LBound(tlCPDat) Then
                                        'Set indicator that break does not exist and Records in DAT do exist
                                        'This indicator would indicate to set the PledgeStatus as Not Carrired
                                        llDATCode = -1
                                    End If
                                    sPdDate = Format$(sLogDate, sgShowDateForm)
                                    sPdSTime = Format$(sLogTime, sgShowTimeWSecForm)
                                    sPdETime = Format$(sLogTime, sgShowTimeWSecForm)  '""
                                    llDateTest = gDateValue(sLogDate)
                                    llTimeTest = gTimeToLong(sLogTime, False)
                                    ilAvailLength = 0
                                    For ilAvail = LBound(tmLSTAst) To UBound(tmLSTAst) - 1 Step 1
                                        If (llDateTest = tmLSTAst(ilAvail).lDate) Then
                                            If llTimeTest = tmLSTAst(ilAvail).lTime Then
                                                ilAvailLength = ilAvailLength + tmLSTAst(ilAvail).tLST.iLen
                                            End If
                                        End If
                                    Next ilAvail
                                    sTPdETime = gLongToTime(gTimeToLong(sPdETime, False) + ilAvailLength)
                                    sPdETime = sTPdETime
                                    Mid(sTPdDays, Weekday(sPdDate, vbMonday), 1) = "Y"
                                    slDayPart = "N"
                                    sAirDate = Format$(sLogDate, sgShowDateForm)
                                    sAirTime = Format$(sLogTime, sgShowTimeWSecForm)
                                    Select Case Weekday(sPdDate)
                                        Case vbMonday
                                            sPdDays = "Mo"
                                        Case vbTuesday
                                            sPdDays = "Tu"
                                        Case vbWednesday
                                            sPdDays = "We"
                                        Case vbThursday
                                            sPdDays = "Th"
                                        Case vbFriday
                                            sPdDays = "Fr"
                                        Case vbSaturday
                                            sPdDays = "Sa"
                                        Case vbSunday
                                            sPdDays = "Su"
                                    End Select
                                    'D.S. 09/11/02 Added If statement to support None Aired
                                    If iCPStatus = 2 Then
                                        '12/6/13: If Not a Game and astCPStatus = 2, then set to 4 not 8.
                                        '         If game, see if by event, if not set to 4.  if game and by event, check pet to see status.
                                        '         If not carry then set to 8 otherwise 4.
                                        'iStatus = 8
                                        If UBound(tlCPDat) <= LBound(tlCPDat) Then
                                            iStatus = mGetNotAirStatus(tmLst.lgsfCode, ilVefCode, ilShttCode, iPledged)
                                        Else
                                            iStatus = 8
                                        End If
                                    Else
                                        'this is where there is no pledge information so use lst
                                        'which is from the airing vehicle.  If there was pledge info,
                                        'but not found set to status 8 not carried
                                        If UBound(tlCPDat) <= LBound(tlCPDat) Then
                                            iStatus = tmLst.iStatus '0
                                        Else
                                            '12/6/13: Other breaks defined, but this break is not found.  Therefore, set at Not Carried
                                            'don't call mGetNotAirStatus
                                            iStatus = 8
                                        End If
                                    End If
                                End If
                                If igExportSource = 2 Then
                                    DoEvents
                                End If
                                ''5/20/11:  set air time if required
                                'If (llVpf <> -1) And (tmLst.iType = 0) And (iPledged <> 2) And (ilDACode <> 2) Then
                                '    If (Asc(tgVpfOptions(llVpf).sUsingFeatures1) And EXPORTPOSTEDTIMES) = EXPORTPOSTEDTIMES Then
                                '        slSQLQuery = "Select sdfTime from SDF_Spot_Detail"
                                '        slSQLQuery = slSQLQuery & " Where (sdfCode = " & tmLst.lSdfCode & ")"
                                '        Set sdf_rst = gSQLSelectCall(slSQLQuery)
                                '        If Not sdf_rst.EOF Then
                                '            slPostedAirTime = Format$(sdf_rst!sdfTime, "h:mm:ssam/pm")
                                '            If iIndex >= 0 Then
                                '                slPostedAirTime = Format$(gLongToTime(gTimeToLong(slPostedAirTime, False) + gTimeToLong(tlCPDat(iIndex).sPdSTime, False) - gTimeToLong(tlCPDat(iIndex).sFdSTime, False)), sgShowTimeWSecForm)
                                '            End If
                                '            llPostedAirDate = DateValue(gAdjYear(tmLst.sLogDate))
                                '            llPostedAirTime = gTimeToLong(slPostedAirTime, False) + 3600 * iLocalAdj
                                '            If llPostedAirTime < 0 Then
                                '                llPostedAirTime = llPostedAirTime + 86400
                                '                llPostedAirDate = llPostedAirDate - 1
                                '            ElseIf llPostedAirTime > 86400 Then
                                '                llPostedAirTime = llPostedAirTime - 86400
                                '                llPostedAirDate = llPostedAirDate + 1
                                '            End If
                                '            'Air time and Date only set if ast is new or astCPStatus = 0 (not posted)
                                '            sAirDate = Format$(llPostedAirDate + ilAdjDay, "m/d/yyyy")
                                '            sAirTime = Format$(gLongToTime(llPostedAirTime), sgShowTimeWSecForm)
                                '        End If
                                '    End If
                                'End If
                                If iIndex >= 0 Then
                                    mGetTrafficPostedTimes llVpf, tmLst, iPledged, ilDACode, iCPStatus, iLocalAdj, ilAdjDay, tlCPDat(iIndex).sPdSTime, tlCPDat(iIndex).sFdSTime, sAirDate, sAirTime
                                Else
                                    mGetTrafficPostedTimes llVpf, tmLst, iPledged, ilDACode, iCPStatus, iLocalAdj, ilAdjDay, "12am", "12am", sAirDate, sAirTime
                                End If
                                If Len(sPdSTime) > 0 Then
                                    If Second(sPdSTime) = 0 Then
                                        sPdSTime = Format$(sPdSTime, sgShowTimeWOSecForm)
                                    End If
                                End If
                                If Len(sPdETime) > 0 Then
                                    If Len(Trim$(sPdETime)) <> 0 Then
                                        If Second(sPdETime) = 0 Then
                                            sPdETime = Format$(sPdETime, sgShowTimeWOSecForm)
                                        End If
                                    End If
                                End If
                                If Second(sAirTime) = 0 Then
                                    sAirTime = Format$(sAirTime, sgShowTimeWOSecForm)
                                End If
                                '4/7/16: Handle case where ast re-created and spot has been posted
                                slPostedAirDate = sAirDate
                                slPostedAirTime = sAirTime
                                'D.S. tmASTInfo contains all of the previously created spots for the agreement.
                                'tlASTInfo contains the new spot records to be created or the previously created
                                'AST records that need to be updated.  As each tmASTInfo record is moved into tlASTInfo
                                'the tmASTInfo(...).lCode is set to zero to indicate that this record has
                                'been processed.
                                'D.S. 10/2/03 Added for loop for the load factor
                                '12/9/08:  Save LST so it can be used for each loaf factor as Blackout replaces it
                                tmSvLst = tmLst
                                For ilLoadIdx = 0 To ilLoadFactor - 1 Step 1
                                    '12/8/14: Set air Play number when using Load
                                    If ilLoadFactor > 1 Then
                                        ilAirPlay = ilLoadIdx + 1
                                    End If
                                    If igExportSource = 2 Then
                                        DoEvents
                                    End If
                                    tmLst = tmSvLst
                                    iFound = False
                                    If Not iFound Then
                                        For iAst = 0 To UBound(tmAstInfo) - 1 Step 1
                                            If igExportSource = 2 Then
                                                DoEvents
                                            End If
                                            'D.S. 10/25 Added If statement
                                           'If (tmAstInfo(iAst).lAttCode = tgCPPosting(iLoop).lAttCode) And llAttValid Then
                                            If (tmAstInfo(iAst).lAttCode = llAtt) And llAttValid Then
                                                If ((tmAstInfo(iAst).lSdfCode = tmLst.lSdfCode) And (tmAstInfo(iAst).lCode <> 0) And (tmLst.iType = 0)) Or ((tmAstInfo(iAst).lLstCode = tmLst.lCode) And (tmAstInfo(iAst).lCode <> 0) And (tmLst.iType = 2)) Then
                                                     'Check for LST match, If none found then use the sdf match
                                                     ilSDFMatchOk = True
                                                     llPdSTime = gTimeToLong(sPdSTime, False)
                                                     If Len(sPdETime) = 0 Then
                                                        llPdETime = llPdSTime
                                                     Else
                                                        llPdETime = gTimeToLong(sPdETime, False)
                                                     End If
                                                     If iIndex >= 0 Then
                                                        ilPledgeStatus = tlCPDat(iIndex).iFdStatus
                                                     Else
                                                        'ilPledgeStatus = iStatus
                                                        '10/10/08:  Don't change pledge from Live (no pledge information defined)
                                                        'Pledge exist but not found.  Not sure if the pledge status should be changed
                                                        If UBound(tlCPDat) > 0 Then
                                                            ilPledgeStatus = iStatus
                                                        Else
                                                            '9/30/11: No pledges, use lst status
                                                            ilPledgeStatus = tmLst.iStatus
                                                        End If
                                                     End If
                                                     'If (tmAstInfo(iAst).lLstCode <> tmLst.lCode) Then
                                                     llCompareLST1 = mCompareLST(tmLst.lCode, tmAstInfo(iAst).lLstCode, tmAstInfo(iAst).lSdfCode)
                                                     If llCompareLST1 = 0 Then
                                                        'Look forward to see if a match exist of both sdf and lst
                                                        'if so, then bypass just the sdf match
                                                        '9/25/14: bypass testing the current one again because the pledge could match and the other spots are not checked
                                                        'For iAst1 = iAst To UBound(tmAstInfo) - 1 Step 1
                                                        For iAst1 = iAst + 1 To UBound(tmAstInfo) - 1 Step 1
                                                            If igExportSource = 2 Then
                                                                DoEvents
                                                            End If
                                                            'If (tmAstInfo(iAst1).lAttCode = tgCPPosting(iLoop).lAttCode) And llAttValid Then
                                                            If (tmAstInfo(iAst1).lAttCode = llAtt) And llAttValid Then
                                                                If (tmAstInfo(iAst1).lSdfCode = tmLst.lSdfCode) And (tmAstInfo(iAst1).lCode <> 0) Then
                                                                    'If (tmAstInfo(iAst1).lLstCode = tmLst.lCode) Then
                                                                    llCompareLST2 = mCompareLST(tmLst.lCode, tmAstInfo(iAst1).lLstCode, tmAstInfo(iAst1).lSdfCode)
                                                                    If llCompareLST2 <> 0 Then
                                                                        ilSDFMatchOk = False
                                                                        Exit For
                                                                    Else
                                                                        If (gTimeToLong(tmAstInfo(iAst1).sPledgeStartTime, False) = llPdSTime) And (gTimeToLong(tmAstInfo(iAst1).sPledgeEndTime, False) = llPdETime) And (tmAstInfo(iAst1).iPledgeStatus = ilPledgeStatus) Then
                                                                            '9/25/14: ast bypassed so this test is not necessary
                                                                            'If iAst1 <> iAst Then
                                                                                ilSDFMatchOk = False
                                                                            'End If
                                                                            Exit For
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Next iAst1
                                                     Else    'Check if fits or is they a better fit (duplicate sdf and lst values)
                                                        If (tmLst.iType = 0) Then
                                                            '11/25/13: Bypass test if pledge match
                                                            If (gTimeToLong(tmAstInfo(iAst).sPledgeStartTime, False) <> llPdSTime) Or (gTimeToLong(tmAstInfo(iAst).sPledgeEndTime, False) <> llPdETime) Or (tmAstInfo(iAst).iPledgeStatus <> ilPledgeStatus) Then
                                                                '9/25/14: bypass testing the current one again as test is not necessary
                                                                'For iAst1 = iAst To UBound(tmAstInfo) - 1 Step 1
                                                                For iAst1 = iAst + 1 To UBound(tmAstInfo) - 1 Step 1
                                                                    If igExportSource = 2 Then
                                                                        DoEvents
                                                                    End If
                                                                    '11/25/13: Add test of sdf
                                                                    'If (gTimeToLong(tmAstInfo(iAst1).sPledgeStartTime, False) = llPdSTime) And (gTimeToLong(tmAstInfo(iAst1).sPledgeEndTime, False) = llPdETime) And (tmAstInfo(iAst1).iPledgeStatus = ilPledgeStatus) Then
                                                                    '    If iAst1 <> iAst Then
                                                                    '        ilSDFMatchOk = False
                                                                    '    End If
                                                                    '    Exit For
                                                                    'End If
                                                                    If (tmAstInfo(iAst1).lAttCode = llAtt) And llAttValid Then
                                                                        If (tmAstInfo(iAst1).lSdfCode = tmLst.lSdfCode) And (tmAstInfo(iAst1).lCode <> 0) Then
                                                                            If (gTimeToLong(tmAstInfo(iAst1).sPledgeStartTime, False) = llPdSTime) And (gTimeToLong(tmAstInfo(iAst1).sPledgeEndTime, False) = llPdETime) And (tmAstInfo(iAst1).iPledgeStatus = ilPledgeStatus) Then
                                                                                '9/25/14: ast bypassed so this test is not necessary
                                                                                'If iAst1 <> iAst Then
                                                                                    ilSDFMatchOk = False
                                                                                'End If
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    End If
                                                                Next iAst1
                                                            End If
                                                        End If
                                                     End If
                                                     '3/16/11: Add test that AST and DAT match
                                                     '         test Feed start time, Pledge Start time, feed day and pledge day
                                                     If (iIndex >= 0) And (ilSDFMatchOk) Then
                                                        '4/7/16: Handle case where ast re-created and spot has been posted
                                                        If tmAstInfo(iAst).iCPStatus = 1 Then
                                                            slPostedAirDate = tmAstInfo(iAst).sAirDate
                                                            slPostedAirTime = tmAstInfo(iAst).sAirTime
                                                        End If
                                                        If iPledged = 0 Then
                                                            ilSDFMatchOk = False
                                                            If (gTimeToLong(tmAstInfo(iAst).sFeedTime, False) = gTimeToLong(tlCPDat(iIndex).sFdSTime, False)) And (gTimeToLong(tmAstInfo(iAst).sPledgeStartTime, False) = gTimeToLong(tlCPDat(iIndex).sPdSTime, False)) Then
                                                                If tlCPDat(iIndex).iFdDay(Weekday(tmAstInfo(iAst).sFeedDate, vbMonday) - 1) <> 0 Then
                                                                    If tlCPDat(iIndex).iPdDay(Weekday(tmAstInfo(iAst).sPledgeDate, vbMonday) - 1) <> 0 Then
                                                                        ilSDFMatchOk = True
                                                                    End If
                                                                End If
                                                            End If
                                                        ElseIf iPledged = 1 Then
                                                            ilSDFMatchOk = False
                                                            If (gTimeToLong(tmAstInfo(iAst).sPledgeStartTime, False) = gTimeToLong(tlCPDat(iIndex).sPdSTime, False)) Then
                                                                If tlCPDat(iIndex).iFdDay(Weekday(tmAstInfo(iAst).sFeedDate, vbMonday) - 1) <> 0 Then
                                                                    If tlCPDat(iIndex).iPdDay(Weekday(tmAstInfo(iAst).sPledgeDate, vbMonday) - 1) <> 0 Then
                                                                        ilSDFMatchOk = True
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                     End If
                                                     If ilSDFMatchOk Then
                                                        If llCompareLST1 > 1 Then
                                                            '11/18/11:  retain BkoutLstCode so that in gGetRegionCopy the blackout can be found
                                                            tmAstInfo(iAst).lPrevBkoutLstCode = tmBkoutLst(llCompareLST1 - 2).tLST.lCode
                                                            tmLst = tmBkoutLst(llCompareLST1 - 2).tLST
                                                            '10/12/14
                                                            tmBkoutLst(llCompareLST1 - 2).bMatched = True
                                                        End If
                                                        lAstCode = tmAstInfo(iAst).lCode
                                                        tmAstInfo(iAst).lDatCode = llDATCode
                                                        'D.S. 2/23/05 If it's posted (iCPStatus = 1) then don't set the date or time back to the pledge info
                                                        If tmAstInfo(iAst).iCPStatus = 0 Then
                                                            tmAstInfo(iAst).sAirDate = sAirDate
                                                            tmAstInfo(iAst).sAirTime = sAirTime
                                                        End If
                                                        tmAstInfo(iAst).sFeedDate = sFdDate
                                                        tmAstInfo(iAst).sFeedTime = sFdTime
                                                        tmAstInfo(iAst).iAdfCode = tmLst.iAdfCode
                                                        'If IsNull(rst!lstProd) Then
                                                        '    tmAstInfo(iAst).sProd = ""
                                                        'Else
                                                            tmAstInfo(iAst).sProd = tmLst.sProd
                                                        'End If
                                                        tmAstInfo(iAst).iAnfCode = tmLst.iAnfCode
                                                        tmAstInfo(iAst).sCart = tmLst.sCart
                                                        tmAstInfo(iAst).sISCI = tmLst.sISCI
                                                        tmAstInfo(iAst).lCifCode = tmLst.lCifCode
                                                        tmAstInfo(iAst).lCrfCsfCode = tmLst.lCrfCsfCode
                                                        tmAstInfo(iAst).lCpfCode = tmLst.lCpfCode
                                                        tmAstInfo(iAst).iVefCode = tmLst.iLogVefCode
                                                        tmAstInfo(iAst).sPdDays = sPdDays
                                                        'D.L. 3/17/21 retain ast length
                                                        'tmAstInfo(iAst).iLen = tmLst.iLen
                                                        tmAstInfo(iAst).lgsfCode = tmLst.lgsfCode
                                                        tmAstInfo(iAst).iLstLnVefCode = tmLst.iLnVefCode
                                                        tmAstInfo(iAst).lLstBkoutLstCode = tmLst.lBkoutLstCode
                                                        '10/9/14: Retain Generic ISCI
                                                        tmAstInfo(iAst).sGISCI = slGISCI
                                                        tmAstInfo(iAst).sGCart = slGCart
                                                        tmAstInfo(iAst).sGProd = slGProd
                                                        '11/2/16: Add Event ID information. Used with XDS ProgramCode:Cue
                                                        tmAstInfo(iAst).lEvtIDCefCode = tmLst.lEvtIDCefCode
                                                        
                                                        tmAstInfo(iAst).sLstStartDate = tmLst.sStartDate
                                                        tmAstInfo(iAst).sLstEndDate = tmLst.sEndDate
                                                        tmAstInfo(iAst).iLstSpotsWk = tmLst.iSpotsWk
                                                        tmAstInfo(iAst).iLstMon = tmLst.iMon
                                                        tmAstInfo(iAst).iLstTue = tmLst.iTue
                                                        tmAstInfo(iAst).iLstWed = tmLst.iWed
                                                        tmAstInfo(iAst).iLstThu = tmLst.iThu
                                                        tmAstInfo(iAst).iLstFri = tmLst.iFri
                                                        tmAstInfo(iAst).iLstSat = tmLst.iSat
                                                        tmAstInfo(iAst).iLstSun = tmLst.iSun
                                                        tmAstInfo(iAst).iLineNo = tmLst.iLineNo
                                                        tmAstInfo(iAst).iSpotType = tmLst.iSpotType
                                                        tmAstInfo(iAst).sSplitNet = tmLst.sSplitNetwork
                                                        tmAstInfo(iAst).iAirPlay = ilAirPlay
                                                        tmAstInfo(iAst).iAgfCode = tmLst.iAgfCode
                                                        tmAstInfo(iAst).iComp = ilAttComp
                                                        tmAstInfo(iAst).sLstLnStartTime = tmLst.sLnStartTime
                                                        tmAstInfo(iAst).sLstLnEndTime = tmLst.sLnEndTime
                                                        tmAstInfo(iAst).sEmbeddedOrROS = ""
                                                        If iIndex >= 0 Then
                                                            tmAstInfo(iAst).sEmbeddedOrROS = tlCPDat(iIndex).sEmbeddedOrROS
                                                        End If
                                                        If Trim$(tmAstInfo(iAst).sEmbeddedOrROS) = "" Then
                                                            tmAstInfo(iAst).sEmbeddedOrROS = smDefaultEmbeddedOrROS
                                                        End If
                                                        If iAst <= UBound(tgDelAst) Then
                                                            If (igTimes > 0) And (igTimes < 3) Then
                                                                tgDelAst(iAst).iAnfCode = tmLst.iAnfCode
                                                                tgDelAst(iAst).sCart = tmLst.sCart
                                                                tgDelAst(iAst).sISCI = tmLst.sISCI
                                                                tgDelAst(iAst).lCifCode = tmLst.lCifCode
                                                                tgDelAst(iAst).lCrfCsfCode = tmLst.lCrfCsfCode
                                                                tgDelAst(iAst).lCpfCode = tmLst.lCpfCode
                                                                tgDelAst(iAst).iVefCode = tmLst.iLogVefCode
                                                                tgDelAst(iAst).sPdDays = sPdDays
                                                                tgDelAst(iAst).iLen = tmLst.iLen
                                                                tgDelAst(iAst).lgsfCode = tmLst.lgsfCode
                                                            End If
                                                        End If
                                                        'tmAstInfo(iAst).lLstCode = rst!lstCode
                                                        '12/17/07(TTP 3109):  Adjust Pledge if changed.
                                                        ilPledgeMatch = True
                                                        If iIndex >= 0 Then
                                                            If tmAstInfo(iAst).iPledgeStatus <> tlCPDat(iIndex).iFdStatus Then
                                                                ilPledgeMatch = False
                                                            End If
                                                        End If
                                                        'If (Not iAddAst) Or (tmAstInfo(iAst).lLstCode <> tmLst.lCode) Or (Not ilPledgeMatch) Then  '8-1-06 insure the ast records are recreated when lst pointers out of sync
                                                        'llCompareLST = mCompareLST(tmLst.lCode, tmAstInfo(iAst).lLstCode)
                                                        If (Not iAddAst) Or (llCompareLST1 = 0) Or (Not ilPledgeMatch) Then  '8-1-06 insure the ast records are recreated when lst pointers out of sync
                                                            '6/29/06: change mGetAstInfo to use API call
                                                            'slSQLQuery = "UPDATE ast SET "
                                                            'slSQLQuery = slSQLQuery + "astlsfCode = " & rst!lstCode
                                                            'slSQLQuery = slSQLQuery + " WHERE (astCode = " & lAstCode & ")"
                                                            'cnn.BeginTrans
                                                            'cnn.Execute slSQLQuery, rdExecDirect
                                                            'cnn.CommitTrans
                                                            tlAstSrchKey.lCode = lAstCode
                                                            ilRet = btrGetEqual(hlAst, tlAst, imAstRecLen, tlAstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
                                                            If ilRet = BTRV_ERR_NONE Then
                                                                tlAst.lLsfCode = tmLst.lCode
                                                                '12/7/13: update and Retain status
                                                                'If iIndex >= 0 Then
                                                                '    If (tmAstInfo(iAst).iCPStatus <> 1) And (iCPStatus <> 1) Then
                                                                '        tmAstInfo(iAst).iPledgeStatus = tlCPDat(iIndex).iFdStatus
                                                                '        tlAst.iStatus = iStatus
                                                                '    End If
                                                                'End If
                                                                If (tmAstInfo(iAst).iCPStatus <> 1) And (iCPStatus <> 1) Then
                                                                    If iIndex >= 0 Then
                                                                        tmAstInfo(iAst).iPledgeStatus = tlCPDat(iIndex).iFdStatus
                                                                        tlAst.iStatus = iStatus
                                                                        '12/9/13
                                                                        'tlAst.iPledgeStatus = tlCPDat(iIndex).iFdStatus
                                                                        tmAstInfo(iAst).iStatus = iStatus
                                                                    Else
                                                                        tmAstInfo(iAst).iPledgeStatus = iStatus
                                                                        tlAst.iStatus = iStatus
                                                                        '12/9/13
                                                                        'tlAst.iPledgeStatus = iStatus
                                                                        tmAstInfo(iAst).iStatus = iStatus
                                                                    End If
                                                                End If
                                                                ilRet = btrUpdate(hlAst, tlAst, imAstRecLen)
                                                            End If
                                                            '6/29/06: End of change
                                                        End If
                                                        tmAstInfo(iAst).lLstCode = tmLst.lCode
                                                        tmAstInfo(iAst).lCntrNo = tmLst.lCntrNo
                                                        tmAstInfo(iAst).lSdfCode = tmLst.lSdfCode
                                                        'If iCPStatus = 1 Then   'Complete or partial posted use ast status instead of lst
                                                        If (tmAstInfo(iAst).iCPStatus = 1) Or (iCPStatus = 1) Then
                                                            iStatus = tmAstInfo(iAst).iStatus
                                                        End If
                                                        iFound = True
                                                        iUpper = iAst
                                                        If igExportSource = 2 Then
                                                            DoEvents
                                                        End If
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        Next iAst
                                    End If
    
                                    If Not iFound Then 'did the AST exist
                                        If tmLst.iType = 0 And llAttValid Then
                                            If igExportSource = 2 Then
                                                DoEvents
                                            End If
                                            iUpper = UBound(tlAstInfo)
                                            tlAstInfo(iUpper).lCode = 0
                                            tlAstInfo(iUpper).lAttCode = tgCPPosting(iLoop).lAttCode 'tlCPDat(iIndex).lAtfCode
                                            tlAstInfo(iUpper).iShttCode = tgCPPosting(iLoop).iShttCode 'tlCPDat(iIndex).iShfCode
                                            tlAstInfo(iUpper).iVefCode = tgCPPosting(iLoop).iVefCode 'tlCPDat(iIndex).iVefCode
                                            tlAstInfo(iUpper).lSdfCode = tmLst.lSdfCode
                                            tlAstInfo(iUpper).lLstCode = tmLst.lCode
                                            tlAstInfo(iUpper).lDatCode = llDATCode
                                            tlAstInfo(iUpper).iStatus = iStatus
                                            tlAstInfo(iUpper).sAirDate = sAirDate
                                            tlAstInfo(iUpper).sAirTime = sAirTime
                                            tlAstInfo(iUpper).sFeedDate = sFdDate
                                            tlAstInfo(iUpper).sFeedTime = sFdTime
                                            tlAstInfo(iUpper).sPledgeDate = sPdDate
                                            tlAstInfo(iUpper).sPledgeStartTime = sPdSTime
                                            tlAstInfo(iUpper).sPledgeEndTime = sPdETime
                                            tlAstInfo(iUpper).sTruePledgeEndTime = sTPdETime
                                            tlAstInfo(iUpper).sTruePledgeDays = sTPdDays
                                            tlAstInfo(iUpper).sPdTimeExceedsFdTime = slDayPart
                                            If iIndex >= 0 Then
                                                tlAstInfo(iUpper).iPledgeStatus = tlCPDat(iIndex).iFdStatus
                                                tlAstInfo(iUpper).sPdDayFed = tlCPDat(iIndex).sPdDayFed
                                            Else
                                                'tlAstInfo(iUpper).iPledgeStatus = iStatus
                                                '10/10/08:  Don't change pledge from Live (no pledge information defined)
                                                'Pledge exist but not found.  Not sure if the pledge status should be changed
                                                If UBound(tlCPDat) > 0 Then
                                                    tlAstInfo(iUpper).iPledgeStatus = iStatus
                                                Else
                                                    '9/30/11: No pledges, use lst status
                                                    tlAstInfo(iUpper).iPledgeStatus = tmLst.iStatus
                                                End If
                                                tlAstInfo(iUpper).sPdDayFed = ""
                                            End If
                                            tlAstInfo(iUpper).iAdfCode = tmLst.iAdfCode
                                            'If IsNull(rst!lstProd) Then
                                            '    tlAstInfo(iUpper).sProd = ""
                                            'Else
                                                tlAstInfo(iUpper).sProd = tmLst.sProd
                                            'End If
                                            tlAstInfo(iUpper).iAnfCode = tmLst.iAnfCode
                                            tlAstInfo(iUpper).sCart = tmLst.sCart
                                            tlAstInfo(iUpper).sISCI = tmLst.sISCI
                                            tlAstInfo(iUpper).lCifCode = tmLst.lCifCode
                                            tlAstInfo(iUpper).lCrfCsfCode = tmLst.lCrfCsfCode
                                            tlAstInfo(iUpper).lCpfCode = tmLst.lCpfCode
                                            tlAstInfo(iUpper).sPdDays = sPdDays
                                            If slServiceAgreement = "Y" Then
                                                tlAstInfo(iUpper).iCPStatus = 1
                                            Else
                                                If iCPStatus <> -1 Then
                                                    tlAstInfo(iUpper).iCPStatus = iCPStatus
                                                    '4/7/16: Handle case where ast re-created and spot has been posted
                                                    tlAstInfo(iUpper).sAirDate = slPostedAirDate
                                                    tlAstInfo(iUpper).sAirTime = slPostedAirTime
                                                Else
                                                    tlAstInfo(iUpper).iCPStatus = 0 'Not Received
                                                End If
                                            End If
                                            tlAstInfo(iUpper).sLstZone = tmLst.sZone
                                            tlAstInfo(iUpper).iLen = tmLst.iLen
                                            tlAstInfo(iUpper).lgsfCode = tmLst.lgsfCode
                                            tlAstInfo(iUpper).lCntrNo = tmLst.lCntrNo
                                            tlAstInfo(iUpper).iLstLnVefCode = tmLst.iLnVefCode
                                            tlAstInfo(iUpper).lLstBkoutLstCode = tmLst.lBkoutLstCode
                                            '10/10/14: Retain Generic ISCI
                                            tlAstInfo(iUpper).sGISCI = slGISCI
                                            tlAstInfo(iUpper).sGCart = slGCart
                                            tlAstInfo(iUpper).sGProd = slGProd
                                            '11/2/16: Add Event ID information. Used with XDS ProgramCode:Cue
                                            tlAstInfo(iUpper).lEvtIDCefCode = tmLst.lEvtIDCefCode
                                            
                                            tlAstInfo(iUpper).sLstStartDate = tmLst.sStartDate
                                            tlAstInfo(iUpper).sLstEndDate = tmLst.sEndDate
                                            tlAstInfo(iUpper).iLstSpotsWk = tmLst.iSpotsWk
                                            tlAstInfo(iUpper).iLstMon = tmLst.iMon
                                            tlAstInfo(iUpper).iLstTue = tmLst.iTue
                                            tlAstInfo(iUpper).iLstWed = tmLst.iWed
                                            tlAstInfo(iUpper).iLstThu = tmLst.iThu
                                            tlAstInfo(iUpper).iLstFri = tmLst.iFri
                                            tlAstInfo(iUpper).iLstSat = tmLst.iSat
                                            tlAstInfo(iUpper).iLstSun = tmLst.iSun
                                            tlAstInfo(iUpper).iLineNo = tmLst.iLineNo
                                            tlAstInfo(iUpper).iSpotType = tmLst.iSpotType
                                            tlAstInfo(iUpper).sSplitNet = tmLst.sSplitNetwork
                                            tlAstInfo(iUpper).lLkAstCode = 0
                                            tlAstInfo(iUpper).iMissedMnfCode = 0
                                            tlAstInfo(iUpper).iAirPlay = ilAirPlay
                                            tlAstInfo(iUpper).iAgfCode = tmLst.iAgfCode
                                            tlAstInfo(iUpper).iComp = ilAttComp
                                            tlAstInfo(iUpper).sLstLnStartTime = tmLst.sLnStartTime
                                            tlAstInfo(iUpper).sLstLnEndTime = tmLst.sLnEndTime
                                            tlAstInfo(iUpper).sEmbeddedOrROS = ""
                                            If iIndex >= 0 Then
                                                tlAstInfo(iUpper).sEmbeddedOrROS = tlCPDat(iIndex).sEmbeddedOrROS
                                            End If
                                            If Trim$(tlAstInfo(iUpper).sEmbeddedOrROS) = "" Then
                                                tlAstInfo(iUpper).sEmbeddedOrROS = smDefaultEmbeddedOrROS
                                            End If
                                            tlAstInfo(iUpper).sStationCompliant = ""
                                            tlAstInfo(iUpper).sAgencyCompliant = ""
                                            tlAstInfo(iUpper).sAffidavitSource = ""
                                            ReDim Preserve tlAstInfo(0 To iUpper + 1) As ASTINFO
                                        End If
                                    Else
                                        If igExportSource = 2 Then
                                            DoEvents
                                        End If
                                        tlAstInfo(UBound(tlAstInfo)) = tmAstInfo(iUpper)
                                        tmAstInfo(iUpper).lCode = 0 'Indicate to ignore record
                                        iUpper = UBound(tlAstInfo)
                                        ReDim Preserve tlAstInfo(0 To iUpper + 1) As ASTINFO
                                        'If tlAstInfo(iUpper).iStatus < 20 Then
                                        If (gIsAstStatus(tlAstInfo(iUpper).iStatus, ASTEXTENDED_MG) = False) And (gIsAstStatus(tlAstInfo(iUpper).iStatus, ASTEXTENDED_BONUS) = False) And (gIsAstStatus(tlAstInfo(iUpper).iStatus, ASTEXTENDED_REPLACEMENT) = False) Then
                                            If iUpdateCpttStatus Then
                                                If (tlAstInfo(iUpper).iCPStatus <> 1) Or (iCPStatus <> -1) Then
                                                    If (tlAstInfo(iUpper).iCPStatus <> iCPStatus) Then
                                                        '6/29/06: change mGetAstInfo to use API call
                                                        'slSQLQuery = "UPDATE ast SET "
                                                        'slSQLQuery = slSQLQuery + "astCPStatus = " & iCPStatus
                                                        'If iCPStatus = 2 Then
                                                        '    slSQLQuery = slSQLQuery + ", " & "astStatus = " & iStatus
                                                        '    tlAstInfo(iUpper).iStatus = iStatus
                                                        'End If
                                                        'slSQLQuery = slSQLQuery + " WHERE (astCode = " & lAstCode & ")"
                                                        'cnn.BeginTrans
                                                        'cnn.Execute slSQLQuery, rdExecDirect
                                                        'cnn.CommitTrans
                                                        tlAstSrchKey.lCode = lAstCode
                                                        ilRet = btrGetEqual(hlAst, tlAst, imAstRecLen, tlAstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            If slServiceAgreement = "Y" Then
                                                                tlAst.iCPStatus = 1
                                                            Else
                                                                If iCPStatus <> -1 Then
                                                                    tlAst.iCPStatus = iCPStatus
                                                                End If
                                                                If iCPStatus = 2 Then
                                                                    tlAstInfo(iUpper).iStatus = iStatus
                                                                    tlAst.iStatus = iStatus
                                                                End If
                                                            End If
                                                            ilRet = btrUpdate(hlAst, tlAst, imAstRecLen)
                                                        End If
                                                        '6/29/06: End of change
                                                        If slServiceAgreement = "Y" Then
                                                            tlAstInfo(iUpper).iCPStatus = 1
                                                        Else
                                                            If iCPStatus <> -1 Then
                                                                tlAstInfo(iUpper).iCPStatus = iCPStatus
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            If igExportSource = 2 Then
                                                DoEvents
                                            End If
                                            tlAstInfo(iUpper).lDatCode = llDATCode
                                            tlAstInfo(iUpper).iAdfCode = tmLst.iAdfCode
                                            tlAstInfo(iUpper).iAnfCode = tmLst.iAnfCode
                                            tlAstInfo(iUpper).sCart = tmLst.sCart
                                            tlAstInfo(iUpper).sISCI = tmLst.sISCI
                                            tlAstInfo(iUpper).lCifCode = tmLst.lCifCode
                                            tlAstInfo(iUpper).lCrfCsfCode = tmLst.lCrfCsfCode
                                            tlAstInfo(iUpper).lCpfCode = tmLst.lCpfCode
                                            tlAstInfo(iUpper).sLstZone = tmLst.sZone
                                            'D.L. 3/17/21 retain the ast length
                                            'tlAstInfo(iUpper).iLen = tmLst.iLen
                                            tlAstInfo(iUpper).lgsfCode = tmLst.lgsfCode
                                            tlAstInfo(iUpper).lCntrNo = tmLst.lCntrNo
                                            tlAstInfo(iUpper).sPledgeDate = sPdDate
                                            tlAstInfo(iUpper).sPledgeStartTime = sPdSTime
                                            tlAstInfo(iUpper).sPledgeEndTime = sPdETime
                                            tlAstInfo(iUpper).sTruePledgeEndTime = sTPdETime
                                            tlAstInfo(iUpper).sTruePledgeDays = sTPdDays
                                            tlAstInfo(iUpper).sPdTimeExceedsFdTime = slDayPart
                                            tlAstInfo(iUpper).iLstLnVefCode = tmLst.iLnVefCode
                                            tlAstInfo(iUpper).lLstBkoutLstCode = tmLst.lBkoutLstCode
                                            '10/10/14: Retain Generic ISCI
                                            tlAstInfo(iUpper).sGISCI = slGISCI
                                            tlAstInfo(iUpper).sGCart = slGCart
                                            tlAstInfo(iUpper).sGProd = slGProd
                                            '11/2/16: Add Event ID information. Used with XDS ProgramCode:Cue
                                            tlAstInfo(iUpper).lEvtIDCefCode = tmLst.lEvtIDCefCode
                                            
                                            tlAstInfo(iUpper).sLstStartDate = tmLst.sStartDate
                                            tlAstInfo(iUpper).sLstEndDate = tmLst.sEndDate
                                            tlAstInfo(iUpper).iLstSpotsWk = tmLst.iSpotsWk
                                            tlAstInfo(iUpper).iLstMon = tmLst.iMon
                                            tlAstInfo(iUpper).iLstTue = tmLst.iTue
                                            tlAstInfo(iUpper).iLstWed = tmLst.iWed
                                            tlAstInfo(iUpper).iLstThu = tmLst.iThu
                                            tlAstInfo(iUpper).iLstFri = tmLst.iFri
                                            tlAstInfo(iUpper).iLstSat = tmLst.iSat
                                            tlAstInfo(iUpper).iLstSun = tmLst.iSun
                                            tlAstInfo(iUpper).iLineNo = tmLst.iLineNo
                                            tlAstInfo(iUpper).iSpotType = tmLst.iSpotType
                                            tlAstInfo(iUpper).sSplitNet = tmLst.sSplitNetwork
                                            tlAstInfo(iUpper).iAirPlay = ilAirPlay
                                            tlAstInfo(iUpper).iAgfCode = tmLst.iAgfCode
                                            tlAstInfo(iUpper).iComp = ilAttComp
                                            tlAstInfo(iUpper).sLstLnStartTime = tmLst.sLnStartTime
                                            tlAstInfo(iUpper).sLstLnEndTime = tmLst.sLnEndTime
                                            tlAstInfo(iUpper).sEmbeddedOrROS = ""
                                            If iIndex >= 0 Then
                                                tlAstInfo(iUpper).sEmbeddedOrROS = tlCPDat(iIndex).sEmbeddedOrROS
                                            End If
                                            If Trim$(tlAstInfo(iUpper).sEmbeddedOrROS) = "" Then
                                                tlAstInfo(iUpper).sEmbeddedOrROS = smDefaultEmbeddedOrROS
                                            End If
                                            'tlAstInfo(iUpper).iPledgeStatus = tlCPDat(iIndex).iFdStatus
                                            If iIndex >= 0 Then
                                                tlAstInfo(iUpper).iPledgeStatus = tlCPDat(iIndex).iFdStatus
                                                tlAstInfo(iUpper).sPdDayFed = tlCPDat(iIndex).sPdDayFed
                                            Else
                                                'tlAstInfo(iUpper).iPledgeStatus = iStatus
                                                '10/10/08:  Don't change pledge from Live (no pledge information defined)
                                                'Pledge exist but not found.  Not sure if the pledge status should be changed
                                                If UBound(tlCPDat) > 0 Then
                                                    tlAstInfo(iUpper).iPledgeStatus = iStatus
                                                Else
                                                    '9/30/11: No Pledge defined, use LST status
                                                    tlAstInfo(iUpper).iPledgeStatus = tmLst.iStatus
                                                End If
                                                tlAstInfo(iUpper).sPdDayFed = ""
                                            End If
                                            tlAstInfo(iUpper).sPdDays = sPdDays
                                            tlAstInfo(iUpper).sLstZone = tmLst.sZone
                                            If (iCPStatus = 0) And (tlAstInfo(iUpper).iCPStatus = 0) Then
                                                'D.S. 12/21/01
                                                'Then no posting has been done
                                                'Now handle the case where pledge info has been altered after the CP
                                                'has been generated
                                                If tlAstInfo(iUpper).iStatus <> iStatus Then
                                                    lAstCode = tlAstInfo(iUpper).lCode
                                                    '6/29/06: change mGetAstInfo to use API call
                                                    'slSQLQuery = "UPDATE ast SET "
                                                    'slSQLQuery = slSQLQuery + "astStatus = " & iStatus
                                                    'slSQLQuery = slSQLQuery + " WHERE (astCode = " & lAstCode & ")"
                                                    'cnn.BeginTrans
                                                    'cnn.Execute slSQLQuery, rdExecDirect
                                                    'cnn.CommitTrans
                                                    tlAstSrchKey.lCode = lAstCode
                                                    ilRet = btrGetEqual(hlAst, tlAst, imAstRecLen, tlAstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
                                                    If ilRet = BTRV_ERR_NONE Then
                                                        tlAst.iStatus = iStatus
                                                        ilRet = btrUpdate(hlAst, tlAst, imAstRecLen)
                                                    End If
                                                    '6/29/06: End of change
                                                    tlAstInfo(iUpper).iStatus = iStatus
                                                    tlAstInfo(iUpper).iCPStatus = iCPStatus
                                                End If
                                            End If
                                        End If 'End if of istatus < 20
                                    End If
                                Next ilLoadIdx
                            'End If
                        End If
                        If ilDatStart = 0 Then
                            Exit Do
                        End If
                        '10/12/14: Restore to handle multi-air plays
                        tmLst = tmSvLst
                    Loop
                End If
                If ilChkLen Then
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    If ilFillLen > 0 Then
                        If ilFindFill Then
                            If ilIssueMoveNext = 1 Then
                                tmLst = tmLastLst
                            Else
                                'rst.MoveNext
                                'ilEof = rst.EOF
                                'If Not ilEof Then
                                '    gCreateUDTforLSTPlusDT rst, tmLst
                                'End If
                                llLstLoop = llLstLoop + 1
                                If llLstLoop < UBound(tmLSTAst) Then
                                    'gCreateUDTforLSTPlusDT rst, tmLst
                                    tmLst = tmLSTAst(llLstLoop).tLST
                                Else
                                    ilEof = True
                                End If
                            End If
                            If Not ilEof Then
                                slSplitNetwork = tmLst.sSplitNetwork
                                If (slSplitNetwork <> "F") Then
                                    ilChkLen = False
                                    ilIssueMoveNext = 1
                                    tmLastLst = tmLst
                                Else
                                    If (tmLst.lRafCode <> llChkLenRafCode) Then
                                        ilIncludeLST = False
                                        ilIssueMoveNext = 0
                                    Else
                                        ilIncludeLST = True
                                        ilIssueMoveNext = 0
                                    End If
                                End If
                            Else
                                ilChkLen = False
                            End If
                        Else
                            ilLen = ilFillLen
                            Do
                                If igExportSource = 2 Then
                                    DoEvents
                                End If
                                tmModelLST = tmLst
                                tmModelLST.iPositionNo = ilChkLenPosition
                                ilRet = gCreateSplitFill(ilLen, llChkLenRafCode, tmModelLST, tmLst)
                                If Not ilRet Then
                                    If ilLen > 60 Then
                                        ilLen = 60
                                    ElseIf ilLen > 30 Then
                                        ilLen = 30
                                    Else
                                        ilLen = ilLen - 5
                                        If ilLen <= 0 Then
                                            ilChkLen = False
                                            Exit Do
                                        End If
                                    End If
                                Else
                                    ilFillLen = ilFillLen - ilLen
                                    ilIncludeLST = True
                                    Exit Do
                                End If
                            Loop While ilRet = False
                        End If
                    Else
                        ilChkLen = False
                        llChkLenRafCode = 0
                    End If
                End If
                If Not ilEof Then
                    If Not ilChkLen Then
                        If igExportSource = 2 Then
                            DoEvents
                        End If
                        ilFillLen = 0
                        llChkLenRafCode = 0
                        ilChkLenPosition = 0
                        If ilIssueMoveNext = 0 Then
                            'rst.MoveNext
                            'ilEof = rst.EOF
                            'If Not ilEof Then
                            '    gCreateUDTforLSTPlusDT rst, tmLst
                            'End If
                            llLstLoop = llLstLoop + 1
                            If llLstLoop < UBound(tmLSTAst) Then
                                tmLst = tmLSTAst(llLstLoop).tLST
                            Else
                                ilEof = True
                            End If
                        ElseIf ilIssueMoveNext = 1 Then
                            'Restore LST
                            tmLst = tmLastLst
                            'ilIssueMoveNext = True
                            ilIssueMoveNext = 0
                            ilEof = False
                        Else
                            ilEof = True
                        End If
                    End If
                End If
            'Wend
            Loop
            'Determine if any records added
'            For iAst = 0 To UBound(tmAstInfo) - 1 Step 1
'                If tmAstInfo(iAst).lCode > 0 Then
'                    iStatus = tmAstInfo(iAst).iStatus
'                    'If iStatus < 20 Then
'                    If (gIsAstStatus(iStatus, ASTEXTENDED_MG) = True) Or (gIsAstStatus(iStatus, ASTEXTENDED_BONUS) = True) Or (gIsAstStatus(iStatus, ASTEXTENDED_REPLACEMENT) = True) Then
'                        tlAstInfo(UBound(tlAstInfo)) = tmAstInfo(iAst)
'                        tmAstInfo(iAst).lCode = 0
'                        ReDim Preserve tlAstInfo(0 To UBound(tlAstInfo) + 1)
'                    End If
'                End If
'            Next iAst
            llCount = 0
            If iAddAst Then
                For iAst = 0 To UBound(tlAstInfo) - 1 Step 1
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    iStatus = tlAstInfo(iAst).iStatus
                    'If iStatus < 20 Then
                    If (gIsAstStatus(iStatus, ASTEXTENDED_MG) = False) And (gIsAstStatus(iStatus, ASTEXTENDED_BONUS) = False) And (gIsAstStatus(iStatus, ASTEXTENDED_REPLACEMENT) = False) Then
                        If tlAstInfo(iAst).lCode <= 0 Then
                            llCount = llCount + 1
                        End If
                    End If
                Next iAst
            End If
            If (tgCPPosting(iLoop).sAstStatus <> "C") Or (llCount > 0) Then
                If iAddAst Then
                    '6/29/06: change mGetAstInfo to use API call
                    ''D.S. 11/11/05 Added 3 lines below to speed up exports
                    'slSQLQuery = "Select MAX(astCode) from ast"
                    'Set rst = gSQLSelectCall(slSQLQuery)
                    'If IsNull(rst(0).Value) = True Then
                    '    llMaxCode = 1
                    'Else
                    '    llMaxCode = rst(0).Value + 1
                    'End If
                    '6/29/06: End of Change
                    
                    For iAst = 0 To UBound(tlAstInfo) - 1 Step 1
                        If igExportSource = 2 Then
                            DoEvents
                        End If
                        lmAttCode = tlAstInfo(iAst).lAttCode
                        imShttCode = tlAstInfo(iAst).iShttCode
                        imVefCode = tlAstInfo(iAst).iVefCode
                        lmSdfCode = tlAstInfo(iAst).lSdfCode
                        lmLstCode = tlAstInfo(iAst).lLstCode
                        sAirDate = Format$(tlAstInfo(iAst).sAirDate, sgShowDateForm)
                        sAirTime = Format$(tlAstInfo(iAst).sAirTime, sgShowTimeWSecForm)
                        sFdDate = Format$(tlAstInfo(iAst).sFeedDate, sgShowDateForm)
                        sFdTime = Format$(tlAstInfo(iAst).sFeedTime, sgShowTimeWSecForm)
                        sPdDate = Format$(tlAstInfo(iAst).sPledgeDate, sgShowDateForm)
                        sPdSTime = Format$(tlAstInfo(iAst).sPledgeStartTime, sgShowTimeWSecForm)
                        sPdETime = Format$(Trim$(tlAstInfo(iAst).sPledgeEndTime), sgShowTimeWSecForm)
                        If Len(Trim$(sPdSTime)) = 0 Then
                            sPdSTime = sFdTime
                        End If
                        If Len(Trim$(sPdETime)) = 0 Then
                            sPdETime = sPdSTime
                        End If
                        iStatus = tlAstInfo(iAst).iStatus
                        'If iStatus < 20 Then
                        If (gIsAstStatus(iStatus, ASTEXTENDED_MG) = False) And (gIsAstStatus(iStatus, ASTEXTENDED_BONUS) = False) And (gIsAstStatus(iStatus, ASTEXTENDED_REPLACEMENT) = False) Then
                            '6/29/06: change mGetAstInfo to use API call
                            'If tlAstInfo(iAst).lCode <= 0 Then
                            '    On Error GoTo InsertErrHand
                            '    ilInsertError = False
                            '    Do
                            '        ilInsertError = False
                            '        slSQLQuery = "INSERT INTO ast"
                            '        slSQLQuery = slSQLQuery + "(astCode, astAtfCode, astShfCode, astVefCode, "
                            '        slSQLQuery = slSQLQuery + "astSdfCode, astLsfCode, astAirDate, astAirTime, "
                            '        slSQLQuery = slSQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, "
                            '        slSQLQuery = slSQLQuery + "astPledgeStartTime, astPledgeEndTime, astPledgeStatus)"
                            '        slSQLQuery = slSQLQuery + " VALUES "
                            '        slSQLQuery = slSQLQuery + "(" & llMaxCode & ", " & lmAttCode & ", " & imShttCode & ", "
                            '        slSQLQuery = slSQLQuery & imVefCode & ", " & lmSdfCode & ", " & lmLstCode & ", "
                            '        slSQLQuery = slSQLQuery + "'" & Format$(sAirDate, sgSQLDateForm) & "', '" & Format$(sAirTime, sgSQLTimeForm) & "', "
                            '        slSQLQuery = slSQLQuery & iStatus & ", " & tlAstInfo(iAst).iCPStatus & ", '" & Format$(sFdDate, sgSQLDateForm) & "', "
                            '        slSQLQuery = slSQLQuery & "'" & Format$(sFdTime, sgSQLTimeForm) & "', '" & Format$(sPdDate, sgSQLDateForm) & "', "
                            '        slSQLQuery = slSQLQuery & "'" & Format$(sPdSTime, sgSQLTimeForm) & "', '" & Format$(sPdETime, sgSQLTimeForm) & "', " & tlAstInfo(iAst).iPledgeStatus & ")"
                            '        'cnn.BeginTrans
                            '        cnn.Execute slSQLQuery, rdExecDirect
                            '        'cnn.CommitTrans
                            '    Loop While ilInsertError = True
                            '    tlAstInfo(iAst).lCode = llMaxCode
                            '    llMaxCode = llMaxCode + 1
                            '    On Error GoTo ErrHand
                            'Else
                            '    'Remove previously added record since FeedDate is part of key and is not modifiable
                            '    cnn.BeginTrans
                            '    slSQLQuery = "DELETE FROM Ast WHERE (astCode = " & tlAstInfo(iAst).lCode & ")"
                            '    cnn.Execute slSQLQuery, rdExecDirect
                            '    cnn.CommitTrans
                            '    slSQLQuery = "INSERT INTO ast"
                            '    slSQLQuery = slSQLQuery + "(astCode, astAtfCode, astShfCode, astVefCode, "
                            '    slSQLQuery = slSQLQuery + "astSdfCode, astLsfCode, astAirDate, astAirTime, "
                            '    slSQLQuery = slSQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, "
                            '    slSQLQuery = slSQLQuery + "astPledgeStartTime, astPledgeEndTime, astPledgeStatus)"
                            '    slSQLQuery = slSQLQuery + " VALUES "
                            '    slSQLQuery = slSQLQuery + "(" & tlAstInfo(iAst).lCode & ", " & lmAttCode & ", " & imShttCode & ", "
                            '    slSQLQuery = slSQLQuery & imVefCode & ", " & lmSdfCode & ", " & lmLstCode & ", "
                            '    slSQLQuery = slSQLQuery + "'" & Format$(sAirDate, sgSQLDateForm) & "', '" & Format$(sAirTime, sgSQLTimeForm) & "', "
                            '    slSQLQuery = slSQLQuery & iStatus & ", " & tlAstInfo(iAst).iCPStatus & ", '" & Format$(sFdDate, sgSQLDateForm) & "', "
                            '    slSQLQuery = slSQLQuery & "'" & Format$(sFdTime, sgSQLTimeForm) & "', '" & Format$(sPdDate, sgSQLDateForm) & "', "
                            '    slSQLQuery = slSQLQuery & "'" & Format$(sPdSTime, sgSQLTimeForm) & "', '" & Format$(sPdETime, sgSQLTimeForm) & "', " & tlAstInfo(iAst).iPledgeStatus & ")"
                            '    'cnn.BeginTrans
                            '    cnn.Execute slSQLQuery, rdExecDirect
                            '    'cnn.CommitTrans
                            'End If
                            If tlAstInfo(iAst).lCode <= 0 Then
                                tmAst.lCode = 0
                                ilAstCPStatus = 0
                            Else
                                tlAstSrchKey.lCode = tlAstInfo(iAst).lCode
                                ilRet = btrGetEqual(hlAst, tlAst, imAstRecLen, tlAstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
                                ilRet = btrDelete(hlAst)
                                tmAst.lCode = tlAstInfo(iAst).lCode
                                ilAstCPStatus = tlAst.iCPStatus
                            End If
                            tmAst.lAtfCode = lmAttCode
                            tmAst.iShfCode = imShttCode
                            tmAst.iVefCode = imVefCode
                            tmAst.lSdfCode = lmSdfCode
                            tmAst.lLsfCode = lmLstCode
                            gPackDate sAirDate, tmAst.iAirDate(0), tmAst.iAirDate(1)
                            gPackTime sAirTime, tmAst.iAirTime(0), tmAst.iAirTime(1)
                            tmAst.iStatus = iStatus
                            tmAst.iCPStatus = tlAstInfo(iAst).iCPStatus
                            gPackDate sFdDate, tmAst.iFeedDate(0), tmAst.iFeedDate(1)
                            gPackTime sFdTime, tmAst.iFeedTime(0), tmAst.iFeedTime(1)
                            '12/9/13
                            'gPackDate sPdDate, tmAst.iPledgeDate(0), tmAst.iPledgeDate(1)
                            'gPackTime sPdSTime, tmAst.iPledgeStartTime(0), tmAst.iPledgeStartTime(1)
                            'gPackTime sPdETime, tmAst.iPledgeEndTime(0), tmAst.iPledgeEndTime(1)
                            'tmAst.iPledgeStatus = tlAstInfo(iAst).iPledgeStatus
                            tmAst.iAdfCode = tlAstInfo(iAst).iAdfCode
                            tmAst.lDatCode = tlAstInfo(iAst).lDatCode
                            tmAst.lCpfCode = tlAstInfo(iAst).lCpfCode
                            tmAst.lRsfCode = 0
                            tmAst.sStationCompliant = ""
                            tmAst.sAgencyCompliant = ""
                            tmAst.sAffidavitSource = ""
                            If ilAstCPStatus = 1 Then
                                ilSchdCount = 0
                                ilAiredCount = 0
                                ilPledgeCompliantCount = 0
                                ilAgyCompliantCount = 0
                                gIncSpotCounts tlAstInfo(iAst), ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount, slServiceAgreement
                                tmAst.sStationCompliant = tlAstInfo(iAst).sStationCompliant
                                tmAst.sAgencyCompliant = tlAstInfo(iAst).sAgencyCompliant
                                tmAst.sAffidavitSource = tlAstInfo(iAst).sAffidavitSource
                            End If
                            tmAst.lCntrNo = tlAstInfo(iAst).lCntrNo
                            tmAst.iLen = tlAstInfo(iAst).iLen
                            tmAst.lLkAstCode = tlAstInfo(iAst).lLkAstCode
                            tmAst.iMissedMnfCode = tlAstInfo(iAst).iMissedMnfCode
                            tmAst.iUstCode = igUstCode
                            ilRet = btrInsert(hlAst, tmAst, imAstRecLen, INDEXKEY0)
                            
                            If ilRet <> 0 Then
                                ilRet = ilRet
                            End If
                            
                            If ilRet >= 30000 Then
                                ilRet = csiHandleValue(0, 7)
                            End If
                            
                            If blErrorMsgLogged = False Then
                                If (ilRet <> 0) Then
                                    gLogMsg "mGetAstInfo: Insert AST Error # " & ilRet, "AffErrorLog.Txt", False
                                    blErrorMsgLogged = True
                                ElseIf tmAst.lCode <= 0 Then
                                    gLogMsg "mGetAstInfo: Insert AST Returned Ast Code " & tmAst.lCode, "AffErrorLog.Txt", False
                                    blErrorMsgLogged = True
                                End If
                            End If
                            
                            tlAstInfo(iAst).lCode = tmAst.lCode
                            '6/28/06: End of Change
                        End If
                    Next iAst
                    '3/5/15
                    'If tgCPPosting(iLoop).lCpttCode > 0 Then
                    If (tgCPPosting(iLoop).lCpttCode > 0) And (iAdfCode <= 0) Then
                        'Update CPTT
                        slSQLQuery = "UPDATE cptt SET "
                        slSQLQuery = slSQLQuery + "cpttAstStatus = " & "'C'"
                        slServiceAgreement = "N"
                        For ilAtt = 0 To UBound(tmATTCrossDates) - 1 Step 1
                            If tmATTCrossDates(ilAtt).lAttCode = tgCPPosting(iLoop).lAttCode Then
                                slServiceAgreement = tmATTCrossDates(ilAtt).sServiceAgreement
                                Exit For
                            End If
                        Next ilAtt
                        If slServiceAgreement = "Y" Then
                            slSQLQuery = slSQLQuery + ", cpttStatus = 1"
                            slSQLQuery = slSQLQuery + ", cpttPostingStatus = 2"
                        End If
                        slSQLQuery = slSQLQuery + " WHERE (cpttCode = " & tgCPPosting(iLoop).lCpttCode & ")"
                        cnn.BeginTrans
                        'cnn.Execute slSQLQuery, rdExecDirect
                        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                            '6/13/16: Replaced GoSub
                            'GoSub ErrHand:
                            Screen.MousePointer = vbDefault
                            gHandleError "AffErrorLog.txt", "modCPReturns-mGetAstInfo"
                            cnn.RollbackTrans
                            mGetAstInfo = False
                            If llLockRec > 0 Then
                                llLockRec = gDeleteLockRec_ByRlfCode(llLockRec)
                            End If
                            mFilterByAdvt tlAstInfo, ilInAdfCode
                            Exit Function
                        End If
                        cnn.CommitTrans
                        gFileChgdUpdate "cptt.mkd", True
                    End If
                End If
            '3/19/13: Moved up here from below
            End If
            
                For iAst = 0 To UBound(tmAstInfo) - 1 Step 1
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    'Remove any records not referenced (found)
                    If tmAstInfo(iAst).lCode > 0 Then
                        If (iAdfCode <= 0) Or (iAdfCode = tmAstInfo(iAst).iAdfCode) Then
                            cnn.BeginTrans
                            slSQLQuery = "DELETE FROM Ast WHERE (astCode = " & tmAstInfo(iAst).lCode & ")"
                            'cnn.Execute slSQLQuery, rdExecDirect
                            If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                                '6/13/16: Replaced GoSub
                                'GoSub ErrHand:
                                Screen.MousePointer = vbDefault
                                gHandleError "AffErrorLog.txt", "modCPReturns-mGetAstInfo"
                                cnn.RollbackTrans
                                mFilterByAdvt tlAstInfo, ilInAdfCode
                                mGetAstInfo = False
                                If llLockRec > 0 Then
                                    llLockRec = gDeleteLockRec_ByRlfCode(llLockRec)
                                End If
                                Exit Function
                            End If
                            cnn.CommitTrans
                        End If
                    End If
                Next iAst
            '3/19/13: Moved up
            'End If
            
            ilRet = gFilterAstInfoBySalesSource(tlAstInfo())
            
            If (igTimes <> 0) And (blRebuildAst) And (blFilterByAirDates) Then
                ilRet = gFilterAstInfoByAirDate(tlAstInfo(), lFWkDate, lLWkDate)
            End If
            
            
            'Assign split copy
            '2/5/09:  Move to gGetRegionCopy.  If false, then bypass blaclout copy
            'If iAddAst Then
            If (blGetRegionCopy) And (UBound(tmRegionAssignmentInfo) > LBound(tmRegionAssignmentInfo) Or bmImportSpot) Then
                'lgSTime8 = timeGetTime
                For iAst = 0 To UBound(tlAstInfo) - 1 Step 1
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    lgCount2 = lgCount2 + 1
                    If tgStatusTypes(gGetAirStatus(tlAstInfo(iAst).iPledgeStatus)).iPledged <> 2 Then
                        If (slForbidSplitLive <> "Y") Or ((slForbidSplitLive = "Y") And (tgStatusTypes(gGetAirStatus(tlAstInfo(iAst).iPledgeStatus)).iPledged <> 0)) Then
                            tlAstInfo(iAst).iRegionType = gGetRegionCopy(hlAst, iAddAst, tlAstInfo(iAst), tlAstInfo(iAst).sRCart, tlAstInfo(iAst).sRProduct, tlAstInfo(iAst).sRISCI, tlAstInfo(iAst).sRCreativeTitle, tlAstInfo(iAst).lRCrfCsfCode, tlAstInfo(iAst).lRCrfCode, tlAstInfo(iAst).lRCifCode, tlAstInfo(iAst).lRRsfCode, tlAstInfo(iAst).lRCpfCode, tlAstInfo(iAst).sReplacementCue)
                        End If
                    End If
                Next iAst
                gClosePoolFiles
                'lgETime8 = timeGetTime
                'lgTtlTime8 = lgTtlTime8 + (lgETime8 - lgSTime8)
            End If
            ''4/18/13: Get range of all breaks.  Use dat and ast
            'For iDay = 0 To 6 Step 1
            '    tgAstTimeRange(iDay).lStartTime = -1
            '    tgAstTimeRange(iDay).lEndTime = -1
            'Next iDay
            'For iDat = 0 To UBound(tlCPDat) - 1 Step 1
            '    For iDay = 0 To 6 Step 1
            '        If tlCPDat(iDat).iFdDay(iDay) <> 0 Then
            '            If tgAstTimeRange(iDay).lStartTime = -1 Then
            '                tgAstTimeRange(iDay).lStartTime = gTimeToLong(tlCPDat(iDat).sFdSTime, False)
            '                tgAstTimeRange(iDay).lEndTime = tgAstTimeRange(iDay).lStartTime + 1
            '            Else
            '                If gTimeToLong(tlCPDat(iDat).sFdSTime, False) < tgAstTimeRange(iDay).lStartTime Then
            '                    tgAstTimeRange(iDay).lStartTime = gTimeToLong(tlCPDat(iDat).sFdSTime, False)
            '                End If
            '                If gTimeToLong(tlCPDat(iDat).sFdETime, True) > tgAstTimeRange(iDay).lEndTime Then
            '                    tgAstTimeRange(iDay).lEndTime = gTimeToLong(tlCPDat(iDat).sFdETime, True)
            '                End If
            '            End If
            '        End If
            '    Next iDay
            'Next iDat
            'For iAst = 0 To UBound(tlAstInfo) - 1 Step 1
            '    If igExportSource = 2 Then
            '        DoEvents
            '    End If
            '    iDay = Weekday(tlAstInfo(iAst).sFeedDate, vbMonday) - 1
            '    If tgAstTimeRange(iDay).lStartTime = -1 Then
            '        tgAstTimeRange(iDay).lStartTime = gTimeToLong(tlAstInfo(iAst).sFeedTime, False)
            '        tgAstTimeRange(iDay).lEndTime = tgAstTimeRange(iDay).lStartTime + 1
            '    Else
            '        If gTimeToLong(tlAstInfo(iAst).sFeedTime, False) < tgAstTimeRange(iDay).lStartTime Then
            '            tgAstTimeRange(iDay).lStartTime = gTimeToLong(tlAstInfo(iAst).sFeedTime, False)
            '        End If
            '        If gTimeToLong(tlAstInfo(iAst).sFeedTime, True) > tgAstTimeRange(iDay).lEndTime Then
             '           tgAstTimeRange(iDay).lEndTime = gTimeToLong(tlAstInfo(iAst).sFeedTime, True)
            '        End If
            '    End If
            'Next iAst
            ''End If
        Else
            ReDim tlAstInfo(0 To 0) As ASTINFO
            '3/5/15
            ReDim tmAstInfo(0 To 0) As ASTINFO
            ReDim tgDelAst(0 To 0) As ASTINFO
        End If
        If igTimes <= 2 Then
            gClearAbf tgCPPosting(iLoop).iVefCode, tgCPPosting(iLoop).iShttCode, sFWkDate, sLWkDate, False
        End If
    Next iLoop
    '2/13/15: Pick up MG and Replacement bypassed
    blExtendExist = False
    For iAst = 0 To UBound(tmAstInfo) - 1 Step 1
        'If (tmAstInfo(iAst).lCode < 0) And (tmAstInfo(iAst).iStatus >= ASTEXTENDED_MG) Then
        If (tmAstInfo(iAst).lCode < 0) And (tmAstInfo(iAst).iStatus >= ASTEXTENDED_MG) And (tmAstInfo(iAst).iStatus <= ASTEXTENDED_REPLACEMENT) Then
            If (gIsAstStatus(tmAstInfo(iAst).iStatus, ASTEXTENDED_MG) = True) Or (gIsAstStatus(tmAstInfo(iAst).iStatus, ASTEXTENDED_BONUS) = True) Or (gIsAstStatus(tmAstInfo(iAst).iStatus, ASTEXTENDED_REPLACEMENT) = True) Then
                '6/1/18
                'If ((iAdfCode = -1) Or (tmAstInfo(iAst).iAdfCode = iAdfCode)) Then
                If ((iAdfCode <= 0) Or (tmAstInfo(iAst).iAdfCode = iAdfCode)) Then
                    blExtendExist = True
                    slSQLQuery = "Select * From lst where lstCode = " & tmAstInfo(iAst).lLstCode
                    Set rst_Genl = gSQLSelectCall(slSQLQuery)
                    gCreateUDTforLST rst_Genl, tmLst
                    'fill in missing fields
                    tmAstInfo(iAst).sProd = tmLst.sProd
                    tmAstInfo(iAst).iAnfCode = tmLst.iAnfCode
                    tmAstInfo(iAst).sCart = tmLst.sCart
                    tmAstInfo(iAst).sISCI = tmLst.sISCI
                    tmAstInfo(iAst).lCifCode = tmLst.lCifCode
                    tmAstInfo(iAst).lCrfCsfCode = tmLst.lCrfCsfCode
                    tmAstInfo(iAst).lCpfCode = tmLst.lCpfCode
                    tmAstInfo(iAst).iVefCode = tmLst.iLogVefCode
                    tmAstInfo(iAst).sPdDays = ""
                    'D.L. 3/17/21 retain ast length
                    'tmAstInfo(iAst).iLen = tmLst.iLen
                    tmAstInfo(iAst).lgsfCode = tmLst.lgsfCode
                    tmAstInfo(iAst).iLstLnVefCode = tmLst.iLnVefCode
                    tmAstInfo(iAst).lLstBkoutLstCode = tmLst.lBkoutLstCode
                    '10/9/14: Retain Generic ISCI
                    tmAstInfo(iAst).sGISCI = ""
                    tmAstInfo(iAst).sGCart = ""
                    tmAstInfo(iAst).sGProd = ""
                    '11/2/16: Add Event ID information. Used with XDS ProgramCode:Cue
                    tmAstInfo(iAst).lEvtIDCefCode = tmLst.lEvtIDCefCode
                    
                    tmAstInfo(iAst).sLstStartDate = tmLst.sStartDate
                    tmAstInfo(iAst).sLstEndDate = tmLst.sEndDate
                    tmAstInfo(iAst).iLstSpotsWk = tmLst.iSpotsWk
                    tmAstInfo(iAst).iLstMon = tmLst.iMon
                    tmAstInfo(iAst).iLstTue = tmLst.iTue
                    tmAstInfo(iAst).iLstWed = tmLst.iWed
                    tmAstInfo(iAst).iLstThu = tmLst.iThu
                    tmAstInfo(iAst).iLstFri = tmLst.iFri
                    tmAstInfo(iAst).iLstSat = tmLst.iSat
                    tmAstInfo(iAst).iLstSun = tmLst.iSun
                    tmAstInfo(iAst).iLineNo = tmLst.iLineNo
                    tmAstInfo(iAst).iSpotType = tmLst.iSpotType
                    tmAstInfo(iAst).sSplitNet = tmLst.sSplitNetwork
                    tmAstInfo(iAst).iAirPlay = 1
                    tmAstInfo(iAst).iAgfCode = tmLst.iAgfCode
                    tmAstInfo(iAst).sLstLnStartTime = tmLst.sLnStartTime
                    tmAstInfo(iAst).sLstLnEndTime = tmLst.sLnEndTime
                    tmAstInfo(iAst).sEmbeddedOrROS = smDefaultEmbeddedOrROS
                    
                    tlAstInfo(UBound(tlAstInfo)) = tmAstInfo(iAst)
                    ReDim Preserve tlAstInfo(0 To UBound(tlAstInfo) + 1) As ASTINFO
                End If
            End If
        End If
    Next iAst
    If blExtendExist Then
        For iAst = 0 To UBound(tlAstInfo) - 1 Step 1
            'Create sort key
            If igTimes = 0 Then
                slSortDate = gDateValue(tlAstInfo(iAst).sAirDate)
                Do While Len(slSortDate) < 6
                    slSortDate = "0" & slSortDate
                Loop
                slSortTime = gTimeToLong(tlAstInfo(iAst).sAirTime, False)
                Do While Len(slSortTime) < 6
                    slSortTime = "0" & slSortTime
                Loop
            ElseIf (igTimes <> 2) And (igTimes <> 4) Then
                slSortDate = gDateValue(tlAstInfo(iAst).sAirDate)
                Do While Len(slSortDate) < 6
                    slSortDate = "0" & slSortDate
                Loop
                slSortTime = gTimeToLong(tlAstInfo(iAst).sAirTime, False)
                Do While Len(slSortTime) < 6
                    slSortTime = "0" & slSortTime
                Loop
            Else
                slSortDate = gDateValue(tlAstInfo(iAst).sFeedDate)
                Do While Len(slSortDate) < 6
                    slSortDate = "0" & slSortDate
                Loop
                slSortTime = gTimeToLong(tlAstInfo(iAst).sFeedTime, False)
                Do While Len(slSortTime) < 6
                    slSortTime = "0" & slSortTime
                Loop
            End If
            slSortPosition = iAst
            Do While Len(slSortPosition) < 5
                slSortPosition = "0" & slSortPosition
            Loop
            tlAstInfo(iAst).sKey = slSortDate & slSortTime & slSortPosition
        Next iAst
        If UBound(tlAstInfo) - 1 >= 1 Then
            ArraySortTyp fnAV(tlAstInfo(), 0), UBound(tlAstInfo), 0, LenB(tlAstInfo(0)), 0, LenB(tlAstInfo(0).sKey), 0
        End If
    End If
    For iAst = 0 To UBound(tlAstInfo) - 1 Step 1
        If tlAstInfo(iAst).lCode < 0 Then
            tlAstInfo(iAst).lCode = -tlAstInfo(iAst).lCode
        End If
        '8/31/16: Check if comment suppressed
        If slHideCommOnWeb = "Y" Then
            tlAstInfo(iAst).lCrfCsfCode = 0
            tlAstInfo(iAst).lRCrfCsfCode = 0
        End If
    Next iAst
    'Erase tmBkoutLst
    On Error Resume Next
    mGetAstInfo = True
    Erase tmAstInfo
    Erase llGsfCode
    lgETime8 = timeGetTime
    lgTtlTime8 = lgTtlTime8 + (lgETime8 - lgSTime8)
    
    If llLockRec > 0 Then
        llLockRec = gDeleteLockRec_ByRlfCode(llLockRec)
    End If
    If iAdfCode <> ilInAdfCode Then
        mFilterByAdvt tlAstInfo, ilInAdfCode
    End If
    
    Exit Function
    
InsertErrHand:
    llMaxCode = llMaxCode + 1
    ilInsertError = True
    Resume Next

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-mGetAstInfo"
    mGetAstInfo = False
    If llLockRec > 0 Then
        llLockRec = gDeleteLockRec_ByRlfCode(llLockRec)
    End If
    Exit Function
CheckForCopyErr:
    ilRet = 1
    Resume Next
End Function

Public Function gAddBonusSpot(llCntrNo As Long, ilAdfCode As Integer, ilVefCode As Integer, slAirDate As String, slAirTime As String, slZone As String, llAttCode As Long, ilShttCode As Integer, slProd As String, slCart As String, slISCI As String, ilLen As Integer) As Integer
    Dim llLstCode As Long
    Dim llDATCode As Long
    Dim llCpfCode As Long
    Dim llRsfCode As Long
    Dim slStationCompliant As String
    Dim slAgencyCompliant As String
    Dim slAffidavitSource As String
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    slSQLQuery = "INSERT INTO lst (lstType, lstSdfCode, lstCntrNo, "
    slSQLQuery = slSQLQuery & "lstAdfCode, lstAgfCode, lstProd, "
    slSQLQuery = slSQLQuery & "lstLineNo, lstLnVefCode, lstStartDate,"
    slSQLQuery = slSQLQuery & "lstEndDate, lstMon, lstTue, "
    slSQLQuery = slSQLQuery & "lstWed, lstThu, lstFri, "
    slSQLQuery = slSQLQuery & "lstSat, lstSun, lstSpotsWk, "
    slSQLQuery = slSQLQuery & "lstPriceType, lstPrice, lstSpotType, "
    slSQLQuery = slSQLQuery & "lstLogVefCode, lstLogDate, lstLogTime, "
    slSQLQuery = slSQLQuery & "lstDemo, lstAud, lstISCI, "
    slSQLQuery = slSQLQuery & "lstWkNo, lstBreakNo, lstPositionNo, "
    slSQLQuery = slSQLQuery & "lstSeqNo, lstZone, lstCart, "
    slSQLQuery = slSQLQuery & "lstCpfCode, lstCrfCsfCode, lstStatus, "
    slSQLQuery = slSQLQuery & "lstLen, lstUnits, lstCifCode, "
    slSQLQuery = slSQLQuery & "lstAnfCode, lstEvtIDCefCode, lstSplitNetwork, "
    '12/28/06
    ''slSQLQuery = slSQLQuery & "lstRafCode, lstUnused)"
    'slSQLQuery = slSQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, lstUnused)"
    slSQLQuery = slSQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, "
    slSQLQuery = slSQLQuery & "lstLnStartTime, lstLnEndTime, lstUnused)"
    slSQLQuery = slSQLQuery & " VALUES (" & 2 & ", " & 0 & ", " & llCntrNo & ", "
    slSQLQuery = slSQLQuery & ilAdfCode & ", " & 0 & ", '" & gFixQuote(slProd) & "', "
    slSQLQuery = slSQLQuery & 0 & ", " & 0 & ", '" & Format$("1/1/1970", sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & "'" & Format$("1/1/1970", sgSQLDateForm) & "', " & 0 & ", " & 0 & ", "
    slSQLQuery = slSQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
    slSQLQuery = slSQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
    slSQLQuery = slSQLQuery & 1 & ", " & 0 & ", " & 5 & ", "
    slSQLQuery = slSQLQuery & ilVefCode & ", '" & Format$(slAirDate, sgSQLDateForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "', "
    slSQLQuery = slSQLQuery & "'" & "" & "', " & 0 & ", '" & gFixQuote(slISCI) & "', "
    slSQLQuery = slSQLQuery & 0 & ", " & 0 & ", " & 0 & ", "
    slSQLQuery = slSQLQuery & 0 & ", '" & slZone & "', '" & gFixQuote(slCart) & "', "
    slSQLQuery = slSQLQuery & 0 & ", " & 0 & ", " & ASTEXTENDED_BONUS & ", "
    slSQLQuery = slSQLQuery & ilLen & ", " & 0 & ", " & 0 & ", "
    slSQLQuery = slSQLQuery & 0 & ", " & 0 & ", '" & "N" & "', "
    '12/28/06
    ''slSQLQuery = slSQLQuery & 0 & ", '" & "" & "'" & ")"
    'slSQLQuery = slSQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "N" & "', " & 0 & ", '" & "" & "'" & ")"
    slSQLQuery = slSQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "N" & "', " & 0 & ", "
    slSQLQuery = slSQLQuery & "'" & Format("12am", sgSQLTimeForm) & "', '" & Format("12am", sgSQLTimeForm) & "', '" & "" & "'" & ")"
    cnn.BeginTrans
    'cnn.Execute slSQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "modCPReturns-gAddBonusSpot"
        cnn.RollbackTrans
        gAddBonusSpot = False
        Exit Function
    End If
    cnn.CommitTrans
    slSQLQuery = "Select MAX(lstCode) from lst"
    Set rst = gSQLSelectCall(slSQLQuery)
    llLstCode = rst(0).Value
    llDATCode = 0
    llCpfCode = 0
    llRsfCode = 0
    slStationCompliant = ""
    slAgencyCompliant = ""
    slAffidavitSource = ""
    slSQLQuery = "INSERT INTO ast"
    slSQLQuery = slSQLQuery + "(astAtfCode, astShfCode, astVefCode, "
    slSQLQuery = slSQLQuery + "astSdfCode, astLsfCode, astAirDate, astAirTime, "
    '12/13/13: Support New AST layout
    'slSQLQuery = slSQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, astPledgeStartTime, astPledgeEndTime)"
    slSQLQuery = slSQLQuery + "astStatus, astCPStatus, astFeedDate, astFeedTime, "
    slSQLQuery = slSQLQuery + "astAdfCode, astDatCode, astCpfCode, astRsfCode, astStationCompliant, astAgencyCompliant, astAffidavitSource, astCntrNo, astLen, astLkAstCode, astMissedMnfCode, astUstCode)"
    slSQLQuery = slSQLQuery + " VALUES "
    slSQLQuery = slSQLQuery + "(" & llAttCode & ", " & ilShttCode & ", "
    slSQLQuery = slSQLQuery & ilVefCode & ", " & 0 & ", " & llLstCode & ", "
    slSQLQuery = slSQLQuery + "'" & Format$(slAirDate, sgSQLDateForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "', "
    slSQLQuery = slSQLQuery & ASTEXTENDED_BONUS & ", " & "1" & ", '" & Format$(slAirDate, sgSQLDateForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "', "
    'slSQLQuery = slSQLQuery & "'" & Format$(slAirTime, sgSQLTimeForm) & "', '" & Format$(slAirDate, sgSQLDateForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "', '" & Format$(slAirTime, sgSQLTimeForm) & "')"
    slSQLQuery = slSQLQuery & ilAdfCode & ", " & llDATCode & ", " & llCpfCode & ", " & llRsfCode & ", "
    slSQLQuery = slSQLQuery & "'" & slStationCompliant & "', '" & slAgencyCompliant & "', '" & slAffidavitSource & "', " & llCntrNo & ", " & ilLen & ", " & 0 & ", " & 0 & ", " & igUstCode & ")"
    cnn.BeginTrans
    'cnn.Execute slSQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "modCPReturns-gAddBonusSpot"
        cnn.RollbackTrans
        gAddBonusSpot = False
        Exit Function
    End If
    cnn.CommitTrans
    On Error GoTo 0
    gAddBonusSpot = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-gAddBonusSpot"
    gAddBonusSpot = False
End Function

Public Function gGetMaxAstCode() As Integer
    Dim slSQLQuery As String

    'D.S. 11/21/05
    'Set global varible lgMaxAstCode to the highest astcode + 1
    
    Dim ast_rst As ADODB.Recordset
                    
    On Error GoTo ErrHand
                    
    gGetMaxAstCode = False
    
    slSQLQuery = "Select MAX(astCode) from ast"
    Set ast_rst = gSQLSelectCall(slSQLQuery)
    If IsNull(ast_rst(0).Value) = True Then
        lgMaxAstCode = 1
    Else
        lgMaxAstCode = ast_rst(0).Value + 1
    End If
    
    gGetMaxAstCode = True

    Exit Function

ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-gGetMaxAstCode"
End Function


'Public Function gGetRegionCopy(ilShttCode As Integer, llSdfCode As Long, ilVefCode As Integer, slCartNo As String, slProduct As String, slISCI As String, slCreativeTitle As String, llCrfCsfCode As Long, llCrfCode As Long) As Integer
Public Function gGetRegionCopy(hlAst As Integer, ilAssginBlackoutCopy As Integer, tlAstInfo As ASTINFO, slCartNo As String, slProduct As String, slISCI As String, slCreativeTitle As String, llCrfCsfCode As Long, llCrfCode As Long, llCifCode As Long, llRsfCode As Long, llCpfCode As Long, slReplacementCue As String) As Integer
    'ilAssginBlackoutCopy(I)- If not creating ast, then blackout can't be created
    'Dim sht_rst As ADODB.Recordset
    Dim llRegionDefinitionIndex As Long
    Dim llUpperSplit As Long
    Dim ilShttCode As Integer
    Dim llSdfCode As Long
    Dim ilVefCode As Integer
    Dim ilMktCode As Integer
    Dim ilMSAMktCode As Integer
    Dim llLstCode As Long
    Dim slState As String
    Dim ilFmtCode As Integer
    Dim ilTztCode As Integer
    Dim ilRet As Integer
    ' dan M 4/19/13 pass as Parameter
    'Dim llCpfCode As Long
    Dim ilCifAdfCode As Integer
    Dim llRegionIndex As Long
    Dim slGroupInfo As String
    Dim slLogDate As String
    Dim slLogTime As String
    Dim ilLoop As Integer
    Dim llCode As Long
    Dim llLogDate As Long
    Dim llLogTime As Long
    Dim llBkoutLstIndex As Long
    Dim llLst As Long
    Dim llBkoutLstCode As Long
    Dim llRet As Long
    Dim ilWeekDay(0 To 6) As Integer
    ReDim tlRegionDefinition(0 To 0) As REGIONDEFINITION
    ReDim tlSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
    ReDim tmRegionDefinition(0 To 0) As REGIONDEFINITION
    ReDim tmSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
    ReDim tlPartRegionDef(0 To 1) As REGIONDEFINITION
    
    Dim llSdf As Long
    Dim llLoop As Long
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    imAstRecLen = Len(tmAst)
    llCifCode = 0
    lgCount1 = lgCount1 + 1
    lgSTime9 = timeGetTime
    lgSTime10 = timeGetTime
    lgSTime11 = timeGetTime
    lgSTime12 = timeGetTime
    gGetRegionCopy = 0  'False
    slCartNo = ""
    slProduct = ""
    slISCI = ""
    slCreativeTitle = ""
    ilShttCode = tlAstInfo.iShttCode
    llSdfCode = tlAstInfo.lSdfCode
    ilVefCode = tlAstInfo.iVefCode
    llLstCode = tlAstInfo.lLstCode
    'slSQLQuery = "SELECT shttCode, shttState, shttMktCode, shttFmtCode, shttTztCode FROM shtt WHERE (shttCode = " & ilShttCode & ")"
    'Set sht_rst = gSQLSelectCall(slSQLQuery)
    'If sht_rst.EOF Then
    '    Exit Function
    'End If
    'ilMktCode = sht_rst!shttMktCode
    'slState = Trim$(sht_rst!shttState)
    'ilFmtCode = sht_rst!shttFmtCode
    'ilTztCode = sht_rst!shttTztCode
    ilRet = gBinarySearchStationInfoByCode(ilShttCode)
    If ilRet = -1 Then
        Exit Function
    End If
    ilMktCode = tgStationInfoByCode(ilRet).iMktCode
    ilMSAMktCode = tgStationInfoByCode(ilRet).iMSAMktCode
    '12/28/15
    'slState = tgStationInfoByCode(ilRet).sPostalName
    If sgSplitState = "L" Then
        slState = tgStationInfoByCode(ilRet).sStateLic
    ElseIf sgSplitState = "P" Then
        slState = tgStationInfoByCode(ilRet).sPhyState
    Else
        slState = tgStationInfoByCode(ilRet).sMailState
    End If
    ilFmtCode = tgStationInfoByCode(ilRet).iFormatCode
    ilTztCode = tgStationInfoByCode(ilRet).iTztCode
'    ilRet = gBuildRegionDefinitions("C", llSdfCode, ilVefCode, tlRegionDefinition(), tlSplitCategoryInfo())
'
'    lgCount6 = lgCount6 + UBound(tlRegionDefinition)  'Total number of reg. def.
'    If UBound(tlRegionDefinition) > 0 Then   'Number of spots w/reg. def.
'        lgCount7 = lgCount7 + 1
'    End If
'    For llRegionDefinitionIndex = 0 To UBound(tlRegionDefinition) - 1 Step 1
'        lgCount3 = lgCount3 + 1  'Num of reg. def. checked to find an applicable def.
'        tlPartRegionDef(0) = tlRegionDefinition(llRegionDefinitionIndex)
'        ReDim tlSplitCategoryInfo(0 To 500) As SPLITCATEGORYINFO
'        llUpperSplit = 0
'        mBuildSplitCategoryInfo "C", llUpperSplit, tlPartRegionDef(0), tlSplitCategoryInfo()
'        ReDim Preserve tlSplitCategoryInfo(0 To llUpperSplit) As SPLITCATEGORYINFO
'
'        ReDim tmRegionDefinition(0 To 0) As REGIONDEFINITION
'        ReDim tmSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO

    '11/9/14
    slReplacementCue = ""
    If tlAstInfo.iSpotType = 5 Then
        slSQLQuery = "SELECT *"
        slSQLQuery = slSQLQuery + " FROM irt"
        slSQLQuery = slSQLQuery + " WHERE (irtShttCode= " & ilShttCode & " And irtLstCode = " & llLstCode & ")"
        Set rst_irt = gSQLSelectCall(slSQLQuery)
        If Not rst_irt.EOF Then
            'Imported region copy
            slReplacementCue = rst_irt!irtXDSCue
            If rst_irt!irtType <> "B" Then
                gGetRegionCopy = 1
                tlAstInfo.lLstBkoutLstCode = 0
                tlAstInfo.lIrtCode = rst_irt!irtCode
                slCartNo = rst_irt!irtCart
                slProduct = rst_irt!irtProduct
                slISCI = rst_irt!irtISCI
                slCreativeTitle = rst_irt!irtCreativeTitle
                llCrfCsfCode = 0
                llCrfCode = 0
                llRsfCode = 0
                llCpfCode = 0
            Else
                ilCifAdfCode = rst_irt!irtAdfCode
                slCartNo = rst_irt!irtCart
                slProduct = rst_irt!irtProduct
                slISCI = rst_irt!irtISCI
                slCreativeTitle = rst_irt!irtCreativeTitle
                llCrfCsfCode = 0
                llCrfCode = 0
                llRsfCode = 0
                llCpfCode = 0
                slSQLQuery = "SELECT * FROM lst WHERE (lstCode = " & llLstCode & ")"
                Set lst_rst = gSQLSelectCall(slSQLQuery)
                If lst_rst.EOF Then
                    llLstCode = tlAstInfo.lLstCode
                    slSQLQuery = "SELECT * FROM lst WHERE (lstCode = " & llLstCode & ")"
                    Set lst_rst = gSQLSelectCall(slSQLQuery)
                End If
                slLogDate = Format$(lst_rst!lstLogDate, sgShowDateForm)
                slLogTime = Format$(lst_rst!lstLogTime, sgShowTimeWOSecForm)
                llCode = mCreateBlackoutLst(hlAst, tlAstInfo, ilVefCode, llLstCode, ilCifAdfCode, slCartNo, slProduct, slISCI, slLogDate, slLogTime, 0, llCpfCode, llCrfCsfCode, 0)
                If llCode > 0 Then
                    tlAstInfo.lIrtCode = rst_irt!irtCode
                    gGetRegionCopy = 2
                Else
                    gGetRegionCopy = -1
                    tlAstInfo.lIrtCode = 0
                End If
            End If
            Exit Function
        End If
    End If
    llSdf = mBinarySearchSdfCode(llSdfCode)
    tlAstInfo.lIrtCode = 0
    Do While llSdf <> -1
        DoEvents
        'If igExportSource = 2 Then
        '    DoEvents
        'End If
        tlPartRegionDef(0) = tmRegionDefinitionForSpots(tmRegionAssignmentInfo(llSdf).lRDIndex)
        ReDim tlSplitCategoryInfo(0 To tmRegionAssignmentInfo(llSdf).lSCIEndIndex - tmRegionAssignmentInfo(llSdf).lSCIStartIndex + 1) As SPLITCATEGORYINFO
        For llLoop = tmRegionAssignmentInfo(llSdf).lSCIStartIndex To tmRegionAssignmentInfo(llSdf).lSCIEndIndex Step 1
            tlSplitCategoryInfo(llLoop - tmRegionAssignmentInfo(llSdf).lSCIStartIndex) = tmSplitCategoryInfoForSpots(llLoop)
        Next llLoop
        
        ReDim tmRegionDefinition(0 To 0) As REGIONDEFINITION
        ReDim tmSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
        
        lgSTime14 = timeGetTime
        gSeparateRegions tlPartRegionDef(), tlSplitCategoryInfo(), tmRegionDefinition(), tmSplitCategoryInfo()

        ilRet = gRegionTestDefinition(ilShttCode, ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, tmRegionDefinition(), tmSplitCategoryInfo(), llRegionIndex, slGroupInfo)
        lgETime14 = timeGetTime
        lgTtlTime14 = lgTtlTime14 + (lgETime14 - lgSTime14)
        
        If ilRet Then
            lgSTime21 = timeGetTime
            ilRet = gGetCopy(tmRegionDefinition(llRegionIndex).sPtType, tmRegionDefinition(llRegionIndex).lCopyCode, tmRegionDefinition(llRegionIndex).lCrfCode, True, slCartNo, slProduct, slISCI, slCreativeTitle, llCrfCsfCode, llCpfCode, ilCifAdfCode, ilVefCode)
            lgETime21 = timeGetTime
            lgTtlTime21 = lgTtlTime21 + (lgETime21 - lgSTime21)
            
            If ilRet Then
                llRsfCode = tmRegionDefinition(llRegionIndex).lRsfCode
                llCrfCode = tmRegionDefinition(llRegionIndex).lCrfCode
                llCifCode = tmRegionDefinition(llRegionIndex).lCopyCode
                lgSTime6 = timeGetTime
                '12/13/13: Update astAdfCode, astRsfCode and astCpfCode
                tmAstSrchKey.lCode = tlAstInfo.lCode
                ilRet = btrGetEqual(hlAst, tmAst, imAstRecLen, tmAstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
                If ilRet = BTRV_ERR_NONE Then
                    tmAst.iAdfCode = ilCifAdfCode
                    tmAst.lRsfCode = llRsfCode
                    tmAst.lCpfCode = llCpfCode
                    ilRet = btrUpdate(hlAst, tmAst, imAstRecLen)
                Else
                    slSQLQuery = "UPDATE ast SET astAdfCode = " & ilCifAdfCode & ","
                    slSQLQuery = slSQLQuery & "astRsfCode = " & llRsfCode & ","
                    slSQLQuery = slSQLQuery & "astCpfCode = " & llCpfCode
                    slSQLQuery = slSQLQuery & " where astCode = " & tlAstInfo.lCode
                    'cnn.Execute slSQLQuery
                    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        gGetRegionCopy = False
                        Exit Function
                    End If
                End If
                lgETime6 = timeGetTime
                lgTtlTime6 = lgTtlTime6 + (lgETime6 - lgSTime6)
                             
                gGetRegionCopy = 1
                mUpdatePoolInfo ilCifAdfCode, llCrfCode, tmRegionDefinition(llRegionIndex).iPoolNextFinal, tmRegionDefinition(llRegionIndex)
                If ilAssginBlackoutCopy Then
                    slSQLQuery = "SELECT * FROM lst WHERE (lstCode = " & llLstCode & ")"
                    Set lst_rst = gSQLSelectCall(slSQLQuery)
                    If Not lst_rst.EOF Then
                        If igExportSource = 2 Then
                            DoEvents
                        End If
                        llBkoutLstIndex = -1
                        slLogDate = Format$(lst_rst!lstLogDate, sgShowDateForm)
                        slLogTime = Format$(lst_rst!lstLogTime, sgShowTimeWOSecForm)
                        llLogDate = DateValue(slLogDate)
                        llLogTime = gTimeToLong(slLogTime, False)
                        '11/18/11: search tmBloutLst only if previously defined
                        If tlAstInfo.lPrevBkoutLstCode > 0 Then
                            For llLst = 0 To UBound(tmBkoutLst) - 1 Step 1
                                If igExportSource = 2 Then
                                    DoEvents
                                End If
                                '11/18/11: replace with looking for Previous Blackout assigned to the ast
                                'If tmBkoutLst(llLst).tLST.lBkoutLstCode = llLstCode Then
                                If tmBkoutLst(llLst).tLST.lCode = tlAstInfo.lPrevBkoutLstCode Then
                                    'If DateValue(tmBkoutLst(llLst).tLST.sLogDate) = llLogDate Then
                                    '    If gTimeToLong(tmBkoutLst(llLst).tLST.sLogTime, False) = llLogTime Then
                                            llBkoutLstIndex = llLst
                                            If (Trim$(slISCI) = Trim$(tmBkoutLst(llLst).tLST.sISCI)) And (Trim$(slCartNo) = Trim$(tmBkoutLst(llLst).tLST.sCart)) Then
                                                If (llCpfCode = tmBkoutLst(llLst).tLST.lCpfCode) And (llCrfCsfCode = tmBkoutLst(llLst).tLST.lCrfCsfCode) Then
                                                    lgCount6 = lgCount6 + 1  'Number of blackuots Unchanged
                                                    If tlAstInfo.lLstCode <> tmBkoutLst(llLst).tLST.lCode Then
                                                        lgSTime6 = timeGetTime
                                                        tmAstSrchKey.lCode = tlAstInfo.lCode
                                                        ilRet = btrGetEqual(hlAst, tmAst, imAstRecLen, tmAstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            tmAst.lLsfCode = tmBkoutLst(llLst).tLST.lCode
                                                            ilRet = btrUpdate(hlAst, tmAst, imAstRecLen)
                                                        Else
                                                            slSQLQuery = "UPDATE ast SET astLsfCode = " & tmBkoutLst(llLst).tLST.lCode & " where astCode = " & tlAstInfo.lCode
                                                            'cnn.Execute slSQLQuery
                                                            If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                                                                gGetRegionCopy = False
                                                                Exit Function
                                                            End If
                                                        End If
                                                        lgETime6 = timeGetTime
                                                        lgTtlTime6 = lgTtlTime6 + (lgETime6 - lgSTime6)
                                                        tlAstInfo.lLstBkoutLstCode = tlAstInfo.lLstCode
                                                    End If
                                                    '1/11/15: Moved up
                                                    'tlAstInfo.lLstBkoutLstCode = tlAstInfo.lLstCode
                                                    tlAstInfo.lLstCode = tmBkoutLst(llLst).tLST.lCode
                                                    tlAstInfo.iAdfCode = tmBkoutLst(llLst).tLST.iAdfCode
                                                    tlAstInfo.sProd = tmBkoutLst(llLst).tLST.sProd
                                                    '1/7/16: Check that product did not change
                                                    llRet = gBinarySearchCifCpf(tmBkoutLst(llLst).tLST.lCifCode)
                                                    If llRet <> -1 Then
                                                        If Trim$(tgCifCpfInfo1(llRet).cpfName) <> Trim$(tmBkoutLst(llLst).tLST.sProd) Then
                                                            tlAstInfo.sProd = tgCifCpfInfo1(llRet).cpfName
                                                            slSQLQuery = "UPDATE lst SET lstProd = '" & tlAstInfo.sProd & "' where lstCode = " & tmBkoutLst(llLst).tLST.lCode
                                                            'cnn.Execute slSQLQuery
                                                            If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                                                                gGetRegionCopy = False
                                                                Exit Function
                                                            End If
                                                        End If
                                                    End If
                                                    tlAstInfo.lCntrNo = 0
                                                    gGetRegionCopy = 2
                                                    lgETime9 = timeGetTime
                                                    lgTtlTime9 = lgTtlTime9 + (lgETime9 - lgSTime9)
                                                    Exit Function
                                                End If
                                            End If
                                    '    End If
                                    'End If
                                End If
                            Next llLst
                        End If
                        If (lst_rst!lstAdfCode <> ilCifAdfCode) And (llBkoutLstIndex = -1) Then
                            If igExportSource = 2 Then
                                DoEvents
                            End If
                            For ilLoop = 0 To 6 Step 1
                                ilWeekDay(ilLoop) = 0
                            Next ilLoop
                            ilWeekDay(Weekday(slLogDate, vbMonday) - 1) = 1
                            '11/18/11:
                            'If lst_rst!lstBkoutLstCode > 0 Then
                            If tlAstInfo.lPrevBkoutLstCode > 0 Then
                                llLstCode = lst_rst!lstBkoutLstCode
                                '11/18/11: Remove blackout
                                'slSQLQuery = "DELETE FROM lst WHERE (lstCode = " & lst_rst!lstCode & ")"
                                slSQLQuery = "DELETE FROM lst WHERE (lstCode = " & tlAstInfo.lPrevBkoutLstCode & ")"
                                ' Dan M 9/18/09 function now returns a long
                                'ilRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                                gSQLWaitNoMsgBox slSQLQuery, False
                                slSQLQuery = "SELECT * FROM lst WHERE (lstCode = " & llLstCode & ")"
                                Set lst_rst = gSQLSelectCall(slSQLQuery)
                                If lst_rst.EOF Then
                                    llLstCode = tlAstInfo.lLstCode
                                    slSQLQuery = "SELECT * FROM lst WHERE (lstCode = " & llLstCode & ")"
                                    Set lst_rst = gSQLSelectCall(slSQLQuery)
                                End If
                            End If
                            ''Create LST
                            'Do
                                If igExportSource = 2 Then
                                    DoEvents
                                End If
                                lgCount5 = lgCount5 + 1  'Number of blackuots LSTs created
                                'slSQLQuery = "SELECT MAX(lstCode) from lst"
                                'Set rst = gSQLSelectCall(slSQLQuery)
                                'If Not rst.EOF Then
                                '    llCode = rst(0).Value + 1
                                'Else
                                '    llCode = 1
                                'End If
'                                ilRet = 0
'                                slSQLQuery = "Insert Into lst ( "
'                                slSQLQuery = slSQLQuery & "lstCode, "
'                                slSQLQuery = slSQLQuery & "lstType, "
'                                slSQLQuery = slSQLQuery & "lstSdfCode, "
'                                slSQLQuery = slSQLQuery & "lstCntrNo, "
'                                slSQLQuery = slSQLQuery & "lstAdfCode, "
'                                slSQLQuery = slSQLQuery & "lstAgfCode, "
'                                slSQLQuery = slSQLQuery & "lstProd, "
'                                slSQLQuery = slSQLQuery & "lstLineNo, "
'                                slSQLQuery = slSQLQuery & "lstLnVefCode, "
'                                slSQLQuery = slSQLQuery & "lstStartDate, "
'                                slSQLQuery = slSQLQuery & "lstEndDate, "
'                                slSQLQuery = slSQLQuery & "lstMon, "
'                                slSQLQuery = slSQLQuery & "lstTue, "
'                                slSQLQuery = slSQLQuery & "lstWed, "
'                                slSQLQuery = slSQLQuery & "lstThu, "
'                                slSQLQuery = slSQLQuery & "lstFri, "
'                                slSQLQuery = slSQLQuery & "lstSat, "
'                                slSQLQuery = slSQLQuery & "lstSun, "
'                                slSQLQuery = slSQLQuery & "lstSpotsWk, "
'                                slSQLQuery = slSQLQuery & "lstPriceType, "
'                                slSQLQuery = slSQLQuery & "lstPrice, "
'                                slSQLQuery = slSQLQuery & "lstSpotType, "
'                                slSQLQuery = slSQLQuery & "lstLogVefCode, "
'                                slSQLQuery = slSQLQuery & "lstLogDate, "
'                                slSQLQuery = slSQLQuery & "lstLogTime, "
'                                slSQLQuery = slSQLQuery & "lstDemo, "
'                                slSQLQuery = slSQLQuery & "lstAud, "
'                                slSQLQuery = slSQLQuery & "lstISCI, "
'                                slSQLQuery = slSQLQuery & "lstWkNo, "
'                                slSQLQuery = slSQLQuery & "lstBreakNo, "
'                                slSQLQuery = slSQLQuery & "lstPositionNo, "
'                                slSQLQuery = slSQLQuery & "lstSeqNo, "
'                                slSQLQuery = slSQLQuery & "lstZone, "
'                                slSQLQuery = slSQLQuery & "lstCart, "
'                                slSQLQuery = slSQLQuery & "lstCpfCode, "
'                                slSQLQuery = slSQLQuery & "lstCrfCsfCode, "
'                                slSQLQuery = slSQLQuery & "lstStatus, "
'                                slSQLQuery = slSQLQuery & "lstLen, "
'                                slSQLQuery = slSQLQuery & "lstUnits, "
'                                slSQLQuery = slSQLQuery & "lstCifCode, "
'                                slSQLQuery = slSQLQuery & "lstAnfCode, "
'                                slSQLQuery = slSQLQuery & "lstEvtIDCefCode, "
'                                slSQLQuery = slSQLQuery & "lstSplitNetwork, "
'                                slSQLQuery = slSQLQuery & "lstRafCode, "
'                                slSQLQuery = slSQLQuery & "lstFsfCode, "
'                                slSQLQuery = slSQLQuery & "lstGsfCode, "
'                                slSQLQuery = slSQLQuery & "lstImportedSpot, "
'                                slSQLQuery = slSQLQuery & "lstBkoutLstCode, "
'                                slSQLQuery = slSQLQuery & "lstLnStartTime, "
'                                slSQLQuery = slSQLQuery & "lstLnEndTime, "
'                                slSQLQuery = slSQLQuery & "lstUnused "
'                                slSQLQuery = slSQLQuery & ") "
'                                slSQLQuery = slSQLQuery & "Values ( "
'                                slSQLQuery = slSQLQuery & "Replace" & ", "
'                                slSQLQuery = slSQLQuery & lst_rst!lstType & ", "
'                                slSQLQuery = slSQLQuery & lst_rst!lstSdfCode & ", "
'                                slSQLQuery = slSQLQuery & 0 & ", "
'                                slSQLQuery = slSQLQuery & ilCifAdfCode & ", "
'                                slSQLQuery = slSQLQuery & 0 & ", "
'                                slSQLQuery = slSQLQuery & "'" & gFixQuote(slProduct) & "', "
'                                slSQLQuery = slSQLQuery & 0 & ", "
'                                slSQLQuery = slSQLQuery & 0 & ", "
'                                slSQLQuery = slSQLQuery & "'" & Format$(slLogDate, sgSQLDateForm) & "', "
'                                slSQLQuery = slSQLQuery & "'" & Format$(slLogDate, sgSQLDateForm) & "', "
'                                slSQLQuery = slSQLQuery & ilWeekDay(0) & ", "
'                                slSQLQuery = slSQLQuery & ilWeekDay(1) & ", "
'                                slSQLQuery = slSQLQuery & ilWeekDay(2) & ", "
'                                slSQLQuery = slSQLQuery & ilWeekDay(3) & ", "
'                                slSQLQuery = slSQLQuery & ilWeekDay(4) & ", "
'                                slSQLQuery = slSQLQuery & ilWeekDay(5) & ", "
'                                slSQLQuery = slSQLQuery & ilWeekDay(6) & ", "
'                                slSQLQuery = slSQLQuery & 1 & ", "
'                                slSQLQuery = slSQLQuery & 1 & ", "
'                                slSQLQuery = slSQLQuery & 0 & ", "
'                                slSQLQuery = slSQLQuery & 5 & ", "
'                                slSQLQuery = slSQLQuery & ilVefCode & ", "
'                                slSQLQuery = slSQLQuery & "'" & Format$(slLogDate, sgSQLDateForm) & "', "
'                                slSQLQuery = slSQLQuery & "'" & Format$(slLogTime, sgSQLTimeForm) & "', "
'                                slSQLQuery = slSQLQuery & "'" & "" & "', "
'                                slSQLQuery = slSQLQuery & 0 & ", "
'                                slSQLQuery = slSQLQuery & "'" & gFixQuote(slISCI) & "', "
'                                slSQLQuery = slSQLQuery & lst_rst!lstWkNo & ", "
'                                slSQLQuery = slSQLQuery & lst_rst!lstBreakNo & ", "
'                                slSQLQuery = slSQLQuery & lst_rst!lstPositionNo & ", "
'                                slSQLQuery = slSQLQuery & lst_rst!lstSeqNo & ", "
'                                slSQLQuery = slSQLQuery & "'" & gFixQuote(lst_rst!lstZone) & "', "
'                                slSQLQuery = slSQLQuery & "'" & gFixQuote(slCartNo) & "', "
'                                slSQLQuery = slSQLQuery & llCpfCode & ", "
'                                slSQLQuery = slSQLQuery & llCrfCsfCode & ", "
'                                slSQLQuery = slSQLQuery & lst_rst!lstStatus & ", "
'                                slSQLQuery = slSQLQuery & lst_rst!lstLen & ", "
'                                slSQLQuery = slSQLQuery & lst_rst!lstUnits & ", "
'                                slSQLQuery = slSQLQuery & tmRegionDefinition(llRegionIndex).lCopyCode & ", "
'                                slSQLQuery = slSQLQuery & lst_rst!lstAnfCode & ", "
'                                slSQLQuery = slSQLQuery & lst_rst!lstEvtIDCefCode & ", "
'                                If (Asc(lst_rst!lstsplitnetwork) <> Asc("N")) And (Asc(lst_rst!lstsplitnetwork) <> Asc("Y")) Then
'                                    slSQLQuery = slSQLQuery & "'" & "N" & "', "
'                                Else
'                                    slSQLQuery = slSQLQuery & "'" & gFixQuote(lst_rst!lstsplitnetwork) & "', "
'                                End If
'                                slSQLQuery = slSQLQuery & tmRegionDefinition(llRegionIndex).lRafCode & ", "
'                                slSQLQuery = slSQLQuery & lst_rst!lstFsfCode & ", "
'                                slSQLQuery = slSQLQuery & lst_rst!lstGsfCode & ", "
'                                If (Asc(lst_rst!lstImportedSpot) <> Asc("N")) And (Asc(lst_rst!lstImportedSpot) <> Asc("Y")) Then
'                                    slSQLQuery = slSQLQuery & "'" & "N" & "', "  'gFixQuote(lst_rst!lstImportedSpot) & "', "
'                                Else
'                                    slSQLQuery = slSQLQuery & "'" & gFixQuote(lst_rst!lstImportedSpot) & "', "
'                                End If
'                                slSQLQuery = slSQLQuery & llLstCode & ", "
'                                slSQLQuery = slSQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
'                                slSQLQuery = slSQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
'                                slSQLQuery = slSQLQuery & "'" & "" & "' "
'                                slSQLQuery = slSQLQuery & ") "
''                                If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
'                                bgIgnoreDuplicateError = True
'                                llCode = gInsertAndReturnCode(slSQLQuery, "lst", "lstCode", "Replace")
'                                bgIgnoreDuplicateError = False
'                                If llCode <= 0 Then
'                                    gGetRegionCopy = -1
'                                    Exit Function
'                                End If
'                            'Loop While ilRet <> 0
'                            lgSTime6 = timeGetTime
'                            tmAstSrchKey.lCode = tlAstInfo.lCode
'                            ilRet = btrGetEqual(hlAst, tmAst, imAstRecLen, tmAstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
'                            If ilRet = BTRV_ERR_NONE Then
'                                tmAst.lLsfCode = llCode
'                                ilRet = btrUpdate(hlAst, tmAst, imAstRecLen)
'                            Else
'                                slSQLQuery = "UPDATE ast SET astLsfCode = " & llCode & " where astCode = " & tlAstInfo.lCode
'                                cnn.Execute slSQLQuery
'                            End If
'                            lgETime6 = timeGetTime
'                            lgTtlTime6 = lgTtlTime6 + (lgETime6 - lgSTime6)
'                            tlAstInfo.lLstBkoutLstCode = tlAstInfo.lLstCode
'                            tlAstInfo.lLstCode = llCode
'                            tlAstInfo.iAdfCode = ilCifAdfCode
'                            tlAstInfo.sProd = slProduct
'                            tlAstInfo.lCntrNo = 0
'                            slSQLQuery = "SELECT * FROM lst WHERE (lstCode = " & llCode & ")"
'                            Set lst_rst = gSQLSelectCall(slSQLQuery)
'                            If Not lst_rst.EOF Then
'                                gCreateUDTforLST lst_rst, tmBkoutLst(UBound(tmBkoutLst)).tLST
'                                tmBkoutLst(UBound(tmBkoutLst)).iDelete = False
'                                '12/12/14
'                                tmBkoutLst(UBound(tmBkoutLst)).bMatched = False
'                                ReDim Preserve tmBkoutLst(0 To UBound(tmBkoutLst) + 1) As BKOUTLST
'                            End If
'                            gGetRegionCopy = 2
'                            lgETime10 = timeGetTime
'                            lgTtlTime10 = lgTtlTime10 + (lgETime10 - lgSTime10)
'                            If igExportSource = 2 Then
'                                DoEvents
'                            End If
                            llCode = mCreateBlackoutLst(hlAst, tlAstInfo, ilVefCode, llLstCode, ilCifAdfCode, slCartNo, slProduct, slISCI, slLogDate, slLogTime, tmRegionDefinition(llRegionIndex).lCopyCode, llCpfCode, llCrfCsfCode, tmRegionDefinition(llRegionIndex).lRafCode)
                            If llCode > 0 Then
                                gGetRegionCopy = 2
                            Else
                                gGetRegionCopy = -1
                                Exit Function
                            End If
                        Else
                            'Update copy if Blackout
                            '11/18/11
                            'If (lst_rst!lstBkoutLstCode > 0) And (llBkoutLstIndex <> -1) Then
                            If (tlAstInfo.lPrevBkoutLstCode > 0) And (llBkoutLstIndex <> -1) Then
                                If igExportSource = 2 Then
                                    DoEvents
                                End If
                                lgCount4 = lgCount4 + 1  'How many blackouts that need to be updated
                                slSQLQuery = "Update lst Set "
                                slSQLQuery = slSQLQuery & "lstISCI = '" & gFixQuote(slISCI) & "', "
                                slSQLQuery = slSQLQuery & "lstCart = '" & gFixQuote(slCartNo) & "', "
                                slSQLQuery = slSQLQuery & "lstCpfCode = " & llCpfCode & ", "
                                slSQLQuery = slSQLQuery & "lstCrfCsfCode = " & llCrfCsfCode & ", "
                                slSQLQuery = slSQLQuery & "lstCifCode = " & tmRegionDefinition(llRegionIndex).lCopyCode & ", "
                                slSQLQuery = slSQLQuery & "lstAdfCode = " & ilCifAdfCode & ", "
                                slSQLQuery = slSQLQuery & "lstProd = '" & gFixQuote(slProduct) & "'"
                                slSQLQuery = slSQLQuery & " where lstCode = " & tlAstInfo.lPrevBkoutLstCode
                                'cnn.Execute slSQLQuery
                                If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                                    gGetRegionCopy = False
                                    Exit Function
                                End If
                                tmBkoutLst(llBkoutLstIndex).tLST.sISCI = slISCI
                                tmBkoutLst(llBkoutLstIndex).tLST.sCart = slCartNo
                                tmBkoutLst(llBkoutLstIndex).tLST.lCpfCode = llCpfCode
                                tmBkoutLst(llBkoutLstIndex).tLST.lCrfCsfCode = llCrfCsfCode
                                tmBkoutLst(llBkoutLstIndex).tLST.iAdfCode = ilCifAdfCode
                                tmBkoutLst(llBkoutLstIndex).tLST.sProd = slProduct
                                If tlAstInfo.lLstCode <> tlAstInfo.lPrevBkoutLstCode Then
                                    '11/18/11
                                    lgSTime6 = timeGetTime
                                    tmAstSrchKey.lCode = tlAstInfo.lCode
                                    ilRet = btrGetEqual(hlAst, tmAst, imAstRecLen, tmAstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
                                    If ilRet = BTRV_ERR_NONE Then
                                        tmAst.lLsfCode = tlAstInfo.lPrevBkoutLstCode
                                        ilRet = btrUpdate(hlAst, tmAst, imAstRecLen)
                                    Else
                                        slSQLQuery = "UPDATE ast SET astLsfCode = " & tlAstInfo.lPrevBkoutLstCode & " where astCode = " & tlAstInfo.lCode
                                        'cnn.Execute slSQLQuery
                                        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                                            gGetRegionCopy = False
                                            Exit Function
                                        End If
                                    End If
                                    lgETime6 = timeGetTime
                                    lgTtlTime6 = lgTtlTime6 + (lgETime6 - lgSTime6)
                                End If
                                '11/18/11
                                tlAstInfo.lLstBkoutLstCode = tlAstInfo.lLstCode
                                tlAstInfo.lLstCode = tlAstInfo.lPrevBkoutLstCode
                                tlAstInfo.iAdfCode = ilCifAdfCode
                                tlAstInfo.sProd = slProduct
                                tlAstInfo.lCntrNo = 0
                                gGetRegionCopy = 2
                                lgETime11 = timeGetTime
                                lgTtlTime11 = lgTtlTime11 + (lgETime11 - lgSTime11)
                            End If
                        End If
                    Else
                    End If
                Else
                    tlAstInfo.lLstBkoutLstCode = 0
                End If
                Exit Function
            End If
        End If
        llSdf = llSdf + 1
        If llSdf >= UBound(tmRegionAssignmentInfo) Then
            llSdf = -1
        Else
            If llSdfCode <> tmRegionAssignmentInfo(llSdf).lSdfCode Then
                llSdf = -1
            End If
        End If
    Loop
    'Next llRegionDefinitionIndex
    'Check if region copy previously defined, if so, remove it
    lgSTime12 = timeGetTime
    If ilAssginBlackoutCopy Then
        'slSQLQuery = "SELECT lstBkoutLstCode FROM lst WHERE (lstCode = " & llLstCode & ")"
        'Set lst_rst = gSQLSelectCall(slSQLQuery)
        'If Not lst_rst.EOF Then
        '    If lst_rst!lstBkoutLstCode > 0 Then
        '        llBkoutLstCode = lst_rst!lstBkoutLstCode
        '        slSQLQuery = "SELECT * FROM lst WHERE (lstCode = " & lst_rst!lstBkoutLstCode & ")"
        For llLst = 0 To UBound(tmBkoutLst) - 1 Step 1
            If igExportSource = 2 Then
                DoEvents
            End If
            If tmBkoutLst(llLst).tLST.lCode = llLstCode Then
                If tmBkoutLst(llLst).tLST.lBkoutLstCode > 0 Then
                    lgCount3 = lgCount3 + 1
                    llBkoutLstCode = tmBkoutLst(llLst).tLST.lBkoutLstCode
                    slSQLQuery = "SELECT * FROM lst WHERE (lstCode = " & tmBkoutLst(llLst).tLST.lBkoutLstCode & ")"
                    Set lst_rst = gSQLSelectCall(slSQLQuery)
                    If Not lst_rst.EOF Then
                        If igExportSource = 2 Then
                            DoEvents
                        End If
                        'Backout that needed to be removed
                        lgCount7 = lgCount7 + 1  'Number of blackuots Removed
                        lgSTime6 = timeGetTime
                        tmAstSrchKey.lCode = tlAstInfo.lCode
                        ilRet = btrGetEqual(hlAst, tmAst, imAstRecLen, tmAstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
                        If ilRet = BTRV_ERR_NONE Then
                            tmAst.iAdfCode = lst_rst!lstAdfCode
                            tmAst.lRsfCode = 0
                            tmAst.lCpfCode = lst_rst!lstCpfCode
                            tmAst.lLsfCode = lst_rst!lstCode
                            ilRet = btrUpdate(hlAst, tmAst, imAstRecLen)
                        Else
                            'slSQLQuery = "UPDATE ast SET astLsfCode = " & lst_rst!lstCode & " where astCode = " & tlAstInfo.lCode
                            slSQLQuery = "UPDATE ast SET astAdfCode = " & lst_rst!lstAdfCode & ","
                            slSQLQuery = slSQLQuery & "astRsfCode = " & 0 & ","
                            slSQLQuery = slSQLQuery & "astCpfCode = " & lst_rst!lstCpfCode & ","
                            slSQLQuery = slSQLQuery & "astLsfCode = " & lst_rst!lstCode
                            slSQLQuery = slSQLQuery & " where astCode = " & tlAstInfo.lCode
                            'cnn.Execute slSQLQuery
                            If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                                gGetRegionCopy = False
                                Exit Function
                            End If
                        End If
                        lgETime6 = timeGetTime
                        lgTtlTime6 = lgTtlTime6 + (lgETime6 - lgSTime6)
                        
                        tlAstInfo.lLstCode = lst_rst!lstCode
                        tlAstInfo.iAdfCode = lst_rst!lstAdfCode
                        tlAstInfo.sProd = lst_rst!lstProd
                        tlAstInfo.lCntrNo = lst_rst!lstCntrNo
                        tlAstInfo.lLstBkoutLstCode = 0
                        tlAstInfo.lCpfCode = lst_rst!lstCpfCode
                        'slSQLQuery = "DELETE FROM lst WHERE (lstCode = " & llLstCode & " AND lstBkoutLstCode = " & llBkoutLstCode & ")"
                        'ilRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                        tmBkoutLst(llLst).iDelete = True
                        lgETime12 = timeGetTime
                        lgTtlTime12 = lgTtlTime12 + (lgETime12 - lgSTime12)
                    End If
                End If
                Exit Function
            End If
        'End If
        Next llLst
    Else
        tlAstInfo.lLstBkoutLstCode = 0
    End If
    lgETime12 = timeGetTime
    lgTtlTime13 = lgTtlTime13 + (lgETime12 - lgSTime12)
    
    'gGetRegionCopy = ilRet
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-gGetRegionCopy"
    gGetRegionCopy = -1
    Exit Function
ErrHand1:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-gGetRegionCopy2"
    gGetRegionCopy = -1
End Function

Public Function gGetCopy(slPtType As String, llCopyCode As Long, llCrfCode As Long, ilWegenerOLA As Integer, slCartNo As String, slProduct As String, slISCI As String, slCreativeTitle As String, llCrfCsfCode As Long, llCpfCode As Long, ilCifAdfCode As Integer, Optional ilVefCode As Integer = -1) As Integer
    'To filter comment, include the vefCode
    Dim ilRet As Integer
    Dim llRet As Long
    Dim llUpper As Long
    '8/31/16: Hide comment
    Dim ilVff As Integer
    Dim slHideCommOnWeb As String
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    
    gGetCopy = False
    slHideCommOnWeb = "N"
    If ilVefCode <> -1 Then
        ilVff = gBinarySearchVff(CLng(ilVefCode))
        If ilVff <> -1 Then
            slHideCommOnWeb = tgVffInfo(ilVff).sHideCommOnWeb
        End If
    End If
    slCartNo = ""
    slProduct = ""
    slISCI = ""
    slCreativeTitle = ""
    llCrfCsfCode = 0
    If slPtType = "1" Then
        llRet = gBinarySearchCifCpf(llCopyCode)
        If llRet <> -1 Then
            If igExportSource = 2 Then
                DoEvents
            End If
            lgCount9 = lgCount9 + 1
            ilCifAdfCode = tgCifCpfInfo1(llRet).cifAdfCode
            If IsNull(tgCifCpfInfo1(llRet).cifName) = False Then
                If Asc(tgCifCpfInfo1(llRet).cifName) <> 0 Then
                    slCartNo = Trim$(tgCifCpfInfo1(llRet).cifName)
                End If
            End If
            If tgCifCpfInfo1(llRet).cifMcfCode > 0 Then
                ilRet = gBinarySearchMcf(tgCifCpfInfo1(llRet).cifMcfCode)
                If ilRet <> -1 Then
                    slCartNo = Trim$(tgMediaCodesInfo(ilRet).sName) & slCartNo
                End If
            Else
                If (slCartNo = "") And (ilWegenerOLA) And (Trim$(tgCifCpfInfo1(llRet).cifReel) <> "") Then
                    slCartNo = tgCifCpfInfo1(llRet).cifReel
                End If
            End If
            llCpfCode = tgCifCpfInfo1(llRet).cifCpfCode
            If llCpfCode > 0 Then
                If IsNull(tgCifCpfInfo1(llRet).cpfName) = False Then
                    If Asc(tgCifCpfInfo1(llRet).cpfName) <> 0 Then
                        slProduct = Trim$(tgCifCpfInfo1(llRet).cpfName)
                    End If
                End If
                If IsNull(tgCifCpfInfo1(llRet).cpfISCI) = False Then
                    If Asc(tgCifCpfInfo1(llRet).cpfISCI) <> 0 Then
                        slISCI = Trim$(tgCifCpfInfo1(llRet).cpfISCI)
                    End If
                End If
                If IsNull(tgCifCpfInfo1(llRet).cpfCreative) = False Then
                    If Asc(tgCifCpfInfo1(llRet).cpfCreative) <> 0 Then
                        slCreativeTitle = Trim$(tgCifCpfInfo1(llRet).cpfCreative)
                    End If
                End If
            End If
            If (llCrfCode > 0) And (slHideCommOnWeb <> "Y") Then
                lgSTime22 = timeGetTime
                llRet = gBinarySearchCrf(llCrfCode)
                If llRet <> -1 Then
                    llCrfCsfCode = tgCrfInfo1(llRet).crfCsfCode
                Else
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    slSQLQuery = "Select crfCsfCode from CRF_Copy_Rot_Header"
                    slSQLQuery = slSQLQuery & " Where (crfCode = " & llCrfCode & ")"
                    Set crf_rst = gSQLSelectCall(slSQLQuery)
                    If Not crf_rst.EOF Then
                        llCrfCsfCode = crf_rst!crfCsfCode
                    End If
                End If
                lgETime22 = timeGetTime
                lgTtlTime22 = lgTtlTime22 + (lgETime22 - lgSTime22)
            End If
            gGetCopy = True
            Exit Function
        Else
            If igExportSource = 2 Then
                DoEvents
            End If
            lgCount10 = lgCount10 + 1
            slSQLQuery = "Select * from CIF_Copy_Inventory"
            slSQLQuery = slSQLQuery & " Where (cifCode = " & llCopyCode & ")"
            Set cif_rst = gSQLSelectCall(slSQLQuery)
            If Not cif_rst.EOF Then
                ilCifAdfCode = cif_rst!cifAdfCode
                If IsNull(cif_rst!cifName) = False Then
                    If Asc(cif_rst!cifName) <> 0 Then
                        slCartNo = Trim$(cif_rst!cifName)
                    End If
                End If
                If cif_rst!cifMcfCode > 0 Then
                    ilRet = gBinarySearchMcf(cif_rst!cifMcfCode)
                    If ilRet <> -1 Then
                        slCartNo = Trim$(tgMediaCodesInfo(ilRet).sName) & slCartNo
                    End If
                Else
                    If (slCartNo = "") And (ilWegenerOLA) Then
                        slCartNo = cif_rst!cifReel
                    End If
                End If
                llCpfCode = cif_rst!cifCpfCode
                If llCpfCode > 0 Then
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    slSQLQuery = "Select * From CPF_Copy_Prodct_ISCI"
                    slSQLQuery = slSQLQuery & " Where (cpfCode = " & llCpfCode & ")"
                    Set cpf_rst = gSQLSelectCall(slSQLQuery)
                    If Not cpf_rst.EOF Then
                        If IsNull(cpf_rst!cpfName) = False Then
                            If Asc(cpf_rst!cpfName) <> 0 Then
                                slProduct = Trim$(cpf_rst!cpfName)
                            End If
                        End If
                        If IsNull(cpf_rst!cpfISCI) = False Then
                            If Asc(cpf_rst!cpfISCI) <> 0 Then
                                slISCI = Trim$(cpf_rst!cpfISCI)
                            End If
                        End If
                        If IsNull(cpf_rst!cpfCreative) = False Then
                            If Asc(cpf_rst!cpfCreative) <> 0 Then
                                slCreativeTitle = Trim$(cpf_rst!cpfCreative)
                            End If
                        End If
                    End If
                End If
                If (llCrfCode > 0) And (slHideCommOnWeb <> "Y") Then
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    lgSTime22 = timeGetTime
                    slSQLQuery = "Select crfCsfCode from CRF_Copy_Rot_Header"
                    slSQLQuery = slSQLQuery & " Where (crfCode = " & llCrfCode & ")"
                    Set crf_rst = gSQLSelectCall(slSQLQuery)
                    If Not crf_rst.EOF Then
                        llCrfCsfCode = crf_rst!crfCsfCode
                    End If
                    lgETime22 = timeGetTime
                    lgTtlTime22 = lgTtlTime22 + (lgETime22 - lgSTime22)
                End If
                llUpper = UBound(tgCifCpfInfo1)
                tgCifCpfInfo1(llUpper).cifCode = cif_rst!cifCode
                tgCifCpfInfo1(llUpper).cifAdfCode = cif_rst!cifAdfCode
                tgCifCpfInfo1(llUpper).cifCode = cif_rst!cifCode
                tgCifCpfInfo1(llUpper).cifCpfCode = cif_rst!cifCpfCode
                tgCifCpfInfo1(llUpper).cifMcfCode = cif_rst!cifMcfCode
                tgCifCpfInfo1(llUpper).cifName = cif_rst!cifName
                tgCifCpfInfo1(llUpper).cifReel = cif_rst!cifReel
                If Not IsNull(cif_rst!cifRotEndDate) Then
                    tgCifCpfInfo1(llUpper).cifRotEndDate = cif_rst!cifRotEndDate
                Else
                    tgCifCpfInfo1(llUpper).cifRotEndDate = ""
                End If
                tgCifCpfInfo1(llUpper).cpfCreative = slCreativeTitle
                tgCifCpfInfo1(llUpper).cpfISCI = slISCI
                tgCifCpfInfo1(llUpper).cpfName = slProduct
                llUpper = llUpper + 1
                ReDim Preserve tgCifCpfInfo1(0 To llUpper) As CIFCPFINFO1
                If UBound(tgCifCpfInfo1) - 1 >= 1 Then
                    ArraySortTyp fnAV(tgCifCpfInfo1(), 0), UBound(tgCifCpfInfo1), 0, LenB(tgCifCpfInfo1(0)), 0, -2, 0
                End If
                If igExportSource = 2 Then
                    DoEvents
                End If
                gGetCopy = True
                Exit Function
            End If
        End If
    End If
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-gGetCopy"
End Function

Private Function mIncludeSplitNetwork(slSplitNetwork As String, ilShttCode As Integer, llRafCode As Long, ilVefCode As Integer) As Integer
    Dim raf_rst As ADODB.Recordset
    Dim sht_rst As ADODB.Recordset
    Dim sef_rst As ADODB.Recordset
    Dim slCategory As String
    Dim slInclExcl As String
    Dim llVpf As Long
    Dim slState As String
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    mIncludeSplitNetwork = True
    
    If ((Asc(sgSpfUsingFeatures2) And SPLITNETWORKS) <> SPLITNETWORKS) Then
        Exit Function
    End If
    If (slSplitNetwork <> "P") And (slSplitNetwork <> "S") Then
        Exit Function
    End If
    If llRafCode <= 0 Then
        Exit Function
    End If
    
    llVpf = gBinarySearchVpf(CLng(ilVefCode))
    If llVpf <> -1 Then
        If igExportSource = 2 Then
            DoEvents
        End If
        If tgVpfOptions(llVpf).sAllowSplitCopy <> "Y" Then
            If slSplitNetwork <> "P" Then
                mIncludeSplitNetwork = False
            End If
            Exit Function
        End If
    End If
    
    If igExportSource = 2 Then
        DoEvents
    End If
    slSQLQuery = "Select rafCategory, rafInclExcl from  RAF_Region_Area"
    slSQLQuery = slSQLQuery & " Where (rafCode = " & llRafCode & ")"
    Set raf_rst = gSQLSelectCall(slSQLQuery)
    If raf_rst.EOF Then
        Exit Function
    End If
    slCategory = raf_rst!rafCategory
    slInclExcl = raf_rst!rafInclExcl
    
    slSQLQuery = "SELECT shttCode, shttState, shttStateLic, shttONState, shttZip, shttMktCode, shttFmtCode, shttOwnerArttCode FROM shtt WHERE (shttCode = " & ilShttCode & ")"
    Set sht_rst = gSQLSelectCall(slSQLQuery)

    slSQLQuery = "Select * from SEF_Split_Entity"
    slSQLQuery = slSQLQuery & " Where (sefRafCode = " & llRafCode & ")"
    slSQLQuery = slSQLQuery + " ORDER BY sefRafCode, sefSeqNo"
    Set sef_rst = gSQLSelectCall(slSQLQuery)
    While Not sef_rst.EOF
        If igExportSource = 2 Then
            DoEvents
        End If
        Select Case slCategory
            Case "M"    'Market
                If sef_rst!sefIntCode = sht_rst!shttMktCode Then
                    If slInclExcl = "E" Then
                        mIncludeSplitNetwork = False
                    End If
                    Exit Function
                End If
            Case "N"    'State Name
                '12/28/15
                If sgSplitState = "L" Then
                    slState = sht_rst!shttStateLic
                ElseIf sgSplitState = "P" Then
                    slState = sht_rst!shttONState
                Else
                    slState = sht_rst!shttState
                End If
                'If StrComp(Trim$(sef_rst!sefName), Trim$(sht_rst!shttState), vbTextCompare) = 0 Then
                If StrComp(Trim$(sef_rst!sefName), Trim$(slState), vbTextCompare) = 0 Then
                    If slInclExcl = "E" Then
                        mIncludeSplitNetwork = False
                    End If
                    Exit Function
                End If
            Case "Z"    'Zip Code
                If StrComp(Trim$(sef_rst!sefName), Trim$(sht_rst!shttZip), vbTextCompare) = 0 Then
                    If slInclExcl = "E" Then
                        mIncludeSplitNetwork = False
                    End If
                    Exit Function
                End If
            'Case "O"    'Owner
            '    If sef_rst!sefIntCode = sht_rst!shttOwnerArttCode Then
            '        If slInclExcl = "E" Then
            '            mIncludeSplitNetwork = False
            '        End If
            '        Exit Function
            '    End If
            Case "F"    'Format
                If sef_rst!sefIntCode = sht_rst!shttFmtCode Then
                    If slInclExcl = "E" Then
                        mIncludeSplitNetwork = False
                    End If
                    Exit Function
                End If
            Case "S"    'Station
                If sef_rst!sefIntCode = sht_rst!shttCode Then
                    If slInclExcl = "E" Then
                        mIncludeSplitNetwork = False
                    End If
                    Exit Function
                End If
        End Select
        sef_rst.MoveNext
    Wend
    If slInclExcl <> "E" Then
        mIncludeSplitNetwork = False
    End If
    
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-mIncludeSplitNetwork"
End Function

Public Function gCreateSplitFill(ilFillLen As Integer, llRafCode As Long, tlInLst As LST, tlOutLst As LST) As Integer
    'tlInLst(I)- Split Network spot to find fill for
    
    Dim ilBof As Integer
    Dim ilLoop As Integer
    Dim ilLast As Integer
    Dim ilDayOk As Integer
    Dim ilLenStart As Integer
    Dim ilStartBOF As Integer
    Dim slCartNo As String
    Dim slProduct As String
    Dim slISCI As String
    Dim slCreativeTitle As String
    Dim llCrfCsfCode As Long
    Dim llCpfCode As Long
    Dim ilRet As Integer
    Dim llLstCode As Long
    Dim ilCifAdfCode As Integer
    Dim ilLastAssign As Integer
    Dim slSQLQuery As String
    Dim lst_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If UBound(tgRBofRec) <= LBound(tgRBofRec) Then
        gCreateSplitFill = False
        Exit Function
    End If
    If UBound(tgSplitNetLastFill) <= LBound(tgSplitNetLastFill) Then
        gCreateSplitFill = False
        Exit Function
    End If
    ilLastAssign = -1
    For ilLoop = 0 To UBound(tgSplitNetLastFill) - 1 Step 1
        If igExportSource = 2 Then
            DoEvents
        End If
        If tgSplitNetLastFill(ilLoop).iFillLen = ilFillLen Then
            ilLastAssign = ilLoop
            Exit For
        End If
    Next ilLoop
    If ilLastAssign = -1 Then
        gCreateSplitFill = False
        Exit Function
    End If
    ilLoop = tgSplitNetLastFill(ilLastAssign).iBofIndex + 1
    If (ilLoop >= UBound(tgRBofRec)) Or (ilLoop < LBound(tgRBofRec)) Then
        ilLoop = LBound(tgRBofRec)
    End If
    ilStartBOF = ilLoop
    ilBof = -1
    Do
        If igExportSource = 2 Then
            DoEvents
        End If
        'If (tgRBofRec(ilLoop).iLen = tlInLst.iLen) And ((tgRBofRec(ilLoop).tBof.iVefCode = tlInLst.iLogVefCode) Or (tgRBofRec(ilLoop).tBof.iVefCode = 0)) Then
        If (tgRBofRec(ilLoop).iLen = ilFillLen) And ((tgRBofRec(ilLoop).tBof.iVefCode = tlInLst.iLogVefCode) Or (tgRBofRec(ilLoop).tBof.iVefCode = 0)) Then
            'Check Dates, Times and Days
            If (DateValue(gAdjYear(tlInLst.sLogDate)) >= DateValue(gAdjYear(tgRBofRec(ilLoop).tBof.sStartDate))) And (DateValue(gAdjYear(tlInLst.sLogDate)) <= DateValue(gAdjYear(tgRBofRec(ilLoop).tBof.sEndDate))) Then
                If (gTimeToLong(tlInLst.sLogTime, False) >= gTimeToLong(tgRBofRec(ilLoop).tBof.sStartTime, False)) And (gTimeToLong(tlInLst.sLogTime, False) <= gTimeToLong(tgRBofRec(ilLoop).tBof.sEndTime, True)) Then
                    ilDayOk = False
                    Select Case Weekday(tlInLst.sLogDate, vbMonday) - 1
                        Case 0  'Monday
                            If tgRBofRec(ilLoop).tBof.sMo = "Y" Then
                                ilDayOk = True
                            End If
                        Case 1
                            If tgRBofRec(ilLoop).tBof.sTu = "Y" Then
                                ilDayOk = True
                            End If
                        Case 2
                            If tgRBofRec(ilLoop).tBof.sWe = "Y" Then
                                ilDayOk = True
                            End If
                        Case 3
                            If tgRBofRec(ilLoop).tBof.sTh = "Y" Then
                                ilDayOk = True
                            End If
                        Case 4
                            If tgRBofRec(ilLoop).tBof.sFr = "Y" Then
                                ilDayOk = True
                            End If
                        Case 5
                            If tgRBofRec(ilLoop).tBof.sSa = "Y" Then
                                ilDayOk = True
                            End If
                        Case 6
                            If tgRBofRec(ilLoop).tBof.sSu = "Y" Then
                                ilDayOk = True
                            End If
                    End Select
                    If ilDayOk Then
                        tgSplitNetLastFill(ilLastAssign).iBofIndex = ilLoop
                        ilBof = ilLoop
                        Exit Do
                    End If
                End If
            End If
        End If
        ilLoop = ilLoop + 1
        If ilLoop >= UBound(tgRBofRec) Then
            ilLoop = LBound(tgRBofRec)
        End If
        If ilLoop = ilStartBOF Then
            Exit Do
        End If
    Loop
    If ilBof = -1 Then
        gCreateSplitFill = False
        Exit Function
    End If
    If igExportSource = 2 Then
        DoEvents
    End If
    'Create LST
    slSQLQuery = "SELECT *"
    slSQLQuery = slSQLQuery + " FROM CHF_Contract_Header"
    slSQLQuery = slSQLQuery + " WHERE (chfCode = " & tgRBofRec(ilBof).tBof.lRChfCode & ")"
    Set chf_rst = gSQLSelectCall(slSQLQuery)
    
    'Note in Traffic Log generation, room is left after split network spot for fill by incrementing PositionNo
    ilRet = gGetCopy("1", tgRBofRec(ilBof).tBof.lCifCode, 0, False, slCartNo, slProduct, slISCI, slCreativeTitle, llCrfCsfCode, llCpfCode, ilCifAdfCode, tlInLst.iLogVefCode)
    slSQLQuery = "INSERT INTO lst (lstType, lstSdfCode, lstCntrNo, "
    slSQLQuery = slSQLQuery & "lstAdfCode, lstAgfCode, lstProd, "
    slSQLQuery = slSQLQuery & "lstLineNo, lstLnVefCode, lstStartDate,"
    slSQLQuery = slSQLQuery & "lstEndDate, lstMon, lstTue, "
    slSQLQuery = slSQLQuery & "lstWed, lstThu, lstFri, "
    slSQLQuery = slSQLQuery & "lstSat, lstSun, lstSpotsWk, "
    slSQLQuery = slSQLQuery & "lstPriceType, lstPrice, lstSpotType, "
    slSQLQuery = slSQLQuery & "lstLogVefCode, lstLogDate, lstLogTime, "
    slSQLQuery = slSQLQuery & "lstDemo, lstAud, lstISCI, "
    slSQLQuery = slSQLQuery & "lstWkNo, lstBreakNo, lstPositionNo, "
    slSQLQuery = slSQLQuery & "lstSeqNo, lstZone, lstCart, "
    slSQLQuery = slSQLQuery & "lstCpfCode, lstCrfCsfCode, lstStatus, "
    slSQLQuery = slSQLQuery & "lstLen, lstUnits, lstCifCode, "
    slSQLQuery = slSQLQuery & "lstAnfCode, lstEvtIDCefCode, lstSplitNetwork, "
    ''slSQLQuery = slSQLQuery & "lstRafCode, lstUnused)"
    'slSQLQuery = slSQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, lstUnused)"
    slSQLQuery = slSQLQuery & "lstRafCode, lstFsfCode, lstGsfCode, lstImportedSpot, lstBkoutLstCode, "
    slSQLQuery = slSQLQuery & "lstLnStartTime, lstLnEndTime, lstUnused)"
    slSQLQuery = slSQLQuery & " VALUES (" & 0 & ", " & 0 & ", " & chf_rst!chfCntrNo & ", "
    slSQLQuery = slSQLQuery & chf_rst!chfAdfCode & ", " & chf_rst!chfAgfCode & ", '" & gFixQuote(chf_rst!chfProduct) & "', "
    slSQLQuery = slSQLQuery & 0 & ", " & 0 & ", '" & Format$(tgRBofRec(ilBof).tBof.sStartDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & "'" & Format$(tgRBofRec(ilBof).tBof.sEndDate, sgSQLDateForm) & "', "
    If tgRBofRec(ilLoop).tBof.sMo = "Y" Then
        slSQLQuery = slSQLQuery & 1 & ", "
    Else
        slSQLQuery = slSQLQuery & 0 & ", "
    End If
    If tgRBofRec(ilLoop).tBof.sTu = "Y" Then
        slSQLQuery = slSQLQuery & 1 & ", "
    Else
        slSQLQuery = slSQLQuery & 0 & ", "
    End If
    If tgRBofRec(ilLoop).tBof.sWe = "Y" Then
        slSQLQuery = slSQLQuery & 1 & ", "
    Else
        slSQLQuery = slSQLQuery & 0 & ", "
    End If
    If tgRBofRec(ilLoop).tBof.sTh = "Y" Then
        slSQLQuery = slSQLQuery & 1 & ", "
    Else
        slSQLQuery = slSQLQuery & 0 & ", "
    End If
    If tgRBofRec(ilLoop).tBof.sFr = "Y" Then
        slSQLQuery = slSQLQuery & 1 & ", "
    Else
        slSQLQuery = slSQLQuery & 0 & ", "
    End If
    If tgRBofRec(ilLoop).tBof.sSa = "Y" Then
        slSQLQuery = slSQLQuery & 1 & ", "
    Else
        slSQLQuery = slSQLQuery & 0 & ", "
    End If
    If tgRBofRec(ilLoop).tBof.sSu = "Y" Then
        slSQLQuery = slSQLQuery & 1 & ", "
    Else
        slSQLQuery = slSQLQuery & 0 & ", "
    End If
    slSQLQuery = slSQLQuery & 1 & ", "
    slSQLQuery = slSQLQuery & 1 & ", " & 0 & ", " & 2 & ", "
    slSQLQuery = slSQLQuery & tlInLst.iLogVefCode & ", '" & Format$(tlInLst.sLogDate, sgSQLDateForm) & "', '" & Format$(tlInLst.sLogTime, sgSQLTimeForm) & "', "
    slSQLQuery = slSQLQuery & "'" & "" & "', " & 0 & ", '" & gFixQuote(slISCI) & "', "
    slSQLQuery = slSQLQuery & tlInLst.iWkNo & ", " & tlInLst.iBreakNo & ", " & tlInLst.iPositionNo + 1 & ", "
    slSQLQuery = slSQLQuery & tlInLst.iSeqNo & ", '" & tlInLst.sZone & "', '" & gFixQuote(slCartNo) & "', "
    slSQLQuery = slSQLQuery & llCpfCode & ", " & llCrfCsfCode & ", " & 0 & ", "
    slSQLQuery = slSQLQuery & ilFillLen & ", " & 0 & ", " & tgRBofRec(ilBof).tBof.lCifCode & ", "
    slSQLQuery = slSQLQuery & tlInLst.iAnfCode & ", " & tlInLst.lEvtIDCefCode & ", 'F', "
    '12/28/06
    ''slSQLQuery = slSQLQuery & 0 & ", '" & "" & "'" & ")"
    'slSQLQuery = slSQLQuery & 0 & ", " & 0 & ", " & 0 & ", '" & "N" & "', " & 0 & ", '" & "" & "'" & ")"
    slSQLQuery = slSQLQuery & 0 & ", " & 0 & ", " & tlInLst.lgsfCode & ", '" & "N" & "', " & 0 & ", "
    slSQLQuery = slSQLQuery & "'" & Format("12am", sgSQLTimeForm) & "', '" & Format("12am", sgSQLTimeForm) & "', '" & "" & "'" & ")"
    cnn.BeginTrans
    'cnn.Execute slSQLQuery, rdExecDirect
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        '6/13/16: Replaced GoSub
        'GoSub ErrHand:
        Screen.MousePointer = vbDefault
        gHandleError "AffErrorLog.txt", "modCPReturns-gCreateSplitFill"
        cnn.RollbackTrans
        gCreateSplitFill = False
        Exit Function
    End If
    cnn.CommitTrans
    If igExportSource = 2 Then
        DoEvents
    End If
    slSQLQuery = "Select MAX(lstCode) from lst"
    Set lst_rst = gSQLSelectCall(slSQLQuery)
    llLstCode = lst_rst(0).Value
    slSQLQuery = "Select * from lst WHERE lstCode = " & llLstCode
    Set lst_rst = gSQLSelectCall(slSQLQuery)
    gCreateUDTforLST lst_rst, tlOutLst
    If igExportSource = 2 Then
        DoEvents
    End If
    gCreateSplitFill = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-gCreateSplitFill"
End Function

Public Function gGetSdfCrfCode(llSdfCode As Long) As Long
    Dim slType As String
    Dim slSQLQuery As String
    'Dim sdf_rst As ADODB.Recordset
    Dim crf_rst As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    gGetSdfCrfCode = 0
    slSQLQuery = "Select * from SDF_Spot_Detail"
    slSQLQuery = slSQLQuery & " Where (sdfCode = " & llSdfCode & ")"
    Set sdf_rst = gSQLSelectCall(slSQLQuery)
    If Not sdf_rst.EOF Then
        If igExportSource = 2 Then
            DoEvents
        End If
        If sdf_rst!sdfSpotType = "C" Then
            slType = "C"
        ElseIf sdf_rst!sdfSpotType = "O" Then
            slType = "O"
        Else
            slType = "A"
        End If
        slSQLQuery = "Select crfCode, crfAdfCode, crfChfCode, crfRotNo from CRF_Copy_Rot_Header"
        slSQLQuery = slSQLQuery & " Where (crfRotType = '" & slType & "' AND "
        slSQLQuery = slSQLQuery & " crfEtfCode = 0 AND crfEnfCode = 0 AND "
        slSQLQuery = slSQLQuery & " crfAdfCode = " & sdf_rst!sdfAdfCode & " AND"
        slSQLQuery = slSQLQuery & " crfChfCode = " & sdf_rst!sdfChfCode & " AND"
        slSQLQuery = slSQLQuery & " crfFsfCode = " & sdf_rst!sdfFsfCode & " AND"
        'Rotation number is independent of vehicle
        'slSQLQuery = slSQLQuery & " crfVefCode = " & sdf_rst!sdfVefCode & " AND"
        slSQLQuery = slSQLQuery & " crfRotNo = " & sdf_rst!sdfRotNo & ")"
        Set crf_rst = gSQLSelectCall(slSQLQuery)
        If Not crf_rst.EOF Then
            If (crf_rst!crfAdfCode = sdf_rst!sdfAdfCode) And (crf_rst!crfChfCode = sdf_rst!sdfChfCode) And (crf_rst!crfRotNo = sdf_rst!sdfRotNo) Then
                gGetSdfCrfCode = crf_rst!crfCode
            End If
        End If
        If igExportSource = 2 Then
            DoEvents
        End If
        crf_rst.Close
    End If
    sdf_rst.Close
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-gGetSdfCrfCode"
End Function

Public Function gBuildRegionDefinitions(slTarget As String, llSdfCode As Long, ilVefCode As Integer, tlRegionDefinition() As REGIONDEFINITION, tlSplitCategoryInfo() As SPLITCATEGORYINFO) As Integer
    'slTarget(I):  C=Copy; W=Wegener; O=OLA.  For Wegener and OLA don't include excludes stations if the station is not part of the wegener or OLA
    Dim slCategory As String
    Dim slInclExcl As String
    Dim llVpf As Long
    Dim llVef As Long
    Dim llVefA As Long
    Dim llPreviousOther As Long
    Dim llPreviousFormat As Long
    Dim llPreviousExclude As Long
    Dim llRegionDefinitionIndex As Long
    Dim ilShtt As Integer
    Dim ilAddExclude As Integer
    Dim llUpperReg As Long
    Dim llUpperSplit As Long
    Dim ilAirRotNo As Integer
    '5/21/15: add check if generic split copy assigned to airing vehicle
    Dim blBypassRsf As Boolean
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    
    gBuildRegionDefinitions = False
    ReDim tlRegionDefinition(0 To 0) As REGIONDEFINITION
    ReDim tlSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO

    If ((Asc(sgSpfUsingFeatures2) And SPLITCOPY) <> SPLITCOPY) Then
        Exit Function
    End If

    llVef = gBinarySearchVef(CLng(ilVefCode))
    If llVef <> -1 Then
        If igExportSource = 2 Then
            DoEvents
        End If
        '11/4/09-  Show Log and Conventional vehicle.  Let client pick which they want agreements to be used for
        'Temporarily include only for Special user until testing is complete
        'If (tgVehicleInfo(llVef).sVehType = "C") Or (tgVehicleInfo(llVef).sVehType = "A") Or (tgVehicleInfo(llVef).sVehType = "G") Then
        If (tgVehicleInfo(llVef).sVehType = "C") Or (tgVehicleInfo(llVef).sVehType = "A") Or (tgVehicleInfo(llVef).sVehType = "G") Or (tgVehicleInfo(llVef).sVehType = "L") Then
            llVpf = gBinarySearchVpf(CLng(ilVefCode))
            If llVpf <> -1 Then
                If tgVpfOptions(llVpf).sAllowSplitCopy <> "Y" Then
                    Exit Function
                End If
            Else
                Exit Function
            End If
        ElseIf (tgVehicleInfo(llVef).sVehType <> "S") And (tgVehicleInfo(llVef).sVehType <> "P") Then
            Exit Function
        End If
    Else
        Exit Function
    End If
    If igExportSource = 2 Then
        DoEvents
    End If
    
    '4/9/15: Obtain rotation number of copy assigned by airing vehicle
    slSQLQuery = "Select rsfRotNo from RSF_Region_Schd_Copy"
    slSQLQuery = slSQLQuery & " Where (rsfSdfCode = " & llSdfCode
    slSQLQuery = slSQLQuery & " AND rsfType = 'A'"     'Airing vehicle copy
    slSQLQuery = slSQLQuery & " AND rsfBVefCode = " & ilVefCode & ")" & " Order By rsfRotNo DESC"
    Set rsf_rst = gSQLSelectCall(slSQLQuery)
    If Not rsf_rst.EOF Then
        ilAirRotNo = rsf_rst!rsfRotNo
    Else
        ilAirRotNo = 0
    End If


    '5/21/15: add check if generic split copy assigned to airing vehicle
    'slSQLQuery = "Select rsfCode, rsfRotNo, rstPtType, rsfCopyCode, rsfCrfCode, rafCode, rafCategory, rafInclExcl, rafName from RSF_Region_Schd_Copy, RAF_Region_Area"
    slSQLQuery = "Select rsfCode, rsfRotNo, rstPtType, rsfCopyCode, rsfCrfCode, rsfBVefCode, rafCode, rafCategory, rafInclExcl, rafName from RSF_Region_Schd_Copy, RAF_Region_Area"
    slSQLQuery = slSQLQuery & " Where (rsfSdfCode = " & llSdfCode
    '4/9/15: Only obtain the rotation that are assigned with rotation numbers greater than the rotation number assigned by airing vehicle
    slSQLQuery = slSQLQuery & " AND rsfRotNo > " & ilAirRotNo
    slSQLQuery = slSQLQuery & " AND rsfType <> 'B'"     'Blackout
    slSQLQuery = slSQLQuery & " AND rsfType <> 'A'"     'Airing vehicle copy
    slSQLQuery = slSQLQuery & " AND rafType = 'C'"     'Split copy
    slSQLQuery = slSQLQuery & " AND rafCode = rsfRafCode" & ")" '& " Order By rsfRotNo DESC"
    Set rsf_rst = gSQLSelectCall(slSQLQuery)
    If rsf_rst.EOF Then
        Exit Function
    End If
    ReDim tlRegionDefinition(0 To 50) As REGIONDEFINITION
    llUpperReg = 0
    Do While Not rsf_rst.EOF
        If igExportSource = 2 Then
            DoEvents
        End If
        '5/21/15: add check if generic split copy assigned to airing vehicle
        blBypassRsf = False
        If rsf_rst!rsfBVefCode <> ilVefCode Then
            If llVef <> -1 Then
                If tgVehicleInfo(llVef).sVehType = "A" Then
                    '6/10/15: The above test was not coded correctly as it should only bypass
                    '         this copy if assigned to airing vehicle (copy assigned to two airing vehicle that don't match)
                    'blBypassRsf = True
                    llVefA = gBinarySearchVef(CLng(rsf_rst!rsfBVefCode))
                    If llVefA <> -1 Then
                        If tgVehicleInfo(llVefA).sVehType = "A" Then
                            blBypassRsf = True
                        End If
                    End If
                End If
            End If
        End If
        If Not blBypassRsf Then
            tlRegionDefinition(llUpperReg).lRotNo = rsf_rst!rsfRotNo
            tlRegionDefinition(llUpperReg).lRafCode = rsf_rst!rafCode
            tlRegionDefinition(llUpperReg).sCategory = Trim$(rsf_rst!rafCategory)
            tlRegionDefinition(llUpperReg).sInclExcl = rsf_rst!rafInclExcl
            tlRegionDefinition(llUpperReg).sRegionName = rsf_rst!rafName
            tlRegionDefinition(llUpperReg).lFormatFirst = -1
            tlRegionDefinition(llUpperReg).lOtherFirst = -1
            tlRegionDefinition(llUpperReg).lExcludeFirst = -1
            tlRegionDefinition(llUpperReg).sPtType = rsf_rst!rstPtType
            tlRegionDefinition(llUpperReg).lCopyCode = rsf_rst!rsfCopyCode
            tlRegionDefinition(llUpperReg).lCrfCode = rsf_rst!rsfCrfCode
            tlRegionDefinition(llUpperReg).lRsfCode = rsf_rst!rsfCode
            tlRegionDefinition(llUpperReg).iStationCount = -1
            tlRegionDefinition(llUpperReg).lStationOtherFirst = -1
            tlRegionDefinition(llUpperReg).iPoolNextFinal = -1
            tlRegionDefinition(llUpperReg).iPoolAdfCode = -1
            tlRegionDefinition(llUpperReg).lPoolCrfCode = -1
            tlRegionDefinition(llUpperReg).bPoolUpdated = False
            llUpperReg = llUpperReg + 1
            If llUpperReg >= UBound(tlRegionDefinition) Then
                ReDim Preserve tlRegionDefinition(0 To UBound(tlRegionDefinition) + 10) As REGIONDEFINITION
            End If
            '7/5/19
            mBuildPoolRegionDefinitions llSdfCode, ilVefCode, llUpperReg, tlRegionDefinition()
        End If
        rsf_rst.MoveNext
    Loop
    ReDim Preserve tlRegionDefinition(0 To llUpperReg) As REGIONDEFINITION
    '2/14/13: Wegener needs to be sorted in ascending order because of how it is assigned
    If slTarget <> "W" Then
        'Descending sort
        ArraySortTyp fnAV(tlRegionDefinition(), 0), UBound(tlRegionDefinition), 1, LenB(tlRegionDefinition(0)), 0, -2, 0
    Else
        'Ascending sort
        ArraySortTyp fnAV(tlRegionDefinition(), 0), UBound(tlRegionDefinition), 0, LenB(tlRegionDefinition(0)), 0, -2, 0
    End If
    
    If igExportSource = 2 Then
        DoEvents
    End If
    If slTarget = "C" Then
        gBuildRegionDefinitions = True
        Exit Function
    End If
    ReDim tlSplitCategoryInfo(0 To 500) As SPLITCATEGORYINFO
    llUpperSplit = 0
    For llRegionDefinitionIndex = 0 To UBound(tlRegionDefinition) - 1 Step 1
        mBuildSplitCategoryInfo slTarget, llUpperSplit, tlRegionDefinition(llRegionDefinitionIndex), tlSplitCategoryInfo()
    Next llRegionDefinitionIndex
    ReDim Preserve tlSplitCategoryInfo(0 To llUpperSplit) As SPLITCATEGORYINFO
    'sef_rst.Close
    'rsf_rst.Close
    gBuildRegionDefinitions = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-gBuildRegionDefinitions"
    ReDim Preserve tlRegionDefinition(0 To llUpperReg) As REGIONDEFINITION
    ReDim Preserve tlSplitCategoryInfo(0 To llUpperSplit) As SPLITCATEGORYINFO
End Function

Public Function gBuildSplitNetRegionDefinitions(ilFirstLst As Integer, tlSplitNetRegion() As SPLITNETREGION, tlRegionDefinition() As REGIONDEFINITION, tlSplitCategoryInfo() As SPLITCATEGORYINFO) As Integer
    Dim slCategory As String
    Dim slInclExcl As String
    Dim llPreviousOther As Long
    Dim llPreviousFormat As Long
    Dim llPreviousExclude As Long
    Dim llRegionDefinitionIndex As Long
    Dim ilShtt As Integer
    Dim ilAddExclude As Integer
    Dim llUpperReg As Long
    Dim llUpperSplit As Long
    Dim ilLst As Integer
    Dim slSQLQuery As String

    On Error GoTo ErrHand
    
    gBuildSplitNetRegionDefinitions = False
    ReDim tlRegionDefinition(0 To 0) As REGIONDEFINITION
    ReDim tlSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
    ilLst = ilFirstLst
    llUpperReg = 0
    Do While ilLst <> -1
        If igExportSource = 2 Then
            DoEvents
        End If
        slSQLQuery = "Select sdfCode, sdfRotNo, sdfPointer, sdfCopy, rafCode, rafCategory, rafInclExcl, rafName from SDF_Spot_Detail, RAF_Region_Area, Lst"
        slSQLQuery = slSQLQuery & " Where (lstCode = " & tlSplitNetRegion(ilLst).lLstCode
        slSQLQuery = slSQLQuery & " AND sdfCode = lstSdfCode"
        slSQLQuery = slSQLQuery & " AND rafCode = lstrafCode )"
        Set lst_rst = gSQLSelectCall(slSQLQuery)
        If lst_rst.EOF Then
            Exit Function
        End If
        If Not lst_rst.EOF Then
            If igExportSource = 2 Then
                DoEvents
            End If
            tlRegionDefinition(llUpperReg).lRotNo = lst_rst!sdfRotNo
            tlRegionDefinition(llUpperReg).lRafCode = lst_rst!rafCode
            tlRegionDefinition(llUpperReg).sCategory = Trim$(lst_rst!rafCategory)
            tlRegionDefinition(llUpperReg).sInclExcl = lst_rst!rafInclExcl
            tlRegionDefinition(llUpperReg).sRegionName = lst_rst!rafName
            tlRegionDefinition(llUpperReg).lFormatFirst = -1
            tlRegionDefinition(llUpperReg).lOtherFirst = -1
            tlRegionDefinition(llUpperReg).lExcludeFirst = -1
            tlRegionDefinition(llUpperReg).sPtType = lst_rst!sdfPointer
            tlRegionDefinition(llUpperReg).lCopyCode = lst_rst!sdfCopy
            tlRegionDefinition(llUpperReg).lCrfCode = gGetSdfCrfCode(lst_rst!sdfcode)
            tlRegionDefinition(llUpperReg).lRsfCode = 0
            tlRegionDefinition(llUpperReg).iStationCount = -1
            tlRegionDefinition(llUpperReg).lStationOtherFirst = -1
            tlRegionDefinition(llUpperReg).iPoolNextFinal = -1
            tlRegionDefinition(llUpperReg).iPoolAdfCode = -1
            tlRegionDefinition(llUpperReg).lPoolCrfCode = -1
            tlRegionDefinition(llUpperReg).bPoolUpdated = False
            llUpperReg = llUpperReg + 1
            If llUpperReg >= UBound(tlRegionDefinition) Then
                ReDim Preserve tlRegionDefinition(0 To UBound(tlRegionDefinition) + 10) As REGIONDEFINITION
            End If
        End If
        ilLst = tlSplitNetRegion(ilLst).iNext
    Loop
    ReDim Preserve tlRegionDefinition(0 To llUpperReg) As REGIONDEFINITION
    'Descending sort
    ArraySortTyp fnAV(tlRegionDefinition(), 0), UBound(tlRegionDefinition), 1, LenB(tlRegionDefinition(0)), 0, -2, 0
    
    If igExportSource = 2 Then
        DoEvents
    End If
    ReDim tlSplitCategoryInfo(0 To 500) As SPLITCATEGORYINFO
    llUpperSplit = 0
    For llRegionDefinitionIndex = 0 To UBound(tlRegionDefinition) - 1 Step 1
        mBuildSplitCategoryInfo "W", llUpperSplit, tlRegionDefinition(llRegionDefinitionIndex), tlSplitCategoryInfo()
    Next llRegionDefinitionIndex
    ReDim Preserve tlSplitCategoryInfo(0 To llUpperSplit) As SPLITCATEGORYINFO
    'sef_rst.Close
    'rsf_rst.Close
    gBuildSplitNetRegionDefinitions = True
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-gBuildSplitNetRegionDefinitions"
    ReDim Preserve tlRegionDefinition(0 To llUpperReg) As REGIONDEFINITION
    ReDim Preserve tlSplitCategoryInfo(0 To llUpperSplit) As SPLITCATEGORYINFO
End Function
Private Function mTestCategorybyStation(slInclExcl As String, ilMktCode As Integer, ilMSAMktCode As Integer, slState As String, ilFmtCode As Integer, ilTztCode As Integer, ilShttCode As Integer, tlSplitCategoryInfo As SPLITCATEGORYINFO, slTypeCode As String) As Integer
    If igExportSource = 2 Then
        DoEvents
    End If
    slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory
    Select Case tlSplitCategoryInfo.sCategory
        Case "M"    'DMA Market
            slTypeCode = slTypeCode & Trim$(Str$(ilMktCode))
            If tlSplitCategoryInfo.iIntCode = ilMktCode Then
                If slInclExcl <> "E" Then
                    mTestCategorybyStation = True
                Else
                    mTestCategorybyStation = False
                End If
                Exit Function
            End If
        Case "A"    'MSA Market
            slTypeCode = slTypeCode & Trim$(Str$(ilMSAMktCode))
            If tlSplitCategoryInfo.iIntCode = ilMSAMktCode Then
                If slInclExcl <> "E" Then
                    mTestCategorybyStation = True
                Else
                    mTestCategorybyStation = False
                End If
                Exit Function
            End If
        Case "N"    'State Name
            slTypeCode = slTypeCode & Trim$(slState)
            If StrComp(Trim$(tlSplitCategoryInfo.sName), Trim$(slState), vbTextCompare) = 0 Then
                If slInclExcl <> "E" Then
                    mTestCategorybyStation = True
                Else
                    mTestCategorybyStation = False
                End If
                Exit Function
            End If
        Case "F"    'Format
            slTypeCode = slTypeCode & Trim$(Str$(ilFmtCode))
            If tlSplitCategoryInfo.iIntCode = ilFmtCode Then
                If slInclExcl <> "E" Then
                    mTestCategorybyStation = True
                Else
                    mTestCategorybyStation = False
                End If
                Exit Function
            End If
        Case "T"    'Time zone
            slTypeCode = slTypeCode & Trim$(Str$(ilTztCode))
            If tlSplitCategoryInfo.iIntCode = ilTztCode Then
                If slInclExcl <> "E" Then
                    mTestCategorybyStation = True
                Else
                    mTestCategorybyStation = False
                End If
                Exit Function
            End If
        Case "S"    'Station
            slTypeCode = slTypeCode & Trim$(Str$(ilShttCode))
            If tlSplitCategoryInfo.iIntCode = ilShttCode Then
                If slInclExcl <> "E" Then
                    mTestCategorybyStation = True
                Else
                    mTestCategorybyStation = False
                End If
                Exit Function
            End If
    End Select
    If igExportSource = 2 Then
        DoEvents
    End If
    If slInclExcl <> "E" Then
        mTestCategorybyStation = False
    Else
        mTestCategorybyStation = True
        Select Case tlSplitCategoryInfo.sCategory
            Case "M"    'DMA Market
                slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory & Trim$(Str$(tlSplitCategoryInfo.iIntCode))
            Case "A"    'MSA arket
                slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory & Trim$(Str$(tlSplitCategoryInfo.iIntCode))
            Case "N"    'State Name
                slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory & Trim$(tlSplitCategoryInfo.sName)
            Case "F"    'Format
                slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory & Trim$(Str$(tlSplitCategoryInfo.iIntCode))
            Case "T"    'Time zone
                slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory & Trim$(Str$(tlSplitCategoryInfo.iIntCode))
            Case "S"    'Station
                slTypeCode = slInclExcl & tlSplitCategoryInfo.sCategory & Trim$(Str$(tlSplitCategoryInfo.iIntCode))
        End Select
    End If
End Function

Public Function gRegionTestDefinition(ilShttCode As Integer, ilMktCode As Integer, ilMSAMktCode As Integer, slState As String, ilFmtCode As Integer, ilTztCode As Integer, tlRegionDefinition() As REGIONDEFINITION, tlSplitCategoryInfo() As SPLITCATEGORYINFO, llRegionIndex As Long, slGroupInfo As String) As Integer
    Dim ilRet As Integer
    Dim llFormatIndex As Long
    Dim llOtherIndex As Long
    Dim llExcludeIndex As Long
    Dim ilExcludeOk As Integer
    Dim llRegion As Long
    Dim slTotalTypeCode As String
    Dim slTypeCode As String
    Dim ilExitDo As Integer
    
    On Error GoTo ErrHand
    
    gRegionTestDefinition = False
    For llRegion = 0 To UBound(tlRegionDefinition) - 1 Step 1
        If igExportSource = 2 Then
            DoEvents
        End If
        slTotalTypeCode = ""
        llFormatIndex = tlRegionDefinition(llRegion).lFormatFirst
            
        If tlRegionDefinition(llRegion).lFormatFirst <> -1 Then
            'Test Format
            ilExitDo = False
            llFormatIndex = tlRegionDefinition(llRegion).lFormatFirst
            Do
                If igExportSource = 2 Then
                    DoEvents
                End If
                ilRet = mTestCategorybyStation("I", ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, ilShttCode, tlSplitCategoryInfo(llFormatIndex), slTypeCode)
                If ilRet Then
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    If slTotalTypeCode = "" Then
                        slTotalTypeCode = slTypeCode
                    Else
                        slTotalTypeCode = slTotalTypeCode & "|" & slTypeCode
                    End If
                    If tlRegionDefinition(llRegion).lOtherFirst <> -1 Then
                        'Test Other
                        llOtherIndex = tlRegionDefinition(llRegion).lOtherFirst
                        Do
                            ilRet = mTestCategorybyStation("I", ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, ilShttCode, tlSplitCategoryInfo(llOtherIndex), slTypeCode)
                            If ilRet Then
                                slTotalTypeCode = slTotalTypeCode & "|" & slTypeCode
                            Else
                                ilExitDo = True
                                Exit Do
                            End If
                            llOtherIndex = tlSplitCategoryInfo(llOtherIndex).lNext
                        Loop While llOtherIndex <> -1
                    End If
                    'Exclude
                    If Not ilExitDo Then
                        ilExcludeOk = True
                        llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
                        Do While llExcludeIndex <> -1
                            If igExportSource = 2 Then
                                DoEvents
                            End If
                            ilRet = mTestCategorybyStation("E", ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, ilShttCode, tlSplitCategoryInfo(llExcludeIndex), slTypeCode)
                            If Not ilRet Then
                                ilExcludeOk = False
                                Exit Do
                            Else
                                slTotalTypeCode = slTotalTypeCode & "|" & slTypeCode
                            End If
                            llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                        Loop
                        If ilExcludeOk Then
                            'ilRet = gGetCopy(tlRegionDefinition(llRegion).sPtType, tlRegionDefinition(llRegion).lCopyCode, tlRegionDefinition(llRegion).lCrfCode, slCartNo, slProduct, slISCI, slCreativeTitle, llCrfCsfCode, llCpfCode)
                            'gRegionTestDefinition = ilRet
                            llRegionIndex = llRegion
                            slGroupInfo = slTotalTypeCode
                            gRegionTestDefinition = True
                            Exit Function
                        Else
                            ilExitDo = True
                        End If
                    End If
                Else
                    ilExitDo = True
                End If
                If ilExitDo Then
                    Exit Do
                End If
                llFormatIndex = tlSplitCategoryInfo(llFormatIndex).lNext
                'Can't have two formats connected
                If llFormatIndex <> -1 Then
                    Exit Do
                End If
            Loop While llFormatIndex <> -1
        ElseIf tlRegionDefinition(llRegion).lOtherFirst <> -1 Then
            ilExitDo = False
            llOtherIndex = tlRegionDefinition(llRegion).lOtherFirst
            Do
                If igExportSource = 2 Then
                    DoEvents
                End If
                ilRet = mTestCategorybyStation("I", ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, ilShttCode, tlSplitCategoryInfo(llOtherIndex), slTypeCode)
                If ilRet Then
                    If slTotalTypeCode = "" Then
                        slTotalTypeCode = slTypeCode
                    Else
                        slTotalTypeCode = slTotalTypeCode & "|" & slTypeCode
                    End If
                Else
                    ilExitDo = True
                    Exit Do
                End If
                llOtherIndex = tlSplitCategoryInfo(llOtherIndex).lNext
            Loop While llOtherIndex <> -1
            If Not ilExitDo Then
                ilExcludeOk = True
                llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
                Do While llExcludeIndex <> -1
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    ilRet = mTestCategorybyStation("E", ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, ilShttCode, tlSplitCategoryInfo(llExcludeIndex), slTypeCode)
                    If Not ilRet Then
                        ilExcludeOk = False
                        Exit Do
                    Else
                        slTotalTypeCode = slTotalTypeCode & "|" & slTypeCode
                    End If
                    llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
                Loop
                If ilExcludeOk Then
                    'ilRet = gGetCopy(tlRegionDefinition(llRegion).sPtType, tlRegionDefinition(llRegion).lCopyCode, tlRegionDefinition(llRegion).lCrfCode, slCartNo, slProduct, slISCI, slCreativeTitle, llCrfCsfCode, llCpfCode)
                    'gRegionTestDefinition = ilRet
                    llRegionIndex = llRegion
                    slGroupInfo = slTotalTypeCode
                    gRegionTestDefinition = True
                    Exit Function
                End If
            End If
        ElseIf tlRegionDefinition(llRegion).lExcludeFirst <> -1 Then
            'Exclude only
            ilExcludeOk = True
            llExcludeIndex = tlRegionDefinition(llRegion).lExcludeFirst
            Do While llExcludeIndex <> -1
                If igExportSource = 2 Then
                    DoEvents
                End If
                ilRet = mTestCategorybyStation("E", ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, ilShttCode, tlSplitCategoryInfo(llExcludeIndex), slTypeCode)
                If Not ilRet Then
                    ilExcludeOk = False
                    Exit Do
                Else
                    If slTotalTypeCode = "" Then
                        slTotalTypeCode = slTypeCode
                    Else
                        slTotalTypeCode = slTotalTypeCode & "|" & slTypeCode
                    End If
                End If
                llExcludeIndex = tlSplitCategoryInfo(llExcludeIndex).lNext
            Loop
            If ilExcludeOk Then
                'ilRet = gGetCopy(tlRegionDefinition(llRegion).sPtType, tlRegionDefinition(llRegion).lCopyCode, tlRegionDefinition(llRegion).lCrfCode, slCartNo, slProduct, slISCI, slCreativeTitle, llCrfCsfCode, llCpfCode)
                'gRegionTestDefinition = ilRet
                llRegionIndex = llRegion
                slGroupInfo = slTotalTypeCode
                gRegionTestDefinition = True
                Exit Function
            End If
        End If
    Next llRegion
    If igExportSource = 2 Then
        DoEvents
    End If
    slGroupInfo = ""
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-gRegionTestDefinition"
End Function
Public Sub gSeparateRegions(tlInRegionDefinition() As REGIONDEFINITION, tlInSplitCategoryInfo() As SPLITCATEGORYINFO, tlOutRegionDefinition() As REGIONDEFINITION, tlOutSplitCategoryInfo() As SPLITCATEGORYINFO, Optional blSplitAllStationDefinition As Boolean = True)
    'If a region is defined as:
    '(Fmt1 or Fmt2 or Fmt3) and (St1 or St2) and (Not K1111 and Not K222)
    'Convert to:
    'Region 1: Fmt1 and St1 And Not K111 and Not K222
    'Region 2: Fmt1 and St2 And Not K111 and Not K222
    'Region 3: Fmt2 and St1 And Not K111 and Not K222
    'Region 4: Fmt2 and St2 And Not K111 and Not K222
    'Region 5: Fmt3 and St1 And Not K111 and Not K222
    'Region 6: Fmt3 and St2 And Not K111 and Not K222
    Dim llFormatIndex As Long
    Dim llRegion As Long
    Dim llOtherIndex As Long
    Dim llExcludeIndex As Long
    '3/3/18
    Dim ilStationCount As Integer
    
    For llRegion = 0 To UBound(tlInRegionDefinition) - 1 Step 1
        If igExportSource = 2 Then
            DoEvents
        End If
        llFormatIndex = tlInRegionDefinition(llRegion).lFormatFirst
            
        If tlInRegionDefinition(llRegion).lFormatFirst <> -1 Then
            'Test Format
            llFormatIndex = tlInRegionDefinition(llRegion).lFormatFirst
            Do
                If igExportSource = 2 Then
                    DoEvents
                End If
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)) = tlInRegionDefinition(llRegion)
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lFormatFirst = UBound(tlOutSplitCategoryInfo)
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lOtherFirst = -1
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = -1
                 '3/3/18
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).iStationCount = -1
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lStationOtherFirst = -1
                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llFormatIndex)
                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                If tlInRegionDefinition(llRegion).lOtherFirst <> -1 Then
                    llOtherIndex = tlInRegionDefinition(llRegion).lOtherFirst
                    Do
                        If igExportSource = 2 Then
                            DoEvents
                        End If
                        tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lOtherFirst = UBound(tlOutSplitCategoryInfo)
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llOtherIndex)
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                        ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        llExcludeIndex = tlInRegionDefinition(llRegion).lExcludeFirst
                        If llExcludeIndex <> -1 Then
                            tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = UBound(tlOutSplitCategoryInfo)
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                            ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                            llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                            Do While llExcludeIndex <> -1
                                If igExportSource = 2 Then
                                    DoEvents
                                End If
                                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo) - 1).lNext = UBound(tlOutSplitCategoryInfo)
                                ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                                llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                            Loop
                        End If
                        ReDim Preserve tlOutRegionDefinition(0 To UBound(tlOutRegionDefinition) + 1) As REGIONDEFINITION
                        llOtherIndex = tlInSplitCategoryInfo(llOtherIndex).lNext
                        If llOtherIndex <> -1 Then
                            tlOutRegionDefinition(UBound(tlOutRegionDefinition)) = tlInRegionDefinition(llRegion)
                            tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lFormatFirst = UBound(tlOutSplitCategoryInfo)
                            tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lOtherFirst = -1
                            tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = -1
                            '3/3/18
                            tlOutRegionDefinition(UBound(tlOutRegionDefinition)).iStationCount = -1
                            tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lStationOtherFirst = -1
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llFormatIndex)
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                            ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        End If
                    Loop While llOtherIndex <> -1
                ElseIf tlInRegionDefinition(llRegion).lExcludeFirst <> -1 Then
                    llExcludeIndex = tlInRegionDefinition(llRegion).lExcludeFirst
                    If llExcludeIndex <> -1 Then
                        tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = UBound(tlOutSplitCategoryInfo)
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                        ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                        Do While llExcludeIndex <> -1
                            If igExportSource = 2 Then
                                DoEvents
                            End If
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo) - 1).lNext = UBound(tlOutSplitCategoryInfo)
                            ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                            llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                        Loop
                        ReDim Preserve tlOutRegionDefinition(0 To UBound(tlOutRegionDefinition) + 1) As REGIONDEFINITION
                    End If
                Else
                    ReDim Preserve tlOutRegionDefinition(0 To UBound(tlOutRegionDefinition) + 1) As REGIONDEFINITION
                End If
                llFormatIndex = tlInSplitCategoryInfo(llFormatIndex).lNext
            Loop While llFormatIndex <> -1
        ElseIf tlInRegionDefinition(llRegion).lOtherFirst <> -1 Then
            '3/3/18: test if only stations.  If so don't expand
            ilStationCount = 0
            If tlInRegionDefinition(llRegion).lExcludeFirst = -1 Then
                If Not blSplitAllStationDefinition Then
                    llOtherIndex = tlInRegionDefinition(llRegion).lOtherFirst
                    Do
                        If tlInSplitCategoryInfo(llOtherIndex).sCategory <> "S" Then
                            ilStationCount = -1
                            Exit Do
                        End If
                        ilStationCount = ilStationCount + 1
                        llOtherIndex = tlInSplitCategoryInfo(llOtherIndex).lNext
                    Loop While llOtherIndex <> -1
                End If
            End If            '3/3/18
            If ilStationCount <= 0 Then
                llOtherIndex = tlInRegionDefinition(llRegion).lOtherFirst
                Do
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    tlOutRegionDefinition(UBound(tlOutRegionDefinition)) = tlInRegionDefinition(llRegion)
                    tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lFormatFirst = -1
                    tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lOtherFirst = UBound(tlOutSplitCategoryInfo)
                    tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = -1
                    '3/3/18
                    tlOutRegionDefinition(UBound(tlOutRegionDefinition)).iStationCount = -1
                    tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lStationOtherFirst = -1
                    tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llOtherIndex)
                    tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                    ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                    llExcludeIndex = tlInRegionDefinition(llRegion).lExcludeFirst
                    If llExcludeIndex <> -1 Then
                        tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = UBound(tlOutSplitCategoryInfo)
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                        tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                        ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                        llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                        Do While llExcludeIndex <> -1
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                            tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo) - 1).lNext = UBound(tlOutSplitCategoryInfo)
                            ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                            llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                        Loop
                    End If
                    ReDim Preserve tlOutRegionDefinition(0 To UBound(tlOutRegionDefinition) + 1) As REGIONDEFINITION
                    llOtherIndex = tlInSplitCategoryInfo(llOtherIndex).lNext
                Loop While llOtherIndex <> -1
            Else
                llOtherIndex = tlInRegionDefinition(llRegion).lOtherFirst
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)) = tlInRegionDefinition(llRegion)
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lFormatFirst = -1
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lOtherFirst = UBound(tlOutSplitCategoryInfo)
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = -1
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).iStationCount = ilStationCount
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lStationOtherFirst = tlInRegionDefinition(llRegion).lOtherFirst
                'Placeholder
                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llOtherIndex)
                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                ReDim Preserve tlOutRegionDefinition(0 To UBound(tlOutRegionDefinition) + 1) As REGIONDEFINITION
            End If
        Else
            'Exclude only
            llExcludeIndex = tlInRegionDefinition(llRegion).lExcludeFirst
            If llExcludeIndex <> -1 Then
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)) = tlInRegionDefinition(llRegion)
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lFormatFirst = -1
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lOtherFirst = -1
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lExcludeFirst = UBound(tlOutSplitCategoryInfo)
                '3/3/18
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).iStationCount = -1
                tlOutRegionDefinition(UBound(tlOutRegionDefinition)).lStationOtherFirst = -1
                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                Do While llExcludeIndex <> -1
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)) = tlInSplitCategoryInfo(llExcludeIndex)
                    tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo)).lNext = -1
                    tlOutSplitCategoryInfo(UBound(tlOutSplitCategoryInfo) - 1).lNext = UBound(tlOutSplitCategoryInfo)
                    ReDim Preserve tlOutSplitCategoryInfo(0 To UBound(tlOutSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                    llExcludeIndex = tlInSplitCategoryInfo(llExcludeIndex).lNext
                Loop
                ReDim Preserve tlOutRegionDefinition(0 To UBound(tlOutRegionDefinition) + 1) As REGIONDEFINITION
            End If
        End If
    Next llRegion

End Sub


Private Function mCompareLST(llLstCode As Long, llAstLstCode As Long, llSdfCode As Long) As Long
    'Return:
    '0=No match
    '1=LST match (lstBkoutLstCode = 0)
    '2+Index = LST match blackout LST
    Dim llLst As Long
    Dim ilRet As Integer
    
    On Error GoTo ErrHand
    
    mCompareLST = 0
    If llLstCode = llAstLstCode Then
        mCompareLST = 1 'LST codes match
        Exit Function
    End If
    For llLst = 0 To UBound(tmBkoutLst) - 1 Step 1
        If igExportSource = 2 Then
            DoEvents
        End If
        '10/12/14
        If (tmBkoutLst(llLst).tLST.lBkoutLstCode = llLstCode) And (tmBkoutLst(llLst).bMatched = False) Then
            If (tmBkoutLst(llLst).tLST.lCode = llAstLstCode) And (tmBkoutLst(llLst).tLST.lSdfCode = llSdfCode) Then
                mCompareLST = 2 + llLst 'Blackout LST codes match
                Exit Function
            End If
        End If
    Next llLst
    For llLst = 0 To UBound(tmBkoutLst) - 1 Step 1
        If igExportSource = 2 Then
            DoEvents
        End If
        '10/12/14
        If (tmBkoutLst(llLst).tLST.lCode = llAstLstCode) And (tmBkoutLst(llLst).bMatched = False) Then
            ''Remove the Blackout record as LST changed
            'slSQLQuery = "DELETE FROM lst WHERE (lstCode = " & llAstLstCode & ")"
            'ilRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            'tmBkoutLst(llLst).lCode = -1
            tmBkoutLst(llLst).iDelete = True
        End If
    Next llLst
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-mCompareLST"
End Function

Public Sub gCloseRegionSQLRst()
    On Error Resume Next
    dat_rst.Close
    lst_rst.Close
    att_rst.Close
    cptt_rst.Close
    sef_rst.Close
    rsf_rst.Close
    gClearASTInfo False
End Sub



Private Sub mBuildSplitCategoryInfo(slTarget As String, llUpperSplit As Long, tlRegionDefinition As REGIONDEFINITION, tlSplitCategoryInfo() As SPLITCATEGORYINFO)
    Dim llPreviousOther As Long
    Dim llPreviousFormat As Long
    Dim llPreviousExclude As Long
    Dim slCategory As String
    Dim slInclExcl As String
    Dim ilAddExclude As Integer
    Dim ilShtt As Integer
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    
    llPreviousFormat = -1
    llPreviousOther = -1
    llPreviousExclude = -1
    slSQLQuery = "Select * from SEF_Split_Entity"
    slSQLQuery = slSQLQuery & " Where (sefRafCode = " & tlRegionDefinition.lRafCode & ")"
    'slSQLQuery = slSQLQuery + " ORDER BY sefRafCode, sefSeqNo"
    Set sef_rst = gSQLSelectCall(slSQLQuery)
    Do While Not sef_rst.EOF
        If igExportSource = 2 Then
            DoEvents
        End If
        slCategory = tlRegionDefinition.sCategory
        slInclExcl = tlRegionDefinition.sInclExcl
        If Trim$(sef_rst!sefCategory) <> "" Then
            slCategory = sef_rst!sefCategory
            slInclExcl = sef_rst!sefInclExcl
        End If
        If slInclExcl <> "E" Then
            If slCategory = "F" Then
                'Add to Format table
                If tlRegionDefinition.lFormatFirst = -1 Then
                    tlRegionDefinition.lFormatFirst = llUpperSplit
                End If
                tlSplitCategoryInfo(llUpperSplit).sCategory = slCategory
                tlSplitCategoryInfo(llUpperSplit).sName = sef_rst!sefName
                tlSplitCategoryInfo(llUpperSplit).iIntCode = sef_rst!sefIntCode
                tlSplitCategoryInfo(llUpperSplit).lLongCode = sef_rst!sefLongCode
                tlSplitCategoryInfo(llUpperSplit).lNext = -1
                If llPreviousFormat <> -1 Then
                    tlSplitCategoryInfo(llPreviousFormat).lNext = llUpperSplit
                End If
                llPreviousFormat = llUpperSplit
                'ReDim Preserve tlSplitCategoryInfo(0 To UBound(tlSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                llUpperSplit = llUpperSplit + 1
                If llUpperSplit >= UBound(tlSplitCategoryInfo) Then
                    ReDim Preserve tlSplitCategoryInfo(0 To UBound(tlSplitCategoryInfo) + 100) As SPLITCATEGORYINFO
                End If
            Else
                'Add to All Category except Format table
                If tlRegionDefinition.lOtherFirst = -1 Then
                    tlRegionDefinition.lOtherFirst = llUpperSplit
                End If
                tlSplitCategoryInfo(llUpperSplit).sCategory = slCategory
                tlSplitCategoryInfo(llUpperSplit).sName = sef_rst!sefName
                tlSplitCategoryInfo(llUpperSplit).iIntCode = sef_rst!sefIntCode
                tlSplitCategoryInfo(llUpperSplit).lLongCode = sef_rst!sefLongCode
                tlSplitCategoryInfo(llUpperSplit).lNext = -1
                If llPreviousOther <> -1 Then
                    tlSplitCategoryInfo(llPreviousOther).lNext = llUpperSplit
                End If
                llPreviousOther = llUpperSplit
                'ReDim Preserve tlSplitCategoryInfo(0 To UBound(tlSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                llUpperSplit = llUpperSplit + 1
                If llUpperSplit >= UBound(tlSplitCategoryInfo) Then
                    ReDim Preserve tlSplitCategoryInfo(0 To UBound(tlSplitCategoryInfo) + 100) As SPLITCATEGORYINFO
                End If
            End If
        Else
            'Add to Exclude table
            ilAddExclude = True
            If (slCategory = "S") And ((slTarget = "W") Or (slTarget = "O")) Then
                ilShtt = gBinarySearchStationInfoByCode(sef_rst!sefIntCode)
                If ilShtt <> -1 Then
                    If slTarget = "W" Then
                        If tgStationInfoByCode(ilShtt).sUsedForWegener <> "Y" Then
                            ilAddExclude = False
                        End If
                    ElseIf slTarget = "O" Then
                        If tgStationInfoByCode(ilShtt).sUsedForOLA <> "Y" Then
                            ilAddExclude = False
                        End If
                    End If
                Else
                    ilAddExclude = False
                End If
            End If
            If ilAddExclude Then
                If tlRegionDefinition.lExcludeFirst = -1 Then
                    tlRegionDefinition.lExcludeFirst = llUpperSplit
                End If
                tlSplitCategoryInfo(llUpperSplit).sCategory = slCategory
                tlSplitCategoryInfo(llUpperSplit).sName = sef_rst!sefName
                tlSplitCategoryInfo(llUpperSplit).iIntCode = sef_rst!sefIntCode
                tlSplitCategoryInfo(llUpperSplit).lLongCode = sef_rst!sefLongCode
                tlSplitCategoryInfo(llUpperSplit).lNext = -1
                If llPreviousExclude <> -1 Then
                    tlSplitCategoryInfo(llPreviousExclude).lNext = llUpperSplit
                End If
                llPreviousExclude = llUpperSplit
                'ReDim Preserve tlSplitCategoryInfo(0 To UBound(tlSplitCategoryInfo) + 1) As SPLITCATEGORYINFO
                llUpperSplit = llUpperSplit + 1
                If llUpperSplit >= UBound(tlSplitCategoryInfo) Then
                    ReDim Preserve tlSplitCategoryInfo(0 To UBound(tlSplitCategoryInfo) + 100) As SPLITCATEGORYINFO
                End If
            End If
        End If
        sef_rst.MoveNext
    Loop
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-mBuildSplitCategoryInfo"
End Sub

Private Sub mBuildRegionForSpots(ilVefCode As Integer)
    Dim ilLstLoop As Integer
    Dim llRegionDefinitionIndex As Long
    Dim llUpperSplit As Long
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim ilSeqNo As Integer
    Dim slSdfCode As String
    Dim slSeqNo As String
    
    
    'ReDim tlRegionDefinition(0 To 0) As REGIONDEFINITION
    'ReDim tlSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
    ReDim tlPartRegionDef(0 To 1) As REGIONDEFINITION
    
    ReDim tmRegionAssignmentInfo(0 To 0) As REGIONASSIGNMENTINFO
    ReDim tmRegionDefinitionForSpots(0 To 0) As REGIONDEFINITION
    ReDim tmSplitCategoryInfoForSpots(0 To 0) As SPLITCATEGORYINFO
    For ilLstLoop = 0 To UBound(lmRCSdfCode) - 1 Step 1
        If igExportSource = 2 Then
            DoEvents
        End If
        ilSeqNo = 0
        ReDim tlRegionDefinition(0 To 0) As REGIONDEFINITION
        ReDim tlSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
        ilRet = gBuildRegionDefinitions("C", lmRCSdfCode(ilLstLoop), ilVefCode, tlRegionDefinition(), tlSplitCategoryInfo())
        If UBound(tlRegionDefinition) > 0 Then
            For llRegionDefinitionIndex = 0 To UBound(tlRegionDefinition) - 1 Step 1
                If igExportSource = 2 Then
                    DoEvents
                End If
                slSdfCode = Trim$(Str$(lmRCSdfCode(ilLstLoop)))
                Do While Len(slSdfCode) < 10
                    slSdfCode = "0" & slSdfCode
                Loop
                ilSeqNo = ilSeqNo + 1
                slSeqNo = Trim$(Str$(ilSeqNo))
                Do While Len(slSeqNo) < 5
                    slSeqNo = "0" & slSeqNo
                Loop
                tmRegionAssignmentInfo(UBound(tmRegionAssignmentInfo)).sKey = slSdfCode & slSeqNo
                tmRegionAssignmentInfo(UBound(tmRegionAssignmentInfo)).lSdfCode = lmRCSdfCode(ilLstLoop)
                tlPartRegionDef(0) = tlRegionDefinition(llRegionDefinitionIndex)
                ReDim tlSplitCategoryInfo(0 To 500) As SPLITCATEGORYINFO
                llUpperSplit = 0
                mBuildSplitCategoryInfo "C", llUpperSplit, tlPartRegionDef(0), tlSplitCategoryInfo()
                'ReDim Preserve tlSplitCategoryInfo(0 To llUpperSplit) As SPLITCATEGORYINFO
                tmRegionAssignmentInfo(UBound(tmRegionAssignmentInfo)).lRDIndex = UBound(tmRegionDefinitionForSpots)
                tmRegionDefinitionForSpots(UBound(tmRegionDefinitionForSpots)) = tlPartRegionDef(0)
                ReDim Preserve tmRegionDefinitionForSpots(0 To UBound(tmRegionDefinitionForSpots) + 1) As REGIONDEFINITION
                tmRegionAssignmentInfo(UBound(tmRegionAssignmentInfo)).lSCIStartIndex = UBound(tmSplitCategoryInfoForSpots)
                tmRegionAssignmentInfo(UBound(tmRegionAssignmentInfo)).lSCIEndIndex = -1
                For llLoop = 0 To llUpperSplit Step 1
                    tmSplitCategoryInfoForSpots(UBound(tmSplitCategoryInfoForSpots)) = tlSplitCategoryInfo(llLoop)
                    ReDim Preserve tmSplitCategoryInfoForSpots(0 To UBound(tmSplitCategoryInfoForSpots) + 1) As SPLITCATEGORYINFO
                    tmRegionAssignmentInfo(UBound(tmRegionAssignmentInfo)).lSCIEndIndex = UBound(tmSplitCategoryInfoForSpots) - 1
                Next llLoop
                If llUpperSplit > 0 Then
                    ReDim Preserve tmRegionAssignmentInfo(0 To UBound(tmRegionAssignmentInfo) + 1) As REGIONASSIGNMENTINFO
                End If
            Next llRegionDefinitionIndex
        End If
    Next ilLstLoop
    If UBound(tmRegionAssignmentInfo) - 1 > 0 Then
        ArraySortTyp fnAV(tmRegionAssignmentInfo(), 0), UBound(tmRegionAssignmentInfo), 0, LenB(tmRegionAssignmentInfo(0)), 0, LenB(tmRegionAssignmentInfo(0).sKey), 0
    End If
End Sub

Public Function mBinarySearchSdfCode(llSdfCode As Long) As Long
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    Dim ilResult As Integer
    Dim slSdfCode As String
    Dim slSeqNo As String
    Dim ilSeqNo As Integer
    Dim slSort As String
    
    
    slSdfCode = Trim$(Str$(llSdfCode))
    Do While Len(slSdfCode) < 10
        slSdfCode = "0" & slSdfCode
    Loop
    ilSeqNo = 1
    slSeqNo = Trim$(Str$(ilSeqNo))
    Do While Len(slSeqNo) < 5
        slSeqNo = "0" & slSeqNo
    Loop
    slSort = slSdfCode & slSeqNo
    mBinarySearchSdfCode = -1    ' Start out as not found.
    llMin = LBound(tmRegionAssignmentInfo)
    llMax = UBound(tmRegionAssignmentInfo) - 1
    Do While llMin <= llMax
        If igExportSource = 2 Then
            DoEvents
        End If
        llMiddle = (llMin + llMax) \ 2
        ilResult = StrComp(Trim$(tmRegionAssignmentInfo(llMiddle).sKey), slSort, vbBinaryCompare)
        Select Case ilResult
            Case 0:
                mBinarySearchSdfCode = llMiddle  ' Found it !
                Exit Function
            Case 1:
                llMax = llMiddle - 1
            Case -1:
                llMin = llMiddle + 1
        End Select
    Loop
    Exit Function

End Function

Public Sub gClearASTInfo(ilClearBlackout As Integer)
    Dim llLst As Long
    Dim ilRet As Integer
    Dim ilLowLimit As Integer
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    
    smGetAstKey = ""
    Erase tmEZLSTAst
    Erase tmCZLSTAst
    Erase tmMZLSTAst
    Erase tmPZLSTAst
    Erase tmNZLSTAst
    Erase tmLSTAst
    Erase lmRCSdfCode
    Erase tmATTCrossDates
    smBuildLstKey = ""
    Erase tmBuildLst
    
    If ilClearBlackout Then
        'On Error GoTo gClearASTInfoErr
        ilRet = 0
        'ilLowLimit = LBound(tmBkoutLst)
        If PeekArray(tmBkoutLst).Ptr <> 0 Then
            ilLowLimit = LBound(tmBkoutLst)
        Else
            ilRet = 1
            ilLowLimit = 0
        End If
        
        If ilRet = 0 Then
            On Error GoTo ErrHand
            For llLst = 0 To UBound(tmBkoutLst) - 1 Step 1
                If igExportSource = 2 Then
                    DoEvents
                End If
                If tmBkoutLst(llLst).iDelete Then
                    slSQLQuery = "DELETE FROM lst WHERE (lstCode = " & tmBkoutLst(llLst).tLST.lCode & ")"
                    'Dan m 9/18/09 now returns long
                    'ilRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                    gSQLWaitNoMsgBox slSQLQuery, False
                End If
            Next llLst
        End If
    End If
    Erase tmBkoutLst
    Exit Sub
gClearASTInfoErr:
    ilRet = 1
    Resume Next
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "AffErrorLog.txt", "modCPReturns-gClearASTInfo"
End Sub

Public Function gRegionCopyExistForSpot(tlAstInfo As ASTINFO, llRsfCode As Long) As Integer
    Dim llSdf As Long
    Dim ilShttCode As Integer
    Dim llSdfCode As Long
    Dim ilVefCode As Integer
    Dim ilMktCode As Integer
    Dim ilMSAMktCode As Integer
    Dim slState As String
    Dim ilFmtCode As Integer
    Dim ilTztCode As Integer
    Dim ilRet As Integer
    Dim llLoop As Long
    Dim llRegionIndex As Long
    Dim slGroupInfo As String
    ReDim tlSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
    ReDim tmRegionDefinition(0 To 0) As REGIONDEFINITION
    ReDim tmSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
    ReDim tlPartRegionDef(0 To 1) As REGIONDEFINITION
    
    gRegionCopyExistForSpot = False
    If igExportSource = 2 Then
        DoEvents
    End If
    llRsfCode = 0
    ReDim lmRCSdfCode(0 To 1) As Long
    lmRCSdfCode(0) = tlAstInfo.lSdfCode
    mBuildRegionForSpots tlAstInfo.iVefCode
    ilShttCode = tlAstInfo.iShttCode
    llSdfCode = tlAstInfo.lSdfCode
    ilVefCode = tlAstInfo.iVefCode
    ilRet = gBinarySearchStationInfoByCode(ilShttCode)
    If ilRet = -1 Then
        Exit Function
    End If
    ilMktCode = tgStationInfoByCode(ilRet).iMktCode
    ilMSAMktCode = tgStationInfoByCode(ilRet).iMSAMktCode
    '12/28/15
    'slState = tgStationInfoByCode(ilRet).sPostalName
    If sgSplitState = "L" Then
        slState = tgStationInfoByCode(ilRet).sStateLic
    ElseIf sgSplitState = "P" Then
        slState = tgStationInfoByCode(ilRet).sPhyState
    Else
        slState = tgStationInfoByCode(ilRet).sMailState
    End If
    ilFmtCode = tgStationInfoByCode(ilRet).iFormatCode
    ilTztCode = tgStationInfoByCode(ilRet).iTztCode
    llSdf = mBinarySearchSdfCode(llSdfCode)
    If llSdf <> -1 Then
        tlPartRegionDef(0) = tmRegionDefinitionForSpots(tmRegionAssignmentInfo(llSdf).lRDIndex)
        ReDim tlSplitCategoryInfo(0 To tmRegionAssignmentInfo(llSdf).lSCIEndIndex - tmRegionAssignmentInfo(llSdf).lSCIStartIndex + 1) As SPLITCATEGORYINFO
        For llLoop = tmRegionAssignmentInfo(llSdf).lSCIStartIndex To tmRegionAssignmentInfo(llSdf).lSCIEndIndex Step 1
            tlSplitCategoryInfo(llLoop - tmRegionAssignmentInfo(llSdf).lSCIStartIndex) = tmSplitCategoryInfoForSpots(llLoop)
        Next llLoop

        ReDim tmRegionDefinition(0 To 0) As REGIONDEFINITION
        ReDim tmSplitCategoryInfo(0 To 0) As SPLITCATEGORYINFO
        
        gSeparateRegions tlPartRegionDef(), tlSplitCategoryInfo(), tmRegionDefinition(), tmSplitCategoryInfo()

        ilRet = gRegionTestDefinition(ilShttCode, ilMktCode, ilMSAMktCode, slState, ilFmtCode, ilTztCode, tmRegionDefinition(), tmSplitCategoryInfo(), llRegionIndex, slGroupInfo)
        If ilRet Then
            llRsfCode = tmRegionDefinition(llRegionIndex).lRsfCode
            gRegionCopyExistForSpot = True
        End If
    End If
End Function

Public Sub gIncSpotCounts(tlAstInfo As ASTINFO, ilOutSchdCount As Integer, ilOutAiredCount As Integer, ilPledgeCompliantCount As Integer, ilAgencyCompliantCount As Integer, Optional slServiceAgreement As String = "N")
    Dim llMonPdDate As Long
    Dim llSunPdDate As Long
    Dim ilAstStatus As Integer
        
    Dim ilInAstPledgeStatus As Integer
    Dim ilInAstStatus As Integer
    Dim ilInAstCPStatus As Integer
    Dim slInPledgeDays As String
    Dim slInAstPledgeDate As String
    Dim slInAstAirDate As String
    Dim slInAstPledgeStartTime As String
    Dim slInAstTruePledgeEndTime As String
    Dim slInAstAirTime As String
    
    Dim slStartDate As String
    Dim slEndDate As String
    Dim slStartTime As String
    Dim slEndTime As String
    Dim ilAllowedDays(0 To 6) As Integer
    Dim ilCompliant As Integer
    Dim ilTimeAdj As Integer
    Dim ilVff As Integer
    
    Dim slStationCompliant As String * 1
    Dim slAgencyCompliant As String * 1
    Dim ilRet As Integer
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand:
    ilInAstPledgeStatus = tlAstInfo.iPledgeStatus
    ilInAstStatus = tlAstInfo.iStatus
    ilInAstCPStatus = tlAstInfo.iCPStatus
    slInPledgeDays = tlAstInfo.sTruePledgeDays
    slInAstPledgeDate = Format$(tlAstInfo.sPledgeDate, "m/d/yy")
    slInAstAirDate = Format$(tlAstInfo.sAirDate, "m/d/yy")
    
    ilVff = gBinarySearchVff(tlAstInfo.iVefCode)
    If ilVff <> -1 Then
        ilTimeAdj = tgVffInfo(ilVff).iLiveCompliantAdj
        If ilTimeAdj <= 0 Then
            ilTimeAdj = 5
        End If
    Else
        ilTimeAdj = 5
    End If
    If tgStatusTypes(gGetAirStatus(ilInAstPledgeStatus)).iPledged = 0 Then   'Live
        slInAstPledgeStartTime = gLongToTime(gTimeToLong(Format$(tlAstInfo.sPledgeStartTime, "h:mm:ssAM/PM"), False) - 60 * ilTimeAdj)
        If gTimeToLong(Format$(tlAstInfo.sTruePledgeEndTime, "h:mm:ssAM/PM"), True) = 86400 Then
            slInAstTruePledgeEndTime = 86400
        Else
            slInAstTruePledgeEndTime = gLongToTime(gTimeToLong(Format$(tlAstInfo.sTruePledgeEndTime, "h:mm:ssAM/PM"), True) + 60 * ilTimeAdj)
        End If
    Else
        slInAstPledgeStartTime = Format$(tlAstInfo.sPledgeStartTime, "h:mm:ssAM/PM")
        slInAstTruePledgeEndTime = Format$(tlAstInfo.sTruePledgeEndTime, "h:mm:ssAM/PM")
    End If
    slInAstAirTime = Format$(tlAstInfo.sAirTime, "h:mm:ssAM/PM")
    
    If igExportSource = 2 Then
        DoEvents
    End If
    If gGetAirStatus(ilInAstStatus) = 6 Then
        ilAstStatus = 1
    Else
        ilAstStatus = ilInAstStatus
    End If
    If tgStatusTypes(gGetAirStatus(ilInAstPledgeStatus)).iPledged < 2 Then
        slAgencyCompliant = ""
        slStationCompliant = ""
        If gGetAirStatus(ilAstStatus) < ASTEXTENDED_MG Or (gGetAirStatus(ilAstStatus) = ASTAIR_MISSED_MG_BYPASS) Then
            If tgStatusTypes(gGetAirStatus(ilInAstPledgeStatus)).iPledged = 0 Then   'Live
                ilOutSchdCount = ilOutSchdCount + 1
            ElseIf tgStatusTypes(gGetAirStatus(ilInAstPledgeStatus)).iPledged = 1 Then    'Delayed
                ilOutSchdCount = ilOutSchdCount + 1
            ElseIf tgStatusTypes(gGetAirStatus(ilInAstPledgeStatus)).iPledged = 2 Then    'Not Carried
            ElseIf tgStatusTypes(gGetAirStatus(ilInAstPledgeStatus)).iPledged = 3 Then    'No Pledged
            End If
            If ilInAstCPStatus >= 1 Then
                If tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 0 Then   'Live
                    ilOutAiredCount = ilOutAiredCount + 1
                ElseIf tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 1 Then    'Delayed
                    ilOutAiredCount = ilOutAiredCount + 1
                ElseIf tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 2 Then    'Not Carried
                ElseIf tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 3 Then    'No Pledged
                    ilOutAiredCount = ilOutAiredCount + 1
                End If
                'Get pledge information from Contract if using Marketron and agreement is by Marketron
                'If sgMarketronCompliant = "A" Then
                    'If tlAstInfo.lAttCode <> lmMktronAttCode Then
                    '    slSQLQuery = "SELECT attExportToMarketron"
                    '    slSQLQuery = slSQLQuery + " FROM att"
                    '    slSQLQuery = slSQLQuery + " WHERE (attCode = " & tlAstInfo.lAttCode & ")"
                    '    Set att_rst = gSQLSelectCall(slSQLQuery)
                    '    If Not att_rst.EOF Then
                    '        lmMktronAttCode = tlAstInfo.lAttCode
                    '        smAttExportToMarketron = att_rst!attExportToMarketron
                    '        If att_rst!attExportToMarketron = "Y" Then
                    '            gGetLineParameters True, tlAstInfo, slStartDate, slEndDate, slStartTime, slEndTime, ilAllowedDays(), ilCompliant
                    '            If ilCompliant Then
                    '                ilAgencyCompliantCount = ilAgencyCompliantCount + 1
                    '            End If
                    '            Exit Sub
                    '        End If
                    '    Else
                    '        lmMktronAttCode = -1
                    '    End If
                    'Else
                        'If smAttExportToMarketron = "Y" Then
                            'gGetLineParameters True, tlAstInfo, slStartDate, slEndDate, slStartTime, slEndTime, ilAllowedDays(), ilCompliant
                            '9/25/14: Blackout spots will always to network compliant
                            '9/29/14: Fill treated as Compliant (SpotType = 2)
                            If (tlAstInfo.lLstBkoutLstCode > 0) Then 'Blackout: Not compliant if missed; Compliant if aired
                                If (tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 2) Or (tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 3) Then
                                    slAgencyCompliant = "N"
                                Else
                                    ilAgencyCompliantCount = ilAgencyCompliantCount + 1
                                    slAgencyCompliant = "Y"
                                End If
                            ElseIf (tlAstInfo.iSpotType = 2) Then   'Filled: Compiant if aired or missed
                                ilAgencyCompliantCount = ilAgencyCompliantCount + 1
                                slAgencyCompliant = "Y"
                            Else
                                If (tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 2) Or (tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 3) Then
                                    slAgencyCompliant = "N"
                                Else
                                    'MG (SpotType = 1; Outside (SpotType = 3): Always Compliant unless missed
                                    If (tlAstInfo.iSpotType = 1) Or (tlAstInfo.iSpotType = 3) Then
                                        ilAgencyCompliantCount = ilAgencyCompliantCount + 1
                                        slAgencyCompliant = "Y"
                                    Else
                                        gGetAgyCompliant tlAstInfo, slStartDate, slEndDate, slStartTime, slEndTime, ilAllowedDays(), ilCompliant
                                        If ilCompliant Then
                                            ilAgencyCompliantCount = ilAgencyCompliantCount + 1
                                            slAgencyCompliant = "Y"
                                        Else
                                            slAgencyCompliant = "N"
                                        End If
                                    End If
                                End If
                            End If
                        '    Exit Sub
                        'End If
                    'End If
                'End If
                '9/25/14: Blackout spots will always to station compliant
                '9/29/14: Fill treated as Compliant (SpotType = 2)
                If (tlAstInfo.lLstBkoutLstCode > 0) Then
                    If (tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 2) Or (tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 3) Then
                        'Non-Compliant
                        slStationCompliant = "N"
                    Else
                        ilPledgeCompliantCount = ilPledgeCompliantCount + 1
                        slStationCompliant = "Y"
                    End If
                ElseIf (tlAstInfo.iSpotType = 2) Then
                    If (tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 2) Or (tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 3) Then
                        'Non-Compliant
                        slStationCompliant = "N"
                    Else
                        ilPledgeCompliantCount = ilPledgeCompliantCount + 1
                        slStationCompliant = "Y"
                    End If
                Else
                    llMonPdDate = DateValue(gObtainPrevMonday(slInAstPledgeDate))
                    llSunPdDate = llMonPdDate + 6
                    'If (ilInAstStatus = 2) Or (ilInAstStatus = 3) Or (ilInAstStatus = 4) Or (ilInAstStatus = 5) Or (ilInAstStatus = 8) Then
                    If (tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 2) Or (tgStatusTypes(gGetAirStatus(ilAstStatus)).iPledged = 3) Then
                        'Non-Compliant
                        slStationCompliant = "N"
                    Else
                        'MG (SpotType = 1; Outside (SpotType = 3): Always Compliant unless missed
                        If (tlAstInfo.iSpotType = 1) Or (tlAstInfo.iSpotType = 3) Then
                            ilPledgeCompliantCount = ilPledgeCompliantCount + 1
                            slStationCompliant = "Y"
                        Else
                            llMonPdDate = DateValue(gObtainPrevMonday(slInAstPledgeDate))
                            llSunPdDate = llMonPdDate + 6
                            If (Weekday(Format$(slInAstPledgeDate, "m/d/yy")) <> Weekday(Format$(slInAstAirDate, "m/d/yy"))) And (tgStatusTypes(gGetAirStatus(ilInAstPledgeStatus)).iPledged = 0) Then
                            'Test if aired on Pledge day
                                slStationCompliant = "N"
                            ElseIf (Mid$(slInPledgeDays, Weekday(Format$(slInAstAirDate, "m/d/yy"), vbMonday), 1) <> "Y") And (tgStatusTypes(gGetAirStatus(ilInAstPledgeStatus)).iPledged = 1) Then
                                'Non-Compliant
                                slStationCompliant = "N"
                            'Test if aired within pledge times
                            ElseIf (gTimeToLong(Format$(slInAstAirTime, "h:mm:ssAM/PM"), False) < gTimeToLong(Format$(slInAstPledgeStartTime, "h:mm:ssAM/PM"), False)) Or (gTimeToLong(Format$(slInAstAirTime, "h:mm:ssAM/PM"), False) > gTimeToLong(Format$(slInAstTruePledgeEndTime, "h:mm:ssAM/PM"), True)) Then
                                'If (ilInAstStatus = 0) Or (ilInAstStatus = 1) Or (ilInAstStatus = 6) Or (ilInAstStatus = 7) Or (ilInAstStatus = 9) Or (ilInAstStatus = 10) Then
                                '    'Non-Compliant
                                'Else
                                '    ilOutCompliantCount = ilOutCompliantCount + 1
                                'End If
                                slStationCompliant = "N"
                            Else
                                'Test if in correct week
                                'llMonPdDate = DateValue(gObtainPrevMonday(slInAstPledgeDate))
                                'If (DateValue(slInAstAirDate) >= llMonPdDate) And (DateValue(slInAstAirDate) <= llMonPdDate + 6) Then
                                If (DateValue(slInAstAirDate) >= llMonPdDate) And (DateValue(slInAstAirDate) <= llSunPdDate) Then
                                    ilPledgeCompliantCount = ilPledgeCompliantCount + 1
                                    slStationCompliant = "Y"
                                Else
                                    slStationCompliant = "N"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            If (slAgencyCompliant <> "") Or (slStationCompliant <> "") Then
                slSQLQuery = "UPDATE ast SET "
                slSQLQuery = slSQLQuery + "astAgencyCompliant = '" & slAgencyCompliant & "',"
                slSQLQuery = slSQLQuery + "astStationCompliant = '" & slStationCompliant & "'"
                slSQLQuery = slSQLQuery + " WHERE (astCode = " & tlAstInfo.lCode & ")"
                ilRet = gSQLWaitNoMsgBox(slSQLQuery, False)
            End If
        End If
    Else
        '8/5/14: Reports will be testing for N and not carried need to be bypassed.
        slSQLQuery = "UPDATE ast SET "
        slSQLQuery = slSQLQuery + "astAgencyCompliant = '" & "Y" & "',"
        slSQLQuery = slSQLQuery + "astStationCompliant = '" & "Y" & "'"
        slSQLQuery = slSQLQuery + " WHERE (astCode = " & tlAstInfo.lCode & ")"
    End If
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "gIncSpotCounts"
End Sub


Public Function gMatchAstAndDat(slFeedStartTime As String, slFeedDate As String, slPledgeStartTime As String, slPledgeDate As String, tlDat() As DATRST) As Long
    Dim llFeedStartTime As Long
    Dim llPledgeStartTime As Long
    Dim llDat As Long
    Dim ilDayOk As Integer
    Dim ilFdDay As Integer
    Dim ilPdDay As Integer
    
    
    gMatchAstAndDat = -1
    llFeedStartTime = gTimeToLong(slFeedStartTime, False)
    llPledgeStartTime = gTimeToLong(slPledgeStartTime, False)
    ilFdDay = Weekday(slFeedDate, vbMonday)
    ilPdDay = Weekday(slPledgeDate, vbMonday)
    For llDat = LBound(tlDat) To UBound(tlDat) - 1 Step 1
        If igExportSource = 2 Then
            DoEvents
        End If
        If tgStatusTypes(gGetAirStatus(tlDat(llDat).iFdStatus)).iPledged <> 2 Then

            If (gTimeToLong(tlDat(llDat).sFdStTime, False) = llFeedStartTime) And (gTimeToLong(tlDat(llDat).sPdStTime, False) = llPledgeStartTime) Then
                ilDayOk = False
                Select Case ilFdDay
                    Case 1  'Monday
                        If tlDat(llDat).iFdMon Then
                            ilDayOk = True
                        End If
                    Case 2  'Tuesday
                        If tlDat(llDat).iFdTue Then
                            ilDayOk = True
                        End If
                    Case 3  'Wednesady
                        If tlDat(llDat).iFdWed Then
                            ilDayOk = True
                        End If
                    Case 4  'Thursday
                        If tlDat(llDat).iFdThu Then
                            ilDayOk = True
                        End If
                    Case 5  'Friday
                        If tlDat(llDat).iFdFri Then
                            ilDayOk = True
                        End If
                    Case 6  'Saturday
                        If tlDat(llDat).iFdSat Then
                            ilDayOk = True
                        End If
                    Case 7  'Sunday
                        If tlDat(llDat).iFdSun Then
                            ilDayOk = True
                        End If
                End Select
                If ilDayOk Then
                    ilDayOk = False
                    Select Case ilPdDay
                        Case 1  'Monday
                            If tlDat(llDat).iPdMon Then
                                ilDayOk = True
                            End If
                        Case 2  'Tuesday
                            If tlDat(llDat).iPdTue Then
                                ilDayOk = True
                            End If
                        Case 3  'Wednesday
                            If tlDat(llDat).iPdWed Then
                                ilDayOk = True
                            End If
                        Case 4  'Thursday
                            If tlDat(llDat).iPdThu Then
                                ilDayOk = True
                            End If
                        Case 5  'Friday
                            If tlDat(llDat).iPdFri Then
                                ilDayOk = True
                            End If
                        Case 6  'Saturday
                            If tlDat(llDat).iPdSat Then
                                ilDayOk = True
                            End If
                        Case 7  'Sunday
                            If tlDat(llDat).iPdSun Then
                                ilDayOk = True
                            End If
                    End Select
                End If
                If ilDayOk Then
                    gMatchAstAndDat = llDat
                    Exit For
                End If
            End If
        End If
    Next llDat
End Function


Public Function gSetCpttCount(llAttCode As Long, slSDate As String, slEDate As String) As Integer
    Dim slMoDate As String
    Dim slSuDate As String
    Dim ilAdfCode As Integer
    Dim ilSchdCount As Integer
    Dim ilAiredCount As Integer
    Dim ilPledgeCompliantCount As Integer
    Dim ilAgyCompliantCount As Integer
    Dim slServiceAgreement As String
    Dim ilRet As Integer
    Dim ilAst As Integer
    Dim hlAst As Integer
    Dim slSQLQuery As String
    ReDim tlCPDat(0 To 0) As DAT
    ReDim tlAstInfo(0 To 0) As ASTINFO
    Dim cprst As ADODB.Recordset
        
    On Error GoTo ErrHand:
    gOpenMKDFile hlAst, "Ast.Mkd"
    slMoDate = gAdjYear(gObtainPrevMonday(slSDate))
    Do
        slSuDate = DateAdd("d", 6, slMoDate)
        slSQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttVefCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate, attServiceAgreement"
        slSQLQuery = slSQLQuery & " FROM shtt, cptt, att"
        slSQLQuery = slSQLQuery & " WHERE (ShttCode = cpttShfCode"
        slSQLQuery = slSQLQuery & " AND attCode = cpttAtfCode"
        slSQLQuery = slSQLQuery & " AND cpttAtfCode = " & llAttCode
        slSQLQuery = slSQLQuery & " AND cpttStartDate = '" & Format$(slMoDate, sgSQLDateForm) & "')"
        Set cprst = gSQLSelectCall(slSQLQuery)
        If Not cprst.EOF Then
            ReDim tgCPPosting(0 To 1) As CPPOSTING
            tgCPPosting(0).lCpttCode = cprst!cpttCode
            tgCPPosting(0).iStatus = cprst!cpttStatus
            tgCPPosting(0).iPostingStatus = cprst!cpttPostingStatus
            tgCPPosting(0).lAttCode = cprst!cpttatfCode
            tgCPPosting(0).iAttTimeType = cprst!attTimeType
            tgCPPosting(0).iVefCode = cprst!cpttvefcode
            tgCPPosting(0).iShttCode = cprst!shttCode
            tgCPPosting(0).sZone = cprst!shttTimeZone
            tgCPPosting(0).sDate = Format$(slMoDate, sgShowDateForm)
            tgCPPosting(0).sAstStatus = cprst!cpttAstStatus
            slServiceAgreement = cprst!attServiceAgreement
            igTimes = 1 'By Week
            ilAdfCode = -1
            'Dan M 9/26/13 6442
            'ilRet = gGetAstInfo(hlAst, tlCPDat(), tlAstInfo(), ilAdfCode, False, False, True, False)
            ilRet = gGetAstInfo(hlAst, tlCPDat(), tlAstInfo(), ilAdfCode, True, False, True, False, , , , , True)
            ilSchdCount = 0
            ilAiredCount = 0
            ilPledgeCompliantCount = 0
            ilAgyCompliantCount = 0
            For ilAst = LBound(tlAstInfo) To UBound(tlAstInfo) - 1 Step 1
                'gIncSpotCounts tlAstInfo(ilAst).iPledgeStatus, tlAstInfo(ilAst).iStatus, tlAstInfo(ilAst).iCPStatus, tlAstInfo(ilAst).sTruePledgeDays, Format$(tlAstInfo(ilAst).sPledgeDate, "m/d/yy"), Format$(tlAstInfo(ilAst).sAirDate, "m/d/yy"), Format$(tlAstInfo(ilAst).sPledgeStartTime, "h:mm:ssAM/PM"), Format$(tlAstInfo(ilAst).sTruePledgeEndTime, "h:mm:ssAM/PM"), Format$(tlAstInfo(ilAst).sAirTime, "h:mm:ssAM/PM"), ilSchdCount, ilAiredCount, ilCompliantCount
                gIncSpotCounts tlAstInfo(ilAst), ilSchdCount, ilAiredCount, ilPledgeCompliantCount, ilAgyCompliantCount, slServiceAgreement
            Next ilAst
            slSQLQuery = "Update cptt Set "
            slSQLQuery = slSQLQuery & "cpttNoSpotsGen = " & ilSchdCount & ", "
            slSQLQuery = slSQLQuery & "cpttNoSpotsAired = " & ilAiredCount & ", "
            slSQLQuery = slSQLQuery & "cpttNoCompliant = " & ilPledgeCompliantCount & ", "
            slSQLQuery = slSQLQuery & "cpttAgyCompliant = " & ilAgyCompliantCount & " "
            slSQLQuery = slSQLQuery & " WHERE cpttAtfCode = " & llAttCode
            slSQLQuery = slSQLQuery & " AND (cpttStartDate >= '" & Format$(slMoDate, sgSQLDateForm) & "' AND cpttStartDate <= '" & Format$(slSuDate, sgSQLDateForm) & "')"
            If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                '6/13/16: Replaced GoSub
                'GoSub ErrHand1:
                gHandleError "AffErrorLog.txt", "modCPReturns-gSetCpttCount"
                On Error Resume Next
                gCloseMKDFile hlAst, "Ast.Mkd"
                cprst.Close
                gSetCpttCount = False
                Exit Function
            End If
        End If
        slMoDate = DateAdd("d", 7, slMoDate)
    Loop While DateValue(gAdjYear(slMoDate)) < DateValue(gAdjYear(slEDate))
    On Error Resume Next
    gCloseMKDFile hlAst, "Ast.Mkd"
    cprst.Close
    gSetCpttCount = True
    Exit Function
ErrHand:
    On Error Resume Next
    gCloseMKDFile hlAst, "Ast.Mkd"
    cprst.Close
    gHandleError "AffErrorLog.txt", "modCPReturns-gSetCpttCount"
    gSetCpttCount = False
    Exit Function
'ErrHand1:
'    On Error Resume Next
'    gCloseMKDFile hlAst, "Ast.Mkd"
'    cprst.Close
'    gHandleError "AffErrorLog.txt", "modCPReturns-gSetCpttCount"
'    gSetCpttCount = False
'    Exit Function
End Function

Private Sub mGetDat(llAttCode As Long, tlCPDat() As DAT)
    Dim ilUpper As Integer
    '5/24/16: Check if Pledge days needs to be reset
    Dim ilPledged As Integer
    Dim ilDay As Integer
    Dim blDayDefined As Boolean
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    
    ' Pledge info is in the dat table.
    ReDim tlCPDat(0 To 0) As DAT
    ilUpper = 0
    slSQLQuery = "SELECT * "
    slSQLQuery = slSQLQuery + " FROM dat"
    slSQLQuery = slSQLQuery + " WHERE datatfCode= " & llAttCode
    slSQLQuery = slSQLQuery + " Order by datFdStTime, datAirPlayNo"
    Set rst = gSQLSelectCall(slSQLQuery)
    If rst.EOF Then
        tlCPDat(0).lAtfCode = llAttCode
        Exit Sub
    End If
    While Not rst.EOF
        If igExportSource = 2 Then
            DoEvents
        End If
        tlCPDat(ilUpper).iStatus = 1         'Used
        tlCPDat(ilUpper).lCode = rst!datCode    '(0).Value
        tlCPDat(ilUpper).lAtfCode = rst!datAtfCode  '(1).Value
        tlCPDat(ilUpper).iShfCode = rst!datShfCode  '(2).Value
        tlCPDat(ilUpper).iVefCode = rst!datVefCode  '(3).Value
        'tlCPDat(ilUpper).iDACode = rst!datDACode    '(4).Value
        tlCPDat(ilUpper).iFdDay(0) = rst!datFdMon   '(5).Value
        tlCPDat(ilUpper).iFdDay(1) = rst!datFdTue   '(6).Value
        tlCPDat(ilUpper).iFdDay(2) = rst!datFdWed   '(7).Value
        tlCPDat(ilUpper).iFdDay(3) = rst!datFdThu   '(8).Value
        tlCPDat(ilUpper).iFdDay(4) = rst!datFdFri   '(9).Value
        tlCPDat(ilUpper).iFdDay(5) = rst!datFdSat   '(10).Value
        tlCPDat(ilUpper).iFdDay(6) = rst!datFdSun   '(11).Value
        If Second(rst!datFdStTime) <> 0 Then
            tlCPDat(ilUpper).sFdSTime = Format$(CStr(rst!datFdStTime), sgShowTimeWSecForm)
        Else
            tlCPDat(ilUpper).sFdSTime = Format$(CStr(rst!datFdStTime), sgShowTimeWOSecForm)
        End If
        If Second(rst!datFdEdTime) <> 0 Then
            tlCPDat(ilUpper).sFdETime = Format$(CStr(rst!datFdEdTime), sgShowTimeWSecForm)
        Else
            tlCPDat(ilUpper).sFdETime = Format$(CStr(rst!datFdEdTime), sgShowTimeWOSecForm)
        End If
        tlCPDat(ilUpper).iFdStatus = rst!datFdStatus    '(14).Value
        tlCPDat(ilUpper).iPdDay(0) = rst!datPdMon   '(15).Value
        tlCPDat(ilUpper).iPdDay(1) = rst!datPdTue   '(16).Value
        tlCPDat(ilUpper).iPdDay(2) = rst!datPdWed   '(17).Value
        tlCPDat(ilUpper).iPdDay(3) = rst!datPdThu   '(18).Value
        tlCPDat(ilUpper).iPdDay(4) = rst!datPdFri   '(19).Value
        tlCPDat(ilUpper).iPdDay(5) = rst!datPdSat   '(20).Value
        tlCPDat(ilUpper).iPdDay(6) = rst!datPdSun   '(21).Value
        tlCPDat(ilUpper).sPdDayFed = rst!datPdDayFed
        If tgStatusTypes(tlCPDat(ilUpper).iFdStatus).iPledged <= 1 Then
            tlCPDat(ilUpper).sPdSTime = Format$(CStr(rst!datPdStTime), sgShowTimeWSecForm)
            tlCPDat(ilUpper).sPdETime = Format$(CStr(rst!datPdEdTime), sgShowTimeWSecForm)
        Else
            tlCPDat(ilUpper).sPdSTime = ""
            tlCPDat(ilUpper).sPdETime = ""
        End If
        tlCPDat(ilUpper).iAirPlayNo = rst!datAirPlayNo
        tlCPDat(ilUpper).sEmbeddedOrROS = rst!datEmbeddedOrROS
        
        '5/24/16: Check if Pledge days needs to be reset
        ilPledged = tgStatusTypes(tlCPDat(ilUpper).iFdStatus).iPledged
        If ilPledged = 0 Then
            For ilDay = 0 To 6 Step 1
                tlCPDat(ilUpper).iPdDay(ilDay) = tlCPDat(ilUpper).iFdDay(ilDay)
            Next ilDay
        ElseIf ilPledged = 1 Then
            blDayDefined = False
            For ilDay = 0 To 6 Step 1
                If tlCPDat(ilUpper).iPdDay(ilDay) <> 0 Then
                    blDayDefined = True
                    Exit For
                End If
            Next ilDay
            If Not blDayDefined Then
                For ilDay = 0 To 6 Step 1
                    tlCPDat(ilUpper).iPdDay(ilDay) = tlCPDat(ilUpper).iFdDay(ilDay)
                Next ilDay
            End If
        ElseIf ilPledged = 2 Then
            For ilDay = 0 To 6 Step 1
                tlCPDat(ilUpper).iPdDay(ilDay) = 0
            Next ilDay
        ElseIf ilPledged = 3 Then
            For ilDay = 0 To 6 Step 1
                tlCPDat(ilUpper).iPdDay(ilDay) = 0
            Next ilDay
        End If
        
        ilUpper = ilUpper + 1
        ReDim Preserve tlCPDat(0 To ilUpper) As DAT
        rst.MoveNext
    Wend
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "mGetDat within gGetAstInfo"
    Exit Sub
End Sub

Public Function gGetProgramTimes(ilInVefCode As Integer, slStartDate As String, slEndDate As String, tlAstTimeRange() As ASTTIMERANGE, Optional blTestMergeOption = True)
    Dim ilVefCode As Integer
    Dim ilVpf As Integer
    Dim slDate As String
    Dim slSDate As String
    Dim slEDate As String
    Dim ilGameNo As Integer
    Dim slSQLQuery As String
   
    On Error GoTo ErrHand
    gGetProgramTimes = False
    ReDim tlAstTimeRange(0 To 0) As ASTTIMERANGE
    If blTestMergeOption Then
        ilVefCode = ilInVefCode
        ilVpf = gBinarySearchVpf(CLng(ilInVefCode))
        If ilVpf <> -1 Then
            If (Asc(tgVpfOptions(ilVpf).sUsingFeatures2) And XDSAPPLYMERGE) <> XDSAPPLYMERGE Then
                gGetProgramTimes = True
                Exit Function
            End If
        End If
    Else
        ilVefCode = ilInVefCode
    End If
    
    slSDate = DateAdd("d", -1, slStartDate)
    slEDate = DateAdd("d", 1, slEndDate)
    slSQLQuery = "SELECT * "
    slSQLQuery = slSQLQuery + " FROM LCF_Log_Calendar"
    slSQLQuery = slSQLQuery + " WHERE ("
    slSQLQuery = slSQLQuery & " lcfStatus = 'C'"
    slSQLQuery = slSQLQuery + " AND lcfLogDate >= '" & Format$(slSDate, sgSQLDateForm) & "' And lcfLogDate <= '" & Format$(slEDate, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " AND lcfVefCode = " & ilVefCode & ")"
    Set rst_lcf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_lcf.EOF
        slDate = Format(rst_lcf!lcfLogDate)
        ilGameNo = rst_lcf!lcfType
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf1, Format$(rst_lcf!lcfTime1, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf2, Format$(rst_lcf!lcfTime2, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf3, Format$(rst_lcf!lcfTime3, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf4, Format$(rst_lcf!lcfTime4, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf5, Format$(rst_lcf!lcfTime5, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf6, Format$(rst_lcf!lcfTime6, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf7, Format$(rst_lcf!lcfTime7, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf8, Format$(rst_lcf!lcfTime8, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf9, Format$(rst_lcf!lcfTime9, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf10, Format$(rst_lcf!lcfTime10, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf11, Format$(rst_lcf!lcfTime11, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf12, Format$(rst_lcf!lcfTime12, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf13, Format$(rst_lcf!lcfTime13, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf14, Format$(rst_lcf!lcfTime14, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf15, Format$(rst_lcf!lcfTime15, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf16, Format$(rst_lcf!lcfTime16, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf17, Format$(rst_lcf!lcfTime17, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf18, Format$(rst_lcf!lcfTime18, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf19, Format$(rst_lcf!lcfTime19, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf20, Format$(rst_lcf!lcfTime20, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf21, Format$(rst_lcf!lcfTime21, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf22, Format$(rst_lcf!lcfTime22, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf23, Format$(rst_lcf!lcfTime23, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf24, Format$(rst_lcf!lcfTime24, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf25, Format$(rst_lcf!lcfTime25, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf26, Format$(rst_lcf!lcfTime26, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf27, Format$(rst_lcf!lcfTime27, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf28, Format$(rst_lcf!lcfTime28, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf29, Format$(rst_lcf!lcfTime29, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf30, Format$(rst_lcf!lcfTime30, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf31, Format$(rst_lcf!lcfTime31, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf32, Format$(rst_lcf!lcfTime32, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf33, Format$(rst_lcf!lcfTime33, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf34, Format$(rst_lcf!lcfTime34, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf35, Format$(rst_lcf!lcfTime35, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf36, Format$(rst_lcf!lcfTime36, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf37, Format$(rst_lcf!lcfTime37, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf38, Format$(rst_lcf!lcfTime38, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf39, Format$(rst_lcf!lcfTime39, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf40, Format$(rst_lcf!lcfTime40, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf41, Format$(rst_lcf!lcfTime41, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf42, Format$(rst_lcf!lcfTime42, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf43, Format$(rst_lcf!lcfTime43, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf44, Format$(rst_lcf!lcfTime44, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf45, Format$(rst_lcf!lcfTime45, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf46, Format$(rst_lcf!lcfTime46, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf47, Format$(rst_lcf!lcfTime47, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf48, Format$(rst_lcf!lcfTime48, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf49, Format$(rst_lcf!lcfTime49, "h:mm:ssAM/PM"), tlAstTimeRange()
        mSetEventTimes slDate, ilGameNo, rst_lcf!lcfLvf50, Format$(rst_lcf!lcfTime50, "h:mm:ssAM/PM"), tlAstTimeRange()
        rst_lcf.MoveNext
    Loop
    gGetProgramTimes = True
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "modCPReturns-gGetProgramTimes"
End Function

Private Sub mSetEventTimes(slDate As String, ilGameNo As Integer, llLvfCode As Long, slLcfStartTime As String, tlAstTimeRange() As ASTTIMERANGE)
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim llLength As Long
    Dim slLength As String
    Dim ilUpper As Integer
    Dim slSQLQuery As String
    
    On Error GoTo mGetEventTimeErr
    If llLvfCode <= 0 Then
        Exit Sub
    End If
    llStartTime = gTimeToLong(slLcfStartTime, False)
    slSQLQuery = "SELECT * "
    slSQLQuery = slSQLQuery + " FROM LVF_Library_Version"
    slSQLQuery = slSQLQuery & " WHERE lvfCode = " & llLvfCode
    Set rst_lvf = gSQLSelectCall(slSQLQuery)
    If Not rst_lvf.EOF Then
        slLength = Format$(rst_lvf!lvfLen, "h:mm:ssAM/PM")
        llLength = gTimeToLong(slLength, False)
        llEndTime = llStartTime + llLength
        ilUpper = UBound(tlAstTimeRange)
        tlAstTimeRange(ilUpper).lDate = gDateValue(slDate)
        tlAstTimeRange(ilUpper).iGameNo = ilGameNo
        tlAstTimeRange(ilUpper).lStartTime = llStartTime
        tlAstTimeRange(ilUpper).lEndTime = llEndTime
        ReDim Preserve tlAstTimeRange(0 To ilUpper + 1) As ASTTIMERANGE
    End If
    Exit Sub
mGetEventTimeErr:
    Exit Sub
End Sub

Public Function gBuildAst(hlAst As Integer, llAttCode As Long, slDate As String, tlAstInfo() As ASTINFO) As Integer
    Dim slMoDate As String
    Dim ilAdfCode As Integer
    Dim ilRet As Integer
    Dim slSQLQuery As String
    ReDim tlCPDat(0 To 0) As DAT
    ReDim tlAstInfo(0 To 0) As ASTINFO
    'Dim cptt_rst As ADODB.Recordset
    
    On Error GoTo ErrHand:
    slMoDate = gAdjYear(gObtainPrevMonday(slDate))
    slSQLQuery = "SELECT shttCallLetters, shttTimeZone, shttCode, cpttStartDate, cpttAtfCode, cpttVefCode, cpttCode, cpttStatus, cpttPostingStatus, cpttAstStatus, attPrintCP, attTimeType, attGenCP, attOnAir, attOffAir, attDropDate"
    slSQLQuery = slSQLQuery & " FROM shtt, cptt, att"
    slSQLQuery = slSQLQuery & " WHERE (ShttCode = cpttShfCode"
    slSQLQuery = slSQLQuery & " AND attCode = cpttAtfCode"
    slSQLQuery = slSQLQuery & " AND cpttAtfCode = " & llAttCode
    slSQLQuery = slSQLQuery & " AND cpttStartDate = '" & Format$(slMoDate, sgSQLDateForm) & "')"
    Set cptt_rst = gSQLSelectCall(slSQLQuery)
    If Not cptt_rst.EOF Then
        ReDim tgCPPosting(0 To 1) As CPPOSTING
        tgCPPosting(0).lCpttCode = cptt_rst!cpttCode
        tgCPPosting(0).iStatus = cptt_rst!cpttStatus
        tgCPPosting(0).iPostingStatus = cptt_rst!cpttPostingStatus
        tgCPPosting(0).lAttCode = cptt_rst!cpttatfCode
        tgCPPosting(0).iAttTimeType = cptt_rst!attTimeType
        tgCPPosting(0).iVefCode = cptt_rst!cpttvefcode
        tgCPPosting(0).iShttCode = cptt_rst!shttCode
        tgCPPosting(0).sZone = cptt_rst!shttTimeZone
        tgCPPosting(0).sDate = Format$(slMoDate, sgShowDateForm)
        tgCPPosting(0).sAstStatus = cptt_rst!cpttAstStatus
        igTimes = 1 'By Week
        ilAdfCode = -1
        'Dan M 9/26/13 6442
        'ilRet = gGetAstInfo(hlAst, tlCPDat(), tlAstInfo(), ilAdfCode, False, False, True, False)
        ilRet = gGetAstInfo(hlAst, tlCPDat(), tlAstInfo(), ilAdfCode, True, True, True, False)
        If UBound(tlAstInfo) > LBound(tlAstInfo) Then
            gBuildAst = ilRet
        Else
            gBuildAst = False
        End If
    Else
        gBuildAst = False
    End If
    On Error Resume Next
    'cptt_rst.Close
    Exit Function
ErrHand:
    On Error Resume Next
    'cptt_rst.Close
    gHandleError "AffErrorLog.txt", "modCPReturns-gBuildAst"
    gBuildAst = False
    Exit Function
End Function

Private Function mGetPledgeByEvent(ilVefCode As Integer) As String
    
    Dim ilVff As Integer
    
    mGetPledgeByEvent = "N"
    If ((Asc(sgSpfSportInfo) And USINGSPORTS) <> USINGSPORTS) Then
        Exit Function
    End If
    If ilVefCode <= 0 Then
        Exit Function
    End If
    ilVff = gBinarySearchVff(ilVefCode)
    If ilVff <> -1 Then
        If Trim$(tgVffInfo(ilVff).sPledgeByEvent) = "" Then
            mGetPledgeByEvent = "N"
        Else
            mGetPledgeByEvent = Trim$(tgVffInfo(ilVff).sPledgeByEvent)
        End If
    End If
End Function

Private Function mGetNotAirStatus(llGsfCode As Long, ilVefCode As Integer, ilShttCode As Integer, ilPledged As Integer) As Integer
    Dim slSQLQuery As String
    '12/6/13: If Not a Game and astCPStatus = 2, then set to 4 not 8.
    '         If game, see if by event, if not set to 4.  if game and by event, check pet to see status.
    '         If not carry then set to 8 otherwise 4.
    'iStatus = 8
    If ilPledged = 2 Then
        mGetNotAirStatus = 8
        Exit Function
    End If
    If (llGsfCode <= 0) Or (mGetPledgeByEvent(ilVefCode) = "N") Then
        mGetNotAirStatus = 4
    Else
        slSQLQuery = "SELECT petClearStatus"
        slSQLQuery = slSQLQuery + " FROM pet"
        slSQLQuery = slSQLQuery & " WHERE (petVefCode = " & ilVefCode & " AND petShttCode = " & ilShttCode & " AND petGsfCode = " & tmLst.lgsfCode & ")"
        Set rst_Pet = gSQLSelectCall(slSQLQuery)
        If Not rst_Pet.EOF Then
            If rst_Pet!petClearStatus = "N" Then
                mGetNotAirStatus = 8
            Else
                mGetNotAirStatus = 4
            End If
        Else
            mGetNotAirStatus = 4
        End If
    End If

End Function

Public Function gBuildAstInfoFromAst(hlAst As Integer, tlCPDat() As DAT, tlAstInfo() As ASTINFO, ilInAdfCode As Integer, blGetRegionCopy As Boolean, llSelGsfCode As Long, blFeedAdjOnReturn As Boolean, blFilterByAirDates As Boolean, blIncludePledgeInfo As Boolean, blCreateServiceATTSpots As Boolean) As Integer
    Dim ilAdfCode As Integer
    Dim ilLoop As Integer
    Dim ilRet As Integer
    Dim ilShttCode As Integer
    Dim ilVefCode As Integer
    Dim slFWkDate As String
    Dim slLWkDate As String
    Dim llFWkDate As Long
    Dim llLWkDate As Long
    Dim ilVef As Integer
    Dim llVpf As Long
    Dim slZone As String
    Dim ilZone As Integer
    Dim ilLocalAdj As Integer
    Dim ilZoneFound As Integer
    Dim ilGsfLoop As Integer
    Dim ilPass As Integer
    Dim ilNumberAsterisk As Integer
    Dim slTemp2 As String
    Dim llAtt As Long
    Dim ilFdIndex As Integer
    Dim ilPdIndex As Integer
    Dim ilAdjIndex As Integer
    'Dim llLogDate As Long
    'Dim llLogTime As Long
    Dim llFeedTimeDiff As Long
    Dim llPledgeTimeDiff As Long
    Dim blPledgeDefined As Boolean
    Dim ilRAdfCode As Integer
    Dim blLivePledge As Boolean
    Dim ilAst As Integer
    Dim ilStartOfBreak As Integer
    Dim ilEndOfBreak As Integer
    Dim blEndFound As Boolean
    Dim llTotalLen As Long
    'Dim ilInner As Integer
    'Dim ilOuter As Integer
    'Dim ilUpper As Integer
    Dim llUpper As Long
    Dim llLst As Long
    Dim slSortDate As String
    Dim slSortTime As String
    Dim slSortBreak As String
    Dim slSortPosition As String
    Dim slSortAirPlay As String
    Dim llDATCode As Long
    Dim llDat As Long
    Dim ilLimit As Integer
    Dim blGetCopy As Boolean
    Dim blAstExist As Boolean
    Dim llCpttSDate As Long
    Dim llCpttEDate As Long
    Dim llFeedDate As Long
    Dim llAirDate As Long
    Dim llBkoutLst As Long
    Dim llGsf As Long
    Dim ilAdjDay As Integer
    '7639 Dan M
    Dim blImportedISCI As Boolean
    '12/10/15: Handle blackouts
    Dim llLstCode As Long
    '3/8/16
    Dim blMGSpot As Boolean
    Dim slMGMissedFeedDate As String
    Dim slMGMissedFeedTime As String
    '8/18/16:
    Dim blExtendExist As Boolean
    '8/31/16: Check if comment suppressed
    Dim ilVff As Integer
    Dim slHideCommOnWeb As String
    '11/8/16: Handle case where AST exist but not for all the days that LST exist
    ReDim llLstDate(0 To 0) As Long
    Dim blDateFd As Boolean
    Dim llDate As Long
    Dim slSQLQuery As String
    Dim blAdvtBkout As Boolean

    On Error GoTo ErrHand
    imAstRecLen = Len(tmAst)
    ilAdfCode = 0
    lgSTime8 = timeGetTime

    ReDim tlAstInfo(0 To 5000) As ASTINFO
    ilAst = 0
    blAstExist = False
    blExtendExist = False
    '11/15/11: Build copy if not alread built
    If blGetRegionCopy Then
        blGetCopy = False
        If lmCopyDate <= 0 Then
            blGetCopy = True
        Else
            If ((igTimes = 3) Or (igTimes = 4)) And (blFeedAdjOnReturn) Then
                If gDateValue(DateAdd("d", -1, gObtainPrevMonday(tgCPPosting(0).sDate))) < lmCopyDate Then
                    blGetCopy = True
                End If
            Else
                If gDateValue(gObtainPrevMonday(tgCPPosting(0).sDate)) < lmCopyDate Then
                    blGetCopy = True
                End If
            End If
        End If
        ilRet = 0
        'On Error GoTo CheckForCopyErr:
        'ilLimit = LBound(tgCifCpfInfo1)
        If PeekArray(tgCifCpfInfo1).Ptr <> 0 Then
            ilLimit = LBound(tgCifCpfInfo1)
            If (UBound(tgCifCpfInfo1) <= LBound(tgCifCpfInfo1)) Then
                ilRet = 1
            End If
        Else
            ilRet = 1
            ilLimit = 0
        End If
        'If (ilRet = 1) Or (UBound(tgCifCpfInfo1) <= LBound(tgCifCpfInfo1)) Or blGetCopy Then
        If (ilRet = 1) Or blGetCopy Then
            If ((igTimes = 3) Or (igTimes = 4)) And (blFeedAdjOnReturn) Then
                lmCopyDate = gDateValue(DateAdd("d", -1, gObtainPrevMonday(tgCPPosting(0).sDate)))
                ilRet = gPopCopy(Format$(DateAdd("d", -1, gObtainPrevMonday(tgCPPosting(0).sDate)), sgShowDateForm), "gGetAstInfo")
            Else
                lmCopyDate = gDateValue(gObtainPrevMonday(tgCPPosting(0).sDate))
                ilRet = gPopCopy(Format$(gObtainPrevMonday(tgCPPosting(0).sDate), sgShowDateForm), "gGetAstInfo")
            End If
        End If
    End If
    On Error GoTo ErrHand
    
    For ilLoop = 0 To UBound(tgCPPosting) - 1 Step 1 ' This loop is always 0 to 0.
        ilShttCode = tgCPPosting(ilLoop).iShttCode
        ilVefCode = tgCPPosting(ilLoop).iVefCode
        llCpttSDate = DateValue(tgCPPosting(ilLoop).sDate)
        llCpttEDate = llCpttSDate + 6
        'If (igTimes = 3) Or (igTimes = 4) Then
        '    If blFeedAdjOnReturn Then
        '        slFWkDate = Format$(DateAdd("d", -1, tgCPPosting(ilLoop).sDate), sgShowDateForm)
        '    Else
        '        slFWkDate = Format$(tgCPPosting(ilLoop).sDate, sgShowDateForm)
        '    End If
        'Else
        '    slFWkDate = Format$(gObtainPrevMonday(tgCPPosting(ilLoop).sDate), sgShowDateForm)
        'End If
        'If igTimes = 0 Then
        '    slLWkDate = Format$(gObtainEndStd(tgCPPosting(ilLoop).sDate), sgShowDateForm)
        'ElseIf (igTimes = 3) Or (igTimes = 4) Then
        '    If blFeedAdjOnReturn Then
        '        slLWkDate = Format$(DateAdd("d", tgCPPosting(ilLoop).iNumberDays, tgCPPosting(ilLoop).sDate), sgShowDateForm)
        '    Else
        '        slLWkDate = Format$(DateAdd("d", tgCPPosting(ilLoop).iNumberDays - 1, tgCPPosting(ilLoop).sDate), sgShowDateForm)
        '    End If
        'Else
        '    slLWkDate = Format$(gObtainNextSunday(tgCPPosting(ilLoop).sDate), sgShowDateForm)
        'End If
        mGetAstDateRange tgCPPosting(ilLoop), blFeedAdjOnReturn, slFWkDate, slLWkDate
        llFWkDate = DateValue(gAdjYear(slFWkDate))
        llLWkDate = DateValue(gAdjYear(slLWkDate))
        blAdvtBkout = mAdvtBkout(ilInAdfCode, ilVefCode, llFWkDate, llLWkDate, blFeedAdjOnReturn)
        If Not blAdvtBkout Then
            ilAdfCode = ilInAdfCode
        End If
        llVpf = gBinarySearchVpf(CLng(tgCPPosting(ilLoop).iVefCode))
        If llVpf <> -1 Then
            smDefaultEmbeddedOrROS = tgVpfOptions(llVpf).sEmbeddedOrROS
        End If
        If Trim$(smDefaultEmbeddedOrROS) = "" Then
            smDefaultEmbeddedOrROS = "R"
        End If
        'slZone = tgCPPosting(ilLoop).sZone
        'ilLocalAdj = 0
        'ilZoneFound = False
        'ilNumberAsterisk = 0
        ilVef = gBinarySearchVef(CLng(ilVefCode))
        '8/31/16: Check if comment suppressed
        slHideCommOnWeb = "N"
        ilVff = gBinarySearchVff(CLng(ilVefCode))
        If ilVff <> -1 Then
            slHideCommOnWeb = tgVffInfo(ilVff).sHideCommOnWeb
        End If
        ' Adjust time zone properly.
        'If Len(Trim$(tgCPPosting(ilLoop).sZone)) <> 0 Then
        '    'Get zone
        '    If ilVef <> -1 Then
        '        For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
        '            'If igExportSource = 2 Then
        '                DoEvents
        '            'End If
        '            If Trim$(tgVehicleInfo(ilVef).sZone(ilZone)) = Trim$(tgCPPosting(ilLoop).sZone) Then
        '                If (tgVehicleInfo(ilVef).sFed(ilZone) <> "*") And (Trim$(tgVehicleInfo(ilVef).sFed(ilZone)) <> "") And (tgVehicleInfo(ilVef).iBaseZone(ilZone) <> -1) Then
        '                    slZone = tgVehicleInfo(ilVef).sZone(tgVehicleInfo(ilVef).iBaseZone(ilZone))
        '                    ilLocalAdj = tgVehicleInfo(ilVef).iLocalAdj(ilZone)
        '                    ilZoneFound = True
        '                End If
        '                Exit For
        '            End If
        '        Next ilZone
        '        For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
        '            If igExportSource = 2 Then
        '                DoEvents
        '            End If
        '            If tgVehicleInfo(ilVef).sFed(ilZone) = "*" Then
        '                ilNumberAsterisk = ilNumberAsterisk + 1
        '            End If
        '        Next ilZone
        '    End If
        'End If
        'If (Not ilZoneFound) And (ilNumberAsterisk <= 1) Then
        '    slZone = ""
        'End If
        
        '4/7/19: lLocalAdj required they the exclusion routine
        slZone = tgCPPosting(ilLoop).sZone
        ilLocalAdj = 0
        ilZoneFound = False
        ilNumberAsterisk = 0
        ilVef = gBinarySearchVef(CLng(ilVefCode))
        
        
        If (Len(Trim$(tgCPPosting(ilLoop).sZone)) <> 0) Then
            'Get zone
            If ilVef <> -1 Then
                For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                    'If igExportSource = 2 Then
                        DoEvents
                    'End If
                    If Trim$(tgVehicleInfo(ilVef).sZone(ilZone)) = Trim$(tgCPPosting(ilLoop).sZone) Then
                        If (tgVehicleInfo(ilVef).sFed(ilZone) <> "*") And (Trim$(tgVehicleInfo(ilVef).sFed(ilZone)) <> "") And (tgVehicleInfo(ilVef).iBaseZone(ilZone) <> -1) Then
                            slZone = tgVehicleInfo(ilVef).sZone(tgVehicleInfo(ilVef).iBaseZone(ilZone))
                            ilLocalAdj = tgVehicleInfo(ilVef).iLocalAdj(ilZone)
                            ilZoneFound = True
                        End If
                        Exit For
                    End If
                Next ilZone
                For ilZone = LBound(tgVehicleInfo(ilVef).sZone) To UBound(tgVehicleInfo(ilVef).sZone) Step 1
                    If igExportSource = 2 Then
                        DoEvents
                    End If
                    If tgVehicleInfo(ilVef).sFed(ilZone) = "*" Then
                        ilNumberAsterisk = ilNumberAsterisk + 1
                    End If
                Next ilZone
            End If
        End If
        If (Not ilZoneFound) And (ilNumberAsterisk <= 1) Then
            slZone = ""
        End If
        
        '4/7/19: Moved here from below
        ReDim tmATTCrossDates(0 To 0) As ATTCrossDates
        bgAnyAttExclusions = False
        llUpper = 0
        slSQLQuery = "SELECT *"
        slSQLQuery = slSQLQuery + " FROM att "
        slSQLQuery = slSQLQuery + " WHERE (attShfCode= " & Trim$(Str(ilShttCode)) & " And attVefCode = " & Trim$(Str(ilVefCode)) & ")"
        slSQLQuery = slSQLQuery + " Order by attOnAir"
        Set att_rst = gSQLSelectCall(slSQLQuery)
        While Not att_rst.EOF
            If igExportSource = 2 Then
                DoEvents
            End If
            If ((blCreateServiceATTSpots) Or (att_rst!attServiceAgreement <> "Y")) Then
                llUpper = UBound(tmATTCrossDates)
                If DateValue(gAdjYear(Trim$(att_rst!attOffAir))) <= DateValue(gAdjYear(Trim$(att_rst!attDropDate))) Then
                    slTemp2 = Trim$(att_rst!attOffAir)
                Else
                    slTemp2 = Trim$(att_rst!attDropDate)
                End If
                tmATTCrossDates(llUpper).lAttCode = att_rst!attCode
                tmATTCrossDates(llUpper).lStartDate = DateValue(gAdjYear(att_rst!attOnAir))
                tmATTCrossDates(llUpper).lEndDate = DateValue(gAdjYear(slTemp2))
                tmATTCrossDates(llUpper).iLoadFactor = att_rst!attLoad
                If tmATTCrossDates(llUpper).iLoadFactor < 1 Then
                    tmATTCrossDates(llUpper).iLoadFactor = 1
                End If
                tmATTCrossDates(llUpper).sForbidSplitLive = att_rst!attForbidSplitLive
                tmATTCrossDates(llUpper).iDACode = 1
                If att_rst!attPledgeType = "D" Then
                    tmATTCrossDates(llUpper).iDACode = 0
                ElseIf att_rst!attPledgeType = "A" Then
                    tmATTCrossDates(llUpper).iDACode = 1
                ElseIf att_rst!attPledgeType = "C" Then
                    tmATTCrossDates(llUpper).iDACode = 2
                End If
                '4/29/19: moved up
                'If ((Not blCreateServiceATTSpots) And (att_rst!attServiceAgreement = "Y")) Then
                '    ReDim tlAstInfo(0 To 0) As ASTINFO
                '    gBuildAstInfoFromAst = True
                '    Exit Function
                ''End If
                '4/3/19
                tmATTCrossDates(llUpper).sExcludeFillSpot = att_rst!attExcludeFillSpot
                tmATTCrossDates(llUpper).sExcludeCntrTypeQ = att_rst!attExcludeCntrTypeQ
                tmATTCrossDates(llUpper).sExcludeCntrTypeR = att_rst!attExcludeCntrTypeR
                tmATTCrossDates(llUpper).sExcludeCntrTypeT = att_rst!attExcludeCntrTypeT
                tmATTCrossDates(llUpper).sExcludeCntrTypeM = att_rst!attExcludeCntrTypeM
                tmATTCrossDates(llUpper).sExcludeCntrTypeS = att_rst!attExcludeCntrTypeS
                tmATTCrossDates(llUpper).sExcludeCntrTypeV = att_rst!attExcludeCntrTypeV
                
                If (igTimes = 3) Or (igTimes = 4) Then
                    If (tmATTCrossDates(llUpper).lEndDate >= llFWkDate) And (tmATTCrossDates(llUpper).lStartDate <= llLWkDate) Then
                        mSetAnyAttExclusions tmATTCrossDates(llUpper)
                        ReDim Preserve tmATTCrossDates(0 To llUpper + 1) As ATTCrossDates
                    End If
                Else
                    If (tmATTCrossDates(llUpper).lEndDate >= llFWkDate - 1) And (tmATTCrossDates(llUpper).lStartDate <= llLWkDate + 1) Then
                        mSetAnyAttExclusions tmATTCrossDates(llUpper)
                        ReDim Preserve tmATTCrossDates(0 To llUpper + 1) As ATTCrossDates
                    End If
                End If
            End If
            att_rst.MoveNext
        Wend
        
        '4/29/19: Added to handle no agreement
        If (UBound(tmATTCrossDates) <= LBound(tmATTCrossDates)) Then
            ReDim tlAstInfo(0 To 0) As ASTINFO
            gBuildAstInfoFromAst = True
            Exit Function
        End If
        
        If igTimes = 0 Then
            'slSQLQuery = "SELECT *"
            'slSQLQuery = slSQLQuery + " FROM lst, ADF_Advertisers, "
            'slSQLQuery = slSQLQuery & "VEF_Vehicles"
            'slSQLQuery = slSQLQuery + " WHERE (adfCode = lstAdfCode"
            'If (iAdfCode > 0) Then
            '    slSQLQuery = slSQLQuery & " AND lstAdfCode = " & iAdfCode
            'End If
            'If llSelGsfCode > 0 Then
            '    slSQLQuery = slSQLQuery & " AND lstGsfCode = " & llSelGsfCode
            'End If
            'slSQLQuery = slSQLQuery + " AND vefCode = lstLogVefCode"
            'slSQLQuery = slSQLQuery + " AND lstLogVefCode = " & tgCPPosting(iLoop).iVefCode
            ''slSQLQuery = slSQLQuery + " AND lstBkoutLstCode = 0"
            'slSQLQuery = slSQLQuery & " AND lstType <> 1"
            ''If Trim$(sZone) <> "" Then
            ''    slSQLQuery = slSQLQuery + " AND lstZone = '" & sZone & "'"
            ''End If
            slSQLQuery = "SELECT * FROM lst"
            slSQLQuery = slSQLQuery + " WHERE (lstLogVefCode = " & tgCPPosting(ilLoop).iVefCode
            slSQLQuery = slSQLQuery & " AND lstType <> 1"
            slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(llFWkDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(llLWkDate + 1, sgSQLDateForm) & "')" & ")"
            'Dan M removed per Dick
            slSQLQuery = slSQLQuery + " ORDER BY lstLogDate, lstLogTime"
            'slSQLQuery = slSQLQuery + " ORDER BY vefName, adfName, lstLogDate, lstLogTime"

        Else
            'slSQLQuery = "SELECT lstProd, lstLogDate, lstLogTime, lstSdfCode, lstStatus, lstAdfCode, lstLogVefCode, lstType, lstZone, lstLen, lstCntrNo, lstCode, lstSplitNetwork, lstRafCode FROM lst"
            slSQLQuery = "SELECT * FROM lst"
            slSQLQuery = slSQLQuery + " WHERE (lstLogVefCode = " & tgCPPosting(ilLoop).iVefCode
            'If (iAdfCode > 0) Then
            '    slSQLQuery = slSQLQuery & " AND lstAdfCode = " & iAdfCode
            'End If
            'If llSelGsfCode > 0 Then
            '    slSQLQuery = slSQLQuery & " AND lstGsfCode = " & llSelGsfCode
            'End If
            'slSQLQuery = slSQLQuery + " AND lstBkoutLstCode = 0"
            slSQLQuery = slSQLQuery & " AND lstType <> 1"
            'If Trim$(sZone) <> "" Then
            '    slSQLQuery = slSQLQuery + " AND lstZone = '" & sZone & "'"
            'End If
            If ((igTimes = 3) Or (igTimes = 4)) And (blFeedAdjOnReturn) Then
                'lFWkDate and lLWkDate has been adjusted because sFWkDate and sLWkDate have been adjusted by one day
                slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(llFWkDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(llLWkDate, sgSQLDateForm) & "')" & ")"
            Else
                slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(llFWkDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(llLWkDate + 1, sgSQLDateForm) & "')" & ")"
            End If
            slSQLQuery = slSQLQuery + " ORDER BY lstZone, lstGsfCode, lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
        End If
        If igExportSource = 2 Then
            DoEvents
        End If
        '5/3/19: If using extlusions, get lst each time
        If StrComp(smBuildLstKey, slSQLQuery, vbTextCompare) <> 0 Or bgAnyAttExclusions Then
            bmImportSpot = False
            lgSTime7 = timeGetTime
            smBuildLstKey = slSQLQuery
            ReDim tmBuildLst(0 To 100) As BUILDLST
            llUpper = 0
            Set lst_rst = gSQLSelectCall(slSQLQuery)
            Do While Not lst_rst.EOF
                If mIncludeLst_ExclusionCheck(lst_rst, ilLocalAdj) Then
                    gCreateUDTforLST lst_rst, tmBuildLst(llUpper).tLST
                    tmBuildLst(llUpper).iBreakLength = 300
                    tmBuildLst(llUpper).bIgnore = False
                    '4/13/18: Added MG also set the SpotTYpe = 5 but ImportSpot is N
                    'If tmBuildLst(llUpper).tLST.iSpotType = 5 Then
                    If (tmBuildLst(llUpper).tLST.iSpotType = 5) And (tmBuildLst(llUpper).tLST.sImportedSpot = "Y") Then
                        bmImportSpot = True
                    End If
                    llUpper = llUpper + 1
                    If llUpper >= UBound(tmBuildLst) Then
                        ReDim Preserve tmBuildLst(0 To UBound(tmBuildLst) + 100) As BUILDLST
                    End If
                End If
                lst_rst.MoveNext
            Loop
            ReDim Preserve tmBuildLst(0 To llUpper) As BUILDLST
'setting BreakLength as 5min (300 sec)
'            'remove those Lst that have a blackout replacement
'            ilStartOfBreak = 0
'            Do While ilStartOfBreak < UBound(tmBuildLst)
'                'Find end of break
'                ilEndOfBreak = ilStartOfBreak + 1
'                blEndFound = False
'                Do
'                    If tmBuildLst(ilStartOfBreak).tLST.sZone <> tmBuildLst(ilEndOfBreak).tLST.sZone Then
'                        blEndFound = True
'                    End If
'                    If tmBuildLst(ilStartOfBreak).tLST.lGsfCode <> tmBuildLst(ilEndOfBreak).tLST.lGsfCode Then
'                        blEndFound = True
'                    End If
'                    If gDateValue(tmBuildLst(ilStartOfBreak).tLST.sLogDate) <> gDateValue(tmBuildLst(ilEndOfBreak).tLST.sLogDate) Then
'                        blEndFound = True
'                    End If
'                    If gTimeToLong(tmBuildLst(ilStartOfBreak).tLST.sLogTime, False) <> gTimeToLong(tmBuildLst(ilEndOfBreak).tLST.sLogTime, False) Then
'                        blEndFound = True
'                    End If
'                    If tmBuildLst(ilStartOfBreak).tLST.iBreakNo <> tmBuildLst(ilEndOfBreak).tLST.iBreakNo Then
'                        blEndFound = True
'                    End If
'                    If blEndFound Then
'                        'For ilOuter = ilStartOfBreak To ilEndOfBreak - 1 Step 1
'                        '    If tmBuildLst(ilLoop).tLST.lBkoutLstCode > 0 Then
'                        '        For ilInner = ilOuter To ilEndOfBreak - 1 Step 1
'                        '            If tmBuildLst(ilOuter).tLST.lBkoutLstCode = tmBuildLst(ilInner).tLST.lCode Then
'                        '                tmBuildLst(ilEndOfBreak).bIgnore = True
'                        '            End If
'                        '        Next ilInner
'                        '    End If
'                        'Next ilOuter
'                        llTotalLen = 0
'                        For ilOuter = ilStartOfBreak To ilEndOfBreak - 1 Step 1
'                            'Bypass Blackouts, MG, Replacements and Bonus spots
'                            If (tmBuildLst(ilOuter).tLST.lBkoutLstCode = 0) And (tmBuildLst(ilOuter).tLST.lCntrNo > 0) Then
'                                llTotalLen = llTotalLen + tmBuildLst(ilOuter).tLST.iLen
'                            End If
'                        Next ilOuter
'                        If llTotalLen = 0 Then
'                            llTotalLen = tmBuildLst(ilStartOfBreak).tLST.iLen
'                        End If
'                        For ilOuter = ilStartOfBreak To ilEndOfBreak - 1 Step 1
'                            tmBuildLst(ilOuter).iBreakLength = llTotalLen
'                        Next ilOuter
'                        ilStartOfBreak = ilEndOfBreak
'                    Else
'                        ilEndOfBreak = ilEndOfBreak + 1
'                    End If
'                Loop While Not blEndFound
'            Loop
            If UBound(tmBuildLst) - 1 >= 1 Then
                ArraySortTyp fnAV(tmBuildLst(), 0), UBound(tmBuildLst), 0, LenB(tmBuildLst(0)), 0, -2, 0
            End If
            lgETime7 = timeGetTime
            lgTtlTime7 = lgTtlTime7 + (lgETime7 - lgSTime7)
        End If
        '11/8/16: Retain all unique LST dates to verify if all ast created
        For llLst = 0 To UBound(tmBuildLst) - 1 Step 1
            '7/8/20: only create dates for days that advertiser airs on
            If (ilInAdfCode <= 0) Or (ilInAdfCode = tmBuildLst(llLst).tLST.iAdfCode) Then
                blDateFd = False
                llDate = gDateValue(tmBuildLst(llLst).tLST.sLogDate)
                '4/7/18: Bypass Extended status
                'If (llDate >= llFWkDate) And (llDate <= llLWkDate) Then
                If (llDate >= llFWkDate) And (llDate <= llLWkDate) And (tmBuildLst(llLst).tLST.iStatus <= 10) Then
                    For llDat = 0 To UBound(llLstDate) - 1 Step 1
                        If llLstDate(llDat) = llDate Then
                            blDateFd = True
                            Exit For
                        End If
                    Next llDat
                    If Not blDateFd Then
                        llLstDate(UBound(llLstDate)) = llDate
                        ReDim Preserve llLstDate(0 To UBound(llLstDate) + 1) As Long
                    End If
                End If
            End If
        Next llLst
        ReDim llGsfCode(0 To 0) As Long
        If ilVef <> -1 Then
            If tgVehicleInfo(ilVef).sVehType = "G" Then
                slSQLQuery = "SELECT * FROM GSF_Game_Schd WHERE (gsfVefCode = " & ilVefCode & " AND gsfAirDate >= '" & Format$(slFWkDate, sgSQLDateForm) & "'" & " AND gsfAirDate <= '" & Format$(slLWkDate, sgSQLDateForm) & "'" & ")"
                Set rst_Genl = gSQLSelectCall(slSQLQuery)
                Do While Not rst_Genl.EOF
                    llGsfCode(UBound(llGsfCode)) = rst_Genl!gsfCode
                    ReDim Preserve llGsfCode(0 To UBound(llGsfCode) + 1) As Long
                    rst_Genl.MoveNext
                Loop
            
            End If
        End If
        

        lgSTime11 = timeGetTime
        
        '4/7/19: Moved above building LST so that the lst can be filtered by exclusions
        'ReDim tmATTCrossDates(0 To 0) As ATTCrossDates
        'bgAnyAttExclusions = False
        'llUpper = 0
        'slSQLQuery = "SELECT *"
        'slSQLQuery = slSQLQuery + " FROM att "
        'slSQLQuery = slSQLQuery + " WHERE (attShfCode= " & Trim$(Str(ilShttCode)) & " And attVefCode = " & Trim$(Str(ilVefCode)) & ")"
        'slSQLQuery = slSQLQuery + " Order by attOnAir"
        'Set att_rst = gSQLSelectCall(slSQLQuery)
        'While Not att_rst.EOF
        '    If igExportSource = 2 Then
        '        DoEvents
        '    End If
        '    llUpper = UBound(tmATTCrossDates)
        '    If DateValue(gAdjYear(Trim$(att_rst!attOffAir))) <= DateValue(gAdjYear(Trim$(att_rst!attDropDate))) Then
        '        slTemp2 = Trim$(att_rst!attOffAir)
        '    Else
        '        slTemp2 = Trim$(att_rst!attDropDate)
        '    End If
        '    tmATTCrossDates(llUpper).lAttCode = att_rst!attCode
        '    tmATTCrossDates(llUpper).lStartDate = DateValue(gAdjYear(att_rst!attOnAir))
        '    tmATTCrossDates(llUpper).lEndDate = DateValue(gAdjYear(slTemp2))
        '    tmATTCrossDates(llUpper).iLoadFactor = att_rst!attLoad
        '    If tmATTCrossDates(llUpper).iLoadFactor < 1 Then
        '        tmATTCrossDates(llUpper).iLoadFactor = 1
        '    End If
        '    tmATTCrossDates(llUpper).sForbidSplitLive = att_rst!attForbidSplitLive
        '    tmATTCrossDates(llUpper).iDACode = 1
        '    If att_rst!attPledgeType = "D" Then
        '        tmATTCrossDates(llUpper).iDACode = 0
        '    ElseIf att_rst!attPledgeType = "A" Then
        '        tmATTCrossDates(llUpper).iDACode = 1
        '    ElseIf att_rst!attPledgeType = "C" Then
        '        tmATTCrossDates(llUpper).iDACode = 2
        '    End If
        '    If ((Not blCreateServiceATTSpots) And (att_rst!attServiceAgreement = "Y")) Then
        '        ReDim tlAstInfo(0 To 0) As ASTINFO
        '        gBuildAstInfoFromAst = True
        '        Exit Function
        '    End If
        '    '4/3/19
        '    tmATTCrossDates(llUpper).sExcludeFillSpot = att_rst!attExcludeFillSpot
        '    tmATTCrossDates(llUpper).sExcludeCntrTypeQ = att_rst!attExcludeCntrTypeQ
        '    tmATTCrossDates(llUpper).sExcludeCntrTypeR = att_rst!attExcludeCntrTypeR
        '    tmATTCrossDates(llUpper).sExcludeCntrTypeT = att_rst!attExcludeCntrTypeT
        '    tmATTCrossDates(llUpper).sExcludeCntrTypeM = att_rst!attExcludeCntrTypeM
        '    tmATTCrossDates(llUpper).sExcludeCntrTypeS = att_rst!attExcludeCntrTypeS
        '    tmATTCrossDates(llUpper).sExcludeCntrTypeV = att_rst!attExcludeCntrTypeV
        '
        '    If (igTimes = 3) Or (igTimes = 4) Then
        '        If (tmATTCrossDates(llUpper).lEndDate >= llFWkDate) And (tmATTCrossDates(llUpper).lStartDate <= llLWkDate) Then
        '            mSetAnyAttExclusions tmATTCrossDates(llUpper)
        '            ReDim Preserve tmATTCrossDates(0 To llUpper + 1) As ATTCrossDates
        '        End If
        '    Else
        '        If (tmATTCrossDates(llUpper).lEndDate >= llFWkDate - 1) And (tmATTCrossDates(llUpper).lStartDate <= llLWkDate + 1) Then
        '            mSetAnyAttExclusions tmATTCrossDates(llUpper)
        '            ReDim Preserve tmATTCrossDates(0 To llUpper + 1) As ATTCrossDates
        '        End If
        '    End If
        '    att_rst.MoveNext
        'Wend
        lgETime11 = timeGetTime
        lgTtlTime11 = lgTtlTime11 + (lgETime11 - lgSTime11)
        
        'att_rst.Close
        'Set att_rst = Nothing
        
        'ReDim tlAstInfo(0 To 5000) As ASTINFO
        'llUpper = 0
        blPledgeDefined = False
        blLivePledge = False
        For ilPass = 0 To UBound(tmATTCrossDates) - 1 Step 1
            lgSTime12 = timeGetTime
            If igExportSource = 2 Then
                DoEvents
            End If
            llAtt = tmATTCrossDates(ilPass).lAttCode
            If lmAttForDat <> llAtt Then
                If blIncludePledgeInfo Then
                    lmAttForDat = llAtt
                    llDat = 0
                    ReDim tmDatRst(0 To 100) As DATRST
                    slSQLQuery = "SELECT * FROM dat WHERE"
                    slSQLQuery = slSQLQuery & " datAtfCode = " & llAtt
                    Set dat_rst = gSQLSelectCall(slSQLQuery)
                    Do While Not dat_rst.EOF
                        gCreateUDTForDat dat_rst, tmDatRst(llDat)
                        llDat = llDat + 1
                        If llDat >= UBound(tmDatRst) Then
                            ReDim Preserve tmDatRst(0 To UBound(tmDatRst) + 30) As DATRST
                        End If
                        dat_rst.MoveNext
                    Loop
                    ReDim Preserve tmDatRst(0 To llDat) As DATRST
                    If UBound(tmDatRst) - 1 >= 1 Then
                        ArraySortTyp fnAV(tmDatRst(), 0), UBound(tmDatRst), 0, LenB(tmDatRst(0)), 0, -2, 0
                    End If
                Else
                    ReDim Preserve tmDatRst(0 To 0) As DATRST
                End If
            End If
            If igTimes = 0 Then
                ' Retrieve an month for an advertiser.
                slSQLQuery = "SELECT astCode,astAtfCode,astShfCode,astVefCode,astSdfCode,astLsfCode,astAirDate,astAirTime,astStatus,astCPStatus,astFeedDate,astFeedTime,astAdfCode,astDatCode,astCpfCode,astRsfCode,astStationCompliant,astAgencyCompliant,astAffidavitSource,astCntrNo,astLen,astLkAstCode,astMissedMnfCode,astUstCode,rsfCode,rsfSdfCode,rstPtType,rsfCopyCode,rsfRotNo,rsfRafCode,rsfSBofCode,rsfRBofCode,rsfType,rsfBVefCode,rsfRChfCode,rsfCrfCode"
                'slSQLQuery = slSQLQuery + " FROM  ast Inner Join lst on astlsfCode = lstCode"
                'slSQLQuery = slSQLQuery & " Left Outer Join VEF_Vehicles On lstLogVefCode = vefCode"
                slSQLQuery = slSQLQuery + " FROM  ast"
                slSQLQuery = slSQLQuery + " left outer join rsf_Region_Schd_Copy on astRsfCode = rsfCode"
                slSQLQuery = slSQLQuery + " WHERE (astatfCode = " & llAtt
                'D.L. 03/07/14
                'slSQLQuery = slSQLQuery + " AND Mod(astStatus,100) <> 22"    '22=Missed part of MG (Status 20)
                slSQLQuery = slSQLQuery + " AND (astFeedDate >= '" & Format$(slFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slLWkDate, sgSQLDateForm) & "')"
                If (ilAdfCode > 0) Then
                    slSQLQuery = slSQLQuery & " AND astAdfCode = " & ilAdfCode
                End If
                'If llSelGsfCode > 0 Then
                '    slSQLQuery = slSQLQuery & " AND lstGsfCode = " & llSelGsfCode
                'End If
                ''slSQLQuery = slSQLQuery + " AND vefCode = lstLogVefCode" & ")"
                slSQLQuery = slSQLQuery + ")"
                'slSQLQuery = slSQLQuery + " ORDER BY vefName, adfName, astAirDate, astAirTime"
                slSQLQuery = slSQLQuery + " ORDER BY astAirDate, astAirTime, astCode"
            Else
                ' Retrieve a weeks worth.
                If llSelGsfCode <= 0 Then
                    slSQLQuery = "SELECT astCode,astAtfCode,astShfCode,astVefCode,astSdfCode,astLsfCode,astAirDate,astAirTime,astStatus,astCPStatus,astFeedDate,astFeedTime,astAdfCode,astDatCode,astCpfCode,astRsfCode,astStationCompliant,astAgencyCompliant,astAffidavitSource,astCntrNo,astLen,astLkAstCode,astMissedMnfCode,astUstCode,rsfCode,rsfSdfCode,rstPtType,rsfCopyCode,rsfRotNo,rsfRafCode,rsfSBofCode,rsfRBofCode,rsfType,rsfBVefCode,rsfRChfCode,rsfCrfCode FROM ast"
                    slSQLQuery = slSQLQuery + " left outer join rsf_Region_Schd_Copy on astRsfCode = rsfCode"
                    slSQLQuery = slSQLQuery + " WHERE (astatfCode= " & llAtt
                    'D.L. 03/07/14
                    'slSQLQuery = slSQLQuery + " AND Mod(astStatus, 100) <> 22"    '22=Missed part of MG (Status 20)
                    If blFilterByAirDates Then
                        slSQLQuery = slSQLQuery + " AND (astAirDate >= '" & Format$(slFWkDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(slLWkDate, sgSQLDateForm) & "')" & ")"
                    Else
                        slSQLQuery = slSQLQuery + " AND (astFeedDate >= '" & Format$(slFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slLWkDate, sgSQLDateForm) & "')" & ")"
                    End If
                    If (ilAdfCode > 0) Then
                        slSQLQuery = slSQLQuery & " AND astAdfCode = " & ilAdfCode
                    End If
                    If igTimes <> 2 Then
                        'slSQLQuery = slSQLQuery + " ORDER BY astAirDate, astAirTime, lstBreakNo, lstPositionNo"
                        slSQLQuery = slSQLQuery + " ORDER BY astAirDate, astAirTime, astCode"
                    Else
                        'slSQLQuery = slSQLQuery + " ORDER BY astFeedDate, astFeedTime, lstBreakNo, lstPositionNo"
                        If tmATTCrossDates(ilPass).iLoadFactor <= 1 Then
                            slSQLQuery = slSQLQuery + " ORDER BY astFeedDate, astFeedTime, astCode"
                        Else
                            slSQLQuery = slSQLQuery + " ORDER BY astFeedDate, astFeedTime, astAdfCode, astCode"
                        End If
                    End If
                Else
                    slSQLQuery = "SELECT astCode,astAtfCode,astShfCode,astVefCode,astSdfCode,astLsfCode,astAirDate,astAirTime,astStatus,astCPStatus,astFeedDate,astFeedTime,astAdfCode,astDatCode,astCpfCode,astRsfCode,astStationCompliant,astAgencyCompliant,astAffidavitSource,astCntrNo,astLen,astLkAstCode,astMissedMnfCode,astUstCode,rsfCode,rsfSdfCode,rstPtType,rsfCopyCode,rsfRotNo,rsfRafCode,rsfSBofCode,rsfRBofCode,rsfType,rsfBVefCode,rsfRChfCode,rsfCrfCode"
                    slSQLQuery = slSQLQuery + " FROM ast"
                    slSQLQuery = slSQLQuery + " left outer join rsf_Region_Schd_Copy on astRsfCode = rsfCode"
                    slSQLQuery = slSQLQuery + " WHERE (astatfCode= " & llAtt
                    'D.L. 03/07/14
                    'slSQLQuery = slSQLQuery + " AND Mod(astStatus, 100) <> 22"    '22=Missed part of MG (Status 20)
                    If blFilterByAirDates Then
                        slSQLQuery = slSQLQuery + " AND (astAirDate >= '" & Format$(slFWkDate, sgSQLDateForm) & "' AND astAirDate <= '" & Format$(slLWkDate, sgSQLDateForm) & "')" & ")"
                    Else
                        slSQLQuery = slSQLQuery + " AND (astFeedDate >= '" & Format$(slFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(slLWkDate, sgSQLDateForm) & "')" & ")"
                    End If
                    If (ilAdfCode > 0) Then
                        slSQLQuery = slSQLQuery & " AND astAdfCode = " & ilAdfCode
                    End If
                    If (igTimes <> 2) And (igTimes <> 4) Then
                        'slSQLQuery = slSQLQuery + " ORDER BY astAirDate, astAirTime, lstBreakNo, lstPositionNo"
                        slSQLQuery = slSQLQuery + " ORDER BY astAirDate, astAirTime, astCode"
                    Else
                        'slSQLQuery = slSQLQuery + " ORDER BY astFeedDate, astFeedTime, lstBreakNo, lstPositionNo"
                        If tmATTCrossDates(ilPass).iLoadFactor <= 1 Then
                            slSQLQuery = slSQLQuery + " ORDER BY astFeedDate, astFeedTime, astCode"
                        Else
                            slSQLQuery = slSQLQuery + " ORDER BY astFeedDate, astFeedTime, astAdfCode, astCode"
                        End If
                    End If
                End If
            End If
            
            Set rst_Genl = gSQLSelectCall(slSQLQuery)
            lgETime12 = timeGetTime
            lgTtlTime12 = lgTtlTime12 + (lgETime12 - lgSTime12)
            If rst_Genl.EOF Then
                '5/29/15: Check only the agreement that was set to Complete
                ''Handle the case where ast not created but lst exist and the complete flag set
                If tgCPPosting(0).lAttCode = llAtt Then
                    gBuildAstInfoFromAst = False
                    Exit Function
                End If
            End If
            Do While Not rst_Genl.EOF
                lgSTime14 = timeGetTime
                
                '11/8/16: clear lst date if ast created
                llDate = DateValue(Format$(rst_Genl!astFeedDate, sgShowDateForm))
                For llLst = 0 To UBound(llLstDate) - 1 Step 1
                    If Abs(llLstDate(llLst)) = llDate Then
                        llLstDate(llLst) = -Abs(llLstDate(llLst))
                        Exit For
                    End If
                Next llLst
                
                llLst = mBinarySearchBuildLst(rst_Genl!astLsfCode)
                If llLst = -1 Then
                    slSQLQuery = "SELECT * FROM lst WHERE (lstCode = " & rst_Genl!astLsfCode & ")"
                    Set lst_rst = gSQLSelectCall(slSQLQuery)
                    If Not lst_rst.EOF Then
                        If mIncludeLst_ExclusionCheck(lst_rst, ilLocalAdj) Then
                            llLst = UBound(tmBuildLst)
                            gCreateUDTforLST lst_rst, tmBuildLst(llLst).tLST
                            tmBuildLst(llLst).iBreakLength = 300    'tmBuildLst(llLst).tLST.iLen
                            tmBuildLst(llLst).bIgnore = False
                            ReDim Preserve tmBuildLst(0 To llLst + 1) As BUILDLST
                            If UBound(tmBuildLst) - 1 >= 1 Then
                                ArraySortTyp fnAV(tmBuildLst(), 0), UBound(tmBuildLst), 0, LenB(tmBuildLst(0)), 0, -2, 0
                            End If
                            llLst = mBinarySearchBuildLst(rst_Genl!astLsfCode)
                        End If
                    End If
                    If llLst = -1 Then
                        'Handle the case where ast exist but lst does not exist and the complete flag set
                        smBuildLstKey = ""
                        gBuildAstInfoFromAst = False
                        Exit Function
                    End If
                End If
                'If tmBuildLst(llLst).tLST.lSdfCode <> rst_Genl!astSdfCode Then
                If (gIsAstStatus(rst_Genl!astStatus, ASTEXTENDED_MG)) Or (gIsAstStatus(rst_Genl!astStatus, ASTEXTENDED_REPLACEMENT)) Then
                    blExtendExist = True
                Else
                    If (tmBuildLst(llLst).tLST.lSdfCode <> rst_Genl!astSdfCode) And (tmBuildLst(llLst).tLST.iStatus <= 10) Then
                        'Handle the case where ast exist and lst exist but reference different spots
                        smBuildLstKey = ""
                        gBuildAstInfoFromAst = False
                        Exit Function
                    End If
                End If
                If (llSelGsfCode <= 0) Or (tmBuildLst(llLst).tLST.lgsfCode = llSelGsfCode) Then
                    If Not blAstExist Then
                        If blFilterByAirDates Then
                            llAirDate = DateValue(Format$(rst_Genl!astAirDate, sgShowDateForm))
                            If (llAirDate >= llCpttSDate) And (llAirDate <= llCpttEDate) Then
                                blAstExist = True
                            End If
                        Else
                            llFeedDate = DateValue(Format$(rst_Genl!astFeedDate, sgShowDateForm))
                            If (llFeedDate >= llCpttSDate) And (llFeedDate <= llCpttEDate) Then
                                blAstExist = True
                            End If
                        End If
                    End If
                    'Dan M for 7639 if isci was imported previously, need to get from different place
                    blImportedISCI = gIsISCIChanged(rst_Genl!astStatus)
                    If blImportedISCI Then
                        tmBuildLst(llLst).tLST.sISCI = ""
                        slSQLQuery = "select cpfIsci from CPF_Copy_Prodct_ISCI where cpfcode = " & rst_Genl!astCpfCode
                        Set cpf_rst = gSQLSelectCall(slSQLQuery)
                        If Not cpf_rst.EOF Then
                            tmBuildLst(llLst).tLST.sISCI = cpf_rst!cpfISCI
                        End If
                        tmBuildLst(llLst).tLST.lCrfCsfCode = 0
                        tmBuildLst(llLst).tLST.lCifCode = 0
                        tmBuildLst(llLst).tLST.sCart = ""
                        tlAstInfo(ilAst).iRegionType = 0
                    End If
                    tlAstInfo(ilAst).lCode = rst_Genl!astCode
                    tlAstInfo(ilAst).lAttCode = rst_Genl!astAtfCode
                    tlAstInfo(ilAst).iShttCode = ilShttCode
                    tlAstInfo(ilAst).iVefCode = ilVefCode
                    tlAstInfo(ilAst).lSdfCode = rst_Genl!astSdfCode
                    tlAstInfo(ilAst).lLstCode = rst_Genl!astLsfCode
                    If (blIncludePledgeInfo) Then
                        tlAstInfo(ilAst).lDatCode = rst_Genl!astDatCode
                    Else
                        tlAstInfo(ilAst).lDatCode = 0
                    End If
                    tlAstInfo(ilAst).iStatus = rst_Genl!astStatus
                    tlAstInfo(ilAst).sAirDate = Format$(rst_Genl!astAirDate, sgShowDateForm)
                    If Second(rst_Genl!astAirTime) <> 0 Then
                        tlAstInfo(ilAst).sAirTime = Format$(rst_Genl!astAirTime, sgShowTimeWSecForm)
                    Else
                        tlAstInfo(ilAst).sAirTime = Format$(rst_Genl!astAirTime, sgShowTimeWOSecForm)
                    End If
        
                    tlAstInfo(ilAst).sFeedDate = Format$(rst_Genl!astFeedDate, sgShowDateForm)
                    If Second(rst_Genl!astFeedTime) <> 0 Then
                        tlAstInfo(ilAst).sFeedTime = Format$(rst_Genl!astFeedTime, sgShowTimeWSecForm)
                    Else
                        tlAstInfo(ilAst).sFeedTime = Format$(rst_Genl!astFeedTime, sgShowTimeWOSecForm)
                    End If
                    'llLogDate = gDateValue(Format$(rst!lstLogDate, sgShowDateForm))
                    'llLogTime = gTimeToLong(Format$(rst!lstLogTime, sgShowTimeWSecForm), False) + 3600 * ilLocalAdj
                    'If llLogTime < 0 Then
                    '    llLogTime = llLogTime + 86400
                    '    llLogDate = llLogDate - 1
                    'ElseIf llLogTime > 86400 Then
                    '    llLogTime = llLogTime - 86400
                    '    llLogDate = llLogDate + 1
                    'End If
                    'tlAstInfo(ilAst).sPledgeDate = Format$(llLogDate, "m/d/yyyy")
                    '3/8/16: Get MG Pledge date from Missed
                    tlAstInfo(ilAst).sPledgeDate = tlAstInfo(ilAst).sFeedDate
                    blMGSpot = gGetMissedPledgeForMG(tlAstInfo(ilAst).iStatus, tlAstInfo(ilAst).sFeedDate, rst_Genl!astLkAstCode, slMGMissedFeedDate, slMGMissedFeedTime)
                    If blMGSpot Then
                        tlAstInfo(ilAst).sPledgeDate = slMGMissedFeedDate
                    End If
                    tlAstInfo(ilAst).iAirPlay = 1
                    llDat = -1
                    If (mGetPledgeByEvent(ilVefCode) <> "Y") Then
                        'If rst!astDatCode > 0 Then
                        If tlAstInfo(ilAst).lDatCode > 0 Then
                            lgSTime9 = timeGetTime
                            'llDatCode = rst!astDatCode
                            'slSQLQuery = "SELECT * FROM dat WHERE datCode = " & llDatCode '& ")"
                            'Set dat_rst = gSQLSelectCall(slSQLQuery)
                            'llDat = mBinarySearchDatRst(rst!astDatCode)
                            llDat = mBinarySearchDatRst(tlAstInfo(ilAst).lDatCode)
                            lgETime9 = timeGetTime
                            lgTtlTime9 = lgTtlTime9 + (lgETime9 - lgSTime9)
                            'If Not dat_rst.EOF Then
                            If llDat <> -1 Then
                                blPledgeDefined = True
                                tlAstInfo(ilAst).iPledgeStatus = tmDatRst(llDat).iFdStatus  'dat_rst!datFdStatus
                                If tgStatusTypes(tlAstInfo(ilAst).iPledgeStatus).iPledged <= 1 Then
                                    tlAstInfo(ilAst).sPledgeStartTime = tmDatRst(llDat).sPdStTime 'Format$(CStr(rst!datPdStTime), sgShowTimeWSecForm)
                                    tlAstInfo(ilAst).sPledgeEndTime = tmDatRst(llDat).sPdEdTime  'Format$(CStr(rst!datPdEdTime), sgShowTimeWSecForm)
                                Else
'                                    tlAstInfo(ilAst).sPledgeStartTime = ""
'                                    tlAstInfo(ilAst).sPledgeEndTime = ""
                                    tlAstInfo(ilAst).sPledgeStartTime = tlAstInfo(ilAst).sFeedTime
                                    tlAstInfo(ilAst).sPledgeEndTime = tlAstInfo(ilAst).sFeedTime
                                End If
                                tlAstInfo(ilAst).iAirPlay = tmDatRst(llDat).iAirPlayNo
                                If tmATTCrossDates(ilPass).iLoadFactor > 1 Then
                                    If ilAst > 0 Then
                                        If rst_Genl!astAdfCode = tlAstInfo(ilAst - 1).iAdfCode Then
                                            If gTimeToLong(tlAstInfo(ilAst).sFeedTime, False) = gTimeToLong(tlAstInfo(ilAst - 1).sFeedTime, False) Then
                                                If gDateValue(tlAstInfo(ilAst).sFeedDate) = gDateValue(tlAstInfo(ilAst - 1).sFeedDate) Then
                                                    If tlAstInfo(ilAst).lDatCode = tlAstInfo(ilAst - 1).lDatCode Then
                                                        tlAstInfo(ilAst).iAirPlay = tlAstInfo(ilAst - 1).iAirPlay + 1
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                '10/19/18: Ast referencing an non-existing dat record. Reset
                                If blIncludePledgeInfo Then
                                    ReDim tlAstInfo(0 To 0) As ASTINFO
                                    gBuildAstInfoFromAst = False
                                    Exit Function
                                End If
                                'If Pledge not defined, then live
                                'If pledges defined, then not caried
                                tlAstInfo(ilAst).iPledgeStatus = 8   'Flag to be set later once all ast read in
                                If Not blMGSpot Then
                                    tlAstInfo(ilAst).sPledgeStartTime = tlAstInfo(ilAst).sFeedTime
                                    tlAstInfo(ilAst).sPledgeEndTime = tlAstInfo(ilAst).sFeedTime
                                Else
                                    tlAstInfo(ilAst).sPledgeStartTime = slMGMissedFeedTime
                                    tlAstInfo(ilAst).sPledgeEndTime = slMGMissedFeedTime
                                End If
                            End If
                       ' ElseIf rst!astDatCode < 0 Then
                        ElseIf tlAstInfo(ilAst).lDatCode < 0 Then
                            tlAstInfo(ilAst).iPledgeStatus = 8
                            If Not blMGSpot Then
                                tlAstInfo(ilAst).sPledgeStartTime = tlAstInfo(ilAst).sFeedTime
                                tlAstInfo(ilAst).sPledgeEndTime = tlAstInfo(ilAst).sFeedTime
                            Else
                                tlAstInfo(ilAst).sPledgeStartTime = slMGMissedFeedTime
                                tlAstInfo(ilAst).sPledgeEndTime = slMGMissedFeedTime
                            End If
                        Else
                            tlAstInfo(ilAst).iPledgeStatus = 0
                            If Not blMGSpot Then
                                tlAstInfo(ilAst).sPledgeStartTime = tlAstInfo(ilAst).sFeedTime
                                tlAstInfo(ilAst).sPledgeEndTime = tlAstInfo(ilAst).sFeedTime
                            Else
                                tlAstInfo(ilAst).sPledgeStartTime = slMGMissedFeedTime
                                tlAstInfo(ilAst).sPledgeEndTime = slMGMissedFeedTime
                            End If
                        End If
                    Else
                        'Live
                        tlAstInfo(ilAst).iPledgeStatus = 0
                        If Not blMGSpot Then
                            tlAstInfo(ilAst).sPledgeStartTime = tlAstInfo(ilAst).sFeedTime
                            tlAstInfo(ilAst).sPledgeEndTime = tlAstInfo(ilAst).sFeedTime
                        Else
                            tlAstInfo(ilAst).sPledgeStartTime = slMGMissedFeedTime
                            tlAstInfo(ilAst).sPledgeEndTime = slMGMissedFeedTime
                        End If
                    End If
                    If tlAstInfo(ilAst).iPledgeStatus = 0 Then
                        blLivePledge = True
                    End If
                    tlAstInfo(ilAst).iAdfCode = rst_Genl!astAdfCode
                    'tlAstInfo(ilAst).iAnfCode = rst!lstAnfCode
                    'tlAstInfo(ilAst).sProd = rst!lstProd
                    'tlAstInfo(ilAst).sCart = rst!lstCart
                    'tlAstInfo(ilAst).sISCI = rst!lstISCI
                    'tlAstInfo(ilAst).lCifCode = rst!lstCifCode
                    'tlAstInfo(ilAst).lCrfCsfCode = rst!lstCrfCsfCode
                    'tlAstInfo(ilAst).lcpfCode = rst!lstCpfCode
                    tlAstInfo(ilAst).iAnfCode = tmBuildLst(llLst).tLST.iAnfCode
                    tlAstInfo(ilAst).sProd = tmBuildLst(llLst).tLST.sProd
                    tlAstInfo(ilAst).sCart = tmBuildLst(llLst).tLST.sCart
                    tlAstInfo(ilAst).sISCI = tmBuildLst(llLst).tLST.sISCI
                    tlAstInfo(ilAst).lCifCode = tmBuildLst(llLst).tLST.lCifCode
                    '8/31/16: Check if comment suppressed
                    'tlAstInfo(ilAst).lCrfCsfCode = tmBuildLst(llLst).tLST.lCrfCsfCode
                    If slHideCommOnWeb <> "Y" Then
                        tlAstInfo(ilAst).lCrfCsfCode = tmBuildLst(llLst).tLST.lCrfCsfCode
                    Else
                        tlAstInfo(ilAst).lCrfCsfCode = 0
                    End If
                    tlAstInfo(ilAst).lCpfCode = tmBuildLst(llLst).tLST.lCpfCode
                    'Clear after obtaining all records if Not Carried
                    Select Case Weekday(tlAstInfo(ilAst).sFeedDate, vbMonday) - 1
                        Case 0
                            tlAstInfo(ilAst).sPdDays = "Mo"
                        Case 1
                            tlAstInfo(ilAst).sPdDays = "Tu"
                        Case 2
                            tlAstInfo(ilAst).sPdDays = "We"
                        Case 3
                            tlAstInfo(ilAst).sPdDays = "Th"
                        Case 4
                            tlAstInfo(ilAst).sPdDays = "Fr"
                        Case 5
                            tlAstInfo(ilAst).sPdDays = "Sa"
                        Case 6
                            tlAstInfo(ilAst).sPdDays = "Su"
                    End Select
                    tlAstInfo(ilAst).iCPStatus = rst_Genl!astCPStatus
                    'tlAstInfo(ilAst).sLstZone = rst!lstZone
                    'tlAstInfo(ilAst).lCntrNo = rst!lstCntrNo
                    'tlAstInfo(ilAst).iLen = rst!lstLen
                    'tlAstInfo(ilAst).lGsfCode = rst!lstGsfCode
                    tlAstInfo(ilAst).sLstZone = tmBuildLst(llLst).tLST.sZone
                    tlAstInfo(ilAst).lCntrNo = tmBuildLst(llLst).tLST.lCntrNo
                    tlAstInfo(ilAst).iLen = rst_Genl!astLen                'tmBuildLst(llLst).tLST.iLen
                    tlAstInfo(ilAst).lgsfCode = tmBuildLst(llLst).tLST.lgsfCode
                    '4/12/15: verify that each game vehicle has ast spots
                    If UBound(llGsfCode) > 0 And llSelGsfCode <= 0 Then
                        For llGsf = 0 To UBound(llGsfCode) - 1 Step 1
                            If llGsfCode(llGsf) = tmBuildLst(llLst).tLST.lgsfCode Then
                                llGsfCode(llGsf) = -llGsfCode(llGsf)
                            End If
                        Next llGsf
                    End If
                    tlAstInfo(ilAst).lLkAstCode = rst_Genl!astLkAstCode
                    tlAstInfo(ilAst).iMissedMnfCode = rst_Genl!astMissedMnfCode
                    tlAstInfo(ilAst).sGISCI = tmBuildLst(llLst).tLST.sISCI
                    tlAstInfo(ilAst).sGCart = tmBuildLst(llLst).tLST.sCart
                    tlAstInfo(ilAst).sGProd = tmBuildLst(llLst).tLST.sProd
                    tlAstInfo(ilAst).lEvtIDCefCode = tmBuildLst(llLst).tLST.lEvtIDCefCode
                    'Dan M 7639
'                    If rst_Genl!astRsfCode > 0 Then
                    If rst_Genl!astRsfCode > 0 And Not blImportedISCI Then
                        '2/14/17: Check that region copy still exist
                        If IsNull(rst_Genl!rsfCode) Then
                            ReDim tlAstInfo(0 To 0) As ASTINFO
                            gBuildAstInfoFromAst = False
                            Exit Function
                        End If
                        lgSTime10 = timeGetTime
                        'If rst!lstBkoutLstCode > 0 Then
                        If tmBuildLst(llLst).tLST.lBkoutLstCode > 0 Then
                            tlAstInfo(ilAst).iRegionType = 2
                            llBkoutLst = mBinarySearchBuildLst(tmBuildLst(llLst).tLST.lBkoutLstCode)
                            If llBkoutLst <> -1 Then
                                tlAstInfo(ilAst).sGISCI = tmBuildLst(llBkoutLst).tLST.sISCI
                                tlAstInfo(ilAst).sGCart = tmBuildLst(llBkoutLst).tLST.sCart
                                tlAstInfo(ilAst).sGProd = tmBuildLst(llBkoutLst).tLST.sProd
                            End If
                        Else
                            tlAstInfo(ilAst).iRegionType = 1
                        End If
                        'Get Split copy
                        tlAstInfo(ilAst).lRRsfCode = rst_Genl!astRsfCode
                        tlAstInfo(ilAst).lIrtCode = 0
                        'slSQLQuery = "SELECT * FROM rsf_Region_Schd_Copy WHERE (rsfCode = " & rst!astRsfCode & ")"
                        'Set rsf_rst = gSQLSelectCall(slSQLQuery)
                        'If Not rsf_rst.EOF Then
                            tlAstInfo(ilAst).lRCrfCode = rst_Genl!rsfCrfCode
                            tlAstInfo(ilAst).lRCifCode = rst_Genl!rsfCopyCode
                            ilRet = gGetCopy(rst_Genl!rstPtType, rst_Genl!rsfCopyCode, rst_Genl!rsfCrfCode, True, tlAstInfo(ilAst).sRCart, tlAstInfo(ilAst).sRProduct, tlAstInfo(ilAst).sRISCI, tlAstInfo(ilAst).sRCreativeTitle, tlAstInfo(ilAst).lRCrfCsfCode, tlAstInfo(ilAst).lRCpfCode, ilRAdfCode, ilVefCode)
                        'Else
                        '    tlAstInfo(ilAst).iRegionType = 0
                        '    tlAstInfo(ilAst).sRCart = ""
                        '    tlAstInfo(ilAst).sRProduct = ""
                        '    tlAstInfo(ilAst).sRISCI = ""
                        '    tlAstInfo(ilAst).sRCreativeTitle = ""
                        '    tlAstInfo(ilAst).lRCrfCsfCode = 0
                        '    tlAstInfo(ilAst).lRCifCode = 0
                        '    tlAstInfo(ilAst).lRCrfCode = 0
                        '    tlAstInfo(ilAst).lRRsfCode = 0
                        '    tlAstInfo(ilAst).lRCpfCode = 0
                        'End If
                        lgETime10 = timeGetTime
                        lgTtlTime10 = lgTtlTime10 + (lgETime10 - lgSTime10)
                    'Dan M 7639
'                    ElseIf bmImportSpot Then
                    '4/13/18: Added MG also set the SpotTYpe = 5 but ImportSpot is N
                    'ElseIf bmImportSpot And Not blImportedISCI Then
                    ElseIf bmImportSpot And (tmBuildLst(llLst).tLST.sImportedSpot = "Y") And Not blImportedISCI Then
                        '12/10/15: Handle Blackouts
                        If tmBuildLst(llLst).tLST.lBkoutLstCode > 0 Then
                            llLstCode = tmBuildLst(llLst).tLST.lBkoutLstCode
                        Else
                            llLstCode = tmBuildLst(llLst).tLST.lCode
                        End If
                        slSQLQuery = "SELECT *"
                        slSQLQuery = slSQLQuery + " FROM irt"
                        'slSQLQuery = slSQLQuery + " WHERE (irtShttCode= " & ilShttCode & " And irtLstCode = " & tmBuildLst(llLst).tLST.lCode & ")"
                        slSQLQuery = slSQLQuery + " WHERE (irtShttCode= " & ilShttCode & " And irtLstCode = " & llLstCode & ")"
                        Set rst_irt = gSQLSelectCall(slSQLQuery)
                        If Not rst_irt.EOF Then
                            tlAstInfo(ilAst).lIrtCode = rst_irt!irtCode
                            If rst_irt!irtType = "B" Then
                                tlAstInfo(ilAst).iRegionType = 2
                                llBkoutLst = mBinarySearchBuildLst(tmBuildLst(llLst).tLST.lBkoutLstCode)
                                If llBkoutLst <> -1 Then
                                    tlAstInfo(ilAst).sGISCI = tmBuildLst(llBkoutLst).tLST.sISCI
                                    tlAstInfo(ilAst).sGCart = tmBuildLst(llBkoutLst).tLST.sCart
                                    tlAstInfo(ilAst).sGProd = tmBuildLst(llBkoutLst).tLST.sProd
                                End If
                            Else
                                tlAstInfo(ilAst).iRegionType = 1
                            End If
                            'Get Split copy
                            tlAstInfo(ilAst).sRCart = rst_irt!irtCart
                            tlAstInfo(ilAst).sRProduct = rst_irt!irtProduct
                            tlAstInfo(ilAst).sRISCI = rst_irt!irtISCI
                            tlAstInfo(ilAst).sRCreativeTitle = rst_irt!irtCreativeTitle
                            tlAstInfo(ilAst).lRCrfCsfCode = 0
                            tlAstInfo(ilAst).lRCifCode = 0
                            tlAstInfo(ilAst).lRCrfCode = 0
                            tlAstInfo(ilAst).lRRsfCode = 0
                            tlAstInfo(ilAst).lRCpfCode = 0
                            tlAstInfo(ilAst).sReplacementCue = rst_irt!irtXDSCue
                        Else
                            tlAstInfo(ilAst).iRegionType = 0
                            tlAstInfo(ilAst).sRCart = ""
                            tlAstInfo(ilAst).sRProduct = ""
                            tlAstInfo(ilAst).sRISCI = ""
                            tlAstInfo(ilAst).sRCreativeTitle = ""
                            tlAstInfo(ilAst).lRCrfCsfCode = 0
                            tlAstInfo(ilAst).lRCifCode = 0
                            tlAstInfo(ilAst).lRCrfCode = 0
                            tlAstInfo(ilAst).lRRsfCode = 0
                            tlAstInfo(ilAst).lRCpfCode = 0
                            tlAstInfo(ilAst).lIrtCode = 0
                            tlAstInfo(ilAst).sReplacementCue = ""
                        End If
                    Else
                        tlAstInfo(ilAst).iRegionType = 0
                        tlAstInfo(ilAst).sRCart = ""
                        tlAstInfo(ilAst).sRProduct = ""
                        tlAstInfo(ilAst).sRISCI = ""
                        tlAstInfo(ilAst).sRCreativeTitle = ""
                        tlAstInfo(ilAst).lRCrfCsfCode = 0
                        tlAstInfo(ilAst).lRCifCode = 0
                        tlAstInfo(ilAst).lRCrfCode = 0
                        tlAstInfo(ilAst).lRRsfCode = 0
                        tlAstInfo(ilAst).lRCpfCode = 0
                        tlAstInfo(ilAst).lIrtCode = 0
                        tlAstInfo(ilAst).sReplacementCue = ""
                    End If
                    lgSTime6 = timeGetTime
                    
                    tlAstInfo(ilAst).sTruePledgeDays = String(7, "N")
                    tlAstInfo(ilAst).sPdTimeExceedsFdTime = "N"
                    tlAstInfo(ilAst).sPdDayFed = ""
                    Select Case tgStatusTypes(tlAstInfo(ilAst).iPledgeStatus).iPledged
                        Case 0  'live
                            'Determine later
                            tlAstInfo(ilAst).sTruePledgeEndTime = ""
                            Mid(tlAstInfo(ilAst).sTruePledgeDays, Weekday(tlAstInfo(ilAst).sPledgeDate, vbMonday), 1) = "Y"
                        Case 1  'Delay
                            tlAstInfo(ilAst).sTruePledgeEndTime = tlAstInfo(ilAst).sPledgeEndTime
                            'If rst!astDatCode > 0 Then
                            If tlAstInfo(ilAst).lDatCode > 0 Then
                                'If Not rst.EOF Then
                                If llDat <> -1 Then
                                    ilPdIndex = -1
                                    If tmDatRst(llDat).iPdMon = 1 Then  ' rst!datPdMon = 1 Then
                                        ilPdIndex = 1
                                    End If
                                    If tmDatRst(llDat).iPdTue = 1 Then  'rst!datPdTue = 1 Then
                                        If ilPdIndex = -1 Then
                                            ilPdIndex = 2
                                        End If
                                    End If
                                    If tmDatRst(llDat).iPdWed = 1 Then  'rst!datPdWed = 1 Then
                                        If ilPdIndex = -1 Then
                                            ilPdIndex = 3
                                        End If
                                    End If
                                    If tmDatRst(llDat).iPdThu = 1 Then  'rst!datPdThu = 1 Then
                                        If ilPdIndex = -1 Then
                                            ilPdIndex = 4
                                        End If
                                    End If
                                    If tmDatRst(llDat).iPdFri = 1 Then  'rst!datPdFri = 1 Then
                                        If ilPdIndex = -1 Then
                                            ilPdIndex = 5
                                        End If
                                    End If
                                    If tmDatRst(llDat).iPdSat = 1 Then  'rst!datPdSat = 1 Then
                                        If ilPdIndex = -1 Then
                                            ilPdIndex = 6
                                        End If
                                    End If
                                    If tmDatRst(llDat).iPdSun = 1 Then  'rst!datPdSun = 1 Then
                                        If ilPdIndex = -1 Then
                                            ilPdIndex = 7
                                        End If
                                    End If
                                    ilFdIndex = -1
                                    If tmDatRst(llDat).iFdMon = 1 Then  'rst!datFdMon = 1 Then
                                        ilFdIndex = 1
                                    End If
                                    If tmDatRst(llDat).iFdTue = 1 Then  'rst!datFdTue = 1 Then
                                        If ilFdIndex = -1 Then
                                            ilFdIndex = 2
                                        End If
                                    End If
                                    If tmDatRst(llDat).iFdWed = 1 Then  'rst!datFdWed = 1 Then
                                        If ilFdIndex = -1 Then
                                            ilFdIndex = 3
                                        End If
                                    End If
                                    If tmDatRst(llDat).iFdThu = 1 Then  'rst!datFdThu = 1 Then
                                        If ilFdIndex = -1 Then
                                            ilFdIndex = 4
                                        End If
                                    End If
                                    If tmDatRst(llDat).iFdFri = 1 Then  'rst!datFdFri = 1 Then
                                        If ilFdIndex = -1 Then
                                            ilFdIndex = 5
                                        End If
                                    End If
                                    If tmDatRst(llDat).iFdSat = 1 Then  'rst!datFdSat = 1 Then
                                        If ilFdIndex = -1 Then
                                            ilFdIndex = 6
                                        End If
                                    End If
                                    If tmDatRst(llDat).iFdSun = 1 Then  'rst!datFdSun = 1 Then
                                        If ilFdIndex = -1 Then
                                            ilFdIndex = 7
                                        End If
                                    End If
                                    'llFeedTimeDiff = gTimeToLong(rst!datFdEdTime, True) - gTimeToLong(rst!datFdStTime, False)
                                    llFeedTimeDiff = gTimeToLong(tmDatRst(llDat).sFdEdTime, True) - gTimeToLong(tmDatRst(llDat).sFdStTime, False)
                                    If llFeedTimeDiff = 0 Then
                                        llFeedTimeDiff = 1
                                    End If
                                    'llPledgeTimeDiff = gTimeToLong(rst!datPdEdTime, True) - gTimeToLong(rst!datPdStTime, False)
                                    llPledgeTimeDiff = gTimeToLong(tmDatRst(llDat).sPdEdTime, True) - gTimeToLong(tmDatRst(llDat).sPdStTime, False)
                                    '6/3/15: Add 5 minute rule
                                    'If llPledgeTimeDiff > llFeedTimeDiff Then
                                    If (llPledgeTimeDiff > llFeedTimeDiff) Or (((llPledgeTimeDiff = llFeedTimeDiff)) And (llPledgeTimeDiff > 300)) Then
                                        tlAstInfo(ilAst).sPdTimeExceedsFdTime = "Y"
                                    End If
                                    tlAstInfo(ilAst).sPdDayFed = tmDatRst(llDat).sPdDayFed  'rst!datPdDayFed
                                    If ilPdIndex >= ilFdIndex Then
                                        ilAdjIndex = ilPdIndex - ilFdIndex
                                    Else
                                        'If rst!datPdDayFed = "B" Then
                                        If tmDatRst(llDat).sPdDayFed = "B" Then
                                            ilAdjIndex = ilPdIndex - ilFdIndex
                                        Else
                                            ilAdjIndex = 7 + ilPdIndex - ilFdIndex
                                        End If
                                    End If
                                    '3/8/16: Use Missed Feed Date if MG
                                    If Not blMGSpot Then
                                        tlAstInfo(ilAst).sPledgeDate = Format$(DateValue(gAdjYear(tlAstInfo(ilAst).sFeedDate)) + ilAdjIndex, sgShowDateForm)
                                    Else
                                        tlAstInfo(ilAst).sPledgeDate = Format$(DateValue(gAdjYear(slMGMissedFeedDate)) + ilAdjIndex, sgShowDateForm)
                                    End If
                                    Mid(tlAstInfo(ilAst).sTruePledgeDays, Weekday(tlAstInfo(ilAst).sPledgeDate, vbMonday), 1) = "Y"
                                    Select Case Weekday(tlAstInfo(ilAst).sPledgeDate, vbMonday) - 1
                                        Case 0
                                            tlAstInfo(ilAst).sPdDays = "Mo"
                                        Case 1
                                            tlAstInfo(ilAst).sPdDays = "Tu"
                                        Case 2
                                            tlAstInfo(ilAst).sPdDays = "We"
                                        Case 3
                                            tlAstInfo(ilAst).sPdDays = "Th"
                                        Case 4
                                            tlAstInfo(ilAst).sPdDays = "Fr"
                                        Case 5
                                            tlAstInfo(ilAst).sPdDays = "Sa"
                                        Case 6
                                            tlAstInfo(ilAst).sPdDays = "Su"
                                    End Select
                                Else
                                    Mid(tlAstInfo(ilAst).sTruePledgeDays, Weekday(tlAstInfo(ilAst).sPledgeDate, vbMonday), 1) = "Y"
                                End If
                            Else
                                Mid(tlAstInfo(ilAst).sTruePledgeDays, Weekday(tlAstInfo(ilAst).sPledgeDate, vbMonday), 1) = "Y"
                            End If
                        Case 2  'Not Carried
                            If Not blMGSpot Then
                                tlAstInfo(ilAst).sTruePledgeEndTime = tlAstInfo(ilAst).sFeedTime
                            Else
                                tlAstInfo(ilAst).sTruePledgeEndTime = slMGMissedFeedTime
                            End If
                            Mid(tlAstInfo(ilAst).sTruePledgeDays, Weekday(tlAstInfo(ilAst).sPledgeDate, vbMonday), 1) = "Y"
                        Case 3  'carried but no pledge
                            If Not blMGSpot Then
                                tlAstInfo(ilAst).sTruePledgeEndTime = tlAstInfo(ilAst).sFeedTime
                            Else
                                tlAstInfo(ilAst).sTruePledgeEndTime = slMGMissedFeedTime
                            End If
                            tlAstInfo(ilAst).sTruePledgeDays = String(7, "Y")
                    End Select

                    tlAstInfo(ilAst).sEmbeddedOrROS = smDefaultEmbeddedOrROS
                    If tlAstInfo(ilAst).lDatCode > 0 Then
                        If llDat <> -1 Then
                            tlAstInfo(ilAst).sEmbeddedOrROS = tmDatRst(llDat).sEmbeddedOrROS
                            If Trim$(tlAstInfo(ilAst).sEmbeddedOrROS = "") Then
                                tlAstInfo(ilAst).sEmbeddedOrROS = smDefaultEmbeddedOrROS
                            End If
                        End If
                    End If
                    tlAstInfo(ilAst).sStationCompliant = rst_Genl!astStationCompliant
                    tlAstInfo(ilAst).sAgencyCompliant = rst_Genl!astAgencyCompliant
                    tlAstInfo(ilAst).sAffidavitSource = gRemoveIllegalChars(rst_Genl!astAffidavitSource)
                    lgETime6 = timeGetTime
                    lgTtlTime6 = lgTtlTime6 + (lgETime6 - lgSTime6)
                    tlAstInfo(ilAst).lPrevBkoutLstCode = 0
                    'tlAstInfo(ilAst).iLstLnVefCode = rst!lstLnVefCode
                    'tlAstInfo(ilAst).lLstBkoutLstCode = rst!lstBkoutLstCode
                    'tlAstInfo(ilAst).sLstStartDate = rst!lstStartDate
                    'tlAstInfo(ilAst).sLstEndDate = rst!lstEndDate
                    'tlAstInfo(ilAst).iLstSpotsWk = rst!lstSpotsWk
                    'tlAstInfo(ilAst).iLstMon = rst!lstMon
                    'tlAstInfo(ilAst).iLstTue = rst!lstTue
                    'tlAstInfo(ilAst).iLstWed = rst!lstWed
                    'tlAstInfo(ilAst).iLstThu = rst!lstThu
                    'tlAstInfo(ilAst).iLstFri = rst!lstFri
                    'tlAstInfo(ilAst).iLstSat = rst!lstSat
                    'tlAstInfo(ilAst).iLstSun = rst!lstSun
                    'tlAstInfo(ilAst).iLineNo = rst!lstLineNo
                    
                    tlAstInfo(ilAst).iLstLnVefCode = tmBuildLst(llLst).tLST.iLnVefCode
                    tlAstInfo(ilAst).lLstBkoutLstCode = tmBuildLst(llLst).tLST.lBkoutLstCode
                    tlAstInfo(ilAst).sLstStartDate = tmBuildLst(llLst).tLST.sStartDate
                    tlAstInfo(ilAst).sLstEndDate = tmBuildLst(llLst).tLST.sEndDate
                    tlAstInfo(ilAst).iLstSpotsWk = tmBuildLst(llLst).tLST.iSpotsWk
                    tlAstInfo(ilAst).iLstMon = tmBuildLst(llLst).tLST.iMon
                    tlAstInfo(ilAst).iLstTue = tmBuildLst(llLst).tLST.iTue
                    tlAstInfo(ilAst).iLstWed = tmBuildLst(llLst).tLST.iWed
                    tlAstInfo(ilAst).iLstThu = tmBuildLst(llLst).tLST.iThu
                    tlAstInfo(ilAst).iLstFri = tmBuildLst(llLst).tLST.iFri
                    tlAstInfo(ilAst).iLstSat = tmBuildLst(llLst).tLST.iSat
                    tlAstInfo(ilAst).iLstSun = tmBuildLst(llLst).tLST.iSun
                    tlAstInfo(ilAst).iLineNo = tmBuildLst(llLst).tLST.iLineNo
                    tlAstInfo(ilAst).iSpotType = tmBuildLst(llLst).tLST.iSpotType
                    tlAstInfo(ilAst).sSplitNet = tmBuildLst(llLst).tLST.sSplitNetwork
                    tlAstInfo(ilAst).iAgfCode = tmBuildLst(llLst).tLST.iAgfCode
                    tlAstInfo(ilAst).sLstLnStartTime = tmBuildLst(llLst).tLST.sLnStartTime
                    tlAstInfo(ilAst).sLstLnEndTime = tmBuildLst(llLst).tLST.sLnEndTime
                    If tgStatusTypes(tlAstInfo(ilAst).iPledgeStatus).iPledged = 0 Then
                        tlAstInfo(ilAst).sTruePledgeEndTime = Format$(gLongToTime(gTimeToLong(Format$(tlAstInfo(ilAst).sPledgeStartTime, "h:mm:ssam/pm"), False) + tmBuildLst(llLst).iBreakLength), sgShowTimeWSecForm)
                    End If
        
                    ''3/7/16: Check if exporting Traffic posted times
                    'If llDat >= 0 Then
                    '    mGetTrafficPostedTimes llVpf, tmBuildLst(llLst).tLST, tgStatusTypes(tlAstInfo(ilAst).iPledgeStatus).iPledged, tmATTCrossDates(ilPass).iDACode, rst_Genl!astCPStatus, ilLocalAdj, ilAdjIndex, tlAstInfo(ilAst).sPledgeStartTime, tlAstInfo(ilAst).sFeedTime, tlAstInfo(ilAst).sAirDate, tlAstInfo(ilAst).sAirTime
                    'Else
                    '    mGetTrafficPostedTimes llVpf, tmBuildLst(llLst).tLST, tgStatusTypes(tlAstInfo(ilAst).iPledgeStatus).iPledged, tmATTCrossDates(ilPass).iDACode, rst_Genl!astCPStatus, ilLocalAdj, ilAdjIndex, "12am", "12am", tlAstInfo(ilAst).sAirDate, tlAstInfo(ilAst).sAirTime
                    'End If
                    
                    'Create sort key
                    If igTimes = 0 Then
                        slSortDate = gDateValue(tlAstInfo(ilAst).sAirDate)
                        Do While Len(slSortDate) < 6
                            slSortDate = "0" & slSortDate
                        Loop
                        slSortTime = gTimeToLong(tlAstInfo(ilAst).sAirTime, False)
                        Do While Len(slSortTime) < 6
                            slSortTime = "0" & slSortTime
                        Loop
                    ''8/18/16: If posted, sort by air date/time otherwise sort by feed date/time
                    ''ElseIf (igTimes <> 2) And (igTimes <> 4) Then
                    'ElseIf (igTimes <> 2) And (igTimes <> 4) And (tlAstInfo(ilAst).iCPStatus = 1) Then
                    '    slSortDate = gDateValue(tlAstInfo(ilAst).sAirDate)
                    '    Do While Len(slSortDate) < 6
                    '        slSortDate = "0" & slSortDate
                    '    Loop
                    '    slSortTime = gTimeToLong(tlAstInfo(ilAst).sAirTime, False)
                    '    Do While Len(slSortTime) < 6
                    '        slSortTime = "0" & slSortTime
                    '    Loop
                    Else
                        slSortDate = gDateValue(tlAstInfo(ilAst).sFeedDate)
                        Do While Len(slSortDate) < 6
                            slSortDate = "0" & slSortDate
                        Loop
                        slSortTime = gTimeToLong(tlAstInfo(ilAst).sFeedTime, False)
                        Do While Len(slSortTime) < 6
                            slSortTime = "0" & slSortTime
                        Loop
                    End If
                    slSortBreak = tmBuildLst(llLst).tLST.iBreakNo
                    Do While Len(slSortBreak) < 3
                        slSortBreak = "0" & slSortBreak
                    Loop
                    slSortPosition = tmBuildLst(llLst).tLST.iPositionNo
                    Do While Len(slSortPosition) < 3
                        slSortPosition = "0" & slSortPosition
                    Loop
                    slSortAirPlay = tlAstInfo(ilAst).iAirPlay
                    Do While Len(slSortAirPlay) < 2
                        slSortAirPlay = "0" & slSortAirPlay
                    Loop
                    
                    tlAstInfo(ilAst).sKey = slSortDate & slSortTime & slSortBreak & slSortPosition & slSortAirPlay
                    ilAst = ilAst + 1
                    If ilAst >= UBound(tlAstInfo) Then
                        ReDim Preserve tlAstInfo(0 To ilAst + 1000) As ASTINFO
                    End If
                End If
                lgETime14 = timeGetTime
                lgTtlTime14 = lgTtlTime14 + (lgETime14 - lgSTime14)
                rst_Genl.MoveNext
            Loop
        Next ilPass
        '4/7/18: Test if ast exist for date
        If blFilterByAirDates Then
            For ilPass = 0 To UBound(tmATTCrossDates) - 1 Step 1
                For llLst = 0 To UBound(llLstDate) - 1 Step 1
                    'If by airing date, check if ast exit before existing
                    If (llLstDate(llLst) > 0) Then
                        slSQLQuery = "SELECT * FROM ast WHERE (astAtfCode = " & tmATTCrossDates(ilPass).lAttCode & " AND astFeedDate = '" & Format$(llLstDate(llLst), sgSQLDateForm) & "'" & ")"
                        'Set rst_Genl = cnn.Execute(slSQLQuery)
                        Set rst_Genl = gSQLSelectCall(slSQLQuery)
                        If Not rst_Genl.EOF Then
                            llLstDate(llLst) = -llLstDate(llLst)
                        End If
                    End If
                Next llLst
            Next ilPass
        End If
    Next ilLoop
    
    '11/8/16: Test if Lst date exist and ast did not exist
    For llLst = 0 To UBound(llLstDate) - 1 Step 1
        If llLstDate(llLst) > 0 Then
            ReDim tlAstInfo(0 To 0) As ASTINFO
            gBuildAstInfoFromAst = False
            Exit Function
        End If
    Next llLst
    
    For llGsf = 0 To UBound(llGsfCode) - 1 Step 1
        If llGsfCode(llGsf) > 0 Then
            ReDim tlAstInfo(0 To 0) As ASTINFO
            gBuildAstInfoFromAst = False
            Exit Function
        End If
    Next llGsf
    ReDim Preserve tlAstInfo(0 To ilAst) As ASTINFO
    If blExtendExist Then
        For ilAst = 0 To UBound(tlAstInfo) - 1 Step 1
            'Create sort key
            If igTimes = 0 Then
                slSortDate = gDateValue(tlAstInfo(ilAst).sAirDate)
                Do While Len(slSortDate) < 6
                    slSortDate = "0" & slSortDate
                Loop
                slSortTime = gTimeToLong(tlAstInfo(ilAst).sAirTime, False)
                Do While Len(slSortTime) < 6
                    slSortTime = "0" & slSortTime
                Loop
            ElseIf (igTimes <> 2) And (igTimes <> 4) Then
                slSortDate = gDateValue(tlAstInfo(ilAst).sAirDate)
                Do While Len(slSortDate) < 6
                    slSortDate = "0" & slSortDate
                Loop
                slSortTime = gTimeToLong(tlAstInfo(ilAst).sAirTime, False)
                Do While Len(slSortTime) < 6
                    slSortTime = "0" & slSortTime
                Loop
            Else
                slSortDate = gDateValue(tlAstInfo(ilAst).sFeedDate)
                Do While Len(slSortDate) < 6
                    slSortDate = "0" & slSortDate
                Loop
                slSortTime = gTimeToLong(tlAstInfo(ilAst).sFeedTime, False)
                Do While Len(slSortTime) < 6
                    slSortTime = "0" & slSortTime
                Loop
            End If
            slSortPosition = ilAst
            Do While Len(slSortPosition) < 5
                slSortPosition = "0" & slSortPosition
            Loop
            tlAstInfo(ilAst).sKey = slSortDate & slSortTime & slSortPosition
        Next ilAst
    End If
    If UBound(tlAstInfo) - 1 >= 1 Then
        ArraySortTyp fnAV(tlAstInfo(), 0), UBound(tlAstInfo), 0, LenB(tlAstInfo(0)), 0, LenB(tlAstInfo(0).sKey), 0
    End If
    ''For Live avails, set True end time
    'lgSTime9 = timeGetTime
    'If blLivePledge Then
    '    For ilLoop = 0 To UBound(tlAstInfo) - 1 Step 1
    '        If tgStatusTypes(tlAstInfo(ilLoop).iPledgeStatus).iPledged = 0 Then
    '            llTotalLen = 0
    '            For ilInner = 0 To UBound(tlAstInfo) - 1 Step 1
    '                If tgStatusTypes(tlAstInfo(ilInner).iPledgeStatus).iPledged = 0 Then
    '                    If gTimeToLong(tlAstInfo(ilLoop).sPledgeStartTime, False) = gTimeToLong(tlAstInfo(ilInner).sPledgeStartTime, False) Then
    '                        llTotalLen = llTotalLen + tlAstInfo(ilInner).iLen
    '                    End If
    '                End If
    '            Next ilInner
    '            tlAstInfo(ilLoop).sTruePledgeEndTime = Format$(gLongToTime(gTimeToLong(Format$(tlAstInfo(ilLoop).sPledgeStartTime, "h:mm:ssam/pm"), False) + llTotalLen), sgShowTimeWSecForm)
    '        End If
    '
    '    Next ilLoop
    'End If
    'lgETime9 = timeGetTime
    'lgTtlTime9 = lgTtlTime9 + (lgETime9 - lgSTime9)
    On Error Resume Next
    Erase llGsfCode
    If blAstExist Then
        If ilAdfCode <> ilInAdfCode Then
            mFilterByAdvt tlAstInfo, ilInAdfCode
        End If
        gBuildAstInfoFromAst = True
    Else
        gBuildAstInfoFromAst = False
    End If
    lgETime8 = timeGetTime
    lgTtlTime8 = lgTtlTime8 + (lgETime8 - lgSTime8)
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "gBuildAstInfoFromAst"
    gBuildAstInfoFromAst = False
    Resume Next
CheckForCopyErr:
    ilRet = 1
    Resume Next
End Function

Public Function gGetPledgeGivenAstInfo(tlDatPledgeInfo As DATPLEDGEINFO) As Integer
    Dim blPledgeDefined As Boolean
    Dim llDat As Long
    '5/29/15: Fix case where pledge day is offset from the feed day
    Dim ilDay As Integer
    Dim ilAdjDay As Integer
    Dim ilFdDay As Integer
    Dim ilPdDay As Integer
    Dim slSQLQuery As String
    
    tlDatPledgeInfo.sPledgeDate = tlDatPledgeInfo.sFeedDate
    If mGetPledgeByEvent(tlDatPledgeInfo.iVefCode) <> "Y" Then
        If tlDatPledgeInfo.lDatCode > 0 Then
            'slSQLQuery = "SELECT * FROM dat WHERE (datCode = " & tlDatPledgeInfo.lDatCode & ")"
            'Set DatPledge_rst = gSQLSelectCall(slSQLQuery)
            'If Not DatPledge_rst.EOF Then
            If lmAttForDat <> tlDatPledgeInfo.lAttCode Then
                lmAttForDat = tlDatPledgeInfo.lAttCode
                llDat = 0
                ReDim tmDatRst(0 To 100) As DATRST
                slSQLQuery = "SELECT * FROM dat WHERE"
                slSQLQuery = slSQLQuery & " datAtfCode = " & lmAttForDat
                Set DatPledge_rst = gSQLSelectCall(slSQLQuery)
                Do While Not DatPledge_rst.EOF
                    gCreateUDTForDat DatPledge_rst, tmDatRst(llDat)
                    llDat = llDat + 1
                    If llDat >= UBound(tmDatRst) Then
                        ReDim Preserve tmDatRst(0 To UBound(tmDatRst) + 30) As DATRST
                    End If
                    DatPledge_rst.MoveNext
                Loop
                ReDim Preserve tmDatRst(0 To llDat) As DATRST
                If UBound(tmDatRst) - 1 >= 1 Then
                    ArraySortTyp fnAV(tmDatRst(), 0), UBound(tmDatRst), 0, LenB(tmDatRst(0)), 0, -2, 0
                End If
            End If
            llDat = mBinarySearchDatRst(tlDatPledgeInfo.lDatCode)
            If llDat <> -1 Then
                blPledgeDefined = True
                tlDatPledgeInfo.iPledgeStatus = tmDatRst(llDat).iFdStatus 'DatPledge_rst!datFdStatus
                If tgStatusTypes(tlDatPledgeInfo.iPledgeStatus).iPledged <= 1 Then
                    tlDatPledgeInfo.sPledgeStartTime = tmDatRst(llDat).sPdStTime 'Format$(CStr(DatPledge_rst!datPdStTime), sgShowTimeWSecForm)
                    tlDatPledgeInfo.sPledgeEndTime = tmDatRst(llDat).sPdEdTime 'Format$(CStr(DatPledge_rst!datPdEdTime), sgShowTimeWSecForm)
                Else
                    'tlDatPledgeInfo.sPledgeStartTime = ""
                    'tlDatPledgeInfo.sPledgeEndTime = ""
                    tlDatPledgeInfo.sPledgeStartTime = tlDatPledgeInfo.sFeedTime
                    tlDatPledgeInfo.sPledgeEndTime = tlDatPledgeInfo.sFeedTime
                End If
                '5/29/15: Fix case where pledge day is offset from the feed day
                If tgStatusTypes(tmDatRst(llDat).iFdStatus).iPledged = 1 Then
                    If tmDatRst(llDat).iFdMon <> 0 Then
                        ilFdDay = 0
                    ElseIf tmDatRst(llDat).iFdTue <> 0 Then
                        ilFdDay = 1
                    ElseIf tmDatRst(llDat).iFdWed <> 0 Then
                        ilFdDay = 2
                    ElseIf tmDatRst(llDat).iFdThu <> 0 Then
                        ilFdDay = 3
                    ElseIf tmDatRst(llDat).iFdFri <> 0 Then
                        ilFdDay = 4
                    ElseIf tmDatRst(llDat).iFdSat <> 0 Then
                        ilFdDay = 5
                    ElseIf tmDatRst(llDat).iFdSun <> 0 Then
                        ilFdDay = 6
                    End If
                    If tmDatRst(llDat).iPdMon <> 0 Then
                        ilPdDay = 0
                    ElseIf tmDatRst(llDat).iPdTue <> 0 Then
                        ilPdDay = 1
                    ElseIf tmDatRst(llDat).iPdWed <> 0 Then
                        ilPdDay = 2
                    ElseIf tmDatRst(llDat).iPdThu <> 0 Then
                        ilPdDay = 3
                    ElseIf tmDatRst(llDat).iPdFri <> 0 Then
                        ilPdDay = 4
                    ElseIf tmDatRst(llDat).iPdSat <> 0 Then
                        ilPdDay = 5
                    ElseIf tmDatRst(llDat).iPdSun <> 0 Then
                        ilPdDay = 6
                    End If
                    ilAdjDay = 0
                    If ilPdDay >= ilFdDay Then
                        ilAdjDay = ilPdDay - ilFdDay
                        If tmDatRst(llDat).sPdDayFed = "B" Then
                            ilAdjDay = ilAdjDay - 7
                        End If
                    Else
                        If tmDatRst(llDat).sPdDayFed = "B" Then
                            ilAdjDay = ilPdDay - ilFdDay
                            If ilAdjDay > 0 Then
                                ilAdjDay = ilAdjDay - 1
                            End If
                        Else
                            ilAdjDay = 7 + ilPdDay - ilFdDay
                        End If
                    End If
                    tlDatPledgeInfo.sPledgeDate = Format$(DateValue(gAdjYear(tlDatPledgeInfo.sFeedDate)) + ilAdjDay, sgShowDateForm)
                End If
            Else
                'If Pledge not defined, then live
                'If pledges defined, then not caried
                tlDatPledgeInfo.iPledgeStatus = 8   'Flag to be set later once all ast read in
                tlDatPledgeInfo.sPledgeStartTime = tlDatPledgeInfo.sFeedTime
                tlDatPledgeInfo.sPledgeEndTime = tlDatPledgeInfo.sFeedTime
            End If
        ElseIf tlDatPledgeInfo.lDatCode < 0 Then
            tlDatPledgeInfo.iPledgeStatus = 8
            tlDatPledgeInfo.sPledgeStartTime = tlDatPledgeInfo.sFeedTime
            tlDatPledgeInfo.sPledgeEndTime = tlDatPledgeInfo.sFeedTime
        Else
            tlDatPledgeInfo.iPledgeStatus = 0
            tlDatPledgeInfo.sPledgeStartTime = tlDatPledgeInfo.sFeedTime
            tlDatPledgeInfo.sPledgeEndTime = tlDatPledgeInfo.sFeedTime
        End If
    Else
        'Live
        tlDatPledgeInfo.iPledgeStatus = 0
        tlDatPledgeInfo.sPledgeStartTime = tlDatPledgeInfo.sFeedTime
        tlDatPledgeInfo.sPledgeEndTime = tlDatPledgeInfo.sFeedTime
    End If

    On Error Resume Next
    gGetPledgeGivenAstInfo = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "gGetPledgeGivenAstInfo"
    gGetPledgeGivenAstInfo = False

End Function

Function mBinarySearchBuildLst(llCode As Long) As Long
    
    'D.S. 01/16/06
    'Returns the index number of tmBuildLst that matches the lstCode that was passed in
    'Note: for this to work tglsttInfo was previously sorted
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tmBuildLst)
    llMax = UBound(tmBuildLst) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tmBuildLst(llMiddle).tLST.lCode Then
            'found the match
            mBinarySearchBuildLst = llMiddle
            Exit Function
        ElseIf llCode < tmBuildLst(llMiddle).tLST.lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchBuildLst = -1
    Exit Function
End Function

Function mBinarySearchDatRst(llCode As Long) As Long
    
    'D.S. 01/16/06
    'Returns the index number of tmDatRst that matches the lstCode that was passed in
    'Note: for this to work tglsttInfo was previously sorted
    
    Dim llMin As Long
    Dim llMax As Long
    Dim llMiddle As Long
    
    llMin = LBound(tmDatRst)
    llMax = UBound(tmDatRst) - 1
    Do While llMin <= llMax
        llMiddle = (llMin + llMax) \ 2
        If llCode = tmDatRst(llMiddle).lCode Then
            'found the match
            mBinarySearchDatRst = llMiddle
            Exit Function
        ElseIf llCode < tmDatRst(llMiddle).lCode Then
            llMax = llMiddle - 1
        Else
            'search the right half
            llMin = llMiddle + 1
        End If
    Loop
    mBinarySearchDatRst = -1
    Exit Function
End Function

Public Function gFilterAstInfoBySalesSource(tlAstInfo() As ASTINFO) As Integer

    Dim iUpper As Integer
    Dim iAst As Integer

    On Error GoTo ErrHand:
    gFilterAstInfoBySalesSource = False
    If (igUstSSMnfCode <> 0) Then
        ReDim tlAstTemp(0 To UBound(tlAstInfo)) As ASTINFO
        iUpper = 0
        For iAst = 0 To UBound(tlAstInfo) - 1 Step 1
            If gCheckSalesSource(tlAstInfo(iAst)) Then
                tlAstTemp(iUpper) = tlAstInfo(iAst)
                iUpper = iUpper + 1
            End If
        Next iAst
        ReDim Preserve tlAstTemp(0 To iUpper) As ASTINFO
        ReDim tlAstInfo(0 To UBound(tlAstTemp)) As ASTINFO
        For iAst = 0 To UBound(tlAstTemp) - 1 Step 1
            tlAstInfo(iAst) = tlAstTemp(iAst)
        Next iAst
    End If
    gFilterAstInfoBySalesSource = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "modCPReturns-gFilterAstInfo"
    gFilterAstInfoBySalesSource = False
    Exit Function
    
End Function


Public Function gCheckSalesSource(tlAstInfo As ASTINFO) As Integer

    Dim slSql As String
    Dim ilRet As Integer
    
    On Error GoTo ErrHand:
    gCheckSalesSource = False
    If (igUstSSMnfCode <> 0) Then
        If tlAstInfo.lCntrNo <> 0 Then
            slSql = " Select chfSlfCode1 from CHF_Contract_Header where (chfSchStatus = 'F' or chfschstatus = 'M') AND chfDelete = 'N' AND  chfCntrNo = " & tlAstInfo.lCntrNo
            Set rst = gSQLSelectCall(slSql)
            If Not rst.EOF Then
                ilRet = gBinarySalesPeopleInfo(rst!chfSlfCode1)
                If ilRet <> -1 Then
                    If igUstSSMnfCode = tgSalesPeopleInfo(ilRet).iSSMnfCode Then
                        gCheckSalesSource = True
                    End If
                End If
            End If
        End If
    Else
        gCheckSalesSource = True
    End If
ErrHand:
    gHandleError "AffErrorLog.txt", "modCPReturns-gCheckSalesSource"
    gCheckSalesSource = False
    Exit Function
End Function

Public Function gFilterAstInfoByAirDate(tlAstInfo() As ASTINFO, llFWkDate As Long, llLWkDate As Long) As Integer

    Dim iUpper As Integer
    Dim iAst As Integer

    On Error GoTo ErrHand:
    gFilterAstInfoByAirDate = False
    ReDim tlAstTemp(0 To UBound(tlAstInfo)) As ASTINFO
    iUpper = 0
    For iAst = 0 To UBound(tlAstInfo) - 1 Step 1
        If (gDateValue(tlAstInfo(iAst).sAirDate) >= llFWkDate) And (gDateValue(tlAstInfo(iAst).sAirDate) <= llLWkDate) Then
            tlAstTemp(iUpper) = tlAstInfo(iAst)
            iUpper = iUpper + 1
        End If
    Next iAst
    ReDim Preserve tlAstTemp(0 To iUpper) As ASTINFO
    ReDim tlAstInfo(0 To UBound(tlAstTemp)) As ASTINFO
    For iAst = 0 To UBound(tlAstTemp) - 1 Step 1
        tlAstInfo(iAst) = tlAstTemp(iAst)
    Next iAst
    gFilterAstInfoByAirDate = True
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "modCPReturns-gFilterAstInfoByAirDate"
    gFilterAstInfoByAirDate = False
    Exit Function
    
End Function

Public Function gCreateLockRec(slType As String, slSubType As String, llRecCode As Long, ilRetryFlag As Integer, slUserNameWithLock As String) As Long
    Dim slNowDate As String
    Dim slNowTime As String
    Dim llRet As Long
    Dim llDate As Long
    Dim ilLoop As Integer
    Dim ilCount As Integer
    Dim ilForceRetry As Integer
    Dim llCode As Long
    Dim slSQLQuery As String

    'Create Lock
    On Error GoTo ErrHand:
    ilCount = 0
    Do
        ilForceRetry = False
        slUserNameWithLock = ""
        
        slSQLQuery = "Insert Into RLF_Record_Lock ( "
        slSQLQuery = slSQLQuery & "rlfCode, "
        slSQLQuery = slSQLQuery & "rlfUrfCode, "
        slSQLQuery = slSQLQuery & "rlfType, "
        slSQLQuery = slSQLQuery & "rlfSubType, "
        slSQLQuery = slSQLQuery & "rlfRecCode, "
        slSQLQuery = slSQLQuery & "rlfEnteredDate, "
        slSQLQuery = slSQLQuery & "rlfEnteredTime, "
        slSQLQuery = slSQLQuery & "rlfUnused "
        slSQLQuery = slSQLQuery & ") "
        slSQLQuery = slSQLQuery & "Values ( "
        slSQLQuery = slSQLQuery & "Replace" & ", "
        slSQLQuery = slSQLQuery & igUstCode & ", "
        slSQLQuery = slSQLQuery & "'" & gFixQuote(slType) & "', "
        slSQLQuery = slSQLQuery & "'" & gFixQuote(slSubType) & "', "
        slSQLQuery = slSQLQuery & llRecCode & ", "
        slSQLQuery = slSQLQuery & "'" & Format$(gNow(), sgSQLDateForm) & "', "
        slSQLQuery = slSQLQuery & "'" & Format$(gNow(), sgSQLTimeForm) & "', "
        slSQLQuery = slSQLQuery & "'" & "" & "' "
        slSQLQuery = slSQLQuery & ") "
        llCode = gInsertAndReturnCode(slSQLQuery, "RLF_Record_Lock", "rlfCode", "Replace", True)
        If (llCode = -1) Then
            slSQLQuery = "SELECT * FROM RLF_Record_Lock"
            slSQLQuery = slSQLQuery + " WHERE (rlfType = '" & slType & "' And rlfRecCode = " & llRecCode & ")"
            Set rst_Rlf = gSQLSelectCall(slSQLQuery)
            If Not rst_Rlf.EOF Then
                slNowDate = Format$(gNow(), "m/d/yy")
                slNowTime = Format$(gNow(), "h:mm:ssAM/PM")
                'Check if longer then Day for lock, if so remove it
                llDate = gDateValue(rst_Rlf!rlfEnteredDate)
                '6/30/06:  Remove checking for same user.  This is being removed to aviod the problem
                '          where the same user is running schedule on two terminals
                'If (llDate < gDateValue(slNowDate)) Or (tlRlf.iurfCode = tgUrf(0).iCode) Then
                If (llDate < gDateValue(slNowDate)) Then
                    slSQLQuery = "DELETE FROM RLF_Record_Lock WHERE (rlfCode = " & rst_Rlf!rlfCode & ")"
                    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
                        '6/13/16: Replaced GoSub
                        'GoSub ErrHand1:
                        gHandleError "AffErrorLog.txt", "modCPReturns-gCreateLockRec"
                        gCreateLockRec = 0
                        Exit Function
                    End If
                    ilForceRetry = True
                Else
                    slSQLQuery = "SELECT * FROM Ust Where ustCode = " & rst_Rlf!rlfUrfCode
                    Set rst_Ust = gSQLSelectCall(slSQLQuery)
                    If Not rst_Ust.EOF Then
                        If Trim$(rst_Ust!ustReportName) <> "" Then
                            slUserNameWithLock = Trim$(rst_Ust!ustReportName)
                        Else
                            slUserNameWithLock = Trim$(rst_Ust!ustname)
                        End If
                    End If
                    If Not ilRetryFlag Then
                        gCreateLockRec = 0
                        Exit Function
                    End If
                End If
            Else
                ilForceRetry = True
            End If
        ElseIf llRet = BTRV_ERR_DEADLOCK_DETECT Then
            ilForceRetry = True
        ElseIf llRet = BTRV_ERR_NONE Then
            gCreateLockRec = llCode
            Exit Function
        End If
        If Not ilForceRetry Then
            If Not ilRetryFlag Then
                gCreateLockRec = 0
                Exit Function
            End If
        End If
        Sleep 100
        ilCount = ilCount + 1
    Loop While ilCount < 20
    gCreateLockRec = 0
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "gCreateLockRec"
    gCreateLockRec = 0
    Exit Function
ErrHand1:
    gHandleError "AffErrorLog.txt", "gCreateLockRec"
    Return
End Function

Public Function gDeleteLockRec_ByRlfCode(llRlfCode As Long) As Integer
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand:
    If (llRlfCode > 0) Then
        slSQLQuery = "DELETE FROM RLF_Record_Lock WHERE (rlfCode = " & llRlfCode & ")"
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            gDeleteLockRec_ByRlfCode = False
        Else
            gDeleteLockRec_ByRlfCode = True
            llRlfCode = 0
        End If
        
    Else
        gDeleteLockRec_ByRlfCode = True
        llRlfCode = 0
    End If
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "gDeleteLockRec_ByRlfCode"
    gDeleteLockRec_ByRlfCode = False
    Exit Function
End Function
Public Function gDeleteLockRec_ByType(slType As String, llRecCode As Long) As Integer
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand:
    slSQLQuery = "DELETE FROM RLF_Record_Lock WHERE (rlfType = '" & slType & "' And rlfRecCode = " & llRecCode & ")"
    If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
        gDeleteLockRec_ByType = False
    Else
        gDeleteLockRec_ByType = True
    End If
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.txt", "gDeleteLockRec_ByType"
    gDeleteLockRec_ByType = False
    Exit Function
End Function


Private Function mCreateBlackoutLst(hlAst As Integer, tlAstInfo As ASTINFO, ilVefCode As Integer, llLstCode As Long, ilCifAdfCode As Integer, slCartNo As String, slProduct As String, slISCI As String, slLogDate As String, slLogTime As String, llCopyCode As Long, llCpfCode As Long, llCrfCsfCode As Long, llRafCode As Long) As Long
    Dim llCode As Long
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim slSQLQuery As String
    Dim ilWeekDay(0 To 6) As Integer
    
    For ilLoop = 0 To 6 Step 1
        ilWeekDay(ilLoop) = 0
    Next ilLoop
    ilWeekDay(Weekday(slLogDate, vbMonday) - 1) = 1
    ilRet = 0
    slSQLQuery = "Insert Into lst ( "
    slSQLQuery = slSQLQuery & "lstCode, "
    slSQLQuery = slSQLQuery & "lstType, "
    slSQLQuery = slSQLQuery & "lstSdfCode, "
    slSQLQuery = slSQLQuery & "lstCntrNo, "
    slSQLQuery = slSQLQuery & "lstAdfCode, "
    slSQLQuery = slSQLQuery & "lstAgfCode, "
    slSQLQuery = slSQLQuery & "lstProd, "
    slSQLQuery = slSQLQuery & "lstLineNo, "
    slSQLQuery = slSQLQuery & "lstLnVefCode, "
    slSQLQuery = slSQLQuery & "lstStartDate, "
    slSQLQuery = slSQLQuery & "lstEndDate, "
    slSQLQuery = slSQLQuery & "lstMon, "
    slSQLQuery = slSQLQuery & "lstTue, "
    slSQLQuery = slSQLQuery & "lstWed, "
    slSQLQuery = slSQLQuery & "lstThu, "
    slSQLQuery = slSQLQuery & "lstFri, "
    slSQLQuery = slSQLQuery & "lstSat, "
    slSQLQuery = slSQLQuery & "lstSun, "
    slSQLQuery = slSQLQuery & "lstSpotsWk, "
    slSQLQuery = slSQLQuery & "lstPriceType, "
    slSQLQuery = slSQLQuery & "lstPrice, "
    slSQLQuery = slSQLQuery & "lstSpotType, "
    slSQLQuery = slSQLQuery & "lstLogVefCode, "
    slSQLQuery = slSQLQuery & "lstLogDate, "
    slSQLQuery = slSQLQuery & "lstLogTime, "
    slSQLQuery = slSQLQuery & "lstDemo, "
    slSQLQuery = slSQLQuery & "lstAud, "
    slSQLQuery = slSQLQuery & "lstISCI, "
    slSQLQuery = slSQLQuery & "lstWkNo, "
    slSQLQuery = slSQLQuery & "lstBreakNo, "
    slSQLQuery = slSQLQuery & "lstPositionNo, "
    slSQLQuery = slSQLQuery & "lstSeqNo, "
    slSQLQuery = slSQLQuery & "lstZone, "
    slSQLQuery = slSQLQuery & "lstCart, "
    slSQLQuery = slSQLQuery & "lstCpfCode, "
    slSQLQuery = slSQLQuery & "lstCrfCsfCode, "
    slSQLQuery = slSQLQuery & "lstStatus, "
    slSQLQuery = slSQLQuery & "lstLen, "
    slSQLQuery = slSQLQuery & "lstUnits, "
    slSQLQuery = slSQLQuery & "lstCifCode, "
    slSQLQuery = slSQLQuery & "lstAnfCode, "
    slSQLQuery = slSQLQuery & "lstEvtIDCefCode, "
    slSQLQuery = slSQLQuery & "lstSplitNetwork, "
    slSQLQuery = slSQLQuery & "lstRafCode, "
    slSQLQuery = slSQLQuery & "lstFsfCode, "
    slSQLQuery = slSQLQuery & "lstGsfCode, "
    slSQLQuery = slSQLQuery & "lstImportedSpot, "
    slSQLQuery = slSQLQuery & "lstBkoutLstCode, "
    slSQLQuery = slSQLQuery & "lstLnStartTime, "
    slSQLQuery = slSQLQuery & "lstLnEndTime, "
    slSQLQuery = slSQLQuery & "lstUnused "
    slSQLQuery = slSQLQuery & ") "
    slSQLQuery = slSQLQuery & "Values ( "
    slSQLQuery = slSQLQuery & "Replace" & ", "
    slSQLQuery = slSQLQuery & lst_rst!lstType & ", "
    slSQLQuery = slSQLQuery & lst_rst!lstSdfCode & ", "
    slSQLQuery = slSQLQuery & 0 & ", "
    slSQLQuery = slSQLQuery & ilCifAdfCode & ", "
    slSQLQuery = slSQLQuery & 0 & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(slProduct) & "', "
    slSQLQuery = slSQLQuery & 0 & ", "
    slSQLQuery = slSQLQuery & 0 & ", "
    slSQLQuery = slSQLQuery & "'" & Format$(slLogDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & "'" & Format$(slLogDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & ilWeekDay(0) & ", "
    slSQLQuery = slSQLQuery & ilWeekDay(1) & ", "
    slSQLQuery = slSQLQuery & ilWeekDay(2) & ", "
    slSQLQuery = slSQLQuery & ilWeekDay(3) & ", "
    slSQLQuery = slSQLQuery & ilWeekDay(4) & ", "
    slSQLQuery = slSQLQuery & ilWeekDay(5) & ", "
    slSQLQuery = slSQLQuery & ilWeekDay(6) & ", "
    slSQLQuery = slSQLQuery & 1 & ", "
    slSQLQuery = slSQLQuery & 1 & ", "
    slSQLQuery = slSQLQuery & 0 & ", "
    slSQLQuery = slSQLQuery & 5 & ", "
    slSQLQuery = slSQLQuery & ilVefCode & ", "
    slSQLQuery = slSQLQuery & "'" & Format$(slLogDate, sgSQLDateForm) & "', "
    slSQLQuery = slSQLQuery & "'" & Format$(slLogTime, sgSQLTimeForm) & "', "
    slSQLQuery = slSQLQuery & "'" & "" & "', "
    slSQLQuery = slSQLQuery & 0 & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(slISCI) & "', "
    slSQLQuery = slSQLQuery & lst_rst!lstWkNo & ", "
    slSQLQuery = slSQLQuery & lst_rst!lstBreakNo & ", "
    slSQLQuery = slSQLQuery & lst_rst!lstPositionNo & ", "
    slSQLQuery = slSQLQuery & lst_rst!lstSeqNo & ", "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(lst_rst!lstZone) & "', "
    slSQLQuery = slSQLQuery & "'" & gFixQuote(slCartNo) & "', "
    slSQLQuery = slSQLQuery & llCpfCode & ", "
    slSQLQuery = slSQLQuery & llCrfCsfCode & ", "
    slSQLQuery = slSQLQuery & lst_rst!lstStatus & ", "
    slSQLQuery = slSQLQuery & lst_rst!lstLen & ", "
    slSQLQuery = slSQLQuery & lst_rst!lstUnits & ", "
    slSQLQuery = slSQLQuery & llCopyCode & ","    'tmRegionDefinition(llRegionIndex).lCopyCode & ", "
    slSQLQuery = slSQLQuery & lst_rst!lstAnfCode & ", "
    slSQLQuery = slSQLQuery & lst_rst!lstEvtIDCefCode & ", "
    If (Asc(lst_rst!lstsplitnetwork) <> Asc("N")) And (Asc(lst_rst!lstsplitnetwork) <> Asc("Y")) Then
        slSQLQuery = slSQLQuery & "'" & "N" & "', "
    Else
        slSQLQuery = slSQLQuery & "'" & gFixQuote(lst_rst!lstsplitnetwork) & "', "
    End If
    slSQLQuery = slSQLQuery & llRafCode & ", "  'tmRegionDefinition(llRegionIndex).lRafCode & ", "
    slSQLQuery = slSQLQuery & lst_rst!lstFsfCode & ", "
    slSQLQuery = slSQLQuery & lst_rst!lstGsfCode & ", "
    If (Asc(lst_rst!lstImportedSpot) <> Asc("N")) And (Asc(lst_rst!lstImportedSpot) <> Asc("Y")) Then
        slSQLQuery = slSQLQuery & "'" & "N" & "', "  'gFixQuote(lst_rst!lstImportedSpot) & "', "
    Else
        slSQLQuery = slSQLQuery & "'" & gFixQuote(lst_rst!lstImportedSpot) & "', "
    End If
    slSQLQuery = slSQLQuery & llLstCode & ", "
    slSQLQuery = slSQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
    slSQLQuery = slSQLQuery & "'" & Format$("12am", sgSQLTimeForm) & "', "
    slSQLQuery = slSQLQuery & "'" & "" & "' "
    slSQLQuery = slSQLQuery & ") "
'                                If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
    bgIgnoreDuplicateError = True
    llCode = gInsertAndReturnCode(slSQLQuery, "lst", "lstCode", "Replace")
    bgIgnoreDuplicateError = False
    If llCode <= 0 Then
        mCreateBlackoutLst = -1
        Exit Function
    End If
    lgSTime6 = timeGetTime
    tmAstSrchKey.lCode = tlAstInfo.lCode
    ilRet = btrGetEqual(hlAst, tmAst, imAstRecLen, tmAstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, 0)
    If ilRet = BTRV_ERR_NONE Then
        tmAst.lLsfCode = llCode
        ilRet = btrUpdate(hlAst, tmAst, imAstRecLen)
    Else
        slSQLQuery = "UPDATE ast SET astLsfCode = " & llCode & " where astCode = " & tlAstInfo.lCode
        'cnn.Execute slSQLQuery
        If gSQLWaitNoMsgBox(slSQLQuery, False) <> 0 Then
            mCreateBlackoutLst = False
            Exit Function
        End If
    End If
    lgETime6 = timeGetTime
    lgTtlTime6 = lgTtlTime6 + (lgETime6 - lgSTime6)
    tlAstInfo.lLstBkoutLstCode = tlAstInfo.lLstCode
    tlAstInfo.lLstCode = llCode
    tlAstInfo.iAdfCode = ilCifAdfCode
    tlAstInfo.sProd = slProduct
    tlAstInfo.lCntrNo = 0
    slSQLQuery = "SELECT * FROM lst WHERE (lstCode = " & llCode & ")"
    Set lst_rst = gSQLSelectCall(slSQLQuery)
    If Not lst_rst.EOF Then
        gCreateUDTforLST lst_rst, tmBkoutLst(UBound(tmBkoutLst)).tLST
        tmBkoutLst(UBound(tmBkoutLst)).iDelete = False
        '12/12/14
        tmBkoutLst(UBound(tmBkoutLst)).bMatched = False
        ReDim Preserve tmBkoutLst(0 To UBound(tmBkoutLst) + 1) As BKOUTLST
    End If
    'gGetRegionCopy = 2
    lgETime10 = timeGetTime
    lgTtlTime10 = lgTtlTime10 + (lgETime10 - lgSTime10)
    If igExportSource = 2 Then
        DoEvents
    End If
    mCreateBlackoutLst = llCode
    Exit Function
ErrHand:
    gHandleError "AffErrorLog.Txt", "CPReturn-mCreateBlackoutLst"
    'Resume Next
    mCreateBlackoutLst = -1
    Exit Function
ErrHand1:
    gHandleError "AffErrorLog.Txt", "CPReturn-mCreateBlackoutLst"
    'Return
    mCreateBlackoutLst = -1
End Function

Public Sub gGetAndAssignRegionToAst(hlAst As Integer, tlAstInfo As ASTINFO)
    'tlAstInfo comes in with information.  Get AttCode and AstCode
    '9452 many changes
    Dim sFWkDate As String
    Dim sLWkDate As String
    Dim lFWkDate As Long
    Dim lLWkDate As Long
    Dim iUpper As Integer
    Dim llSdfCode As Long
    Dim ilFound As Integer
    Dim ilLoop As Integer
    Dim iAst As Integer
    Dim slForbidSplitLive As String
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand
    iUpper = 0
    ReDim Preserve lmRCSdfCode(0 To 0) As Long
    ReDim tmBkoutLst(0 To 0) As BKOUTLST
    slSQLQuery = "SELECT *"
    slSQLQuery = slSQLQuery + " FROM att"
'    slSQLQuery = slSQLQuery + " WHERE (attCode= " & tgCPPosting(0).lAttCode & ")"
    slSQLQuery = slSQLQuery + " WHERE (attCode= " & tlAstInfo.lAttCode & ")"
    Set rst = gSQLSelectCall(slSQLQuery)
    If rst.EOF Then
        slForbidSplitLive = "N"
    Else
        slForbidSplitLive = rst!attForbidSplitLive
    End If
    sFWkDate = Format$(gObtainPrevMonday(tgCPPosting(0).sDate), sgShowDateForm)
    sLWkDate = Format$(gObtainNextSunday(tgCPPosting(0).sDate), sgShowDateForm)
    lFWkDate = DateValue(gAdjYear(sFWkDate))
    lLWkDate = DateValue(gAdjYear(sLWkDate))
    '9452
   ' slSQLQuery = "SELECT astAtfCode, astShfCode, astVefCode, astlsfCode, astAirDate, astAirTime, astStatus, astCPStatus, astFeedDate, astFeedTime, astPledgeDate, astPledgeStartTime, astPledgeEndTime, astPledgeStatus, astSdfCode, astCode FROM ast"
    slSQLQuery = "SELECT astAtfCode, astShfCode, astVefCode, astlsfCode, astAirDate, astAirTime, astStatus, astCPStatus, astFeedDate, astFeedTime, datPdStTime as astPledgeStartTime, datPdEdTime as astPledgeEndTime, astStatus as astPledgeStatus, astSdfCode, astCode,astdatcode FROM ast left outer join dat on astdatcode = datcode "
    slSQLQuery = slSQLQuery + " WHERE (astCode= " & tlAstInfo.lCode
    slSQLQuery = slSQLQuery + " AND (astFeedDate >= '" & Format$(sFWkDate, sgSQLDateForm) & "' AND astFeedDate <= '" & Format$(sLWkDate, sgSQLDateForm) & "')" & ")"
    slSQLQuery = slSQLQuery + " ORDER BY astFeedDate, astFeedTime"
    Set rst = gSQLSelectCall(slSQLQuery)
    While Not rst.EOF
        'tlAstInfo.lCode = rst!astCode
        tlAstInfo.lLstCode = rst!astLsfCode
        tlAstInfo.iStatus = rst!astStatus
        tlAstInfo.iCPStatus = rst!astCPStatus
        tlAstInfo.sAirDate = Format$(rst!astAirDate, sgShowDateForm)
        If Second(rst!astAirTime) <> 0 Then
            tlAstInfo.sAirTime = Format$(rst!astAirTime, sgShowTimeWSecForm)
        Else
            tlAstInfo.sAirTime = Format$(rst!astAirTime, sgShowTimeWOSecForm)
        End If
        tlAstInfo.sFeedDate = Format$(rst!astFeedDate, sgShowDateForm)
        If Second(rst!astFeedTime) <> 0 Then
            tlAstInfo.sFeedTime = Format$(rst!astFeedTime, sgShowTimeWSecForm)
        Else
            tlAstInfo.sFeedTime = Format$(rst!astFeedTime, sgShowTimeWOSecForm)
        End If
        '9452
      '  tlAstInfo(iUpper).sPledgeDate = Format$(rst!astPledgeDate, sgShowDateForm)
        tlAstInfo.sPledgeDate = Format$(rst!astFeedDate, sgShowDateForm)
        If Second(rst!astPledgeStartTime) <> 0 Then
            tlAstInfo.sPledgeStartTime = Format$(rst!astPledgeStartTime, sgShowTimeWSecForm)
        Else
            tlAstInfo.sPledgeStartTime = Format$(rst!astPledgeStartTime, sgShowTimeWOSecForm)
        End If
        If Not IsNull(rst!astPledgeEndTime) Then
            If Second(rst!astPledgeEndTime) <> 0 Then
                tlAstInfo.sPledgeEndTime = Format$(rst!astPledgeEndTime, sgShowTimeWSecForm)
            Else
                tlAstInfo.sPledgeEndTime = Format$(rst!astPledgeEndTime, sgShowTimeWOSecForm)
            End If
        Else
            tlAstInfo.sPledgeEndTime = ""
        End If
        tlAstInfo.iPledgeStatus = rst!astPledgeStatus
        tlAstInfo.lAttCode = rst!astAtfCode
        tlAstInfo.iShttCode = rst!astShfCode
        tlAstInfo.iVefCode = rst!astVefCode
        tlAstInfo.lSdfCode = rst!astSdfCode
        tlAstInfo.lDatCode = rst!astDatCode ' 9452 changed from  0
        '11/18/11: Set to blackout that was previous defined as a replacement.
        '          astLstCode could be referecing the Blackout currently assigned
        '          each ast will reference a unique blackout lst.  Those lst will reference the original lst from traffic
        '          the field is set later if required
        tlAstInfo.lPrevBkoutLstCode = 0
        tlAstInfo.iComp = 0
'        iUpper = iUpper + 1
'        ReDim Preserve tlAstInfo(0 To iUpper) As ASTINFO
        llSdfCode = rst!astSdfCode
        ilFound = False
        For ilLoop = 0 To UBound(lmRCSdfCode) - 1 Step 1
            If lmRCSdfCode(ilLoop) = llSdfCode Then
                ilFound = True
                Exit For
            End If
        Next ilLoop
        If Not ilFound Then
            lmRCSdfCode(UBound(lmRCSdfCode)) = llSdfCode
            ReDim Preserve lmRCSdfCode(0 To UBound(lmRCSdfCode) + 1) As Long
        End If
        rst.MoveNext
    Wend
    slSQLQuery = "SELECT * FROM lst"
    slSQLQuery = slSQLQuery + " WHERE (lstLogVefCode = " & tlAstInfo.iVefCode
    slSQLQuery = slSQLQuery + " AND lstBkoutLstCode <> 0"
    slSQLQuery = slSQLQuery & " AND lstType <> 1"
    slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(lFWkDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(lLWkDate + 1, sgSQLDateForm) & "')" & ")"
    slSQLQuery = slSQLQuery + " ORDER BY lstLogDate, lstLogTime, lstBreakNo, lstPositionNo"
    Set rst = gSQLSelectCall(slSQLQuery)
    Do While Not rst.EOF
        gCreateUDTforLST rst, tmBkoutLst(UBound(tmBkoutLst)).tLST
        tmBkoutLst(UBound(tmBkoutLst)).iDelete = False
        '10/12/14
        tmBkoutLst(UBound(tmBkoutLst)).bMatched = False
        ReDim Preserve tmBkoutLst(0 To UBound(tmBkoutLst) + 1) As BKOUTLST
        rst.MoveNext
    Loop
   ' mBuildRegionForSpots tgCPPosting(0).iVefCode
    mBuildRegionForSpots tlAstInfo.iVefCode
    If (slForbidSplitLive <> "Y") Or ((slForbidSplitLive = "Y") And (tgStatusTypes(gGetAirStatus(tlAstInfo.iPledgeStatus)).iPledged <> 0)) Then
        tlAstInfo.iRegionType = gGetRegionCopy(hlAst, True, tlAstInfo, tlAstInfo.sRCart, tlAstInfo.sRProduct, tlAstInfo.sRISCI, tlAstInfo.sRCreativeTitle, tlAstInfo.lRCrfCsfCode, tlAstInfo.lRCrfCode, tlAstInfo.lRCifCode, tlAstInfo.lRRsfCode, tlAstInfo.lRCpfCode, tlAstInfo.sReplacementCue)
    End If
'    For iAst = 0 To UBound(tlAstInfo) - 1 Step 1
'        If igExportSource = 2 Then
'            DoEvents
'        End If
'        lgCount2 = lgCount2 + 1
'        If tgStatusTypes(gGetAirStatus(tlAstInfo(iAst).iPledgeStatus)).iPledged <> 2 Then
'            If (slForbidSplitLive <> "Y") Or ((slForbidSplitLive = "Y") And (tgStatusTypes(gGetAirStatus(tlAstInfo(iAst).iPledgeStatus)).iPledged <> 0)) Then
'                tlAstInfo(iAst).iRegionType = gGetRegionCopy(hlAst, True, tlAstInfo(iAst), tlAstInfo(iAst).sRCart, tlAstInfo(iAst).sRProduct, tlAstInfo(iAst).sRISCI, tlAstInfo(iAst).sRCreativeTitle, tlAstInfo(iAst).lRCrfCsfCode, tlAstInfo(iAst).lRCrfCode, tlAstInfo(iAst).lRCifCode, tlAstInfo(iAst).lRRsfCode, tlAstInfo(iAst).lRCpfCode, tlAstInfo(iAst).sReplacementCue)
'            End If
'        End If
'    Next iAst
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    gHandleError "", "gGetAndAssignRegionToAst"
    Resume Next
End Sub

Public Function gDeterminePledgeDateTime(dat_rst As ADODB.Recordset, llDATCode As Long, slFdDate As String, slFdTime As String, slAirDate As String, slPdDate As String, slPdTime As String) As Boolean
    Dim ilDay As Integer
    Dim ilPledged As Integer
    Dim ilPdDay As Integer
    Dim ilFdDay As Integer
    Dim ilAdjDay As Integer
    Dim slSQLQuery As String
    Dim tlDatRst As DATRST
    
    gDeterminePledgeDateTime = False
    slPdDate = slFdDate
    slPdTime = slFdTime
    If llDATCode > 0 Then
        slSQLQuery = "Select * From Dat Where datCode = " & llDATCode
        Set dat_rst = gSQLSelectCall(slSQLQuery)
        If Not dat_rst.EOF Then
            gCreateUDTForDat dat_rst, tlDatRst
            ilPledged = tgStatusTypes(tlDatRst.iFdStatus).iPledged
            If ilPledged = 1 Then   'Delay
                If tlDatRst.iFdMon <> 0 Then
                    ilFdDay = 0
                ElseIf tlDatRst.iFdTue Then
                    ilFdDay = 1
                ElseIf tlDatRst.iFdWed Then
                    ilFdDay = 2
                ElseIf tlDatRst.iFdThu Then
                    ilFdDay = 3
                ElseIf tlDatRst.iFdFri Then
                    ilFdDay = 4
                ElseIf tlDatRst.iFdSat Then
                    ilFdDay = 5
                ElseIf tlDatRst.iFdSun Then
                    ilFdDay = 6
                End If
                If tlDatRst.iPdMon <> 0 Then
                    ilPdDay = 0
                ElseIf tlDatRst.iPdTue Then
                    ilPdDay = 1
                ElseIf tlDatRst.iPdWed Then
                    ilPdDay = 2
                ElseIf tlDatRst.iPdThu Then
                    ilPdDay = 3
                ElseIf tlDatRst.iPdFri Then
                    ilPdDay = 4
                ElseIf tlDatRst.iPdSat Then
                    ilPdDay = 5
                ElseIf tlDatRst.iPdSun Then
                    ilPdDay = 6
                End If
                If (ilPdDay >= ilFdDay) Then
                    ilAdjDay = ilPdDay - ilFdDay
                    If tlDatRst.sPdDayFed = "B" Then
                        ilAdjDay = ilAdjDay - 7
                    End If
                Else
                    If tlDatRst.sPdDayFed = "B" Then
                        ilAdjDay = ilPdDay - ilFdDay
                        If ilAdjDay > 0 Then
                            ilAdjDay = ilAdjDay - 1
                        End If
                    Else
                        ilAdjDay = 7 + ilPdDay - ilFdDay
                    End If
                End If
                slPdDate = Format$(DateValue(gAdjYear(Format(slFdDate, sgShowDateForm))) + ilAdjDay, sgShowDateForm)
                slPdTime = tlDatRst.sPdStTime
            End If
        End If
    End If
    If gObtainPrevMonday(slPdDate) <> gObtainPrevMonday(slAirDate) Then
        gDeterminePledgeDateTime = True
    End If
End Function

Private Sub mGetTrafficPostedTimes(llVpf As Long, tlLST As LST, ilPledged As Integer, ilDACode As Integer, ilCPStatus As Integer, ilLocalAdj As Integer, ilAdjDay As Integer, slPdSTime As String, slFdSTime As String, slAirDate As String, slAirTime As String)
    Dim slPostedAirTime As String
    Dim llPostedAirDate As Long
    Dim llPostedAirTime As Long
    Dim slSQLQuery As String
    
    If (llVpf <> -1) And (tlLST.iType = 0) And (ilPledged <> 2) And (ilDACode <> 2) And (ilCPStatus = 0) Then
        If (Asc(tgVpfOptions(llVpf).sUsingFeatures1) And EXPORTPOSTEDTIMES) = EXPORTPOSTEDTIMES Then
            slSQLQuery = "Select sdfTime from SDF_Spot_Detail"
            slSQLQuery = slSQLQuery & " Where (sdfCode = " & tlLST.lSdfCode & ")"
            Set sdf_rst = gSQLSelectCall(slSQLQuery)
            If Not sdf_rst.EOF Then
                slPostedAirTime = Format$(sdf_rst!sdfTime, "h:mm:ssam/pm")
                'If iIndex >= 0 Then
                    slPostedAirTime = Format$(gLongToTime(gTimeToLong(slPostedAirTime, False) + gTimeToLong(slPdSTime, False) - gTimeToLong(slFdSTime, False)), sgShowTimeWSecForm)
                'End If
                llPostedAirDate = DateValue(gAdjYear(tlLST.sLogDate))
                llPostedAirTime = gTimeToLong(slPostedAirTime, False) + 3600 * ilLocalAdj
                If llPostedAirTime < 0 Then
                    llPostedAirTime = llPostedAirTime + 86400
                    llPostedAirDate = llPostedAirDate - 1
                ElseIf llPostedAirTime > 86400 Then
                    llPostedAirTime = llPostedAirTime - 86400
                    llPostedAirDate = llPostedAirDate + 1
                End If
                'Air time and Date only set if ast is new or astCPStatus = 0 (not posted)
                slAirDate = Format$(llPostedAirDate + ilAdjDay, "m/d/yyyy")
                slAirTime = Format$(gLongToTime(llPostedAirTime), sgShowTimeWSecForm)
            End If
        End If
    End If

End Sub
Public Function gGetMissedPledgeForMG(ilStatus As Integer, slFeedDate As String, llLkAstCode As Long, slMissedFeedDate As String, slMissedFeedTime As String) As Boolean
    Dim slSQLQuery As String
    
    gGetMissedPledgeForMG = False
    If (gIsAstStatus(ilStatus, ASTEXTENDED_MG)) Or (gIsAstStatus(ilStatus, ASTEXTENDED_REPLACEMENT)) Then
        slSQLQuery = "Select astFeedDate, astFeedTime FROM ast WHERE (AstCode = " & llLkAstCode & ")"
        Set ast_rst = gSQLSelectCall(slSQLQuery)
        If Not ast_rst.EOF Then
            gGetMissedPledgeForMG = True
            slMissedFeedDate = Format$(ast_rst!astFeedDate, sgShowDateForm)
            If Second(ast_rst!astFeedTime) <> 0 Then
                slMissedFeedTime = Format$(ast_rst!astFeedTime, sgShowTimeWSecForm)
            Else
                slMissedFeedTime = Format$(ast_rst!astFeedTime, sgShowTimeWOSecForm)
            End If
        End If
    End If
End Function

Public Sub gFilterAstExtendedTypes(tlAstInfo() As ASTINFO)
    Dim llCount As Long
    Dim llAst As Long
    
    ReDim tlTempAstInfo(0 To UBound(tlAstInfo)) As ASTINFO
    llCount = 0
    For llAst = 0 To UBound(tlAstInfo) - 1 Step 1
        If (tlAstInfo(llAst).iStatus Mod 100 <= 10) Or (tlAstInfo(llAst).iStatus Mod 100 = ASTAIR_MISSED_MG_BYPASS) Then
            tlTempAstInfo(llCount) = tlAstInfo(llAst)
            llCount = llCount + 1
        End If
    Next llAst
    If llCount < UBound(tlAstInfo) Then
        ReDim tlAstInfo(0 To llCount) As ASTINFO
        For llAst = 0 To llCount Step 1
            tlAstInfo(llAst) = tlTempAstInfo(llAst)
        Next llAst
    End If
    Erase tlTempAstInfo
End Sub

Public Function gSplitFillDefined(ilVefCode As Integer, slStartDate As String, slEndDate As String) As Boolean
    Dim llStartDate As Long
    Dim llEndDate As Long
    Dim blDayOk As Boolean
    Dim llDate As Long
    Dim tlBof As BOF
    Dim bof_rst As ADODB.Recordset
    Dim lst_rst As ADODB.Recordset
    Dim slSQLQuery As String
    
    On Error GoTo ErrHand

    If ((Asc(sgSpfUsingFeatures2) And SPLITNETWORKS) <> SPLITNETWORKS) Then
        gSplitFillDefined = True
        Exit Function
    End If
    slSQLQuery = "SELECT Count(*) as SplitCount"
    slSQLQuery = slSQLQuery + " FROM lst"
    slSQLQuery = slSQLQuery + " WHERE "
    slSQLQuery = slSQLQuery + "lstLogVefCode = " & ilVefCode
    slSQLQuery = slSQLQuery + " And lstLogDate >= '" & Format(slStartDate, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery + " And lstLogDate <= '" & Format(slEndDate, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery + " And lstSplitNetwork In('P', 'S')"
    'Set lst_rst = cnn.Execute(slSQLQuery)
    Set lst_rst = gSQLSelectCall(slSQLQuery)
    If lst_rst.EOF Then
        gSplitFillDefined = True
        Exit Function
    End If
    If lst_rst!SplitCount <= 0 Then
        gSplitFillDefined = True
        Exit Function
    End If
    llStartDate = DateValue(gAdjYear(slStartDate))
    llEndDate = DateValue(gAdjYear(slEndDate))
    slSQLQuery = "SELECT *"
    slSQLQuery = slSQLQuery + " FROM BOF_Blackout"
    slSQLQuery = slSQLQuery + " WHERE (bofType = 'R'" & ")"
    'Set bof_rst = cnn.Execute(slSQLQuery)
    Set bof_rst = gSQLSelectCall(slSQLQuery)
    While Not bof_rst.EOF
        gCreateUDTforBOF bof_rst, tlBof
        If tlBof.lCifCode > 0 Then
            If (tlBof.iVefCode = ilVefCode) Or (tlBof.iVefCode = 0) Then
                If (DateValue(gAdjYear(tlBof.sEndDate)) >= llStartDate) And (DateValue(gAdjYear(tlBof.sStartDate)) <= llEndDate) Then
                    If llStartDate + 6 >= llEndDate Then
                        gSplitFillDefined = True
                        Exit Function
                    End If
                    For llDate = llStartDate To llEndDate Step 1
                        blDayOk = False
                        Select Case gWeekDayLong(llDate)
                            Case 0  'Monday
                                If tlBof.sMo = "Y" Then
                                    blDayOk = True
                                End If
                            Case 1
                                If tlBof.sTu = "Y" Then
                                    blDayOk = True
                                End If
                            Case 2
                                If tlBof.sWe = "Y" Then
                                    blDayOk = True
                                End If
                            Case 3
                                If tlBof.sTh = "Y" Then
                                    blDayOk = True
                                End If
                            Case 4
                                If tlBof.sFr = "Y" Then
                                    blDayOk = True
                                End If
                            Case 5
                                If tlBof.sSa = "Y" Then
                                    blDayOk = True
                                End If
                            Case 6
                                If tlBof.sSu = "Y" Then
                                    blDayOk = True
                                End If
                        End Select
                        If blDayOk Then
                            gSplitFillDefined = True
                            Exit Function
                        End If
                    Next llDate
                End If
            End If
        End If
        bof_rst.MoveNext
    Wend
    gSplitFillDefined = False
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "modCPReturns-gSplitFillDefined"
End Function

Public Function gProgramDefined(ilInVefCode As Integer, slStartDate As String, slEndDate As String)
    Dim ilVef As Integer
    Dim slSQLQuery As String
    ReDim ilVefCode(0 To 0) As Integer
    
    On Error GoTo ErrHand
    gProgramDefined = False
    ilVef = gBinarySearchVef(CLng(ilInVefCode))
    If ilVef = -1 Then
        Exit Function
    End If
    If tgVehicleInfo(ilVef).sVehType = "A" Then
        gGetSellingVehicles ilInVefCode, slStartDate, ilVefCode()
    ElseIf tgVehicleInfo(ilVef).sVehType = "L" Then
        gGetLogVehicles ilInVefCode, ilVefCode()
    Else
        ilVefCode(0) = ilInVefCode
        ReDim Preserve ilVefCode(0 To 1) As Integer
    End If
    For ilVef = LBound(ilVefCode) To UBound(ilVefCode) - 1 Step 1
        slSQLQuery = "SELECT * "
        slSQLQuery = slSQLQuery + " FROM LCF_Log_Calendar"
        slSQLQuery = slSQLQuery + " WHERE ("
        slSQLQuery = slSQLQuery & " lcfStatus = 'C'"
        slSQLQuery = slSQLQuery + " AND lcfLogDate >= '" & Format$(slStartDate, sgSQLDateForm) & "' And lcfLogDate <= '" & Format$(slEndDate, sgSQLDateForm) & "'"
        slSQLQuery = slSQLQuery & " AND lcfVefCode = " & ilVefCode(ilVef) & ")"
        'Set rst_lcf = cnn.Execute(slSQLQuery)
        Set rst_lcf = gSQLSelectCall(slSQLQuery)
        If Not rst_lcf.EOF Then
            gProgramDefined = True
            Exit Function
        End If
    Next ilVef
    Exit Function
    
ErrHand:
    gHandleError "AffErrorLog.txt", "modCPReturns-gProgramDefined"
End Function

Public Sub gGetSellingVehicles(ilAirVefCode As Integer, slDate As String, ilSellVefCode() As Integer)
    Dim slSQLQuery As String
    Dim vlf_rst As ADODB.Recordset
    On Error GoTo ErrHand
    slSQLQuery = "Select distinct vlfSellCode from VLF_Vehicle_Linkages where vlfAirCode = " & ilAirVefCode
    slSQLQuery = slSQLQuery & " And vlfEffDate <= '" & Format(slDate, sgSQLDateForm) & "'  and (vlfTermDate >= '" & Format(slDate, sgSQLDateForm) & " ' or vlfTermDate is null)"
    'Set vlf_rst = cnn.Execute(slSQLQuery)
    Set vlf_rst = gSQLSelectCall(slSQLQuery)
    Do While Not vlf_rst.EOF
        ilSellVefCode(UBound(ilSellVefCode)) = vlf_rst!vlfSellCode
        ReDim Preserve ilSellVefCode(0 To UBound(ilSellVefCode) + 1) As Integer
        vlf_rst.MoveNext
    Loop
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "modCPReturns-gGetSellingVehicles"
End Sub

Public Sub gGetLogVehicles(ilLogVefCode As Integer, ilVefCode() As Integer)
    Dim slSQLQuery As String
    Dim vef_rst As ADODB.Recordset
    On Error GoTo ErrHand
    slSQLQuery = "Select vefCode From vef_Vehicles Where vefVefCode = " & ilLogVefCode
    'Set vef_rst = cnn.Execute(slSQLQuery)
    Set vef_rst = gSQLSelectCall(slSQLQuery)
    Do While Not vef_rst.EOF
        If mMergeWithLog(vef_rst!vefCode) Then
            ilVefCode(UBound(ilVefCode)) = vef_rst!vefCode
            ReDim Preserve ilVefCode(0 To UBound(ilVefCode) + 1) As Integer
        End If
        vef_rst.MoveNext
    Loop
    Exit Sub
ErrHand:
    gHandleError "AffErrorLog.txt", "modCPReturns-gGetSellingVehicles"

End Sub
Private Function mMergeWithLog(ilVefCode As Integer) As Integer
    Dim ilVff As Integer
    
    mMergeWithLog = True
    For ilVff = LBound(tgVffInfo) To UBound(tgVffInfo) Step 1
        If ilVefCode = tgVffInfo(ilVff).iVefCode Then
            If tgVffInfo(ilVff).sMergeTraffic = "S" Then
                mMergeWithLog = False
            End If
            Exit For
        End If
    Next ilVff

End Function
Public Sub gClearAbf(ilVefCode As Integer, ilShttCode As Integer, slFromDate As String, slToDate As String, blCheckServiceAgreement As Boolean)
    Dim slMoDate As String
    Dim slSQLQuery As String
    Dim llRet As Long
    Dim att_rst As ADODB.Recordset
    
    If blCheckServiceAgreement Then
        slSQLQuery = "SELECT Count(*) as ServiceCount"
        slSQLQuery = slSQLQuery + " FROM att"
        slSQLQuery = slSQLQuery + " WHERE attVefCode = " & Trim$(Str(ilVefCode))
        slSQLQuery = slSQLQuery & " And attServiceAgreement = 'Y'"
        slSQLQuery = slSQLQuery & " AND attOffAir >= attOnAir"
        slSQLQuery = slSQLQuery & " AND attOnAir <= " & "'" & Format$(slToDate, sgSQLDateForm) & "'"
        slSQLQuery = slSQLQuery & " And attOffAir >= '" & Format(slFromDate, sgSQLDateForm) & "' "
        slSQLQuery = slSQLQuery & " And attDropDate >= '" & Format(slFromDate, sgSQLDateForm) & "' "
        'Set att_rst = cnn.Execute(slSQLQuery)
        Set att_rst = gSQLSelectCall(slSQLQuery)
        If att_rst!ServiceCount > 0 Then
            Exit Sub
        End If
    End If
    
    slMoDate = gObtainPrevMonday(slFromDate)
    
    Do While gDateValue(slMoDate) <= gDateValue(slToDate)
        slSQLQuery = "SELECT * FROM abf_AST_Build_Queue "
        slSQLQuery = slSQLQuery & " WHERE abfStatus In ('G', 'P', 'H') "
        slSQLQuery = slSQLQuery & " And abfVefCode = " & ilVefCode
        slSQLQuery = slSQLQuery & " And abfShttCode = " & ilShttCode
        slSQLQuery = slSQLQuery & " And abfMondayDate = '" & Format(slMoDate, sgSQLDateForm) & "' "
        Set rst_abf = gSQLSelectCall(slSQLQuery)
        Do While Not rst_abf.EOF
            If gDateValue(Format(rst_abf!abfGenStartDate, sgShowDateForm)) >= gDateValue(slFromDate) And gDateValue(Format(rst_abf!abfGenStartDate, sgShowDateForm)) <= gDateValue(slToDate) Then
                If gDateValue(Format(rst_abf!abfGenEndDate, sgShowDateForm)) >= gDateValue(slFromDate) And gDateValue(Format(rst_abf!abfGenEndDate, sgShowDateForm)) <= gDateValue(slToDate) Then
                    slSQLQuery = "UPDATE ABF_AST_Build_Queue SET "
                    slSQLQuery = slSQLQuery & "abfStatus = 'C'" & ", "  'Completed
                    slSQLQuery = slSQLQuery & "abfCompletedDate = '" & Format$(gNow(), sgSQLDateForm) & "', "
                    slSQLQuery = slSQLQuery & "abfCompletedTime = '" & Format$(gNow(), sgSQLTimeForm) & "', "
                    slSQLQuery = slSQLQuery & "abfUstCode = " & igUstCode & " "
                    slSQLQuery = slSQLQuery & "WHERE abfCode = " & rst_abf!abfCode
                    llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
                End If
            End If
            rst_abf.MoveNext
        Loop
        slMoDate = DateAdd("d", 7, slMoDate)
    Loop
End Sub

Private Function mIncludeLst_ExclusionCheck(rst_Lst As ADODB.Recordset, ilLocalAdj As Integer) As Boolean   'ilATTCrossDates As Integer, ilSpotType As Integer, llCntrNo As Long) As Boolean
    Dim slSQLQuery As String
    Dim ilSpotType As Integer
    Dim llCntrNo As Long
    Dim ilLogVefCode As Integer
    Dim slLogDate As String
    Dim slLogTime As String
    Dim llLogDate As Long
    Dim llLogTime As Long
    Dim ilAtt As Integer
    
    mIncludeLst_ExclusionCheck = True
    If (bgAnyAttExclusions <> True) Then Exit Function
    If rst_Lst!lstType = 1 Then Exit Function   'Avail
    
    ilSpotType = rst_Lst!lstSpotType
    llCntrNo = rst_Lst!lstCntrNo
    ilLogVefCode = rst_Lst!lstLogVefCode
    slLogDate = Format$(rst_Lst!lstLogDate, sgShowDateForm)
    slLogTime = Format$(rst_Lst!lstLogTime, sgShowTimeWSecForm)
    llLogDate = DateValue(gAdjYear(slLogDate))
    llLogTime = gTimeToLong(slLogTime, False)
    llLogTime = llLogTime + 3600 * ilLocalAdj
    If llLogTime < 0 Then
        llLogTime = llLogTime + 86400
        llLogDate = llLogDate - 1
    ElseIf llLogTime > 86400 Then
        llLogTime = llLogTime - 86400
        llLogDate = llLogDate + 1
    End If
    
    'Determine which agreement
    For ilAtt = 0 To UBound(tmATTCrossDates) - 1 Step 1
        If (llLogDate >= tmATTCrossDates(ilAtt).lStartDate) And (llLogDate <= tmATTCrossDates(ilAtt).lEndDate) Then
    
            If (ilSpotType = 2) And (tmATTCrossDates(ilAtt).sExcludeFillSpot = "Y") Then
                mIncludeLst_ExclusionCheck = False
                Exit Function
            End If
            
            slSQLQuery = " Select chfType from CHF_Contract_Header where (chfSchStatus = 'F' or chfschstatus = 'M') AND chfDelete = 'N' AND  chfCntrNo = " & llCntrNo
            Set chf_rst = gSQLSelectCall(slSQLQuery)
            If Not chf_rst.EOF Then
                If (chf_rst!chfType = "Q") And (tmATTCrossDates(ilAtt).sExcludeCntrTypeQ = "Y") Then mIncludeLst_ExclusionCheck = False
                If (chf_rst!chfType = "R") And (tmATTCrossDates(ilAtt).sExcludeCntrTypeR = "Y") Then mIncludeLst_ExclusionCheck = False
                If (chf_rst!chfType = "T") And (tmATTCrossDates(ilAtt).sExcludeCntrTypeT = "Y") Then mIncludeLst_ExclusionCheck = False
                If (chf_rst!chfType = "M") And (tmATTCrossDates(ilAtt).sExcludeCntrTypeM = "Y") Then mIncludeLst_ExclusionCheck = False
                If (chf_rst!chfType = "S") And (tmATTCrossDates(ilAtt).sExcludeCntrTypeS = "Y") Then mIncludeLst_ExclusionCheck = False
                If (chf_rst!chfType = "V") And (tmATTCrossDates(ilAtt).sExcludeCntrTypeV = "Y") Then mIncludeLst_ExclusionCheck = False
            End If
            Exit For
        End If
    Next ilAtt
End Function
Private Sub mSetAnyAttExclusions(tlATTCrossDates As ATTCrossDates)
    If (tlATTCrossDates.sExcludeFillSpot = "Y") Then bgAnyAttExclusions = True
    If (tlATTCrossDates.sExcludeCntrTypeQ = "Y") Then bgAnyAttExclusions = True
    If (tlATTCrossDates.sExcludeCntrTypeR = "Y") Then bgAnyAttExclusions = True
    If (tlATTCrossDates.sExcludeCntrTypeT = "Y") Then bgAnyAttExclusions = True
    If (tlATTCrossDates.sExcludeCntrTypeM = "Y") Then bgAnyAttExclusions = True
    If (tlATTCrossDates.sExcludeCntrTypeS = "Y") Then bgAnyAttExclusions = True
    If (tlATTCrossDates.sExcludeCntrTypeV = "Y") Then bgAnyAttExclusions = True
End Sub

Private Sub mFilterByAdvt(tlAstInfo() As ASTINFO, iAdfCode As Integer)
    Dim ilAst As Integer
    Dim ilUpper As Integer
    
    If iAdfCode > 0 Then
        ReDim tlTmp(0 To UBound(tlAstInfo)) As ASTINFO
        ilUpper = 0
        For ilAst = 0 To UBound(tlAstInfo) Step 1
            If tlAstInfo(ilAst).iAdfCode = iAdfCode Then
                LSet tlTmp(ilUpper) = tlAstInfo(ilAst)
                ilUpper = ilUpper + 1
            End If
        Next ilAst
        ReDim tlAstInfo(0 To ilUpper) As ASTINFO
        For ilAst = 0 To ilUpper Step 1
            LSet tlAstInfo(ilAst) = tlTmp(ilAst)
        Next ilAst
    End If

End Sub

Private Sub mGetAstDateRange(tlCPPosting As CPPOSTING, blFeedAdjOnReturn As Boolean, slFWkDate As String, slLWkDate As String)

    If (igTimes = 3) Or (igTimes = 4) Then
        If blFeedAdjOnReturn Then
            slFWkDate = Format$(DateAdd("d", -1, tlCPPosting.sDate), sgShowDateForm)
        Else
            slFWkDate = Format$(tlCPPosting.sDate, sgShowDateForm)
        End If
    Else
        slFWkDate = Format$(gObtainPrevMonday(tlCPPosting.sDate), sgShowDateForm)
    End If
    If igTimes = 0 Then
        slLWkDate = Format$(gObtainEndStd(tlCPPosting.sDate), sgShowDateForm)
    ElseIf (igTimes = 3) Or (igTimes = 4) Then
        If blFeedAdjOnReturn Then
            slLWkDate = Format$(DateAdd("d", tlCPPosting.iNumberDays, tlCPPosting.sDate), sgShowDateForm)
        Else
            slLWkDate = Format$(DateAdd("d", tlCPPosting.iNumberDays - 1, tlCPPosting.sDate), sgShowDateForm)
        End If
    Else
        slLWkDate = Format$(gObtainNextSunday(tlCPPosting.sDate), sgShowDateForm)
    End If

End Sub

Private Function mAdvtBkout(ilAdfCode As Integer, ilVefCode As Integer, llFWkDate As Long, llLWkDate As Long, blFeedAdjOnReturn As Boolean) As Boolean
    Dim slSQLQuery As String
    Dim slKey As String
    
    On Error GoTo AdvtCountErr:
    mAdvtBkout = False
    If ilAdfCode <= 0 Then
        Exit Function
    End If
    slKey = ilAdfCode & ilVefCode & llFWkDate & llLWkDate & blFeedAdjOnReturn
    If slKey = sgAdvtBkoutKey Then
        mAdvtBkout = bgAdvtBlout
        Exit Function
    End If
    sgAdvtBkoutKey = slKey
    'slSQLQuery = "Select Count(1) as BkoutCount From lst As A Left Outer Join Lst As B on A.lstBkoutLstCode = B.LstCode"
    'slSQLQuery = slSQLQuery & " Where B.lstAdfCode = " & ilAdfCode
    'slSQLQuery = slSQLQuery & " AND A.lstLogVefCode = " & ilVefCode
    'slSQLQuery = slSQLQuery & " AND A.lstBkoutLstCode > 0"
    'slSQLQuery = slSQLQuery & " AND A.lstType <> 1"
    'If igTimes = 0 Then
    '    slSQLQuery = slSQLQuery + " AND (A.lstLogDate >= '" & Format$(llFWkDate - 1, sgSQLDateForm) & "' AND A.lstLogDate <= '" & Format$(llLWkDate + 1, sgSQLDateForm) & "')" & ")"
    'Else
    '    If ((igTimes = 3) Or (igTimes = 4)) And (blFeedAdjOnReturn) Then
    '        slSQLQuery = slSQLQuery + " AND (A.lstLogDate >= '" & Format$(llFWkDate, sgSQLDateForm) & "' AND A.lstLogDate <= '" & Format$(llLWkDate, sgSQLDateForm) & "')" & ")"
    '    Else
    '        slSQLQuery = slSQLQuery + " AND (A.lstLogDate >= '" & Format$(llFWkDate - 1, sgSQLDateForm) & "' AND A.lstLogDate <= '" & Format$(llLWkDate + 1, sgSQLDateForm) & "')"
    '    End If
    'End If
    
    slSQLQuery = "Select Count(1) as BkoutCount From lst "
    slSQLQuery = slSQLQuery & " Where lstAdfCode = " & ilAdfCode
    slSQLQuery = slSQLQuery & " AND lstLogVefCode = " & ilVefCode
    slSQLQuery = slSQLQuery & " AND lstType <> 1"
    If igTimes = 0 Then
        slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(llFWkDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(llLWkDate + 1, sgSQLDateForm) & "')"
    Else
        If ((igTimes = 3) Or (igTimes = 4)) And (blFeedAdjOnReturn) Then
            slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(llFWkDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(llLWkDate, sgSQLDateForm) & "')"
        Else
            slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(llFWkDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(llLWkDate + 1, sgSQLDateForm) & "')"
        End If
    End If
    slSQLQuery = slSQLQuery & " And lstCode In ("
    slSQLQuery = slSQLQuery & "Select lstBkoutLstCode From lst "
    slSQLQuery = slSQLQuery & " Where lstLogVefCode = " & ilVefCode
    slSQLQuery = slSQLQuery & " AND lstBkoutLstCode > 0"
    slSQLQuery = slSQLQuery & " AND lstType <> 1"
    If igTimes = 0 Then
        slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(llFWkDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(llLWkDate + 1, sgSQLDateForm) & "')"
    Else
        If ((igTimes = 3) Or (igTimes = 4)) And (blFeedAdjOnReturn) Then
            slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(llFWkDate, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(llLWkDate, sgSQLDateForm) & "')"
        Else
            slSQLQuery = slSQLQuery + " AND (lstLogDate >= '" & Format$(llFWkDate - 1, sgSQLDateForm) & "' AND lstLogDate <= '" & Format$(llLWkDate + 1, sgSQLDateForm) & "')"
        End If
    End If
    slSQLQuery = slSQLQuery & ")"
    Set lst_rst = gSQLSelectCall(slSQLQuery)
    If lst_rst.EOF Then
        bgAdvtBlout = False
        Exit Function
    End If
    If lst_rst!BkoutCount <= 0 Then
        bgAdvtBlout = False
        Exit Function
    End If
    bgAdvtBlout = True
    mAdvtBkout = True
    Exit Function
AdvtCountErr:
    Resume Next
    Exit Function
End Function

Public Sub gGetPoolAdf()
    Dim slSQLQuery As String
    If Not gFileChgd("adf.btr") Then
        Exit Sub
    End If
    ReDim igPoolAdfCode(0 To 0) As Integer
    slSQLQuery = "Select adfCode from Adf_Advertisers"
    slSQLQuery = slSQLQuery & " Where "
    slSQLQuery = slSQLQuery & " (adfBkoutPoolStatus = 'A'"
    slSQLQuery = slSQLQuery & " Or adfBkoutPoolStatus = 'U'" & ")"
    Set rst_adf = gSQLSelectCall(slSQLQuery)
    Do While Not rst_adf.EOF
        igPoolAdfCode(UBound(igPoolAdfCode)) = rst_adf!adfCode
        ReDim Preserve igPoolAdfCode(0 To UBound(igPoolAdfCode) + 1) As Integer
        rst_adf.MoveNext
    Loop
    ReDim Preserve igSortPoolAdfCode(0 To UBound(igPoolAdfCode)) As Integer
    rst_adf.Close
End Sub


Private Sub mBuildPoolRegionDefinitions(llSdfCode As Long, ilVefCode As Integer, llUpper As Long, tlRegionDefinition() As REGIONDEFINITION)
    'ilVefCode = agreement vehicle
    Dim slSQLQuery As String
    Dim ilAdf As Integer
    Dim blUpdateFirst As Boolean
    Dim blVefOk As Boolean
    Dim blAdfOk As Boolean
    Dim slSdfDate As String
    Dim slSdfTime As String
    Dim slLen As String
    Dim crf_sub As ADODB.Recordset
    Dim crf_pool As ADODB.Recordset
    Dim ilDay As Integer
    Dim slDay As String
    Dim llSdfTime As Long
    Dim llEndTime As Long
    Dim ilIndex As Integer
    Dim ilLoop As Integer
    Dim ilPriorAdjAdfCode As Integer
    Dim ilNextAdjAdfCode As Integer
    Dim ilAdfCode As Integer
    Dim blPoolFound As Boolean
    Dim llVef As Long
    Dim llAdf As Long
    Dim ilRet As Integer

    On Error GoTo Err:
    If UBound(igPoolAdfCode) > LBound(igPoolAdfCode) Then
        blPoolFound = True
        mGetAdjAdvt llSdfCode, ilVefCode, ilPriorAdjAdfCode, ilNextAdjAdfCode
        blUpdateFirst = True
        slSQLQuery = "Select crfBkoutInstAdfCode, crfCode from CRF_Copy_Rot_Header"
        slSQLQuery = slSQLQuery & " Where (crfCode = " & rsf_rst!rsfCrfCode
        slSQLQuery = slSQLQuery & ")"
        slSQLQuery = slSQLQuery & " Order by crfRotNo Desc"
        Set crf_pool = gSQLSelectCall(slSQLQuery)
        Do While Not crf_pool.EOF
            If crf_pool!crfBkoutInstAdfCode = POOLROTATION Then
                'Find replacement rotations
                slSQLQuery = "Select sdfDate, sdfTime, sdfSpotType, sdfLen, sdfAdfCode from SDF_Spot_Detail"
                slSQLQuery = slSQLQuery & " Where sdfCode = " & llSdfCode
                Set sdf_rst = gSQLSelectCall(slSQLQuery)
                If Not sdf_rst.EOF Then
                    blPoolFound = False
                    slSdfDate = Format(sdf_rst!sdfDate, sgSQLDateForm)
                    slSdfTime = Format(sdf_rst!sdfTime, sgSQLTimeForm)
                    llSdfTime = gTimeToLong(Format(sdf_rst!sdfTime, "h:mm:ssAM/PM"), False)
                    slLen = sdf_rst!sdfLen
                    ilDay = gWeekDayLong(gDateValue(sdf_rst!sdfDate))
                    ilAdfCode = sdf_rst!sdfAdfCode
                    If igLastPoolAdfCode > 0 Then
                        For ilAdf = 0 To UBound(igPoolAdfCode) - 1 Step 1
                            If igLastPoolAdfCode = igPoolAdfCode(ilAdf) Then
                                ilIndex = 0
                                For ilLoop = ilAdf + 1 To UBound(igPoolAdfCode) - 1 Step 1
                                    igSortPoolAdfCode(ilIndex) = igPoolAdfCode(ilLoop)
                                    ilIndex = ilIndex + 1
                                Next ilLoop
                                For ilLoop = 0 To ilAdf Step 1
                                    igSortPoolAdfCode(ilIndex) = igPoolAdfCode(ilLoop)
                                    ilIndex = ilIndex + 1
                                Next ilLoop
                                Exit For
                            End If
                        Next ilAdf
                    Else
                        For ilAdf = 0 To UBound(igPoolAdfCode) - 1 Step 1
                            igSortPoolAdfCode(ilAdf) = igPoolAdfCode(ilAdf)
                        Next ilAdf
                    End If
                    For ilAdf = 0 To UBound(igSortPoolAdfCode) - 1 Step 1
                        If ilAdfCode = igSortPoolAdfCode(ilAdf) Then
                            blAdfOk = False
                        ElseIf ilAdf = UBound(igSortPoolAdfCode) - 1 Then
                            blAdfOk = True
                        ElseIf (igLastPoolAdfCode <> igSortPoolAdfCode(ilAdf)) And (ilPriorAdjAdfCode <> igSortPoolAdfCode(ilAdf)) And (ilNextAdjAdfCode <> igSortPoolAdfCode(ilAdf)) Then
                            blAdfOk = True
                        Else
                            blAdfOk = False
                        End If
                        If blAdfOk Then
                            slSQLQuery = "Select * from CRF_Copy_Rot_Header"
                            slSQLQuery = slSQLQuery & " Where (crfAdfCode = " & igSortPoolAdfCode(ilAdf)
                            If sdf_rst!sdfSpotType = "O" Then
                                slSQLQuery = slSQLQuery & " And crfRotType = 'O'"
                            ElseIf sdf_rst!sdfSpotType = "C" Then
                                slSQLQuery = slSQLQuery & " And crfRotType = 'C'"
                            Else
                                slSQLQuery = slSQLQuery & " And crfRotType = 'A'"
                            End If
                            slSQLQuery = slSQLQuery & " And crfStartDate <= '" & slSdfDate & "'"
                            slSQLQuery = slSQLQuery & " And crfEndDate >= '" & slSdfDate & "'"
                            slSQLQuery = slSQLQuery & " And crfStartTime <= '" & slSdfTime & "'"
                            'slSQLQuery = slSQLQuery & " And crfEndTime >= '" & slSdfTime & "'"
                            slSQLQuery = slSQLQuery & " And crfAnfCode = " & 0
                            slSQLQuery = slSQLQuery & " And crfRafCode = " & 0
                            slSQLQuery = slSQLQuery & " And crfBkoutInstAdfCode = " & 0
                            slSQLQuery = slSQLQuery & " And crfLen = " & slLen
                            slSQLQuery = slSQLQuery & ")"
                            slSQLQuery = slSQLQuery & " Order by crfRotNo Desc"
                            Set crf_sub = gSQLSelectCall(slSQLQuery)
                            Do While Not crf_sub.EOF
                                slDay = Switch(ilDay = 0, crf_sub!crfMo, ilDay = 1, crf_sub!crfTu, ilDay = 2, crf_sub!crfWe, _
                                               ilDay = 3, crf_sub!crfTh, ilDay = 4, crf_sub!crfFr, ilDay = 5, crf_sub!crfSa, ilDay = 6, crf_sub!crfSu)
                                If slDay = "Y" Then
                                    llEndTime = gTimeToLong(Format(crf_sub!crfEndTime, "h:mm:ssAM/PM"), True)
                                    If llSdfTime <= llEndTime Then
                                        'Test if rotation vehicle
                                        'blVefOk = False
                                        'If ilVefCode = crf_sub!CrfVefCode Then
                                        '    blVefOk = True
                                        'ElseIf crf_sub!CrfVefCode <= 0 Then
                                        '    blVefOk = mTestCvf(ilVefCode, crf_sub!crfCode)
                                        'Else
                                        '    blVefOk = mTestPvf(ilVefCode, crf_sub!CrfVefCode)
                                        'End
                                        blVefOk = mVefOk(crf_sub!crfChfCode, ilVefCode)
                                        If blVefOk Then
                                            blPoolFound = True
                                            If blUpdateFirst Then
                                                tlRegionDefinition(llUpper - 1).lCrfCode = crf_sub!crfCode
                                                mGetPoolCopy igSortPoolAdfCode(ilAdf), crf_sub!crfNextFinal, crf_pool!crfCode, tlRegionDefinition(llUpper - 1)
                                                blUpdateFirst = False
                                                mUpdatePoolInfo igSortPoolAdfCode(ilAdf), crf_sub!crfCode, crf_sub!crfNextFinal, tlRegionDefinition(llUpper - 1)
                                                'Only need one valid advertiser
                                                Exit Sub
                                            Else
                                                tlRegionDefinition(llUpper) = tlRegionDefinition(llUpper - 1)
                                                tlRegionDefinition(llUpper).lCrfCode = crf_sub!crfCode
                                                mGetPoolCopy igSortPoolAdfCode(ilAdf), crf_sub!crfNextFinal, crf_pool!crfCode, tlRegionDefinition(llUpper)
                                                llUpper = llUpper + 1
                                                If llUpper >= UBound(tlRegionDefinition) Then
                                                    ReDim Preserve tlRegionDefinition(0 To UBound(tlRegionDefinition) + 10) As REGIONDEFINITION
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                crf_sub.MoveNext
                            Loop
                        End If
                    Next ilAdf
                End If
            End If
            crf_pool.MoveNext
        Loop
        If Not blPoolFound Then
            llVef = gBinarySearchVef(CLng(ilVefCode))
            If llVef <> -1 Then
                llAdf = gBinarySearchAdf(CLng(ilAdfCode))
                If llAdf <> -1 Then
                    gLogMsg "Pool Not Assigned to " & Trim(tgVehicleInfo(llVef).sVehicle) & " " & Trim$(tgAdvtInfo(llAdf).sAdvtName) & " " & slSdfDate & " " & slSdfTime & " as no Copy Rotation found", "PoolUnassignedLog_" & Format(Now, "mm-dd-yy") & ".txt", False
                Else
                    gLogMsg "Pool Not Assigned to " & Trim(tgVehicleInfo(llVef).sVehicle) & " " & slSdfDate & " " & slSdfTime & " as no Copy Rotation found", "PoolUnassignedLog_" & Format(Now, "mm-dd-yy") & ".txt", False
                End If
                ilRet = gAlertAdd("U", "P", 0, Format(Now, "ddddd"))
            End If
        End If
    End If
    Exit Sub
Err:
    Resume Next
End Sub
Private Function mVefOk(llChfCode As Long, ilVefCode As Integer) As Boolean
    Dim slSQLQuery As String
    
    mVefOk = False
    slSQLQuery = "SELECT Count(1) as ClfCount"
    slSQLQuery = slSQLQuery + " FROM clf_Contract_Line"
    slSQLQuery = slSQLQuery + " WHERE (clfchfCode = " & llChfCode
    slSQLQuery = slSQLQuery & " And clfVefCode = " & ilVefCode
    slSQLQuery = slSQLQuery & ")"
    Set chf_rst = gSQLSelectCall(slSQLQuery)
    If Not chf_rst.EOF Then
        If chf_rst!ClfCount > 0 Then
            mVefOk = True
        End If
    End If
End Function
Private Function mTestCvf(ilVefCode As Integer, llCrfCode As Long) As Boolean
    Dim ilTestVefCode As Integer
    Dim slSQLQuery As String
    Dim ilCvf As Integer
    
    On Error GoTo Err:
    mTestCvf = False
    slSQLQuery = "Select * from cvf_Copy_Vehicles"
    slSQLQuery = slSQLQuery & " Where cvfCrfCode = " & llCrfCode
    Set cvf_rst = gSQLSelectCall(slSQLQuery)
    Do While Not cvf_rst.EOF
        For ilCvf = 1 To 100 Step 1
            If ilCvf <= 50 Then
                ilTestVefCode = Switch(ilCvf = 1, cvf_rst!cvfVefCode1, ilCvf = 2, cvf_rst!cvfVefCode2, ilCvf = 3, cvf_rst!cvfVefCode3, ilCvf = 4, cvf_rst!cvfVefCode4, ilCvf = 5, cvf_rst!cvfVefCode5, _
                ilCvf = 6, cvf_rst!cvfVefCode6, ilCvf = 7, cvf_rst!cvfVefCode7, ilCvf = 8, cvf_rst!cvfVefCode8, ilCvf = 9, cvf_rst!cvfVefCode9, ilCvf = 10, cvf_rst!cvfVefCode10, _
                ilCvf = 11, cvf_rst!cvfVefCode11, ilCvf = 12, cvf_rst!cvfVefCode12, ilCvf = 13, cvf_rst!cvfVefCode13, ilCvf = 14, cvf_rst!cvfVefCode14, ilCvf = 15, cvf_rst!cvfVefCode15, _
                ilCvf = 16, cvf_rst!cvfVefCode16, ilCvf = 17, cvf_rst!cvfVefCode17, ilCvf = 18, cvf_rst!cvfVefCode18, ilCvf = 19, cvf_rst!cvfVefCode19, ilCvf = 20, cvf_rst!cvfVefCode20, _
                ilCvf = 21, cvf_rst!cvfVefCode21, ilCvf = 22, cvf_rst!cvfVefCode22, ilCvf = 23, cvf_rst!cvfVefCode23, ilCvf = 24, cvf_rst!cvfVefCode24, ilCvf = 25, cvf_rst!cvfVefCode25, _
                ilCvf = 26, cvf_rst!cvfVefCode26, ilCvf = 27, cvf_rst!cvfVefCode27, ilCvf = 28, cvf_rst!cvfVefCode28, ilCvf = 29, cvf_rst!cvfVefCode29, ilCvf = 30, cvf_rst!cvfVefCode30, _
                ilCvf = 31, cvf_rst!cvfVefCode31, ilCvf = 32, cvf_rst!cvfVefCode32, ilCvf = 33, cvf_rst!cvfVefCode33, ilCvf = 34, cvf_rst!cvfVefCode34, ilCvf = 35, cvf_rst!cvfVefCode35, _
                ilCvf = 36, cvf_rst!cvfVefCode36, ilCvf = 37, cvf_rst!cvfVefCode37, ilCvf = 38, cvf_rst!cvfVefCode38, ilCvf = 39, cvf_rst!cvfVefCode39, ilCvf = 40, cvf_rst!cvfVefCode40, _
                ilCvf = 41, cvf_rst!cvfVefCode41, ilCvf = 42, cvf_rst!cvfVefCode42, ilCvf = 43, cvf_rst!cvfVefCode43, ilCvf = 44, cvf_rst!cvfVefCode44, ilCvf = 45, cvf_rst!cvfVefCode45, _
                ilCvf = 46, cvf_rst!cvfVefCode46, ilCvf = 47, cvf_rst!cvfVefCode47, ilCvf = 48, cvf_rst!cvfVefCode48, ilCvf = 49, cvf_rst!cvfVefCode49, ilCvf = 50, cvf_rst!cvfVefCode50)
            Else
                ilTestVefCode = Switch(ilCvf = 51, cvf_rst!cvfVefCode51, ilCvf = 52, cvf_rst!cvfVefCode52, ilCvf = 53, cvf_rst!cvfVefCode53, ilCvf = 54, cvf_rst!cvfVefCode54, ilCvf = 55, cvf_rst!cvfVefCode55, _
                ilCvf = 56, cvf_rst!cvfVefCode56, ilCvf = 57, cvf_rst!cvfVefCode57, ilCvf = 58, cvf_rst!cvfVefCode58, ilCvf = 59, cvf_rst!cvfVefCode59, ilCvf = 60, cvf_rst!cvfVefCode60, _
                ilCvf = 61, cvf_rst!cvfVefCode61, ilCvf = 62, cvf_rst!cvfVefCode62, ilCvf = 63, cvf_rst!cvfVefCode63, ilCvf = 64, cvf_rst!cvfVefCode64, ilCvf = 65, cvf_rst!cvfVefCode65, _
                ilCvf = 66, cvf_rst!cvfVefCode66, ilCvf = 67, cvf_rst!cvfVefCode67, ilCvf = 68, cvf_rst!cvfVefCode68, ilCvf = 69, cvf_rst!cvfVefCode69, ilCvf = 70, cvf_rst!cvfVefCode70, _
                ilCvf = 71, cvf_rst!cvfVefCode71, ilCvf = 72, cvf_rst!cvfVefCode72, ilCvf = 73, cvf_rst!cvfVefCode73, ilCvf = 74, cvf_rst!cvfVefCode74, ilCvf = 75, cvf_rst!cvfVefCode75, _
                ilCvf = 76, cvf_rst!cvfVefCode76, ilCvf = 77, cvf_rst!cvfVefCode77, ilCvf = 78, cvf_rst!cvfVefCode78, ilCvf = 79, cvf_rst!cvfVefCode79, ilCvf = 80, cvf_rst!cvfVefCode80, _
                ilCvf = 81, cvf_rst!cvfVefCode81, ilCvf = 82, cvf_rst!cvfVefCode82, ilCvf = 83, cvf_rst!cvfVefCode83, ilCvf = 84, cvf_rst!cvfVefCode84, ilCvf = 85, cvf_rst!cvfVefCode85, _
                ilCvf = 86, cvf_rst!cvfVefCode86, ilCvf = 87, cvf_rst!cvfVefCode87, ilCvf = 88, cvf_rst!cvfVefCode88, ilCvf = 89, cvf_rst!cvfVefCode89, ilCvf = 90, cvf_rst!cvfVefCode90, _
                ilCvf = 91, cvf_rst!cvfVefCode91, ilCvf = 92, cvf_rst!cvfVefCode92, ilCvf = 93, cvf_rst!cvfVefCode93, ilCvf = 44, cvf_rst!cvfVefCode94, ilCvf = 95, cvf_rst!cvfVefCode95, _
                ilCvf = 96, cvf_rst!cvfVefCode96, ilCvf = 97, cvf_rst!cvfVefCode97, ilCvf = 98, cvf_rst!cvfVefCode98, ilCvf = 49, cvf_rst!cvfVefCode99, ilCvf = 100, cvf_rst!cvfVefCode100)
            End If
            If ilTestVefCode > 0 Then
                If ilVefCode = ilTestVefCode Then
                    mTestCvf = True
                    Exit Function
                End If
                If mTestPvf(ilVefCode, ilTestVefCode) Then
                    mTestCvf = True
                    Exit Function
                End If
            End If
        Next ilCvf
        cvf_rst.MoveNext
    Loop
    Exit Function
Err:
    Resume Next
End Function

Private Function mTestPvf(ilVefCode As Integer, ilPkgVefCode As Integer) As Boolean
    Dim ilTestVefCode As Integer
    Dim slSQLQuery As String
    Dim ilPvf As Integer
    Dim llPvfCode As Long
    
    On Error GoTo Err:
    mTestPvf = False
    slSQLQuery = "Select vefType, vefPvfCode from vef_Vehicles"
    slSQLQuery = slSQLQuery & " Where vefCode = " & ilPkgVefCode
    Set vef_rst = gSQLSelectCall(slSQLQuery)
    If vef_rst!vefType = "P" Then
        llPvfCode = vef_rst!vefPvfCode
        Do While llPvfCode > 0
            slSQLQuery = "Select * from PVF_Package_Vehicle"
            slSQLQuery = slSQLQuery & " Where pvfCode = " & llPvfCode
            Set pvf_rst = gSQLSelectCall(slSQLQuery)
            If Not pvf_rst.EOF Then
                For ilPvf = 1 To 25 Step 1
                    ilTestVefCode = Switch(ilPvf = 1, pvf_rst!pvfVefCode1, ilPvf = 2, pvf_rst!pvfVefCode2, ilPvf = 3, pvf_rst!pvfVefCode3, ilPvf = 4, pvf_rst!pvfVefCode4, ilPvf = 5, pvf_rst!pvfVefCode5, _
                    ilPvf = 6, pvf_rst!pvfVefCode6, ilPvf = 7, pvf_rst!pvfVefCode7, ilPvf = 8, pvf_rst!pvfVefCode8, ilPvf = 9, pvf_rst!pvfVefCode9, ilPvf = 10, pvf_rst!pvfVefCode10, _
                    ilPvf = 11, pvf_rst!pvfVefCode11, ilPvf = 12, pvf_rst!pvfVefCode12, ilPvf = 13, pvf_rst!pvfVefCode13, ilPvf = 14, pvf_rst!pvfVefCode14, ilPvf = 15, pvf_rst!pvfVefCode15, _
                    ilPvf = 16, pvf_rst!pvfVefCode16, ilPvf = 17, pvf_rst!pvfVefCode17, ilPvf = 18, pvf_rst!pvfVefCode18, ilPvf = 19, pvf_rst!pvfVefCode19, ilPvf = 20, pvf_rst!pvfVefCode20, _
                    ilPvf = 21, pvf_rst!pvfVefCode21, ilPvf = 22, pvf_rst!pvfVefCode22, ilPvf = 23, pvf_rst!pvfVefCode23, ilPvf = 24, pvf_rst!pvfVefCode24, ilPvf = 25, pvf_rst!pvfVefCode25)
                    If ilTestVefCode > 0 Then
                        If ilVefCode = ilTestVefCode Then
                            mTestPvf = True
                            Exit Function
                        End If
                    End If
                Next ilPvf
                llPvfCode = pvf_rst!pvfLkPvfCode
            Else
                Exit Do
            End If
        Loop
    End If
    Exit Function
Err:
    Resume Next
End Function


Public Sub gClosePoolFiles()
    On Error Resume Next
    vef_rst.Close
    pvf_rst.Close
    cvf_rst.Close
    sdf_rst.Close
    cnf_rst.Close
End Sub

Private Sub mGetPoolCopy(ilAdfCode As Integer, ilNextFinal As Integer, llPoolCrfCode As Long, tlRegionDefinition As REGIONDEFINITION)
    Dim slSQLQuery As String
    Dim blUpdateCrf As Boolean
    Dim ilNextInstrNo As Integer
    Dim ilRet As Integer
    
    blUpdateCrf = False
    'Check if already assigned
    If tlRegionDefinition.lCopyCode > 0 And tlRegionDefinition.sPtType = "1" Then
        tlRegionDefinition.bPoolUpdated = True
        tlRegionDefinition.iPoolNextFinal = ilNextInstrNo
        tlRegionDefinition.iPoolAdfCode = ilAdfCode
        tlRegionDefinition.lPoolCrfCode = llPoolCrfCode
        Exit Sub
    End If
    slSQLQuery = "Select * FROM cnf_Copy_Instruction"
    slSQLQuery = slSQLQuery & " Where (cnfCrfCode = " & tlRegionDefinition.lCrfCode
    slSQLQuery = slSQLQuery & " And cnfInstrNo = " & ilNextFinal
    slSQLQuery = slSQLQuery & ")"
    Set cnf_rst = gSQLSelectCall(slSQLQuery)
    If Not cnf_rst.EOF Then
        tlRegionDefinition.sPtType = "1"
        tlRegionDefinition.lCopyCode = cnf_rst!cnfCifCode
        blUpdateCrf = True
        ilNextInstrNo = ilNextFinal + 1
    Else
        slSQLQuery = "Select * FROM cnf_Copy_Instruction"
        slSQLQuery = slSQLQuery & " Where (cnfCrfCode = " & tlRegionDefinition.lCrfCode
        slSQLQuery = slSQLQuery & " And cnfInstrNo = " & 1
        slSQLQuery = slSQLQuery & ")"
        Set cnf_rst = gSQLSelectCall(slSQLQuery)
        If Not cnf_rst.EOF Then
            tlRegionDefinition.sPtType = "1"
            tlRegionDefinition.lCopyCode = cnf_rst!cnfCifCode
            blUpdateCrf = True
            ilNextInstrNo = 2
        End If
    End If
    If blUpdateCrf Then
        tlRegionDefinition.iPoolNextFinal = ilNextInstrNo
        tlRegionDefinition.iPoolAdfCode = ilAdfCode
        tlRegionDefinition.lPoolCrfCode = llPoolCrfCode
    End If
End Sub

Private Sub mUpdatePoolInfo(ilAdfCode As Integer, llCrfCode As Long, ilNextInstrNo As Integer, tlRegionDefinition As REGIONDEFINITION)
    Dim slSQLQuery As String
    Dim ilRet As Integer
    If ilAdfCode <= 0 Then
        Exit Sub
    End If
    
    If tlRegionDefinition.lPoolCrfCode <= 0 Then
        Exit Sub
    End If
    'Check if already assigned
    If tlRegionDefinition.lCopyCode > 0 And tlRegionDefinition.sPtType = "1" And tlRegionDefinition.bPoolUpdated Then
        Exit Sub
    End If

    tlRegionDefinition.bPoolUpdated = True
    
    slSQLQuery = "Update adf_Advertisers Set adfBkoutPoolStatus = 'A'"
    slSQLQuery = slSQLQuery & " Where adfCode = " & ilAdfCode
    ilRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    
    igLastPoolAdfCode = ilAdfCode
    
    slSQLQuery = "Update RSF_Region_Schd_Copy Set rstPtType = '" & tlRegionDefinition.sPtType & "',"
    slSQLQuery = slSQLQuery & "rsfCopyCode = " & tlRegionDefinition.lCopyCode
    slSQLQuery = slSQLQuery & " Where rsfCode = " & tlRegionDefinition.lRsfCode
    ilRet = gSQLWaitNoMsgBox(slSQLQuery, False)
    
End Sub

Private Sub mGetAdjAdvt(llSdfCode As Long, ilVefCode As Integer, ilPriorAdjAdfCode As Integer, ilNextAdjAdfCode As Integer)
    Dim slSQLQuery As String
    Dim llLstLogTime As Long
    Dim lst_pool As ADODB.Recordset
    Dim lst_Adj As ADODB.Recordset

    ilPriorAdjAdfCode = -1
    ilNextAdjAdfCode = -1
    slSQLQuery = "Select lstLogVefCode, lstLogDate, LstLogTime, lstBreakNo, lstPositionNo From lst"
    slSQLQuery = slSQLQuery & " Where lstSdfCode = " & llSdfCode
    slSQLQuery = slSQLQuery & " And lstLogVefCode = " & ilVefCode
    Set lst_pool = gSQLSelectCall(slSQLQuery)
    If Not lst_pool.EOF Then
        llLstLogTime = gTimeToLong(Format(lst_pool!lstLogTime, "h:mm:ssAM/PM"), False)
        If (llSdfCode = lmAdjSdfCode) And (ilVefCode = imAdjVefCode) Then
            If llLstLogTime = lmAdjLstLogTime Then
                ilPriorAdjAdfCode = imPriorAdjAdfCode
                ilNextAdjAdfCode = imNextAdjAdfCode
                Exit Sub
            End If
        End If
        'Find previous
        'igLastPoolAdfCode contain the previous blackout assigned
        'Find the original previous spot to avoid possible advt matching with the next blackout advt to be assigned
        If lst_pool!lstPositionNo = 1 Then
            slSQLQuery = "Select lstAdfCode From lst"
            slSQLQuery = slSQLQuery & " Where  lstLogVefCode = " & ilVefCode
            slSQLQuery = slSQLQuery & " And lstLogDate = '" & Format(lst_pool!lstLogDate, sgSQLDateForm) & "'"
            slSQLQuery = slSQLQuery & " And lstLogTime < '" & Format(lst_pool!lstLogTime, sgSQLTimeForm) & "'"
            slSQLQuery = slSQLQuery & " And lstSdfCode > 0"
            slSQLQuery = slSQLQuery & " And lstBkoutLstCode = " & 0
            slSQLQuery = slSQLQuery & " Order By lstPositionNo Desc"
        Else
            'The prior blackout would have been just assigned
            slSQLQuery = "Select lstAdfCode From lst"
            slSQLQuery = slSQLQuery & " Where  lstLogVefCode = " & ilVefCode
            slSQLQuery = slSQLQuery & " And lstLogDate = '" & Format(lst_pool!lstLogDate, sgSQLDateForm) & "'"
            slSQLQuery = slSQLQuery & " And lstLogTime = '" & Format(lst_pool!lstLogTime, sgSQLTimeForm) & "'"
            slSQLQuery = slSQLQuery & " And lstPositionNo < " & lst_pool!lstPositionNo
            slSQLQuery = slSQLQuery & " And lstSdfCode > 0"
            slSQLQuery = slSQLQuery & " And lstBkoutLstCode = " & 0
            slSQLQuery = slSQLQuery & " Order By lstPositionNo Desc"
        End If
        Set lst_Adj = gSQLSelectCall(slSQLQuery)
        If Not lst_Adj.EOF Then
            ilPriorAdjAdfCode = lst_Adj!lstAdfCode
        End If
        'Find Next
        'The next spot will be assigned and check against this spot, so get original next spot to avois possible advt match
        slSQLQuery = "Select lstAdfCode From lst"
        slSQLQuery = slSQLQuery & " Where  lstLogVefCode = " & ilVefCode
        slSQLQuery = slSQLQuery & " And lstLogDate = '" & Format(lst_pool!lstLogDate, sgSQLDateForm) & "'"
        slSQLQuery = slSQLQuery & " And lstLogTime = '" & Format(lst_pool!lstLogTime, sgSQLTimeForm) & "'"
        slSQLQuery = slSQLQuery & " And lstSdfCode > 0"
        slSQLQuery = slSQLQuery & " And lstBkoutLstCode = " & 0
        slSQLQuery = slSQLQuery & " And lstPositionNo = " & lst_pool!lstPositionNo + 1
        Set lst_Adj = gSQLSelectCall(slSQLQuery)
        If lst_Adj.EOF Then
            slSQLQuery = "Select lstAdfCode From lst"
            slSQLQuery = slSQLQuery & " Where  lstLogVefCode = " & ilVefCode
            slSQLQuery = slSQLQuery & " And lstLogDate = '" & Format(lst_pool!lstLogDate, sgSQLDateForm) & "'"
            slSQLQuery = slSQLQuery & " And lstLogTime > '" & Format(lst_pool!lstLogTime, sgSQLTimeForm) & "'"
            slSQLQuery = slSQLQuery & " And lstSdfCode > 0"
            slSQLQuery = slSQLQuery & " And lstBkoutLstCode = 0"
            slSQLQuery = slSQLQuery & " And lstPositionNo = 1"
            Set lst_Adj = gSQLSelectCall(slSQLQuery)
        End If
        If Not lst_Adj.EOF Then
            ilNextAdjAdfCode = lst_Adj!lstAdfCode
        End If
        
    End If
    lmAdjSdfCode = llSdfCode
    lmAdjLstLogTime = llLstLogTime
    imPriorAdjAdfCode = ilPriorAdjAdfCode
    imNextAdjAdfCode = ilNextAdjAdfCode
End Sub

Public Function gGetAstInfo(hlAst As Integer, tlCPDat() As DAT, tlAstInfo() As ASTINFO, ilInAdfCode As Integer, iAddAst As Integer, iUpdateCpttStatus As Integer, ilBuildAstInfo As Integer, Optional blInGetRegionCopy As Boolean = True, Optional llSelGsfCode As Long = -1, Optional blFeedAdjOnReturn As Boolean = False, Optional blFilterByAirDates As Boolean = False, Optional blIncludePledgeInfo As Boolean = True, Optional blCreateServiceATTSpots As Boolean = False) As Boolean
    Dim ilMShttCode As Integer
    Dim slAstStatus As String
    Dim llMCpttCode As Long
    Dim blRet As Boolean
    Dim slATTMultcast As String
    
    smStatiomType = "N"
'    slATTMultcast = mIsAttMulticast(tgCPPosting(0).lAttCode)
'    If slATTMultcast = "Y" Then
'        ilMShttCode = mGetMasterMulticast(tgCPPosting(0).iShttCode)
'        If ilMShttCode > 0 Then
'            If ilMShttCode <> tgCPPosting(0).iShttCode Then
'                slAstStatus = mGetCPTTAstStatus(ilMShttCode, tgCPPosting(0).iVefCode, tgCPPosting(0).sDate, llMCpttCode, lmMATTCode)
'                If slAstStatus <> "C" Then
'                    blRet = mBuildCPPosting(llMCpttCode, tgCPPosting(0).sDate)
'                    If blRet Then
'                        smStatiomType = "M"
'                        tmSvCPPosting = tgCPPosting(0)
'                        tgCPPosting(0) = tmCPPosting
'                        tgCPPosting(0).iNumberDays = tmSvCPPosting.iNumberDays
'                        blRet = mGetAstInfo(hlAst, tlCPDat(), tlAstInfo(), ilInAdfCode, iAddAst, iUpdateCpttStatus, ilBuildAstInfo, blInGetRegionCopy, llSelGsfCode, blFeedAdjOnReturn, blFilterByAirDates, blIncludePledgeInfo, blCreateServiceATTSpots)
'                        smStatiomType = "C"
'                        tgCPPosting(0) = tmSvCPPosting
'                        tgCPPosting(0).sAstStatus = "N"
'                    Else
'                        smStatiomType = "C"
'                    End If
'                Else
'                    smStatiomType = "C"
'                End If
'            Else
'                smStatiomType = "M"
'            End If
'        End If
'    End If
    blRet = mGetAstInfo(hlAst, tlCPDat(), tlAstInfo(), ilInAdfCode, iAddAst, iUpdateCpttStatus, ilBuildAstInfo, blInGetRegionCopy, llSelGsfCode, blFeedAdjOnReturn, blFilterByAirDates, blIncludePledgeInfo, blCreateServiceATTSpots)
    gGetAstInfo = blRet
End Function

Private Function mGetMasterMulticast(ilShttCode As Integer) As Integer
    'ilShttCode (I): Station to determine if multicast and which station is the master
    'mGetMasterMulticast (O): -1 if not multicast; > 0 if Multicast and station code that is the master
    Dim ilShtt As Integer
    Dim slSQLQuery As String
    
    mGetMasterMulticast = -1
    ilShtt = gBinarySearchStationInfoByCode(ilShttCode)
    If ilShtt = -1 Then
        Exit Function
    End If
    If tgStationInfoByCode(ilShtt).lMultiCastGroupID <= 0 Then
        Exit Function
    End If
    If tgStationInfoByCode(ilShtt).sMasterCluster = "Y" Then
        mGetMasterMulticast = ilShttCode
        Exit Function
    End If
    slSQLQuery = "Select shttCode FROM shtt"
    slSQLQuery = slSQLQuery & " where shttMultiCastGroupID = " & tgStationInfoByCode(ilShtt).lMultiCastGroupID
    slSQLQuery = slSQLQuery & " where shttMasterCluster = 'Y'"
    Set rst_Shtt = gSQLSelectCall(slSQLQuery)
    If rst_Shtt.EOF Then
        Exit Function
    End If
    mGetMasterMulticast = rst_Shtt!shttCode
    
End Function

Private Function mGetCPTTAstStatus(ilShttCode As Integer, ilVefCode As Integer, slDate As String, llCpttCode As Long, llAttCode As Long) As String
    Dim slSQLQuery As String
    
    mGetCPTTAstStatus = ""
    llCpttCode = -1
    llAttCode = -1
    'Get attCode
    slSQLQuery = "SELECT attCode, attTimeType FROM att "
    slSQLQuery = slSQLQuery & " Where attShfCode = " & ilShttCode
    slSQLQuery = slSQLQuery & " AND attVefCode = " & ilVefCode
    slSQLQuery = slSQLQuery & " AND attOnAir <= '" & Format(slDate, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " AND attOffAir >= '" & Format(slDate, sgSQLDateForm) & "'"
    slSQLQuery = slSQLQuery & " AND attDropDate >= '" & Format(slDate, sgSQLDateForm) & "'"
    Set att_rst = gSQLSelectCall(slSQLQuery)
    If att_rst.EOF Then
        Exit Function
    End If
    'Get Cptt
    slSQLQuery = "SELECT cpttCode, cpttAstStatus"
    slSQLQuery = slSQLQuery & " FROM cptt WHERE ("
    slSQLQuery = slSQLQuery & " AND cpttAtfCode = " & att_rst!attCode
    slSQLQuery = slSQLQuery & " AND cpttStartDate = '" & Format$(gObtainPrevMonday(slDate), sgSQLDateForm) & "')"
    Set cptt_rst = gSQLSelectCall(slSQLQuery)
    If Not cptt_rst.EOF Then
        llAttCode = att_rst!attCode
        llCpttCode = cptt_rst!cpttCode
        mGetCPTTAstStatus = cptt_rst!cpttAstStatus
    End If
End Function

Private Function mBuildCPPosting(llCpttCode As Long, slDate As String) As Boolean
    Dim slSQLQuery As String
    
    mBuildCPPosting = False
    'Get Cptt
    slSQLQuery = "SELECT cpttCode,cpttStatus,cpttPostingStatus,cpttAstStatus,shttTimeZone,ShttackDaylight as Daylight, shttTztCode as TimeZone, cpttshfcode, cpttvefcode, attCode, attTimeType"
    slSQLQuery = slSQLQuery & " FROM cptt "
    slSQLQuery = slSQLQuery & " left outer join shtt on cpttShfCode = shttCode "
    slSQLQuery = slSQLQuery & " Left Outer Join att On cpttAtfCode = attCode "
    slSQLQuery = slSQLQuery & " WHERE cpttCode = " & llCpttCode
    slSQLQuery = slSQLQuery & " And cpttStartDate = '" & Format$(gObtainPrevMonday(slDate), sgSQLDateForm) & "')"
    Set cptt_rst = gSQLSelectCall(slSQLQuery)
    If Not cptt_rst.EOF Then
        mBuildCPPosting = True
        tmCPPosting.lCpttCode = cptt_rst!cpttCode
        tmCPPosting.iStatus = cptt_rst!cpttStatus
        tmCPPosting.iPostingStatus = cptt_rst!cpttPostingStatus
        tmCPPosting.lAttCode = cptt_rst!attCode
        tmCPPosting.iAttTimeType = cptt_rst!attTimeType
        tmCPPosting.iVefCode = cptt_rst!cpttvefcode
        tmCPPosting.iShttCode = cptt_rst!cpttshfcode
        tmCPPosting.sZone = cptt_rst!shttTimeZone
        tmCPPosting.sDate = slDate
        tmCPPosting.sAstStatus = cptt_rst!cpttAstStatus
    End If
End Function


Private Function mIsAttMulticast(llAttCode As Long) As String
    Dim slSQLQuery As String
    
    mIsAttMulticast = "N"
    slSQLQuery = "SELECT attMulticast FROM att "
    slSQLQuery = slSQLQuery & " Where attCode = " & llAttCode
    Set att_rst = gSQLSelectCall(slSQLQuery)
    If Not att_rst.EOF Then
        mIsAttMulticast = att_rst!attMulticast
    End If
End Function

Public Function gPoolExist() As Boolean
    Dim slSQLQuery As String
    gPoolExist = False
    slSQLQuery = "Select Count(1) as BkoutCount from Adf_Advertisers"
    slSQLQuery = slSQLQuery & " Where "
    slSQLQuery = slSQLQuery & " (adfBkoutPoolStatus = 'A'"
    slSQLQuery = slSQLQuery & " Or adfBkoutPoolStatus = 'U'" & ")"
    Set rst_adf = gSQLSelectCall(slSQLQuery)
    If Not rst_adf.EOF Then
        If rst_adf!BkoutCount > 0 Then
            gPoolExist = True
        End If
    End If
    rst_adf.Close
End Function
