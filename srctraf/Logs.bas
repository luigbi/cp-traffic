Attribute VB_Name = "LOGSSUBS"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of Logs.bas on Wed 6/17/09 @ 12:56 PM **
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************

' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: Logs.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Form subs and functions
Option Explicit
Option Compare Text

'******************************************************************************
' abf_Ast_Build_Queue Record Definition
'
'******************************************************************************
Type ABF
    lCode                 As Long            ' Auto-Increment AST Build Queue
    sSource               As String * 1      ' Source: L=Log; A=Agreement;
                                             ' P=Affiliate Post Log; F=Fast Add
    iVefCode              As Integer         ' Vehicle reference code
    iShttCode             As Integer         ' Station reference code (aqfSource
                                             ' = A or P or F; otherwise 0)
    sStatus               As String * 1      ' H=Hold; G=Generate AST;
                                             ' C=Completed; A= AST Generating
    iMondayDate(0 To 1)   As Integer         ' Monday date
    iGenStartDate(0 To 1) As Integer         ' Generation start date
    iGenEndDate(0 To 1)   As Integer         ' Generation end date
    iEnteredDate(0 To 1)  As Integer         ' Entered date
    iEnteredTime(0 To 1)  As Integer         ' Entered time
    iCompletedDate(0 To 1) As Integer        ' Completed date
    iCompletedTime(0 To 1) As Integer        ' Completed time
    iUrfCode              As Integer         ' Urf reference code (aqfSource =
                                             ' L)
    iUstCode              As Integer         ' User Reference code (aqfSource =
                                             ' A; P or F)
    sUnused               As String * 10
End Type


'Type ABFKEY0
'    lCode                 As Long
'End Type

Type ABFKEY1
    sStatus               As String * 1
    iGenStartDate(0 To 1) As Integer
    iEnteredDate(0 To 1)  As Integer
    iEnteredTime(0 To 1)  As Integer
End Type

Type ABFKEY2
    sStatus               As String * 1
    iVefCode              As Integer
    iShttCode             As Integer
    iMondayDate(0 To 1)   As Integer
End Type


Type ABFINFO
    sStatus As String * 1   'N=New; C=Changed; S=Saved
    lCode As Long
    iVefCode As Integer
    lMondayDate As Long
    lStartDate As Long
    lEndDate As Long
End Type

Public tgAbfInfo() As ABFINFO
Public sgGenDate As String
Public igGenDate(0 To 1) As Integer     'Log Generation Date, used to date stamp ODF
Public sgGenTime As String
Public igGenTime(0 To 1) As Integer     'Log Generation time, used to time stamp ODF
Public lgGenTime As Long                 'log generation time
Type LSTUPDATEINFO
    iType As Integer        '0=Spot; 1=Copy
    iVefCode As Integer     'Vehicle Code with Alter
    lSDate As Long          'Week start date
    lEDate As Long          'Week end date
End Type
Public tgLSTUpdateInfo() As LSTUPDATEINFO     '
Type LOGSEL
    sKey As String * 80
    iChk As Integer  'Checked Y(1) or N(0)
    iInitChk As Integer  'Checked Y(1) or N(0)
    sWrkDate As String * 10
    sVehicle As String * 40
    iVefCode As Integer
    iVpfIndex As Integer
    sLLD As String * 10  'Last Log Date
    iLLDChgAllowed As Integer   'If LLD was zero, allow LLD to be changed
    iLeadTime As Integer    'Lead Time
    iCycle As Integer   'Cycle
    lStartDate As Long  'Next Closing Start Date
    lEndDate As Long    'Next Closing End Date
    iLog As Integer     'Index for Log
    iCP As Integer      'Index for CP
    iLogo As Integer    'Index for Logo
    iOther As Integer    'Index for Play List
    iZone As Integer    'Index for Zone
    iChg As Integer
    iStatus As Integer  '0=Allow Log to be generate; 1=Links missing; 2=Selling not scheduled;
                        '3=Conventional missing; 4=Conventional for Logs not scheduled; 5=Not scheduled
End Type
Type LOGGEN
    iGenVefCode As Integer
    iSimVefCode As Integer
End Type

Type LOGEXPORTLOC           '6-4-19
    sKey As String * 5
    iVefCode As Integer
    sExportPath As String * 250
End Type

'In RptGen
'Type COPYROTNO
'    iRotNo As Integer
'    sZone As String * 3
'End Type
Type ZONEINFO
    sZone As String * 1
    lGLocalAdj As Long
    lGFeedAdj As Long
    sFed As String * 1
    lcpfCode As Long
End Type
Type PAGEEJECT
    lTime As Long
    ianfCode As Integer
End Type
Type BBVEFINFO
    iVefCode As Integer
    lStartDate As Long
    lEndDate As Long
End Type
Type BBSDFINFO
    sType As String * 1 'A=Avail; O=Open; C=Close
    lChfCode As Long
    iTime(0 To 1) As Integer
    iLen As Integer
    iBreakNo As Integer
    iPositionNo As Integer
End Type
Public tgLogExportLoc() As LOGEXPORTLOC           '6-4-19
Public tgSel() As LOGSEL
Public tgLogSel() As LOGSEL
Public tgRPSel() As LOGSEL
Public tgAlertSel() As LOGSEL
Public tgChkMsg() As LOGSEL
Public tgLogChkMsg() As LOGSEL
Dim tmPageEject() As PAGEEJECT
Public igSVefCode() As Integer

Public bgReprintLogType As Integer

Public tgStations() As SHTTINFO
Public sgStationsStamp As String

Type REGIONSTATIONINFO
    sKey As String * 40     'Call letters
    iShttCode As Integer    'Auto Inc
    sCallLetters As String * 40 'Call letters or Persons Name
    iMktCode As Integer       ' Reference to Market Table
    sMarket As String * 60  'Market
End Type

Type COMBINELSTINFO
    lDate As Long
    tLst As LST
End Type

Type SPLITBOFREC
    sKey As String * 20 'Random number
    tBof As BOF
    iLen As Integer
End Type

Type SPLITNETLASTFILL
    iBofIndex As Integer
    iFillLen As Integer
End Type

Type SPLITODFREC
    sKey As String * 30
    tOdf As ODF
End Type

'1-22-14        array of ntr items for L87 indicating the vehicle and its starting and ending index in tgSBFInfo()
Type NTRSORTINFO
    iVefCode As Integer
    iLoInx As Integer
    iHiInx As Integer
End Type

Type NTRINFO
    iBillVefCode As Integer     'use key, since using the index into array isnt predictable
    tNTR As SBF
End Type

'1-28-14 this type statment moved from rptrec due to Logs requireing the NTR records and other projects have Log.bas in their project leaving unreferenced items
'6-9-14 moved again with sbf recd definition
'Type SBFTypes                   '9-30-02
'    iNTR As Integer             'include "I" SBFTypes
'    iInstallment As Integer     'include "F" SBFTypes
'    iImport As Integer          'include "T" SBFTypes
'End Type

Public tgNtrSortInfo() As NTRSORTINFO
Public tgNTRInfo() As NTRINFO

'Log Spot record
Dim hmLst As Integer        'Log Spots file
Dim tmLst As LST
Dim imLstRecLen As Integer
Dim tmLstSrchKey As LONGKEY0
Dim tmLstSrchKey2 As LSTKEY2
Dim imLstExist As Integer       'False=LST Does NOT exist, don't update
Dim tmLstCode() As LST
Public tmCombineLstInfo() As COMBINELSTINFO
'Media code record information
Dim hmMcf As Integer
Dim tmMcfSrchKey As INTKEY0 'MCF key record image
Dim imMcfRecLen As Integer  'MCF record length
Dim tmMcf As MCF            'MCF record image
'Line record
Dim hmClf As Integer        'Line file
Dim tmClfSrchKey As CLFKEY0 'CLF key record image
Dim imClfRecLen As Integer  'CLF record length
Dim tmClf As CLF
'Flight record
Dim hmCff As Integer        'Flight file
Dim tmCff As CFF
Dim tmFCff() As CFF
'MG Spec record
Dim hmSmf As Integer        'MG Spec file
'Demo Data record
Dim hmDrf As Integer        'Research Data file
'Demo Plus Data record
Dim hmDpf As Integer        'Research Demo Plus Data file
'Research Estimate
Dim hmDef As Integer
Dim hmRaf As Integer
'Product/ISCI record
Dim hmCpf As Integer        'Product file
Dim tmCpf As CPF
Dim tmCpfSrchKey As LONGKEY0
Dim imCpfRecLen As Integer
'MultiName record
Dim hmMnf As Integer        'MNF file
Dim tmMnf As MNF
Dim tmMnfSrchKey As INTKEY0
Dim imMnfRecLen As Integer

Dim hmCvf As Integer

Dim hmSsf As Integer
Dim tmSsf As SSF                'SSF record image
Dim imSsfRecLen As Integer

'Dim tmSsfOld As SSF
Dim tmProg As PROGRAMSS
Dim tmAvail As AVAILSS
Dim tmSpot As CSPOTSS
Dim tmAvailTest As AVAILSS
Dim tmCTSsf As SSF               'Ssf for conflict test
Dim tmAvAvail As AVAILSS
Dim tmAvSpot As CSPOTSS
Dim tmOpenAvail As AVAILSS
Dim tmCloseAvail As AVAILSS
'Spot record
Dim hmSdf As Integer
Dim tmSdf As SDF
Dim imSdfRecLen As Integer
Dim tmSdfSrchKey3 As LONGKEY0
Dim tmSdfSrchKey0 As SDFKEY0
Dim tmBBSdfInfo() As BBSDFINFO
'Contract record information
Dim hmCHF As Integer
Dim tmChfSrchKey As LONGKEY0 'CHF key record image
Dim imCHFRecLen As Integer  'CHF record length
Dim tmChf As CHF            'CHF record image
'One day file (ODF)
Dim hmOdf As Integer        'One day file
Dim imOdfRecLen As Integer  'ODF record length
Dim tmOdfSrchKey0 As ODFKEY0 'ODF key record image
Dim tmOdfSrchKey1 As ODFKEY1 'ODF key record image
Dim tmOdfSrchKey2 As ODFKEY2 'ODF key record image
Dim tmOdf As ODF            'ODF record image
' Vehicle File
Dim hmVef As Integer        'Vehicle file handle
Dim tmVef As VEF            'VEF record image
Dim tmVefSrchKey As INTKEY0 'VEF key record image
Dim imVefRecLen As Integer     'VEF record length

Dim tmVpf As VPF                'VPF record image
Dim imVpfRecLen As Integer      'VPF record length
Dim hmVpf As Integer            'Vehicle preference file handle

'Delivery file (DLF)
Dim imDlfRecLen As Integer  'DLF record length
Dim tmDlfSrchKey As DLFKEY0 'DLF key record image
Dim tmDlf As DLF            'DLF record image
'Vehicle linkage record information
Dim hmVlf As Integer
Dim tmVlfSrchKey1 As VLFKEY1 'VLF key record image
Dim imVlfRecLen As Integer  'VLF record length
Dim tmVlf As VLF            'VLF record image
' Virtual Vehicle File

'Feed
Dim hmFsf As Integer
Dim imFsfRecLen As Integer
Dim tmFsfSrchKey0 As LONGKEY0
Dim tmFsf As FSF

'Feed Names
Dim hmFnf As Integer
Dim imFnfRecLen As Integer
Dim tmFnf As FNF

'Product
Dim hmPrf As Integer
Dim imPrfRecLen As Integer
Dim tmPrf As PRF

Dim tmRdf As RDF


Dim hmGhf As Integer
Dim tmGhf As GHF        'GHF record image
Dim tmCombineGhf As GHF        'GHF record image
Dim tmGhfSrchKey1 As GHFKEY1    'GHF key record image
Dim imGhfRecLen As Integer        'GHF record length

Dim hmGsf As Integer
Dim tmGsf As GSF        'GSF record image
Dim tmCombineGsf As GSF        'GSF record image
Dim tmGsfSrchKey1 As GSFKEY1    'GSF key record image
Dim tmGsfSrchKey3 As GSFKEY3    'GSF key record image
Dim imGsfRecLen As Integer        'GSF record length

'Copy inventory
Dim hmCif As Integer        'Copy inventory file handle
Dim tmCif As CIF            'CIF record image
Dim tmCifSrchKey As LONGKEY0 'CIF key record image
Dim imCifRecLen As Integer     'CIF record length

Dim hmBof As Integer        'Blackout file
Dim tmBof As BOF
Dim imBofRecLen As Integer  ' record length

Dim hmRsf As Integer        'Copy replacement file handle
Dim tmRsf As RSF
Dim imRsfRecLen As Integer
Dim tmRsfSrchKey1 As LONGKEY0 'RSF key record image

Dim hmSif As Integer        'short title file
Dim tmSif As SIF
Dim imSifRecLen As Integer  ' record length

Dim hmCrf As Integer        'Copy Rotation file handle
Dim tmCrf As CRF
Dim imCrfRecLen As Integer

Dim hmCnf As Integer        'Copy Instr file
Dim imCnfRecLen As Integer  'Copy Instr record length
Dim tmCnf As CNF
Dim tmCnfSrchKey As CNFKEY0

'1-21-14
' Special billing
Dim hmSbf As Integer        'SBF file handle
Dim tmSbf As SBF            'SBF record image
Dim tmSbfSrchKey2 As SBFKEY2 'SBF key record image
Dim imSbfRecLen As Integer     'SBF record length
Dim tmSbfSrchKey4 As SBFKEY4


Dim hmMsg As Integer        'message handle

'11/6/14: Correlate ODF and Sdf
Type ODFSDFCODES
    lOdfCode As Long
    lSdfCode As Long
End Type
Public tgOdfSdfCodes() As ODFSDFCODES

Private tmDate1ODF() As ODF
Private imDateAdj As Integer
Private smLockDate As String
Private smLockStartTime As String
Private smLockEndTime As String
Private imLockVefCode As Integer
Private lmLockStartTime As Long
Private bmCreatingLstDate1 As Boolean
Private smAdjDate As String
Private imGenODF As Integer     'True=Generate ODF
Private lmLPTime As Long     'Start Time if ilLPLocalAdj > 0
Private lmLNTime As Long     'End Time if ilLNLocalAdj < 0
Private lmGLocalAdj As Long
Private lm010570 As Long
Private imBreakNo As Integer    'Reset to zero for each program
Private imPositionNo As Integer 'Reset to zero for each avail
Private tmZoneInfo() As ZONEINFO
Private smBonus As String * 1        '2-15-01 bonus flag (B)
Private lmCrfCsfCode As Long
Private lmAvailCefCode As Long
Private lmEvtIDCefCode As Long
Private lmEvtCefCode As Long
Private lmEvtTime As Long
Private smEDIDays As String
Private imDaySort As Integer
Private imEvtCefSort As Integer     'Incremented whenever a Other Comment (14) exist within a day
Private lmAvailTime As Long
Private lmPrevAvailTime As Long
Private smLogType As String
Private imClearLstSdf As Integer
Private imCreateNewLST As Integer
Private imWegenerOLA As Integer
Private smAlertStatus As String * 1 'C=Check if alert required; A=Alert check Added
Private smXMid As String
Private hmRdf As Integer
Private hmAdf As Integer
Private hmCxf As Integer

Public bgLogFirstCallToVpfFind As Boolean
Public sgAutomationLogBuffer As String 'Log buffer
Public sgMessageFile As String   'The current Log File (as defined in various mOpenMsgFile() functions)
'
'       Update LCF record with either "I" (not complete) or "C" (complete
'       <input>  Date to search
'                ilType (game # or 0 for regular programming)
'               ilVehCode - vehicle code
'       <return> None
Public Sub gUpdateLCFCompleteFlag(hlLcf As Integer, tlLcf As LCF, ilDate0 As Integer, ilDate1 As Integer, ilType As Integer, ilVehCode As Integer, slCompleteFlag As String)
Dim ilLcfRet As Integer
Dim tlLcfSrchKey0 As LCFKEY0

        tlLcfSrchKey0.iLogDate(0) = ilDate0
        tlLcfSrchKey0.iLogDate(1) = ilDate1
        tlLcfSrchKey0.iSeqNo = 1
        tlLcfSrchKey0.iType = ilType
        tlLcfSrchKey0.iVefCode = ilVehCode
        tlLcfSrchKey0.sStatus = "C"
        ilLcfRet = btrGetEqual(hlLcf, tlLcf, Len(tlLcf), tlLcfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        If ilLcfRet = BTRV_ERR_NONE Then
            tlLcf.sAffPost = slCompleteFlag
            ilLcfRet = btrUpdate(hlLcf, tlLcf, Len(tlLcf))
        End If
        Exit Sub
End Sub
'
'           For Airing vehicles only - test the log date day to see if
'           it is a valid airing date to gather spots.  This wont
'           handle if only TFN exists without the start date day
'           SSF records.  Also, wont handle case if a day is terminated
'           i.e. M-Fr and take away Tue and Thur.  It will still
'           find those days
'
'           gTestAirVefValidDay
'           <input> slDate - log date
'                   ilvefCode - airing vehicle code
'                   tlVlf - vlf, only used if blTestAirTimeUnits is true 10-14-16
'                   blTestAirTimeUnits - true if processing airing vehicle that is to ignore air avail that is defined with 0 units or 0 seconds
'                                        if true, will loop thru the ssf to find the matching selling time with airing link.
'           <return> true if valid day
Public Function gTestAirVefValidDay(hlSsf As Integer, slDate As String, ilVehCode As Integer, tlVlf As VLF, Optional blTestAirTimeUnits As Boolean = False) As Integer
Dim ilDay As Integer
Dim ilLogDate0 As Integer
Dim ilLogDate1 As Integer
Dim ilSsfRecLen As Integer
Dim slSsfDate As String
Dim tlSsfSrchKey As SSFKEY0 'SSF key record image
Dim ilRet As Integer
Dim llLinkAirTime As Long
Dim llEvtTime As Long
Dim ilVff As Integer
Dim ilEvt As Integer

    gTestAirVefValidDay = False
    ilDay = gWeekDayStr(slDate)
    gPackDate slDate, ilLogDate0, ilLogDate1
    'If airing- then use first Ssf prior to date defined
    ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
    tlSsfSrchKey.iType = 0
    tlSsfSrchKey.iVefCode = ilVehCode
    tlSsfSrchKey.iDate(0) = ilLogDate0
    tlSsfSrchKey.iDate(1) = ilLogDate1
    tlSsfSrchKey.iStartTime(0) = 0
    tlSsfSrchKey.iStartTime(1) = 6144   '24*256
    ilRet = gSSFGetLessOrEqual(hlSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = 0) And (tmSsf.iVefCode = ilVehCode)
        gUnpackDate tmSsf.iDate(0), tmSsf.iDate(1), slSsfDate
        If (ilDay = gWeekDayStr(slSsfDate)) And (tmSsf.iStartTime(0) = 0) And (tmSsf.iStartTime(1) = 0) Then
            'test the airing vehicle to determine if honoring zero units.  if so, do not include the selling spot
            If blTestAirTimeUnits Then
                gTestAirVefValidDay = True
                'get time of airing
                gUnpackTimeLong tlVlf.iAirTime(0), tlVlf.iAirTime(1), False, llLinkAirTime
                'search the ssf for matching airing avail time to see if zero units defined
                ilEvt = 1
                Do While ilEvt <= tmSsf.iCount
                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                    If ((tmAvail.iRecType >= 2) And (tmAvail.iRecType <= 9)) Then
                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llEvtTime
                        
                        If llLinkAirTime = llEvtTime Then
                            If (tmAvail.iAvInfo And &H1F <= 0) Or (tmAvail.iLen <= 0) Then
                                gTestAirVefValidDay = False         'found a matching avail and its zero units and/or second
                            End If
                            Exit Do
                        End If
                      
                    End If
                    ilEvt = ilEvt + 1
                Loop
                
            Else            'no testing for zero units, include spot
                gTestAirVefValidDay = True      'found an SSF record equal or previous to the log date
                Exit Do
            End If
            
        End If
        ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
        ilRet = gSSFGetPrevious(hlSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    Exit Function
End Function
'************************************************************
'*
'*      Procedure Name:gBuildODFSpotDay
'*
'*             Created:5/18/93       By:D. LeVine
'*            Modified:12/11/97      By:D. Hosaka
'*                 Place avail comment into spot record
'*            1/26/99 Include comment or other events
'*              when it is not encompassd by pgm
'*
'*            Comments:Build a day into ODF
'             1-10-01 Update page skip records with the
'               named avail  ID
'*      9-23-04 if not using affiliate system, logs showing
'               time zone are shown.
'*************************************************************
Function gBuildODFSpotDay(slFor As String, ilSSFType As Integer, sLCP As String, ilCallCode As Integer, slSDate As String, slEDate As String, slStartTime As String, slEndTime As String, ilEvtType() As Integer, ilSimVefCode As Integer, slInLogType As String, hlLst As Integer, hlMcf As Integer, ilGenLST As Integer, ilExportType As Integer, ilODFVefCode As Integer, slAdjLocalOrFeed As String, ilODFGameNo As Integer, llGsfCode As Long, ilLSTForLogVeh As Integer, Optional blTestMerge As Boolean = False) As Integer
'
'   ilRet = gBuildODFSpotDay (slFor, slType, slCp, ilCallCode, slSDate, slEDate, slStartTime, slEndTime, ilEvtType())
'
'   Where:
'       slFor (I)- "L"=Log; "C"=Commercial; "D"=Delivery (not valid for delivery with tmDlf.sCmmlSched = "N" and tmDlf.iMnfSubFeed = 0 test prior to slFor = "D" test)
'       slType (I)- "O"=On air; "A"=Alternate
'       slCP (I)- "C"=Current only; "P"=Pending only; "B"=Both
'       ilCallCode (I)-Vehicle code number(slFor = L or C) or feed code (slFor = D)
'       slSDate (I)- Start Date that events are to be obtained
'       slEDate (I)- Start Date that events are to be obtained
'       slStartTime (I)- Start Time (included)
'       slEndTime (I)- End time (not included)
'       ilEvtType (I)- Array of which events are to be included (True or False)
'                       Index description
'                         0   Library
'                         1   Program event
'                         2   Contract avail
'                         3   Open BB
'                         4   Floating BB
'                         5   Close BB
'                         6   Cmml promo
'                         7   Feed avail
'                         8   PSA avail
'                         9   Promo avail
'                         10  Page eject
'                         11  Line space 1
'                         12  Line space 2
'                         13  Line space 3
'                         14  Other event types
'       ilSimVefCode(I)- Simulcast vehicle code otherwisw 0 (if defined, generate for ilCallCode but store
'                        this vehicle code as the generation vehicle)
'       slInLogType(I)- P=Preliminary; F=Final; R=Reprint, A=Alert; I=Internal
'       hlLst (I)- LST File handle
'       hlMcf(I)- MCF File handle
'       ilGenLST(I)- True=Generate LST; False-Don't generate LST(only CP Return requested)
'       ilExportType(I)- 0=Manual; 1=Web; 2=Marketron.  This field used to know if alert should be created
'       ilUpdateLST(I)- True=Update LST (don't delete and create); False= Delete and Create because major change
'
'       slAdjLocalOrFeed(I)- L=Adjust Date via Local Time; F= Adjust Date by Feed time
'                            W=Same as L plus use Ft1CefCode for ChfCode (This is for Export Airwave)
'       ilLSTForLogVeh(I)- 0=No, 1=First time and get LST, 2=Second, third etc, 3=Clear unused LST
'       blTestMerge(I)- Test vehicle option with Log vehicles
'
'       Odf (O)- Event and spot records
'
'
    Dim ilRet As Integer
    Dim ilLogDate0 As Integer
    Dim ilLogDate1 As Integer
    Dim ilSsfDate0 As Integer
    Dim ilSsfDate1 As Integer
    Dim ilEvt As Integer
    Dim ilDay As Integer
    Dim slDay As String
    Dim slStr As String
    Dim ilSpot As Integer
    Dim ilLoop As Integer
    Dim ilVpf As Integer
    Dim ilIndex As Integer
    Dim slDate As String
    Dim llDate As Long
    Dim llSDate As Long
    Dim llEDate As Long
    Dim ilFound As Integer
    Dim ilTerminated As Integer
    Dim ilStartTime0 As Integer
    Dim ilStartTime1 As Integer
    Dim ilEndTime0 As Integer
    Dim ilEndTime1 As Integer
    Dim ilSeqNo As Integer
    Dim slTime As String
    Dim slLength As String
    Dim ilVpfIndex As Integer
    Dim ilVehCode As Integer
    Dim ilVeh As Integer
    Dim ilDlfDate0 As Integer
    Dim ilDlfDate1 As Integer
    Dim ilDlfFound As Integer
    Dim ilVlfDate0 As Integer
    Dim ilVlfDate1 As Integer
    Dim ilSIndex As Integer
    Dim slSsfDate As String
    'Spot summary
    Dim hlSsf As Integer        'Spot summary file handle
    Dim hlCTSsf As Integer        'Spot summary file handle
    Dim tlSsfSrchKey As SSFKEY0 'SSF key record image
    Dim ilSsfRecLen As Integer  'SSF record length
    Dim llStartTime As Long
    Dim llEndTime As Long
    Dim ilWithinTime As Integer
    Dim ilEvtFdIndex As Integer
    Dim llCefCode As Long
    Dim ilAirHour As Integer
    Dim ilLocalHour As Integer
    Dim ilFeedHour As Integer
    Dim ilType As Integer   '1=Prog; 2=Comment; 3=Avail; 4=Spot
    'Spot detail record information
    Dim llFirstEventTime As Long    '1/26/99

    Dim hlSdf As Integer        'Spot detail file handle
    Dim tlOdf As ODF
    Dim ilEvtRet As Integer
    Dim ilTestVefCode As Integer
    'Copy rotation record information
    Dim hlCrf As Integer        'Copy rotation file handle
    Dim ilCrfVefCode As Integer
    Dim ilPkgVefCode As Integer
    Dim ilLnVefCode As Integer
    Dim slLive As String
    Dim ilRdfCode As Integer
    'Copy inventory record information
    'Dim hmCif As Integer        'Copy inventory file handle
    'Dim ilCifRecLen As Integer  'CIF record length
    'Dim tlCif As CIF            'CIF record image
    'Copy
    Dim hlCnf As Integer
    'Short Title record information
    Dim hlSif As Integer        'Short Title file handle
    'Dim tlSifSrchKey As LONGKEY0 'SIF key record image
    'Dim ilSifRecLen As Integer  'SIF record length
    'Dim tlSif As SIF            'SIF record image
    'Dim llSifCode As Long
    'Air Copy
    'Advertiser record information
    Dim tlAdfSrchKey As INTKEY0 'ADF key record image
    Dim ilAdfRecLen As Integer  'ADF record length
    Dim tlAdf As ADF            'ADF record image
    Dim hlTzf As Integer
    Dim llCifCode As Long
    Dim slZone As String
    'Dim ilWkNo As Integer   'Week number from 1/5/70
    Dim llWkDateSet As Long
    Dim ilOther As Integer
    Dim ilZone As Integer
    Dim ilNoZones As Integer
    Dim ilFirstEvtShown As Integer
    Dim ilUpper As Integer
    Dim llUpper As Long
    Dim ilZoneFd As Integer
    Dim ilTZone As Integer
    Dim ilOtherGen As Integer
    Dim llAirTime As Long
    Dim llTime As Long
    Dim llTermDate As Long
    Dim ilPE As Integer
    Dim ilAvEvt As Integer
    Dim ilUnits As Integer
    Dim slUnits As String
    Dim slSpotLen As String
    Dim ilSec As Integer
    Dim ilAvVefCode As Integer
    Dim ilAvVpfIndex As Integer
    Dim ilNoDefZones As Integer
    Dim ilCopyReplaced As Integer
    Dim ilReplaceRotNo As Integer
    Dim llReplaceCopyCode As Long
    Dim ilRotNo As Integer
    Dim ilLPLocalAdj As Integer  'Largest Positive Local Adj (check day before start date)
    Dim ilLNLocalAdj As Integer  'Largest Negative Local Adj (check day after end date)
    Dim llLocalTime As Long
    Dim llFeedTime As Long
    Dim ilGenDate As Integer    'True = Generate odf for date
    Dim llMoDate As Long
    Dim hlChf As Integer
    Dim hlODF As Integer
    Dim hlVef As Integer
    Dim hlVlf As Integer
    Dim hlDlf As Integer
    Dim hlVsf As Integer
    Dim tlSdf As SDF
    Dim slAvailComm As String
    Dim ilLCFType As Integer
    Dim hlCef As Integer
    Dim tlCef As CEF
    Dim ilCefRecLen As Integer
    Dim tlCefSrchKey As LONGKEY0
    Dim ilCxfRecLen As Integer      'CXF record length
    Dim tlCxfSrchKey As LONGKEY0    'CXF key record image
    Dim tlCxf As CXF
    Dim ilBBPass As Integer
    Dim llBBTime As Long
    Dim ilAddSpot As Integer
    Dim ilBB As Integer
    Dim ilFdOpen As Integer
    Dim ilFdClose As Integer
    Dim llCrfCode As Long
    Dim ilLastWasSplit As Integer
    Dim ilCombineVefCode As Integer
    Dim ilLst As Integer
    Dim ilBBLen As Integer
    Dim ilMove As Integer
    Dim llLastLogDate As Long
    Dim slLSTCreateMsg As String
    Dim ilVff As Integer
    Dim blLstFound As Boolean
    ReDim llPrevOpenFdBBSpots(0 To 0) As Long
    ReDim llPrevCloseFdBBSpots(0 To 0) As Long
    Dim slSplitCopyFlag As String * 1       'blank if not split copy, else S
    Dim llCheckedGsfCode As Long


    '8/13/14: Generate spots for last date+1 into lst with Fed as *
    Dim llOdf As Long
    ReDim tmDate1ODF(0 To 0) As ODF
    Dim ilSortSeq As Integer

    Dim blBypassZeroUnits As Boolean
    
    'If ilSSFType = 0 Then
    '    slType = "O"
    'Else
    '    slType = "A"
    'End If
    If (ilLSTForLogVeh = 3) Then
        ReDim tmLstCode(0 To UBound(tmCombineLstInfo)) As LST
        For ilLst = LBound(tmCombineLstInfo) To UBound(tmCombineLstInfo) Step 1
            tmLstCode(ilLst) = tmCombineLstInfo(ilLst).tLst
        Next ilLst
        mDeleteUnusedLST slInLogType, ilExportType
        Erase tmCombineLstInfo
        Erase tmLstCode
        Erase llPrevOpenFdBBSpots
        Erase llPrevCloseFdBBSpots
        gBuildODFSpotDay = True
        Exit Function
    End If
    
    ilLCFType = ilSSFType
    llStartTime = CLng(gTimeToCurrency(slStartTime, False))
    llEndTime = CLng(gTimeToCurrency(slEndTime, True)) - 1
    ilSsfRecLen = Len(tmSsf)  'Get and save SSF record length
    hlSsf = CBtrvTable(ONEHANDLE)        'Create SSF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlSsf
        gBuildODFSpotDay = False
        Exit Function
    End If
    hlCTSsf = CBtrvTable(ONEHANDLE)        'Create SSF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlCTSsf, "", sgDBPath & "Ssf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        gBuildODFSpotDay = False
        Exit Function
    End If
    imSdfRecLen = Len(tmSdf)  'Get and save SDF record length
    hlSdf = CBtrvTable(ONEHANDLE)        'Create SDF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSdf, "", sgDBPath & "Sdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        gBuildODFSpotDay = False
        Exit Function
    End If
    imCHFRecLen = Len(tmChf)  'Get and save ADF record length
    hlChf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlChf, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        gBuildODFSpotDay = False
        Exit Function
    End If
    imOdfRecLen = Len(tmOdf)  'Get and save ODF record length
    hlODF = CBtrvTable(ONEHANDLE)        'Create ODF object handle- only maintained on Retrieval DB
    On Error GoTo 0
    ilRet = btrOpen(hlODF, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        gBuildODFSpotDay = False
        Exit Function
    End If
    imVefRecLen = Len(tmVef)  'Get and save VEF record length
    hlVef = CBtrvTable(ONEHANDLE)        'Create VLF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        gBuildODFSpotDay = False
        Exit Function
    End If
    imVlfRecLen = Len(tmVlf)  'Get and save VLF record length
    hlVlf = CBtrvTable(ONEHANDLE)        'Create VLF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlVlf, "", sgDBPath & "Vlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        btrDestroy hlVlf
        gBuildODFSpotDay = False
        Exit Function
    End If
    imDlfRecLen = Len(tmDlf)  'Get and save DLF record length
    hlDlf = CBtrvTable(ONEHANDLE)        'Create VLF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlDlf, "", sgDBPath & "Dlf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        btrDestroy hlVlf
        btrDestroy hlDlf
        gBuildODFSpotDay = False
        Exit Function
    End If
    tmCif.lCode = 0
    imCifRecLen = Len(tmCif)  'Get and save DLF record length
    hmCif = CBtrvTable(TWOHANDLES)        'Create VLF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        btrDestroy hlVlf
        btrDestroy hlDlf
        btrDestroy hmCif
        gBuildODFSpotDay = False
        Exit Function
    End If
    hlTzf = CBtrvTable(ONEHANDLE)        'Create TZF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlTzf, "", sgDBPath & "Tzf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        btrDestroy hlVlf
        btrDestroy hlDlf
        btrDestroy hmCif
        btrDestroy hlTzf
        gBuildODFSpotDay = False
        Exit Function
    End If
    hlCrf = CBtrvTable(TWOHANDLES)        'Create TZF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        btrDestroy hlVlf
        btrDestroy hlDlf
        btrDestroy hmCif
        btrDestroy hlTzf
        btrDestroy hlCrf
        gBuildODFSpotDay = False
        Exit Function
    End If
    'ilSifRecLen = Len(tlSif)  'Get and save SIF record length
    hlSif = CBtrvTable(ONEHANDLE)        'Create SIF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        btrDestroy hlVlf
        btrDestroy hlDlf
        btrDestroy hmCif
        btrDestroy hlTzf
        btrDestroy hlCrf
        btrDestroy hlSif
        gBuildODFSpotDay = False
        Exit Function
    End If
    ilAdfRecLen = Len(tlAdf)  'Get and save ADF record length
    hmAdf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmAdf, "", sgDBPath & "Adf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        btrDestroy hlVlf
        btrDestroy hlDlf
        btrDestroy hmCif
        btrDestroy hlTzf
        btrDestroy hlCrf
        btrDestroy hlSif
        btrDestroy hmAdf
        gBuildODFSpotDay = False
        Exit Function
    End If
    hmClf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        btrDestroy hlVlf
        btrDestroy hlDlf
        btrDestroy hmCif
        btrDestroy hlTzf
        btrDestroy hlCrf
        btrDestroy hlSif
        btrDestroy hmAdf
        btrDestroy hmClf
        gBuildODFSpotDay = False
        Exit Function
    End If
    hlVsf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    On Error GoTo 0
    ilRet = btrOpen(hlVsf, "", sgDBPath & "Vsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        btrDestroy hlVlf
        btrDestroy hlDlf
        btrDestroy hmCif
        btrDestroy hlTzf
        btrDestroy hlCrf
        btrDestroy hlSif
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hlVsf
        gBuildODFSpotDay = False
        Exit Function
    End If

    hmRdf = CBtrvTable(ONEHANDLE)        'Create Rdf object handle
    On Error GoTo 0
    ilRet = btrOpen(hmRdf, "", sgDBPath & "Rdf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmRdf
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        btrDestroy hlVlf
        btrDestroy hlDlf
        btrDestroy hmCif
        btrDestroy hlTzf
        btrDestroy hlCrf
        btrDestroy hlSif
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hlVsf
        gBuildODFSpotDay = False
        Exit Function
    End If

    hmCff = CBtrvTable(ONEHANDLE)        'Create Cff object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmCff
        btrDestroy hmRdf
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        btrDestroy hlVlf
        btrDestroy hlDlf
        btrDestroy hmCif
        btrDestroy hlTzf
        btrDestroy hlCrf
        btrDestroy hlSif
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hlVsf
        gBuildODFSpotDay = False
        Exit Function
    End If

    hmSmf = CBtrvTable(ONEHANDLE)        'Create SMF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        btrDestroy hmSmf
        btrDestroy hmCff
        btrDestroy hmRdf
        btrDestroy hlCTSsf
        btrDestroy hlSsf
        btrDestroy hlSdf
        btrDestroy hlChf
        btrDestroy hlODF
        btrDestroy hlVef
        btrDestroy hlVlf
        btrDestroy hlDlf
        btrDestroy hmCif
        btrDestroy hlTzf
        btrDestroy hlCrf
        btrDestroy hlSif
        btrDestroy hmAdf
        btrDestroy hmClf
        btrDestroy hlVsf
        gBuildODFSpotDay = False
        Exit Function
    End If
    ReDim ilVehicle(0 To 0) As Integer
    ReDim tlLLC(0 To 0) As LLC  'Image
    hlCnf = CBtrvTable(TWOHANDLES)        'Create CE object handle
    On Error GoTo 0
    ilRet = btrOpen(hlCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        '6/4/16: Replaced GoSub
        'GoSub CloseFiles
        mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
        gBuildODFSpotDay = False
        Exit Function
    End If
    hlCef = CBtrvTable(ONEHANDLE)        'Create CE object handle
    On Error GoTo 0
    ilRet = btrOpen(hlCef, "", sgDBPath & "Cef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        '6/4/16: Replaced GoSub
        'GoSub CloseFiles
        mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
        gBuildODFSpotDay = False
        Exit Function
    End If
    hmCxf = CBtrvTable(ONEHANDLE)        '2-15-01 Create CE object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCxf, "", sgDBPath & "Cxf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        '6/4/16: Replaced GoSub
        'GoSub CloseFiles
        mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
        gBuildODFSpotDay = False
        Exit Function
    End If
    hmRsf = CBtrvTable(TWOHANDLES)        'Create CE object handle
    On Error GoTo 0
    ilRet = btrOpen(hmRsf, "", sgDBPath & "Rsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        '6/4/16: Replaced GoSub
        'GoSub CloseFiles
        mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
        gBuildODFSpotDay = False
        Exit Function
    End If

    hmCvf = CBtrvTable(ONEHANDLE)        'Create CE object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        '6/4/16: Replaced GoSub
        'GoSub CloseFiles
        mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
        gBuildODFSpotDay = False
        Exit Function
    End If

    imCpfRecLen = Len(tmCpf)  'Get and save ADF record length
    hmCpf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmCpf, "", sgDBPath & "Cpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        '6/4/16: Replaced GoSub
        'GoSub CloseFiles
        mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
        gBuildODFSpotDay = False
        Exit Function
    End If

    If tgSpf.sSystemType = "R" Then         'radio system, need feed files
        hmFsf = CBtrvTable(TWOHANDLES)        'Create CE object handle
        On Error GoTo 0
        ilRet = btrOpen(hmFsf, "", sgDBPath & "Fsf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            '6/4/16: Replaced GoSub
            'GoSub CloseFiles
            mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
            gBuildODFSpotDay = False
            Exit Function
        End If
        imFsfRecLen = Len(tmFsf)

        hmPrf = CBtrvTable(TWOHANDLES)        'Create CE object handle
        On Error GoTo 0
        ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            '6/4/16: Replaced GoSub
            'GoSub CloseFiles
            mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
            gBuildODFSpotDay = False
            Exit Function
        End If
        imPrfRecLen = Len(tmPrf)

        hmFnf = CBtrvTable(TWOHANDLES)        'Create CE object handle
        On Error GoTo 0
        ilRet = btrOpen(hmFnf, "", sgDBPath & "Fnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            '6/4/16: Replaced GoSub
            'GoSub CloseFiles
            mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
            gBuildODFSpotDay = False
            Exit Function
        End If
        imFnfRecLen = Len(tmFnf)

    End If

    If (tgSpf.sGUseAffSys = "Y") And ilGenLST Then
        imLstExist = True
        imLstRecLen = Len(tmLst)  'Get and save ADF record length
        imMcfRecLen = Len(tmMcf)
        hmLst = hlLst   'CBtrvTable(ONEHANDLE)        'Create ADF object handle
        'On Error GoTo 0
        'ilRet = btrOpen(hmLst, "", sgDBPath & "Lst.Mkd", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'If ilRet <> BTRV_ERR_NONE Then
        '    MsgBox "LST.MKD Missing", vbOkOnly + vbExclamation, "Open Error"
        '    GoSub CloseFiles
        '    gBuildODFSpotDay = False
        '    Exit Function
        'End If
        imMnfRecLen = Len(tmMnf)  'Get and save ADF record length
        hmMnf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        On Error GoTo 0
        ilRet = btrOpen(hmMnf, "", sgDBPath & "Mnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            '6/4/16: Replaced GoSub
            'GoSub CloseFiles
            mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
            gBuildODFSpotDay = False
            Exit Function
        End If
        'commented out 1/18/99 - cff & smf added to main file open above
        'hmCff = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        'On Error GoTo 0
        'ilRet = btrOpen(hmCff, "", sgDBPath & "Cff.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'If ilRet <> BTRV_ERR_NONE Then
        '    GoSub CloseFiles
        '    gBuildODFSpotDay = False
        '    Exit Function
        'End If
        'hmSmf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        'On Error GoTo 0
        'ilRet = btrOpen(hmSmf, "", sgDBPath & "Smf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        'If ilRet <> BTRV_ERR_NONE Then
        '    GoSub CloseFiles
        '    gBuildODFSpotDay = False
        '    Exit Function
        'End If
        hmDrf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        On Error GoTo 0
        ilRet = btrOpen(hmDrf, "", sgDBPath & "Drf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            '6/4/16: Replaced GoSub
            'GoSub CloseFiles
            mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
            gBuildODFSpotDay = False
            Exit Function
        End If
        '8-4-01
        hmDpf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        On Error GoTo 0
        ilRet = btrOpen(hmDpf, "", sgDBPath & "Dpf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            '6/4/16: Replaced GoSub
            'GoSub CloseFiles
            mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
            gBuildODFSpotDay = False
            Exit Function
        End If
        ' setup global variable for Demo Plus file (to see if any exists)
        lgDpfNoRecs = btrRecords(hmDpf)
        If lgDpfNoRecs = 0 Then
            lgDpfNoRecs = -1
        End If
        hmDef = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        On Error GoTo 0
        ilRet = btrOpen(hmDef, "", sgDBPath & "Def.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            '6/4/16: Replaced GoSub
            'GoSub CloseFiles
            mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
            gBuildODFSpotDay = False
            Exit Function
        End If
        hmRaf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        On Error GoTo 0
        ilRet = btrOpen(hmRaf, "", sgDBPath & "Raf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            '6/4/16: Replaced GoSub
            'GoSub CloseFiles
            mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
            gBuildODFSpotDay = False
            Exit Function
        End If
    Else
        imLstExist = False
    End If
    smLogType = slInLogType
    If smLogType = "I" Then 'Internal
        smLogType = "R"
    ElseIf smLogType = "A" Then 'Alert
        smLogType = "R"
    End If
    llCheckedGsfCode = -1
    If slFor = "D" Then
        ReDim ilVehicle(0 To 0) As Integer
        ilRet = btrGetFirst(hlVef, tmVef, imVefRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While ilRet = BTRV_ERR_NONE
            If (tmVef.sType = "A") Then
                ilVpfIndex = -1
                'For ilLoop = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
                '    If tmVef.iCode = tgVpf(ilLoop).iVefKCode Then
                    ilLoop = gBinarySearchVpf(tmVef.iCode)
                    If ilLoop <> -1 Then
                        ilVpfIndex = ilLoop
                '        Exit For
                    End If
                'Next ilLoop
                If ilVpfIndex >= 0 Then
                    gBuildLinkArray hlVlf, tmVef, slSDate, igSVefCode()
                    For ilLoop = LBound(tgVpf(ilVpfIndex).iGMnfNCode) To UBound(tgVpf(ilVpfIndex).iGMnfNCode) Step 1
                        If tgVpf(ilVpfIndex).iGMnfNCode(ilLoop) = ilCallCode Then
                            'For ilIndex = LBound(tgVpf(ilVpfIndex).iGLink) To UBound(tgVpf(ilVpfIndex).iGLink) Step 1
                            '    If tgVpf(ilVpfIndex).iGLink(ilIndex) > 0 Then
                            For ilIndex = LBound(igSVefCode) To UBound(igSVefCode) - 1 Step 1
                                    ilFound = False
                                    For ilVeh = LBound(ilVehicle) To UBound(ilVehicle) - 1 Step 1
                                        If ilVehicle(ilVeh) = igSVefCode(ilIndex) Then  'tgVpf(ilVpfIndex).iGLink(ilIndex) Then
                                            ilFound = True
                                            Exit For
                                        End If
                                    Next ilVeh
                                    If Not ilFound Then
                                        ilVehicle(UBound(ilVehicle)) = igSVefCode(ilIndex)  'tgVpf(ilVpfIndex).iGLink(ilIndex)
                                        ReDim Preserve ilVehicle(LBound(ilVehicle) To UBound(ilVehicle) + 1) As Integer
                                    End If
                            '    End If
                            Next ilIndex
                        End If
                    Next ilLoop
                End If
            ElseIf (tmVef.sType = "C") Then
                ilVpfIndex = -1
                'For ilLoop = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
                '    If tmVef.iCode = tgVpf(ilLoop).iVefKCode Then
                    ilLoop = gBinarySearchVpf(tmVef.iCode)
                    If ilLoop <> -1 Then
                        ilVpfIndex = ilLoop
                '        Exit For
                    End If
                'Next ilLoop
                If ilVpfIndex >= 0 Then
                    For ilLoop = LBound(tgVpf(ilVpfIndex).iGMnfNCode) To UBound(tgVpf(ilVpfIndex).iGMnfNCode) Step 1
                        If tgVpf(ilVpfIndex).iGMnfNCode(ilLoop) = ilCallCode Then
                            ilFound = False
                            For ilVeh = LBound(ilVehicle) To UBound(ilVehicle) - 1 Step 1
                                If ilVehicle(ilVeh) = tmVef.iCode Then
                                    ilFound = True
                                    Exit For
                                End If
                            Next ilVeh
                            If Not ilFound Then
                                ilVehicle(UBound(ilVehicle)) = tmVef.iCode
                                ReDim Preserve ilVehicle(LBound(ilVehicle) To UBound(ilVehicle) + 1) As Integer
                            End If
                        End If
                    Next ilLoop
                End If
            End If
            ilRet = btrGetNext(hlVef, tmVef, imVefRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
            DoEvents
        Loop
    Else
        ReDim ilVehicle(0 To 1) As Integer
        ilVehicle(0) = ilCallCode
    End If
    tmCTSsf.iVefCode = 0    'Force read of ssf
    tmVef.iCode = 0
    llSDate = gDateValue(slSDate)
    llEDate = gDateValue(slEDate)
    lm010570 = gDateValue("1/5/70")
    'Don't need to look at llSDate - 1 as those spots have been moved forward from the previous week generation
    '(Generate for PST, and EST is + 3)
    tmOdf.iGenDate(0) = igGenDate(0)
    tmOdf.iGenDate(1) = igGenDate(1)
    '10-9-01
    tmOdf.lGenTime = lgGenTime
    'tmOdf.iGenTime(0) = igGenTime(0)
    'tmOdf.iGenTime(1) = igGenTime(1)
    llMoDate = 0
    ilLastWasSplit = False
    If ilODFVefCode <= 0 Then
        '11/4/09:  Re-add generation of LST for Log vehicles
        'ReDim tmCombineLstInfo(0 To 0) As COMBINELSTINFO
        If (ilLSTForLogVeh = 0) Or (tmVef.iVefCode <= 0) Then
            ReDim tmCombineLstInfo(0 To 0) As COMBINELSTINFO
        ElseIf (ilLSTForLogVeh = 1) And (tmVef.iVefCode > 0) Then
            ReDim tmCombineLstInfo(0 To 0) As COMBINELSTINFO
        End If
    End If
    blLstFound = False

    '7/11/14: Local avails between 12am-3am
    imLockVefCode = -1
    lmLockStartTime = -1
    
    '8/13/14: Generate spots for last date+1 into lst with Fed as *
    'For llDate = llSDate To llEDate + 1 Step 1
    bmCreatingLstDate1 = False
    llDate = llSDate
    Do While llDate <= llEDate + 1
        ReDim tmBBSdfInfo(0 To 0) As BBSDFINFO
        ReDim tlLLC(0 To 0) As LLC  'Image
        'ReDim tmLstCode(0 To 0) As LST
        'If (llGsfCode = 0) Or ((llDate = llSDate) And llGsfCode > 0) Then
        If (llGsfCode = 0) Or ((Not blLstFound) And llGsfCode > 0) Then
            ReDim tmLstCode(0 To 0) As LST
        End If
        ilWithinTime = False
        slDate = Format$(llDate, "m/d/yy")
        imDaySort = gWeekDayLong(llDate)
        If imDaySort <= 4 Then
            imDaySort = 1
        Else
            imDaySort = imDaySort + 1
        End If
        ilDay = gWeekDayStr(slDate)
        gPackDate slDate, ilLogDate0, ilLogDate1
        lmPrevAvailTime = -1
        For ilVeh = LBound(ilVehicle) To UBound(ilVehicle) - 1 Step 1
            ilSortSeq = 0                               '8-18-14
            smAlertStatus = "C"
            ilVehCode = ilVehicle(ilVeh)
            If ilVehCode <> tmVef.iCode Then
                tmVefSrchKey.iCode = ilVehCode
                ilRet = btrGetEqual(hlVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                If (ilRet <> BTRV_ERR_NONE) Then
                    '6/4/16: Replaced GoSub
                    'GoSub CloseFiles
                    mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
                    gBuildODFSpotDay = False
                    Exit Function
                End If
                ilVpfIndex = -1
                'For ilLoop = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
                '    If tmVef.iCode = tgVpf(ilLoop).iVefKCode Then
                    ilLoop = gBinarySearchVpf(tmVef.iCode)
                    If ilLoop <> -1 Then
                        ilVpfIndex = ilLoop
                '        Exit For
                    End If
                'Next ilLoop
                If ilVpfIndex = -1 Then
                    '6/4/16: Replaced GoSub
                    'GoSub CloseFiles
                    mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
                    gBuildODFSpotDay = False
                    Exit Function
                End If
                blBypassZeroUnits = False
                If (tmVef.sType = "A") Then
                    ilVff = gBinarySearchVff(tmVef.iCode)
                    If ilVff <> -1 Then
                        If tgVff(ilVff).sHonorZeroUnits = "Y" Then
                            blBypassZeroUnits = True
                        End If
                    End If
                End If
                If tmVef.iVefCode > 0 And blTestMerge Then
                    For ilVff = LBound(tgVff) To UBound(tgVff) Step 1
                        If tmVef.iCode = tgVff(ilVff).iVefCode Then
                            If tgVff(ilVff).sMergeTraffic = "S" Then
                                tmVef.iVefCode = 0
                            End If
                            Exit For
                        End If
                    Next ilVff
                End If
                gUnpackDateLong tgVpf(ilVpfIndex).iLLD(0), tgVpf(ilVpfIndex).iLLD(1), llLastLogDate
            End If
            
            '12/29/08:  If cart # not defined, then use the Reel # if Wegener or OLA
            imWegenerOLA = False
            'If (tgVpf(ilVpfIndex).sWegenerExport = "Y") Or (tgVpf(ilVpfIndex).sOLAExport = "Y") Then
            If (tgVpf(ilVpfIndex).sOLAExport = "Y") Then
                imWegenerOLA = True
            End If
            ilCombineVefCode = tmVef.iCombineVefCode
            tmOdf.lHd1CefCode = 0   'tgVpf(ilVpfIndex).lLgHd1CefCode
            tmOdf.iAlternateVefCode = 0 '1-17-14 was tmosdf.llgnmcefcode, tgVpf(ilVpfIndex).lLgNmCefCode
            tmOdf.lFt1CefCode = 0   'tgVpf(ilVpfIndex).lLgFt1CefCode
            tmOdf.lFt2CefCode = 0   'tgVpf(ilVpfIndex).lLgFt2CefCode
            ilLPLocalAdj = 0
            ilLNLocalAdj = 0
            ilNoDefZones = 0
            ReDim tmZoneInfo(0 To 0) As ZONEINFO
            For ilLoop = LBound(tgVpf(ilVpfIndex).sGZone) To UBound(tgVpf(ilVpfIndex).sGZone) Step 1
                If Trim$(tgVpf(ilVpfIndex).sGZone(ilLoop)) <> "" Then
                    tmZoneInfo(ilNoDefZones).sZone = UCase$(Trim$(tgVpf(ilVpfIndex).sGZone(ilLoop)))
                    tmZoneInfo(ilNoDefZones).lGLocalAdj = 3600 * tgVpf(ilVpfIndex).iGLocalAdj(ilLoop)
                    tmZoneInfo(ilNoDefZones).lGFeedAdj = 3600 * tgVpf(ilVpfIndex).iGFeedAdj(ilLoop)
                    tmZoneInfo(ilNoDefZones).sFed = UCase$(Trim$(tgVpf(ilVpfIndex).sGFed(ilLoop)))
                    tmZoneInfo(ilNoDefZones).lcpfCode = 0
                    If (tgSpf.sGUseAffSys = "Y") Then
                        If slAdjLocalOrFeed <> "F" Then
                            If tgVpf(ilVpfIndex).iGLocalAdj(ilLoop) < 0 Then
                                If tgVpf(ilVpfIndex).iGLocalAdj(ilLoop) < ilLNLocalAdj Then
                                    ilLNLocalAdj = tgVpf(ilVpfIndex).iGLocalAdj(ilLoop)
                                End If
                            ElseIf tgVpf(ilVpfIndex).iGLocalAdj(ilLoop) > 0 Then
                                If tgVpf(ilVpfIndex).iGLocalAdj(ilLoop) > ilLPLocalAdj Then
                                    ilLPLocalAdj = tgVpf(ilVpfIndex).iGLocalAdj(ilLoop)
                                End If
                            End If
                        Else
                            If tgVpf(ilVpfIndex).iGFeedAdj(ilLoop) < 0 Then
                                If tgVpf(ilVpfIndex).iGFeedAdj(ilLoop) < ilLNLocalAdj Then
                                    ilLNLocalAdj = tgVpf(ilVpfIndex).iGFeedAdj(ilLoop)
                                End If
                            ElseIf tgVpf(ilVpfIndex).iGFeedAdj(ilLoop) > 0 Then
                                If tgVpf(ilVpfIndex).iGFeedAdj(ilLoop) > ilLPLocalAdj Then
                                    ilLPLocalAdj = tgVpf(ilVpfIndex).iGFeedAdj(ilLoop)
                                End If
                            End If
                        End If
                    End If
                    ilNoDefZones = ilNoDefZones + 1
                    ReDim Preserve tmZoneInfo(0 To ilNoDefZones) As ZONEINFO
                End If
            Next ilLoop
            If ilNoDefZones = 0 Then
                tmZoneInfo(ilNoDefZones).sZone = ""
                tmZoneInfo(ilNoDefZones).lGLocalAdj = 0
                tmZoneInfo(ilNoDefZones).lGFeedAdj = 0
                tmZoneInfo(ilNoDefZones).lcpfCode = 0
                tmZoneInfo(ilNoDefZones).sFed = "*"
                ilNoDefZones = ilNoDefZones + 1
                ReDim Preserve tmZoneInfo(0 To ilNoDefZones) As ZONEINFO
            End If
            ilGenDate = True
            If llDate < llSDate Then
                If (tgSpf.sGUseAffSys = "Y") And (ilLPLocalAdj > 0) Then
                    '8/12/08: Convert to Local Time, not air time
                    'llLPTime = 86400 - ilLPLocalAdj * CLng(3600)
                    'llLNTime = 86400
                    lmLPTime = 0
                    lmLNTime = 3600 * CLng(ilLNLocalAdj)
                Else
                    ilGenDate = False
                End If
            ElseIf llDate > llEDate Then
                If (tgSpf.sGUseAffSys = "Y") And (ilLNLocalAdj < 0) Then
                    '8/12/08: Convert to Local Time, not air time
                    'llLPTime = 0
                    'llLNTime = -3600 * CLng(ilLNLocalAdj)
                    lmLPTime = 86400 + ilLPLocalAdj * CLng(3600)
                    lmLNTime = 86400
                Else
                    ilGenDate = False
                End If
            Else
                lmLPTime = 0
                lmLNTime = 86400
            End If
            If ilGenDate Then
                imCreateNewLST = False
                '8/10/10:  Handle games that cross midnight
                'ReDim tmLstCode(0 To 0) As LST
                'If (llGsfCode = 0) Or ((llDate = llSDate) And llGsfCode > 0) Then
                If (llGsfCode = 0) Or ((Not blLstFound) And llGsfCode > 0) Then
                    ReDim tmLstCode(0 To 0) As LST
                End If
                ReDim llPrevOpenFdBBSpots(0 To 0) As Long
                ReDim llPrevCloseFdBBSpots(0 To 0) As Long
                slLSTCreateMsg = ""
                If imLstExist And ilGenLST Then

                    'Determine if LST should be deleted, then saved or only updated
                    If llDate < llEDate + 1 Then
                        '1/11/08:  Match how odf is created for combination games (remove game test). i.e. produce lst as a combination of game vehicles
                        'If (ilODFVefCode <= 0) Or (tmVef.sType = "G") Then
                        If (ilODFVefCode <= 0) Then
                            If ilSimVefCode <= 0 Then
                                '11/4/09:  Re-ass the generation of LST by Log vehicle
                                'If gAlertFound("L", "S", 0, ilVehCode, slDate) Then
                                '    ilCreateNewLST = True
                                'Else
                                '    If gAlertFound("L", "C", 0, ilVehCode, slDate) Then
                                '        ilCreateNewLST = True
                                '    End If
                                'End If
                                If (ilLSTForLogVeh > 0) And (tmVef.iVefCode > 0) Then
                                    If gAlertFound("L", "S", 0, tmVef.iVefCode, slDate) Then
                                        imCreateNewLST = True
                                    Else
                                        If gAlertFound("L", "C", 0, tmVef.iVefCode, slDate) Then
                                            imCreateNewLST = True
                                        End If
                                    End If
                                Else
                                    If gAlertFound("L", "S", 0, ilVehCode, slDate) Then
                                        imCreateNewLST = True
                                    Else
                                        If gAlertFound("L", "C", 0, ilVehCode, slDate) Then
                                            imCreateNewLST = True
                                        End If
                                    End If
                                End If
                            Else
                                For ilLoop = 0 To UBound(tgLSTUpdateInfo) - 1 Step 1
                                    If ilVehCode = tgLSTUpdateInfo(ilLoop).iVefCode Then
                                        If (llDate >= tgLSTUpdateInfo(ilLoop).lSDate) And (llDate <= tgLSTUpdateInfo(ilLoop).lEDate) Then
                                            If tgLSTUpdateInfo(ilLoop).iType = 0 Then
                                                imCreateNewLST = True
                                                Exit For
                                            ElseIf tgLSTUpdateInfo(ilLoop).iType = 1 Then
                                                imCreateNewLST = True
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next ilLoop
                            End If
                        Else
                            If gAlertFound("L", "S", 0, ilODFVefCode, slDate) Then
                                imCreateNewLST = True
                            Else
                                If gAlertFound("L", "C", 0, ilODFVefCode, slDate) Then
                                    imCreateNewLST = True
                                End If
                            End If
                        End If
                    End If
                    'Added when changed to generate Affiliate with each Conventional Vehicle instead of
                    'the Log Vehicle.  11/20/03.  Use ilVehCode instead of tmVef.iVefCode
                    'If tmVef.iVefCode > 0 Then
                    '    mClearLst tmVef.iVefCode, llDate, llLPTime, llLNTime
                    'Else
                    '    mClearLst ilVehCode, llDate, llLPTime, llLNTime
                    'End If
                    '1/11/08:  Match how odf is created for combination games (remove game test). i.e. produce lst as a combination of game vehicles
                    'If (ilODFVefCode <= 0) Or (tmVef.sType = "G") Then
                    If (ilODFVefCode <= 0) Then
                        If ilSimVefCode <= 0 Then
                            '11/4/09: re-add generation of lst by log vehicle
                            'mClearLst ilVehCode, llDate, llLPTime, llLNTime, ilCreateNewLST, llGsfCode
                            If (ilLSTForLogVeh > 0) And (tmVef.iVefCode > 0) Then
                                If ilLSTForLogVeh = 1 Then
                                    'mClearLst tmVef.iVefCode, llDate, llLPTime, llLNTime, ilCreateNewLST, llGsfCode, slLSTCreateMsg, blLstFound
                                    '8/13/14: Generate spots for last date+1 into lst with Fed as *
                                    If Not bmCreatingLstDate1 Then
                                        mClearLst tmVef.iVefCode, llDate, lmLPTime, lmLNTime, imCreateNewLST, llGsfCode, slLSTCreateMsg, blLstFound
                                    Else
                                        mClearLst tmVef.iVefCode, llDate, 0, -3600 * CLng(ilLNLocalAdj), imCreateNewLST, llGsfCode, slLSTCreateMsg, blLstFound
                                    End If
                                Else
                                    ReDim tmLstCode(0 To 0) As LST
                                    For ilLst = LBound(tmCombineLstInfo) To UBound(tmCombineLstInfo) Step 1
                                        If tmCombineLstInfo(ilLst).lDate = llDate Then
                                            '1/29/10:  Spots set in CreateLst as to retain
                                            'If tmCombineLstInfo(ilLst).tLst.lSdfCode > 0 Then
                                            '    tmCombineLstInfo(ilLst).tLst.lCode = 0
                                            'End If
                                            tmLstCode(UBound(tmLstCode)) = tmCombineLstInfo(ilLst).tLst
                                            ReDim Preserve tmLstCode(0 To UBound(tmLstCode) + 1) As LST
                                        End If
                                    Next ilLst
                                    '1/29/10:  remove dates moved to lstspots
                                    ilLst = LBound(tmCombineLstInfo)
                                    ilUpper = UBound(tmCombineLstInfo)
                                    Do While ilLst < ilUpper
                                        If tmCombineLstInfo(ilLst).lDate = llDate Then
                                            For ilMove = ilLst To UBound(tmCombineLstInfo) - 1 Step 1
                                                tmCombineLstInfo(ilMove) = tmCombineLstInfo(ilMove + 1)
                                            Next ilMove
                                            ReDim Preserve tmCombineLstInfo(0 To UBound(tmCombineLstInfo) - 1) As COMBINELSTINFO
                                            ilUpper = UBound(tmCombineLstInfo)
                                        Else
                                            ilLst = ilLst + 1
                                        End If
                                    Loop
                                End If
                            Else
                                ''8/10/10:  Handle games that cross midnight
                                ''mClearLst ilVehCode, llDate, llLPTime, llLNTime, ilCreateNewLST, llGsfCode
                                ''If (llGsfCode = 0) Or ((llDate = llSDate) And llGsfCode > 0) Then
                                '5/9/14: In mClearLST if it is a game, the spots on the next day are obtained.
                                '        What happens here is that the spots created on the next day are removed if generating Final logs
                                '        as blLstFound is false
                                'If (llGsfCode = 0) Or ((Not blLstFound) And llGsfCode > 0) Then
                                If (llGsfCode = 0) Or ((Not blLstFound) And (llGsfCode > 0) And (llCheckedGsfCode <> llGsfCode)) Then
                                    '8/13/14: Generate spots for last date+1 into lst with Fed as *
                                    'mClearLst ilVehCode, llDate, llLPTime, llLNTime, ilCreateNewLST, llGsfCode, slLSTCreateMsg, blLstFound
                                    If Not bmCreatingLstDate1 Then
                                        mClearLst ilVehCode, llDate, lmLPTime, lmLNTime, imCreateNewLST, llGsfCode, slLSTCreateMsg, blLstFound
                                    Else
                                        mClearLst ilVehCode, llDate, 0, -3600 * CLng(ilLNLocalAdj), imCreateNewLST, llGsfCode, slLSTCreateMsg, blLstFound
                                    End If
                                    llCheckedGsfCode = llGsfCode
                                End If
                            End If
                        Else
                            '8/13/14: Generate spots for last date+1 into lst with Fed as *
                            'mClearLst ilSimVefCode, llDate, llLPTime, llLNTime, ilCreateNewLST, llGsfCode, slLSTCreateMsg, blLstFound
                            If Not bmCreatingLstDate1 Then
                                mClearLst ilSimVefCode, llDate, lmLPTime, lmLNTime, imCreateNewLST, llGsfCode, slLSTCreateMsg, blLstFound
                            Else
                                mClearLst ilSimVefCode, llDate, 0, -3600 * CLng(ilLNLocalAdj), imCreateNewLST, llGsfCode, slLSTCreateMsg, blLstFound
                            End If
                        End If
                    Else
                        'LST obtained for first part of combine
                        'mClearLst ilODFVefCode, llDate, llLPTime, llLNTime, ilCreateNewLST
                        ReDim tmLstCode(0 To 0) As LST
                        For ilLst = LBound(tmCombineLstInfo) To UBound(tmCombineLstInfo) Step 1
                            If tmCombineLstInfo(ilLst).lDate = llDate Then
                                tmLstCode(UBound(tmLstCode)) = tmCombineLstInfo(ilLst).tLst
                                ReDim Preserve tmLstCode(0 To UBound(tmLstCode) + 1) As LST
                            End If
                        Next ilLst
                    End If
                    'Add of cgange:  Removed if statement and added If statement with ilSimVefCode
                End If
                lmEvtCefCode = 0
                imEvtCefSort = 0
                ilDlfFound = False
                'If (slFor = "D") Or (((tmVef.sType = "A") Or ((tmVef.sType = "C") And (tgVpf(ilVpfIndex).iGMnfNCode(1) <> 0)))) Then
                If (slFor = "D") Or (((tmVef.sType = "A") Or ((tmVef.sType = "C") And (tgVpf(ilVpfIndex).iGMnfNCode(0) <> 0)))) Then
                    'Obtain delivery records for date
                    If (ilDay >= 0) And (ilDay <= 4) Then
                        slDay = "0"
                    ElseIf ilDay = 5 Then
                        slDay = "6"
                    Else
                        slDay = "7"
                    End If
                    'Obtain the start date of DLF
                    tmDlfSrchKey.iVefCode = ilVehCode
                    tmDlfSrchKey.sAirDay = slDay
                    tmDlfSrchKey.iStartDate(0) = ilLogDate0
                    tmDlfSrchKey.iStartDate(1) = ilLogDate1
                    tmDlfSrchKey.iAirTime(0) = 0
                    tmDlfSrchKey.iAirTime(1) = 6144 '24*256
                    ilRet = btrGetLessOrEqual(hlDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                    If (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) Then
                        ilDlfDate0 = tmDlf.iStartDate(0)
                        ilDlfDate1 = tmDlf.iStartDate(1)
                        ilDlfFound = True
                    Else
                        ilDlfDate0 = 0
                        ilDlfDate1 = 0
                    End If
                    'Obtain the start date of VLF
                    If tmVef.sType = "A" Then
                        ilVlfDate0 = 0
                        ilVlfDate1 = 0
                        tmVlfSrchKey1.iAirCode = ilVehCode
                        tmVlfSrchKey1.iAirDay = Val(slDay)
                        tmVlfSrchKey1.iEffDate(0) = ilLogDate0
                        tmVlfSrchKey1.iEffDate(1) = ilLogDate1
                        tmVlfSrchKey1.iAirTime(0) = 0
                        tmVlfSrchKey1.iAirTime(1) = 6144    '24*256
                        tmVlfSrchKey1.iAirPosNo = 32000
                        tmVlfSrchKey1.iAirSeq = 32000
                        ilRet = btrGetLessOrEqual(hlVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
                        Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode)
                            If (tmVlf.iAirDay = Val(slDay)) Then
                                ilTerminated = False
                                If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                    If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                        ilTerminated = True
                                    End If
                                End If
                                If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                'If (tmVlf.sStatus <> "P") Then
'                                    ilVlfDate0 = tmVlf.iEffDate(0)
'                                    ilVlfDate1 = tmVlf.iEffDate(1)
                                    gUnpackDateLong tmVlf.iTermDate(0), tmVlf.iTermDate(1), llTermDate
                                    If (llTermDate = 0) Or (llDate <= llTermDate) Then
                                        ilVlfDate0 = tmVlf.iEffDate(0)
                                        ilVlfDate1 = tmVlf.iEffDate(1)
                                    End If
                                    Exit Do
                                End If
                            End If
                            ilRet = btrGetPrevious(hlVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                        Loop
                    End If
                End If
                DoEvents
                'gObtainVlf hlVlf, ilVehCode, llDate, tlVlf0(), tlVlf5(), tlVlf6()
                ilDay = gWeekDayStr(slDate)
                gPackDate slDate, ilLogDate0, ilLogDate1
                ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                ilSsfDate0 = ilLogDate0
                ilSsfDate1 = ilLogDate1
                tlSsfSrchKey.iType = ilSSFType
                tlSsfSrchKey.iVefCode = ilVehCode
                tlSsfSrchKey.iDate(0) = ilSsfDate0
                tlSsfSrchKey.iDate(1) = ilSsfDate1
                tlSsfSrchKey.iStartTime(0) = 0
                tlSsfSrchKey.iStartTime(1) = 0
                ilRet = gSSFGetEqual(hlSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                If (ilRet <> BTRV_ERR_NONE) Or (tmSsf.iType <> ilSSFType) Or (tmSsf.iVefCode <> ilVehCode) Or (tmSsf.iDate(0) <> ilSsfDate0) Or (tmSsf.iDate(1) <> ilSsfDate1) Then
                    'If airing- then use first Ssf prior to date defined
                    If tmVef.sType = "A" Then
                        'Check that airing vehicle is running in the log date
                        ilFound = False
                        'First check for TFN, if it does not exist, then check for SSF image
                        gUnpackDate ilLogDate0, ilLogDate1, slSsfDate
                        Select Case gWeekDayStr(slSsfDate)
                            Case 0
                                slSsfDate = "TFNMO"
                            Case 1
                                slSsfDate = "TFNTU"
                            Case 2
                                slSsfDate = "TFNWE"
                            Case 3
                                slSsfDate = "TFNTH"
                            Case 4
                                slSsfDate = "TFNFR"
                            Case 5
                                slSsfDate = "TFNSA"
                            Case 6
                                slSsfDate = "TFNSU"
                        End Select
                        ReDim tlTFNLLC(0 To 0) As LLC  'Image
                        If (ilEvtType(0) = True) Or (ilEvtType(10) = True) Or (ilEvtType(11) = True) Or (ilEvtType(12) = True) Or (ilEvtType(13) = True) Or (ilEvtType(14) = True) Then
                            ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, slSsfDate, slStartTime, slEndTime, ilEvtType(), tlTFNLLC())
                        Else
                            If ilEvtType(1) = True Then
                                ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, slSsfDate, slStartTime, slEndTime, ilEvtType(), tlTFNLLC())
                            End If
                        End If
                        If UBound(tlTFNLLC) <= 0 Then
                            ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                            tlSsfSrchKey.iType = ilSSFType
                            tlSsfSrchKey.iVefCode = ilVehCode
                            tlSsfSrchKey.iDate(0) = ilLogDate0
                            tlSsfSrchKey.iDate(1) = ilLogDate1
                            tlSsfSrchKey.iStartTime(0) = 0
                            tlSsfSrchKey.iStartTime(1) = 0  '6144   '24*256
                            ilRet = gSSFGetGreaterOrEqual(hlSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = ilVehCode)
                                gUnpackDate tmSsf.iDate(0), tmSsf.iDate(1), slSsfDate
                                If (ilDay = gWeekDayStr(slSsfDate)) And (tmSsf.iStartTime(0) = 0) And (tmSsf.iStartTime(1) = 0) Then
                                    ilFound = True
                                    Exit Do
                                End If
                                ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                                ilRet = gSSFGetNext(hlSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                        Else
                            ilFound = True
                        End If
                        If ilFound Then
                            ilSsfDate0 = 0
                            ilSsfDate1 = 0
                            ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                            tlSsfSrchKey.iType = ilSSFType
                            tlSsfSrchKey.iVefCode = ilVehCode
                            tlSsfSrchKey.iDate(0) = ilLogDate0
                            tlSsfSrchKey.iDate(1) = ilLogDate1
                            tlSsfSrchKey.iStartTime(0) = 0
                            tlSsfSrchKey.iStartTime(1) = 6144   '24*256
                            ilRet = gSSFGetLessOrEqual(hlSsf, tmSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                            Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = ilVehCode)
                                gUnpackDate tmSsf.iDate(0), tmSsf.iDate(1), slSsfDate
                                If (ilDay = gWeekDayStr(slSsfDate)) And (tmSsf.iStartTime(0) = 0) And (tmSsf.iStartTime(1) = 0) Then
                                    ilSsfDate0 = tmSsf.iDate(0)
                                    ilSsfDate1 = tmSsf.iDate(1)
                                    Exit Do
                                End If
                                ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                                ilRet = gSSFGetPrevious(hlSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                            Loop
                        Else
                            ilSsfDate0 = ilLogDate0
                            ilSsfDate1 = ilLogDate1
                        End If
                    End If
                End If
                DoEvents
                If (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = ilVehCode) Then
                    '3/8/12: Output error message move here from mClearLST so that the message only is generated
                    '        if day exist
                    If (slLSTCreateMsg <> "") And (bgReprintLogType) Then
                        'Check last log date
                        If llDate <= llLastLogDate Then
                            gLogMsg slLSTCreateMsg, "TrafficErrors.Txt", False
                        End If
                    End If
                    gUnpackDate ilSsfDate0, ilSsfDate1, slSsfDate
                    If (ilEvtType(0) = True) Or (ilEvtType(10) = True) Or (ilEvtType(11) = True) Or (ilEvtType(12) = True) Or (ilEvtType(13) = True) Or (ilEvtType(14) = True) Then
                        ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, slSsfDate, slStartTime, slEndTime, ilEvtType(), tlLLC())
                    Else
                        If ilEvtType(1) = True Then
                            ilEvtRet = gBuildEventDay(ilLCFType, sLCP, ilVehCode, slSsfDate, slStartTime, slEndTime, ilEvtType(), tlLLC())
                        End If
                    End If
                    ReDim tmPageEject(0 To 0) As PAGEEJECT
                    llFirstEventTime = -1       '1/26/99   do one time only
                    For ilLoop = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                        'Get Page Ejects
                        If tlLLC(ilLoop).iEtfCode = 10 Then
                            tmPageEject(UBound(tmPageEject)).lTime = gTimeToLong(tlLLC(ilLoop).sStartTime, False)
                            tmPageEject(UBound(tmPageEject)).ianfCode = Val(tlLLC(ilLoop).sName)
                            ReDim Preserve tmPageEject(0 To UBound(tmPageEject) + 1) As PAGEEJECT
                        End If
                        lmEvtTime = gTimeToLong(tlLLC(ilLoop).sStartTime, False)        '1/26/99
                        If tlLLC(ilLoop).iEtfCode > 9 And llFirstEventTime = -1 And lmEvtTime >= llStartTime Then     '1/26/99   save time first time only
                            llFirstEventTime = lmEvtTime
                        End If
                    Next ilLoop
                    imBreakNo = 0
                    imPositionNo = 0
                    ilFirstEvtShown = False         '1/26/99
                    Do While (ilRet = BTRV_ERR_NONE) And (tmSsf.iType = ilSSFType) And (tmSsf.iVefCode = ilVehCode) And (tmSsf.iDate(0) = ilSsfDate0) And (tmSsf.iDate(1) = ilSsfDate1)
                        'Loop thru Ssf and move records to tmOdf
                        ilEvt = 1
                        'ilFirstEvtShown = False
                        Do While ilEvt <= tmSsf.iCount
                            ilSortSeq = ilSortSeq + 1
                           LSet tmProg = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                            If (tmProg.iRecType = 1) Or ((tmProg.iRecType >= 2) And (tmProg.iRecType <= 9)) Then
                                gUnpackTimeLong tmProg.iStartTime(0), tmProg.iStartTime(1), False, lmEvtTime
                                If lmEvtTime > llEndTime Then
                                    ilWithinTime = False
                                    Exit Do
                                End If
                                If lmEvtTime >= llStartTime Then
                                    '1/26/99 see if first event is a other comment outside the first avail, if so need to look at LLC array later
                                    If llFirstEventTime = -1 Or llFirstEventTime > lmEvtTime Then
                                        ilWithinTime = True
                                        ilFirstEvtShown = True
                                        llFirstEventTime = -1
                                    Else
                                        ilEvt = ilEvt - 1
                                        llFirstEventTime = -1
                                        tmProg.iRecType = 0 'Aviod Break and Position increment
                                    End If
                                End If
                            End If
                            If blBypassZeroUnits And (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then
                               LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                If (tmAvail.iAvInfo And &H1F <= 0) Or (tmAvail.iLen <= 0) Then
                                    tmProg.iRecType = 0
                                End If
                            End If
                            llCefCode = 0
                            lmEvtIDCefCode = 0
                            lmAvailCefCode = 0
                            If tmProg.iRecType = 1 Then
                                imBreakNo = 0
                                imPositionNo = 0
                            ElseIf (tmProg.iRecType >= 2) And (tmProg.iRecType <= 9) Then
                                imBreakNo = imBreakNo + 1
                                imPositionNo = 0
                            End If
                            ilEvtFdIndex = -1
                            If ilWithinTime Then
                                If tmProg.iRecType = 1 Then    'Program
                                    For ilLoop = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                                        'Match start time and length
                                        If tlLLC(ilLoop).iEtfCode = 1 Then
                                            gPackTime tlLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                                            If (ilStartTime0 = tmProg.iStartTime(0)) And (ilStartTime1 = tmProg.iStartTime(1)) Then
                                                gAddTimeLength tlLLC(ilLoop).sStartTime, tlLLC(ilLoop).sLength, "A", "1", slTime, smXMid
                                                gPackTime slTime, ilEndTime0, ilEndTime1
                                                If (ilEndTime0 = tmProg.iEndTime(0)) And (ilEndTime1 = tmProg.iEndTime(1)) Then
                                                    smXMid = tlLLC(ilLoop).sXMid
                                                    ilEvtFdIndex = ilLoop
                                                    llCefCode = tlLLC(ilLoop).lCefCode
                                                    lmEvtIDCefCode = tlLLC(ilLoop).lEvtIDCefCode
                                                    ilType = 1
                                                    If slFor = "D" Then
                                                        'Only spot are sent to delivery
                                                    'ElseIf (tmVef.sType = "A") Or (ilDlfFound) Then
                                                    ElseIf (ilDlfFound) Then
                                                        'Obtain delivery entry to see is prog is sent
                                                        tmDlfSrchKey.iVefCode = ilVehCode
                                                        tmDlfSrchKey.sAirDay = slDay
                                                        tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                        tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                        tmDlfSrchKey.iAirTime(0) = tmProg.iStartTime(0)
                                                        tmDlfSrchKey.iAirTime(1) = tmProg.iStartTime(1)
                                                        ilRet = btrGetEqual(hlDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                        Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmProg.iStartTime(0)) And (tmDlf.iAirTime(1) = tmProg.iStartTime(1))
                                                            ilTerminated = False
                                                            If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
                                                                ilTerminated = True
                                                            Else
                                                                If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                                    If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                                        ilTerminated = True
                                                                    End If
                                                                End If
                                                            End If
                                                            If Not ilTerminated Then
                                                                If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
                                                                    'If slFor = "C" Then
                                                                        If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0) Then
                                                                            tmDlf.iMnfFeed = 0
                                                                            '6/4/16: Replaced GoSub
                                                                            'GoSub lProcProg
                                                                            mProcProg ilDlfFound, ilOtherGen, ilVpfIndex, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode, ilType, ilSSFType, ilLogDate0, ilLogDate1, ilStartTime0, ilStartTime1, ilSortSeq, tlLLC(), ilLoop, llCefCode
                                                                            DoEvents
                                                                        End If
                                                                    'End If
                                                                End If
                                                            End If
                                                            ilRet = btrGetNext(hlDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        Loop
                                                    Else
                                                        tmDlf.iVefCode = ilVehCode
                                                        tmDlf.iLocalTime(0) = tmProg.iStartTime(0)
                                                        tmDlf.iLocalTime(1) = tmProg.iStartTime(1)
                                                        tmDlf.iFeedTime(0) = tmProg.iStartTime(0)
                                                        tmDlf.iFeedTime(1) = tmProg.iStartTime(1)
                                                        tmDlf.sZone = ""
                                                        tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode
                                                        tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode
                                                        tmDlf.sProgCode = ""
                                                        tmDlf.iMnfFeed = 0
                                                        tmDlf.sBus = ""
                                                        tmDlf.sSchedule = ""
                                                        tmDlf.iMnfSubFeed = 0
                                                        '6/4/16: Replaced GoSub
                                                        'GoSub lProcProg
                                                        mProcProg ilDlfFound, ilOtherGen, ilVpfIndex, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode, ilType, ilSSFType, ilLogDate0, ilLogDate1, ilStartTime0, ilStartTime1, ilSortSeq, tlLLC(), ilLoop, llCefCode
                                                        DoEvents
                                                    End If
                                                    'Determine if any events that are
                                                    tlLLC(ilLoop).iEtfCode = -1 'Remove event
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next ilLoop
                                ElseIf (tmProg.iRecType = 2) Or ((tmProg.iRecType >= 6) And (tmProg.iRecType <= 9)) Then 'Avail
                                   LSet tmAvail = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                    'If tgSpf.sUsingBBs = "Y" Then
                                    '    ilBreakNo = 0
                                    '    imPositionNo = 0
                                    'End If
                                    For ilLoop = LBound(tlLLC) To UBound(tlLLC) - 1 Step 1
                                        'Match start time and length
                                        'If (tlLLC(ilLoop).iEtfCode = 1) And (tgSpf.sUsingBBs = "Y") Then
                                        '    ilBreakNo = 0
                                        'End If
                                        If (tlLLC(ilLoop).iEtfCode >= 2) And (tlLLC(ilLoop).iEtfCode <= 9) Then
                                            'If tgSpf.sUsingBBs = "Y" Then
                                            '    ilBreakNo = ilBreakNo + 1
                                            '    If tlLLC(ilLoop).iEtfCode = 3 Then
                                            '        tmOpenAvail.iAvInfo = ilBreakNo
                                            '    End If
                                            'End If
                                            gPackTime tlLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                                            If (ilStartTime0 = tmAvail.iTime(0)) And (ilStartTime1 = tmAvail.iTime(1)) Then
                                                ilEvtFdIndex = ilLoop
                                                llCefCode = tlLLC(ilLoop).lCefCode
                                                lmEvtIDCefCode = tlLLC(ilLoop).lEvtIDCefCode
                                                lmAvailCefCode = tlLLC(ilLoop).lCefCode
                                                'Scan to find Open & Close avail
                                                tmOpenAvail.iRecType = -1
                                                tmCloseAvail.iRecType = -1
                                                If tgSpf.sUsingBBs = "Y" Then
                                                    For ilFdOpen = ilEvtFdIndex - 1 To LBound(tlLLC) Step -1
                                                        If tlLLC(ilFdOpen).iEtfCode = 3 Then
                                                            tmOpenAvail.iRecType = Val(tlLLC(ilFdOpen).sType)
                                                            gPackTime tlLLC(ilFdOpen).sStartTime, tmOpenAvail.iTime(0), tmOpenAvail.iTime(1)
                                                            tmOpenAvail.iLtfCode = tlLLC(ilFdOpen).iLtfCode
                                                            tmOpenAvail.iLen = 0
                                                            tmOpenAvail.ianfCode = Val(tlLLC(ilFdOpen).sName)
                                                            tmOpenAvail.iNoSpotsThis = 0
                                                            tmOpenAvail.iOrigUnit = 0
                                                            tmOpenAvail.iOrigLen = 0
                                                            Exit For
                                                        End If
                                                    Next ilFdOpen
                                                    tmCloseAvail.iAvInfo = 0    'ilBreakNo
                                                    For ilFdClose = ilEvtFdIndex + 1 To UBound(tlLLC) - 1 Step 1
                                                        If tlLLC(ilFdClose).iEtfCode = 1 Then
                                                            tmCloseAvail.iAvInfo = 0
                                                        End If
                                                        If (tlLLC(ilFdClose).iEtfCode >= 2) And (tlLLC(ilFdClose).iEtfCode <= 9) Then
                                                            tmCloseAvail.iAvInfo = tmCloseAvail.iAvInfo + 1
                                                        End If
                                                        If tlLLC(ilFdClose).iEtfCode = 5 Then
                                                            tmCloseAvail.iRecType = Val(tlLLC(ilFdClose).sType)
                                                            gPackTime tlLLC(ilFdClose).sStartTime, tmCloseAvail.iTime(0), tmCloseAvail.iTime(1)
                                                            tmCloseAvail.iLtfCode = tlLLC(ilFdClose).iLtfCode
                                                            tmCloseAvail.iAvInfo = 0
                                                            tmCloseAvail.iLen = 0
                                                            tmCloseAvail.ianfCode = Val(tlLLC(ilFdClose).sName)
                                                            tmCloseAvail.iNoSpotsThis = 0
                                                            tmCloseAvail.iOrigUnit = 0
                                                            tmCloseAvail.iOrigLen = 0
                                                            Exit For
                                                        End If
                                                    Next ilFdClose
                                                End If
                                                tlCefSrchKey.lCode = llCefCode
                                                ilCefRecLen = Len(tlCef)
                                                ilRet = btrGetEqual(hlCef, tlCef, ilCefRecLen, tlCefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                If ilRet = BTRV_ERR_NONE Then
                                                    'slAvailComm = Trim$(Left$(tlCef.sComment, tlCef.iStrLen)) 'tlCef.sComment
                                                    slAvailComm = gStripChr0(tlCef.sComment)
                                                Else
                                                    slAvailComm = ""
                                                End If
                                                ilType = 3
                                                If slFor = "D" Then
                                                    'Only spot are sent to delivery
                                                'ElseIf (tmVef.sType = "A") Or (ilDlfFound) Then
                                                ElseIf (ilDlfFound) Then
                                                    'Don't include avails

                                                    'Obtain delivery entry to see is avail is sent
                                                    'tmDlfSrchKey.iVefCode = ilVehCode
                                                    'tmDlfSrchKey.sAirDay = slDay
                                                    'tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                    'tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                    'tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                    'tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                    'ilRet = btrGetEqual(hlDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                                    'Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                    '        ilTerminated = False
                                                    '        If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
                                                    '            ilTerminated = True
                                                    '        Else
                                                    '           If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                    '               If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                    '                   ilTerminated = True
                                                    '               End If
                                                    '           End If
                                                    '        End If
                                                    '        If Not ilTerminated Then
                                                    '        If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
                                                    '            'If slFor = "C" Then
                                                    '                If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0)  Then
                                                    '                    tmDlf.iMnfFeed = 0
                                                    '                    GoSub lProcAvail
                                                    '                End If
                                                    '            'End If
                                                    '        End If
                                                    '    End If
                                                    '   ilRet = btrGetNext(hlDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE)
                                                    'Loop
                                                Else
                                                    tmDlf.iVefCode = ilVehCode
                                                    tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                    tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                    tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                    tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                    tmDlf.sZone = ""
                                                    tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode
                                                    tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode
                                                    tmDlf.sProgCode = ""
                                                    tmDlf.iMnfFeed = 0
                                                    tmDlf.sBus = ""
                                                    tmDlf.sSchedule = ""
                                                    tmDlf.iMnfSubFeed = 0
                                                    'GoSub lProcAvail
                                                End If
                                                'Loop on spots, then add conflicting spots
                                                If (tmVef.sType = "A") Then
                                                    ilType = 4
                                                    tmVlfSrchKey1.iAirCode = ilVehCode
                                                    tmVlfSrchKey1.iAirDay = Val(slDay)
                                                    tmVlfSrchKey1.iEffDate(0) = ilVlfDate0
                                                    tmVlfSrchKey1.iEffDate(1) = ilVlfDate1
                                                    tmVlfSrchKey1.iAirTime(0) = tmAvail.iTime(0)
                                                    tmVlfSrchKey1.iAirTime(1) = tmAvail.iTime(1)
                                                    tmVlfSrchKey1.iAirPosNo = 0
                                                    tmVlfSrchKey1.iAirSeq = 1
                                                    ilRet = btrGetGreaterOrEqual(hlVlf, tmVlf, imVlfRecLen, tmVlfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
                                                    Do While (ilRet = BTRV_ERR_NONE) And (tmVlf.iAirCode = ilVehCode) And (tmVlf.iAirDay = Val(slDay)) And (tmVlf.iEffDate(0) = ilVlfDate0) And (tmVlf.iEffDate(1) = ilVlfDate1) And (tmVlf.iAirTime(0) = tmAvail.iTime(0)) And (tmVlf.iAirTime(1) = tmAvail.iTime(1))
                                                        ilTerminated = False
                                                        If (tmVlf.iTermDate(1) <> 0) Or (tmVlf.iTermDate(0) <> 0) Then
                                                            If (tmVlf.iTermDate(1) < tmVlf.iEffDate(1)) Or ((tmVlf.iEffDate(1) = tmVlf.iTermDate(1)) And (tmVlf.iTermDate(0) < tmVlf.iEffDate(0))) Then
                                                                ilTerminated = True
                                                            End If
                                                        End If
                                                        If (tmVlf.sStatus <> "P") And (Not ilTerminated) Then
                                                            If (tmCTSsf.iType <> ilSSFType) Or (tmCTSsf.iVefCode <> tmVlf.iSellCode) Or (tmCTSsf.iDate(0) <> ilLogDate0) Or (tmCTSsf.iDate(1) <> ilLogDate1) Then
                                                                ilSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                                                                tlSsfSrchKey.iType = ilSSFType
                                                                tlSsfSrchKey.iVefCode = tmVlf.iSellCode
                                                                tlSsfSrchKey.iDate(0) = ilLogDate0
                                                                tlSsfSrchKey.iDate(1) = ilLogDate1
                                                                tlSsfSrchKey.iStartTime(0) = 0
                                                                tlSsfSrchKey.iStartTime(1) = 0
                                                                ilRet = gSSFGetEqual(hlCTSsf, tmCTSsf, ilSsfRecLen, tlSsfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                            End If
                                                            Do While (ilRet = BTRV_ERR_NONE) And (tmCTSsf.iType = ilSSFType) And (tmCTSsf.iVefCode = tmVlf.iSellCode) And (tmCTSsf.iDate(0) = ilLogDate0) And (tmCTSsf.iDate(1) = ilLogDate1)
                                                                For ilSIndex = 1 To tmCTSsf.iCount Step 1
                                                                    tmAvailTest = tmCTSsf.tPas(ADJSSFPASBZ + ilSIndex)
                                                                    If ((tmAvailTest.iRecType >= 2) And (tmAvailTest.iRecType <= 9)) Then
                                                                        If (tmAvailTest.iTime(0) = tmVlf.iSellTime(0)) And (tmAvailTest.iTime(1) = tmVlf.iSellTime(1)) Then
                                                                            tmAvAvail = tmAvailTest
                                                                            ilAvEvt = ilSIndex
                                                                            ilAvVefCode = tmVlf.iSellCode
                                                                            '6/5/16: Replaced GoSub
                                                                            'GoSub lChkOpenAvail
                                                                            mChkOpenAvail ilAvVefCode, ilAvEvt, ilDlfFound, slDay, ilLoop, tlLLC(), hlDlf, ilDlfDate0, ilDlfDate1, ilODFVefCode, ilSimVefCode, ilVpfIndex, ilSSFType, ilLogDate0, ilLogDate1, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlODF, tlOdf, ilVehCode, ilGenLST, ilLSTForLogVeh, ilSec, ilUnits, llGsfCode, ilOtherGen, ilTerminated
                                                                            For ilSpot = 1 To tmAvailTest.iNoSpotsThis Step 1
                                                                                ilType = 4
                                                                               LSet tmSpot = tmCTSsf.tPas(ADJSSFPASBZ + ilSpot + ilSIndex)
                                                                                imClearLstSdf = True
                                                                                tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                                                ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                                If ilRet = BTRV_ERR_NONE Then
                                                                                    If tmSdf.lChfCode = 0 Then
                                                                                        tmFsfSrchKey0.lCode = tmSdf.lFsfCode
                                                                                        ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                                        gMoveFeedToCntr tmFsf, tmRdf, tmChf, tmClf, tmFCff(), hmFnf, hmPrf
                                                                                    Else
                                                                                        tmChfSrchKey.lCode = tmSdf.lChfCode
                                                                                        ilRet = btrGetEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                                                    End If
                                                                                    If ilRet = BTRV_ERR_NONE Then
                                                                                        imPositionNo = imPositionNo + 1
                                                                                        
                                                                                        '4/30/11:  Add Region copy by Airing vehicle
                                                                                        tlSdf = tmSdf
                                                                                        '10451
                                                                                        gTestForRegionAirCopy tmVef.sType, tmVef.iCode, ilVpfIndex, tlSdf, tmAvailTest, hlCrf, hlCnf, hmCif, hmRsf, hmCvf, hmClf
                                                                                        
                                                                                        slZone = "Oth"
                                                                                        tlSdf = tmSdf
                                                                                        '10451
                                                                                        gTestForAirCopy 0, tmVef.sType, tmVef.iCode, ilVpfIndex, tlSdf, tmAvailTest, hlCrf, hlCnf, hmCif, hmRsf, hmCvf, hmClf, slZone, ilZoneFd, ilCopyReplaced, ilRotNo
                                                                                        If ilCopyReplaced Then
                                                                                            llReplaceCopyCode = tlSdf.lCopyCode
                                                                                            ilReplaceRotNo = ilRotNo
                                                                                        Else
                                                                                            llReplaceCopyCode = 0
                                                                                        End If
                                                                                        If ilDlfFound Then
                                                                                            tlSdf = tmSdf
                                                                                            slZone = "EST"  'Use EST as standard, if not found, use OTH
                                                                                            '10451
                                                                                            gTestForAirCopy 2, tmVef.sType, tmVef.iCode, ilVpfIndex, tlSdf, tmAvailTest, hlCrf, hlCnf, hmCif, hmRsf, hmCvf, hmClf, slZone, ilZoneFd, ilCopyReplaced, ilRotNo
                                                                                            If (Not ilCopyReplaced) And (llReplaceCopyCode > 0) Then
                                                                                                tlSdf.sPtType = "1"
                                                                                                tlSdf.lCopyCode = llReplaceCopyCode
                                                                                            'Need RotNo set so that CSF is obtained
                                                                                                tlSdf.iRotNo = ilReplaceRotNo
                                                                                            'ElseIf (ilCopyReplaced) And ilZoneFd Then
                                                                                            ElseIf (ilCopyReplaced) Then
                                                                                                tlSdf.iRotNo = ilRotNo
                                                                                            End If
                                                                                            llCifCode = mObtainCifCode(tlSdf, slZone, hlTzf, ilOther)
                                                                                            ilRet = gGetCrfVefCode(hmClf, tlSdf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
                                                                                            lmCrfCsfCode = mObtainCrfCsfCode(tlSdf, slZone, hlCrf, hlTzf, hmCvf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, tmVef.iCode, llCrfCode)
                                                                                            ''Remove comment
                                                                                            'llCrfCsfCode = 0
                                                                                            smBonus = mSetBonusFlag(tlSdf)  '2-15-01
                                                                                            tmDlfSrchKey.iVefCode = ilVehCode
                                                                                            tmDlfSrchKey.sAirDay = slDay
                                                                                            tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                                                            tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                                                            tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                                                            tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                                                            ilRet = btrGetEqual(hlDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                                            Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                                                                ilTerminated = False
                                                                                                If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
                                                                                                    ilTerminated = True
                                                                                                Else
                                                                                                    If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                                                                        If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                                                                            ilTerminated = True
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                                If Not ilTerminated Then
                                                                                                    If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
                                                                                                        If slFor = "D" Then
                                                                                                            If tmDlf.sFed = "Y" Then
                                                                                                                '6/5/16: Replaced GoSub
                                                                                                                'GoSub lProcSpot
                                                                                                                mProcSpot slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, slFor, ilSeqNo, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode, ilSSFType, ilLogDate0, ilLogDate1, llCifCode, tlCxf, ilType, ilVpfIndex, tlAdf, hlVsf, hlSif, ilGenLST, ilLSTForLogVeh, hlMcf, ilExportType, llGsfCode, llCrfCode
                                                                                                                DoEvents
                                                                                                            End If
                                                                                                        'ElseIf slFor = "C" Then
                                                                                                        Else
                                                                                                            If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0) Then
                                                                                                                tmDlf.iMnfFeed = 0
                                                                                                                '6/5/16: Replaced GoSub
                                                                                                                'GoSub lProcSpot
                                                                                                                mProcSpot slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, slFor, ilSeqNo, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode, ilSSFType, ilLogDate0, ilLogDate1, llCifCode, tlCxf, ilType, ilVpfIndex, tlAdf, hlVsf, hlSif, ilGenLST, ilLSTForLogVeh, hlMcf, ilExportType, llGsfCode, llCrfCode
                                                                                                                DoEvents
                                                                                                            End If
                                                                                                        'Else
                                                                                                        '    GoSub lProcSpot
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                                ilRet = btrGetNext(hlDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                                            Loop
                                                                                        Else
                                                                                            'Assign copy for all zones
                                                                                            'ilNoZones = 0
                                                                                            'For ilZone = 1 To 5 Step 1
                                                                                            '    Select Case ilZone
                                                                                            '        Case 1
                                                                                            '            slZone = "EST"
                                                                                            '        Case 2
                                                                                            '            slZone = "MST"
                                                                                            '        Case 3
                                                                                            '            slZone = "CST"
                                                                                            '        Case 4
                                                                                            '            slZone = "PST"
                                                                                            '        Case 5
                                                                                            '            slZone = "Oth"
                                                                                            '    End Select
                                                                                            ilNoZones = 0
                                                                                            gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llAirTime
                                                                                            ilOtherGen = False
                                                                                            For ilZone = 0 To UBound(tmZoneInfo) - 1 Step 1
                                                                                                'If ((tmZoneInfo(ilZone).sFed = "*") And (tgSpf.sGUseAffSys = "Y")) Or (tgSpf.sGUseAffSys <> "Y") Then
                                                                                                    Select Case tmZoneInfo(ilZone).sZone
                                                                                                        Case "E"
                                                                                                            slZone = "EST"
                                                                                                        Case "M"
                                                                                                            slZone = "MST"
                                                                                                        Case "C"
                                                                                                            slZone = "CST"
                                                                                                        Case "P"
                                                                                                            slZone = "PST"
                                                                                                        Case Else
                                                                                                            If ilOtherGen Then
                                                                                                                Exit For
                                                                                                            End If
                                                                                                            slZone = "Oth"
                                                                                                    End Select
                                                                                                    tlSdf = tmSdf
                                                                                                    If slZone <> "Oth" Then
                                                                                                        '10451
                                                                                                        gTestForAirCopy 2, tmVef.sType, tmVef.iCode, ilVpfIndex, tlSdf, tmAvailTest, hlCrf, hlCnf, hmCif, hmRsf, hmCvf, hmClf, slZone, ilZoneFd, ilCopyReplaced, ilRotNo
                                                                                                    Else
                                                                                                        ilCopyReplaced = False
                                                                                                    End If
                                                                                                    If (Not ilCopyReplaced) And (llReplaceCopyCode > 0) Then
                                                                                                        tlSdf.sPtType = "1"
                                                                                                        tlSdf.lCopyCode = llReplaceCopyCode
                                                                                                    'Need RotNo set so that CSF is obtained
                                                                                                        tlSdf.iRotNo = ilReplaceRotNo
                                                                                                    'ElseIf ilZoneFd Then
                                                                                                    '    tlSdf.iRotNo = 0
                                                                                                    ElseIf ilCopyReplaced Then
                                                                                                        tlSdf.iRotNo = ilRotNo
                                                                                                    End If
                                                                                                    llCifCode = mObtainCifCode(tlSdf, slZone, hlTzf, ilOther)
                                                                                                    'If ilZoneFd Then
                                                                                                    '    ilOther = False
                                                                                                    'End If
                                                                                                    If (Not ilOther) Or (Not ilOtherGen) Or (tgSpf.sGUseAffSys = "Y") Then
                                                                                                        'To handle all vehicles with zone defined, use the code with sGUseAffSys = "Y", remove else
                                                                                                        '9-23-04 allow logs to show time zone when not using affiliate system
                                                                                                        'If (tgSpf.sGUseAffSys = "Y") Then
                                                                                                            If slZone <> "Oth" Then
                                                                                                                tmDlf.sZone = slZone
                                                                                                            Else
                                                                                                                tmDlf.sZone = ""
                                                                                                            End If
                                                                                                        'Else
                                                                                                        '   If ilOther Then
                                                                                                        '        ilOtherGen = True
                                                                                                        '        tmDlf.sZone = ""
                                                                                                        '    Else
                                                                                                        '        tmDlf.sZone = slZone
                                                                                                        '    End If
                                                                                                        'End If
                                                                                                        ilRet = gGetCrfVefCode(hmClf, tlSdf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
                                                                                                        lmCrfCsfCode = mObtainCrfCsfCode(tlSdf, slZone, hlCrf, hlTzf, hmCvf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, tmVef.iCode, llCrfCode)
                                                                                                        smBonus = mSetBonusFlag(tlSdf)  '2-15-01
                                                                                                        'tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                                                                        'tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                                                                        'tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                                                                        'tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                                                                        If (tgSpf.sGUseAffSys = "Y") Then
                                                                                                            llTime = llAirTime + tmZoneInfo(ilZone).lGLocalAdj
                                                                                                            If llTime < 0 Then
                                                                                                                llTime = llTime + 86400
                                                                                                            ElseIf llTime > 86400 Then
                                                                                                                llTime = llTime - 86400
                                                                                                            End If
                                                                                                            gPackTimeLong llTime, tmDlf.iLocalTime(0), tmDlf.iLocalTime(1)
                                                                                                            llTime = llAirTime + tmZoneInfo(ilZone).lGFeedAdj
                                                                                                            If llTime < 0 Then
                                                                                                                llTime = llTime + 86400
                                                                                                            ElseIf llTime > 86400 Then
                                                                                                                llTime = llTime - 86400
                                                                                                            End If
                                                                                                            gPackTimeLong llTime, tmDlf.iFeedTime(0), tmDlf.iFeedTime(1)
                                                                                                        Else
                                                                                                            tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                                                                            tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                                                                            tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                                                                            tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                                                                        End If
                                                                                                        'If (ilOther) And (ilNoZones = 0) Then
                                                                                                        '    tmDlf.sZone = ""
                                                                                                        'Else
                                                                                                        '    tmDlf.sZone = slZone
                                                                                                        '    If ilNoZones = 0 Then
                                                                                                        '        ilType = 4
                                                                                                        '        ilNoZones = 1
                                                                                                        '    Else
                                                                                                        '        ilType = 5
                                                                                                        '        ilOther = False
                                                                                                        '        ilNoZones = ilNoZones + 1
                                                                                                        '    End If
                                                                                                        'End If
                                                                                                        tmDlf.iEtfCode = 0
                                                                                                        tmDlf.iEnfCode = 0
                                                                                                        ''tmDlf.sProgCode = ""
                                                                                                        ''1/10/98 AMFM wants comment in all spots in break
                                                                                                        ''If imPositionNo = 1 Then        '1st position of spot has the avail spot code
                                                                                                            tmDlf.sProgCode = slAvailComm
                                                                                                        ''Else
                                                                                                        ''    tmDlf.sProgCode = ""
                                                                                                        ''End If
                                                                                                        tmDlf.iMnfFeed = 0
                                                                                                        tmDlf.sBus = ""
                                                                                                        tmDlf.sSchedule = ""
                                                                                                        tmDlf.iMnfSubFeed = 0
                                                                                                        '6/5/16: Replaced GoSub
                                                                                                        'GoSub lProcSpot
                                                                                                        mProcSpot slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, slFor, ilSeqNo, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode, ilSSFType, ilLogDate0, ilLogDate1, llCifCode, tlCxf, ilType, ilVpfIndex, tlAdf, hlVsf, hlSif, ilGenLST, ilLSTForLogVeh, hlMcf, ilExportType, llGsfCode, llCrfCode
                                                                                                        DoEvents
                                                                                                        'If (ilOther) Or (ilNoZones = 4) Then
                                                                                                        '    Exit For
                                                                                                        'End If
                                                                                                    End If
                                                                                                'End If
                                                                                            Next ilZone
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            Next ilSpot
                                                                            Exit Do
                                                                        End If
                                                                    End If
                                                                Next ilSIndex
                                                                ilSsfRecLen = Len(tmCTSsf) 'Max size of variable length record
                                                                ilRet = gSSFGetNext(hlCTSsf, tmCTSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                            Loop
                                                        End If
                                                        ilRet = btrGetNext(hlVlf, tmVlf, imVlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                    Loop
                                                Else
                                                    tmCTSsf = tmSsf
                                                    tmAvAvail = tmAvail
                                                    ilAvEvt = ilEvt
                                                    ilAvVefCode = ilVehCode
                                                    '6/5/16: Replaced GoSub
                                                    'GoSub lChkOpenAvail
                                                    mChkOpenAvail ilAvVefCode, ilAvEvt, ilDlfFound, slDay, ilLoop, tlLLC(), hlDlf, ilDlfDate0, ilDlfDate1, ilODFVefCode, ilSimVefCode, ilVpfIndex, ilSSFType, ilLogDate0, ilLogDate1, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlODF, tlOdf, ilVehCode, ilGenLST, ilLSTForLogVeh, ilSec, ilUnits, llGsfCode, ilOtherGen, ilTerminated
                                                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llBBTime
                                                    For ilSpot = 1 To tmAvail.iNoSpotsThis Step 1
                                                        ilType = 4
                                                        ilEvt = ilEvt + 1
                                                       LSet tmSpot = tmSsf.tPas(ADJSSFPASBZ + ilEvt)
                                                        imClearLstSdf = True
                                                        tmSdfSrchKey3.lCode = tmSpot.lSdfCode
                                                        ilRet = btrGetEqual(hlSdf, tmSdf, imSdfRecLen, tmSdfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)
                                                        If ilRet = BTRV_ERR_NONE Then
                                                            If tmSdf.lChfCode = 0 Then
                                                                tmFsfSrchKey0.lCode = tmSdf.lFsfCode
                                                                ilRet = btrGetEqual(hmFsf, tmFsf, imFsfRecLen, tmFsfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                gMoveFeedToCntr tmFsf, tmRdf, tmChf, tmClf, tmFCff(), hmFnf, hmPrf
                                                            Else
                                                                tmChfSrchKey.lCode = tmSdf.lChfCode
                                                                ilRet = btrGetEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                                                                If tgSpf.sUsingBBs = "Y" Then
                                                                    tmClfSrchKey.lChfCode = tmSdf.lChfCode
                                                                    tmClfSrchKey.iLine = tmSdf.iLineNo
                                                                    tmClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
                                                                    tmClfSrchKey.iPropVer = 32000 ' Plug with very high number
                                                                    imClfRecLen = Len(tmClf)
                                                                    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, imClfRecLen, tmClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
                                                                    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
                                                                        ilRet = btrGetNext(hmClf, tmClf, imClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                    Loop
                                                                End If
                                                            End If
                                                            If ilRet = BTRV_ERR_NONE Then
                                                                imPositionNo = imPositionNo + 1
                                                                If (tmSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
                                                                    ilLastWasSplit = True
                                                                    If imPositionNo > 1 Then
                                                                        'Leave room for Split Fill
                                                                        imPositionNo = imPositionNo + 1
                                                                    End If
                                                                ElseIf (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
                                                                    ilLastWasSplit = True
                                                                Else
                                                                    If (ilLastWasSplit) And (imPositionNo > 1) Then
                                                                        imPositionNo = imPositionNo + 1
                                                                    End If
                                                                    ilLastWasSplit = False
                                                                End If
                                                                slZone = "EST"  'Use EST as standard, if not found, use OTH
                                                                If ilDlfFound Then
                                                                    tlSdf = tmSdf
                                                                    'gObtainAirCopy 0, tmVef.sType, tmVef.iCode, ilVpfIndex, tlSdf, tmAvailTest, hlCrf, hlCnf, hmCif, slZone, ilZoneFd
                                                                    llCifCode = mObtainCifCode(tlSdf, slZone, hlTzf, ilOther)
                                                                    ilRet = gGetCrfVefCode(hmClf, tlSdf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
                                                                    lmCrfCsfCode = mObtainCrfCsfCode(tlSdf, slZone, hlCrf, hlTzf, hmCvf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, 0, llCrfCode)
                                                                    smBonus = mSetBonusFlag(tlSdf)  '2-15-01
                                                                    ''Remove Comment
                                                                    'llCrfCsfCode = 0
                                                                    'Obtain delivery entry to see is avail is sent
                                                                    tmDlfSrchKey.iVefCode = ilVehCode
                                                                    tmDlfSrchKey.sAirDay = slDay
                                                                    tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                                    tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                                    tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0)
                                                                    tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1)
                                                                    ilRet = btrGetEqual(hlDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                                    Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvail.iTime(1))
                                                                        ilTerminated = False
                                                                        If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
                                                                            ilTerminated = True
                                                                        Else
                                                                            If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                                                If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                                                    ilTerminated = True
                                                                                End If
                                                                            End If
                                                                        End If
                                                                        If Not ilTerminated Then
                                                                            If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
                                                                                If slFor = "D" Then
                                                                                    If tmDlf.sFed = "Y" Then
                                                                                        '6/5/16: Replaced GoSub
                                                                                        'GoSub lProcSpot
                                                                                        mProcSpot slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, slFor, ilSeqNo, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode, ilSSFType, ilLogDate0, ilLogDate1, llCifCode, tlCxf, ilType, ilVpfIndex, tlAdf, hlVsf, hlSif, ilGenLST, ilLSTForLogVeh, hlMcf, ilExportType, llGsfCode, llCrfCode
                                                                                        DoEvents
                                                                                    End If
                                                                                'ElseIf slFor = "C" Then
                                                                                Else
                                                                                    If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0) Then
                                                                                        tmDlf.iMnfFeed = 0
                                                                                        '6/5/16: Replaced GoSub
                                                                                        'GoSub lProcSpot
                                                                                        mProcSpot slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, slFor, ilSeqNo, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode, ilSSFType, ilLogDate0, ilLogDate1, llCifCode, tlCxf, ilType, ilVpfIndex, tlAdf, hlVsf, hlSif, ilGenLST, ilLSTForLogVeh, hlMcf, ilExportType, llGsfCode, llCrfCode
                                                                                        DoEvents
                                                                                    End If
                                                                                'Else
                                                                                '    GoSub lProcSpot
                                                                                End If
                                                                            End If
                                                                        End If
                                                                        ilRet = btrGetNext(hlDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                                    Loop
                                                                Else
                                                                    'Assign copy for all zones
                                                                    'ilNoZones = 0
                                                                    'For ilZone = 1 To 5 Step 1
                                                                    '    Select Case ilZone
                                                                    '        Case 1
                                                                    '            slZone = "EST"
                                                                    '        Case 2
                                                                    '            slZone = "MST"
                                                                    '        Case 3
                                                                    '            slZone = "CST"
                                                                    '        Case 4
                                                                    '            slZone = "PST"
                                                                    '        Case 5
                                                                    '            slZone = "Oth"
                                                                    '    End Select
                                                                    ilBBPass = 0
                                                                    Do
                                                                        ilNoZones = 0
                                                                        gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llAirTime
                                                                        ilOtherGen = False
                                                                        For ilZone = 0 To UBound(tmZoneInfo) - 1 Step 1
                                                                            'If ((tmZoneInfo(ilZone).sFed = "*") And (tgSpf.sGUseAffSys = "Y")) Or (tgSpf.sGUseAffSys <> "Y") Then
                                                                                Select Case tmZoneInfo(ilZone).sZone
                                                                                    Case "E"
                                                                                        slZone = "EST"
                                                                                    Case "M"
                                                                                        slZone = "MST"
                                                                                    Case "C"
                                                                                        slZone = "CST"
                                                                                    Case "P"
                                                                                        slZone = "PST"
                                                                                    Case Else
                                                                                        If ilOtherGen Then
                                                                                            Exit For
                                                                                        End If
                                                                                        slZone = "Oth"
                                                                                End Select
                                                                                tlSdf = tmSdf
                                                                                'gObtainAirCopy 0, tmVef.sType, tmVef.iCode, ilVpfIndex, tlSdf, tmAvailTest, hlCrf, hlCnf, hmCif, slZone, ilZoneFd
                                                                                llCifCode = mObtainCifCode(tlSdf, slZone, hlTzf, ilOther)
                                                                                If (Not ilOther) Or (Not ilOtherGen) Or (tgSpf.sGUseAffSys = "Y") Then
                                                                                    'To handle all vehicles with zone defined, use the code with sGUseAffSys = "Y", remove else
                                                                                    '9-23-04 allow logs to show time zone if not using affiliate system
                                                                                    'If (tgSpf.sGUseAffSys = "Y") Then
                                                                                        If slZone <> "Oth" Then
                                                                                            tmDlf.sZone = slZone
                                                                                        Else
                                                                                            tmDlf.sZone = ""
                                                                                        End If
                                                                                   'Else
                                                                                   '     If ilOther Then
                                                                                   '         ilOtherGen = True
                                                                                   '         tmDlf.sZone = ""
                                                                                   '     Else
                                                                                   '         tmDlf.sZone = slZone
                                                                                   '     End If
                                                                                   ' End If
                                                                                    ilRet = gGetCrfVefCode(hmClf, tlSdf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, slLive, ilRdfCode)
                                                                                    lmCrfCsfCode = mObtainCrfCsfCode(tlSdf, slZone, hlCrf, hlTzf, hmCvf, ilCrfVefCode, ilPkgVefCode, ilLnVefCode, 0, llCrfCode)
                                                                                    smBonus = mSetBonusFlag(tlSdf)  '2-15-01

                                                                                    If (tgSpf.sGUseAffSys = "Y") Then
                                                                                        llTime = llAirTime + tmZoneInfo(ilZone).lGLocalAdj
                                                                                        If llTime < 0 Then
                                                                                            llTime = llTime + 86400
                                                                                        ElseIf llTime > 86400 Then
                                                                                            llTime = llTime - 86400
                                                                                        End If
                                                                                        gPackTimeLong llTime, tmDlf.iLocalTime(0), tmDlf.iLocalTime(1)
                                                                                        llTime = llAirTime + tmZoneInfo(ilZone).lGFeedAdj
                                                                                        If llTime < 0 Then
                                                                                            llTime = llTime + 86400
                                                                                        ElseIf llTime > 86400 Then
                                                                                            llTime = llTime - 86400
                                                                                        End If
                                                                                        gPackTimeLong llTime, tmDlf.iFeedTime(0), tmDlf.iFeedTime(1)
                                                                                    Else
                                                                                        tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                                                                                        tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                                                                                        tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                                                                                        tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                                                                                    End If
                                                                                    'If (ilOther) And (ilNoZones = 0) Then
                                                                                    '    tmDlf.sZone = ""
                                                                                    'Else
                                                                                    '    tmDlf.sZone = slZone
                                                                                    '    If ilNoZones = 0 Then
                                                                                    '        ilType = 4
                                                                                    '        ilNoZones = 1
                                                                                    '    Else
                                                                                    '        ilType = 5
                                                                                    '        ilOther = False
                                                                                    '        ilNoZones = ilNoZones + 1
                                                                                    '    End If
                                                                                    'End If
                                                                                    tmDlf.iEtfCode = 0
                                                                                    tmDlf.iEnfCode = 0
                                                                                    ''tmDlf.sProgCode = ""
                                                                                    ''1/10/98 AMFM wants comment in all spots in break
                                                                                    ''If imPositionNo = 1 Then        '1st position of spot has the avail spot code
                                                                                        tmDlf.sProgCode = slAvailComm
                                                                                    ''Else
                                                                                    ''    tmDlf.sProgCode = ""
                                                                                    ''End If
                                                                                    tmDlf.iMnfFeed = 0
                                                                                    tmDlf.sBus = ""
                                                                                    tmDlf.sSchedule = ""
                                                                                    tmDlf.iMnfSubFeed = 0
                                                                                    '6/5/16: Replaced GoSub
                                                                                    'GoSub lProcSpot
                                                                                    mProcSpot slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, slFor, ilSeqNo, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode, ilSSFType, ilLogDate0, ilLogDate1, llCifCode, tlCxf, ilType, ilVpfIndex, tlAdf, hlVsf, hlSif, ilGenLST, ilLSTForLogVeh, hlMcf, ilExportType, llGsfCode, llCrfCode
                                                                                    DoEvents
                                                                                    'If (ilOther) Or (ilNoZones = 4) Then
                                                                                    '    Exit For
                                                                                    'End If
                                                                                End If
                                                                            'End If
                                                                        Next ilZone
                                                                       LSet tmAvail = tmAvAvail

                                                                        If (tgSpf.sUsingBBs <> "Y") Or (ilBBPass >= 2) Then
                                                                            Exit Do
                                                                        End If
                                                                        ilBBLen = tmClf.iBBOpenLen
                                                                        If (ilBBLen > 0) And (tmOpenAvail.iRecType <> -1) And (ilBBPass = 0) Then
                                                                            'Find BB spot
                                                                            ilFound = gFindBBSpot(hlSdf, "O", tmSdf.iVefCode, tmSdf.lChfCode, tmSdf.iLineNo, llDate, llBBTime, tlSdf, llPrevOpenFdBBSpots())
                                                                            If Not ilFound Then
                                                                                ilBBPass = 1
                                                                            Else
                                                                               LSet tmAvail = tmOpenAvail
                                                                                tmSdf = tlSdf
                                                                                tmSdf.iTime(0) = tmOpenAvail.iTime(0)
                                                                                tmSdf.iTime(1) = tmOpenAvail.iTime(1)
                                                                            End If
                                                                        Else
                                                                            ilBBPass = 1
                                                                            If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                                                                                ilBBLen = tmClf.iBBOpenLen
                                                                            Else
                                                                                ilBBLen = tmClf.iBBCloseLen
                                                                            End If
                                                                        End If
                                                                        If (ilBBLen > 0) And (tmCloseAvail.iRecType <> -1) And (ilBBPass = 1) Then
                                                                            'Find BB spot
                                                                            ilFound = gFindBBSpot(hlSdf, "C", tmSdf.iVefCode, tmSdf.lChfCode, tmSdf.iLineNo, llDate, llBBTime, tlSdf, llPrevCloseFdBBSpots())
                                                                            If Not ilFound Then
                                                                                Exit Do
                                                                            Else
                                                                               LSet tmAvail = tmCloseAvail
                                                                                tmSdf = tlSdf
                                                                                tmSdf.iTime(0) = tmCloseAvail.iTime(0)
                                                                                tmSdf.iTime(1) = tmCloseAvail.iTime(1)
                                                                            End If
                                                                        Else
                                                                            If ilBBPass = 1 Then
                                                                                Exit Do
                                                                            End If
                                                                        End If
                                                                        ilBBPass = ilBBPass + 1
                                                                    Loop
                                                                End If
                                                            End If
                                                        End If
                                                    Next ilSpot
                                                End If
                                                tlLLC(ilLoop).iEtfCode = -1 'Remove event
                                                Exit For
                                            End If
                                        End If
                                    Next ilLoop
                                End If
                            End If
                            'Any other events to be sent out
                            If (ilEvtFdIndex >= 0) Or (Not ilFirstEvtShown) Then
                                If Not ilFirstEvtShown Then
                                    ilEvtFdIndex = LBound(tlLLC) - 1
                                    ilFirstEvtShown = True
                                End If
                                'Output all event until Program or Avail found
                                For ilLoop = ilEvtFdIndex + 1 To UBound(tlLLC) - 1 Step 1
                                    'Match start time and length
                                    If (tlLLC(ilLoop).iEtfCode >= 0) Then
                                        If (tlLLC(ilLoop).iEtfCode > 9) Then
                                            smXMid = tlLLC(ilLoop).sXMid
                                            'Handle like program
                                            ilType = 2
                                            lmEvtTime = CLng(gTimeToCurrency(tlLLC(ilLoop).sStartTime, False))
                                            If lmEvtTime > llEndTime Then
                                                ilWithinTime = False
                                                Exit Do
                                            End If
                                            If lmEvtTime >= llStartTime Then
                                                ilWithinTime = True
                                            End If
                                            If ilWithinTime Then
                                                If (tlLLC(ilLoop).iEtfCode > 13) Then
                                                    lmEvtCefCode = tlLLC(ilLoop).lCefCode
                                                    imEvtCefSort = imEvtCefSort + 1
                                                End If
                                                llCefCode = tlLLC(ilLoop).lCefCode
                                                lmEvtIDCefCode = tlLLC(ilLoop).lEvtIDCefCode
                                                lmAvailCefCode = tlLLC(ilLoop).lCefCode
                                                If slFor = "D" Then
                                                    'Only spot are sent to delivery
                                                ElseIf (ilDlfFound) Then
                                                    'Obtain delivery entry to see is prog is sent
                                                    tmDlfSrchKey.iVefCode = ilVehCode
                                                    tmDlfSrchKey.sAirDay = slDay
                                                    tmDlfSrchKey.iStartDate(0) = ilDlfDate0
                                                    tmDlfSrchKey.iStartDate(1) = ilDlfDate1
                                                    gPackTime tlLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                                                    tmDlfSrchKey.iAirTime(0) = ilStartTime0
                                                    tmDlfSrchKey.iAirTime(1) = ilStartTime1
                                                    ilRet = btrGetEqual(hlDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                                                    Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = ilStartTime0) And (tmDlf.iAirTime(1) = ilStartTime1)
                                                        ilTerminated = False
                                                        If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
                                                            ilTerminated = True
                                                        Else
                                                            If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                                                                If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                                                                    ilTerminated = True
                                                                End If
                                                            End If
                                                        End If
                                                        If Not ilTerminated Then
                                                            If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
                                                                'If slFor = "C" Then
                                                                    If tmDlf.sCmmlSched = "Y" Then
                                                                        tmDlf.iMnfFeed = 0
                                                                        '6/4/16: Replaced GoSub
                                                                        'GoSub lProcProg
                                                                        mProcProg ilDlfFound, ilOtherGen, ilVpfIndex, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode, ilType, ilSSFType, ilLogDate0, ilLogDate1, ilStartTime0, ilStartTime1, ilSortSeq, tlLLC(), ilLoop, llCefCode
                                                                        DoEvents
                                                                    End If
                                                                'End If
                                                            End If
                                                        End If
                                                        ilRet = btrGetNext(hlDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                                                    Loop
                                                Else
                                                    tmDlf.iVefCode = ilVehCode
                                                    gPackTime tlLLC(ilLoop).sStartTime, ilStartTime0, ilStartTime1
                                                    tmDlf.iLocalTime(0) = ilStartTime0
                                                    tmDlf.iLocalTime(1) = ilStartTime1
                                                    tmDlf.iFeedTime(0) = ilStartTime0
                                                    tmDlf.iFeedTime(1) = ilStartTime1
                                                    tmDlf.sZone = ""
                                                    tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode
                                                    tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode
                                                    tmDlf.sProgCode = ""
                                                    tmDlf.iMnfFeed = 0
                                                    tmDlf.sBus = ""
                                                    tmDlf.sSchedule = ""
                                                    tmDlf.iMnfSubFeed = 0
                                                    '6/4/16: Replaced GoSub
                                                    'GoSub lProcProg
                                                    mProcProg ilDlfFound, ilOtherGen, ilVpfIndex, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode, ilType, ilSSFType, ilLogDate0, ilLogDate1, ilStartTime0, ilStartTime1, ilSortSeq, tlLLC(), ilLoop, llCefCode
                                                    DoEvents
                                                End If
                                            End If
                                            tlLLC(ilLoop).iEtfCode = -1 'Remove event
                                        Else
                                            Exit For
                                        End If
                                    End If
                                Next ilLoop
                            End If

                            ilEvt = ilEvt + 1
                        Loop
                        ilSsfRecLen = Len(tmSsf) 'Max size of variable length record
                        ilRet = gSSFGetNext(hlSsf, tmSsf, ilSsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
                    Loop
                End If
                If (Not imCreateNewLST) And (imLstExist) And (ilGenLST) Then
                    'Remove unused LST if not combine vehicle or it is the second part of the combine
                    If ilCombineVefCode = 0 Then
                        '11/4/09: Re-add Generation of LST for Log Vehicles
                        'mDeleteUnusedLST
                        If (ilLSTForLogVeh > 0) And (tmVef.iVefCode > 0) Then
                            For ilLst = LBound(tmLstCode) To UBound(tmLstCode) - 1 Step 1
                                tmCombineLstInfo(UBound(tmCombineLstInfo)).lDate = llDate
                                tmCombineLstInfo(UBound(tmCombineLstInfo)).tLst = tmLstCode(ilLst)
                                ReDim Preserve tmCombineLstInfo(0 To UBound(tmCombineLstInfo) + 1) As COMBINELSTINFO
                            Next ilLst
                        Else
                            '8/10/10:  Handle games that cross midnight
                            'mDeleteUnusedLST
                            If llGsfCode = 0 Then
                                mDeleteUnusedLST slInLogType, ilExportType
                            End If
                        End If
                    Else
                        For ilLst = LBound(tmLstCode) To UBound(tmLstCode) - 1 Step 1
                            tmCombineLstInfo(UBound(tmCombineLstInfo)).lDate = llDate
                            tmCombineLstInfo(UBound(tmCombineLstInfo)).tLst = tmLstCode(ilLst)
                            ReDim Preserve tmCombineLstInfo(0 To UBound(tmCombineLstInfo) + 1) As COMBINELSTINFO
                        Next ilLst
                    End If
                ElseIf (imLstExist) And (ilGenLST) Then
                    If ilCombineVefCode = 0 Then
                        '11/4/09: Re-add Generation of LST for Log Vehicles
                        'mDeleteUnusedLST
                        If (ilLSTForLogVeh > 0) And (tmVef.iVefCode > 0) Then
                            For ilLst = LBound(tmLstCode) To UBound(tmLstCode) - 1 Step 1
                                tmCombineLstInfo(UBound(tmCombineLstInfo)).lDate = llDate
                                tmCombineLstInfo(UBound(tmCombineLstInfo)).tLst = tmLstCode(ilLst)
                                ReDim Preserve tmCombineLstInfo(0 To UBound(tmCombineLstInfo) + 1) As COMBINELSTINFO
                            Next ilLst
                        End If
                    End If
                End If
            End If
            'create NTR entries only podcast vehicle (or podcast logvehicle)
            gCreateODFForNTR hlODF, hlChf, ilVehCode, tmZoneInfo()
        Next ilVeh
    '8/13/14: Generate spots for last date+1 into lst with Fed as *
    'Next llDate
        If llDate = llEDate + 1 Then
            If (bmCreatingLstDate1) Or (ilLNLocalAdj >= 0) Or (tgSpf.sGUseAffSys <> "Y") Or (slAdjLocalOrFeed = "F") Then
                Exit Do
            End If
            bmCreatingLstDate1 = True
        Else
            llDate = llDate + 1
        End If
    Loop
    '8/10/10:  Handle games that cross midnight
    If llGsfCode > 0 Then
        mDeleteUnusedLST slInLogType, ilExportType
    End If
    '6/4/16: Replaced GoSub
    'GoSub CloseFiles
    mCloseFiles hlCnf, hlCef, hlCTSsf, hlSsf, hlSdf, hlChf, hlODF, hlVef, hlVlf, hlDlf, hlTzf, hlCrf, hlSif, hlVsf, ilVehicle(), tlLLC()
    On Error Resume Next
    Erase tmBBSdfInfo
    Erase tmLstCode
    Erase llPrevOpenFdBBSpots
    Erase llPrevCloseFdBBSpots
    '6/12/09: Moved to Logs.frm because more then one game can be combined
    'If ilODFVefCode > 0 Then
    '    Erase tmCombineLstInfo
    'End If
    Erase tmDate1ODF
    gBuildODFSpotDay = True
    Exit Function
'lProcProg:
'    tmOdf.iUrfCode = tgUrf(0).iCode
'    If ilODFVefCode <= 0 Then
'        If ilSimVefCode <= 0 Then
'            If tmVef.iVefCode > 0 Then  'Log vehicle defined
'                tmOdf.iVefCode = tmVef.iVefCode
'            Else
'                tmOdf.iVefCode = ilVehCode
'            End If
'        Else
'            tmOdf.iVefCode = ilSimVefCode
'        End If
'    Else
'        tmOdf.iVefCode = ilODFVefCode
'    End If
'    If (tmVef.sType = "G") And (tgVpf(ilVpfIndex).sGenLog <> "M") Then
'        tmOdf.iGameNo = ilSSFType
'    Else
'        tmOdf.iGameNo = 0
'    End If
'    tmOdf.iAirDate(0) = ilLogDate0
'    tmOdf.iAirDate(1) = ilLogDate1
'    If tlLLC(ilLoop).sXMid = "Y" Then
'        smXMid = "Y"
'    Else
'        smXMid = "N"
'    End If
'
'    If tmDlf.iEtfCode = 1 Then
'        tmOdf.iAirTime(0) = tmProg.iStartTime(0)
'        tmOdf.iAirTime(1) = tmProg.iStartTime(1)
'    Else
'        tmOdf.iAirTime(0) = ilStartTime0
'        tmOdf.iAirTime(1) = ilStartTime1
'    End If
'    tmOdf.iLocalTime(0) = tmDlf.iLocalTime(0)
'    tmOdf.iLocalTime(1) = tmDlf.iLocalTime(1)
'    tmOdf.iFeedTime(0) = tmDlf.iFeedTime(0)
'    tmOdf.iFeedTime(1) = tmDlf.iFeedTime(1)
'    tmOdf.sZone = tmDlf.sZone
'    tmOdf.iEtfCode = tmDlf.iEtfCode
'    tmOdf.iEnfCode = tmDlf.iEnfCode
'    tmOdf.sProgCode = tmDlf.sProgCode
'    tmOdf.iMnfFeed = tmDlf.iMnfFeed
'    tmOdf.iDPSort = 0   '5-31-01 tmOdf.sUnused1 = "" 'tmDlf.sBus
'    'tmOdf.iWkNo = 0    'tmDlf.sSchedule
'    tmOdf.ianfCode = 0
'    tmOdf.iSortSeq = ilSortSeq                              '8-18-14 This field is required for C88, to keep programs with same back to back times apart.
'                                                            'gL14PageSkips & gSetSeqL29 subroutines modify this field for some other logs.  Do not call them for C88.
'    If ilType = 2 And tlLLC(ilLoop).sType = "A" Then         '1-10-01
'        tmOdf.ianfCode = Val(tlLLC(ilLoop).sName)
'    End If
'    tmOdf.iUnits = 0
'    If tmDlf.iEtfCode = 1 Then
'        gPackLength tlLLC(ilLoop).sLength, tmOdf.iLen(0), tmOdf.iLen(1)
'    Else
'        If (tmDlf.iEtfCode > 13) And (Trim$(tlLLC(ilLoop).sLength) <> "") Then
'            If gValidLength(tlLLC(ilLoop).sLength) Then
'                gPackLength tlLLC(ilLoop).sLength, tmOdf.iLen(0), tmOdf.iLen(1)
'            Else
'                tmOdf.iLen(0) = 1
'                tmOdf.iLen(1) = 0
'            End If
'        Else
'            tmOdf.iLen(0) = 1
'            tmOdf.iLen(1) = 0
'        End If
'    End If
'    tmOdf.lHd1CefCode = 0
'    tmOdf.lFt1CefCode = 0
'    tmOdf.lFt2CefCode = 0
'    tmOdf.iAlternateVefCode = 0 '1-17-14 was tmodfvefnmcefcode
'    tmOdf.iAdfCode = 0
'    tmOdf.lCifCode = 0
'    tmOdf.sProduct = ""
'    tmOdf.iMnfSubFeed = tmDlf.iMnfSubFeed
'    tmOdf.lCntrNo = 0
'    tmOdf.lchfcxfCode = 0
'    tmOdf.sDPDesc = ""
'    tmOdf.iRdfSortCode = 0
'    tmOdf.iBreakNo = 0
'    tmOdf.iPositionNo = 0
'    tmOdf.iType = ilType
'    tmOdf.lCefCode = llCefCode
'    tmOdf.lEvtIDCefCode = lmEvtIDCefCode
'    tmOdf.sDupeAvailID = tmDlf.sBus
'    tmOdf.lAvailcefCode = lmAvailCefCode            'comment from avail placed into spots
'    tmOdf.sShortTitle = ""
'    tmOdf.imnfSeg = 0               '6-19-01
'    tmOdf.sPageEjectFlag = "N"
'    'Determine seq number
'    '6/4/16: Replaced GoSub
'    'GoSub lProcAdjDate
'    mProcAdjDate slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType
'    '6/4/16: Replaced GoSub
'    'GoSub lProcSeqNo
'    mProcSeqNo ilSeqNo, slFor, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode
'    tmOdf.iSeqNo = ilSeqNo + 1
'    tmOdf.iDaySort = imDaySort
'    tmOdf.lEvtCefCode = lmEvtCefCode
'    tmOdf.iEvtCefSort = imEvtCefSort
'    tmOdf.sLogType = smLogType
'    tmOdf.sBBDesc = ""
'    '10/31/05:  Add zones to program and other events
'    'ilRet = btrInsert(hlOdf, tmOdf, imOdfRecLen, INDEXKEY0)
'    If ilDlfFound Then
'        '8/13/14: Generate spots for last date+1 into lst with Fed as *
'        If Not bmCreatingLstDate1 Then
'            tmOdf.lCode = 0
'            ilRet = btrInsert(hlODF, tmOdf, imOdfRecLen, INDEXKEY3)
'            gLogBtrError ilRet, "gBuildODFSpotDay: Insert #1"
'        End If
'    Else
'        ilOtherGen = False
'        For ilZone = 0 To UBound(tmZoneInfo) - 1 Step 1
'            Select Case tmZoneInfo(ilZone).sZone
'                Case "E"
'                    slZone = "EST"
'                Case "M"
'                    slZone = "MST"
'                Case "C"
'                    slZone = "CST"
'                Case "P"
'                    slZone = "PST"
'                Case Else
'                    If ilOtherGen Then
'                        Exit For
'                    End If
'                    slZone = ""
'            End Select
'            tmOdf.iAirDate(0) = ilLogDate0
'            tmOdf.iAirDate(1) = ilLogDate1
'            If tmDlf.iEtfCode = 1 Then
'                tmOdf.iAirTime(0) = tmProg.iStartTime(0)
'                tmOdf.iAirTime(1) = tmProg.iStartTime(1)
'                gUnpackTimeLong tmProg.iStartTime(0), tmProg.iStartTime(1), False, llAirTime
'            Else
'                tmOdf.iAirTime(0) = ilStartTime0
'                tmOdf.iAirTime(1) = ilStartTime1
'                gUnpackTimeLong ilStartTime0, ilStartTime1, False, llAirTime
'            End If
'            If (tgSpf.sGUseAffSys = "Y") Then
'                llTime = llAirTime + tmZoneInfo(ilZone).lGLocalAdj
'                If llTime < 0 Then
'                    llTime = llTime + 86400
'                ElseIf llTime > 86400 Then
'                    llTime = llTime - 86400
'                End If
'                gPackTimeLong llTime, tmDlf.iLocalTime(0), tmDlf.iLocalTime(1)
'                llTime = llAirTime + tmZoneInfo(ilZone).lGFeedAdj
'                If llTime < 0 Then
'                    llTime = llTime + 86400
'                ElseIf llTime > 86400 Then
'                    llTime = llTime - 86400
'                End If
'                gPackTimeLong llTime, tmDlf.iFeedTime(0), tmDlf.iFeedTime(1)
'            End If
'            tmOdf.iLocalTime(0) = tmDlf.iLocalTime(0)
'            tmOdf.iLocalTime(1) = tmDlf.iLocalTime(1)
'            tmOdf.iFeedTime(0) = tmDlf.iFeedTime(0)
'            tmOdf.iFeedTime(1) = tmDlf.iFeedTime(1)
'            tmOdf.sZone = slZone
'            '6/4/16: Replaced GoSub
'            'GoSub lProcAdjDate
'            mProcAdjDate slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType
'            '6/4/16: Replaced GoSub
'            'GoSub lProcSeqNo
'            mProcSeqNo ilSeqNo, slFor, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode
'            tmOdf.iSeqNo = ilSeqNo + 1
'            If (tgSpf.sGUseAffSys <> "Y") Or (imGenODF) Then
'                '8/13/14: Generate spots for last date+1 into lst with Fed as *
'                If Not bmCreatingLstDate1 Then
'                    tmOdf.lCode = 0
'                    ilRet = btrInsert(hlODF, tmOdf, imOdfRecLen, INDEXKEY0)
'                    gLogBtrError ilRet, "gBuildODFSpotDay: Insert #2"
'                End If
'            End If
'        Next ilZone
'    End If
'    Return

'    tmOdf.iUrfCode = tgUrf(0).iCode
'    If ilSimVefCode <= 0 Then
'        If tmVef.iVefCode > 0 Then  'Log vehicle defined
'            tmOdf.iVefCode = tmVef.iVefCode
'        Else
'            tmOdf.iVefCode = ilVehCode
'        End If
'    Else
'        tmOdf.iVefCode = ilSimVefCode
'    End If
'    tmOdf.iAirDate(0) = ilLogDate0
'    tmOdf.iAirDate(1) = ilLogDate1
'    tmOdf.iAirTime(0) = tmAvail.iTime(0)
'    tmOdf.iAirTime(1) = tmAvail.iTime(1)
'    tmOdf.iLocalTime(0) = tmDlf.iLocalTime(0)
'    tmOdf.iLocalTime(1) = tmDlf.iLocalTime(1)
'    tmOdf.iFeedTime(0) = tmDlf.iFeedTime(0)
'    tmOdf.iFeedTime(1) = tmDlf.iFeedTime(1)
'    tmOdf.sZone = tmDlf.sZone
'    tmOdf.iEtfCode = tmDlf.iEtfCode
'    tmOdf.iEnfCode = tmDlf.iEnfCode
'    tmOdf.sProgCode = tmDlf.sProgCode
'    tmOdf.iMnfFeed = tmDlf.iMnfFeed
'    tmOdf.iDPSort = 0   '5-31-01tmOdf.sUnused1 = "" 'tmDlf.sBus
'    'tmOdf.iWkNo = 0    'tmDlf.sSchedule
'    tmOdf.ianfCode = tmAvail.ianfCode
'    tmOdf.iUnits = tmAvail.iAvInfo And &H1F
'    slLength = Trim$(str$(tmAvail.iLen)) & "s"
'    gPackLength slLength, tmOdf.iLen(0), tmOdf.iLen(1)
'    tmOdf.iAdfCode = 0
'    tmOdf.lCifCode = 0
'    tmOdf.sProduct = ""
'    tmOdf.iMnfSubFeed = tmDlf.iMnfSubFeed
'    tmOdf.lCntrNo = 0
'    tmOdf.lchfcxfCode = 0
'    tmOdf.iRdfSortCode = 0
'    tmOdf.sDPDesc = ""
'    tmOdf.iBreakNo = imBreakNo
'    tmOdf.iPositionNo = 0
'    tmOdf.iType = ilType
'    tmOdf.lCefCode = llCefCode
'    tmOdf.lEvtIDCefCode = lmEvtIDCefCode
'    tmOdf.sDupeAvailID = tmDlf.sBus
'    tmOdf.sShortTitle = ""
'    tmOdf.imnfSeg = 0               '6-19-01
'    tmOdf.sPageEjectFlag = "N"
'    'Determine seq number
'    '6/4/16: Replaced GoSub
'    'GoSub lProcAdjDate
'    mProcAdjDate slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType
'    '6/4/16: Replaced GoSub
'    'GoSub lProcSeqNo
'    mProcSeqNo ilSeqNo, slFor, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode
'    tmOdf.iSeqNo = ilSeqNo + 1
'    tmOdf.iDaySort = imDaySort
'    tmOdf.lEvtCefCode = lmEvtCefCode
'    tmOdf.iEvtCefSort = imEvtCefSort
'    tmOdf.sLogType = smLogType
'    'ilRet = btrInsert(hlOdf, tmOdf, imOdfRecLen, INDEXKEY0)
'    If (tgSpf.sCBlackoutLog = "Y") And (imLstExist) Then
'        'Added when changed to generate Affiliate with each Conventional Vehicle instead of
'        'the Log Vehicle.  11/20/03.  Passed vefCode instead of getting it from tmOdf.iVefCode
'        If ilSimVefCode <= 0 Then
'            '11/4/09: Re-add generation of lst by Log vehicle
'            'mCreateAvailLst ilVehCode, tmAvail.iLen, tmOdf.iUnits, llGsfCode
'            If (ilLSTForLogVeh > 0) And (tmVef.iVefCode > 0) Then
'                mCreateAvailLst tmVef.iVefCode, tmAvail.iLen, tmOdf.iUnits, llGsfCode
'            Else
'                mCreateAvailLst ilVehCode, tmAvail.iLen, tmOdf.iUnits, llGsfCode
'            End If
'        Else
'            mCreateAvailLst ilSimVefCode, tmAvail.iLen, tmOdf.iUnits, llGsfCode
'        End If
'        'End If Change
'    End If
'    Return
'lChkOpenAvail:
'    If (tgSpf.sCBlackoutLog <> "Y") Or (Not imLstExist) Or (Not ilGenLST) Then
'        Return
'    End If
'
'    ilAvVpfIndex = -1
'    'For ilVpf = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
'    '    If ilAvVefCode = tgVpf(ilVpf).iVefKCode Then
'        ilVpf = gBinarySearchVpf(ilAvVefCode)
'        If ilVpf <> -1 Then
'            ilAvVpfIndex = ilVpf
'    '        Exit For
'        End If
'    'Next ilVpf
'    If ilAvVpfIndex < 0 Then
'        Return
'    End If
'    If (tgVpf(ilAvVpfIndex).sSSellOut = "B") Or (tgVpf(ilAvVpfIndex).sSSellOut = "U") Or (tgVpf(ilAvVpfIndex).sSSellOut = "M") Then
'        ilUnits = tmAvAvail.iAvInfo And &H1F
'        slUnits = Trim$(str$(ilUnits)) & ".0"   'For units as thirty
'        ilSec = tmAvAvail.iLen
'    Else
'        ilUnits = tmAvAvail.iAvInfo And &H1F
'        ilSec = 0
'    End If
'    For ilSpot = 1 To tmAvAvail.iNoSpotsThis Step 1
'        tmAvSpot = tmCTSsf.tPas(ADJSSFPASBZ + ilAvEvt + ilSpot)
'        If (tmAvSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
'            If (tgVpf(ilAvVpfIndex).sSSellOut = "B") Or (tgVpf(ilAvVpfIndex).sSSellOut = "U") Then
'                ilUnits = ilUnits - 1
'                ilSec = ilSec - (tmAvSpot.iPosLen And &HFFF)
'            ElseIf tgVpf(ilAvVpfIndex).sSSellOut = "M" Then
'                ilUnits = ilUnits - 1
'                ilSec = ilSec - (tmAvSpot.iPosLen And &HFFF)
'            ElseIf tgVpf(ilAvVpfIndex).sSSellOut = "T" Then
'                slSpotLen = Trim$(str$(tmAvSpot.iPosLen And &HFFF))
'                slStr = gDivStr(slSpotLen, "30.0")
'                slUnits = gSubStr(slUnits, slSpotLen)
'            End If
'        End If
'    Next ilSpot
'    If (tgVpf(ilAvVpfIndex).sSSellOut = "B") Or (tgVpf(ilAvVpfIndex).sSSellOut = "U") Or (tgVpf(ilAvVpfIndex).sSSellOut = "M") Then
'        If (ilUnits > 0) And (ilSec > 0) Then
'            '6/5/16: Replaced GoSub
'            'GoSub lGetDlfAvail
'            mGetDlfAvail ilDlfFound, slDay, ilLoop, tlLLC(), hlDlf, ilDlfDate0, ilDlfDate1, ilODFVefCode, ilSimVefCode, ilVpfIndex, ilSSFType, ilLogDate0, ilLogDate1, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlOdf, tlOdf, ilVehCode, ilGenLST, ilLSTForLogVeh, ilSec, ilUnits, llGsfCode, ilOtherGen, ilTerminated
'        End If
'    ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
'        If gCompNumberStr(slUnits, "0.0") > 0 Then
'            ilSec = Val(slUnits)
'            '6/5/16: Replaced GoSub
'            'GoSub lGetDlfAvail
'            mGetDlfAvail ilDlfFound, slDay, ilLoop, tlLLC(), hlDlf, ilDlfDate0, ilDlfDate1, ilODFVefCode, ilSimVefCode, ilVpfIndex, ilSSFType, ilLogDate0, ilLogDate1, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlOdf, tlOdf, ilVehCode, ilGenLST, ilLSTForLogVeh, ilSec, ilUnits, llGsfCode, ilOtherGen, ilTerminated
'        End If
'    End If
'    Return
'lGetDlfAvail:
'    If (ilDlfFound) Then
'        'Obtain delivery entry to see is avail is sent
'        tmDlfSrchKey.iVefCode = ilVehCode
'        tmDlfSrchKey.sAirDay = slDay
'        tmDlfSrchKey.iStartDate(0) = ilDlfDate0
'        tmDlfSrchKey.iStartDate(1) = ilDlfDate1
'        tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0) 'tmAvAvail.iTime(0)
'        tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1) 'tmAvAvail.iTime(1)
'        ilRet = btrGetEqual(hlDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
'        Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvAvail.iTime(1))
'            ilTerminated = False
'            If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
'                ilTerminated = True
'            Else
'                If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
'                    If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
'                        ilTerminated = True
'                    End If
'                End If
'            End If
'            If Not ilTerminated Then
'                If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
'                    'If slFor = "C" Then
'                        If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0) Then
'                            tmDlf.iMnfFeed = 0
'                            '6/5/16: Replaced GoSub
'                            'GoSub lMakeAvails
'                            mMakeAvails ilODFVefCode, ilSimVefCode, ilVpfIndex, ilSSFType, ilLogDate0, ilLogDate1, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlOdf, tlOdf, ilVehCode, ilGenLST, ilLSTForLogVeh, ilSec, ilUnits, llGsfCode
'                        End If
'                    'End If
'                End If
'            End If
'            ilRet = btrGetNext(hlDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'    Else
'        For ilZone = 0 To UBound(tmZoneInfo) - 1 Step 1
'            If ((tmZoneInfo(ilZone).sFed = "*") And (tgSpf.sGUseAffSys = "Y")) Or (tgSpf.sGUseAffSys <> "Y") Then
'                Select Case tmZoneInfo(ilZone).sZone
'                    Case "E"
'                        slZone = "EST"
'                    Case "M"
'                        slZone = "MST"
'                    Case "C"
'                        slZone = "CST"
'                    Case "P"
'                        slZone = "PST"
'                    Case Else
'                        If ilOtherGen Then
'                            Exit For
'                        End If
'                        slZone = ""
'                End Select
'                If (tgSpf.sGUseAffSys = "Y") Then
'                    'gUnpackTimeLong tmAvAvail.iTime(0), tmAvAvail.iTime(1), False, llAirTime
'                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llAirTime
'                    llTime = llAirTime + tmZoneInfo(ilZone).lGLocalAdj
'                    If llTime < 0 Then
'                        llTime = llTime + 86400
'                    ElseIf llTime > 86400 Then
'                        llTime = llTime - 86400
'                    End If
'                    gPackTimeLong llTime, tmDlf.iLocalTime(0), tmDlf.iLocalTime(1)
'                    llTime = llAirTime + tmZoneInfo(ilZone).lGFeedAdj
'                    If llTime < 0 Then
'                        llTime = llTime + 86400
'                    ElseIf llTime > 86400 Then
'                        llTime = llTime - 86400
'                    End If
'                    gPackTimeLong llTime, tmDlf.iFeedTime(0), tmDlf.iFeedTime(1)
'                Else
'                    'tmDlf.iLocalTime(0) = tmAvAvail.iTime(0)
'                    'tmDlf.iLocalTime(1) = tmAvAvail.iTime(1)
'                    'tmDlf.iFeedTime(0) = tmAvAvail.iTime(0)
'                    'tmDlf.iFeedTime(1) = tmAvAvail.iTime(1)
'                    tmDlf.iLocalTime(0) = tmAvail.iTime(0)
'                    tmDlf.iLocalTime(1) = tmAvail.iTime(1)
'                    tmDlf.iFeedTime(0) = tmAvail.iTime(0)
'                    tmDlf.iFeedTime(1) = tmAvail.iTime(1)
'                End If
'                tmDlf.sZone = slZone
'                tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode
'                tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode
'                tmDlf.sProgCode = ""
'                tmDlf.iMnfFeed = 0
'                tmDlf.sBus = ""
'                tmDlf.sSchedule = ""
'                tmDlf.iMnfSubFeed = 0
'                '6/5/16: Replaced GoSub
'                'GoSub lMakeAvails
'                mMakeAvails ilODFVefCode, ilSimVefCode, ilVpfIndex, ilSSFType, ilLogDate0, ilLogDate1, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlOdf, tlOdf, ilVehCode, ilGenLST, ilLSTForLogVeh, ilSec, ilUnits, llGsfCode
'            End If
'        Next ilZone
'    End If
'    Return
'lMakeAvails:
'    If ilODFVefCode <= 0 Then
'        If ilSimVefCode <= 0 Then
'            If tmVef.iVefCode > 0 Then  'Log vehicle defined
'                tmOdf.iVefCode = tmVef.iVefCode
'            Else
'                tmOdf.iVefCode = ilVehCode
'            End If
'        Else
'            tmOdf.iVefCode = ilSimVefCode
'        End If
'    Else
'        tmOdf.iVefCode = ilODFVefCode
'    End If
'    If (tmVef.sType = "G") And (tgVpf(ilVpfIndex).sGenLog <> "M") Then
'        tmOdf.iGameNo = ilSSFType
'    Else
'        tmOdf.iGameNo = 0
'    End If
'    tmOdf.iAirDate(0) = ilLogDate0
'    tmOdf.iAirDate(1) = ilLogDate1
'    If (tmAvail.iAvInfo And SSXMID) = SSXMID Then
'        smXMid = "Y"
'    Else
'        smXMid = "N"
'    End If
'    tmOdf.iAirTime(0) = tmAvail.iTime(0)  'tmAvAvail.iTime(0)
'    tmOdf.iAirTime(1) = tmAvail.iTime(1)  'tmAvAvail.iTime(1)
'    tmOdf.iLocalTime(0) = tmDlf.iLocalTime(0)
'    tmOdf.iLocalTime(1) = tmDlf.iLocalTime(1)
'    tmOdf.iFeedTime(0) = tmDlf.iFeedTime(0)
'    tmOdf.iFeedTime(1) = tmDlf.iFeedTime(1)
'    tmOdf.sZone = tmDlf.sZone
'    tmOdf.iBreakNo = imBreakNo
'    tmOdf.iPositionNo = imPositionNo
'    tmOdf.ianfCode = tmAvail.ianfCode
'    '6/4/16: Replaced GoSub
'    'GoSub lProcAdjDate
'    mProcAdjDate slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType
'    '6/4/16: Replaced GoSub
'    'GoSub lProcSeqNo
'    mProcSeqNo ilSeqNo, slFor, hlOdf, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode
'    tmOdf.iSeqNo = ilSeqNo + 1
'    'imGenODF set for affiliate system
'    If (tgSpf.sGUseAffSys <> "Y") Or (imGenODF) Then
'        If imLstExist And ilGenLST Then
'            For ilTZone = 0 To UBound(tmZoneInfo) - 1 Step 1
'                If Left$(tmOdf.sZone, 1) = tmZoneInfo(ilTZone).sZone Then
'                    If tmZoneInfo(ilTZone).sFed = "*" Then
'                        'Added when changed to generate Affiliate with each Conventional Vehicle instead of
'                        'the Log Vehicle.  11/20/03.  Passed vefCode instead of getting it from tmOdf.iVefCode
'                        '1/11/08:  Match how odf is created for combination games (remove game test). i.e. produce lst as a combination of game vehicles
'                        'If (ilODFVefCode <= 0) Or (tmVef.sType = "G") Then
'                        If (ilODFVefCode <= 0) Then
'                            If ilSimVefCode <= 0 Then
'                                '11/4/09: Re-add generation of lst by Log vehicle
'                                'mCreateAvailLst ilVehCode, ilSec, ilUnits, llGsfCode
'                                If (ilLSTForLogVeh > 0) And (tmVef.iVefCode > 0) Then
'                                    mCreateAvailLst tmVef.iVefCode, ilSec, ilUnits, llGsfCode
'                                Else
'                                    mCreateAvailLst ilVehCode, ilSec, ilUnits, llGsfCode
'                                End If
'                            Else
'                                mCreateAvailLst ilSimVefCode, ilSec, ilUnits, llGsfCode
'                            End If
'                        Else
'                            mCreateAvailLst ilODFVefCode, ilSec, ilUnits, llGsfCode
'                        End If
'                        'End of Change
'                    End If
'                    Exit For
'                End If
'            Next ilTZone
'        End If
'    End If
'    Return
'lProcSpot:              '1/17/99 added DP description & sort code, & cntr hdr Other comment code
'    '8/13/14: Generate spots for last date+1 into lst with Fed as *
'    If bmCreatingLstDate1 Then
'        For ilTZone = 0 To UBound(tmZoneInfo) - 1 Step 1
'            If Left$(tmDlf.sZone, 1) = tmZoneInfo(ilTZone).sZone Then
'                If tmZoneInfo(ilTZone).sFed <> "*" Then
'                    Return
'                End If
'            End If
'        Next ilTZone
'        lmGLocalAdj = 0
'        For ilTZone = 0 To UBound(tmZoneInfo) - 1 Step 1
'            If Left$(tmDlf.sZone, 1) = tmZoneInfo(ilTZone).sFed Then
'                If tmZoneInfo(ilTZone).lGLocalAdj < lmGLocalAdj Then
'                    lmGLocalAdj = tmZoneInfo(ilTZone).lGLocalAdj
'                End If
'            End If
'        Next ilTZone
'        If lmGLocalAdj >= 0 Then
'            Return
'        End If
'        lmGLocalAdj = -lmGLocalAdj
'    End If
'    If ((Asc(tgSpf.sUsingFeatures6) And BBNOTSEPARATELINE) = BBNOTSEPARATELINE) Then
'        If (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
'            Return
'        End If
'    End If
'    If (tmAvail.iAvInfo And SSXMID) = SSXMID Then
'        smXMid = "Y"
'    Else
'        smXMid = "N"
'    End If
'    tmOdf.iUrfCode = tgUrf(0).iCode
'    If ilODFVefCode <= 0 Then
'        If ilSimVefCode <= 0 Then
'            If tmVef.iVefCode > 0 Then  'Log vehicle defined
'                tmOdf.iVefCode = tmVef.iVefCode
'                tmOdf.iAlternateVefCode = ilVehCode     '1-15-14 for log vehicles, put name in this field to separate the vehicles on output
'            Else
'                tmOdf.iVefCode = ilVehCode
'                tmOdf.iAlternateVefCode = ilVehCode
'            End If
'        Else
'            tmOdf.iVefCode = ilSimVefCode
'            tmOdf.iAlternateVefCode = ilSimVefCode
'        End If
'    Else
'        tmOdf.iVefCode = ilODFVefCode
'        tmOdf.iAlternateVefCode = ilODFVefCode
'    End If
'    If (tmVef.sType = "G") And (tgVpf(ilVpfIndex).sGenLog <> "M") Then
'        tmOdf.iGameNo = ilSSFType
'    Else
'        tmOdf.iGameNo = 0
'    End If
'    tmOdf.iAirDate(0) = ilLogDate0
'    tmOdf.iAirDate(1) = ilLogDate1
'    tmOdf.iAirTime(0) = tmAvail.iTime(0)
'    tmOdf.iAirTime(1) = tmAvail.iTime(1)
'    tmOdf.iLocalTime(0) = tmDlf.iLocalTime(0)
'    tmOdf.iLocalTime(1) = tmDlf.iLocalTime(1)
'    tmOdf.iFeedTime(0) = tmDlf.iFeedTime(0)
'    tmOdf.iFeedTime(1) = tmDlf.iFeedTime(1)
'    tmOdf.sZone = tmDlf.sZone
'    tmOdf.iEtfCode = 0
'    tmOdf.iEnfCode = tmDlf.iEnfCode 'Required for Commercial schedule
'    tmOdf.sProgCode = tmDlf.sProgCode
'    tmOdf.iMnfFeed = tmDlf.iMnfFeed
'    tmOdf.iDPSort = 0   '5-31-01 tmOdf.sUnused1 = "" 'tmDlf.sBus
'    'tmOdf.iWkNo = 0    'tmDlf.sSchedule
'    tmOdf.ianfCode = tmAvail.ianfCode
'    tmOdf.iUnits = 0
'    slLength = Trim$(str$(tmSdf.iLen)) & "s"
'    gPackLength slLength, tmOdf.iLen(0), tmOdf.iLen(1)
'    tmOdf.sBBDesc = ""
'    If ((Asc(tgSpf.sUsingFeatures6) And BBNOTSEPARATELINE) = BBNOTSEPARATELINE) Then
'        If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
'            If tmClf.iBBOpenLen > 0 Then
'                tmOdf.sBBDesc = "BB"
'            End If
'        Else
'            If (tmClf.iBBOpenLen > 0) And (tmClf.iBBCloseLen > 0) Then
'                tmOdf.sBBDesc = "O/C BB"
'            ElseIf tmClf.iBBOpenLen > 0 Then
'                tmOdf.sBBDesc = "O BB"
'            ElseIf tmClf.iBBCloseLen > 0 Then
'                tmOdf.sBBDesc = "C BB"
'            End If
'        End If
'    Else
'        If (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
'            If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
'                tmOdf.sBBDesc = "BB"
'            Else
'                If tmSdf.sSpotType = "O" Then
'                    tmOdf.sBBDesc = "O BB"
'                ElseIf tmSdf.sSpotType = "C" Then
'                    tmOdf.sBBDesc = "C BB"
'                End If
'            End If
'        End If
'    End If
'    tmOdf.iAdfCode = tmChf.iAdfCode
'    'Test tmSdf.sPtType
'    tmOdf.lCifCode = llCifCode
'    'If Trim$(tmChf.sProduct) <> "" Then
'    '    tmOdf.sProduct = tmChf.sProduct
'    'Else
'    '    tmOdf.sProduct = "" '"???? Name ????"
'    'End If
'    mGetCpf hmCif, llCifCode
'    If Trim$(tmCpf.sName) <> "" Then
'        tmOdf.sProduct = tmCpf.sName
'    Else
'        If Trim$(tmChf.sProduct) <> "" Then
'            tmOdf.sProduct = tmChf.sProduct
'        Else
'            tmOdf.sProduct = "" '"???? Name ????"
'        End If
'    End If
'    If tmChf.lCode = 0 Then                 'feed spot, indicate it with advt/prod
'        tmOdf.sProduct = Trim$(tmOdf.sProduct) & " (Feed)"
'    End If
'
'    tmOdf.iMnfSubFeed = tmDlf.iMnfSubFeed
'    tmOdf.lCntrNo = tmChf.lCntrNo
'
'    '2-15-01 Setup comment pointers only if show = yes on Log
'    tmOdf.lchfcxfCode = 0                     'assume no comment
'    If tmChf.lCxfCode > 0 Then
'        tlCxfSrchKey.lCode = tmChf.lCxfCode      'comment  code
'        ilCxfRecLen = Len(tlCxf)
'        ilRet = btrGetEqual(hmCxf, tlCxf, ilCxfRecLen, tlCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching comment recd
'        If ilRet = BTRV_ERR_NONE Then
'            If tlCxf.sShSpot = "Y" Then         'show comment on log
'                tmOdf.lchfcxfCode = tmChf.lCxfCode
'            End If
'        End If
'    End If
'    If slAdjLocalOrFeed = "W" Then
'        tmOdf.lFt1CefCode = tmChf.lCode
'    End If
'    'tmOdf.lChfCxfCode = tmChf.lCxfCode      '2-15-01, only place if showing on log-Other comment code (1/17/99)
'    If (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
'        tmOdf.iBreakNo = 0
'    Else
'        tmOdf.iBreakNo = imBreakNo
'    End If
'    tmOdf.iPositionNo = imPositionNo
'    tmOdf.iType = ilType
'    tmOdf.lCefCode = lmCrfCsfCode   '0
'    tmOdf.sBonus = smBonus          '2-15-01 Bonus flag (B= bonus, f = fill)
'    tmOdf.lAvailcefCode = lmAvailCefCode            'comment ptr from avail
'    If tlAdf.iCode <> tmChf.iAdfCode Then
'        tlAdfSrchKey.iCode = tmChf.iAdfCode
'        ilRet = btrGetEqual(hmAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'    End If
'    tmOdf.lEvtIDCefCode = lmEvtIDCefCode
'    tmOdf.sDupeAvailID = tmDlf.sBus
'    tmOdf.sShortTitle = gGetShortTitle(hlVsf, hmClf, hlSif, tmChf, tlAdf, tmSdf)
'    tmOdf.imnfSeg = tmChf.imnfSeg       '6-19-01
'    tmOdf.sPageEjectFlag = "N"
'    For ilPE = LBound(tmPageEject) To UBound(tmPageEject) - 1 Step 1
'        If (lmEvtTime >= tmPageEject(ilPE).lTime) And ((tmPageEject(ilPE).ianfCode = tmAvail.ianfCode) Or (tmPageEject(ilPE).ianfCode = 0)) Then
'            tmOdf.sPageEjectFlag = "Y"
'            tmPageEject(ilPE).lTime = 999999
'            tmPageEject(ilPE).ianfCode = -1
'            Exit For
'        End If
'    Next ilPE
'    '11-11-09 show the line comment on logs that are coded for it.  Dual use of the field OdfEvtCefCode .  A spot record will contain the line comment ptr
'    mDPDaysTimes hmRdf, smEDIDays, lmEvtCefCode             'Read clf & cff & rdf to format the DP description (or sch line override)
'    tmOdf.lRafCode = 0              '5-16-08
'    tmOdf.sSplitNetwork = "N"
'    '10/23/13: Check for split network, then split copy
'    '10/24/13: added test if using split network
'    'split networks
'    If (Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS Then
'        If (tmSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
'            tmOdf.sSplitNetwork = "P"
'            tmOdf.lRafCode = tmClf.lRafCode
'        ElseIf (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
'            tmOdf.sSplitNetwork = "S"
'            tmOdf.lRafCode = tmClf.lRafCode
'        End If
'    End If
'    slSplitCopyFlag = ""
'    'If (Asc(tgSpf.sUsingFeatures2) And SPLITCOPY = SPLITCOPY) Then
'    If ((Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) = SPLITCOPY) And (tmOdf.sSplitNetwork = "N") Then
'        tmRsfSrchKey1.lCode = tmSdf.lCode
'        ilRet = btrGetEqual(hmRsf, tmRsf, Len(tmRsf), tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
'        Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
'            If (tmRsf.sType <> "A") Then
'                'tmOdf.sSplitNetwork = "S"
'                slSplitCopyFlag = "S"
'                Exit Do
'            End If
'            ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'    '10/23/13: Moved above Split Copy
'    'Else
'    '    'split networks
'    '    If (tmSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
'    '        tmOdf.sSplitNetwork = "P"
'    '        tmOdf.lRafCode = tmClf.lRafCode
'    '    ElseIf (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
'    '        tmOdf.sSplitNetwork = "S"
'    '        tmOdf.lRafCode = tmClf.lRafCode
'    '    End If
'    End If
'
'    'Same code in gGetShortTitle- it is here to save execution time
'    'If tgSpf.sUseProdSptScr = "P" Then
'    '    If llSifCode <= 0 Then  'llSifCode obtained from Crf in mObtainCrfCsfCode
'    '        llSifCode = tmChf.lSifCode
'    '    End If
'    '    If llSifCode > 0 Then
'    '        tlSifSrchKey.lCode = llSifCode
'    '        ilRet = btrGetEqual(hlSif, tlSif, ilSifRecLen, tlSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'    '        If ilRet = BTRV_ERR_NONE Then
'    '            tmOdf.sShortTitle = tlSif.sName
'    '        Else
'    '            tmOdf.sShortTitle = ""  '"???? Name ????"
'    '        End If
'    '    Else
'    '        tmOdf.sShortTitle = ""  '"???? Name ????"
'    '    End If
'    'Else
'    '    If tlAdf.iCode <> tmChf.iAdfCode Then
'    '        tlAdfSrchKey.iCode = tmChf.iAdfCode
'    '        ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
'    '    Else
'    '        ilRet = BTRV_ERR_NONE
'    '    End If
'    '    If ilRet = BTRV_ERR_NONE Then
'    '        tmOdf.sShortTitle = Trim$(tlAdf.sAbbr) & "," & Trim$(tmChf.sProduct)
'    '    Else
'    '        tmOdf.sShortTitle = ""
'    '    End If
''    'End If
'    'Determine seq number
'    '6/4/16: Replaced GoSub
'    'GoSub lProcTestForDuplBB
'    mProcTestForDuplBB ilAddSpot
'    If Not ilAddSpot Then
'        Return
'    End If
'    '6/4/16: Replaced GoSub
'    'GoSub lProcAdjDate
'    mProcAdjDate slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType
'    '6/4/16: Replaced GoSub
'    'GoSub lProcSeqNo
'    mProcSeqNo ilSeqNo, slFor, hlOdf, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode
'
'    tmOdf.iSortSeq = 0                      'ensure all the spots are not separated
'    tmOdf.iSeqNo = ilSeqNo + 1
'    tmOdf.iDaySort = imDaySort
'    tmOdf.lEvtCefCode = lmEvtCefCode
'    tmOdf.iEvtCefSort = imEvtCefSort
'    tmOdf.sLogType = smLogType
''    tmOdf.sSplitNetwork = "N"
''    tmOdf.lRafCode = 0              '5-16-08
''    If (tmSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
''        tmOdf.sSplitNetwork = "P"
''        tmOdf.lRafCode = tmClf.lRafCode
''    ElseIf (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
''        tmOdf.sSplitNetwork = "S"
''        tmOdf.lRafCode = tmClf.lRafCode
''    End If
'
'    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, lmAvailTime
'    If lmPrevAvailTime <> lmAvailTime Then
'        tmOdf.iAvailLen = tmAvail.iLen
'        lmPrevAvailTime = lmAvailTime
'        'BB avail
'        If tmAvail.iLen = 0 Then
'            tmOdf.iAvailLen = -1
'        End If
'    Else
'        tmOdf.iAvailLen = 0
'    End If
'    tmOdf.sAvailLock = "N"
'    '7-25-13 combine 2 definitions into this field:  Locked avails/spots + split copy defined
'    If ((tmAvail.iAvInfo And SSLOCK) = SSLOCK) And ((tmAvail.iAvInfo And SSLOCKSPOT) = SSLOCKSPOT) Then
'        tmOdf.sAvailLock = "B"
'        If slSplitCopyFlag = "S" Then
'            tmOdf.sAvailLock = "D"              'locked avail & spot plus a split copy
'        End If
'    ElseIf ((tmAvail.iAvInfo And SSLOCK) = SSLOCK) Then
'        tmOdf.sAvailLock = "A"
'        If slSplitCopyFlag = "S" Then
'            tmOdf.sAvailLock = "E"              'locked avail  plus a split copy
'        End If
'    ElseIf ((tmAvail.iAvInfo And SSLOCKSPOT) = SSLOCKSPOT) Then
'        tmOdf.sAvailLock = "S"
'        If slSplitCopyFlag = "S" Then
'            tmOdf.sAvailLock = "E"              'locked  spot plus a split copy
'        End If
'    Else
'        If slSplitCopyFlag = "S" Then
'            tmOdf.sAvailLock = "F"              'nothing locked, but split copy
'        End If
'    End If
'    'imGenODF set for affiliate system
'    If (tgSpf.sGUseAffSys <> "Y") Or (imGenODF) Then
'        If imLstExist And ilGenLST Then
'            If (tmSdf.sSpotType <> "O") Or (tmSdf.sSpotType <> "C") Or ((tmSdf.sSpotType = "O") And (tgSpf.sBBsToAff = "Y")) Or ((tmSdf.sSpotType = "C") And (tgSpf.sBBsToAff = "Y")) Then
'                For ilTZone = 0 To UBound(tmZoneInfo) - 1 Step 1
'                    If Left$(tmOdf.sZone, 1) = tmZoneInfo(ilTZone).sZone Then
'                        If tmZoneInfo(ilTZone).sFed = "*" Then
'                            'Added when changed to generate Affiliate with each Conventional Vehicle instead of
'                            'the Log Vehicle.  11/20/03.  Passed vefCode instead of getting it from tmOdf.iVefCode
'                            '1/11/08:  Match how odf is created for combination games (remove game test). i.e. produce lst as a combination of game vehicles
'                            'If (ilODFVefCode <= 0) Or (tmVef.sType = "G") Then
'                            If (ilODFVefCode <= 0) Then
'                                If ilSimVefCode <= 0 Then
'                                    '11/4/09: re-add generation of LST by Log vehicle
'                                    'mCreateLst ilVehCode, ilClearLstSdf, hmClf, hmCif, hlMcf, slLogType, ilCreateNewLST, ilExportType, llGsfCode, ilWegenerOLA
'                                    If (ilLSTForLogVeh > 0) And (tmVef.iVefCode > 0) Then
'                                        mCreateLst tmVef.iVefCode, imClearLstSdf, hmClf, hmCif, hlMcf, slInLogType, imCreateNewLST, ilExportType, llGsfCode, imWegenerOLA, smAlertStatus
'                                    Else
'                                        mCreateLst ilVehCode, imClearLstSdf, hmClf, hmCif, hlMcf, slInLogType, imCreateNewLST, ilExportType, llGsfCode, imWegenerOLA, smAlertStatus
'                                    End If
'                                Else
'                                    mCreateLst ilSimVefCode, imClearLstSdf, hmClf, hmCif, hlMcf, slInLogType, imCreateNewLST, ilExportType, llGsfCode, imWegenerOLA, smAlertStatus
'                                End If
'                            Else
'                                mCreateLst ilODFVefCode, imClearLstSdf, hmClf, hmCif, hlMcf, slInLogType, imCreateNewLST, ilExportType, llGsfCode, imWegenerOLA, smAlertStatus
'                            End If
'                            'End of Change.
'
'                            ' 3-25-10 remove overriding short title with cart (thats only in affiliate system)
'                            ' L31A was created on Affiliate side to show the short title, change L31 back to using the copy pointers
'                            ' due to Music of your Life import, copy must be obtained from
'                            ' the short title field as text rather than using a copy pointer
'                            'tmOdf.sShortTitle = tmLst.sCart
'                        End If
'                        Exit For
'                    End If
'                Next ilTZone
'            End If
'        End If
'        '8/13/14: Generate spots for last date+1 into lst with Fed as *
'        If Not bmCreatingLstDate1 Then
'            tmOdf.lCode = 0
'            ilRet = btrInsert(hlOdf, tmOdf, imOdfRecLen, INDEXKEY3)
'            tgOdfSdfCodes(UBound(tgOdfSdfCodes)).lOdfCode = tmOdf.lCode
'            tgOdfSdfCodes(UBound(tgOdfSdfCodes)).lSdfCode = tmSdf.lCode
'            ReDim Preserve tgOdfSdfCodes(0 To UBound(tgOdfSdfCodes) + 1) As ODFSDFCODES
'            gLogBtrError ilRet, "gBuildODFSpotDay: Insert #3"
'            If (tgSpf.sCBlackoutLog = "Y") Or (igBkgdProg = 3) Then
'                llUpper = UBound(tgSpotSum)
'                tgSpotSum(llUpper).iVefCode = tmOdf.iVefCode
'                gUnpackDateLong tmOdf.iAirDate(0), tmOdf.iAirDate(1), tgSpotSum(llUpper).lDate
'                tgSpotSum(llUpper).lChfCode = tmChf.lCode
'                tgSpotSum(llUpper).iMnfComp(0) = tmChf.iMnfComp(0)
'                tgSpotSum(llUpper).iMnfComp(1) = tmChf.iMnfComp(1)
'                tgSpotSum(llUpper).iAdfCode = tmChf.iAdfCode
'                tgSpotSum(llUpper).iLen = tmSdf.iLen
'                tgSpotSum(llUpper).sProduct = tmChf.sProduct
'                If Not imLstExist Or Not ilGenLST Then
'                    tgSpotSum(llUpper).sShortTitle = tmOdf.sShortTitle
'                End If
'                tgSpotSum(llUpper).imnfSeg = tmChf.imnfSeg          '6-19-01
'                gUnpackTimeLong tmOdf.iLocalTime(0), tmOdf.iLocalTime(1), False, tgSpotSum(llUpper).lTime
'                tgSpotSum(llUpper).sZone = tmOdf.sZone   'tmDlf.sZone
'                tgSpotSum(llUpper).iSeqNo = tmOdf.iSeqNo
'                If imLstExist Then
'                    tgSpotSum(llUpper).lLstCode = tmLst.lCode
'                Else
'                    tgSpotSum(llUpper).lLstCode = 0
'                End If
'                tgSpotSum(llUpper).iLnVefCode = tmClf.iVefCode
'                tgSpotSum(llUpper).lSdfCode = tmSdf.lCode
'                tgSpotSum(llUpper).sLogType = smLogType
'                tgSpotSum(llUpper).lCrfCode = llCrfCode
'                tgSpotSum(llUpper).sDays = smEDIDays
'                tgSpotSum(llUpper).iOrigAirDate(0) = ilLogDate0
'                tgSpotSum(llUpper).iOrigAirDate(1) = ilLogDate1
'                ReDim Preserve tgSpotSum(0 To llUpper + 1) As SPOTSUM
'            End If
'        '8/13/14: Generate spots for last date+1 into lst with Fed as *
'        Else
'            tmDate1ODF(UBound(tmDate1ODF)) = tmOdf
'            ReDim Preserve tmDate1ODF(0 To UBound(tmDate1ODF) + 1) As ODF
'        End If
'    End If
'    Return
'lProcSeqNo:
'    ilSeqNo = 0
'    If slFor = "D" Then
'        'tmOdfSrchKey1.iUrfCode = tgUrf(0).iCode
'        tmOdfSrchKey1.iMnfFeed = tmOdf.iMnfFeed 'tmDlf.iMnfFeed
'        tmOdfSrchKey1.iAirDate(0) = tmOdf.iAirDate(0)   'ilLogDate0
'        tmOdfSrchKey1.iAirDate(1) = tmOdf.iAirDate(1)   'ilLogDate1
'        tmOdfSrchKey1.iFeedTime(0) = tmOdf.iFeedTime(0) 'tmDlf.iFeedTime(0)
'        tmOdfSrchKey1.iFeedTime(1) = tmOdf.iFeedTime(1) 'tmDlf.iFeedTime(1)
'        tmOdfSrchKey1.sZone = tmOdf.sZone   'tmDlf.sZone
'        tmOdfSrchKey1.iSeqNo = 32000
'        ilRet = btrGetLessOrEqual(hlOdf, tlOdf, imOdfRecLen, tmOdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'        'Do While (ilRet = BTRV_ERR_NONE) And (tlOdf.iUrfCode = tgUrf(0).iCode) And (tlOdf.iMnfFeed = tmOdf.iMnfFeed) And (tlOdf.iAirDate(0) = tmOdf.iAirDate(0)) And (tlOdf.iAirDate(1) = tmOdf.iAirDate(1))
'        Do While (ilRet = BTRV_ERR_NONE) And (tlOdf.iMnfFeed = tmOdf.iMnfFeed) And (tlOdf.iAirDate(0) = tmOdf.iAirDate(0)) And (tlOdf.iAirDate(1) = tmOdf.iAirDate(1))
'            If (tlOdf.iLocalTime(0) <> tmOdf.iLocalTime(0)) Or (tlOdf.iLocalTime(1) <> tmOdf.iLocalTime(1)) Or (tlOdf.sZone <> tmOdf.sZone) Then
'                Exit Do
'            End If
'            If ilODFVefCode <= 0 Then
'                If ilSimVefCode <= 0 Then
'                    'This test is OK on Log vehicle since ODF is created for LOG vehicle
'                    If tmVef.iVefCode > 0 Then  'Log vehicle defined
'                        If tlOdf.iVefCode = tmVef.iVefCode Then
'                            ilSeqNo = tlOdf.iSeqNo
'                            Exit Do
'                        End If
'                    Else
'                        If tlOdf.iVefCode = ilVehCode Then
'                            ilSeqNo = tlOdf.iSeqNo
'                            Exit Do
'                        End If
'                    End If
'                Else
'                    If tlOdf.iVefCode = ilSimVefCode Then
'                        ilSeqNo = tlOdf.iSeqNo
'                        Exit Do
'                    End If
'                End If
'            Else
'                If tlOdf.iVefCode = ilODFVefCode Then
'                    ilSeqNo = tlOdf.iSeqNo
'                    Exit Do
'                End If
'            End If
'            ilRet = btrGetPrevious(hlOdf, tlOdf, imOdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
'        Loop
'    Else
'        'tmOdfSrchKey0.iUrfCode = tgUrf(0).iCode
'        If ilODFVefCode <= 0 Then
'            If ilSimVefCode <= 0 Then
'                If tmVef.iVefCode > 0 Then  'Log vehicle defined
'                    tmOdfSrchKey0.iVefCode = tmVef.iVefCode
'                    ilTestVefCode = tmVef.iVefCode
'                Else
'                    tmOdfSrchKey0.iVefCode = ilVehCode
'                    ilTestVefCode = ilVehCode
'                End If
'            Else
'                tmOdfSrchKey0.iVefCode = ilSimVefCode
'                ilTestVefCode = ilSimVefCode
'            End If
'        Else
'            tmOdfSrchKey0.iVefCode = ilODFVefCode
'            ilTestVefCode = ilODFVefCode
'        End If
'        '8/13/14: Generate spots for last date+1 into lst with Fed as *
'        If Not blCreatingLstDate1 Then
'            tmOdfSrchKey0.iAirDate(0) = tmOdf.iAirDate(0)   'ilLogDate0
'            tmOdfSrchKey0.iAirDate(1) = tmOdf.iAirDate(1)   'ilLogDate1
'            tmOdfSrchKey0.iLocalTime(0) = tmOdf.iLocalTime(0)   'tmDlf.iLocalTime(0)
'            tmOdfSrchKey0.iLocalTime(1) = tmOdf.iLocalTime(1)   'tmDlf.iLocalTime(1)
'            tmOdfSrchKey0.sZone = tmOdf.sZone   'tmDlf.sZone
'            tmOdfSrchKey0.iSeqNo = 32000
'            ilRet = btrGetLessOrEqual(hlOdf, tlOdf, imOdfRecLen, tmOdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
'            'If (ilRet = BTRV_ERR_NONE) And (tlOdf.iUrfCode = tgUrf(0).iCode) And (tlOdf.iVefCode = ilTestVefCode) And (tlOdf.iAirDate(0) = tmOdf.iAirDate(0)) And (tlOdf.iAirDate(1) = tmOdf.iAirDate(1)) Then
'            If (ilRet = BTRV_ERR_NONE) And (tlOdf.iVefCode = ilTestVefCode) And (tlOdf.iAirDate(0) = tmOdf.iAirDate(0)) And (tlOdf.iAirDate(1) = tmOdf.iAirDate(1)) Then
'                If (tlOdf.iLocalTime(0) = tmOdf.iLocalTime(0)) And (tlOdf.iLocalTime(1) = tmOdf.iLocalTime(1)) And (tlOdf.sZone = tmOdf.sZone) Then
'                    ilSeqNo = tlOdf.iSeqNo
'                End If
'            End If
'        Else
'            For llOdf = 0 To UBound(tlDate1ODF) - 1 Step 1
'                tlOdf = tlDate1ODF(llOdf)
'                If (tlOdf.iVefCode = ilTestVefCode) And (tlOdf.iAirDate(0) = tmOdf.iAirDate(0)) And (tlOdf.iAirDate(1) = tmOdf.iAirDate(1)) Then
'                    If (tlOdf.iLocalTime(0) = tmOdf.iLocalTime(0)) And (tlOdf.iLocalTime(1) = tmOdf.iLocalTime(1)) And (tlOdf.sZone = tmOdf.sZone) Then
'                        If tlOdf.iSeqNo > ilSeqNo Then
'                            ilSeqNo = tlOdf.iSeqNo
'                        End If
'                    End If
'                End If
'            Next llOdf
'        End If
'    End If
'    Return
'lProcAdjDate:
'    ilDateAdj = False
'    'Test if Air time is AM and Local Time is PM. If so, adjust date
'    ilAirHour = tmOdf.iAirTime(1) \ 256  'Obtain month
'    ilLocalHour = tmOdf.iLocalTime(1) \ 256  'Obtain month
'    ilFeedHour = tmOdf.iFeedTime(1) \ 256
'    If (tgSpf.sGUseAffSys = "Y") Then
'        If slAdjLocalOrFeed <> "F" Then
'            If (ilAirHour < 6) And (ilLocalHour > 17) Then
'
'                '7/11/14: Lock avail between 12am-3am
'                If (slInLogType = "F") Or (slInLogType = "R") Or (slInLogType = "A") Then
'                    If llDate = llEDate + 1 Then
'                        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), slLockDate
'                        gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", slLockStartTime
'                        slLockEndTime = gFormatTimeLong(gTimeToLong(slLockStartTime, False) + 1, "A", "1")
'                        If (ilLockVefCode <> tmSdf.iVefCode) Or (gTimeToLong(slLockStartTime, False) <> llLockStartTime) Then
'                            gSetLockStatus tmSdf.iVefCode, 1, -1, slLockDate, slLockDate, tmSdf.iGameNo, slLockStartTime, slLockEndTime
'                            ilLockVefCode = tmSdf.iVefCode
'                            llLockStartTime = gTimeToLong(slLockStartTime, False)
'                        End If
'                    End If
'                End If
'
'                ilDateAdj = True
'                gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slAdjDate
'                slAdjDate = gDecOneDay(slAdjDate)
'                gPackDate slAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
'                If (llDate > llSDate) Then
'                    ilGenODF = True
'                Else
'                    ilGenODF = False
'                End If
'            ElseIf (ilLocalHour < 6) And (ilAirHour > 17) Then
'                ilDateAdj = True
'                gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slAdjDate
'                slAdjDate = gIncOneDay(slAdjDate)
'                gPackDate slAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
'                If (llDate >= llSDate) And (llDate <= llEDate) Then
'                    ilGenODF = True
'                Else
'                    ilGenODF = False
'                End If
'            Else
'                gUnpackTimeLong tmOdf.iLocalTime(0), tmOdf.iLocalTime(1), False, llLocalTime
'                '8/13/14: Generate spots for last date+1 into lst with Fed as *
'                'If (llLocalTime >= llLPTime) And (llLocalTime < llLNTime) Then
'                '    ilGenODF = True
'                'Else
'                '    ilGenODF = False
'                'End If
'                If Not blCreatingLstDate1 Then
'                    If (llLocalTime >= llLPTime) And (llLocalTime < llLNTime) Then
'                        ilGenODF = True
'                    Else
'                        ilGenODF = False
'                    End If
'                Else
'                    If (llLocalTime >= 0) And (llLocalTime < llGLocalAdj) Then
'                        ilGenODF = True
'                    Else
'                        ilGenODF = False
'                    End If
'                End If
'            End If
'        Else
'            If (ilAirHour < 6) And (ilFeedHour > 17) Then
'                ilDateAdj = True
'                gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slAdjDate
'                slAdjDate = gDecOneDay(slAdjDate)
'                gPackDate slAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
'                If (llDate > llSDate) Then
'                    ilGenODF = True
'                Else
'                    ilGenODF = False
'                End If
'            ElseIf (ilFeedHour < 6) And (ilAirHour > 17) Then
'                ilDateAdj = True
'                gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slAdjDate
'                slAdjDate = gIncOneDay(slAdjDate)
'                gPackDate slAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
'                If (llDate >= llSDate) And (llDate <= llEDate) Then
'                    ilGenODF = True
'                Else
'                    ilGenODF = False
'                End If
'            Else
'                gUnpackTimeLong tmOdf.iFeedTime(0), tmOdf.iFeedTime(1), False, llFeedTime
'                If (llFeedTime >= llLPTime) And (llFeedTime < llLNTime) Then
'                    ilGenODF = True
'                Else
'                    ilGenODF = False
'                End If
'            End If
'        End If
'    Else
'        ilGenODF = True
'        If slAdjLocalOrFeed <> "F" Then
'            If (ilAirHour < 6) And (ilLocalHour > 17) Then
'                'If monday convert to next sunday- this is wrong but the same spot
'                'runs each sunday (the spot should have show on the previous week sunday)
'                'If not monday, then subtract one day
'                ilDateAdj = True
'                gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slAdjDate
'                If gWeekDayStr(slAdjDate) = 0 Then
'                    slAdjDate = gObtainNextSunday(slAdjDate)
'                Else
'                    slAdjDate = gDecOneDay(slAdjDate)
'                End If
'                gPackDate slAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
'            End If
'        Else
'            If (ilAirHour < 6) And (ilFeedHour > 17) Then
'                'If monday convert to next sunday- this is wrong but the same spot
'                'runs each sunday (the spot should have show on the previous week sunday)
'                'If not monday, then subtract one day
'                ilDateAdj = True
'                gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slAdjDate
'                If gWeekDayStr(slAdjDate) = 0 Then
'                    slAdjDate = gObtainNextSunday(slAdjDate)
'                Else
'                    slAdjDate = gDecOneDay(slAdjDate)
'                End If
'                gPackDate slAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
'            End If
'        End If
'    End If
'    gUnpackDateLong tmOdf.iAirDate(0), tmOdf.iAirDate(1), llWkDateSet
'    tmOdf.iWkNo = (llWkDateSet - ll010570) \ 7 + 1
'    If (ilDateAdj = False) And (slXMid = "Y") Then
'        gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slAdjDate
'        slAdjDate = gIncOneDay(slAdjDate)
'        gPackDate slAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
'    End If
'    Return
'lProcTestForDuplBB:
'    ilAddSpot = True
'    If tgSpf.sUsingBBs <> "Y" Then
'        Return
'    End If
'    If (tmSdf.sSpotType <> "O") Or (tmSdf.sSpotType <> "C") Then
'        Return
'    End If
'    For ilBB = 0 To UBound(tmBBSdfInfo) - 1 Step 1
'        If (tmBBSdfInfo(ilBB).lChfCode = tmSdf.lChfCode) And (tmBBSdfInfo(ilBB).iLen = tmSdf.iLen) And (tmBBSdfInfo(ilBB).sType = tmSdf.sSpotType) Then
'            If (tmBBSdfInfo(ilBB).iTime(0) = tmSdf.iTime(0)) And (tmBBSdfInfo(ilBB).iTime(1) = tmSdf.iTime(1)) Then
'                ilAddSpot = False
'                Return
'            End If
'        End If
'    Next ilBB
'    tmBBSdfInfo(UBound(tmBBSdfInfo)).sType = tmSdf.sSpotType
'    tmBBSdfInfo(UBound(tmBBSdfInfo)).lChfCode = tmSdf.lChfCode
'    tmBBSdfInfo(UBound(tmBBSdfInfo)).iLen = tmSdf.iLen
'    tmBBSdfInfo(UBound(tmBBSdfInfo)).iTime(0) = tmSdf.iTime(0)
'    tmBBSdfInfo(UBound(tmBBSdfInfo)).iTime(1) = tmSdf.iTime(1)
'    tmOdf.iBreakNo = 0  'tmAvail.iAvInfo
'    tmOdf.iPositionNo = 0
'    For ilBB = 0 To UBound(tmBBSdfInfo) - 1 Step 1
'        If (tmBBSdfInfo(ilBB).iTime(0) = tmSdf.iTime(0)) And (tmBBSdfInfo(ilBB).iTime(1) = tmSdf.iTime(1)) Then
'            tmOdf.iPositionNo = tmOdf.iPositionNo + 1
'        End If
'    Next ilBB
'    ReDim Preserve tmBBSdfInfo(0 To UBound(tmBBSdfInfo) + 1) As BBSDFINFO
'    Return
'CloseFiles:
'    On Error Resume Next
'    Erase tmLstCode
'    If imLstExist Then
'        'ilRet = btrClose(hmLst)
'        btrDestroy hmMnf
'        'btrDestroy hmCff
'        'btrDestroy hmSmf
'        btrDestroy hmDrf
'        btrDestroy hmDpf
'        btrDestroy hmDef
'        btrDestroy hmRaf
'        'btrDestroy hmLst
'    End If
'    ilRet = btrClose(hmCpf)
'    btrDestroy hmCpf
'    ilRet = btrClose(hmCvf)
'    ilRet = btrClose(hlRsf)
'    ilRet = btrClose(hlCnf)
'    ilRet = btrClose(hmCff)
'    ilRet = btrClose(hmSmf)
'    ilRet = btrClose(hlRdf)
'    ilRet = btrClose(hlCef)
'    ilRet = btrClose(hlCTSsf)
'    ilRet = btrClose(hlSsf)
'    ilRet = btrClose(hlSdf)
'    ilRet = btrClose(hlChf)
'    ilRet = btrClose(hlOdf)
'    ilRet = btrClose(hlVef)
'    ilRet = btrClose(hlVlf)
'    ilRet = btrClose(hlDlf)
'    ilRet = btrClose(hmCif)
'    ilRet = btrClose(hlTzf)
'    ilRet = btrClose(hlCrf)
'    ilRet = btrClose(hlSif)
'    ilRet = btrClose(hlAdf)
'    ilRet = btrClose(hmClf)
'    ilRet = btrClose(hlVsf)
'    ilRet = btrClose(hlcxf)
'
'    btrDestroy hmCff
'    btrDestroy hmSmf
'    btrDestroy hlRdf
'    btrDestroy hlCTSsf
'    btrDestroy hlSsf
'    btrDestroy hlSdf
'    btrDestroy hlChf
'    btrDestroy hlOdf
'    btrDestroy hlVef
'    btrDestroy hlVlf
'    btrDestroy hlDlf
'    btrDestroy hmCif
'    btrDestroy hlTzf
'    btrDestroy hlCrf
'    btrDestroy hlSif
'    btrDestroy hlAdf
'    btrDestroy hmClf
'    btrDestroy hlVsf
'    btrDestroy hlCef
'    btrDestroy hlCnf
'    btrDestroy hlRsf
'    btrDestroy hmCvf
'    btrDestroy hlcxf
'
'    If tgSpf.sSystemType = "R" Then
'        ilRet = btrClose(hmFsf)
'        ilRet = btrClose(hmPrf)
'        ilRet = btrClose(hmFnf)
'        btrDestroy hmFsf
'        btrDestroy hmPrf
'        btrDestroy hmFnf
'    End If
'
'    Erase ilVehicle
'    Erase tmPageEject
'    Erase tlLlc
'    Return
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:gBuildODFSpotDay                *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Build a day into ODF            *
'*                                                     *
'*******************************************************
Sub gDeleteOdf(slFor As String, ilType As Integer, sLCP As String, ilCallCode As Integer)
'
'   gDeleteODF slFor, ilType, slCp, ilCallCode
'
'   Where:
'       slFor (I)- "L"=Log; "C"=Commercial; "D"=Delivery; G=Gen Date  and time (igGenDate and igGenTime used)
'       ilType (I)- 0=On air; 1=Alternate
'       slCP (I)- "C"=Current only; "P"=Pending only; "B"=Both
'       ilCallCode (I)-Vehicle code number(slFor = L or C) or feed code (slFor = D)
'       slSDate (I)- Start Date that events are to be obtained
'       slEDate (I)- Start Date that events are to be obtained
    Dim ilRet As Integer
    Dim ilCRet As Integer
    Dim ilVefCode As Integer
    Dim llOdfRecPos As Long
    imOdfRecLen = Len(tmOdf)  'Get and save ODF record length
    hmOdf = CBtrvTable(TEMPHANDLE)        'Create ODF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmOdf, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmOdf)
        btrDestroy hmOdf
        Exit Sub
    End If
    imVefRecLen = Len(tmVef)  'Get and save VEF record length
    hmVef = CBtrvTable(ONEHANDLE)        'Create VLF object handle
    On Error GoTo 0
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    If ilRet <> BTRV_ERR_NONE Then
        ilRet = btrClose(hmOdf)
        ilRet = btrClose(hmVef)
        btrDestroy hmOdf
        btrDestroy hmVef
        Exit Sub
    End If
    tmVef.iCode = 0
    'llSDate = gDateValue(slSDate)  'Dates removed from call
    'llEDate = gDateValue(slEDate)
    'Clear ODF
    If slFor = "D" Then
        'For llDate = llSDate To llEDate Step 1
            'slDate = Format$(llDate, "m/d/yy")
            'ilDay = gWeekDayStr(slDate)
            'gPackDate slDate, ilLogDate0, ilLogDate1
            'tmOdfSrchKey1.iUrfCode = tgUrf(0).iCode
            tmOdfSrchKey1.iMnfFeed = ilCallCode
            tmOdfSrchKey1.iAirDate(0) = 0   'ilLogDate0
            tmOdfSrchKey1.iAirDate(1) = 0   'ilLogDate1
            tmOdfSrchKey1.iFeedTime(0) = 0
            tmOdfSrchKey1.iFeedTime(1) = 0
            tmOdfSrchKey1.sZone = ""
            tmOdfSrchKey1.iSeqNo = 0
            ilRet = btrGetGreaterOrEqual(hmOdf, tmOdf, imOdfRecLen, tmOdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            'Do While (ilRet = BTRV_ERR_NONE) And (tmOdf.iUrfCode = tgUrf(0).iCode) And (tmOdf.iMnfFeed = ilCallCode) And (tmOdf.iAirDate(0) = ilLogDate0) And (tmOdf.iAirDate(1) = ilLogDate1)
            'Do While (ilRet = BTRV_ERR_NONE) And (tmOdf.iMnfFeed = ilCallCode) And (tmOdf.iAirDate(0) = ilLogDate0) And (tmOdf.iAirDate(1) = ilLogDate1)
            Do While (ilRet = BTRV_ERR_NONE) And (tmOdf.iMnfFeed = ilCallCode)
                ilRet = btrGetPosition(hmOdf, llOdfRecPos)
                Do
                    ilRet = btrDelete(hmOdf)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        ilCRet = btrGetDirect(hmOdf, tmOdf, imOdfRecLen, llOdfRecPos, INDEXKEY1, BTRV_LOCK_NONE)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                ilRet = btrGetNext(hmOdf, tmOdf, imOdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
            Loop
        'Next llDate
    ElseIf slFor = "G" Then
        tmOdfSrchKey2.iGenDate(0) = igGenDate(0)   'ilLogDate0
        tmOdfSrchKey2.iGenDate(1) = igGenDate(1)
        '10-9-01
        tmOdfSrchKey2.lGenTime = lgGenTime
        'tmOdfSrchKey2.iGenTime(0) = igGenTime(0)
        'tmOdfSrchKey2.iGenTime(1) = igGenTime(1)
        ilRet = btrGetEqual(hmOdf, tmOdf, imOdfRecLen, tmOdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
        Do While (ilRet = BTRV_ERR_NONE)
            ilRet = btrDelete(hmOdf)
            If (ilRet = BTRV_ERR_CONFLICT) Or (ilRet = BTRV_ERR_NONE) Then
                tmOdfSrchKey2.iGenDate(0) = igGenDate(0)   'ilLogDate0
                tmOdfSrchKey2.iGenDate(1) = igGenDate(1)
                '10-9-01
                tmOdfSrchKey2.lGenTime = lgGenTime
                'tmOdfSrchKey2.iGenTime(0) = igGenTime(0)
                'tmOdfSrchKey2.iGenTime(1) = igGenTime(1)
                ilRet = btrGetEqual(hmOdf, tmOdf, imOdfRecLen, tmOdfSrchKey2, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
            End If
        Loop
    Else
        tmVefSrchKey.iCode = ilCallCode
        ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        If (ilRet <> BTRV_ERR_NONE) Then
            ilRet = btrClose(hmOdf)
            ilRet = btrClose(hmVef)
            btrDestroy hmOdf
            btrDestroy hmVef
            Exit Sub
        End If
        If tmVef.iVefCode > 0 Then
            ilVefCode = tmVef.iVefCode
            tmOdfSrchKey0.iVefCode = tmVef.iVefCode
        Else
            ilVefCode = ilCallCode
            tmOdfSrchKey0.iVefCode = ilCallCode
        End If
        tmOdfSrchKey0.iAirDate(0) = 0   'ilLogDate0
        tmOdfSrchKey0.iAirDate(1) = 0   'ilLogDate1
        tmOdfSrchKey0.iLocalTime(0) = 0
        tmOdfSrchKey0.iLocalTime(1) = 0
        tmOdfSrchKey0.sZone = ""
        tmOdfSrchKey0.iSeqNo = 0
        ilRet = btrGetGreaterOrEqual(hmOdf, tmOdf, imOdfRecLen, tmOdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmOdf.iVefCode = ilVefCode)
            If tmOdf.iMnfFeed = 0 Then  'Not generated for delivery
                ilRet = btrGetPosition(hmOdf, llOdfRecPos)
                Do
                    ilRet = btrDelete(hmOdf)
                    If ilRet = BTRV_ERR_CONFLICT Then
                        ilCRet = btrGetDirect(hmOdf, tmOdf, imOdfRecLen, llOdfRecPos, INDEXKEY0, BTRV_LOCK_NONE)
                    End If
                Loop While ilRet = BTRV_ERR_CONFLICT
                'ilRet = btrGetNext(hmOdf, tmOdf, imOdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                tmOdfSrchKey0.iVefCode = ilVefCode
                tmOdfSrchKey0.iAirDate(0) = 0   'ilLogDate0
                tmOdfSrchKey0.iAirDate(1) = 0   'ilLogDate1
                tmOdfSrchKey0.iLocalTime(0) = 0
                tmOdfSrchKey0.iLocalTime(1) = 0
                tmOdfSrchKey0.sZone = ""
                tmOdfSrchKey0.iSeqNo = 0
                ilRet = btrGetGreaterOrEqual(hmOdf, tmOdf, imOdfRecLen, tmOdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            Else
                ilRet = btrGetNext(hmOdf, tmOdf, imOdfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
            End If
        Loop
    End If
    ilRet = btrClose(hmOdf)
    ilRet = btrClose(hmVef)
    btrDestroy hmOdf
    btrDestroy hmVef
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mClearLst                       *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Remove all avails/Spots         *
'*                                                     *
'*******************************************************
Sub mClearLst(ilVefCode As Integer, llDate As Long, llSTime As Long, llETime As Long, ilCreateLST As Integer, llGsfCode As Long, slLSTCreateMsg As String, blLstFound As Boolean)
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llLogTime As Long
    Dim llLogDate As Long
    Dim slMsg As String
    Dim ilCRet As Integer
    ReDim ilDate(0 To 1) As Integer
    ReDim tmLstCode(0 To 0) As LST
    slLSTCreateMsg = ""
    tmLstSrchKey2.iLogVefCode = ilVefCode
    gPackDateLong llDate, ilDate(0), ilDate(1)
    tmLstSrchKey2.iLogDate(0) = ilDate(0)
    tmLstSrchKey2.iLogDate(1) = ilDate(1)
    'ilRet = btrGetEqual(hmLst, tmLst, imLstRecLen, tmLstSrchKey, INDEXKEY2, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
    ilRet = btrGetGreaterOrEqual(hmLst, tmLst, imLstRecLen, tmLstSrchKey2, INDEXKEY2, BTRV_LOCK_NONE)   'Get last current record to obtain date
    ''8/10/10:  Handle games that cross midnight
    ''Do While (ilRet = BTRV_ERR_NONE) And (tmLst.iLogVefCode = ilVefCode) And (((tmLst.iLogDate(0) = ilDate(0)) And (tmLst.iLogDate(1) = ilDate(1))) Or (tmLst.lGsfCode = llGsfCode))
    'Do While (ilRet = BTRV_ERR_NONE) And (tmLst.iLogVefCode = ilVefCode) And (((tmLst.iLogDate(0) = ilDate(0)) And (tmLst.iLogDate(1) = ilDate(1)) And (llGsfCode = 0)) Or ((tmLst.lGsfCode = llGsfCode) And (llGsfCode > 0)))
    '9/14/10:  Handle two sporting events on same day
    If bgReprintLogType Then
        slMsg = "Reprint: Date " & Format(llDate, "m/d/yy") & "; VefCode " & ilVefCode & "; GsfCode " & llGsfCode & " LST Not Found"
        If (ilRet <> BTRV_ERR_NONE) Or (tmLst.iLogVefCode <> ilVefCode) Then
            If (ilRet = 30000) Or (ilRet = 30001) Or (ilRet = 30002) Or (ilRet = 30003) Then
                ilCRet = csiHandleValue(0, 7)
            Else
                ilCRet = ilRet
            End If
            slMsg = slMsg & " BTR Status " & ilCRet
            'gLogMsg slMsg, "TrafficErrors.Txt", False
            slLSTCreateMsg = slMsg
        Else
            If (llGsfCode = 0) Or ((tmLst.lGsfCode = llGsfCode) And (llGsfCode > 0)) Then
                gUnpackDateLong tmLst.iLogDate(0), tmLst.iLogDate(1), llLogDate
                If llGsfCode = 0 Then
                    If llDate <> llLogDate Then
                        'gLogMsg slMsg, "TrafficErrors.Txt", False
                        slLSTCreateMsg = slMsg
                    End If
                Else
                    If llLogDate > llDate + 1 Then
                        'gLogMsg slMsg, "TrafficErrors.Txt", False
                        slLSTCreateMsg = slMsg
                    End If
                End If
            Else
                'gLogMsg slMsg, "TrafficErrors.Txt", False
                slLSTCreateMsg = slMsg
            End If
        End If
    End If
    Do While (ilRet = BTRV_ERR_NONE) And (tmLst.iLogVefCode = ilVefCode)
        gUnpackDateLong tmLst.iLogDate(0), tmLst.iLogDate(1), llLogDate
        If llGsfCode = 0 Then
            If llDate <> llLogDate Then
                Exit Do
            End If
        Else
            If llLogDate > llDate + 1 Then
                Exit Do
            End If
        End If
        'If tmLst.iType = 1 Then
        '12/9/08: Retain blackout replacement.  If not required, they will be removed in affiliate system
        '9/14/10:  Bypass games that don't match the game that is being processed
        '3/9/16: Bypass MG, Replacement and Bonus LST
        'If tmLst.iStatus Mod 100 < ASTEXTENDED_MG Then
        If (tmLst.iStatus Mod 100 < ASTEXTENDED_MG) And (tmLst.iType <> 2) Then
            If (llGsfCode = 0) Or ((tmLst.lGsfCode = llGsfCode) And (llGsfCode > 0)) Then
                '10/30/18: Added checking if valid blackout
                If Not mRemovedBadLST(tmLst) Then
                    If (tmLst.lBkoutLstCode <= 0) Or (ilCreateLST) Then
                        gUnpackTimeLong tmLst.iLogTime(0), tmLst.iLogTime(1), False, llLogTime
                        If (llLogTime >= llSTime) And (llLogTime < llETime) And tmLst.lGsfCode = llGsfCode Then
                            blLstFound = True
                            tmLstCode(UBound(tmLstCode)) = tmLst
                            ReDim Preserve tmLstCode(0 To UBound(tmLstCode) + 1) As LST
                        End If
                    End If
                End If
            End If
        End If
        'End If
        ilRet = btrGetNext(hmLst, tmLst, imLstRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If ilCreateLST Then
        'If ilCreateLst is false, then unused removed in mDeleteUnusedLST
        For ilLoop = 0 To UBound(tmLstCode) - 1 Step 1
            tmLstSrchKey.lCode = tmLstCode(ilLoop).lCode
            ilRet = btrGetEqual(hmLst, tmLst, imLstRecLen, tmLstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) Then
                ilRet = btrDelete(hmLst)
                gLogBtrError ilRet, "mClearLst: Delete"
            End If
        Next ilLoop
        Erase tmLstCode
        ReDim tmLstCode(0 To 0) As LST
    End If
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateAvailLst                 *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Create LST for open avail       *
'*                                                     *
'*******************************************************
Sub mCreateAvailLst(ilVefCode As Integer, ilLen As Integer, ilUnits As Integer, llGsfCode As Long)
    Dim ilRet As Integer
    Dim ilAnf As Integer
    ilAnf = gBinarySearchAnf(tmOdf.ianfCode, tgAvailAnf())
    If ilAnf <> -1 Then
        If tgAvailAnf(ilAnf).sTrafToAff = "N" Then
            Exit Sub
        End If
    End If
    tmLst.lCode = 0
    tmLst.iType = 1
    tmLst.iStatus = 0
    tmLst.lSdfCode = 0
    tmLst.lCntrNo = 0
    tmLst.lFsfCode = 0
    tmLst.lGsfCode = llGsfCode
    tmLst.iAdfCode = 0
    tmLst.iAgfCode = 0
    tmLst.sProd = ""
    tmLst.iLineNo = 0
    tmLst.iLen = ilLen
    tmLst.iUnits = ilUnits
    tmLst.iLnVefCode = 0
    tmLst.iStartDate(0) = 0
    tmLst.iStartDate(1) = 0
    tmLst.iEndDate(0) = 0
    tmLst.iEndDate(1) = 0
    tmLst.iDays(0) = 0
    tmLst.iDays(0) = 0
    tmLst.iDays(1) = 0
    tmLst.iDays(2) = 0
    tmLst.iDays(3) = 0
    tmLst.iDays(4) = 0
    tmLst.iDays(5) = 0
    tmLst.iDays(6) = 0
    tmLst.iSpotsWk = 0
    tmLst.iPriceType = 0
    tmLst.lPrice = 0
    tmLst.iSpotType = 0
    tmLst.iLogVefCode = ilVefCode   'tmOdf.iVefCode
    'Air date is associated with LocalTime
    tmLst.iLogDate(0) = tmOdf.iAirDate(0)
    tmLst.iLogDate(1) = tmOdf.iAirDate(1)
    tmLst.iLogTime(0) = tmOdf.iLocalTime(0) 'tmOdf.iAirTime(0)
    tmLst.iLogTime(1) = tmOdf.iLocalTime(1) 'tmOdf.iAirTime(1)
    tmLst.sDemo = ""
    tmLst.lAud = 0
    tmLst.sISCI = ""
    tmLst.iWkNo = tmOdf.iWkNo
    tmLst.iBreakNo = tmOdf.iBreakNo
    tmLst.iPositionNo = tmOdf.iPositionNo
    tmLst.iSeqNo = tmOdf.iSeqNo
    tmLst.sZone = tmOdf.sZone
    tmLst.sCart = ""
    tmLst.lcpfCode = 0
    tmLst.lCrfCsfcode = 0
    tmLst.ianfCode = tmOdf.ianfCode
    tmLst.lCifCode = 0
    tmLst.lEvtIDCefCode = tmOdf.lEvtIDCefCode
    tmLst.lBkoutLstCode = 0
    ilRet = btrInsert(hmLst, tmLst, imLstRecLen, INDEXKEY0)
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mCreateLst                      *
'*                                                     *
'*             Created:5/18/93       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments:Create LST from ODF             *
'*      1/18/99 DH  Comment reading of Clf & Cff which *
'*      has already been read when formatting the DP   *
'*      description                                    *
'*                                                     *
'*******************************************************
Private Sub mCreateLst(ilVefCode As Integer, ilClearLstSdf As Integer, hlClf As Integer, hlCif As Integer, hlMcf As Integer, slLogType As String, ilCreateLST As Integer, ilExportType As Integer, llGsfCode As Long, ilWegenerOLA As Integer, slAlertStatus As String)
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim ilFound As Integer
    Dim ilDay As Integer
    Dim slDate As String
    ReDim ilInputDay(0 To 6) As Integer
    Dim llOvStartTime As Long
    Dim llOvEndTime As Long
    Dim llStartTime As Long
    Dim llDate As Long
    Dim llLSTTime As Long
    Dim llLstDate As Long
    Dim ilMnfDemo As Integer
    Dim ilDnfCode As Integer
    Dim ilMnfSocEco As Integer
    Dim llAvgAud As Long
    ''Dim tlClfSrchKey As CLFKEY0 'CLF key record image
    ''Dim ilClfRecLen As Integer  'CLF record length
    'Dim tmCifSrchKey As LONGKEY0 'CIF key record image
    'Dim ilCifRecLen As Integer  'CIF record length
    'Dim tmCif As CIF            'CIF record image
    ''Dim tlClf As CLF            'CLF record image
    Dim llPopEst As Long
    Dim ilAudFromSource As Integer
    Dim llAudFromCode As Long
    Dim ilAnf As Integer
    Dim ilVff As Integer
    
    'Dim blProvider As Boolean
    ilAnf = gBinarySearchAnf(tmOdf.ianfCode, tgAvailAnf())
    If ilAnf <> -1 Then
        If tgAvailAnf(ilAnf).sTrafToAff = "N" Then
            Exit Sub
        End If
    End If

    gUnpackTimeLong tmOdf.iLocalTime(0), tmOdf.iLocalTime(1), False, llStartTime
    gUnpackDateLong tmOdf.iAirDate(0), tmOdf.iAirDate(1), llDate
    If ilCreateLST Then
        tmLst.lCode = 0
    Else
        tmLst.lCode = 0
        'Look for match, if found us its lstCode value.  If not found, this set alert on
        ilFound = False
        For ilLoop = 0 To UBound(tmLstCode) - 1 Step 1
            If (tmLstCode(ilLoop).lSdfCode = tmSdf.lCode) And (tmLstCode(ilLoop).sZone = tmOdf.sZone) Then
                gUnpackTimeLong tmLstCode(ilLoop).iLogTime(0), tmLstCode(ilLoop).iLogTime(1), False, llLSTTime
                gUnpackDateLong tmLstCode(ilLoop).iLogDate(0), tmLstCode(ilLoop).iLogDate(1), llLstDate
                If (llDate = llLstDate) And (llStartTime = llLSTTime) And (tmLstCode(ilLoop).lGsfCode = llGsfCode) Then
                    tmLst.lCode = tmLstCode(ilLoop).lCode
                    'Delete so that the update will work
                    tmLstSrchKey.lCode = tmLstCode(ilLoop).lCode
                    ilRet = btrGetEqual(hmLst, tmLst, imLstRecLen, tmLstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
                    If (ilRet = BTRV_ERR_NONE) Then
                        ilRet = btrDelete(hmLst)
                    End If
                    'Flag record so that it will not be used again and not deleted in mDeleteUsedLST
                    tmLstCode(ilLoop).lCode = 0
                    ilFound = True
                    Exit For
                End If
            End If
        Next ilLoop
        '7/20/12: Only create alert if exported previously, move to end of this routine
        'If (Not ilFound) And (slLogType = "R") Then
        '    'Checking if date bewteen todays date and last log is not required
        '    'It is Ok if Alert set on when generating Final Logs (Normally ilCreateLST would be true)
        '    slDate = Format$(llDate, "m/d/yy")
        '    If ilExportType > 0 Then
        '
        '        ilRet = gAlertAdd(slLogType, "S", 0, ilVefCode, slDate)
        '    End If
        '    'Only update ISCI Alert if Provider defined and not embedded or Provider and Produce defined and embedded
        '    'Same logic in affiliate when showing vehicles to generate ISCI for.
        '    ilRet = gBinarySearchVpf(ilVefCode)
        '    If ilRet <> -1 Then
        '        If tgVpf(ilRet).iCommProvArfCode > 0 Then
        '            '1/21/10:  Removed embedded from producer definition (Vehicle-Option).
        '            'If (tgVpf(ilRet).sEmbeddedComm <> "Y") Or ((tgVpf(ilRet).sEmbeddedComm = "Y") And (tgVpf(ilRet).iProducerArfCode > 0)) Then
        '                ilRet = gAlertAdd(slLogType, "I", 0, ilVefCode, slDate)
        '            'End If
        '        End If
        '    End If
        'End If
    End If
    tmLst.iType = 0
    tmLst.iStatus = 0 'Aired
    tmLst.lSdfCode = tmSdf.lCode
    tmLst.lCntrNo = tmChf.lCntrNo
    tmLst.lFsfCode = tmSdf.lFsfCode
    tmLst.lGsfCode = llGsfCode
    tmLst.iAdfCode = tmChf.iAdfCode
    tmLst.iAgfCode = tmChf.iAgfCode
    tmLst.sProd = tmChf.sProduct
    tmLst.iLineNo = tmSdf.iLineNo
    tmLst.iLen = tmSdf.iLen
    tmLst.iUnits = 0

    mGetLineTimes tmLst.iLnStartTime(), tmLst.iLnEndTime()

    'tlClfSrchKey.lChfCode = tmSdf.lChfCode
    'tlClfSrchKey.iLine = tmSdf.iLineNo
    'tlClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
    'tlClfSrchKey.iPropVer = 32000 ' Plug with very high number
    'ilClfRecLen = Len(tlClf)
    'ilRet = btrGetGreaterOrEqual(hlClf, tlClf, ilClfRecLen, tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    'Do While (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tmSdf.lChfCode) And (tlClf.iLine = tmSdf.iLineNo) And (tlClf.sSchStatus = "A")
    '    ilRet = btrGetNext(hlClf, tlClf, ilClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    'Loop
    'If (ilRet = BTRV_ERR_NONE) And (tlClf.lChfCode = tmSdf.lChfCode) And (tlClf.iLine = tmSdf.iLineNo) Then
        tmLst.iLnVefCode = tmClf.iVefCode   'changed from tlclf to tmclf
     '   ilRet = gGetSpotFlight(tmSdf, tlClf, hmCff, hmSmf, tmCff)
     '   If ilRet Then
            tmLst.iStartDate(0) = tmCff.iStartDate(0)
            tmLst.iStartDate(1) = tmCff.iStartDate(1)
            tmLst.iEndDate(0) = tmCff.iEndDate(0)
            tmLst.iEndDate(1) = tmCff.iEndDate(1)
            For ilLoop = 0 To 6 Step 1
                tmLst.iDays(ilLoop) = tmCff.iDay(ilLoop)
            Next ilLoop
            tmLst.iSpotsWk = tmCff.iSpotsWk
            If tmSdf.sSpotType = "X" Then
                tmLst.iSpotType = 2
                tmLst.iPriceType = 1
                tmLst.lPrice = 0
            ElseIf tmSdf.sSpotType = "O" Then
                tmLst.iSpotType = 6
                tmLst.iPriceType = 1
                tmLst.lPrice = 0
            ElseIf tmSdf.sSpotType = "C" Then
                tmLst.iSpotType = 7
                tmLst.iPriceType = 1
                tmLst.lPrice = 0
            Else
                If tmSdf.sSchStatus = "G" Then
                    tmLst.iSpotType = 1
                ElseIf tmSdf.sSchStatus = "O" Then
                    tmLst.iSpotType = 3
                Else
                    tmLst.iSpotType = 0
                End If
                Select Case tmCff.sPriceType
                    Case "B"    'Bonus
                        tmLst.iPriceType = 1
                        tmLst.lPrice = 0
                    Case "N"    'No Charge
                        tmLst.iPriceType = 2
                        tmLst.lPrice = 0
                    Case "M"    'MG
                        tmLst.iPriceType = 3
                        tmLst.lPrice = 0
                    Case "S"    'Spinoff
                        tmLst.iPriceType = 4
                        tmLst.lPrice = 0
                    Case "R"    'Recapturable
                        tmLst.iPriceType = 5
                        tmLst.lPrice = 0
                    Case "A"    'A=Audience Deficiency
                        tmLst.iPriceType = 6
                        tmLst.lPrice = 0
                    Case Else
                        tmLst.iPriceType = 0
                        tmLst.lPrice = tmCff.lActPrice
                End Select
            End If
            tmLst.iLogVefCode = ilVefCode   'tmOdf.iVefCode
            tmLst.iLogDate(0) = tmOdf.iAirDate(0)   'This is Local Date
            tmLst.iLogDate(1) = tmOdf.iAirDate(1)
            tmLst.iLogTime(0) = tmOdf.iLocalTime(0) 'tmOdf.iAirTime(0)
            tmLst.iLogTime(1) = tmOdf.iLocalTime(1) 'tmOdf.iAirTime(1)
            ilMnfDemo = tmClf.iMnfDemo  'tmChf.iMnfDemo(0), changed from tlClf to tmClf
            ilDnfCode = tmClf.iDnfCode  'changed from tlClf to tmClf
            ilMnfSocEco = 0
            llAvgAud = 0
            If (ilMnfDemo > 0) And (ilDnfCode > 0) Then
                tmMnfSrchKey.iCode = ilMnfDemo
                ilRet = btrGetEqual(hmMnf, tmMnf, imMnfRecLen, tmMnfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
                If (ilRet = BTRV_ERR_NONE) Then
                    tmLst.sDemo = tmMnf.sName
                    'If ((tmClf(ilClf).ClfRec.iStartTime(0) <> 1) Or (tmClf(ilClf).ClfRec.iStartTime(1) <> 0)) And ((tmClf(ilClf).ClfRec.iEndTime(0) <> 1) Or (tmClf(ilClf).ClfRec.iEndTime(1) <> 0)) Then
                    '    gUnpackTimeLong tlClf.iStartTime(0), tlClf.iStartTime(1), False, llOvStartTime
                    '    gUnpackTimeLong tlClf.iEndTime(0), tlClf.iEndTime(1), True, llOvEndTime
                    'Else
                    '    llOvStartTime = 0
                    '    llOvEndTime = 0
                    'End If
                    gUnpackTimeLong tmOdf.iLocalTime(0), tmOdf.iLocalTime(1), False, llOvStartTime
                    gUnpackTimeLong tmOdf.iLocalTime(0), tmOdf.iLocalTime(1), True, llOvEndTime
                    For ilDay = 0 To 6 Step 1
                        ilInputDay(ilDay) = False
                    Next ilDay
                    gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), slDate
                    ilDay = gWeekDayStr(slDate)
                    ilInputDay(ilDay) = True
                    'ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, ilDnfCode, tmOdf.iVefCode, ilMnfSocEco, ilMnfDemo, tmClf.iRdfcode, llOvStartTime, llOvEndTime, ilInputDay(), llAvgAud) 'chged from tlClf to tmClf
                    ilRet = gGetDemoAvgAud(hmDrf, hmMnf, hmDpf, hmDef, hmRaf, ilDnfCode, ilVefCode, ilMnfSocEco, ilMnfDemo, llDate, llDate, tmClf.iRdfCode, llOvStartTime, llOvEndTime, ilInputDay(), tmClf.sType, tmClf.lRafCode, llAvgAud, llPopEst, ilAudFromSource, llAudFromCode) 'chged from tlClf to tmClf
                Else
                    tmLst.sDemo = ""
                End If
            Else
                tmLst.sDemo = ""
            End If
            tmLst.lAud = llAvgAud
            'tlCif.lcpfCode = 0
            'tlCif.iMcfCode = 0
            'If tmOdf.lCifCode > 0 Then
            '    ilCifRecLen = Len(tlCif)
            '    tlCifSrchKey.lCode = tmOdf.lCifCode
            '    ilRet = btrGetEqual(hlCif, tlCif, ilCifRecLen, tlCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            '    If (ilRet = BTRV_ERR_NONE) Then
            '        tmCpfSrchKey.lCode = tlCif.lcpfCode
            '        ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            '        If (ilRet = BTRV_ERR_NONE) Then
            '            tmLst.sISCI = tmCpf.sISCI
            '        Else
            '            tmLst.sISCI = ""
            '        End If
            '    Else
            '        tmLst.sISCI = ""
            '    End If
            'Else
            '    tmLst.sISCI = ""
            'End If
            'tmCif.lcpfCode = 0
            'tmCif.iMcfCode = 0
            mGetCpf hlCif, tmOdf.lCifCode
            'If tmOdf.lCifCode > 0 Then
            '    ilCifRecLen = Len(tmCif)
            '    tmCifSrchKey.lCode = tmOdf.lCifCode
            '    ilRet = btrGetEqual(hlCif, tmCif, ilCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            '    If (ilRet = BTRV_ERR_NONE) Then
            '        tmCpfSrchKey.lCode = tmCif.lcpfCode
            '        ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            '        If (ilRet = BTRV_ERR_NONE) Then
            '            tmLst.sISCI = tmCpf.sISCI
            '        Else
            '            tmLst.sISCI = ""
            '        End If
            '    Else
            '        tmLst.sISCI = ""
            '    End If
            'Else
            '    tmLst.sISCI = ""
            'End If
            tmLst.sISCI = tmCpf.sISCI
            tmLst.iWkNo = tmOdf.iWkNo
            tmLst.iBreakNo = tmOdf.iBreakNo
            tmLst.iPositionNo = tmOdf.iPositionNo
            tmLst.iSeqNo = tmOdf.iSeqNo
            tmLst.sZone = tmOdf.sZone
            tmLst.sCart = ""
            'If tmCif.iMcfCode > 0 Then
            '    If (tmMcf.iCode <> tmCif.iMcfCode) Then
            '        tmMcfSrchKey.iCode = tmCif.iMcfCode
            If tmCif.iMcfCode > 0 Then
                If (tmMcf.iCode <> tmCif.iMcfCode) Then
                    tmMcfSrchKey.iCode = tmCif.iMcfCode
                    ilRet = btrGetEqual(hlMcf, tmMcf, imMcfRecLen, tmMcfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    If ilRet <> BTRV_ERR_NONE Then
                        tmMcf.sName = "C"
                        tmMcf.sPrefix = "C"
                    End If
                End If
            Else
                tmMcf.iCode = 0
                tmMcf.sName = ""
                tmMcf.sPrefix = ""
            End If
            If (tgSpf.sUseCartNo <> "N") And (tmCif.iMcfCode <> 0) Then
                If Trim$(tmCif.sCut) = "" Then
                    tmLst.sCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & " "
                Else
                    tmLst.sCart = Trim$(tmMcf.sPrefix) & Trim$(tmCif.sName) & "-" & Trim$(tmCif.sCut) & " "
                End If
            Else
                '12/29/08: Use Reel number if Wegener or OLA
                If (ilWegenerOLA) Then
                    tmLst.sCart = tmCif.sReel
                End If
            End If
            tmLst.lcpfCode = tmCif.lcpfCode
            tmLst.lCrfCsfcode = tmOdf.lCefCode
            ilVff = gBinarySearchVff(ilVefCode)
            If ilVff <> -1 Then
                If tgVff(ilVff).sHideCommOnWeb = "Y" Then
                    tmLst.lCrfCsfcode = 0
                End If
            End If
            ''Test if this is a replacement or new
            'If ilClearLstSdf Then
            '    Do
            '        tmLstSrchKey.lCode = tmSdf.lCode
            '        ilRet = btrGetEqual(hmLst, tlLst, imLstRecLen, tmLstSrchKey, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
            '        If (ilRet = BTRV_ERR_NONE) Then
            '            tmLstSrchKey.lCode = tlLst.lCode
            '            ilRet = btrGetEqual(hmLst, tlLst, imLstRecLen, tmLstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
            '            If (ilRet = BTRV_ERR_NONE) Then
            '                ilRet = btrDelete(hmLst)
            '            Else
            '                Exit Do
            '            End If
            '        Else
            '            Exit Do
            '        End If
            '    Loop
            'End If
            tmLst.lCifCode = tmOdf.lCifCode
            tmLst.ianfCode = tmOdf.ianfCode
            tmLst.lEvtIDCefCode = tmOdf.lEvtIDCefCode
            tmLst.sSplitNetwork = tmOdf.sSplitNetwork
            tmLst.sImportedSpot = "N"
            tmLst.lRafCode = 0
            If (tmLst.sSplitNetwork = "P") Or (tmLst.sSplitNetwork = "S") Then
                tmLst.lRafCode = tmClf.lRafCode
            End If
            tmLst.lBkoutLstCode = 0
            ilRet = btrInsert(hmLst, tmLst, imLstRecLen, INDEXKEY0)
            gLogBtrError ilRet, "mCreateLst: Insert"
            ''7/20/12: Only create alert if exported previously
            'blProvider = False
            'ilRet = gBinarySearchVpf(ilVefCode)
            'If ilRet <> -1 Then
            '    If tgVpf(ilRet).iCommProvArfCode > 0 Then
            '        blProvider = True
            '    End If
            'End If
            'If ((ilExportType > 0) Or (blProvider)) And (slAlertStatus <> "A") Then
            '    slAlertStatus = "A"
            '    slDate = Format$(llDate, "m/d/yy")
            '    ilFound = False
            '    If gAlertFound("L", "S", 0, ilVefCode, slDate) Then
            '        ilFound = True
            '    Else
            '        If gAlertFound("L", "C", 0, ilVefCode, slDate) Then
            '            ilFound = True
            '        End If
            '    End If
            '    If (ilFound) Or (slLogType = "R") Then
            '        'Checking if date bewteen todays date and last log is not required
            '        'It is Ok if Alert set on when generating Final Logs (Normally ilCreateLST would be true)
            '        If ilExportType > 0 Then
            '            gMakeExportAlert ilVefCode, llDate, "S"
            '        End If
            '        'Only update ISCI Alert if Provider defined and not embedded or Provider and Produce defined and embedded
            '        'Same logic in affiliate when showing vehicles to generate ISCI for.
            '        ilRet = gBinarySearchVpf(ilVefCode)
            '        If ilRet <> -1 Then
            '            If tgVpf(ilRet).iCommProvArfCode > 0 Then
            '                gMakeExportAlert ilVefCode, llDate, "I"
            '            End If
            '        End If
            '    End If
            'End If
            mCreateAlert ilVefCode, slLogType, ilExportType, slAlertStatus, llDate
        'End If
    'End If
    ilClearLstSdf = False
    mBuildAbfInfo ilVefCode, llDate
End Sub
'
'
'               mDPDaystimes - Format Days & Times to print for each
'                   spot based on Schedule line, overrides and DP.
'
'               created:  1/17/99 (The DP descriptions prints same as BR)
'
'
'               11-11-09 Get the line remark code for a spot.
Sub mDPDaysTimes(hlRdf As Integer, slEDIDays As String, llEvtCefCode As Long)
Dim ilRet As Integer
Dim ilLoop2 As Integer
Dim ilDay As Integer
ReDim ilInputDays(0 To 6) As Integer
Dim slStartTime As String
Dim slEndTime As String
Dim ilShowOVDays As Integer
Dim ilShowOVTimes As Integer
Dim slRemove As String
Dim slStr As String
Dim slTempDays As String
Dim tlClfSrchKey As CLFKEY0 'CLF key record image
Dim ilClfRecLen As Integer  'CLF record length
Dim tlRdfSrchKey As INTKEY0
Dim tlRdf As RDF
Dim ilRdfRecLen As Integer
    tlClfSrchKey.lChfCode = tmSdf.lChfCode
    tlClfSrchKey.iLine = tmSdf.iLineNo
    tlClfSrchKey.iCntRevNo = 32000 ' Plug with very high number
    tlClfSrchKey.iPropVer = 32000 ' Plug with very high number
    ilClfRecLen = Len(tmClf)
    ilRet = btrGetGreaterOrEqual(hmClf, tmClf, ilClfRecLen, tlClfSrchKey, INDEXKEY0, BTRV_LOCK_NONE)
    Do While (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) And ((tmClf.sSchStatus <> "M") And (tmClf.sSchStatus <> "F"))  'And (tmClf.sSchStatus = "A")
        ilRet = btrGetNext(hmClf, tmClf, ilClfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
    Loop
    If (ilRet = BTRV_ERR_NONE) And (tmClf.lChfCode = tmSdf.lChfCode) And (tmClf.iLine = tmSdf.iLineNo) Then
        ilRet = gGetSpotFlight(tmSdf, tmClf, hmCff, hmSmf, tmCff)
        If ilRet Then
            If tmCff.sDyWk = "W" Then            'weekly
                For ilDay = 0 To 6 Step 1
                    If tmCff.iDay(ilDay) > 0 Or tmCff.sXDay(ilDay) = "1" Then
                        ilInputDays(ilDay) = True
                    Else
                        ilInputDays(ilDay) = False
                    End If
                Next ilDay
            Else                                        'daily
                For ilDay = 0 To 6 Step 1
                    If tmCff.iDay(ilDay) > 0 Then
                        ilInputDays(ilDay) = True
                    Else
                        ilInputDays(ilDay) = False
                    End If
                Next ilDay
            End If

            ilRdfRecLen = Len(tlRdf)
            tlRdfSrchKey.iCode = tmClf.iRdfCode
            ilRet = btrGetEqual(hlRdf, tlRdf, ilRdfRecLen, tlRdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching r/c recd
            If ilRet = BTRV_ERR_NONE Then
                For ilLoop2 = 0 To 6 Step 1
                    'If tlRdf.sWkDays(7, ilLoop2 + 1) = "Y" Then             'is DP is a valid day
                    If tlRdf.sWkDays(6, ilLoop2) = "Y" Then             'is DP is a valid day
                        If ilInputDays(ilLoop2) = 0 Then         'is flight a valid day? 0=invalid day
                            ilShowOVDays = True
                            Exit For
                        Else
                            ilShowOVDays = False
                        End If
                    End If
                Next ilLoop2


                'format the Days of the week, then remove commas and blanks
                slTempDays = gDayNames(tmCff.iDay(), tmCff.sXDay(), 2, slEDIDays)
                slStr = " "
                For ilLoop2 = 1 To Len(slTempDays)
                    slRemove = Mid$(slTempDays, ilLoop2, 1)
                    If slRemove <> " " And slRemove <> "," Then
                        slStr = Trim$(slStr) & Trim$(slRemove)
                    End If
                Next ilLoop2

                'Times
                ilShowOVTimes = False
                If ((tmClf.iStartTime(0) <> 1) Or (tmClf.iStartTime(1) <> 0)) And ((tmClf.iEndTime(0) <> 1) Or (tmClf.iEndTime(1) <> 0)) Then
                    gUnpackTime tmClf.iStartTime(0), tmClf.iStartTime(1), "A", "1", slStartTime
                    gUnpackTime tmClf.iEndTime(0), tmClf.iEndTime(1), "A", "1", slEndTime
                    ilShowOVTimes = True
                Else
                    'Add times
                    For ilLoop2 = LBound(tlRdf.iStartTime, 2) To UBound(tlRdf.iStartTime, 2) Step 1 'Row
                        If (tlRdf.iStartTime(0, ilLoop2) <> 1) Or (tlRdf.iStartTime(1, ilLoop2) <> 0) Then
                            gUnpackTime tlRdf.iStartTime(0, ilLoop2), tlRdf.iStartTime(1, ilLoop2), "A", "1", slStartTime
                            gUnpackTime tlRdf.iEndTime(0, ilLoop2), tlRdf.iEndTime(1, ilLoop2), "A", "1", slEndTime
                            Exit For
                        End If
                    Next ilLoop2
                End If
                If ilShowOVDays Or ilShowOVTimes Then
                    tmOdf.sDPDesc = Trim$(slStr) & " " & Trim$(slStartTime) & "-" & Trim$(slEndTime)
                Else
                    tmOdf.sDPDesc = tlRdf.sName
                End If
                tmOdf.iRdfSortCode = tlRdf.iSortCode
                tmOdf.lClfCode = tmClf.lCode                '1-17-14 save line reference to get to audio and header type
            Else                    'missing DP record
                tmOdf.sDPDesc = "Missing Daypart"
            End If
        End If
        llEvtCefCode = tmClf.lCxfCode       '11-11-09 program events (type 14 and greater will use llEvtCefCode for its event comment
    End If                      'ilret = BTRV_err_none
End Sub
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainCifCode                  *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Copy                    *
'*                                                     *
'*******************************************************
Private Function mObtainCifCode(tlSdf As SDF, slZone As String, hlTzf As Integer, ilOther As Integer) As Long
'
'   mObtainCifCode
'       Where:
'           tlSdf(I)- Spot record
'           slZone(I)-Zone
'           hlTzf(I)- TZF Handle
'           ilOther(O)- Output for other
'           mObtainCifCode(O)- Cif Code or Zero if not found
'
'           tmVef(I)- Vehicle record
'
    Dim ilIndex As Integer
    Dim ilRet As Integer
    Dim tlTzf As TZF
    Dim ilTzfRecLen As Integer
    Dim tlTzfSrchKey As LONGKEY0
    ilOther = True
    If tlSdf.sPtType = "1" Then  '  Single Copy
        ' Read CIF using lCopyCode from SDF
        mObtainCifCode = tlSdf.lCopyCode
        Exit Function
    ElseIf tlSdf.sPtType = "2" Then  '  Combo Copy
    ElseIf tlSdf.sPtType = "3" Then  '  Time Zone Copy
        ' Read TZF using lCopyCode from SDF
        ilTzfRecLen = Len(tlTzf)
        tlTzfSrchKey.lCode = tlSdf.lCopyCode
        ilRet = btrGetEqual(hlTzf, tlTzf, ilTzfRecLen, tlTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ' Look for the first positive lZone value
        For ilIndex = 1 To 6 Step 1
            If tlTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                If StrComp(tlTzf.sZone(ilIndex - 1), slZone, 1) = 0 Then
                    ' Read CIF using lCopyCode from SDF
                    mObtainCifCode = tlTzf.lCifZone(ilIndex - 1)
                    If StrComp(slZone, "Oth", 1) <> 0 Then
                        ilOther = False
                    End If
                    Exit Function
                End If
            End If
        Next ilIndex
        For ilIndex = 1 To 6 Step 1
            If tlTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                If StrComp(tlTzf.sZone(ilIndex - 1), "Oth", 1) = 0 Then
                    ' Read CIF using lCopyCode from SDF
                    mObtainCifCode = tlTzf.lCifZone(ilIndex - 1)
                    Exit Function
                End If
            End If
        Next ilIndex
    End If
    mObtainCifCode = 0
    Exit Function
End Function
'*******************************************************
'*                                                     *
'*      Procedure Name:mObtainCrfCsfCode               *
'*                                                     *
'*             Created:3/01/94       By:D. LeVine      *
'*            Modified:              By:               *
'*                                                     *
'*            Comments: Obtain Rotation Comment Code   *
'*                                                     *
'*******************************************************
Private Function mObtainCrfCsfCode(tlSdf As SDF, slZone As String, hlCrf As Integer, hlTzf As Integer, hlCvf As Integer, ilCrfVefCode As Integer, ilInPkgVefCode As Integer, ilInLnVefCode As Integer, ilInAirVefCode As Integer, llCrfCode As Long) As Long
'
'   mObtainCrfCsfCode
'       Where:
'           tlSdf(I)- Spot record
'           hlCrf(I)- CRF Handle
'           hlTzf(I)- TZF Handle
'           mObtainCrfCsfCode(O)- Crf CsfCode or Zero if not found
'
    Dim ilRet As Integer
    Dim ilIndex As Integer
    'Time zone
    Dim tlTzf As TZF
    Dim ilTzfRecLen As Integer
    Dim tlTzfSrchKey As LONGKEY0
    'Copy rotation record information
    Dim tlCrfSrchKey1 As CRFKEY1 'CRF key record image
    Dim ilCrfRecLen As Integer  'CRF record length
    Dim tlCrf As CRF            'CRF record image
    Dim ilRotNo As Integer
    Dim ilVefCode As Integer
    Dim slType As String
    Dim ilLnVefCode As Integer
    Dim ilPkgVefCode As Integer
    Dim ilAirVefCode As Integer

    mObtainCrfCsfCode = 0
    ilLnVefCode = ilInLnVefCode
    ilAirVefCode = ilInAirVefCode
    ilPkgVefCode = ilInPkgVefCode
    llCrfCode = 0
    ilRotNo = -1
    If tlSdf.sPtType = "3" Then  '  Time zone
        ilTzfRecLen = Len(tlTzf)
        tlTzfSrchKey.lCode = tlSdf.lCopyCode
        ilRet = btrGetEqual(hlTzf, tlTzf, ilTzfRecLen, tlTzfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
        ' Look for the first positive lZone value
        For ilIndex = 1 To 6 Step 1
            If tlTzf.lCifZone(ilIndex - 1) > 0 Then ' Process just the first positive Zone
                If StrComp(tlTzf.sZone(ilIndex - 1), slZone, 1) = 0 Then
                    ilRotNo = tlTzf.iRotNo(ilIndex - 1)
                    Exit For
                End If
            End If
        Next ilIndex
        If ilRotNo = -1 Then
            Exit Function
        End If
    Else
        If (tlSdf.sPtType <> "1") And (tlSdf.sPtType <> "2") Then
            Exit Function
        End If
        ilRotNo = tlSdf.iRotNo
    End If
    ilCrfRecLen = Len(tlCrf)
    ilVefCode = ilCrfVefCode
    Do
        slType = "A"
        If tlSdf.sSpotType = "O" Then           'determine if open/close bb
            slType = "O"
        ElseIf tlSdf.sSpotType = "C" Then
            slType = "C"
        End If

        tlCrfSrchKey1.sRotType = slType
        tlCrfSrchKey1.iEtfCode = 0
        tlCrfSrchKey1.iEnfCode = 0
        tlCrfSrchKey1.iAdfCode = tlSdf.iAdfCode
        tlCrfSrchKey1.lChfCode = tlSdf.lChfCode
        tlCrfSrchKey1.lFsfCode = 0
        tlCrfSrchKey1.iVefCode = ilVefCode   'tlSdf.iVefCode
        tlCrfSrchKey1.iRotNo = ilRotNo  '32000
        'ilRet = btrGetGreaterOrEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get last current record to obtain date
        'Do While (ilRet = BTRV_ERR_NONE) And (tlCrf.sRotType = slType) And (tlCrf.iEtfCode = 0) And (tlCrf.iEnfCode = 0) And (tlCrf.iAdfCode = tlSdf.iAdfCode) And (tlCrf.lChfCode = tlSdf.lChfCode) And (tlCrf.iVefCode = ilVefCode)    'tlSdf.iVefCode)
        '    If ilRotNo = tlCrf.iRotNo Then
        '        mObtainCrfCsfCode = tlCrf.lCsfCode
        '        Exit Function
        '    End If
        '    ilRet = btrGetNext(hlCrf, tlCrf, ilCrfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        'Loop
        ilRet = btrGetEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            mObtainCrfCsfCode = tlCrf.lCsfCode
            llCrfCode = tlCrf.lCode
            Exit Function
        End If
        If ilPkgVefCode > 0 Then
            ilVefCode = ilPkgVefCode
            ilPkgVefCode = 0
        ElseIf ilAirVefCode > 0 Then
            ilVefCode = ilAirVefCode
            ilAirVefCode = 0
        Else
            If (ilCrfVefCode = ilLnVefCode) Or (ilLnVefCode = 0) Then
                Exit Do
            End If
            ilVefCode = ilLnVefCode
            ilLnVefCode = 0
        End If
    Loop While ilVefCode > 0
    ilLnVefCode = ilInLnVefCode
    ilAirVefCode = ilInAirVefCode
    ilPkgVefCode = ilInPkgVefCode
    ilVefCode = ilCrfVefCode
    Do
        slType = "A"
        If tlSdf.sSpotType = "O" Then           'determine if open/close bb
            slType = "O"
        ElseIf tlSdf.sSpotType = "C" Then
            slType = "C"
        End If

        tlCrfSrchKey1.sRotType = slType
        tlCrfSrchKey1.iEtfCode = 0
        tlCrfSrchKey1.iEnfCode = 0
        tlCrfSrchKey1.iAdfCode = tlSdf.iAdfCode
        tlCrfSrchKey1.lChfCode = tlSdf.lChfCode
        tlCrfSrchKey1.lFsfCode = 0
        tlCrfSrchKey1.iVefCode = 0
        tlCrfSrchKey1.iRotNo = ilRotNo  '32000
        ilRet = btrGetEqual(hlCrf, tlCrf, ilCrfRecLen, tlCrfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)
        If ilRet = BTRV_ERR_NONE Then
            If gCheckCrfVehicle(ilVefCode, tlCrf, hlCvf) Then
                mObtainCrfCsfCode = tlCrf.lCsfCode
                llCrfCode = tlCrf.lCode
                Exit Function
            End If
        End If
        If ilPkgVefCode > 0 Then
            ilVefCode = ilPkgVefCode
            ilPkgVefCode = 0
        ElseIf ilAirVefCode > 0 Then
            ilVefCode = ilAirVefCode
            ilAirVefCode = 0
        Else
            If (ilCrfVefCode = ilLnVefCode) Or (ilLnVefCode = 0) Then
                Exit Do
            End If
            ilVefCode = ilLnVefCode
            ilLnVefCode = 0
        End If
    Loop While ilVefCode > 0
    Exit Function
End Function
'
'
'           Set flag for ODF if spot is a Bonus or Fill;
'           otherwise initialize field to blank.  REquired to show
'           on CP
'
'           D Hosaka 2-15-01
'           3-22-03 Determine fill/bonus by advt (not spot)
'
Function mSetBonusFlag(tlSdf As SDF) As String
Dim slBonus As String * 1
Dim slBonusOnInv As String
Dim ilLoop As Integer

    slBonus = ""

    If tlSdf.sSpotType = "X" Then           '2-15-01 flag to denote on Log
        slBonusOnInv = "Y"              'if no advt found, assume to show on Inv
        ilLoop = gBinarySearchAdf(tlSdf.iAdfCode)
        If ilLoop <> -1 Then
            slBonusOnInv = tgCommAdf(ilLoop).sBonusOnInv
        End If
        'If tlSdf.sPriceType <> "N" Then
        If tlSdf.sPriceType = "+" Then
            slBonus = "B"
        ElseIf tlSdf.sPriceType = "-" Then
            slBonus = "F"
        Else
            If slBonusOnInv = "Y" Then
                slBonus = "B"
            Else
                slBonus = "F"
            End If
        End If
    End If
    mSetBonusFlag = slBonus
End Function

Private Sub mDeleteUnusedLST(slLogType As String, ilExportType As Integer)
    Dim ilRet As Integer
    Dim ilLoop As Integer
    Dim llDate As Long
    Dim slAlertStatus As String

    For ilLoop = 0 To UBound(tmLstCode) - 1 Step 1
        If tmLstCode(ilLoop).lCode > 0 Then
            tmLstSrchKey.lCode = tmLstCode(ilLoop).lCode
            ilRet = btrGetEqual(hmLst, tmLst, imLstRecLen, tmLstSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORWRITE)   'Get last current record to obtain date
            If (ilRet = BTRV_ERR_NONE) Then
                ilRet = btrDelete(hmLst)
                If ilRet <> BTRV_ERR_NONE Then
                    gLogBtrError ilRet, "mDeleteUnusedLST: Delete"
                Else
                    slAlertStatus = "C"
                    gUnpackDateLong tmLst.iLogDate(0), tmLst.iLogDate(1), llDate
                    mCreateAlert tmLst.iLogVefCode, slLogType, ilExportType, slAlertStatus, llDate
                End If
            Else
                gLogBtrError ilRet, "mDeleteUnusedLST: GetEqual"
            End If
        End If
    Next ilLoop

End Sub

Public Function gFindBBSpot(hlSdf As Integer, slType As String, ilVefCode As Integer, llChfCode As Long, ilLineNo As Integer, llDate As Long, llTime As Long, tlSdf As SDF, llPrevFdBBSpots() As Long) As Integer
    Dim ilRet As Integer
    Dim ilDate0 As Integer
    Dim ilDate1 As Integer
    Dim ilLoop As Integer

    gFindBBSpot = False
    imSdfRecLen = Len(tlSdf)
    tmSdfSrchKey0.iVefCode = ilVefCode
    tmSdfSrchKey0.lChfCode = llChfCode
    tmSdfSrchKey0.iLineNo = ilLineNo
    tmSdfSrchKey0.lFsfCode = 0
    gPackDateLong llDate, tmSdfSrchKey0.iDate(0), tmSdfSrchKey0.iDate(1)
    ilDate0 = tmSdfSrchKey0.iDate(0)
    ilDate1 = tmSdfSrchKey0.iDate(1)
    tmSdfSrchKey0.sSchStatus = "S"
    gPackTimeLong llTime, tmSdfSrchKey0.iTime(0), tmSdfSrchKey0.iTime(1)
    If slType = "O" Then
        ilRet = btrGetLessOrEqual(hlSdf, tlSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    Else
        ilRet = btrGetGreaterOrEqual(hlSdf, tlSdf, imSdfRecLen, tmSdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
    End If
    Do While (ilRet = BTRV_ERR_NONE) And (tlSdf.iVefCode = ilVefCode) And (tlSdf.iDate(0) = ilDate0) And (tlSdf.iDate(1) = ilDate1)
        If (tlSdf.sSpotType = slType) And (tlSdf.lChfCode = llChfCode) And (tlSdf.iLineNo = ilLineNo) And (tlSdf.iVefCode = ilVefCode) Then
            For ilLoop = 0 To UBound(llPrevFdBBSpots) - 1 Step 1
                If tlSdf.lCode = llPrevFdBBSpots(ilLoop) Then
                    Exit Do
                End If
            Next ilLoop
            gFindBBSpot = True
            llPrevFdBBSpots(UBound(llPrevFdBBSpots)) = tlSdf.lCode
            ReDim Preserve llPrevFdBBSpots(0 To UBound(llPrevFdBBSpots) + 1) As Long
            Exit Do
        End If
        If slType = "O" Then
            ilRet = btrGetPrevious(hlSdf, tlSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Else
            ilRet = btrGetNext(hlSdf, tlSdf, imSdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        End If
    Loop
End Function

'
'           gGenODFDay - generate the ODF events for a given date span for one or more vehicles
'           <input> Form - name of form calling this routine
'                   slInputStartDate - earliest date to create ODF
'                   slInputEndDate - latest date to create ODF
'                   tlLbcSelection - list box of vehicles
'                   lbcSort - list sorted vehicles
'                   tlSortCode() - vehicle codes (sorted by vehicle name)
'                   slUseLocalOrFeed - "L" for local time conversions from Zone table, or
'                                      "F" for feed time conversion from Zone table
'                   slSource - "E" for Export Enco; "C" for Copy Book
'           <return> -
'           <output> ODF records
Public Function gGenODFDay(Form As Form, slInputStartDate As String, slInputEndDate As String, tlLbcSelection As Control, tlLbcSort As Control, tlSortCode() As SORTCODE, slUseLocalOrFeed As String, slSource As String) As Integer
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Local Variables (Removed)                                                              *
'*  ilLogGen                      llCycleDate                   ilCycle                   *
'*  slCycleDate                   slUseAffSys                   ilUpdateLst               *
'*                                                                                        *
'******************************************************************************************

Dim ilType As Integer
Dim sLCP  As String
Dim slStartTime As String
Dim slEndTime As String
Dim ilLoop As Integer
Dim slNameCode As String
Dim slCode As String
ReDim ilEvtAllowed(0 To 14) As Integer
Dim ilVpfIndex As Integer
Dim lgStartIndex As Integer
Dim ilRet As Integer
Dim ilPass As Integer
Dim ilODFVefCode As Integer
Dim ilFound As Integer
Dim ilVef As Integer
Dim ilGenLST As Integer
Dim ilVefCode As Integer
Dim llStartDate As Long
Dim llEndDate As Long
Dim llDate As Long
Dim slStartDate As String
Dim slEndDate As String
Dim slLogType As String
Dim ilExportType As Integer
Dim ilValue As Integer
Dim ilCRet As Integer
Dim llSeasonStart As Long
Dim llSeasonEnd As Long
Dim hlRsf As Integer
Dim slNewLines(0 To 0) As String * 72   'required as parameter to gBlackoutTest


    'Open btrieve files
    imVefRecLen = Len(tmVef)    'Save VEF record length
    hmVef = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmVef, "", sgDBPath & "Vef.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenODFDayErr
    gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: VEF.BTR)", Form
    On Error GoTo 0
    imVpfRecLen = Len(tmVpf)    'Save VEF record length
    hmVpf = CBtrvTable(TWOHANDLES)          'Save VEF handle
    ilRet = btrOpen(hmVpf, "", sgDBPath & "Vpf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenODFDayErr
    gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: VPF.BTR)", Form
    On Error GoTo 0
    imVlfRecLen = Len(tmVlf)    'Save VEF record length
    hmVlf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmVlf, "", sgDBPath & "Vlf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenODFDayErr
    gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: VLF.BTR)", Form
    On Error GoTo 0

    imSsfRecLen = Len(tmSsf)    'Save VEF record length
    hmSsf = CBtrvTable(ONEHANDLE)          'Save VEF handle
    ilRet = btrOpen(hmSsf, "", sgDBPath & "Ssf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenODFDayErr
    gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: SSF.BTR)", Form
    On Error GoTo 0
    hmSdf = CBtrvTable(TWOHANDLES)          'Save VEF handle
    ilRet = btrOpen(hmSdf, "", sgDBPath & "Sdf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
    On Error GoTo gGenODFDayErr
    gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: SDF.BTR)", Form
    On Error GoTo 0

    ilValue = Asc(tgSpf.sSportInfo)
    If (ilValue And USINGSPORTS) = USINGSPORTS Then
        imGhfRecLen = Len(tmGhf)    'Save VEF record length
        hmGhf = CBtrvTable(ONEHANDLE)          'Save VEF handle
        ilRet = btrOpen(hmGhf, "", sgDBPath & "Ghf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenODFDayErr
        gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: GHF.BTR)", Form
        On Error GoTo 0
        imGsfRecLen = Len(tmGsf)    'Save GSF record length
        hmGsf = CBtrvTable(ONEHANDLE)          'Save VEF handle
        ilRet = btrOpen(hmGsf, "", sgDBPath & "Gsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenODFDayErr
        gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: GSF.BTR)", Form
        On Error GoTo 0
    End If

    If tgSpf.sCBlackoutLog = "Y" Then
        imMcfRecLen = Len(tmMcf)  'Get and save ADF record length
        hmMcf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmMcf, "", sgDBPath & "Mcf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenODFDayErr
        gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: Mcf.Btr)", Form
        On Error GoTo 0
        imCifRecLen = Len(tmCif)  'Get and save ADF record length
        hmCif = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmCif, "", sgDBPath & "Cif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenODFDayErr
        gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: Cif.Btr)", Form
        On Error GoTo 0
        imCHFRecLen = Len(tmChf)  'Get and save ADF record length
        hmCHF = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmCHF, "", sgDBPath & "Chf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenODFDayErr
        gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: Chf.Btr)", Form
        On Error GoTo 0
        imClfRecLen = Len(tmClf)  'Get and save ADF record length
        hmClf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmClf, "", sgDBPath & "Clf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenODFDayErr
        gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: Clf.Btr)", Form
        On Error GoTo 0
        imBofRecLen = Len(tmBof)  'Get and save ADF record length
        hmBof = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmBof, "", sgDBPath & "Bof.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenODFDayErr
        gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: Bof.Btr)", Form
        On Error GoTo 0
        imRsfRecLen = Len(tmRsf)    'Save  record length
        hlRsf = CBtrvTable(TWOHANDLES)          'Save RNF handle
        ilRet = btrOpen(hlRsf, "", sgDBPath & "Rsf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenODFDayErr
        gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: RSF.BTR)", Form
        On Error GoTo 0
        imPrfRecLen = Len(tmPrf)  'Get and save ADF record length
        hmPrf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmPrf, "", sgDBPath & "Prf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenODFDayErr
        gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: Prf.Btr)", Form
        On Error GoTo 0
        imSifRecLen = Len(tmSif)  'Get and save ADF record length
        hmSif = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmSif, "", sgDBPath & "Sif.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenODFDayErr
        gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: Sif.Btr)", Form
        On Error GoTo 0
        imCrfRecLen = Len(tmCrf)  'Get and save ADF record length
        hmCrf = CBtrvTable(TWOHANDLES)        'Create ADF object handle
        ilRet = btrOpen(hmCrf, "", sgDBPath & "Crf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenODFDayErr
        gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: Crf.Btr)", Form
        On Error GoTo 0
        imCnfRecLen = Len(tmCnf)  'Get and save ADF record length
        hmCnf = CBtrvTable(ONEHANDLE)        'Create ADF object handle
        ilRet = btrOpen(hmCnf, "", sgDBPath & "Cnf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        On Error GoTo gGenODFDayErr
        gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: Cnf.Btr)", Form
        On Error GoTo 0
    End If

    igGenDate(0) = igNowDate(0)     'these dates and times are to be stored in ODF, used for selection
    igGenDate(1) = igNowDate(1)
    lgGenTime = lgNowTime

    'slUseAffSys = tgSpf.sGUseAffSys     'save flag to indicate whether Affiliate system is used.  Force to NO for this reprot
    'tgSpf.sGUseAffSys = "N"


    ilType = 0          'On-air vs Alternate
    sLCP = "C"          'Current vs pending
    slLogType = "R"     'assume reprint
    ilExportType = 0            'no exporting , and no alerts to create
    slStartTime = "12M"
    slEndTime = "12M"
    ilGenLST = False                    'dont care about affiliate system in this mode
    For ilLoop = LBound(ilEvtAllowed) To UBound(ilEvtAllowed) Step 1
        ilEvtAllowed(ilLoop) = True
    Next ilLoop
    ilEvtAllowed(0) = False 'Don't include library names
    'ilevtallowed(1) = False 'Don't include programs   -- pgms must be included to gather spots
    ilEvtAllowed(10) = False 'Don't include page eject
    ilEvtAllowed(11) = False 'Don't include line skip
    ilEvtAllowed(12) = False 'Don't include line skip
    ilEvtAllowed(13) = False 'Don't include line skip
    ilEvtAllowed(14) = False 'Don't include other events

    slStartDate = slInputStartDate  'Form!edcSelCFrom.Text
    llStartDate = gDateValue(slStartDate)
    slEndDate = slInputEndDate  'Form!edcSelCTo.Text
    llEndDate = gDateValue(slEndDate)

    tlLbcSort.Clear
    If tgSpf.sCBlackoutLog = "Y" Then
        If Not mOpenMsgFile() Then
            Screen.MousePointer = vbDefault
            Exit Function
        End If

        ilRet = gReadBofRec(1, hmBof, hmCif, hmPrf, hmSif, hmCHF, "B", slStartDate, 1)
        igStartBofIndex = LBound(tgRBofRec) - 1
        ig30StartBofIndex = igStartBofIndex
        ig60StartBofIndex = igStartBofIndex

    End If
    bgLogFirstCallToVpfFind = True
    For ilLoop = 0 To tlLbcSelection.ListCount - 1 Step 1
        If (tlLbcSelection.Selected(ilLoop)) Then
            ilType = 0          'On-air vs Alternate
            sLCP = "C"          'Current vs pending
            slLogType = "R"     'assume reprint
            ilExportType = 0            'no exporting , and no alerts to create
            slStartTime = "12M"
            slEndTime = "12M"
            ilGenLST = False                    'dont care about affiliate system in this mode
            slNameCode = tlSortCode(ilLoop).sKey 'Form!lbcAirNameCode.List(ilVehicle)
            ilRet = gParseItem(slNameCode, 2, "\", slCode)
            tmVefSrchKey.iCode = Val(slCode)
            ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
            ilVefCode = tmVef.iCode
            If bgLogFirstCallToVpfFind Then
                ilVpfIndex = gVpfFind(Form, tmVef.iCode)
                bgLogFirstCallToVpfFind = False
            Else
                ilVpfIndex = gVpfFindIndex(tmVef.iCode)
            End If

            ReDim tmLogGen(0 To 1) As LOGGEN
            tmLogGen(0).iGenVefCode = tmVef.iCode
            tmLogGen(0).iSimVefCode = tmVef.iCode
            '3/26/15: Required by gBuildODFDay
            ReDim tgOdfSdfCodes(0 To 0) As ODFSDFCODES
            ReDim tgSpotSum(0 To 0) As SPOTSUM
            lgStartIndex = UBound(tgSpotSum)

            If tgSpf.sUsingBBs = "Y" Then
                'Determine vehicles to create Billboard spots
                ilRet = gMakeBBAndAssignCopy(hmSdf, hmVlf, ilVefCode, llStartDate, llEndDate)
            End If

            For ilPass = 0 To 1 Step 1      'look for the combined vehicle in 2nd pass, if it exists
                If ilPass = 1 Then          'look for the vehicle to merge spots with
                    If tmVef.iCombineVefCode <= 0 Then      'no need to combine any other vehicle
                        Exit For
                    End If
                    ilFound = False
                    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        If (tgMVef(ilVef).iCode = tmVef.iCombineVefCode) Then
                            If (tgMVef(ilVef).sType = "C") Or (tgMVef(ilVef).sType = "A") Then
                                ilFound = True
                                ilODFVefCode = tmVef.iCode  'Retain veCode that ODF should be created within
                                tmVef = tgMVef(ilVef)
                            End If
                            Exit For
                        End If
                    Next ilVef
                    If Not ilFound Then
                        Exit For
                    End If
                    ilVefCode = tmVef.iCode
                    If bgLogFirstCallToVpfFind Then
                        ilVpfIndex = gVpfFind(Form, tmVef.iCode)
                        bgLogFirstCallToVpfFind = False
                    Else
                        ilVpfIndex = gVpfFindIndex(tmVef.iCode)
                    End If
                Else
                    ilODFVefCode = 0
                End If
                If tmVef.sType = "L" Then

                    For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                        If (tgMVef(ilVef).sType = "C") And (tgMVef(ilVef).iVefCode = tmVef.iCode) Then
                            ilRet = gBuildODFSpotDay("L", ilType, sLCP, tgMVef(ilVef).iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, slLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, slUseLocalOrFeed, 0, 0, 0)   '11-2-06
                            If Not ilRet Then
                                gGenODFDay = False
                                Exit Function
                            End If
                        '7/27/12: Include Sports within Log vehicles
                        ElseIf (tgMVef(ilVef).sType = "G") And (tgMVef(ilVef).iVefCode = tmVef.iCode) Then
                            tmGsfSrchKey3.iVefCode = tgMVef(ilVef).iCode
                            tmGsfSrchKey3.iGameNo = 0
                            ilCRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey3, INDEXKEY3, BTRV_LOCK_NONE)
                            Do While (ilCRet = BTRV_ERR_NONE) And (tmGsf.iVefCode = tgMVef(ilVef).iCode)
                                gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
                                If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                                    If tmGsf.sGameStatus <> "C" Then
                                        ilType = tmGsf.iGameNo
                                        ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmGsf.iVefCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, slLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, slUseLocalOrFeed, 0, 0, 0)
                                        If Not ilRet Then
                                            gGenODFDay = False
                                            Exit Function
                                        End If
                                    End If
                                End If
                                ilCRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                            Loop
                            ilType = 0
                        End If
                    Next ilVef

                ElseIf tmVef.sType = "A" Then
                    gBuildLinkArray hmVlf, tmVef, slStartDate, igSVefCode() 'Build igSVefCode so that gBuildODFSpotDay can use it
                    DoEvents
                    ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmVef.iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, slLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, slUseLocalOrFeed, 0, 0, 0)   '11-2-06
                    If Not ilRet Then
                        gGenODFDay = False
                        Exit Function
                    End If
                ElseIf tmVef.sType = "G" Then
                    DoEvents
                    tmGhfSrchKey1.iVefCode = tmVef.iCode
                    ilRet = btrGetEqual(hmGhf, tmGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                    'If ilRet = BTRV_ERR_NONE Then
                    Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.iVefCode = tmVef.iCode)
                        gUnpackDateLong tmGhf.iSeasonStartDate(0), tmGhf.iSeasonStartDate(1), llSeasonStart
                        gUnpackDateLong tmGhf.iSeasonEndDate(0), tmGhf.iSeasonEndDate(1), llSeasonEnd
                        If (llEndDate >= llSeasonStart) And (llStartDate <= llSeasonEnd) Then
                            tmGsfSrchKey1.lghfcode = tmGhf.lCode
                            tmGsfSrchKey1.iGameNo = 0
                            ilRet = btrGetGreaterOrEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                            Do While (ilRet = BTRV_ERR_NONE) And (tmGhf.lCode = tmGsf.lghfcode)
                                gUnpackDateLong tmGsf.iAirDate(0), tmGsf.iAirDate(1), llDate
                                If (llDate >= llStartDate) And (llDate <= llEndDate) Then
                                    '9/21/09:  Bypass Live Log if generating export
                                    '          E=Export Enco; C=Copy Book
                                    '          For the Export, Live Log vehicles bypassed
                                    If (slSource <> "E") Or ((slSource = "E") And (((tgVpf(ilVpfIndex).sGenLog <> "L") And (tgVpf(ilVpfIndex).sGenLog <> "A")) Or ((tgVpf(ilVpfIndex).sGenLog = "A") And (tmGsf.sLiveLogMerge = "M")))) Then
                                        ilType = tmGsf.iGameNo
                                        ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmVef.iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, slLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, slUseLocalOrFeed, tmGsf.iGameNo, tmGsf.lCode, 0)   '11-2-06
                                        If tmVef.iCombineVefCode > 0 Then
                                            For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                                                If (tgMVef(ilVef).iCode = tmVef.iCombineVefCode) Then
                                                    If (tgMVef(ilVef).sType = "G") Then
                                                        ilODFVefCode = tmVef.iCode  'Retain veCode that ODF should be created within
                                                        tmGhfSrchKey1.iVefCode = tgMVef(ilVef).iCode
                                                        ilCRet = btrGetEqual(hmGhf, tmCombineGhf, imGhfRecLen, tmGhfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                                        'If ilCRet = BTRV_ERR_NONE Then
                                                        Do While (ilCRet = BTRV_ERR_NONE) And (tmCombineGhf.iVefCode = tgMVef(ilVef).iCode)
                                                            gUnpackDateLong tmCombineGhf.iSeasonStartDate(0), tmCombineGhf.iSeasonStartDate(1), llSeasonStart
                                                            gUnpackDateLong tmCombineGhf.iSeasonEndDate(0), tmCombineGhf.iSeasonEndDate(1), llSeasonEnd
                                                            If (llEndDate >= llSeasonStart) And (llStartDate <= llSeasonEnd) Then
                                                                tmGsfSrchKey1.lghfcode = tmCombineGhf.lCode
                                                                tmGsfSrchKey1.iGameNo = 0
                                                                ilCRet = btrGetGreaterOrEqual(hmGsf, tmCombineGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)
                                                                Do While (ilCRet = BTRV_ERR_NONE) And (tmCombineGhf.lCode = tmCombineGsf.lghfcode)
                                                                    If (tmGsf.iAirDate(0) = tmCombineGsf.iAirDate(0)) And (tmGsf.iAirDate(1) = tmCombineGsf.iAirDate(1)) And (tmGsf.iAirTime(0) = tmCombineGsf.iAirTime(0)) And (tmGsf.iAirTime(1) = tmCombineGsf.iAirTime(1)) Then
                                                                        ilType = tmCombineGsf.iGameNo
                                                                        ilRet = gBuildODFSpotDay("L", ilType, sLCP, tgMVef(ilVef).iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, slLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, slUseLocalOrFeed, tmGsf.iGameNo, tmGsf.lCode, 0)
                                                                        Exit For
                                                                    End If
                                                                    ilCRet = btrGetNext(hmGsf, tmCombineGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                                                Loop
                                                            End If
                                                        Loop
                                                        ilCRet = btrGetNext(hmGhf, tmCombineGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                                                    End If
                                                    Exit For
                                                End If
                                            Next ilVef
                                            'Reset key so that getNext works
                                            tmGsfSrchKey1.lghfcode = tmGhf.lCode
                                            tmGsfSrchKey1.iGameNo = tmGsf.iGameNo
                                            ilCRet = btrGetEqual(hmGsf, tmGsf, imGsfRecLen, tmGsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORWRITE)
                                        End If
                                    End If
                                End If
                                ilRet = btrGetNext(hmGsf, tmGsf, imGsfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                            Loop
                        End If
                        ilRet = btrGetNext(hmGhf, tmGhf, imGhfRecLen, BTRV_LOCK_NONE, SETFORWRITE)
                    Loop
                Else
                    DoEvents
                    ilRet = gBuildODFSpotDay("L", ilType, sLCP, tmVef.iCode, slStartDate, slEndDate, slStartTime, slEndTime, ilEvtAllowed(), 0, slLogType, hmLst, hmMcf, ilGenLST, ilExportType, ilODFVefCode, slUseLocalOrFeed, 0, 0, 0)       '11-2-06
                    If Not ilRet Then
                        gGenODFDay = False
                        Exit Function
                    End If
                    'Simulcast Vehicle Log Generation array creation
                    'Removed 1/6/99: Shadow request- to reinstate remove comments on lines below
                    '                Code has previously been tested and it works to generate Logs
                    '                for all simulcast vehicles
                    'For ilVef = LBound(tgMVef) To UBound(tgMVef) - 1 Step 1
                    '    If (tgMVef(ilVef).sType = "T") And (tgMVef(ilVef).iVefCode = tmVef.iCode) And (tgMVef(ilVef).sState = "A") Then
                    '        tmLogGen(UBound(tmLogGen)).iGenVefCode = tmVef.iCode
                    '        tmLogGen(UBound(tmLogGen)).iSimVefCode = tgMVef(ilVef).iCode
                    '        ReDim Preserve tmLogGen(0 To UBound(tmLogGen) + 1) As LOGGEN
                    '    End If
                    'Next ilVef
                End If

                If ilPass = 1 Then
                    slNameCode = tlSortCode(ilLoop).sKey 'Form!lbcAirNameCode.List(ilVehicle)
                    ilRet = gParseItem(slNameCode, 2, "\", slCode)
                    tmVefSrchKey.iCode = Val(slCode)
                    ilRet = btrGetEqual(hmVef, tmVef, imVefRecLen, tmVefSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
                    ilVefCode = tmVef.iCode
                    If bgLogFirstCallToVpfFind Then
                        ilVpfIndex = gVpfFind(Form, tmVef.iCode)
                        bgLogFirstCallToVpfFind = False
                    Else
                        ilVpfIndex = gVpfFindIndex(tmVef.iCode)
                    End If
                End If
            Next ilPass
            If (tgSpf.sCBlackoutLog = "Y") And (tmVef.sType <> "G") Then
                imOdfRecLen = Len(tmOdf)    'Save ODF record length
                hmOdf = CBtrvTable(ONEHANDLE)          'Save odf handle
                ilRet = btrOpen(hmOdf, "", sgDBPath & "Odf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                On Error GoTo gGenODFDayErr
                gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: ODF.BTR)", Form
                On Error GoTo 0
                hmCvf = CBtrvTable(ONEHANDLE)          'Save odf handle
                ilRet = btrOpen(hmCvf, "", sgDBPath & "Cvf.btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
                On Error GoTo gGenODFDayErr
                gBtrvErrorMsg ilRet, "gGenODFDay (btrOpen: CVF.BTR)", Form
                On Error GoTo 0
                lgEndIndex = UBound(tgSpotSum)
                gBlackoutTest 1, hmCif, hmMcf, hmOdf, hlRsf, hmCpf, hmCrf, hmCnf, hmClf, hmLst, hmCvf, slNewLines(), hmMsg, tlLbcSort
                ilRet = btrClose(hmOdf)
                btrDestroy hmOdf
                ilRet = btrClose(hmCvf)
                btrDestroy hmCvf
            End If
        End If
    Next ilLoop
    'tgSpf.sGUseAffSys = slUseAffSys     'restore Using Affiliate system flag to original state
    
    On Error Resume Next
    Erase tgOdfSdfCodes
    Erase tgSpotSum
    
    'close all files
    ilRet = btrClose(hmSdf)
    btrDestroy hmSdf
    ilRet = btrClose(hmSsf)
    btrDestroy hmSsf

    ilRet = btrClose(hmVlf)
    btrDestroy hmVlf
    ilRet = btrClose(hmVpf)
    btrDestroy hmVpf
    ilRet = btrClose(hmVef)
    btrDestroy hmVef

    ilValue = Asc(tgSpf.sSportInfo)
    If (ilValue And USINGSPORTS) = USINGSPORTS Then
        ilRet = btrClose(hmGhf)
        btrDestroy hmGhf
        ilRet = btrClose(hmGsf)
        btrDestroy hmGsf
    End If

    If tgSpf.sCBlackoutLog = "Y" Then
        ilRet = btrClose(hmMcf)
        btrDestroy hmMcf
        ilRet = btrClose(hmCif)
        btrDestroy hmCif
        ilRet = btrClose(hmCHF)
        btrDestroy hmCHF
        ilRet = btrClose(hmClf)
        btrDestroy hmClf
        ilRet = btrClose(hmBof)
        btrDestroy hmBof
        ilRet = btrClose(hmPrf)
        btrDestroy hmPrf
        ilRet = btrClose(hmSif)
        btrDestroy hmSif
        ilRet = btrClose(hmCrf)
        btrDestroy hmCrf
        ilRet = btrClose(hmCnf)
        btrDestroy hmCnf
        ilRet = btrClose(hlRsf)
        btrDestroy hlRsf
        'Print #hmMsg, "Copy Book Completed: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        gAutomationAlertAndLogHandler "Copy Book Completed: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
        Close #hmMsg

    End If
    Erase tgOdfSdfCodes
    
    Exit Function

gGenODFDayErr:
    gGenODFDay = False
    Exit Function
End Function

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
Private Function mOpenMsgFile()
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim llNowDate As Long

    'On Error GoTo mOpenMsgFileErr:

    slDate = Format$(gNow(), "m/d/yy") 'Correctly format current date
    llNowDate = gDateValue(slDate)

    slToFile = sgDBPath & "Messages\" & "CopyBook" & CStr(tgUrf(0).iCode) & ".Txt"
    sgMessageFile = slToFile
    'slDateTime = gFileDateTime(slToFile)
    ilRet = gFileExist(slToFile)
    If ilRet = 0 Then
        slDateTime = gFileDateTime(slToFile)
        slFileDate = Format$(slDateTime, "m/d/yy")
        If gDateValue(slFileDate) = llNowDate Then  'Append
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Append As hmMsg
            'ilRet = gFileOpen(slToFile, "Append", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        Else
            Kill slToFile
            On Error GoTo 0
            ilRet = 0
            'On Error GoTo mOpenMsgFileErr:
            'hmMsg = FreeFile
            'Open slToFile For Output As hmMsg
            'ilRet = gFileOpen(slToFile, "Output", hmMsg)
            If ilRet <> 0 Then
                Screen.MousePointer = vbDefault
                MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
                mOpenMsgFile = False
                Exit Function
            End If
        End If
    Else
        On Error GoTo 0
        ilRet = 0
        'On Error GoTo mOpenMsgFileErr:
        'hmMsg = FreeFile
        'Open slToFile For Output As hmMsg
        'ilRet = gFileOpen(slToFile, "Output", hmMsg)
        If ilRet <> 0 Then
            Screen.MousePointer = vbDefault
            MsgBox "Open " & slToFile & ", Error #" & str$(ilRet), vbOkOnly + vbCritical + vbApplicationModal, "Open Error"
            mOpenMsgFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0
    'Print #hmMsg, ""
    
    'Print #hmMsg, "Copy Book: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    gAutomationAlertAndLogHandler "Copy Book: " & Format$(gNow(), "m/d/yy") & " at " & Format$(gNow(), "h:mm:ssAM/PM")
    mOpenMsgFile = True
    Exit Function
'mOpenMsgFileErr:
'    ilRet = Err.Number
'    Resume Next
End Function

'********************************************************
'*      Procedure Name: gAutomationLogMessage           *
'*             Created: 5/18/93       By:J. White       *
'*            Comments: Writes Logs, handles MessageBox *
'*                      or stores the message in Mem    *
'*                      until a Filename is computed    *
'*                                                      *
'* TTP 10342: Automation Alerting and Logging           *
'*  - Revised 7/11/2022 to close file, not use handles  *
'*                                                      *
'* If slTitle is empty (""),                            *
'*    This function is used to Write to a Log file      *
'*                                                      *
'* If slTitle is not empty (<> ""),                     *
'*  - This function is used show a message box unless   *
'*     we are in Automation mode, then the alert        *
'*     is logged instead shown to user.                 *
'*  - Use buttons to determine which message box        *
'*     buttons will be shown on the alert dialog        *
'*                                                      *
'* If sgMessageFile is empty (""),                      *
'*  - message is buffered into sgAutomationLogBuffer    *
'*                                                      *
'* If sgMessageFile is not empty (<> ""),               *
'*  - a New File Handle is created,                     *
'*     File is appended with the Date/Time and message  *
'*     The file is then Closed.                         *
'*                                                      *
'********************************************************
Public Sub gAutomationAlertAndLogHandler(sMessage As String, Optional buttons As VbMsgBoxStyle = vbInformation, Optional slTitle As String = "")
    Dim slToFile As String
    Dim slDateTime As String
    Dim slFileDate As String
    Dim ilRet As Integer
    Dim slDate As String
    Dim llNowDate As Long
    Dim hlFileHandle As Integer
    On Error GoTo AutomationAlertAndLogHandlerError
    slDate = Format$(gNow(), "m/d/yy") 'Correctly format current date
    llNowDate = gDateValue(slDate)
    slDateTime = slDate & " " & Format(Now, "hh:mm:ss AMPM")
    
    If igExportType > 1 Then
        '----------------------------------
        'Automation in Progress  (Never use Message boxes, Log the message instead)
        '----------------------------------
        'Buffer Message until LogFileName (sgMessageFile) is set
        If sgMessageFile = "" Then
            If sMessage = "" Then
                sgAutomationLogBuffer = sgAutomationLogBuffer & "-----------------------------------------------------------------------------" & vbCrLf
            Else
                sgAutomationLogBuffer = sgAutomationLogBuffer & slDateTime & " - " & sMessage & vbCrLf
            End If
            Exit Sub
        Else
            'we have a Filename; Open a New Log or Append to Existing, and Write sMessage
            If InStr(1, sgMessageFile, sgDBPath & "Messages\") > 0 Or InStr(1, sgMessageFile, sgDBPath & "\Messages\") > 0 Then
                slToFile = sgMessageFile
            Else
                slToFile = sgDBPath & "Messages\" & sgMessageFile
            End If
            If hlFileHandle > 0 Then
                Close #hlFileHandle
            End If
            ilRet = gFileExist(slToFile)
            If ilRet = 0 Then
                slFileDate = Format$(gFileDateTime(slToFile), "m/d/yy")
                If gDateValue(slFileDate) = llNowDate Then  'Append
                    ilRet = 0
                    ilRet = gFileOpen(slToFile, "Append", hlFileHandle)
                    If ilRet <> 0 Then
                        gLogMsg "Open " & slToFile & ", Error #" & str$(ilRet), "Exports.txt", False
                    End If
                Else
                    Kill slToFile
                    ilRet = 0
                    ilRet = gFileOpen(slToFile, "Output", hlFileHandle)
                    If ilRet <> 0 Then
                        gLogMsg "Open " & slToFile & ", Error #" & str$(ilRet), "Exports.txt", False
                    End If
                End If
            Else
                ilRet = 0
                ilRet = gFileOpen(slToFile, "Output", hlFileHandle)
                If ilRet <> 0 Then
                    gLogMsg "Open " & slToFile & ", Error #" & str$(ilRet), "Exports.txt", False
                End If
            End If
            If sgAutomationLogBuffer <> "" Then
                'Print the Buffered Log Data
                'If InStr(1, slDateTime, sgAutomationLogBuffer) = 0 Then sgAutomationLogBuffer = slDateTime & " - " & sgAutomationLogBuffer
                If right(sgAutomationLogBuffer, 2) = vbCrLf Then
                    Print #hlFileHandle, Mid(sgAutomationLogBuffer, 1, Len(sgAutomationLogBuffer) - 2)
                Else
                    Print #hlFileHandle, sgAutomationLogBuffer
                End If
                sgAutomationLogBuffer = ""
            End If
            
            If sMessage = "" Then
                Print #hlFileHandle, "-----------------------------------------------------------------------------"
            Else
                Print #hlFileHandle, slDateTime & " - " & sMessage
            End If
            Close #hlFileHandle
            hlFileHandle = 0
        End If
    Else
        '----------------------------------
        'No Automation (interactive mode)
        '----------------------------------
        If slTitle <> "" Then
            'This is a MsgBox
            If sMessage <> "" Then MsgBox sMessage, buttons, slTitle
        Else
            'This is a Log Entry (no Msgbox)
            If sgMessageFile = "" Then
                'Buffer Message until LogFileName (sgMessageFile) is set
                If sMessage = "" Then
                    sgAutomationLogBuffer = sgAutomationLogBuffer & "-----------------------------------------------------------------------------" & vbCrLf
                Else
                    sgAutomationLogBuffer = sgAutomationLogBuffer & slDateTime & " - " & sMessage & vbCrLf
                End If
                Exit Sub
            Else
                'we have a Filename; Open a New Log or Append to Existing, and Write sMessage
                If InStr(1, sgMessageFile, sgDBPath & "Messages\") > 0 Or InStr(1, sgMessageFile, sgDBPath & "\Messages\") > 0 Then
                    slToFile = sgMessageFile
                Else
                    slToFile = sgDBPath & "Messages\" & sgMessageFile
                End If
                
                If hlFileHandle > 0 Then
                    Close #hlFileHandle
                End If
                ilRet = gFileExist(slToFile)
                If ilRet = 0 Then
                    slFileDate = Format$(gFileDateTime(slToFile), "m/d/yy")
                    If gDateValue(slFileDate) = llNowDate Then  'Append
                        ilRet = 0
                        ilRet = gFileOpen(slToFile, "Append", hlFileHandle)
                        If ilRet <> 0 Then
                            gLogMsg "Open " & slToFile & ", Error #" & str$(ilRet), "Exports.txt", False
                        End If
                    Else
                        Kill slToFile
                        ilRet = 0
                        ilRet = gFileOpen(slToFile, "Output", hlFileHandle)
                        If ilRet <> 0 Then
                            gLogMsg "Open " & slToFile & ", Error #" & str$(ilRet), "Exports.txt", False
                        End If
                    End If
                Else
                    ilRet = 0
                    ilRet = gFileOpen(slToFile, "Output", hlFileHandle)
                    If ilRet <> 0 Then
                        gLogMsg "Open " & slToFile & ", Error #" & str$(ilRet), "Exports.txt", False
                    End If
                End If
                If sgAutomationLogBuffer <> "" Then
                    'Print the Buffered Log Data
                    If right(sgAutomationLogBuffer, 2) = vbCrLf Then
                        Print #hlFileHandle, Mid(sgAutomationLogBuffer, 1, Len(sgAutomationLogBuffer) - 2)
                    Else
                        Print #hlFileHandle, sgAutomationLogBuffer
                    End If
                    sgAutomationLogBuffer = ""
                End If
                If sMessage = "" Then
                    Print #hlFileHandle, "-----------------------------------------------------------------------------"
                Else
                    Print #hlFileHandle, slDateTime & " - " & sMessage
                End If
                Close #hlFileHandle
            End If
        End If
    End If
    Exit Sub
    
AutomationAlertAndLogHandlerError:
    'OOPS!
    Beep
End Sub

Private Sub mCreateAlert(ilVefCode As Integer, slLogType As String, ilExportType As Integer, slAlertStatus As String, llDate As Long)
    Dim ilVpf As Integer
    Dim ilFound As Integer
    Dim slDate As String
    Dim blProvider As Boolean
    '7/20/12: Only create alert if exported previously
    blProvider = False
    ilVpf = gBinarySearchVpf(ilVefCode)
    If ilVpf <> -1 Then
        If tgVpf(ilVpf).iCommProvArfCode > 0 Then
            blProvider = True
        End If
    End If
    If ((ilExportType > 0) Or (blProvider)) And (slAlertStatus <> "A") Then
        slAlertStatus = "A"
        slDate = Format$(llDate, "m/d/yy")
        ilFound = False
        If gAlertFound("L", "S", 0, ilVefCode, slDate) Then
            ilFound = True
        Else
            If gAlertFound("L", "C", 0, ilVefCode, slDate) Then
                ilFound = True
            End If
        End If
        If (ilFound) Or (slLogType = "F") Or (slLogType = "R") Or (slLogType = "A") Then 'F=Final; R=Reprint; A=Alert
            'Checking if date bewteen todays date and last log is not required
            'It is Ok if Alert set on when generating Final Logs (Normally ilCreateLST would be true)
            If ilExportType > 0 Then
                If slLogType = "F" Then
                    gMakeExportAlert ilVefCode, llDate, slLogType, "S"
                Else
                    gMakeExportAlert ilVefCode, llDate, "R", "S"
                End If
            End If
            'Only update ISCI Alert if Provider defined and not embedded or Provider and Produce defined and embedded
            'Same logic in affiliate when showing vehicles to generate ISCI for.
            ilVpf = gBinarySearchVpf(ilVefCode)
            If ilVpf <> -1 Then
                If tgVpf(ilVpf).iCommProvArfCode > 0 Then
                    If slLogType = "F" Then
                        gMakeExportAlert ilVefCode, llDate, slLogType, "I"
                    Else
                        gMakeExportAlert ilVefCode, llDate, "R", "I"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub mBuildAbfInfo(ilVefCode As Integer, llLogDate As Long)
    Dim ilRet As Integer
    Dim llUpper As Long
    Dim blFound As Boolean
    Dim llLoop As Long
    Dim llMoDate As Long
    
    On Error GoTo mBuildAbfInfoErr:
    ilRet = 0
    llUpper = UBound(tgAbfInfo)
    If ilRet = 1 Then
        ReDim tgAbfInfo(0 To 0) As ABFINFO
        llUpper = 0
    End If
    llMoDate = llLogDate
    Do While gWeekDayLong(llMoDate) <> 0
        llMoDate = llMoDate - 1
    Loop
    blFound = False
    For llLoop = 0 To llUpper - 1 Step 1
        If (tgAbfInfo(llLoop).iVefCode = ilVefCode) And (tgAbfInfo(llLoop).lMondayDate = llMoDate) Then
            blFound = True
            If llLogDate < tgAbfInfo(llLoop).lStartDate Then
                tgAbfInfo(llLoop).lStartDate = llLogDate
                If tgAbfInfo(llLoop).sStatus = "S" Then
                    tgAbfInfo(llLoop).sStatus = "C"
                End If
            End If
            If llLogDate > tgAbfInfo(llLoop).lEndDate Then
                tgAbfInfo(llLoop).lEndDate = llLogDate
                If tgAbfInfo(llLoop).sStatus = "S" Then
                    tgAbfInfo(llLoop).sStatus = "C"
                End If
            End If
            Exit For
        End If
    Next llLoop
    If Not blFound Then
        tgAbfInfo(llUpper).sStatus = "N"
        tgAbfInfo(llUpper).lCode = 0
        tgAbfInfo(llUpper).iVefCode = ilVefCode
        tgAbfInfo(llUpper).lMondayDate = llMoDate
        tgAbfInfo(llUpper).lStartDate = llLogDate
        tgAbfInfo(llUpper).lEndDate = llLogDate
        ReDim Preserve tgAbfInfo(0 To llUpper + 1) As ABFINFO
    End If
    Exit Sub
mBuildAbfInfoErr:
    ilRet = 1
    Resume Next
End Sub

'
'       Log created for podcast client that needs to include the NTR with
'       the air time spots.
'
Public Function gGetNTRForLog(Form As Form, ilVpfIndex As Integer, slStartDate As String, slEndDate As String, blGenNTR As Boolean) As Integer
Dim ilRet As Integer
Dim ilWhichKey As Integer
Dim tlSBFTypes As SBFTypes
Dim ilPrevVefCode As Integer
Dim llLoopOnNTR As Long
Dim ilLoopOnVehicle As Integer
Dim ilVef As Integer
Dim ilfirstTime As Integer
Dim ilUpper As Integer
Dim ilNTRInx As Integer
Dim ilLoop As Integer
ReDim tlNTRInfo(0 To 0) As SBF

        gGetNTRForLog = True
        ReDim tgNTRInfo(0 To 0) As NTRINFO                      'array of NTR (sbf) records but with the keyfield in position 0 for sorting
        ReDim tgNtrSortInfo(0 To 0) As NTRSORTINFO              'array of vehicles with start/end indices pointing to TGNTRInfo array of SBF records
        
        'retain 1 entry in array to prevent subscript out of range in global arrays
        If tgVpf(ilVpfIndex).sGMedium <> "P" Or (Not blGenNTR) Then            'only podcast vehicle, read all NTRs for this period for the log; or its not a podcast log
            Exit Function
        End If
  
        ReDim tlNTRUnsortedInfo(0 To 0) As SBF                          'SBF (NTR) records without keyfield array
        
        hmSbf = CBtrvTable(ONEHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmSbf, "", sgDBPath & "Sbf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            gGetNTRForLog = False
            Exit Function
        End If
        imSbfRecLen = Len(tmSbf)
        
        hmOdf = CBtrvTable(TEMPHANDLE) 'CBtrvObj()
        ilRet = btrOpen(hmOdf, "", sgDBPath & "Odf.Btr", BTRV_OPEN_NORMAL, BTRV_OPEN_NONSHARE, BTRV_LOCK_NONE)
        If ilRet <> BTRV_ERR_NONE Then
            gGetNTRForLog = False
            Exit Function
        End If
        imOdfRecLen = Len(tmOdf)

        ilWhichKey = 2      'use key 2, by sbfdate, trantype (I)
        tlSBFTypes.iImport = False
        tlSBFTypes.iInstallment = False
        tlSBFTypes.iNTR = True
        
        'search sbf by date

        ilRet = gObtainSBF(Form, hmSbf, 0, slStartDate, slEndDate, tlSBFTypes, tlNTRUnsortedInfo(), ilWhichKey)
        If Not ilRet Then
            gGetNTRForLog = False
            Exit Function
        End If
        
        For llLoopOnNTR = LBound(tlNTRUnsortedInfo) To UBound(tlNTRUnsortedInfo) - 1
            tgNTRInfo(llLoopOnNTR).tNTR = tlNTRUnsortedInfo(llLoopOnNTR)
            tgNTRInfo(llLoopOnNTR).iBillVefCode = tlNTRUnsortedInfo(llLoopOnNTR).iBillVefCode
            ReDim Preserve tgNTRInfo(LBound(tgNTRInfo) To UBound(tgNTRInfo) + 1) As NTRINFO
        Next llLoopOnNTR
        
        If UBound(tgNTRInfo) - 1 > 1 Then
            ArraySortTyp fnAV(tgNTRInfo(), 0), UBound(tgNTRInfo), 0, LenB(tgNTRInfo(0)), 0, -1, 0
        End If
        
        ilUpper = 0
        ilfirstTime = True
        For ilNTRInx = LBound(tgNTRInfo) To UBound(tgNTRInfo) - 1
            If ilfirstTime Then
                ilfirstTime = False
                tgNtrSortInfo(ilUpper).iLoInx = ilNTRInx
                tgNtrSortInfo(ilUpper).iHiInx = ilNTRInx
                tgNtrSortInfo(ilUpper).iVefCode = tgNTRInfo(ilNTRInx).iBillVefCode
                ilPrevVefCode = tgNTRInfo(ilNTRInx).iBillVefCode
            End If
            'is current vehicle same as previous
            If tgNTRInfo(ilNTRInx).iBillVefCode = ilPrevVefCode Then
                tgNtrSortInfo(ilUpper).iHiInx = ilNTRInx
            Else
                ilUpper = ilUpper + 1
                ReDim Preserve tgNtrSortInfo(0 To ilUpper) As NTRSORTINFO
                tgNtrSortInfo(ilUpper).iLoInx = ilNTRInx
                tgNtrSortInfo(ilUpper).iHiInx = ilNTRInx
                tgNtrSortInfo(ilUpper).iVefCode = tgNTRInfo(ilNTRInx).iBillVefCode
                ilPrevVefCode = tgNTRInfo(ilNTRInx).iBillVefCode
            End If
        Next ilNTRInx
        ReDim Preserve tgNtrSortInfo(0 To UBound(tgNtrSortInfo) + 1) As NTRSORTINFO     '6-25-15 adjust for the last vehicle to process
        
        ilRet = btrClose(hmOdf)
        btrDestroy hmOdf
        ilRet = btrClose(hmSbf)
        btrDestroy hmSbf

        Erase tlNTRUnsortedInfo
        Exit Function

End Function
'
'               Create ODF records for the vehicle processing
'               tgNTRInfo contains the sorted array of NTR records.  First element is the bill vehicle that
'               was used to sort the array in vehicle order.
'               tgNTRSortInfo is array that contains the bill vehicle and the start/end index into tgNTRInfo
'
'               <input> ilLogVefcode - log vehicle code; otherwise conventional vehicle code
'                       ilVefCode - convention vehicle code when log vehicle is generated.  If not using log vehicle, 2 vehicle codes are same
'                       slZone - time zone generated
'
'  ODF Fields special for NTR record:
'       odfVefCode - vehicle code (conventional or the log vehicle)
'       odfalternatevefcode - vehicle code if using log vehicle (ability to sort major by thelog vehicle, then break each conventional vehicle within it)
'       odfairdate - sbfdate
'       odfzone - Zone probably isnt used, but will generate an NTR for each zone if it is used
'       odfAdfCode - contract advertiser
'       odfcntrno - contract #
'       odfProduct -
'       odfType = 5 is a new type for NTR designation
'       odfClfCode = air time spots, pointer to line ; if NTR, pointer to SBF
'
Public Sub gCreateODFForNTR(hlODF As Integer, hlChf As Integer, ilVefCode As Integer, tlzoneInfo() As ZONEINFO)
Dim ilLoopOnNTR As Integer
Dim ilLoopOnVehicle As Integer
Dim ilLoInx As Integer
Dim ilHiInx As Integer
Dim ilRet As Integer
Dim ilZone As Integer
Dim ilVefInx As Integer
Dim ilLogVefCode As Integer
Dim slZone As String

            For ilLoopOnVehicle = LBound(tgNtrSortInfo) To UBound(tgNtrSortInfo) - 1
                If ilVefCode >= tgNtrSortInfo(ilLoopOnVehicle).iVefCode Then
                    If tgNtrSortInfo(ilLoopOnVehicle).iVefCode = ilVefCode Then
                        ilLoInx = tgNtrSortInfo(ilLoopOnVehicle).iLoInx
                        ilHiInx = tgNtrSortInfo(ilLoopOnVehicle).iHiInx
                        tgNtrSortInfo(ilLoopOnVehicle).iVefCode = -tgNtrSortInfo(ilLoopOnVehicle).iVefCode    'negate to indicate already processed, so its not done for each day of the log date span
                        tmOdf.iUrfCode = tgUrf(0).iCode
                        ilVefInx = gBinarySearchVef(ilVefCode)            'determine if this belongs to a log vehicle
                        'If ilVefInx <= 0 Then
                        If ilVefInx < 0 Then
                            ilLogVefCode = ilVefCode
                        Else
                            ilLogVefCode = tgMVef(ilVefInx).iVefCode     'this vehicle belongs to log vehicle
                            If ilLogVefCode = 0 Then
                                ilLogVefCode = ilVefCode
                            End If
                        End If
                        tmOdf.iAlternateVefCode = ilVefCode
                        tmOdf.iVefCode = ilLogVefCode
                        tmOdf.iGenDate(0) = igGenDate(0)
                        tmOdf.iGenDate(1) = igGenDate(1)
                        tmOdf.lGenTime = lgGenTime
                        tmOdf.iType = 5                 'new type (NTR)
                        tmOdf.iAirTime(0) = 0
                        tmOdf.iAirTime(1) = 0
                        tmOdf.iLocalTime(0) = 0
                        tmOdf.iLocalTime(1) = 0
                        tmOdf.iFeedTime(0) = 0
                        tmOdf.iFeedTime(1) = 0
                        tmOdf.iEtfCode = 0
                        tmOdf.iEnfCode = 0
                        tmOdf.sProgCode = 0
                        tmOdf.iMnfFeed = 0
                        tmOdf.iDPSort = 0
                        tmOdf.iWkNo = 0
                        tmOdf.ianfCode = 0
                        tmOdf.iUnits = 1
                        tmOdf.iLen(0) = 0
                        tmOdf.iLen(1) = 0
                        tmOdf.iAdfCode = 0
                        tmOdf.lCifCode = 0
                        tmOdf.sProduct = ""
                        tmOdf.iMnfSubFeed = 0
                        tmOdf.lCntrNo = 0
                        tmOdf.lchfcxfCode = 0
                        tmOdf.iRdfSortCode = 0
                        tmOdf.sDPDesc = ""
                        tmOdf.iBreakNo = 0
                        tmOdf.iPositionNo = 0
                        tmOdf.lCefCode = 0
                        tmOdf.lEvtIDCefCode = 0
                        tmOdf.sDupeAvailID = 0
                        tmOdf.sShortTitle = ""
                        tmOdf.imnfSeg = 0
                        tmOdf.sPageEjectFlag = "N"
                        
                        tmOdf.iSeqNo = 0
                        tmOdf.iDaySort = 0
                        tmOdf.lEvtCefCode = 0
                        tmOdf.iEvtCefSort = 0
                        tmOdf.sLogType = ""
                        
                        For ilLoopOnNTR = ilLoInx To ilHiInx
                            tmSbf = tgNTRInfo(ilLoopOnNTR).tNTR
                            tmChfSrchKey.lCode = tmSbf.lChfCode
                            ilRet = btrGetEqual(hlChf, tmChf, imCHFRecLen, tmChfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
                            gLogBtrError ilRet, "gCreateODFForNTR: btrGetEqualChf"
                            '2-4-15 ignore proposals on the log for NTRs
                            If tmChf.sDelete = "N" And (tmChf.sStatus = "H" Or tmChf.sStatus = "O" Or tmChf.sStatus = "G" Or tmChf.sStatus = "N") Then
                                tmOdf.iAdfCode = tmChf.iAdfCode
                                tmOdf.lCntrNo = tmChf.lCntrNo
                                If Trim$(tmChf.sProduct) <> "" Then
                                    tmOdf.sProduct = tmChf.sProduct
                                End If
                                tmOdf.iAirDate(0) = tmSbf.iDate(0)
                                tmOdf.iAirDate(1) = tmSbf.iDate(1)
                                'tmOdf.lCefCode = tmSbf.lCode
                                tmOdf.lClfCode = tmSbf.lCode
                                
                                slZone = ""
                                For ilZone = 0 To UBound(tlzoneInfo) - 1 Step 1
                                    Select Case tlzoneInfo(ilZone).sZone
                                        Case "E"
                                            slZone = "EST"
                                        Case "M"
                                            slZone = "MST"
                                        Case "C"
                                            slZone = "CST"
                                        Case "P"
                                            slZone = "PST"
            
                                    End Select
                                    tmOdf.sZone = slZone
                                    tmOdf.lCode = 0
                                    ilRet = btrInsert(hlODF, tmOdf, Len(tmOdf), INDEXKEY0)
                                    gLogBtrError ilRet, "gCreateODFForNTR: btrInsert ODF"
                                Next ilZone
                            End If
                        Next ilLoopOnNTR
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next ilLoopOnVehicle
        Exit Sub
    End Sub
'*************************************************************
'*                                                           *
'*      Procedure Name:gObtainSBF                            *
'*      Extended read to get all matching SBF                *
'*      records with passed transaction date                 *
'*      <input>  hlSbf - SBF handle (file must open          *
'*               llChfCode - 0 if all contracts, else        *
'*                           selective contract code         *
'*               slEarliestDate - start date to gather       *
'*               slLatestDate - end date to gather           *
'*               tlSBFTypes - Types of SBF records to        *
'*                        include (NTR, Installment,         *
'*                        Import types                       *
'*               ilWhichKey - # indicating key to use        *
'*               ilVefCode - if key 4, vehicle code to match *
'*     I/O       tlSbf() - array of matching SBF recds       *
'*                                                           *
'*             Created:8-30-02       By:D. Hosaka            *
'*            Modified:              By:                     *
'*                                                           *
'*            Comments: Read all of sbf by transaction       *
'*                      date                                 *
'*      9-30-02 Modify for sbftypes and selective or all     *
'               contracts                                    *
'       11-28-06 send flag for which key to use
'       1-28-14 this type statment moved from rptrec due to Logs
'            requireing the NTR records and other projects
'            have Log.bas in their project leaving unreferenced
'            items
''*************************************************************
Function gObtainSBF(RptForm As Form, hlSbf As Integer, llChfCode As Long, slEarliestDate As String, slLatestDate As String, tlSBFTypes As SBFTypes, tlSbf() As SBF, ilWhichKey As Integer, Optional ilVefCode = 0) As Integer
'
'    gObtainSBF (hlSBF, llChfCode, slStartDate, slEndDate, tlSBF(), ilWhichKey)
'
    Dim ilRet As Integer    'Return status
    Dim slStr As String
    Dim ilExtLen As Integer
    Dim llNoRec As Long
    Dim llRecPos As Long
    Dim ilOffSet As Integer
    'TTP 10855 - prevent overflow due to too many NTR items
    'Dim ilSBFUpper As Integer
    Dim llSBFUpper As Long
    Dim ilEarliestDate(0 To 1) As Integer
    Dim ilLatestDate(0 To 1) As Integer
    Dim tlDateTypeBuff As POPDATETYPE   'Type field record
    Dim tlLongTypeBuff As POPLCODE
    Dim tlStrTypeBuff As POPCHARTYPE   'Type field record
    Dim tlIntTypeBuff As POPICODE
    Dim ilRetGetFirst As Integer


    gPackDate slEarliestDate, ilEarliestDate(0), ilEarliestDate(1)
    gPackDate slLatestDate, ilLatestDate(0), ilLatestDate(1)

    ReDim tlSbf(0 To 0) As SBF
    btrExtClear hlSbf   'Clear any previous extend operation
    ilExtLen = Len(tlSbf(0))  'Extract operation record size
    imSbfRecLen = Len(tlSbf(0))
    llSBFUpper = UBound(tlSbf)

    If ilWhichKey = 0 Then          'use key0
        ilRetGetFirst = btrGetFirst(hlSbf, tmSbf, imSbfRecLen, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation

        If ilRetGetFirst <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
            Call btrExtSetBounds(hlSbf, llNoRec, -1, "UC", "SBF", "") '"EG") 'Set extract limits (all records)

            tlDateTypeBuff.iDate0 = ilEarliestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilEarliestDate(1)
            ilOffSet = gFieldOffset("Sbf", "SbfDate")
           ' ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)        '11-28-06 should be date test, not int.
            On Error GoTo mObtainSBFErr
            gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
            On Error GoTo 0


            If llChfCode = 0 Then                   'retrieve all contracts
                tlDateTypeBuff.iDate0 = ilLatestDate(0)                       'retrieve all trans between given dates
                tlDateTypeBuff.iDate1 = ilLatestDate(1)
                ilOffSet = gFieldOffset("Sbf", "SbfDate")
                'ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_INT, ilOffset, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)
                ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_LAST_TERM, tlDateTypeBuff, 4)      'should be date test, not int.
                On Error GoTo mObtainSBFErr
                gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
                On Error GoTo 0
            Else                'find matching contract
                tlDateTypeBuff.iDate0 = ilLatestDate(0)                       'retrieve all trans equal or prior to this date for pacing
                tlDateTypeBuff.iDate1 = ilLatestDate(1)
                ilOffSet = gFieldOffset("Sbf", "SbfDate")
                ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
                On Error GoTo mObtainSBFErr
                gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
                On Error GoTo 0

                tlLongTypeBuff.lCode = llChfCode                       'retrieve matching contract            tlDateTypeBuff.iDate1 = ilLatestDate(1)
                ilOffSet = gFieldOffset("Sbf", "SbfchfCode")
                ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_INT, ilOffSet, 4, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlLongTypeBuff, 4)
                On Error GoTo mObtainSBFErr
                gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
                On Error GoTo 0
            End If
        End If
    ElseIf ilWhichKey = 2 Then         'key2 : trantype, bill date
        ilRetGetFirst = btrGetFirst(hlSbf, tmSbf, imSbfRecLen, INDEXKEY2, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRetGetFirst <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
            Call btrExtSetBounds(hlSbf, llNoRec, -1, "UC", "SBF", "") '"EG") 'Set extract limits (all records)

            tlDateTypeBuff.iDate0 = ilEarliestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilEarliestDate(1)
            ilOffSet = gFieldOffset("Sbf", "SbfDate")
            ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            On Error GoTo mObtainSBFErr
            gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
            On Error GoTo 0

            tlDateTypeBuff.iDate0 = ilLatestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilLatestDate(1)
            ilOffSet = gFieldOffset("Sbf", "SbfDate")
            ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            On Error GoTo mObtainSBFErr
            gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
            On Error GoTo 0

            tlStrTypeBuff.sType = "I"               'NTRs
            ilOffSet = gFieldOffset("Sbf", "SbfTranType")
            ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlStrTypeBuff, 1)
            On Error GoTo mObtainSBFErr
            gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
            On Error GoTo 0
        End If
    ElseIf ilWhichKey = 3 Then                'use key 3:  tran type, then post date
        ilRetGetFirst = btrGetFirst(hlSbf, tmSbf, imSbfRecLen, INDEXKEY3, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        If ilRetGetFirst <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
            Call btrExtSetBounds(hlSbf, llNoRec, -1, "UC", "SBF", "") '"EG") 'Set extract limits (all records)

            tlDateTypeBuff.iDate0 = ilEarliestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilEarliestDate(1)
            ilOffSet = gFieldOffset("Sbf", "SbfPostDate")
            ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            On Error GoTo mObtainSBFErr
            gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
            On Error GoTo 0

            tlDateTypeBuff.iDate0 = ilLatestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilLatestDate(1)
            ilOffSet = gFieldOffset("Sbf", "SbfPostDate")
            ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            On Error GoTo mObtainSBFErr
            gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
            On Error GoTo 0

            tlStrTypeBuff.sType = "T"    'Extract all matching records for rep only
            ilOffSet = gFieldOffset("Sbf", "SbfTranType")
            ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlStrTypeBuff, 1)
            On Error GoTo mObtainSBFErr
            gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
            On Error GoTo 0
        End If
    Else                    'key 4: vehicle, trantype & postdate
        tmSbfSrchKey4.iBillVefCode = ilVefCode
        tmSbfSrchKey4.sTranType = "T"
        tmSbfSrchKey4.iPostDate(0) = ilEarliestDate(0)
        tmSbfSrchKey4.iPostDate(1) = ilEarliestDate(1)
        'TTP 10190 - Barter Payments report error - typo on variable used when obtaining SBF records using key 4: vehicle, trantype & postdate
        'ilRet = btrGetGreaterOrEqual(hlSbf, tmSbf, imSbfRecLen, tmSbfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)  'Get first record as starting point of extend operation
        ilRetGetFirst = btrGetGreaterOrEqual(hlSbf, tmSbf, imSbfRecLen, tmSbfSrchKey4, INDEXKEY4, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        If ilRetGetFirst <> BTRV_ERR_END_OF_FILE Then
            llNoRec = gExtNoRec(ilExtLen)               'Obtain number of records
            Call btrExtSetBounds(hlSbf, llNoRec, -1, "UC", "SBF", "") '"EG") 'Set extract limits (all records)
            
            If ilVefCode > 0 Then           'get selective or all if 0
                tlIntTypeBuff.iCode = ilVefCode
                ilOffSet = gFieldOffset("Sbf", "SbfBillVefCode")
                ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_INT, ilOffSet, 2, BTRV_EXT_EQUAL, BTRV_EXT_AND, tlIntTypeBuff, 2)
                On Error GoTo mObtainSBFErr
                gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
                On Error GoTo 0
            End If
           
            
            tlDateTypeBuff.iDate0 = ilEarliestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilEarliestDate(1)
            ilOffSet = gFieldOffset("Sbf", "SbfPostDate")
            ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_GTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            On Error GoTo mObtainSBFErr
            gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
            On Error GoTo 0

            tlDateTypeBuff.iDate0 = ilLatestDate(0)                       'retrieve all trans equal or prior to this date for pacing
            tlDateTypeBuff.iDate1 = ilLatestDate(1)
            ilOffSet = gFieldOffset("Sbf", "SbfPostDate")
            ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_DATE, ilOffSet, 4, BTRV_EXT_LTE, BTRV_EXT_AND, tlDateTypeBuff, 4)
            On Error GoTo mObtainSBFErr
            gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
            On Error GoTo 0

            tlStrTypeBuff.sType = "T"    'Extract all matching records for rep only
            ilOffSet = gFieldOffset("Sbf", "SbfTranType")
            ilRet = btrExtAddLogicConst(hlSbf, BTRV_KT_STRING, ilOffSet, 1, BTRV_EXT_EQUAL, BTRV_EXT_LAST_TERM, tlStrTypeBuff, 1)
            On Error GoTo mObtainSBFErr
            gBtrvErrorMsg ilRet, "gObtainSbf (btrExtAddLogicConst):" & "Sbf.Btr", RptForm
            On Error GoTo 0
        End If
    End If

    If ilRetGetFirst <> BTRV_ERR_END_OF_FILE Then
        ilRet = btrExtAddField(hlSbf, 0, ilExtLen) 'Extract the whole record
        On Error GoTo mObtainSBFErr
        gBtrvErrorMsg ilRet, "gObtainSBF (btrExtAddField):" & "SBF.Btr", RptForm
        On Error GoTo 0
        ilRet = btrExtGetNext(hlSbf, tmSbf, ilExtLen, llRecPos)
        If (ilRet <> BTRV_ERR_END_OF_FILE) And (ilRet <> BTRV_ERR_FILTER_LIMIT) Then
            On Error GoTo mObtainSBFErr
            gBtrvErrorMsg ilRet, "gObtainSBF (btrExtGetNextExt):" & "SBF.Btr", RptForm
            On Error GoTo 0
            ilExtLen = Len(tmSbf)  'Extract operation record size
            Do While ilRet = BTRV_ERR_REJECT_COUNT
                ilRet = btrExtGetNext(hlSbf, tmSbf, ilExtLen, llRecPos)
            Loop
            Do While ilRet = BTRV_ERR_NONE
                'Filter the proper SBF Types and if a single contract vs. all contracts for the date span
                If ((tmSbf.sTranType = "I" And tlSBFTypes.iNTR) Or (tmSbf.sTranType = "F" And tlSBFTypes.iInstallment) Or (tmSbf.sTranType = "T" And tlSBFTypes.iImport)) And ((llChfCode = 0) Or (llChfCode > 0 And tmSbf.lChfCode = llChfCode)) Then
                    slStr = ""
                    tlSbf(UBound(tlSbf)) = tmSbf           'save entire record
                    ReDim Preserve tlSbf(0 To UBound(tlSbf) + 1) As SBF
                End If
                ilRet = btrExtGetNext(hlSbf, tmSbf, ilExtLen, llRecPos)
                Do While ilRet = BTRV_ERR_REJECT_COUNT
                    ilRet = btrExtGetNext(hlSbf, tmSbf, ilExtLen, llRecPos)
                Loop
            Loop
        End If
    End If
    gObtainSBF = True
    Exit Function
mObtainSBFErr:
    On Error GoTo 0
    gObtainSBF = False
    Exit Function
End Function

Private Sub mGetLineTimes(ilLnStartTime() As Integer, ilLnEndTime() As Integer)
    Dim ilRdf As Integer
    Dim llTBStartTime As Long
    Dim llTBEndTime As Long
    Dim llLnStartTime As Long
    Dim llLnEndTime As Long
    Dim blFirstRdfTime As Boolean
    Dim ilLoop As Integer
    
    If (tmClf.iStartTime(0) = 1) And (tmClf.iStartTime(1) = 0) Then
        blFirstRdfTime = True
        ilRdf = gBinarySearchRdf(tmClf.iRdfCode)
        If ilRdf <> -1 Then
            llLnStartTime = 0
            llLnEndTime = 86400
            tmRdf = tgMRdf(ilRdf)
            For ilLoop = LBound(tmRdf.iStartTime, 2) To UBound(tmRdf.iStartTime, 2) Step 1
                If (tmRdf.iStartTime(0, ilLoop) <> 1) Or (tmRdf.iStartTime(1, ilLoop) <> 0) Then
                    gUnpackTimeLong tmRdf.iStartTime(0, ilLoop), tmRdf.iStartTime(1, ilLoop), False, llTBStartTime
                    gUnpackTimeLong tmRdf.iEndTime(0, ilLoop), tmRdf.iEndTime(1, ilLoop), True, llTBEndTime
                    If blFirstRdfTime Then
                        llLnStartTime = llTBStartTime
                        llLnEndTime = llTBEndTime
                        blFirstRdfTime = False
                    Else
                        'Test if adjacent.
                        'If not then return the first one found
                        'Expand time if adjacent
                        If (llTBStartTime = llLnEndTime) Or (llTBStartTime + 86400 = llLnEndTime) Then
                            llLnEndTime = llTBEndTime
                        ElseIf (llLnStartTime = llTBEndTime) Or (llLnStartTime + 86400 = llTBEndTime) Then
                            llLnStartTime = llTBStartTime
                        End If
                    End If
                End If
            Next ilLoop
            gPackTimeLong llLnStartTime, ilLnStartTime(0), ilLnStartTime(1)
            gPackTimeLong llLnEndTime, ilLnEndTime(0), ilLnEndTime(1)
        End If
    Else
        ilLnStartTime(0) = tmClf.iStartTime(0)
        ilLnStartTime(1) = tmClf.iStartTime(1)
        ilLnEndTime(0) = tmClf.iEndTime(0)
        ilLnEndTime(1) = tmClf.iEndTime(1)
    End If

End Sub


Private Sub mGetCpf(hlCif As Integer, llCifCode As Long)
    Dim ilRet As Integer
    If llCifCode <= 0 Then
        tmCif.lCode = 0
        tmCif.lcpfCode = 0
        tmCif.iMcfCode = 0
        tmCif.sReel = ""
        tmCpf.sName = ""
        tmCpf.sISCI = ""
        tmCpf.sCreative = ""
        Exit Sub
    End If
    If tmCif.lCode <> llCifCode Then
        imCifRecLen = Len(tmCif)
        tmCifSrchKey.lCode = llCifCode
        ilRet = btrGetEqual(hlCif, tmCif, imCifRecLen, tmCifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        If (ilRet = BTRV_ERR_NONE) Then
            tmCpfSrchKey.lCode = tmCif.lcpfCode
            ilRet = btrGetEqual(hmCpf, tmCpf, imCpfRecLen, tmCpfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
            If (ilRet <> BTRV_ERR_NONE) Then
                tmCpf.sName = ""
                tmCpf.sISCI = ""
                tmCpf.sCreative = ""
            End If
        Else
            tmCif.lcpfCode = 0
            tmCif.iMcfCode = 0
            tmCif.sReel = ""
            tmCpf.sName = ""
            tmCpf.sISCI = ""
            tmCpf.sCreative = ""
        End If
    End If
End Sub

Private Sub mCloseFiles(hlCnf As Integer, hlCef As Integer, hlCTSsf As Integer, hlSsf As Integer, hlSdf As Integer, hlChf As Integer, hlODF As Integer, hlVef As Integer, hlVlf As Integer, hlDlf As Integer, hlTzf As Integer, hlCrf As Integer, hlSif As Integer, hlVsf As Integer, ilVehicle() As Integer, tlLLC() As LLC)
    Dim ilRet As Integer
    
    On Error Resume Next
    Erase tmLstCode
    If imLstExist Then
        'ilRet = btrClose(hmLst)
        btrDestroy hmMnf
        'btrDestroy hmCff
        'btrDestroy hmSmf
        btrDestroy hmDrf
        btrDestroy hmDpf
        btrDestroy hmDef
        btrDestroy hmRaf
        'btrDestroy hmLst
    End If
    ilRet = btrClose(hmCpf)
    btrDestroy hmCpf
    ilRet = btrClose(hmCvf)
    ilRet = btrClose(hmRsf)
    ilRet = btrClose(hlCnf)
    ilRet = btrClose(hmCff)
    ilRet = btrClose(hmSmf)
    ilRet = btrClose(hmRdf)
    ilRet = btrClose(hlCef)
    ilRet = btrClose(hlCTSsf)
    ilRet = btrClose(hlSsf)
    ilRet = btrClose(hlSdf)
    ilRet = btrClose(hlChf)
    ilRet = btrClose(hlODF)
    ilRet = btrClose(hlVef)
    ilRet = btrClose(hlVlf)
    ilRet = btrClose(hlDlf)
    ilRet = btrClose(hmCif)
    ilRet = btrClose(hlTzf)
    ilRet = btrClose(hlCrf)
    ilRet = btrClose(hlSif)
    ilRet = btrClose(hmAdf)
    ilRet = btrClose(hmClf)
    ilRet = btrClose(hlVsf)
    ilRet = btrClose(hmCxf)

    btrDestroy hmCff
    btrDestroy hmSmf
    btrDestroy hmRdf
    btrDestroy hlCTSsf
    btrDestroy hlSsf
    btrDestroy hlSdf
    btrDestroy hlChf
    btrDestroy hlODF
    btrDestroy hlVef
    btrDestroy hlVlf
    btrDestroy hlDlf
    btrDestroy hmCif
    btrDestroy hlTzf
    btrDestroy hlCrf
    btrDestroy hlSif
    btrDestroy hmAdf
    btrDestroy hmClf
    btrDestroy hlVsf
    btrDestroy hlCef
    btrDestroy hlCnf
    btrDestroy hmRsf
    btrDestroy hmCvf
    btrDestroy hmCxf

    If tgSpf.sSystemType = "R" Then
        ilRet = btrClose(hmFsf)
        ilRet = btrClose(hmPrf)
        ilRet = btrClose(hmFnf)
        btrDestroy hmFsf
        btrDestroy hmPrf
        btrDestroy hmFnf
    End If

    Erase ilVehicle
    Erase tmPageEject
    Erase tlLLC
End Sub

Private Sub mProcTestForDuplBB(ilAddSpot As Integer)
    Dim ilBB As Integer
    
    ilAddSpot = True
    If tgSpf.sUsingBBs <> "Y" Then
        'Return
        Exit Sub
    End If
    If (tmSdf.sSpotType <> "O") Or (tmSdf.sSpotType <> "C") Then
        'Return
        Exit Sub
    End If
    For ilBB = 0 To UBound(tmBBSdfInfo) - 1 Step 1
        If (tmBBSdfInfo(ilBB).lChfCode = tmSdf.lChfCode) And (tmBBSdfInfo(ilBB).iLen = tmSdf.iLen) And (tmBBSdfInfo(ilBB).sType = tmSdf.sSpotType) Then
            If (tmBBSdfInfo(ilBB).iTime(0) = tmSdf.iTime(0)) And (tmBBSdfInfo(ilBB).iTime(1) = tmSdf.iTime(1)) Then
                ilAddSpot = False
                'Return
                Exit Sub
            End If
        End If
    Next ilBB
    tmBBSdfInfo(UBound(tmBBSdfInfo)).sType = tmSdf.sSpotType
    tmBBSdfInfo(UBound(tmBBSdfInfo)).lChfCode = tmSdf.lChfCode
    tmBBSdfInfo(UBound(tmBBSdfInfo)).iLen = tmSdf.iLen
    tmBBSdfInfo(UBound(tmBBSdfInfo)).iTime(0) = tmSdf.iTime(0)
    tmBBSdfInfo(UBound(tmBBSdfInfo)).iTime(1) = tmSdf.iTime(1)
    tmOdf.iBreakNo = 0  'tmAvail.iAvInfo
    tmOdf.iPositionNo = 0
    For ilBB = 0 To UBound(tmBBSdfInfo) - 1 Step 1
        If (tmBBSdfInfo(ilBB).iTime(0) = tmSdf.iTime(0)) And (tmBBSdfInfo(ilBB).iTime(1) = tmSdf.iTime(1)) Then
            tmOdf.iPositionNo = tmOdf.iPositionNo + 1
        End If
    Next ilBB
    ReDim Preserve tmBBSdfInfo(0 To UBound(tmBBSdfInfo) + 1) As BBSDFINFO
End Sub

Private Sub mProcAdjDate(slAdjLocalOrFeed As String, llDate As Long, llSDate As Long, llEDate As Long, slInLogType As String)
    Dim ilAirHour As Integer, ilLocalHour As Integer, ilFeedHour As Integer
    Dim llLocalTime As Long
    Dim llFeedTime As Long
    Dim llWkDateSet As Long
    
    imDateAdj = False
    'Test if Air time is AM and Local Time is PM. If so, adjust date
    ilAirHour = tmOdf.iAirTime(1) \ 256  'Obtain month
    ilLocalHour = tmOdf.iLocalTime(1) \ 256  'Obtain month
    ilFeedHour = tmOdf.iFeedTime(1) \ 256
    If (tgSpf.sGUseAffSys = "Y") Then
        If slAdjLocalOrFeed <> "F" Then
            If (ilAirHour < 6) And (ilLocalHour > 17) Then
            
                '7/11/14: Lock avail between 12am-3am
                If (slInLogType = "F") Or (slInLogType = "R") Or (slInLogType = "A") Then
                    If llDate = llEDate + 1 Then
                        gUnpackDate tmSdf.iDate(0), tmSdf.iDate(1), smLockDate
                        gUnpackTime tmSdf.iTime(0), tmSdf.iTime(1), "A", "1", smLockStartTime
                        smLockEndTime = gFormatTimeLong(gTimeToLong(smLockStartTime, False) + 1, "A", "1")
                        If (imLockVefCode <> tmSdf.iVefCode) Or (gTimeToLong(smLockStartTime, False) <> lmLockStartTime) Then
                            gSetLockStatus tmSdf.iVefCode, 1, -1, smLockDate, smLockDate, tmSdf.iGameNo, smLockStartTime, smLockEndTime
                            imLockVefCode = tmSdf.iVefCode
                            lmLockStartTime = gTimeToLong(smLockStartTime, False)
                        End If
                    End If
                End If
                
                imDateAdj = True
                gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), smAdjDate
                smAdjDate = gDecOneDay(smAdjDate)
                gPackDate smAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
                If (llDate > llSDate) Then
                    imGenODF = True
                Else
                    imGenODF = False
                End If
            ElseIf (ilLocalHour < 6) And (ilAirHour > 17) Then
                imDateAdj = True
                gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), smAdjDate
                smAdjDate = gIncOneDay(smAdjDate)
                gPackDate smAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
                If (llDate >= llSDate) And (llDate <= llEDate) Then
                    imGenODF = True
                Else
                    imGenODF = False
                End If
            Else
                gUnpackTimeLong tmOdf.iLocalTime(0), tmOdf.iLocalTime(1), False, llLocalTime
                '8/13/14: Generate spots for last date+1 into lst with Fed as *
                'If (llLocalTime >= llLPTime) And (llLocalTime < llLNTime) Then
                '    imGenODF = True
                'Else
                '    imGenODF = False
                'End If
                If Not bmCreatingLstDate1 Then
                    If (llLocalTime >= lmLPTime) And (llLocalTime < lmLNTime) Then
                        imGenODF = True
                    Else
                        imGenODF = False
                    End If
                Else
                    If (llLocalTime >= 0) And (llLocalTime < lmGLocalAdj) Then
                        imGenODF = True
                    Else
                        imGenODF = False
                    End If
                End If
            End If
        Else
            If (ilAirHour < 6) And (ilFeedHour > 17) Then
                imDateAdj = True
                gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), smAdjDate
                smAdjDate = gDecOneDay(smAdjDate)
                gPackDate smAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
                If (llDate > llSDate) Then
                    imGenODF = True
                Else
                    imGenODF = False
                End If
            ElseIf (ilFeedHour < 6) And (ilAirHour > 17) Then
                imDateAdj = True
                gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), smAdjDate
                smAdjDate = gIncOneDay(smAdjDate)
                gPackDate smAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
                If (llDate >= llSDate) And (llDate <= llEDate) Then
                    imGenODF = True
                Else
                    imGenODF = False
                End If
            Else
                gUnpackTimeLong tmOdf.iFeedTime(0), tmOdf.iFeedTime(1), False, llFeedTime
                If (llFeedTime >= lmLPTime) And (llFeedTime < lmLNTime) Then
                    imGenODF = True
                Else
                    imGenODF = False
                End If
            End If
        End If
    Else
        imGenODF = True
        If slAdjLocalOrFeed <> "F" Then
            If (ilAirHour < 6) And (ilLocalHour > 17) Then
                'If monday convert to next sunday- this is wrong but the same spot
                'runs each sunday (the spot should have show on the previous week sunday)
                'If not monday, then subtract one day
                imDateAdj = True
                gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), smAdjDate
                If gWeekDayStr(smAdjDate) = 0 Then
                    smAdjDate = gObtainNextSunday(smAdjDate)
                Else
                    smAdjDate = gDecOneDay(smAdjDate)
                End If
                gPackDate smAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
            End If
        Else
            If (ilAirHour < 6) And (ilFeedHour > 17) Then
                'If monday convert to next sunday- this is wrong but the same spot
                'runs each sunday (the spot should have show on the previous week sunday)
                'If not monday, then subtract one day
                imDateAdj = True
                gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), smAdjDate
                If gWeekDayStr(smAdjDate) = 0 Then
                    smAdjDate = gObtainNextSunday(smAdjDate)
                Else
                    smAdjDate = gDecOneDay(smAdjDate)
                End If
                gPackDate smAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
            End If
        End If
    End If
    gUnpackDateLong tmOdf.iAirDate(0), tmOdf.iAirDate(1), llWkDateSet
    tmOdf.iWkNo = (llWkDateSet - lm010570) \ 7 + 1
    If (imDateAdj = False) And (smXMid = "Y") Then
        gUnpackDate tmOdf.iAirDate(0), tmOdf.iAirDate(1), smAdjDate
        smAdjDate = gIncOneDay(smAdjDate)
        gPackDate smAdjDate, tmOdf.iAirDate(0), tmOdf.iAirDate(1)
    End If
End Sub

Private Sub mProcSeqNo(ilSeqNo As Integer, slFor As String, hlODF As Integer, tlOdf As ODF, ilODFVefCode As Integer, ilSimVefCode As Integer, ilVehCode As Integer)
    Dim ilRet As Integer
    Dim ilTestVefCode As Integer
    Dim llOdf As Long
    
    ilSeqNo = 0
    If slFor = "D" Then
        'tmOdfSrchKey1.iUrfCode = tgUrf(0).iCode
        tmOdfSrchKey1.iMnfFeed = tmOdf.iMnfFeed 'tmDlf.iMnfFeed
        tmOdfSrchKey1.iAirDate(0) = tmOdf.iAirDate(0)   'ilLogDate0
        tmOdfSrchKey1.iAirDate(1) = tmOdf.iAirDate(1)   'ilLogDate1
        tmOdfSrchKey1.iFeedTime(0) = tmOdf.iFeedTime(0) 'tmDlf.iFeedTime(0)
        tmOdfSrchKey1.iFeedTime(1) = tmOdf.iFeedTime(1) 'tmDlf.iFeedTime(1)
        tmOdfSrchKey1.sZone = tmOdf.sZone   'tmDlf.sZone
        tmOdfSrchKey1.iSeqNo = 32000
        ilRet = btrGetLessOrEqual(hlODF, tlOdf, imOdfRecLen, tmOdfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
        'Do While (ilRet = BTRV_ERR_NONE) And (tlOdf.iUrfCode = tgUrf(0).iCode) And (tlOdf.iMnfFeed = tmOdf.iMnfFeed) And (tlOdf.iAirDate(0) = tmOdf.iAirDate(0)) And (tlOdf.iAirDate(1) = tmOdf.iAirDate(1))
        Do While (ilRet = BTRV_ERR_NONE) And (tlOdf.iMnfFeed = tmOdf.iMnfFeed) And (tlOdf.iAirDate(0) = tmOdf.iAirDate(0)) And (tlOdf.iAirDate(1) = tmOdf.iAirDate(1))
            If (tlOdf.iLocalTime(0) <> tmOdf.iLocalTime(0)) Or (tlOdf.iLocalTime(1) <> tmOdf.iLocalTime(1)) Or (tlOdf.sZone <> tmOdf.sZone) Then
                Exit Do
            End If
            If ilODFVefCode <= 0 Then
                If ilSimVefCode <= 0 Then
                    'This test is OK on Log vehicle since ODF is created for LOG vehicle
                    If tmVef.iVefCode > 0 Then  'Log vehicle defined
                        If tlOdf.iVefCode = tmVef.iVefCode Then
                            ilSeqNo = tlOdf.iSeqNo
                            Exit Do
                        End If
                    Else
                        If tlOdf.iVefCode = ilVehCode Then
                            ilSeqNo = tlOdf.iSeqNo
                            Exit Do
                        End If
                    End If
                Else
                    If tlOdf.iVefCode = ilSimVefCode Then
                        ilSeqNo = tlOdf.iSeqNo
                        Exit Do
                    End If
                End If
            Else
                If tlOdf.iVefCode = ilODFVefCode Then
                    ilSeqNo = tlOdf.iSeqNo
                    Exit Do
                End If
            End If
            ilRet = btrGetPrevious(hlODF, tlOdf, imOdfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Else
        'tmOdfSrchKey0.iUrfCode = tgUrf(0).iCode
        If ilODFVefCode <= 0 Then
            If ilSimVefCode <= 0 Then
                If tmVef.iVefCode > 0 Then  'Log vehicle defined
                    tmOdfSrchKey0.iVefCode = tmVef.iVefCode
                    ilTestVefCode = tmVef.iVefCode
                Else
                    tmOdfSrchKey0.iVefCode = ilVehCode
                    ilTestVefCode = ilVehCode
                End If
            Else
                tmOdfSrchKey0.iVefCode = ilSimVefCode
                ilTestVefCode = ilSimVefCode
            End If
        Else
            tmOdfSrchKey0.iVefCode = ilODFVefCode
            ilTestVefCode = ilODFVefCode
        End If
        '8/13/14: Generate spots for last date+1 into lst with Fed as *
        If Not bmCreatingLstDate1 Then
            tmOdfSrchKey0.iAirDate(0) = tmOdf.iAirDate(0)   'ilLogDate0
            tmOdfSrchKey0.iAirDate(1) = tmOdf.iAirDate(1)   'ilLogDate1
            tmOdfSrchKey0.iLocalTime(0) = tmOdf.iLocalTime(0)   'tmDlf.iLocalTime(0)
            tmOdfSrchKey0.iLocalTime(1) = tmOdf.iLocalTime(1)   'tmDlf.iLocalTime(1)
            tmOdfSrchKey0.sZone = tmOdf.sZone   'tmDlf.sZone
            tmOdfSrchKey0.iSeqNo = 32000
            ilRet = btrGetLessOrEqual(hlODF, tlOdf, imOdfRecLen, tmOdfSrchKey0, INDEXKEY0, BTRV_LOCK_NONE)   'Get first record as starting point of extend operation
            'If (ilRet = BTRV_ERR_NONE) And (tlOdf.iUrfCode = tgUrf(0).iCode) And (tlOdf.iVefCode = ilTestVefCode) And (tlOdf.iAirDate(0) = tmOdf.iAirDate(0)) And (tlOdf.iAirDate(1) = tmOdf.iAirDate(1)) Then
            If (ilRet = BTRV_ERR_NONE) And (tlOdf.iVefCode = ilTestVefCode) And (tlOdf.iAirDate(0) = tmOdf.iAirDate(0)) And (tlOdf.iAirDate(1) = tmOdf.iAirDate(1)) Then
                If (tlOdf.iLocalTime(0) = tmOdf.iLocalTime(0)) And (tlOdf.iLocalTime(1) = tmOdf.iLocalTime(1)) And (tlOdf.sZone = tmOdf.sZone) Then
                    ilSeqNo = tlOdf.iSeqNo
                End If
            End If
        Else
            For llOdf = 0 To UBound(tmDate1ODF) - 1 Step 1
                tlOdf = tmDate1ODF(llOdf)
                If (tlOdf.iVefCode = ilTestVefCode) And (tlOdf.iAirDate(0) = tmOdf.iAirDate(0)) And (tlOdf.iAirDate(1) = tmOdf.iAirDate(1)) Then
                    If (tlOdf.iLocalTime(0) = tmOdf.iLocalTime(0)) And (tlOdf.iLocalTime(1) = tmOdf.iLocalTime(1)) And (tlOdf.sZone = tmOdf.sZone) Then
                        If tlOdf.iSeqNo > ilSeqNo Then
                            ilSeqNo = tlOdf.iSeqNo
                        End If
                    End If
                End If
            Next llOdf
        End If
    End If
End Sub

Private Sub mProcSpot(slAdjLocalOrFeed As String, llDate As Long, llSDate As Long, llEDate As Long, slInLogType As String, slFor As String, ilSeqNo As Integer, hlODF As Integer, tlOdf As ODF, ilODFVefCode As Integer, ilSimVefCode As Integer, ilVehCode As Integer, ilSSFType As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, llCifCode As Long, tlCxf As CXF, ilType As Integer, ilVpfIndex As Integer, tlAdf As ADF, hlVsf As Integer, hlSif As Integer, ilGenLST As Integer, ilLSTForLogVeh As Integer, hlMcf As Integer, ilExportType As Integer, llGsfCode As Long, llCrfCode As Long)
    Dim ilRet As Integer
    Dim ilTZone As Integer
    Dim ilPE As Integer
    Dim slLength As String
    Dim tlCxfSrchKey As LONGKEY0    'CXF key record image
    Dim tlAdfSrchKey As INTKEY0 'ADF key record image
    Dim slSplitCopyFlag As String * 1
    Dim ilAddSpot As Integer
    Dim ilCxfRecLen As Integer
    Dim ilAdfRecLen As Integer
    Dim llUpper As Long
    Dim ilOdfVff As Integer
    
    ilCxfRecLen = Len(tlCxf)
    ilAdfRecLen = Len(tlAdf)
    '8/13/14: Generate spots for last date+1 into lst with Fed as *
    If bmCreatingLstDate1 Then
        For ilTZone = 0 To UBound(tmZoneInfo) - 1 Step 1
            If Left$(tmDlf.sZone, 1) = tmZoneInfo(ilTZone).sZone Then
                If tmZoneInfo(ilTZone).sFed <> "*" Then
                    'Return
                    Exit Sub
                End If
            End If
        Next ilTZone
        lmGLocalAdj = 0
        For ilTZone = 0 To UBound(tmZoneInfo) - 1 Step 1
            If Left$(tmDlf.sZone, 1) = tmZoneInfo(ilTZone).sFed Then
                If tmZoneInfo(ilTZone).lGLocalAdj < lmGLocalAdj Then
                    lmGLocalAdj = tmZoneInfo(ilTZone).lGLocalAdj
                End If
            End If
        Next ilTZone
        If lmGLocalAdj >= 0 Then
            'Return
            Exit Sub
        End If
        lmGLocalAdj = -lmGLocalAdj
    End If
    If ((Asc(tgSpf.sUsingFeatures6) And BBNOTSEPARATELINE) = BBNOTSEPARATELINE) Then
        If (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
            'Return
            Exit Sub
        End If
    End If
    If (tmAvail.iAvInfo And SSXMID) = SSXMID Then
        smXMid = "Y"
    Else
        smXMid = "N"
    End If
    tmOdf.iUrfCode = tgUrf(0).iCode
    If ilODFVefCode <= 0 Then
        If ilSimVefCode <= 0 Then
            If tmVef.iVefCode > 0 Then  'Log vehicle defined
                tmOdf.iVefCode = tmVef.iVefCode
                tmOdf.iAlternateVefCode = ilVehCode     '1-15-14 for log vehicles, put name in this field to separate the vehicles on output
            Else
                tmOdf.iVefCode = ilVehCode
                tmOdf.iAlternateVefCode = ilVehCode
            End If
        Else
            tmOdf.iVefCode = ilSimVefCode
            tmOdf.iAlternateVefCode = ilSimVefCode
        End If
    Else
        tmOdf.iVefCode = ilODFVefCode
        tmOdf.iAlternateVefCode = ilODFVefCode
    End If
    If (tmVef.sType = "G") And (tgVpf(ilVpfIndex).sGenLog <> "M") Then
        tmOdf.iGameNo = ilSSFType
    Else
        tmOdf.iGameNo = 0
    End If
    tmOdf.iAirDate(0) = ilLogDate0
    tmOdf.iAirDate(1) = ilLogDate1
    tmOdf.iAirTime(0) = tmAvail.iTime(0)
    tmOdf.iAirTime(1) = tmAvail.iTime(1)
    tmOdf.iLocalTime(0) = tmDlf.iLocalTime(0)
    tmOdf.iLocalTime(1) = tmDlf.iLocalTime(1)
    tmOdf.iFeedTime(0) = tmDlf.iFeedTime(0)
    tmOdf.iFeedTime(1) = tmDlf.iFeedTime(1)
    tmOdf.sZone = tmDlf.sZone
    tmOdf.iEtfCode = 0
    tmOdf.iEnfCode = tmDlf.iEnfCode 'Required for Commercial schedule
    tmOdf.sProgCode = tmDlf.sProgCode
    tmOdf.iMnfFeed = tmDlf.iMnfFeed
    tmOdf.iDPSort = 0   '5-31-01 tmOdf.sUnused1 = "" 'tmDlf.sBus
    'tmOdf.iWkNo = 0    'tmDlf.sSchedule
    tmOdf.ianfCode = tmAvail.ianfCode
    tmOdf.iUnits = 0
    slLength = Trim$(str$(tmSdf.iLen)) & "s"
    gPackLength slLength, tmOdf.iLen(0), tmOdf.iLen(1)
    tmOdf.sBBDesc = ""
    If ((Asc(tgSpf.sUsingFeatures6) And BBNOTSEPARATELINE) = BBNOTSEPARATELINE) Then
        If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
            If tmClf.iBBOpenLen > 0 Then
                tmOdf.sBBDesc = "BB"
            End If
        Else
            If (tmClf.iBBOpenLen > 0) And (tmClf.iBBCloseLen > 0) Then
                tmOdf.sBBDesc = "O/C BB"
            ElseIf tmClf.iBBOpenLen > 0 Then
                tmOdf.sBBDesc = "O BB"
            ElseIf tmClf.iBBCloseLen > 0 Then
                tmOdf.sBBDesc = "C BB"
            End If
        End If
    Else
        If (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
            If ((Asc(tgSpf.sUsingFeatures6) And BBCLOSEST) = BBCLOSEST) Then
                tmOdf.sBBDesc = "BB"
            Else
                If tmSdf.sSpotType = "O" Then
                    tmOdf.sBBDesc = "O BB"
                ElseIf tmSdf.sSpotType = "C" Then
                    tmOdf.sBBDesc = "C BB"
                End If
            End If
        End If
    End If
    tmOdf.iAdfCode = tmChf.iAdfCode
    'Test tmSdf.sPtType
    tmOdf.lCifCode = llCifCode
    'If Trim$(tmChf.sProduct) <> "" Then
    '    tmOdf.sProduct = tmChf.sProduct
    'Else
    '    tmOdf.sProduct = "" '"???? Name ????"
    'End If
    mGetCpf hmCif, llCifCode
    If Trim$(tmCpf.sName) <> "" Then
        tmOdf.sProduct = tmCpf.sName
    Else
        If Trim$(tmChf.sProduct) <> "" Then
            tmOdf.sProduct = tmChf.sProduct
        Else
            tmOdf.sProduct = "" '"???? Name ????"
        End If
    End If
    If tmChf.lCode = 0 Then                 'feed spot, indicate it with advt/prod
        tmOdf.sProduct = Trim$(tmOdf.sProduct) & " (Feed)"
    End If

    tmOdf.iMnfSubFeed = tmDlf.iMnfSubFeed
    tmOdf.lCntrNo = tmChf.lCntrNo

    '2-15-01 Setup comment pointers only if show = yes on Log
    tmOdf.lchfcxfCode = 0                     'assume no comment
    If tmChf.lCxfCode > 0 Then
        tlCxfSrchKey.lCode = tmChf.lCxfCode      'comment  code
        ilCxfRecLen = Len(tlCxf)
        ilRet = btrGetEqual(hmCxf, tlCxf, ilCxfRecLen, tlCxfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY) 'find matching comment recd
        If ilRet = BTRV_ERR_NONE Then
            If tlCxf.sShSpot = "Y" Then         'show comment on log
                tmOdf.lchfcxfCode = tmChf.lCxfCode
            End If
        End If
    End If
    If slAdjLocalOrFeed = "W" Then
        tmOdf.lFt1CefCode = tmChf.lCode
    End If
    'tmOdf.lChfCxfCode = tmChf.lCxfCode      '2-15-01, only place if showing on log-Other comment code (1/17/99)
    If (tmSdf.sSpotType = "O") Or (tmSdf.sSpotType = "C") Then
        tmOdf.iBreakNo = 0
    Else
        tmOdf.iBreakNo = imBreakNo
    End If
    tmOdf.iPositionNo = imPositionNo
    tmOdf.iType = ilType
    tmOdf.lCefCode = lmCrfCsfCode   '0
    tmOdf.sBonus = smBonus          '2-15-01 Bonus flag (B= bonus, f = fill)
    tmOdf.lAvailcefCode = lmAvailCefCode            'comment ptr from avail
    If tlAdf.iCode <> tmChf.iAdfCode Then
        tlAdfSrchKey.iCode = tmChf.iAdfCode
        ilRet = btrGetEqual(hmAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    End If
    tmOdf.lEvtIDCefCode = lmEvtIDCefCode
    tmOdf.sDupeAvailID = tmDlf.sBus
    tmOdf.sShortTitle = gGetShortTitle(hlVsf, hmClf, hlSif, tmChf, tlAdf, tmSdf)
    If tgSpf.sCUseSegments = "Y" Then
        tmOdf.imnfSeg = tmChf.imnfSeg       '6-19-01
    Else            'L75 used for Podcast client wants to see word IMPR on the log
        tmOdf.imnfSeg = 0
        If tgSaf(0).sHideDemoOnBR = "Y" And tmChf.sHideDemo = "Y" Then      '6-13-19 if hiding the demo , show IMPR?
            tmOdf.imnfSeg = 1       'flag to indicate to show word IMPR on L75.  Individual logs need to be changed to test the flag and implement
        End If
    End If
    tmOdf.sPageEjectFlag = "N"
    For ilPE = LBound(tmPageEject) To UBound(tmPageEject) - 1 Step 1
        If (lmEvtTime >= tmPageEject(ilPE).lTime) And ((tmPageEject(ilPE).ianfCode = tmAvail.ianfCode) Or (tmPageEject(ilPE).ianfCode = 0)) Then
            tmOdf.sPageEjectFlag = "Y"
            tmPageEject(ilPE).lTime = 999999
            tmPageEject(ilPE).ianfCode = -1
            Exit For
        End If
    Next ilPE
    '11-11-09 show the line comment on logs that are coded for it.  Dual use of the field OdfEvtCefCode .  A spot record will contain the line comment ptr
    mDPDaysTimes hmRdf, smEDIDays, lmEvtCefCode             'Read clf & cff & rdf to format the DP description (or sch line override)
    tmOdf.lRafCode = 0              '5-16-08
    tmOdf.sSplitNetwork = "N"
    '10/23/13: Check for split network, then split copy
    '10/24/13: added test if using split network
    'split networks
    If (Asc(tgSpf.sUsingFeatures2) And SPLITNETWORKS) = SPLITNETWORKS Then
        If (tmSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
            tmOdf.sSplitNetwork = "P"
            tmOdf.lRafCode = tmClf.lRafCode
        ElseIf (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
            tmOdf.sSplitNetwork = "S"
            tmOdf.lRafCode = tmClf.lRafCode
        End If
    End If
    slSplitCopyFlag = ""
    'If (Asc(tgSpf.sUsingFeatures2) And SPLITCOPY = SPLITCOPY) Then
    If ((Asc(tgSpf.sUsingFeatures2) And SPLITCOPY) = SPLITCOPY) And (tmOdf.sSplitNetwork = "N") Then
        tmRsfSrchKey1.lCode = tmSdf.lCode
        ilRet = btrGetEqual(hmRsf, tmRsf, Len(tmRsf), tmRsfSrchKey1, INDEXKEY1, BTRV_LOCK_NONE, SETFORREADONLY)   'Get first record as starting point of extend operation
        Do While (ilRet = BTRV_ERR_NONE) And (tmRsf.lSdfCode = tmSdf.lCode)
            If (tmRsf.sType <> "A") Then
                'tmOdf.sSplitNetwork = "S"
                slSplitCopyFlag = "S"
                Exit Do
            End If
            ilRet = btrGetNext(hmRsf, tmRsf, imRsfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    '10/23/13: Moved above Split Copy
    'Else
    '    'split networks
    '    If (tmSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
    '        tmOdf.sSplitNetwork = "P"
    '        tmOdf.lRafCode = tmClf.lRafCode
    '    ElseIf (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
    '        tmOdf.sSplitNetwork = "S"
    '        tmOdf.lRafCode = tmClf.lRafCode
    '    End If
    End If
    
    'Same code in gGetShortTitle- it is here to save execution time
    'If tgSpf.sUseProdSptScr = "P" Then
    '    If llSifCode <= 0 Then  'llSifCode obtained from Crf in mObtainCrfCsfCode
    '        llSifCode = tmChf.lSifCode
    '    End If
    '    If llSifCode > 0 Then
    '        tlSifSrchKey.lCode = llSifCode
    '        ilRet = btrGetEqual(hlSif, tlSif, ilSifRecLen, tlSifSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '        If ilRet = BTRV_ERR_NONE Then
    '            tmOdf.sShortTitle = tlSif.sName
    '        Else
    '            tmOdf.sShortTitle = ""  '"???? Name ????"
    '        End If
    '    Else
    '        tmOdf.sShortTitle = ""  '"???? Name ????"
    '    End If
    'Else
    '    If tlAdf.iCode <> tmChf.iAdfCode Then
    '        tlAdfSrchKey.iCode = tmChf.iAdfCode
    '        ilRet = btrGetEqual(hlAdf, tlAdf, ilAdfRecLen, tlAdfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)
    '    Else
    '        ilRet = BTRV_ERR_NONE
    '    End If
    '    If ilRet = BTRV_ERR_NONE Then
    '        tmOdf.sShortTitle = Trim$(tlAdf.sAbbr) & "," & Trim$(tmChf.sProduct)
    '    Else
    '        tmOdf.sShortTitle = ""
    '    End If
'    'End If
    'Determine seq number
    '6/4/16: Replaced GoSub
    'GoSub lProcTestForDuplBB
    mProcTestForDuplBB ilAddSpot
    If Not ilAddSpot Then
        'Return
        Exit Sub
    End If
    '6/4/16: Replaced GoSub
    'GoSub lProcAdjDate
    mProcAdjDate slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType
    '6/4/16: Replaced GoSub
    'GoSub lProcSeqNo
    mProcSeqNo ilSeqNo, slFor, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode
    
    tmOdf.iSortSeq = 0                      'ensure all the spots are not separated
    tmOdf.iSeqNo = ilSeqNo + 1
    tmOdf.iDaySort = imDaySort
    tmOdf.lEvtCefCode = lmEvtCefCode
    tmOdf.iEvtCefSort = imEvtCefSort
    tmOdf.sLogType = smLogType
'    tmOdf.sSplitNetwork = "N"
'    tmOdf.lRafCode = 0              '5-16-08
'    If (tmSpot.iRecType And SSSPLITPRI) = SSSPLITPRI Then
'        tmOdf.sSplitNetwork = "P"
'        tmOdf.lRafCode = tmClf.lRafCode
'    ElseIf (tmSpot.iRecType And SSSPLITSEC) = SSSPLITSEC Then
'        tmOdf.sSplitNetwork = "S"
'        tmOdf.lRafCode = tmClf.lRafCode
'    End If

    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, lmAvailTime
    If lmPrevAvailTime <> lmAvailTime Then
        tmOdf.iAvailLen = tmAvail.iLen
        lmPrevAvailTime = lmAvailTime
        'BB avail
        If tmAvail.iLen = 0 Then
            tmOdf.iAvailLen = -1
        End If
    Else
        tmOdf.iAvailLen = 0
    End If
    tmOdf.sAvailLock = "N"
    '7-25-13 combine 2 definitions into this field:  Locked avails/spots + split copy defined
    If ((tmAvail.iAvInfo And SSLOCK) = SSLOCK) And ((tmAvail.iAvInfo And SSLOCKSPOT) = SSLOCKSPOT) Then
        tmOdf.sAvailLock = "B"
        If slSplitCopyFlag = "S" Then
            tmOdf.sAvailLock = "D"              'locked avail & spot plus a split copy
        End If
    ElseIf ((tmAvail.iAvInfo And SSLOCK) = SSLOCK) Then
        tmOdf.sAvailLock = "A"
        If slSplitCopyFlag = "S" Then
            tmOdf.sAvailLock = "E"              'locked avail  plus a split copy
        End If
    ElseIf ((tmAvail.iAvInfo And SSLOCKSPOT) = SSLOCKSPOT) Then
        tmOdf.sAvailLock = "S"
        If slSplitCopyFlag = "S" Then
            tmOdf.sAvailLock = "E"              'locked  spot plus a split copy
        End If
    Else
        If slSplitCopyFlag = "S" Then
            tmOdf.sAvailLock = "F"              'nothing locked, but split copy
        End If
    End If
    'imGenODF set for affiliate system
    If (tgSpf.sGUseAffSys <> "Y") Or (imGenODF) Then
        If imLstExist And ilGenLST Then
            If (tmSdf.sSpotType <> "O") Or (tmSdf.sSpotType <> "C") Or ((tmSdf.sSpotType = "O") And (tgSpf.sBBsToAff = "Y")) Or ((tmSdf.sSpotType = "C") And (tgSpf.sBBsToAff = "Y")) Then
                For ilTZone = 0 To UBound(tmZoneInfo) - 1 Step 1
                    If Left$(tmOdf.sZone, 1) = tmZoneInfo(ilTZone).sZone Then
                        If tmZoneInfo(ilTZone).sFed = "*" Then
                            'Added when changed to generate Affiliate with each Conventional Vehicle instead of
                            'the Log Vehicle.  11/20/03.  Passed vefCode instead of getting it from tmOdf.iVefCode
                            '1/11/08:  Match how odf is created for combination games (remove game test). i.e. produce lst as a combination of game vehicles
                            'If (ilODFVefCode <= 0) Or (tmVef.sType = "G") Then
                            If (ilODFVefCode <= 0) Then
                                If ilSimVefCode <= 0 Then
                                    '11/4/09: re-add generation of LST by Log vehicle
                                    'mCreateLst ilVehCode, ilClearLstSdf, hmClf, hmCif, hlMcf, slLogType, ilCreateNewLST, ilExportType, llGsfCode, ilWegenerOLA
                                    If (ilLSTForLogVeh > 0) And (tmVef.iVefCode > 0) Then
                                        mCreateLst tmVef.iVefCode, imClearLstSdf, hmClf, hmCif, hlMcf, slInLogType, imCreateNewLST, ilExportType, llGsfCode, imWegenerOLA, smAlertStatus
                                    Else
                                        mCreateLst ilVehCode, imClearLstSdf, hmClf, hmCif, hlMcf, slInLogType, imCreateNewLST, ilExportType, llGsfCode, imWegenerOLA, smAlertStatus
                                    End If
                                Else
                                    mCreateLst ilSimVefCode, imClearLstSdf, hmClf, hmCif, hlMcf, slInLogType, imCreateNewLST, ilExportType, llGsfCode, imWegenerOLA, smAlertStatus
                                End If
                            Else
                                mCreateLst ilODFVefCode, imClearLstSdf, hmClf, hmCif, hlMcf, slInLogType, imCreateNewLST, ilExportType, llGsfCode, imWegenerOLA, smAlertStatus
                            End If
                            'End of Change.
                            
                            ' 3-25-10 remove overriding short title with cart (thats only in affiliate system)
                            ' L31A was created on Affiliate side to show the short title, change L31 back to using the copy pointers
                            ' due to Music of your Life import, copy must be obtained from
                            ' the short title field as text rather than using a copy pointer
                            'tmOdf.sShortTitle = tmLst.sCart
                        End If
                        Exit For
                    End If
                Next ilTZone
            End If
        End If
        '8/29/16: Test if comment should be suppressed on the Log
        ilOdfVff = gBinarySearchVff(tmOdf.iVefCode)
        If ilOdfVff <> -1 Then
            If tgVff(ilOdfVff).sHideCommOnLog = "Y" Then
                tmOdf.lCefCode = 0
            End If
        End If
        '8/13/14: Generate spots for last date+1 into lst with Fed as *
        If Not bmCreatingLstDate1 Then
            tmOdf.lCode = 0
            ilRet = btrInsert(hlODF, tmOdf, imOdfRecLen, INDEXKEY3)
            tgOdfSdfCodes(UBound(tgOdfSdfCodes)).lOdfCode = tmOdf.lCode
            tgOdfSdfCodes(UBound(tgOdfSdfCodes)).lSdfCode = tmSdf.lCode
            ReDim Preserve tgOdfSdfCodes(0 To UBound(tgOdfSdfCodes) + 1) As ODFSDFCODES
            gLogBtrError ilRet, "gBuildODFSpotDay: Insert #3"
            If (tgSpf.sCBlackoutLog = "Y") Or (igBkgdProg = 3) Then
                llUpper = UBound(tgSpotSum)
                tgSpotSum(llUpper).iVefCode = tmOdf.iVefCode
                gUnpackDateLong tmOdf.iAirDate(0), tmOdf.iAirDate(1), tgSpotSum(llUpper).lDate
                tgSpotSum(llUpper).lChfCode = tmChf.lCode
                tgSpotSum(llUpper).iMnfComp(0) = tmChf.iMnfComp(0)
                tgSpotSum(llUpper).iMnfComp(1) = tmChf.iMnfComp(1)
                tgSpotSum(llUpper).iAdfCode = tmChf.iAdfCode
                tgSpotSum(llUpper).iLen = tmSdf.iLen
                tgSpotSum(llUpper).sProduct = tmChf.sProduct
                If Not imLstExist Or Not ilGenLST Then
                    tgSpotSum(llUpper).sShortTitle = tmOdf.sShortTitle
                End If
                tgSpotSum(llUpper).imnfSeg = tmChf.imnfSeg          '6-19-01
                gUnpackTimeLong tmOdf.iLocalTime(0), tmOdf.iLocalTime(1), False, tgSpotSum(llUpper).lTime
                tgSpotSum(llUpper).sZone = tmOdf.sZone   'tmDlf.sZone
                tgSpotSum(llUpper).iSeqNo = tmOdf.iSeqNo
                If imLstExist Then
                    tgSpotSum(llUpper).lLstCode = tmLst.lCode
                Else
                    tgSpotSum(llUpper).lLstCode = 0
                End If
                tgSpotSum(llUpper).iLnVefCode = tmClf.iVefCode
                tgSpotSum(llUpper).lSdfCode = tmSdf.lCode
                tgSpotSum(llUpper).sLogType = smLogType
                tgSpotSum(llUpper).lCrfCode = llCrfCode
                tgSpotSum(llUpper).sDays = smEDIDays
                tgSpotSum(llUpper).iOrigAirDate(0) = ilLogDate0
                tgSpotSum(llUpper).iOrigAirDate(1) = ilLogDate1
                ReDim Preserve tgSpotSum(0 To llUpper + 1) As SPOTSUM
            End If
        '8/13/14: Generate spots for last date+1 into lst with Fed as *
        Else
            tmDate1ODF(UBound(tmDate1ODF)) = tmOdf
            ReDim Preserve tmDate1ODF(0 To UBound(tmDate1ODF) + 1) As ODF
        End If
    End If
End Sub

Private Sub mMakeAvails(ilODFVefCode As Integer, ilSimVefCode As Integer, ilVpfIndex As Integer, ilSSFType As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, slAdjLocalOrFeed As String, llDate As Long, llSDate As Long, llEDate As Long, slInLogType As String, ilSeqNo As Integer, slFor As String, hlODF As Integer, tlOdf As ODF, ilVehCode As Integer, ilGenLST As Integer, ilLSTForLogVeh As Integer, ilSec As Integer, ilUnits As Integer, llGsfCode As Long)
    Dim ilTZone As Integer
    If ilODFVefCode <= 0 Then
        If ilSimVefCode <= 0 Then
            If tmVef.iVefCode > 0 Then  'Log vehicle defined
                tmOdf.iVefCode = tmVef.iVefCode
            Else
                tmOdf.iVefCode = ilVehCode
            End If
        Else
            tmOdf.iVefCode = ilSimVefCode
        End If
    Else
        tmOdf.iVefCode = ilODFVefCode
    End If
    If (tmVef.sType = "G") And (tgVpf(ilVpfIndex).sGenLog <> "M") Then
        tmOdf.iGameNo = ilSSFType
    Else
        tmOdf.iGameNo = 0
    End If
    tmOdf.iAirDate(0) = ilLogDate0
    tmOdf.iAirDate(1) = ilLogDate1
    If (tmAvail.iAvInfo And SSXMID) = SSXMID Then
        smXMid = "Y"
    Else
        smXMid = "N"
    End If
    tmOdf.iAirTime(0) = tmAvail.iTime(0)  'tmAvAvail.iTime(0)
    tmOdf.iAirTime(1) = tmAvail.iTime(1)  'tmAvAvail.iTime(1)
    tmOdf.iLocalTime(0) = tmDlf.iLocalTime(0)
    tmOdf.iLocalTime(1) = tmDlf.iLocalTime(1)
    tmOdf.iFeedTime(0) = tmDlf.iFeedTime(0)
    tmOdf.iFeedTime(1) = tmDlf.iFeedTime(1)
    tmOdf.sZone = tmDlf.sZone
    tmOdf.iBreakNo = imBreakNo
    tmOdf.iPositionNo = imPositionNo
    tmOdf.ianfCode = tmAvail.ianfCode
    '6/4/16: Replaced GoSub
    'GoSub lProcAdjDate
    mProcAdjDate slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType
    '6/4/16: Replaced GoSub
    'GoSub lProcSeqNo
    mProcSeqNo ilSeqNo, slFor, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode
    tmOdf.iSeqNo = ilSeqNo + 1
    'imGenODF set for affiliate system
    If (tgSpf.sGUseAffSys <> "Y") Or (imGenODF) Then
        If imLstExist And ilGenLST Then
            For ilTZone = 0 To UBound(tmZoneInfo) - 1 Step 1
                If Left$(tmOdf.sZone, 1) = tmZoneInfo(ilTZone).sZone Then
                    If tmZoneInfo(ilTZone).sFed = "*" Then
                        'Added when changed to generate Affiliate with each Conventional Vehicle instead of
                        'the Log Vehicle.  11/20/03.  Passed vefCode instead of getting it from tmOdf.iVefCode
                        '1/11/08:  Match how odf is created for combination games (remove game test). i.e. produce lst as a combination of game vehicles
                        'If (ilODFVefCode <= 0) Or (tmVef.sType = "G") Then
                        If (ilODFVefCode <= 0) Then
                            If ilSimVefCode <= 0 Then
                                '11/4/09: Re-add generation of lst by Log vehicle
                                'mCreateAvailLst ilVehCode, ilSec, ilUnits, llGsfCode
                                If (ilLSTForLogVeh > 0) And (tmVef.iVefCode > 0) Then
                                    mCreateAvailLst tmVef.iVefCode, ilSec, ilUnits, llGsfCode
                                Else
                                    mCreateAvailLst ilVehCode, ilSec, ilUnits, llGsfCode
                                End If
                            Else
                                mCreateAvailLst ilSimVefCode, ilSec, ilUnits, llGsfCode
                            End If
                        Else
                            mCreateAvailLst ilODFVefCode, ilSec, ilUnits, llGsfCode
                        End If
                        'End of Change
                    End If
                    Exit For
                End If
            Next ilTZone
        End If
    End If
End Sub

Private Sub mGetDlfAvail(ilDlfFound As Integer, slDay As String, ilLoop As Integer, tlLLC() As LLC, hlDlf As Integer, ilDlfDate0 As Integer, ilDlfDate1 As Integer, ilODFVefCode As Integer, ilSimVefCode As Integer, ilVpfIndex As Integer, ilSSFType As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, slAdjLocalOrFeed As String, llDate As Long, llSDate As Long, llEDate As Long, slInLogType As String, ilSeqNo As Integer, slFor As String, hlODF As Integer, tlOdf As ODF, ilVehCode As Integer, ilGenLST As Integer, ilLSTForLogVeh As Integer, ilSec As Integer, ilUnits As Integer, llGsfCode As Long, ilOtherGen As Integer, ilTerminated As Integer)
    Dim ilZone As Integer
    Dim slZone As String
    Dim llAirTime As Long
    Dim llTime As Long
    Dim ilRet As Integer
    
    If (ilDlfFound) Then
        'Obtain delivery entry to see is avail is sent
        tmDlfSrchKey.iVefCode = ilVehCode
        tmDlfSrchKey.sAirDay = slDay
        tmDlfSrchKey.iStartDate(0) = ilDlfDate0
        tmDlfSrchKey.iStartDate(1) = ilDlfDate1
        tmDlfSrchKey.iAirTime(0) = tmAvail.iTime(0) 'tmAvAvail.iTime(0)
        tmDlfSrchKey.iAirTime(1) = tmAvail.iTime(1) 'tmAvAvail.iTime(1)
        ilRet = btrGetEqual(hlDlf, tmDlf, imDlfRecLen, tmDlfSrchKey, INDEXKEY0, BTRV_LOCK_NONE, SETFORREADONLY)   'Get last current record to obtain date
        Do While (ilRet = BTRV_ERR_NONE) And (tmDlf.iVefCode = ilVehCode) And (tmDlf.sAirDay = slDay) And (tmDlf.iStartDate(0) = ilDlfDate0) And (tmDlf.iStartDate(1) = ilDlfDate1) And (tmDlf.iAirTime(0) = tmAvAvail.iTime(0)) And (tmDlf.iAirTime(1) = tmAvAvail.iTime(1))
            ilTerminated = False
            If (tmDlf.sCmmlSched = "N") And (tmDlf.iMnfSubFeed = 0) Then
                ilTerminated = True
            Else
                If (tmDlf.iTermDate(1) <> 0) Or (tmDlf.iTermDate(0) <> 0) Then
                    If (tmDlf.iTermDate(1) < tmDlf.iStartDate(1)) Or ((tmDlf.iStartDate(1) = tmDlf.iTermDate(1)) And (tmDlf.iTermDate(0) < tmDlf.iStartDate(0))) Then
                        ilTerminated = True
                    End If
                End If
            End If
            If Not ilTerminated Then
                If (tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode) And (tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode) Then
                    'If slFor = "C" Then
                        If (tmDlf.sCmmlSched = "Y") Or (tmDlf.iMnfSubFeed <> 0) Then
                            tmDlf.iMnfFeed = 0
                            '6/5/16: Replaced GoSub
                            'GoSub lMakeAvails
                            mMakeAvails ilODFVefCode, ilSimVefCode, ilVpfIndex, ilSSFType, ilLogDate0, ilLogDate1, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlODF, tlOdf, ilVehCode, ilGenLST, ilLSTForLogVeh, ilSec, ilUnits, llGsfCode
                        End If
                    'End If
                End If
            End If
            ilRet = btrGetNext(hlDlf, tmDlf, imDlfRecLen, BTRV_LOCK_NONE, SETFORREADONLY)
        Loop
    Else
        For ilZone = 0 To UBound(tmZoneInfo) - 1 Step 1
            If ((tmZoneInfo(ilZone).sFed = "*") And (tgSpf.sGUseAffSys = "Y")) Or (tgSpf.sGUseAffSys <> "Y") Then
                Select Case tmZoneInfo(ilZone).sZone
                    Case "E"
                        slZone = "EST"
                    Case "M"
                        slZone = "MST"
                    Case "C"
                        slZone = "CST"
                    Case "P"
                        slZone = "PST"
                    Case Else
                        If ilOtherGen Then
                            Exit For
                        End If
                        slZone = ""
                End Select
                If (tgSpf.sGUseAffSys = "Y") Then
                    'gUnpackTimeLong tmAvAvail.iTime(0), tmAvAvail.iTime(1), False, llAirTime
                    gUnpackTimeLong tmAvail.iTime(0), tmAvail.iTime(1), False, llAirTime
                    llTime = llAirTime + tmZoneInfo(ilZone).lGLocalAdj
                    If llTime < 0 Then
                        llTime = llTime + 86400
                    ElseIf llTime > 86400 Then
                        llTime = llTime - 86400
                    End If
                    gPackTimeLong llTime, tmDlf.iLocalTime(0), tmDlf.iLocalTime(1)
                    llTime = llAirTime + tmZoneInfo(ilZone).lGFeedAdj
                    If llTime < 0 Then
                        llTime = llTime + 86400
                    ElseIf llTime > 86400 Then
                        llTime = llTime - 86400
                    End If
                    gPackTimeLong llTime, tmDlf.iFeedTime(0), tmDlf.iFeedTime(1)
                Else
                    'tmDlf.iLocalTime(0) = tmAvAvail.iTime(0)
                    'tmDlf.iLocalTime(1) = tmAvAvail.iTime(1)
                    'tmDlf.iFeedTime(0) = tmAvAvail.iTime(0)
                    'tmDlf.iFeedTime(1) = tmAvAvail.iTime(1)
                    tmDlf.iLocalTime(0) = tmAvail.iTime(0)
                    tmDlf.iLocalTime(1) = tmAvail.iTime(1)
                    tmDlf.iFeedTime(0) = tmAvail.iTime(0)
                    tmDlf.iFeedTime(1) = tmAvail.iTime(1)
                End If
                tmDlf.sZone = slZone
                tmDlf.iEtfCode = tlLLC(ilLoop).iEtfCode
                tmDlf.iEnfCode = tlLLC(ilLoop).iEnfCode
                tmDlf.sProgCode = ""
                tmDlf.iMnfFeed = 0
                tmDlf.sBus = ""
                tmDlf.sSchedule = ""
                tmDlf.iMnfSubFeed = 0
                '6/5/16: Replaced GoSub
                'GoSub lMakeAvails
                mMakeAvails ilODFVefCode, ilSimVefCode, ilVpfIndex, ilSSFType, ilLogDate0, ilLogDate1, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlODF, tlOdf, ilVehCode, ilGenLST, ilLSTForLogVeh, ilSec, ilUnits, llGsfCode
            End If
        Next ilZone
    End If
End Sub

Private Sub mChkOpenAvail(ilAvVefCode As Integer, ilAvEvt As Integer, ilDlfFound As Integer, slDay As String, ilLoop As Integer, tlLLC() As LLC, hlDlf As Integer, ilDlfDate0 As Integer, ilDlfDate1 As Integer, ilODFVefCode As Integer, ilSimVefCode As Integer, ilVpfIndex As Integer, ilSSFType As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, slAdjLocalOrFeed As String, llDate As Long, llSDate As Long, llEDate As Long, slInLogType As String, ilSeqNo As Integer, slFor As String, hlODF As Integer, tlOdf As ODF, ilVehCode As Integer, ilGenLST As Integer, ilLSTForLogVeh As Integer, ilSec As Integer, ilUnits As Integer, llGsfCode As Long, ilOtherGen As Integer, ilTerminated As Integer)
    Dim ilAvVpfIndex As Integer
    Dim ilVpf As Integer
    Dim slUnits As String
    Dim ilSpot As Integer
    Dim slStr As String
    Dim slSpotLen As String
    
    If (tgSpf.sCBlackoutLog <> "Y") Or (Not imLstExist) Or (Not ilGenLST) Then
        'Return
        Exit Sub
    End If

    ilAvVpfIndex = -1
    'For ilVpf = 0 To UBound(tgVpf) Step 1  'gVpfRead called in signon
    '    If ilAvVefCode = tgVpf(ilVpf).iVefKCode Then
        ilVpf = gBinarySearchVpf(ilAvVefCode)
        If ilVpf <> -1 Then
            ilAvVpfIndex = ilVpf
    '        Exit For
        End If
    'Next ilVpf
    If ilAvVpfIndex < 0 Then
        'Return
        Exit Sub
    End If
    If (tgVpf(ilAvVpfIndex).sSSellOut = "B") Or (tgVpf(ilAvVpfIndex).sSSellOut = "U") Or (tgVpf(ilAvVpfIndex).sSSellOut = "M") Then
        ilUnits = tmAvAvail.iAvInfo And &H1F
        slUnits = Trim$(str$(ilUnits)) & ".0"   'For units as thirty
        ilSec = tmAvAvail.iLen
    Else
        ilUnits = tmAvAvail.iAvInfo And &H1F
        ilSec = 0
    End If
    For ilSpot = 1 To tmAvAvail.iNoSpotsThis Step 1
        LSet tmAvSpot = tmCTSsf.tPas(ADJSSFPASBZ + ilAvEvt + ilSpot)
        If (tmAvSpot.iRecType And SSSPLITSEC) <> SSSPLITSEC Then
            If (tgVpf(ilAvVpfIndex).sSSellOut = "B") Or (tgVpf(ilAvVpfIndex).sSSellOut = "U") Then
                ilUnits = ilUnits - 1
                ilSec = ilSec - (tmAvSpot.iPosLen And &HFFF)
            ElseIf tgVpf(ilAvVpfIndex).sSSellOut = "M" Then
                ilUnits = ilUnits - 1
                ilSec = ilSec - (tmAvSpot.iPosLen And &HFFF)
            ElseIf tgVpf(ilAvVpfIndex).sSSellOut = "T" Then
                slSpotLen = Trim$(str$(tmAvSpot.iPosLen And &HFFF))
                slStr = gDivStr(slSpotLen, "30.0")
                slUnits = gSubStr(slUnits, slSpotLen)
            End If
        End If
    Next ilSpot
    If (tgVpf(ilAvVpfIndex).sSSellOut = "B") Or (tgVpf(ilAvVpfIndex).sSSellOut = "U") Or (tgVpf(ilAvVpfIndex).sSSellOut = "M") Then
        If (ilUnits > 0) And (ilSec > 0) Then
            '6/5/16: Replaced GoSub
            'GoSub lGetDlfAvail
            mGetDlfAvail ilDlfFound, slDay, ilLoop, tlLLC(), hlDlf, ilDlfDate0, ilDlfDate1, ilODFVefCode, ilSimVefCode, ilVpfIndex, ilSSFType, ilLogDate0, ilLogDate1, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlODF, tlOdf, ilVehCode, ilGenLST, ilLSTForLogVeh, ilSec, ilUnits, llGsfCode, ilOtherGen, ilTerminated
        End If
    ElseIf tgVpf(ilVpfIndex).sSSellOut = "T" Then
        If gCompNumberStr(slUnits, "0.0") > 0 Then
            ilSec = Val(slUnits)
            '6/5/16: Replaced GoSub
            'GoSub lGetDlfAvail
            mGetDlfAvail ilDlfFound, slDay, ilLoop, tlLLC(), hlDlf, ilDlfDate0, ilDlfDate1, ilODFVefCode, ilSimVefCode, ilVpfIndex, ilSSFType, ilLogDate0, ilLogDate1, slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType, ilSeqNo, slFor, hlODF, tlOdf, ilVehCode, ilGenLST, ilLSTForLogVeh, ilSec, ilUnits, llGsfCode, ilOtherGen, ilTerminated
        End If
    End If
End Sub

Private Sub mProcProg(ilDlfFound As Integer, ilOtherGen As Integer, ilVpfIndex As Integer, slAdjLocalOrFeed As String, llDate As Long, llSDate As Long, llEDate As Long, slInLogType As String, ilSeqNo As Integer, slFor As String, hlODF As Integer, tlOdf As ODF, ilODFVefCode As Integer, ilSimVefCode As Integer, ilVehCode As Integer, ilType As Integer, ilSSFType As Integer, ilLogDate0 As Integer, ilLogDate1 As Integer, ilStartTime0 As Integer, ilStartTime1 As Integer, ilSortSeq As Integer, tlLLC() As LLC, ilLoop As Integer, llCefCode As Long)
    Dim ilRet As Integer
    Dim ilZone As Integer
    Dim llAirTime As Long
    Dim llTime As Long
    Dim slZone As String
    
    tmOdf.iUrfCode = tgUrf(0).iCode
    If ilODFVefCode <= 0 Then
        If ilSimVefCode <= 0 Then
            If tmVef.iVefCode > 0 Then  'Log vehicle defined
                tmOdf.iVefCode = tmVef.iVefCode
            Else
                tmOdf.iVefCode = ilVehCode
            End If
        Else
            tmOdf.iVefCode = ilSimVefCode
        End If
    Else
        tmOdf.iVefCode = ilODFVefCode
    End If
    If (tmVef.sType = "G") And (tgVpf(ilVpfIndex).sGenLog <> "M") Then
        tmOdf.iGameNo = ilSSFType
    Else
        tmOdf.iGameNo = 0
    End If
    tmOdf.iAirDate(0) = ilLogDate0
    tmOdf.iAirDate(1) = ilLogDate1
    If tlLLC(ilLoop).sXMid = "Y" Then
        smXMid = "Y"
    Else
        smXMid = "N"
    End If

    If tmDlf.iEtfCode = 1 Then
        tmOdf.iAirTime(0) = tmProg.iStartTime(0)
        tmOdf.iAirTime(1) = tmProg.iStartTime(1)
    Else
        tmOdf.iAirTime(0) = ilStartTime0
        tmOdf.iAirTime(1) = ilStartTime1
    End If
    tmOdf.iLocalTime(0) = tmDlf.iLocalTime(0)
    tmOdf.iLocalTime(1) = tmDlf.iLocalTime(1)
    tmOdf.iFeedTime(0) = tmDlf.iFeedTime(0)
    tmOdf.iFeedTime(1) = tmDlf.iFeedTime(1)
    tmOdf.sZone = tmDlf.sZone
    tmOdf.iEtfCode = tmDlf.iEtfCode
    tmOdf.iEnfCode = tmDlf.iEnfCode
    tmOdf.sProgCode = tmDlf.sProgCode
    tmOdf.iMnfFeed = tmDlf.iMnfFeed
    tmOdf.iDPSort = 0   '5-31-01 tmOdf.sUnused1 = "" 'tmDlf.sBus
    'tmOdf.iWkNo = 0    'tmDlf.sSchedule
    tmOdf.ianfCode = 0
    tmOdf.iSortSeq = ilSortSeq                              '8-18-14 This field is required for C88, to keep programs with same back to back times apart.
                                                            'gL14PageSkips & gSetSeqL29 subroutines modify this field for some other logs.  Do not call them for C88.
    If ilType = 2 And tlLLC(ilLoop).sType = "A" Then         '1-10-01
        tmOdf.ianfCode = Val(tlLLC(ilLoop).sName)
    End If
    tmOdf.iUnits = 0
    If tmDlf.iEtfCode = 1 Then
        gPackLength tlLLC(ilLoop).sLength, tmOdf.iLen(0), tmOdf.iLen(1)
    Else
        If (tmDlf.iEtfCode > 13) And (Trim$(tlLLC(ilLoop).sLength) <> "") Then
            If gValidLength(tlLLC(ilLoop).sLength) Then
                gPackLength tlLLC(ilLoop).sLength, tmOdf.iLen(0), tmOdf.iLen(1)
            Else
                tmOdf.iLen(0) = 1
                tmOdf.iLen(1) = 0
            End If
        Else
            tmOdf.iLen(0) = 1
            tmOdf.iLen(1) = 0
        End If
    End If
    tmOdf.lHd1CefCode = 0
    tmOdf.lFt1CefCode = 0
    tmOdf.lFt2CefCode = 0
    tmOdf.iAlternateVefCode = 0 '1-17-14 was tmodfvefnmcefcode
    tmOdf.iAdfCode = 0
    tmOdf.lCifCode = 0
    tmOdf.sProduct = ""
    tmOdf.iMnfSubFeed = tmDlf.iMnfSubFeed
    tmOdf.lCntrNo = 0
    tmOdf.lchfcxfCode = 0
    tmOdf.sDPDesc = ""
    tmOdf.iRdfSortCode = 0
    tmOdf.iBreakNo = 0
    tmOdf.iPositionNo = 0
    tmOdf.iType = ilType
    tmOdf.lCefCode = llCefCode
    tmOdf.lEvtIDCefCode = lmEvtIDCefCode
    tmOdf.sDupeAvailID = tmDlf.sBus
    tmOdf.lAvailcefCode = lmAvailCefCode            'comment from avail placed into spots
    tmOdf.sShortTitle = ""
    tmOdf.imnfSeg = 0               '6-19-01
    tmOdf.sPageEjectFlag = "N"
    'Determine seq number
    '6/4/16: Replaced GoSub
    'GoSub lProcAdjDate
    mProcAdjDate slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType
    '6/4/16: Replaced GoSub
    'GoSub lProcSeqNo
    mProcSeqNo ilSeqNo, slFor, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode
    tmOdf.iSeqNo = ilSeqNo + 1
    tmOdf.iDaySort = imDaySort
    tmOdf.lEvtCefCode = lmEvtCefCode
    tmOdf.iEvtCefSort = imEvtCefSort
    tmOdf.sLogType = smLogType
    tmOdf.sBBDesc = ""
    '10/31/05:  Add zones to program and other events
    'ilRet = btrInsert(hlOdf, tmOdf, imOdfRecLen, INDEXKEY0)
    If ilDlfFound Then
        '8/13/14: Generate spots for last date+1 into lst with Fed as *
        If Not bmCreatingLstDate1 Then
            tmOdf.lCode = 0
            ilRet = btrInsert(hlODF, tmOdf, imOdfRecLen, INDEXKEY3)
            gLogBtrError ilRet, "gBuildODFSpotDay: Insert #1"
        End If
    Else
        ilOtherGen = False
        For ilZone = 0 To UBound(tmZoneInfo) - 1 Step 1
            Select Case tmZoneInfo(ilZone).sZone
                Case "E"
                    slZone = "EST"
                Case "M"
                    slZone = "MST"
                Case "C"
                    slZone = "CST"
                Case "P"
                    slZone = "PST"
                Case Else
                    If ilOtherGen Then
                        Exit For
                    End If
                    slZone = ""
            End Select
            tmOdf.iAirDate(0) = ilLogDate0
            tmOdf.iAirDate(1) = ilLogDate1
            If tmDlf.iEtfCode = 1 Then
                tmOdf.iAirTime(0) = tmProg.iStartTime(0)
                tmOdf.iAirTime(1) = tmProg.iStartTime(1)
                gUnpackTimeLong tmProg.iStartTime(0), tmProg.iStartTime(1), False, llAirTime
            Else
                tmOdf.iAirTime(0) = ilStartTime0
                tmOdf.iAirTime(1) = ilStartTime1
                gUnpackTimeLong ilStartTime0, ilStartTime1, False, llAirTime
            End If
            If (tgSpf.sGUseAffSys = "Y") Then
                llTime = llAirTime + tmZoneInfo(ilZone).lGLocalAdj
                If llTime < 0 Then
                    llTime = llTime + 86400
                ElseIf llTime > 86400 Then
                    llTime = llTime - 86400
                End If
                gPackTimeLong llTime, tmDlf.iLocalTime(0), tmDlf.iLocalTime(1)
                llTime = llAirTime + tmZoneInfo(ilZone).lGFeedAdj
                If llTime < 0 Then
                    llTime = llTime + 86400
                ElseIf llTime > 86400 Then
                    llTime = llTime - 86400
                End If
                gPackTimeLong llTime, tmDlf.iFeedTime(0), tmDlf.iFeedTime(1)
            End If
            tmOdf.iLocalTime(0) = tmDlf.iLocalTime(0)
            tmOdf.iLocalTime(1) = tmDlf.iLocalTime(1)
            tmOdf.iFeedTime(0) = tmDlf.iFeedTime(0)
            tmOdf.iFeedTime(1) = tmDlf.iFeedTime(1)
            tmOdf.sZone = slZone
            '6/4/16: Replaced GoSub
            'GoSub lProcAdjDate
            mProcAdjDate slAdjLocalOrFeed, llDate, llSDate, llEDate, slInLogType
            '6/4/16: Replaced GoSub
            'GoSub lProcSeqNo
            mProcSeqNo ilSeqNo, slFor, hlODF, tlOdf, ilODFVefCode, ilSimVefCode, ilVehCode
            tmOdf.iSeqNo = ilSeqNo + 1
            If (tgSpf.sGUseAffSys <> "Y") Or (imGenODF) Then
                '8/13/14: Generate spots for last date+1 into lst with Fed as *
                If Not bmCreatingLstDate1 Then
                    tmOdf.lCode = 0
                    ilRet = btrInsert(hlODF, tmOdf, imOdfRecLen, INDEXKEY0)
                    gLogBtrError ilRet, "gBuildODFSpotDay: Insert #2"
                End If
            End If
        Next ilZone
    End If
End Sub

Private Function mRemovedBadLST(tlLst As LST) As Boolean
    'Bad = sdfChfCode <> crfChfCode
    Dim llRet As Long
    Dim slSQLQuery As String
    Dim tmp_rst As ADODB.Recordset


    On Error GoTo ErrHand
    mRemovedBadLST = False
    slSQLQuery = "Select lstCode, sdfchfcode, crfchfcode, rsfCode, rsfRChfCode from lst left outer join sdf_spot_detail on lstsdfcode = sdfcode left outer join rsf_Region_Schd_Copy On sdfCode = rsfSdfCode left outer join crf_Copy_Rot_Header on rsfCrfCode = crfCode where lstCode = " & tlLst.lCode & " And sdfChfCode <> crfchfCode And rsfType = 'R'" 'and lstsdfcode = " & tlLst.lSdfCode
    Set tmp_rst = gSQLSelectCall(slSQLQuery)
    Do While Not tmp_rst.EOF
        slSQLQuery = "Delete From rsf_Region_Schd_Copy Where rsfCode = " & tmp_rst!rsfCode
        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
        If llRet = 0 Then
            mRemovedBadLST = True
        End If
        slSQLQuery = "Delete From lst Where lstCode = " & tmp_rst!lstCode
        llRet = gSQLWaitNoMsgBox(slSQLQuery, False)
        If llRet = 0 Then
            mRemovedBadLST = True
        End If
        tmp_rst.MoveNext
    Loop
    tmp_rst.Close
    Exit Function
ErrHand:
    On Error GoTo 0
End Function


'***************************************************************************************
'*
'* Function Name: gLogMsg
'*
'* Created: August, 2004  By: D. Smith
'*
'* Modified:              By:
'*
'* Comments: A general file routine that shows: Date and Time followed by a message
'* Moved from GENSUBS
'***************************************************************************************
Public Sub gLogMsg(sMsg As String, sFileName As String, iKill As Integer)
    'Params
    'sMsg is the string to be written out
    'sFileName is the name of the file to be written to in the Messages directory
    'iKill = True then delete the file first, iKill = False then append to the file

    Dim slFullMsg As String
    Dim hlLogFile As Integer
    Dim ilRet As Integer
    Dim slDateTime As String
    Dim slToFile As String
    
    slToFile = sgDBPath & "Messages\" & sFileName
    'On Error GoTo Error

    If iKill = True Then
        ilRet = 0
        'slDateTime = FileDateTime(slToFile)
        ilRet = gFileExist(slToFile)
        If ilRet = 0 Then
            Kill slToFile
        End If
    End If

    'hlLogFile = FreeFile
    'Open slToFile For Append As hlLogFile
    ilRet = gFileOpen(slToFile, "Append", hlLogFile)
    If sMsg = "" Then
        slFullMsg = "-----------------------------------------------------------------------------"
    Else
        slFullMsg = Format$(Now, "mm/dd/yyyy") & " " & Format$(Now, "hh:mm:ssam/pm") & " - " & sMsg
    End If
    Print #hlLogFile, slFullMsg
    Close hlLogFile

    slFullMsg = UCase(slFullMsg)
    If InStr(1, slFullMsg, "ERROR", vbTextCompare) > 0 Then
        gSaveStackTrace slToFile
    End If

    Exit Sub

'Error:
'    ilRet = 1
'    Resume Next

End Sub

Public Sub gLogMsgWODT(slAction As String, hlFileHandle As Integer, slMsg As String)
    'Moved from GENSUBS
    'Add line to file without adding Date and time like gLogMsg
    
    'Params
    'slAction:  "ON" open as New (kill any previous version); "OA"=Open in append mode (retain previous version); "W"=Write message to file; "C" = Close handle
    'slMsg= If Open, Drive\Path\File name; If Write, message to write to file
    'hlFileHandle: If Open, return value; If Write or Close, handle of file to write to or close
    Dim ilRet As Integer
    Dim slDateTime As String
    
    'On Error GoTo Error
    If slAction = "ON" Then 'Open as New
        ilRet = 0
        'slDateTime = FileDateTime(slMsg)
        ilRet = gFileExist(slMsg)
        If ilRet = 0 Then
            Kill slMsg
        End If
    End If
    Select Case Left$(slAction, 1)
        Case "O"
            'hlFileHandle = FreeFile
            'Open slMsg For Append As hlFileHandle
            ilRet = gFileOpen(slMsg, "Append", hlFileHandle)
        Case "W"
            Print #hlFileHandle, slMsg
        Case "C"
            Close hlFileHandle
    End Select
    Exit Sub
    
'Error:
'    ilRet = 1
'    Resume Next
    
End Sub

