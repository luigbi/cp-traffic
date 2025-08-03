Attribute VB_Name = "RecDefAff"

' Proprietary Software, Do not copy
'
' File Name: RecDefAff.Bas
'
' Release: 1.0
'
' Description:
'   This file contains the Invoice support functions
Option Explicit
Option Compare Text

Public tgUst() As UST               '5-9-11


'******************************************************************************
' aet Record Definition
'
'******************************************************************************
Type AET 'VBC NR
    lCode                 As Long 'VBC NR
    lAtfCode              As Long 'VBC NR
    iShfCode              As Integer 'VBC NR
    iVefCode              As Integer 'VBC NR
    lSdfCode              As Long 'VBC NR
    iFeedDate(0 To 1)     As Integer 'VBC NR
    iFeedTime(0 To 1)     As Integer 'VBC NR
    iPledgeStartDate(0 To 1) As Integer 'VBC NR
    iPledgeEndDate(0 To 1) As Integer 'VBC NR
    iPledgeStartTime(0 To 1) As Integer 'VBC NR
    iPledgeEndTime(0 To 1) As Integer 'VBC NR
    sAdvt                 As String * 30 'VBC NR
    sProd                 As String * 35 'VBC NR
    sCart                 As String * 12 'VBC NR
    sISCI                 As String * 20 'VBC NR
    sCreative             As String * 30 'VBC NR
    lAstCode              As Long 'VBC NR
    iLen                  As Integer 'VBC NR
    lCntrNo               As Long 'VBC NR
    sStatus               As String * 1 'VBC NR
    'sUnused               As String * 10 'VBC NR
    sPledgeDays           As String * 7      ' Pledge days (Mo-Su).  Y or N for
                                             ' each pledge day
    sUnused               As String * 3
End Type 'VBC NR


'Type AETKEY0
'    lCode                 As Long
'End Type

Type AETKEY1 'VBC NR
    lAtfCode              As Long 'VBC NR
    iFeedDate(0 To 1)     As Integer 'VBC NR
End Type 'VBC NR

'Type AETKEY2
'    lAstCode              As Long
'End Type


'******************************************************************************
' ast Record Definition
'
'******************************************************************************
Type AST 'VBC NR
    lCode                 As Long
    lAtfCode              As Long
    iShfCode              As Integer
    iVefCode              As Integer
    lSdfCode              As Long
    lLsfCode              As Long
    iAirDate(0 To 1)      As Integer
    iAirTime(0 To 1)      As Integer
    iStatus               As Integer         ' 0=Live;1=Delay;2-5=Not
                                             ' Carry;7=Aired not Pledged;8=Not
                                             ' Carried:9=Delay;10=cmml
                                             ' only;11=MG;12=Bonus;13=Replacemen
                                             ' t;100=Missed Reason;1000=ISCI
    iCPStatus             As Integer         ' 0=Not Received;1=Received;2=CP
                                             ' Not Aired
    iFeedDate(0 To 1)     As Integer
    iFeedTime(0 To 1)     As Integer
    iAdfCode              As Integer         ' Advertiser reference
    lDatCode              As Long            ' Pledge information reference
    lcpfCode              As Long            ' Product/ISCI reference of spot
                                             ' that was on air. ALT will have
                                             ' what was scheduled to air if
                                             ' different
    lRsfCode              As Long            ' Region spot assigned reference
    sStationCompliant     As String * 1      ' Station Compliant: A=Aired within pledge time;
                                             ' O=Aired outside pledge time; N=Did not
                                             ' air; Blank=Not set
    sAgencyCompliant      As String * 1      ' Agency status: A=Aired as sold;
                                             ' O=Aired outside sold; N=Did not
                                             ' air; Blank=Not set
    sAffidavitSource      As String * 2      ' Affidavit source
                                             ' (hierarchy):1A=Electronically;1B=
                                             ' Third-party;2A=From Auto;2B=To
                                             ' Auto;3A=Station
                                             ' Comfirmed;3B=Station
                                             ' Posted;4A=Schd by Network;4B=Schd
                                             ' by Station
    lCntrNo               As Long            ' Contract Number
    iLen                  As Integer         ' Spot Length
    lLkAstCode            As Long            ' Missed to MG/Replacement Link or
                                             ' MG/Replacement Link to Missed
    iMissedMnfCode        As Integer         ' Missed Reason reference code
    iUstCode              As Integer         ' Affiliate User reference
End Type 'VBC NR


'Type ASTKEY0
'    lCode                 As Long
'End Type

Type ASTKEY1
    lAtfCode              As Long
    iFeedDate(0 To 1)     As Integer
End Type

Type ASTKEY2
    lAtfCode              As Long
    iAirDate(0 To 1)      As Integer
End Type

Type ASTKEY3
    iAdfCode              As Integer
    iAirDate(0 To 1)      As Integer
End Type

'******************************************************************************
' att Record Definition
'
'******************************************************************************
Type ATT
    lCode                 As Long            ' AutoInc
    iShfCode              As Integer         ' Station Code
    iVefCode              As Integer         ' Vehicle code
    iAgreeStart(0 To 1)   As Integer         ' Agreement start date (Blank is
                                             ' 1/1/1901)
    iAgreeEnd(0 To 1)     As Integer         ' Agreement end date (TFN =
                                             ' 1/1/2100)
    iOnAir(0 To 1)        As Integer         ' On Air date (Blank is 1/1/1901)
    iOffAir(0 To 1)       As Integer         ' Off Air date (TFN = 1/1/2100)
    iSigned               As Integer         ' Agreement Signed (0=Signed;
                                             ' 1=Signed; 2=Rejected)
    iSignDate(0 To 1)     As Integer         ' Signed date
    iLoad                 As Integer         ' Load Multiplier
    iTimeType             As Integer         ' 0=Daypart defined; 1=Avails
                                             ' Defined
    iComp                 As Integer         ' 0=Network; 1=Affiliate; 2=Barter
    iStartTime(0 To 1)    As Integer
    iBarCode              As Integer         ' 0=Yes; 1=No
    iDropDate(0 To 1)     As Integer         ' Dropped date
    iCPType               As Integer         ' CP Type
    iUsfCode              As Integer         ' User Code
    iEnterDate(0 To 1)    As Integer         ' Entered date
    iEnterTime(0 To 1)    As Integer         ' Entered Time
    sNotice               As String * 3      ' Cancellation Number of days
                                             ' notice required
    iCarryCmml            As Integer         ' Carry commercials: 0=Yes; 1=No
    iNoCDs                As Integer         ' number of CD sent
    iSendTape             As Integer         ' 0-Yes; 1= No
    sACName               As String * 40
    sACPhone              As String * 20
    sGenLog               As String * 3
    sGenCP                As String * 3
    iPostingType          As Integer         ' 0=CP Return only; CP by Spot
                                             ' Count; Spots by Date; Spots by
                                             ' Advertiser
    iPrintCP              As Integer
    sComments             As String * 120
    sGenOther             As String * 3
    lAgreementID          As Long
    sWklyClear            As String * 1      ' Unused it was defined as 12
                                             ' reduced to 1
    sHrlyClear            As String * 1      ' Unused it was defined as 12
                                             ' reduced to 1
    sHrUsed               As String * 1      ' Unused it was defined as 12
                                             ' reduced to 1
    sUnused               As String * 33     ' Unused
    iExportType           As Integer         ' 0=By Network; 1=By Station
    iLogType              As Integer
    iPostType             As Integer
    sWebPW                As String * 10
    sWebEMail             As String * 240
    iSendLogEMail         As Integer         ' 0=No; 1=Yes: Send EMail that Log
                                             ' is on Web Site
    sSuppressNotice       As String * 1      ' Suppress Notices (Y/N)
    sLabelID              As String * 10     ' Lable ID
    sLabelShipInfo        As String * 40     ' Label Shipping Information
    sMulticast            As String * 1      ' Multicast (Y or N)
    sRadarClearType       As String * 1      ' Radar Clearance Type
                                             ' (C=Commercial Only;
                                             ' P=Program+Commercial).  This
                                             ' field is used with Radar Export
    lArttCode             As Long            ' Reference to the Affiliate A/E
                                             ' that is responsible for this
                                             ' Agreement.
    sStatus               As String * 1      ' C or Blank = Current; H=History
    sNCR                  As String * 1      '7-8-09 agreement non-compliant (y/n)
    sFormerNCR            As String * 1      '7-8-09  agreement became compliant after becoming non-compliant (y/n), then non-compliant again.
                                             'show them as a repeat offender
    sForbidSplitLive      As String * 1      'Prohibit Split Copy in Pledge avails defined with status of Air Live
    'sUnused               As String * 14     '7-8-09 chg from 17 to 15
    lXDReceiverID         As Long            'X-Digital Receiver ID
    sVoiceTracked         As String * 1      'Voice Tracked. Test for Y.  If Y don't export to X-Digital.  Commercial sent not thru X-Digital
    sMonthlyWebPost As String * 1
    sWebInterface         As String * 1      'C=Cumulus
    'sUnused               As String * 5      '7-8-09 chg from 17 to 15
    sContractPrinted      As String * 1      ' Contract Printed. Y or N. Test
                                             ' for Y
    iMktRepUstCode        As Integer         ' MarketRep User Reference code
                                             ' override of station value
    iServRepUstCode       As Integer         ' Service Rep User Reference code
                                             ' override of station value
    iVehProgStartTime(0 To 1) As Integer     ' Vehicle program Start Time
    iVehProgEndTime(0 To 1) As Integer       ' Vehicle program end time
    sExportToWeb          As String * 1      ' Y or N. Test for Y
    sExportToUnivision    As String * 1      ' Y or N. Test for Y
    sExportToMarketron    As String * 1      ' Y or N. Test for Y
    sExportToCBS          As String * 1      ' Y or N. Test for Y
    sExportToClearCh      As String * 1      ' Clear Channel. Y or N. Test for Y
    lContractPFTCode      As Long            ' Contract PDF file table reference
    sPledgeType           As String * 1      ' Pledge Type: A=Avail; D=Dayprt; C=CD Tape
    iNoAirPlays           As Integer         ' Number of air plays
    iDesignVersion        As Integer         ' Agreement design version number, Initial version is 1; Adding management version is 2
    sIDCReceiverID        As String * 5      ' IDC Receiver ID
    iMktronActiveDate(0 To 1) As Integer     ' Marketron Activation date
    sSentToXDSStatus      As String * 1      ' Sent to X-Digital Status (Y, N, M). Test for N and M. Y=Yes, N=No, M=Modified
    sPetPrinted           As String * 1      ' Pledge by Events (Sports Declaration) printed. Y=Yes, N = No; I or Blank = Ignore.  Test for N
    sIDCGroupType         As String * 1      ' IDC Group Type: S=Group by
                                             ' Station; L=Group by Location; N
                                             ' or Blank = None. Test for S or L.
    sExportToJelli        As String * 1      ' Export to Jelli (Y/N). Test for Y
    sAudioDelivery        As String * 1      ' Audio Delivery (X=XDS-ISCI;B=XDS-Break; W=Wegener; I=IDC; P=iPump)
    sServiceAgreement     As String * 1      ' Service type agreement (Limited functionality: XDS Spot insertion Export only). Y=Yes; N=No. Test for Y.
    sSendDelayToXDS       As String * 1      ' Send delay pledges to XDS (Y/N. Test for Y
End Type


'Type ATTKEY0
'    lCode                 As Long
'End Type

Type ATTKEY1 'VBC NR
    iVefCode              As Integer 'VBC NR
End Type 'VBC NR

Type ATTKEY2 'VBC NR
    iShfCode              As Integer 'VBC NR
End Type 'VBC NR

Type ATTKEY3
    iMktRepUstCode        As Integer
End Type

Type ATTKEY4
    iServRepUstCode       As Integer
End Type

Type ATTKEY5
    sContractPrinted      As String * 1
End Type

Type ATTDATEINFO 'VBC NR
    lCode As Long 'VBC NR
    iOnAir(0 To 1)        As Integer         ' On Air date (Blank is 1/1/1901) 'VBC NR
    iOffAir(0 To 1)       As Integer         ' Off Air date (TFN = 1/1/2100) 'VBC NR
    iDropDate(0 To 1)     As Integer         ' Dropped date 'VBC NR
    iExportType           As Integer         ' 0=Manual; 1=Web; 2=Marketron 'VBC NR
End Type 'VBC NR
Public Const ATTDATEEXTPK As String = "LIIIIIII" 'VBC NR
Type ATTDATEARCHIVE 'VBC NR
    lCode As Long 'VBC NR
    lStartDate As Long 'VBC NR
    lEndDate As Long 'VBC NR
    iExportType As Integer 'VBC NR
End Type 'VBC NR

'
'Cert of Perforance Table
'
Type CPTT
    lCode                 As Long
    lAtfCode              As Long            ' Agreement code (att above)
    iShfCode              As Integer         ' Station Code
    iVefCode              As Integer         ' Vehicle Code
    iCreateDate(0 To 1)   As Integer         ' Date record created date
    iStartDate(0 To 1)    As Integer         ' Date on log
    iReturnDate(0 To 1)   As Integer         ' Date CP Reurned (Todays date when
                                             ' posted)
    iStatus               As Integer         ' 0=Not posted
                                             ' (cpttPostingStatus=0) or
                                             ' Partially
                                             ' posed(cpttPostingStatus=1);
                                             ' 1=Posting Completed; 2=Posting
                                             ' Completed as None Aired
    iUsfCode              As Integer         ' User code
    iNoSpotsGen           As Integer         ' Get count when posting cp return
                                             ' (use lst and pledge information)
    iNoSpotsAired         As Integer
    iPostingStatus        As Integer         ' 0=Not posted; 1=Partially posted;
                                             ' 2=Posting Completed
    sAstStatus            As String * 1      ' Status: N=Ast never Created; R=Recreate AST; C= Created
    iNoCompliant          As Integer         ' Number of spots that are pledge
                                             ' compliant
    iAgyCompliant         As Integer         ' number of spots that are agency
                                             ' compliant
    sUnused               As String * 13
End Type
'CPTT key record layout- use LONGKEY0
'Type CPTTKEY0
'    lCode As Integer
'End Type

Type CPTTKEY1
    iVefCode As Integer 'Vehicle Code
    iStartDate(0 To 1) As Integer 'date on log
End Type

Type CPTTKEY2 'VBC NR
    lAtfCode As Long 'Agreement code (att above) 'VBC NR
    iStartDate(0 To 1) As Integer 'date on log 'VBC NR
End Type 'VBC NR

Type CPTTKEY3
    lAtfCode              As Long
End Type

Type CPTTKEY4
    iStartDate(0 To 1)    As Integer
    sAstStatus            As String * 1
End Type

Type CPTTINFO
    lAtfCode As Long 'Agreement code (att above)
    iShfCode As Integer 'Station Code
    iVefCode As Integer 'Vehicle Code
    iReturnDate(0 To 1) As Integer 'Date CP Reurned (Todays date when posted)
    iStatus As Integer  'cpttStatus (0=Not Received or Partially posted; 1= Received(receipt Only) or Received and Completely Posted(Manual or Export); 2=Not Aired)
    'iPrintStatus As Integer     '0=CP Not printed on affiliate system; 1= CP printed; 2=Reprint required
    sAstStatus As String * 1       'N=Never Created; R=Recreate AST; C=Created
    iPostingStatus As Integer   'cpttPostingStatus(0=Not Posted; 1= Partially Posted; 2= Posted)
    lCycleDate As Long
End Type


'******************************************************************************
' cct Record Definition - added 7/24/09
'
'******************************************************************************
Type CCT
    lCode                 As Long
    iShfCode              As Integer
    iVefCode              As Integer
    iActionDate(0 To 1)   As Integer
    iUstCode              As Integer
    iEnteredDate(0 To 1)  As Integer
    iEnteredTime(0 To 1)  As Integer
    'sUnused               As String * 20
    'sComment              As String * 255
    sDone                 As String * 1      ' Comment task completed: Y or N.
                                             ' Test for Y
    iDoneUstCode          As Integer         ' Who set comment as Done User
                                             ' reference code
    iDoneDate(0 To 1)     As Integer         ' Date when set to Done
    iDoneTime(0 To 1)     As Integer         ' Time when set to Done
    iChgdUstCode          As Integer         ' Last person who changed the
                                             ' Comment User Reference code
    iChgdDate(0 To 1)     As Integer         ' Last date comment changed
    iChgdTime(0 To 1)     As Integer         ' Time comment last changed
    iToEMailUstCode       As Integer         ' E-Mail sent as 'To' User
                                             ' reference code.  Used to pass
                                             ' along task to Market or Service
                                             ' Rep
    iCstCode              As Integer         ' Comment Source Reference code
    sUnused               As String * 20
    sComment              As String * 1101    ' Comment. If generated from
                                             ' E-Mail comment includes Sent To
                                             ' and Subject information
End Type


Type CCTKEY0
    lCode                 As Long
End Type

Type CCTKEY1
    iShfCode              As Integer
    iVefCode              As Integer
    iActionDate(0 To 1)   As Integer
End Type


' dat Record Definition
'
'******************************************************************************
Type DAT 'VBC NR
    lCode                 As Long 'VBC NR
    lAtfCode              As Long 'VBC NR
    iShfCode              As Integer 'VBC NR
    iVefCode              As Integer 'VBC NR
    iDACode               As Integer         ' 0=Daypart;1=Avail 'VBC NR
    iFdMon                As Integer         ' 0=N; 1=Y 'VBC NR
    iFdTue                As Integer 'VBC NR
    iFdWed                As Integer 'VBC NR
    iFdThu                As Integer 'VBC NR
    iFdFri                As Integer 'VBC NR
    iFdSat                As Integer 'VBC NR
    iFdSun                As Integer 'VBC NR
    iFdStTime(0 To 1)     As Integer 'VBC NR
    iFdEdTime(0 To 1)     As Integer 'VBC NR
    iFdStatus             As Integer         ' 0=Carried;1=Delay;2-5=Not;7=Specl 'VBC NR
                                             ' ;8=Off air
    iPdMon                As Integer 'VBC NR
    iPdTue                As Integer 'VBC NR
    iPdWed                As Integer 'VBC NR
    iPdThu                As Integer 'VBC NR
    iPdFri                As Integer 'VBC NR
    iPdSat                As Integer 'VBC NR
    iPdSun                As Integer 'VBC NR
    iPdStTime(0 To 1)     As Integer 'VBC NR
    iPdEdTime(0 To 1)     As Integer 'VBC NR
    sPdDayFed             As String * 1      ' Clarify which days spots is 'VBC NR
                                             ' airing on when Pledge day prior
                                             ' to Feed day.  B=Before; A=After
                                             ' (default).
    iAirPlayNo            As Integer         ' Air play number
    sNominalTime          As String * 1      ' Nominal Time allowed to be defined (Y or N). Test for Y
    sEmbeddedOrROS        As String * 1      ' Embedded spots or ROS spots (E/R).  Test for E. Blank is the same as R.
    sUnused               As String * 15     ' Unused
End Type 'VBC NR


'Type DATKEY0
'    lCode                 As Long
'End Type

'Type DATKEY1
'    lAtfCode              As Long
'End Type

Type DATKEY2 'VBC NR
    iShfCode              As Integer 'VBC NR
    iVefCode              As Integer 'VBC NR
End Type 'VBC NR


'
'   Log Spot
'
Type LST
    lCode As Long
    iType As Integer    '0=Spot; 1=Avail; 2= Internal added (MG; Bonus or Replacement)
    lSdfCode As Long    'SdfCode
    lCntrNo As Long     'Contract Number
    iAdfCode As Integer 'AdfCode
    iAgfCode As Integer 'AgyCode
    sProd As String * 35    'Product Name
    iLineNo As Integer  'Line number
    iLnVefCode As Integer   'Line Vehicle vefcode
    iStartDate(0 To 1) As Integer   'Flight Start Date
    iEndDate(0 To 1) As Integer     'Flight End Date
    iDays(0 To 6) As Integer    'Index 0:Monday, 1=Tuesday,...6=Sunday; 0=No; 1=Yes
    iSpotsWk As Integer     'Spots per week
    iPriceType As Integer   'Price type (0=True; 1=Bonus,...)
    lPrice As Long          'Price (xxxxx.xx)
    iSpotType As Integer    'Spot Type (0=Schd; 1=MG; 2=Filled; 3=Outside; 4=  ;5=Added; 6=Open BB; 7=Close BB)
    iLogVefCode As Integer  'Log (airing) vehicle code
    iLogDate(0 To 1) As Integer 'Log date
    iLogTime(0 To 1) As Integer 'Log Time
    sDemo As String * 6     'Demo Name
    lAud As Long    'Audience for primary demo
    sISCI As String * 20    'ISCI code from copy
    'sUnused As String * 30
    iWkNo As Integer
    iBreakNo As Integer
    iPositionNo As Integer
    iSeqNo As Integer
    sZone As String * 3
    sCart As String * 7
    lcpfCode As Long
    lCrfCsfcode As Long
    iStatus As Integer      '0=Carried; 1=Delay; 2-5=Not Carried; 7=Special; 8=Off Air
    iLen As Integer         'Spot length or unsold avail length
    iUnits As Integer       'Avail Units (Unsold)
    lCifCode As Long        'CifCode (required by odf)
    ianfCode As Integer     'Avail name (required by odf to determine if avail starts with N)
    lEvtIDCefCode As Long   'Event ID
    '6/29/06
    sSplitNetwork         As String * 1      ' Split Network Flag: N=Not split
                                             ' network spot; P=Primary Split
                                             ' Network spot; S=Secondary Split
                                             ' Network Spot
    lRafCode              As Long            ' Split network Region (Rafcode)
                                             ' reference or zero if not a split
                                             ' network spot
    lFsfCode As Long    'Feed Spot (0 if contract spot)
    lGsfCode As Long    'Game schedule (0 if not game spot)
    sImportedSpot As String * 1     'M=MYL spot import; F=Feed Spot import (fsfCode will be zero and Contract number will be blank)
    lBkoutLstCode As Long   'Used to reference the lst blackout becuase of copy split by this spot.
                            'This record create so that ast will show the create advertiser name.
    iLnStartTime(0 To 1) As Integer   'Line Start time.  Override time if defined, otherwise daypart time
    iLnEndTime(0 To 1) As Integer     'Line End time
    sUnused               As String * 20     ' Unused
End Type

Type LSTKEY2
    iLogVefCode As Integer  'Log (airing) vehicle code
    iLogDate(0 To 1) As Integer 'Log date
End Type

Type LSTKEY3
    iLogVefCode           As Integer
    lGsfCode              As Long
End Type
'
'   Station info
'
Type SHTT
    iCode As Integer    'Auto Inc
    sCallLetters As String * 40 'Call letters or Persons Name
    sAddress(0 To 1) As String * 40
    sCity As String * 40    'City
    sState As String * 40    'State
    sCountry As String * 40 'Country
    sZip As String * 20     'Zip
    iSelected As Integer
    sEMail As String * 70   'E-Mail Address
    sFax As String * 20     'Fax
    sPhone As String * 20   'Phone number
    sTimeZone As String * 3
    sHomePage As String * 50
    'sPDName As String * 80  'Not Used: Program Director
    sPDName               As String * 69     ' Not Used,  moved to artt: Program
                                             ' Director (changed size from 80 to
                                             ' 54 to add iPump fields)
    sPDPhone              As String * 20
    sUsedForIPump         As String * 1      ' Used for iPump (Y/N)
    sIPumpID              As String * 10     ' iPump Station ID
    sTDName As String * 62  'Not Used: Traffic Director
    sTDPhone As String * 20
    'sMDName As String * 80  'Not Used: Music Director
    ''sMDPhone As String * 20
    '''iPC As Integer          'Has PC- not used
    '''iHdDrive As Integer     'Hard Drive- Not Used
    ''sMonthlyWebPost As String * 1
    ''sUnused As String * 3
    lOnCityMntCode        As Long            ' On City Mnt reference code
    sOnCountry            As String * 40     ' Country Name
    lCityMntCode          As Long            ' Mailing City mnt reference code
    lCityLicMntCode       As Long            ' City License mnt reference code
    lCountyLicMntCode     As Long            ' County License mnt reference code
    sMDName As String * 1
    sAgreementExist       As String * 1      ' Agreement exist. Y or N. Test for
                                             ' Y
    sCommentExist         As String * 1      ' Comment Exist. Y or N. Test for Y
    iMktRepUstCode        As Integer         ' Market Rep User Reference Code
    iServRepUstCode       As Integer         ' Service Rep User Reference Code
    lOperatorMntCode      As Long            ' Operator of Station Multi-Name
                                             ' Reference code
    lAreaMntCode          As Long            ' Area Multi-Name Reference code
    iHistStartDate(0 To 1) As Integer        ' Historical start date of
                                             ' agreements.
    lAudP12Plus           As Long            ' Audience value of persons 12+
    lMonikerMntCode       As Long            ' Moniker Multi-Name reference
                                             ' code
    lMultiCastGroupID     As Long            ' Group ID assigned to Multi-Cast
                                             ' stations
    lClusterGroupID       As Long            ' Sister station Group ID assigned
                                             ' to stations
    sMasterCluster        As String * 1      ' Master cluster station (Y or N)
    sVieroID              As String * 5      ' Viero Transact Enterprise ID
    sFrequency As String * 6    'Station Frequency (AM: xxxx; FM: xxx.xx)
    lPermStationID As Long  'Permanent station ID
    'sUnused As String * 12
    sOnAir                As String * 1      ' Station On Air. Y or N.  Test for
                                             ' N
    sStationType          As String * 1      ' C=Commercial station;
                                             ' N=Non-Commercial station. Test
                                             ' for N(Assume Commercial Station).
    lWatts As Long  'Watts
    sPledgeVsAir As String * 1  'Include station in the Pledge vs Air CSV Affiliate export (Y/N). Test for Y
    sSentToXDSStatus      As String * 1      ' Sent to X-Digital Status (Y, N, M). Test for N and M. Y=Yes, N=No, M=Modified
    sUnused               As String * 1      ' Unused
    iSpotsPerWebPage      As Integer         ' Number of spots per web page.
    sACName As String * 80  'Not Used: Affidavit Contact
    sACPhone As String * 20
    lMntCode As Long    'Territory Reference (mntType = "T")
    iChecked As Integer     'Affiliate contact(0=PD; 1=MD and 2=TD)
    sMarket As String * 60  'Market
    iRank As Integer        'Rank
    iUsfCode As Integer     'User code
    iEnterDate(0 To 1) As Integer 'Entered date
    iEnterTime(0 To 1) As Integer 'Entered Time
    iType As Integer        'Type: 0=Station; 1=Person
    sONAddress(0 To 1) As String * 40   'Over night address
    sONCity As String * 40    'City     'Over night city
    sONState As String * 40    'State    'Over night State
    sONZip As String * 20     'Zip      'Over night zip
    lStationID As Long      'Station ID
    sCityLic As String * 40 'City of License
    sStateLic As String * 2 'State of Licenece
    iAckDaylight As Integer 'Acknowledge Daylight (0=Yes; 1=No)
    sWebEMail As String * 240
    sWebPW As String * 10
    lOwnerArttCode As Long    ' Reference to Owner stored into the ARTT file
    iMktCode As Integer       ' DMA (Designated market area) reference to mkt.mkd
    'sACFax As String * 20     ' Not Used: Affidavit Contact Fax #
    'sACEMail As String * 70   ' Not Used: Affidavit Contact E-Maill address
    sWebAddress As String * 90  'Station Web Address (URL)
    iFmtCode As Integer       'Format
    sSerialNo1 As String * 10 'Serial number 1 (For  starguide, this is Primary EDAS.  For wegener, this is the reciever serail number)
    sSerialNo2 As String * 10  'Serial number 2 (For  starguide, this is Backup EDAS)
    iTztCode As Integer 'Time zone reference
    sWebNumber            As String * 1      ' Electronic Affidavit number (1=Prior to Web MG system; 2=Web MG system). Test for 2
    sUsedForAtt           As String * 1      ' Used with Agreements (Y/N).  Test for N
    sUsedForXDigital      As String * 1      ' Used for X-Digital export (Y/N). Test for Y
    sUsedForWegener       As String * 1      ' Used for Wegener export (Y/N). Test for Y
    sUsedForOLA           As String * 1      ' Used for OLA export (Y/N).  Test for Y
    sPost As String * 1
    iMetCode As Integer 'MSA(Metropolitan statistical area) reference to met.mkd
    'sUnused As String * 2
End Type
'SHTT key record layout- use INTKEY0
'Type `KEY0
'    iCode As Integer
'End Type
Type SHTTINFO
    iCode As Integer    'Auto Inc
    sCallLetters As String * 40 'Call letters or Persons Name
    sState As String * 40    'State
    sTimeZone As String * 3
    sAgreementExist       As String * 1      ' Agreement exist. Y or N. Test for
    lPermStationID As Long  'Permanent station ID
    iMktCode As Integer       ' Reference to Market Table
    iFmtCode As Integer       'Format
    iTztCode As Integer 'Time zone reference
    iMetCode As Integer 'Metro Market
    iShttVefCode As Integer
    lAudP12Plus  As Long            '12-17-18
    lMultiCastGroupID As Long
End Type

'******************************************************************************
' ust Record Definition
'
'******************************************************************************
Type UST
    iCode                 As Integer
    sName                 As String * 20
    sReportName           As String * 20
    sPassword             As String * 20
    iState                As Integer
    iChgPast              As Integer
    iPassDate(0 To 1)     As Integer
    sChangeCSIDate As String * 1
    sActivityLog          As String * 1      ' Activity Log: V=View; H=Hide
    sUnused2               As String * 6      ' Unused
    iSnfCode              As Integer
    sWin1(0 To 9)          As String * 1      ' Station (I=Input allowed; V=View
                                             ' Only;H=Hide)
    sPedge                As String * 1
    sExptSpotAlert        As String * 1
    sExptISCIAlert        As String * 1
    sTrafLogAlert         As String * 1
    sWin2(0 To 4)         As String * 1      ' E-Mail(I=Input allowed; V=View
                                             ' Only;H=Hide)
    lEMailCefCode         As Long
    sPhoneNo              As String * 25     ' Phone number
    sCity                 As String * 50     ' City
    sAllowedToBlock       As String * 1      ' Allowed to block system
    'sUnused               As String * 32
    sWin16                As String * 1      ' Affiliate CRM permission (I=Input
                                             ' allowed; V=View Only; H=Hide).
    sUserInitials         As String * 3      ' Users initials
    iDntCode              As Integer         ' Department name reference code
    sAllowCmmtChg         As String * 1      ' Allowed to change comments. Y or
                                             ' N. Test for Y
    sAllowCmmtDelete      As String * 1      ' Allowed to delete comments. Y or
                                             ' N. Test for Y
    sWin17                As String * 1      ' Export (I=Input
                                             ' allowed; V=View Only; H=Hide).
    sChgExptPriority      As String * 1      ' Allowed to change Export priority
                                             ' Y=Yes; N=No. Test for
                                             ' Y
    sExptSpec             As String * 1      ' Allwed to enter/update export
                                             ' specifications (Y/N). Test for Y
    sChgRptPriority       As String * 1      ' Allowed to change Report Queue
                                             ' priority. Y=Yes; N=No. Test for
                                             ' Y
    iSSMnfCode            As Integer         ' Affiliate spot Sales Source
                                             ' MnfCode reference filter
    sAdHocType            As String * 1      ' Ad Hoc Type: C=Create and Run;
                                             ' R=Run Only; N or Blank=Not
    sVendorAlert          As String * 1      ' Allow user to view Vendor Alert
                                             ' (Y or blank/N). Test for N
    sUnused               As String * 16
End Type


'Type USTKEY0
'    iCode                 As Integer
'End Type

'******************************************************************************
' site Record Definition
'
'******************************************************************************
Type SITE
    lCode                 As Long
    iMarketron            As Integer
    iWeb                  As Integer
    iNoDaysDelq           As Integer
    'iDaysResendISCI       As Integer
    iUnused3              As Integer
    iCMCmtCode            As Integer
    iWMCmtCode            As Integer
    iPMCmtCode            As Integer
    iOMCmtCode            As Integer
    iOMNoWeeks            As Integer
    iDaysRetainSpots      As Integer
    lAdminArttCode        As Long
    sVehicleStn           As String * 1
    'sISCIExport           As String * 1
    ''iNoMnthRetain         As Integer
    ''iLastDateArch(0 To 1) As Integer
    'sUnused2              As String * 6
    sUsingServAgree       As String * 1      ' Using Service Agreements (Y/N).
                                             ' Test for Y
    sSyncMulticast        As String * 1      ' Sync multicast stations. Part 1:
                                             ' Child Copy = Parent Copy; Web
                                             ' post to all parent/child (Y/N or
                                             ' blank. Test for Y)
    iBandCmtCode          As Integer         ' Allowed import bands cmt internal
                                             ' reference (bands strored
                                             ' separated by commas)
    sUnused2              As String * 1
    sShowVehType          As String * 1
    sExportCart4          As String * 1      ' Use 4 character Cart Numbers
    sExportCart5          As String * 1      ' Use 4 character Cart Numbers
    sCCEMail              As String * 70
    iDayRetainPost        As Integer         ' Number of days to retain web
                                             ' posted spots
    sChngPswd             As String * 1      ' Flag to allow web users to change
                                             ' passwords or not - Y/N
    iOMMinDate(0 To 1)    As Integer         ' indicates oldest date that will
                                             ' be used when sending out Overdue
                                             ' CP emails
    sEmailHost            As String * 80     ' Host name/SMTP server address or
                                             ' URL
    iEmailPort            As Integer         ' Port Number for SMTP server
    sEmailAcctName        As String * 80     ' Account name for SMTP
                                             ' credentials.
    sEmailPassword        As String * 80     ' Password for the SMTP credentials
    sEmailFromAddress     As String * 80     ' The email address of the sender.
    sEmailFromName        As String * 80     ' Used as from email name.  Like
                                             ' "Dial Global Netowrk" or "ABC
                                             ' Clearance Dept."
    iNCRWks               As Integer        '7-8-09 # weeks behind considered to be non-compliant
    'sUnused               As String * 4     '7-8-09 chg from 6 to 4
    sWebSuppressLog As String * 1
    sMultiVehWebPost As String * 1
    'sUnused As String * 2
    iWebNoDyKeepMiss      As Integer         ' Numner of days to retain missed
                                             ' spots on the web.
    sAllowBonusSpots      As String * 1      ' Allow Bonus spots on the Web
                                             ' Site. Y or N.
    sDDF092710            As String * 1      ' Conversion program run for DDF
                                             ' vhanges on 9/26/10.  Y or N or I. Test
                                             ' for Y. I= In progress
    sShowContrDate        As String * 1      ' Show agreement start and end date. Y or N. Test for Y
    sUsingViero           As String * 1      ' Using Viero Transact Enterprise.
                                             ' Y or N. Test for Y.
    sUsingStationID       As String * 1      ' Using station ID
    sStationToXDS         As String * 1      ' Send Station information to
                                             ' X-Digital (Y/N).  Test for Y
    sAgreementToXDS       As String * 1      ' Send Agreement information to
                                             ' X-Digital (Y/N).  Test for Y
    sMissedDateTime       As String * 1      ' Show missed date/time on reports
                                             ' (Y/N). Test for Y
    sCompliantBy          As String * 1      ' Compliant by: A=Advertiser;
                                             ' P=Pledge.  Test for A.  Used by
                                             ' Marketron export only
    iWebNoDyViewPost      As Integer         ' Number of days to View Posted
                                             ' spots on the Web
    iRqtDate(0 To 1)      As Integer         ' Active date of Report Queue.
                                             ' Used to determine if Report Queue
                                             ' is running.  It is update every x
                                             ' seconds
    iRqtTime(0 To 1)      As Integer         ' Active timesiteRqtTime of Report
                                             ' Queue.  Used to determine if
                                             ' Report Queue is running.  It is
                                             ' update every x seconds
    sGenTransparent       As String * 1      'Generate Transparent file as part of XDS spot insertion (Y/N). Test for Y
    sProgToXDS            As String * 1      ' Export Program (Vehicle) Information to XDS (Y/N). Test for Y
    iSSBDate(0 To 1)      As Integer         ' Active date of Station Spot
                                             ' Builder
    iSSBTime(0 To 1)      As Integer         ' Active time of Station Spot
                                             ' Builder
    sSupportXDSDelay      As String * 1      ' Support XDS Delay (playback)
                                             ' (Y/N). Test for Y
    sAllowAutoPost        As String * 1      ' Station not required to approval
                                             ' Auto Posting prior to
                                             ' transmitting back to Affiliate
                                             ' (Y/N). Test for N
    sWithinMissMonth      As String * 1      ' Within missed standard month
                                             ' (Y=Yes; N or Blank=No). Test for
                                             ' Y. If Y, than siteMGDays will be
                                             ' -1
    sLastWk1stWk          As String * 1      ' Miised in last week of the
                                             ' standard month, MG allowed in 1st
                                             ' week of next month (Y=Yes; N or
                                             ' Blank = No). Test for Y
    sSkipHiatusWk         As String * 1      ' Disallow MG in skipped Weeks
                                             ' (Hiatus). (Y=Yes; N or Blank=No).
                                             ' Test for Y
    sValidDaysOnly        As String * 1      ' Book only into valid flight days.
                                             ' (Y=Yes; N or Blank=No). Test for
                                             ' Y
    sTimeRange            As String * 1      ' MG with time range (O=Flight
                                             ' Order Times; P=Pledge Times; N or
                                             ' Blank=Any Time).Test for O or P
    sISCIPolicy           As String * 1      ' ISCI Policy: O or Blank=Order;
                                             ' A=Advertiser after Order failed.
                                             ' Test for A
    iMGDays               As Integer         ' Number of days that the MG must
                                             ' be issued by. If this is defined,
                                             ' than siteWithinMissMonth must be
                                             ' N for no.
    iCompetSepTime        As Integer         ' Competitive separation time in
                                             ' minutes displayed on the Web
    sAllowMGSpots         As String * 1      ' Allow MG spots on the Web (Y/N).
                                             ' Test for N.  If yes, than check
                                             ' at the vehicle level if MG's
                                             ' allowed
    sAllowReplSpots       As String * 1      ' Allow replacement spots on the
                                             ' web (Y/N). Test for N.  Check at
                                             ' the vehicle level to determine if
                                             ' the vehicle is allow replacements
    sNoMissedReason       As String * 1      ' No Missed reason required on the
                                             ' Web (Y/N).  If Yes bypass all
                                             ' missed reason fields on the web
    sDefaultEstDay        As String * 1      ' Default Estimate Day (Y/N). Test
                                             ' for Y
    sWebPostInFuture      As String * 1      ' Allow Web posing in Today and Future (N=No; Y or Blank = Yes. Test for N)
    sMissedMGBypass       As String * 1      ' Missed MG Bypassed Approval (Y=Yes; N or Blank = No). Test for Y
    'sUnused               As String * 1
    sRADARMultiAir        As String * 1      ' RADAR Export: Blank
                                             ' or S = Multi-Airplay by Spot ID; P=Multi-Airplay by RADAR Program Code. Test for P; A=by Air time. Test for A
    lAstMaxLastValue      As Long            ' astCode max value last obtained
                                             ' for date defined in
                                             ' siteAstMaxDate
    iAstMaxDate(0 To 1)   As Integer         ' Date last obtained max astCode
                                             ' value
    iUMCmtCode            As Integer         ' Unresolved Missed Message stored
                                             ' in Cmt table
    sUnused               As String * 18     ' Unused
  '  sUnused               As String * 20     ' Unused
End Type


'Type SITEKEY0
'    lCode                 As Long
'End Type


'******************************************************************************
' RBT Record Definition
'
'******************************************************************************
Type RBT
    lCode                 As Long            ' Auto Increment.  Blackout
                                             ' advertiser reference
    lAstCode              As Long            ' Affiliate Spot reference
    iAdfCode              As Integer         ' Replacement advertiser reference
    iUstCode              As Integer         ' Affiliate user reference code
    sUnused               As String * 20     ' Unused
End Type


'Type RBTKEY0
'    lCode                 As Long
'End Type

Type RBTKEY1 'VBC NR
    lAstCode              As Long 'VBC NR
End Type 'VBC NR

'******************************************************************************
' EDF_Export_Detail Record Definition
'
'******************************************************************************
Type EDF
    lCode                 As Long
    lEsfCode              As Long            ' Link to ESF - Export Summary
    lAttCode              As Long            ' Link to Att - Agreement File
    iVefCode              As Integer         ' Link to Vef - Vehicle Code
    iShttCode             As Integer         ' Link to Shtt - Station File
    iStartTime(0 To 1)    As Integer         ' Time the spots were created for
                                             ' export
    iStartDate(0 To 1)    As Integer         ' Date the spots were created fort
                                             ' export
    lTtlAdd               As Long            ' Total number of Add records for
                                             ' that station/vehicle combination
    lTtlDel               As Long            ' Total number of Delete records
                                             ' for that station/vehicle
                                             ' combination
    lTtlAddDel            As Long            ' Total number of Add and Delete
                                             ' records for that station/vehicle
                                             ' combination
    sUser                 As String * 10     ' The users name that they logged
                                             ' in with
    sAlert                As String * 40     ' If no spots were exported it may
                                             ' be that the weelk already had
                                             ' spots posted against it.
    sUnused               As String * 40     ' Extra space
End Type


'Type EDFKEY0
'    lCode                 As Long
'End Type

'Type EDFKEY1
'    lEsfCode              As Long
'End Type

'Type EDFKEY2
'    lattCode              As Long
'End Type


'******************************************************************************
' ESF_Export_Summary Record Definition
'
'******************************************************************************
Type ESF
    lCode                 As Long
    lTtlAdd               As Long            ' Total number of add records for
                                             ' the entire export
    lTtlDel               As Long            ' Total number of delete records
                                             ' for the entire export
    lTtlAddDel            As Long            ' Total number of add and delete
                                             ' records for the entire export
    lTtlEmails            As Long            ' Total number of emails sent for
                                             ' the entire export
    lTtlHdrs              As Long            ' Total number of headers sent out
                                             ' the entire export
    lTtlComments          As Long            ' Total number of spot comments
                                             ' sent out
    iStartTime(0 To 1)    As Integer         ' The time the export was started
    iStartDate(0 To 1)    As Integer         ' The date the export was started
    iEndTime(0 To 1)      As Integer         ' The time the export ended
    iEndDate(0 To 1)      As Integer         ' The date the export ended
    sElspdTime            As String * 8      ' The total elapsed time the export
                                             ' took start to finish
    iExpDate(0 To 1)      As Integer         ' The date the user entered for the
                                             ' start date of the export.
    iNumDays              As Integer         ' The number of days the user
                                             ' entered to export.
    sUser                 As String * 10     ' The name of the user as they are
                                             ' logged in as
    sMachine              As String * 10     ' The name of the machine that was
                                             ' used to do the export
    sFileName             As String * 80     ' The name of the FIRST file in the
                                             ' export.  An export can send many
                                             ' files all with different names
    sErrors               As String * 80     ' If an error occurs it gets logged
                                             ' here
    sExpSuccess           As String * 1      ' Export was successful or not.  Y
                                             ' = successful, N = not successful
    sUnused               As String * 40     ' Extra space
End Type


'Type ESFKEY0
'    lCode                 As Long
'End Type

Type ESFKEY1
    iExpDate(0 To 1)      As Integer
End Type


'******************************************************************************
' webl Record Definition
'
'******************************************************************************
Type WEBL
    lCode                 As Long
    iType                 As Integer
    lAttCode              As Long
    sCallLetters          As String * 40
    sVehicleName          As String * 40
    sUserName             As String * 40
    iPostDay(0 To 1)      As Integer
    iDate(0 To 1)         As Integer
    iTime(0 To 1)         As Integer
    sIP                   As String * 40
    sCPUName              As String * 40
End Type


'Type WEBLKEY0
'    lCode                 As Long
'End Type

Type WEBLKEY1
    lAttCode              As Long
    iPostDay(0 To 1)      As Integer
End Type

Type WEBLKEY2
    iPostDay(0 To 1)      As Integer
End Type

