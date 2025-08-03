Attribute VB_Name = "modAffiliate"
'******************************************************
'*  modAffiliate - various global declarations
'*
'*  Copyright Counterpoint Software, Inc.
'******************************************************
Option Explicit

'11/16/17
Public bgTaskBlocked As Boolean 'True=Unable to create ast in gGetAstInfo. Set to false prior to calling gGetAstInfo
Public sgTaskBlockedName As String     'Which task called gGetAstInfo. Set prior to calling gGetAstInfo
Public sgTaskBlockedDate As String  'File extension date
Public sgCopyComment As String
' Dan M 4/14/09 changes to guide

Public bgLimitedGuide As Boolean
Public bgStrongPassword As Boolean
Public igExitAff As Integer
Public sgPassUserName As String
'Public Const STRONGPASSWORD = 16
Public bgNoRADAR As Boolean
Public igChangesAllowed As Integer
Public sgPasswordPasser As String   'to communicate with affnewpw
Public bgShowCurrentPassword As Boolean 'set certain controls visible affnewpw
'Current web site information
Public sgWebSiteVersion As String
Public sgWebSiteDate As String
Public sgWebSiteNeedsUpdating As String
'Web site version the Affiliate expects to see
Public sgWebSiteExpectedByAffiliate As String
Public bgMarketRepDefinedByStation As Boolean
Public bgServiceRepDefinedByStation As Boolean
Public igGGFlag As Integer
Public igRptGGFlag As Integer   '0=Disallow Report; 1=Allow reports. test only when igGGFlag = 0
'4/21/18: SQL Trace
Public sgSQLTrace As String
Public hgSQLTrace As Integer

Public sgUserNameToPassToTraffic As String
Public sgUserPasswordToPassToTraffic As String

'added to pass parameter to report
Public sgDate As String
'added to pass parameter to report
Public ifFlag As Boolean
Public imChkListBoxIgnore As Integer
Public sgStdDate As String
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function timeGetTime Lib "winmm.dll" () As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function LockWindowUpdate Lib "user32" (ByVal hKey As Long) As Long
Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hwnd%, ByVal wMsg%, ByVal wParam%, lParam As Any)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
Declare Function GetFocus Lib "user32" () As Long
Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
' FTP Operations
Declare Function csiFTPInit Lib "CSI_Utils.dll" (ByRef FTPInfo As CSIFTPINFO) As Integer
Declare Function csiFTPFileToServer Lib "CSI_Utils.dll" (ByVal slFileName$) As Integer
Declare Function csiFTPFileFromServer Lib "CSI_Utils.dll" (ByVal slFileName$) As Integer
Declare Function csiFTPGetStatus Lib "CSI_Utils.dll" (ByRef FTPStatus As CSIFTPSTATUS) As Integer
Declare Function csiFTPGetError Lib "CSI_Utils.dll" (ByRef FTPErrorInfo As CSIFTPERRORINFO) As Integer
' End FTP Operation
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const STILL_ACTIVE = &H103&  'Result of GetExitCodeProcess
Public Const PROCESS_QUERY_INFORMATION = &H400  'Used by OpenProcess
Global Const LB_SETHORIZONTALEXTENT = &H194 'Added for Email-Messages feature
Public sgFileAttachment As String
Public sgFileAttachmentName As String
Public Const LB_FINDSTRING = &H18F
Public Const LB_SELITEMRANGE = &H19B
Public Const CB_FINDSTRING = &H14C
Public Const LB_GETITEMHEIGHT = &H1A1
Public Const LB_GETHORIZONTALEXTENT = &H193
Public Const INTERNET_FLAG_PASSIVE = &H8000000
'List View
Public Const LV_GRIDLINES = 1
Global Const LV_SETEXTENDEDLISTVIEWSTYLE = 4150
Global Const LV_FULLROWSSELECT = 32
Public Const LIGHTYELLOW = &HC0FFFF '&HBFFFFF '&H80FFFF '&HBFFFFF
Public DARKGREEN As Long ' = RGB(0, 128, 0)      'rgb(0,128,0)
Public Const LIGHTGREEN = &H80FF80
Public Const LIGHTBLUE = &HFDFFD7
Public Const GRAY = &HC0C0C0
Public Const LIGHTGRAY = &HDCDCDC
Public Const ORANGE = 33023 '&H80FF
Public Const GREEN = 49152
Public Const BROWN = 128
Public igShowMsgBox As Integer
Public sgLogoPath As String
Public sgEMailGenericTitle As String

' Global indicator whether using the Novelty web system
Public gIsUsingNovelty As Boolean

'5676
Public sgRootDrive As String
'Const G_MAX_ARRAYDIMS = 60      'VB limit on array dimensions
'Changed to 59 because the array is now 0 based instead on 1
Const G_MAX_ARRAYDIMS = 59      'VB limit on array dimensions
Public Const AVAIL_OR_DP_TIME = 300    'Length of time to consider a break as daypart instead of avail
                                '300 or less equal avail; greater then 300 equals daypart
'7967
Public dgWvImportLast As Date
Public igWVImportElapsed As Integer
Type TASKINFO
    sTaskCode As String * 3
    sTaskName As String * 30
    sSortCode As String * 1
    iMenuIndex As Integer
    lRunningDate As Long
    lRunningTime As Long
    lElapsedTime As Long
    lColor As Long
End Type
Public hgTmf As Integer
'7967 on/off here  change to 15/16
'D.S. 03/21/18
Public tgTaskInfo(0 To 20) As TASKINFO      '7-9-15 add tableau '6/21/16 7967-WVI '12/15/20 TTP9992
Public lgDateMonitorChecked As Long
Public lgTimeMonitorChecked As Long


' FTP Operations
Type CSIFTPINFO
   nPort As Integer
   sIPAddress As String * 64
   sUID As String * 40
   sPWD As String * 40
   sSendFolder As String * 128
   sRecvFolder As String * 128
   sServerDstFolder As String * 128
   sServerSrcFolder As String * 128
   sLogPathName As String * 128
End Type
 
Type CSIFTPSTATUS
   iState As Integer    ' 0=Complete, 1=Busy.
   iStatus As Integer   ' 0=Success, 1=Errors occured
   iJobCount As Integer ' The # of files yet to process.
   lLastError As Long   ' Contains the results of GetLastError if an error occurs.
End Type
 
Type CSIFTPERRORINFO
    sInfo As String * 1024
    sFileThatFailed As String * 128
End Type
 
' End FTP Operation - See delcares above

' FTP Operations

Public igManUnload As Integer   'Force unload even if file (records) not saved

Type CSIFTPFILELISTING
   nPort As Integer
   sIPAddress As String * 64
   sUID As String * 40
   sPWD As String * 40
   sPathFileMask As String * 128
   sSavePathFileName As String * 128
   sLogPathName As String * 128
   nTotalFiles As Integer
End Type

Public tgCsiFtpFileListing As CSIFTPFILELISTING


Type tagAV                      'Array and Vector in 1 compact unit
    PPSA As Long                'Address of pointer to SAFEARRAY
    NumDims As Long         'Number of dimensions
    sCode As Long               'Error info
    Flags As Long               'Reserved
    'Subscripts(1 To G_MAX_ARRAYDIMS) As Long        'rgIndices Vector
    Subscripts(0 To G_MAX_ARRAYDIMS) As Long        'rgIndices Vector
End Type
Declare Function fnAV Lib "QPRO32.DLL" (ByRef A() As Any, ParamArray SubscriptsVector()) As tagAV
Public Declare Sub ArraySortTyp Lib "QPRO32.DLL" (ByRef AV As tagAV, ByVal NumEls As Long, ByVal bDirection As Long, ByVal ElSize As Long, ByVal MbrOff As Long, ByVal MbrSiz As Long, ByVal CaseSensitive As Long)


Public sgCommand As String
Public sgNowDate As String  'If blank, then use Now date, otherwise use this date

'Just some temporary items used for debugging and testing or whatever you like
Public lgSTime1 As Long
Public lgETime1 As Long
Public lgTtlTime1 As Long

Public lgSTime2 As Long
Public lgETime2 As Long
Public lgTtlTime2 As Long

Public lgSTime3 As Long
Public lgETime3 As Long
Public lgTtlTime3 As Long

Public lgSTime4 As Long
Public lgETime4 As Long
Public lgTtlTime4 As Long

Public lgSTime5 As Long
Public lgETime5 As Long
Public lgTtlTime5 As Long

Public lgSTime6 As Long
Public lgETime6 As Long
Public lgTtlTime6 As Long

Public lgSTime7 As Long
Public lgETime7 As Long
Public lgTtlTime7 As Long

Public lgSTime8 As Long
Public lgETime8 As Long
Public lgTtlTime8 As Long   'Overall time getting regions

Public lgSTime9 As Long
Public lgETime9 As Long
Public lgTtlTime9 As Long   'Time checking LST region definition and its Ok

Public lgSTime10 As Long
Public lgETime10 As Long
Public lgTtlTime10 As Long  'Time adding LST region

Public lgSTime11 As Long
Public lgETime11 As Long
Public lgTtlTime11 As Long  'Time update LST region

Public lgSTime12 As Long
Public lgETime12 As Long
Public lgTtlTime12 As Long  'Time removing LST region
Public lgTtlTime13 As Long  'Time to check if LST should be removing

Public lgSTime14 As Long
Public lgETime14 As Long
Public lgTtlTime14 As Long  'Time checking for split copy

Public lgSTime15 As Long
Public lgETime15 As Long
Public lgTtlTime15 As Long

Public lgSTime16 As Long
Public lgETime16 As Long
Public lgTtlTime16 As Long

Public lgSTime17 As Long
Public lgETime17 As Long
Public lgTtlTime17 As Long

Public lgSTime18 As Long
Public lgETime18 As Long
Public lgTtlTime18 As Long

Public lgSTime19 As Long
Public lgETime19 As Long
Public lgTtlTime19 As Long

Public lgSTime20 As Long
Public lgETime20 As Long
Public lgTtlTime20 As Long

Public lgSTime21 As Long
Public lgETime21 As Long
Public lgTtlTime21 As Long

Public lgSTime22 As Long
Public lgETime22 As Long
Public lgTtlTime22 As Long

Public lgSTime23 As Long
Public lgETime23 As Long
Public lgTtlTime23 As Long

Public lgSTime24 As Long
Public lgETime24 As Long
Public lgTtlTime24 As Long

Public lgSTime25 As Long
Public lgETime25 As Long
Public lgTtlTime25 As Long

Public lgSTimeSQL As Long
Public lgETimeSQL As Long
Public lgTtlTimeSQL As Long

Public lgCount1 As Long
Public lgCount2 As Long
Public lgCount3 As Long
Public lgCount4 As Long
Public lgCount5 As Long
Public lgCount6 As Long
Public lgCount7 As Long
Public lgCount8 As Long
Public lgCount9 As Long
Public lgCount10 As Long
Public lgCount11 As Long
Public lgCount12 As Long

'Login variables
Public igAutoImport As Integer
Public igCompelAutoImport As Integer
Public igDemoMode As Integer
Public igSmallFiles As Integer
Public igTestSystem As Integer
Public igShowVersionNo As Integer   'Show version number status (0=Client; 1=internal not debug; 2=internal debug)
Public sgDDFDateInfo As String
Public sgWallpaper As String    'Obtained from ini
Public UID As String
Public PWD As String
Public sgPasswordAddition As String 'Message to add to Previous Password
Public sgSpecialPassword As String
Public bgInternalGuide As Boolean
Public igPasswordOk As Boolean
Public sgLastDateArch As String
Public igEmailNeedsConv As Integer
Public sgNumMoToRetain As String
Public igNumMoBehind As Integer
Public igWarning As Integer
'10000
Public sgWebServerSection As String
Public bgTestSystemAllowWebVendors As Boolean
'SQL varaibles
Public SQLQuery As String

'ADO variables
Public cnn As ADODB.Connection
Public rst As ADODB.Recordset
Public rst2 As ADODB.Recordset
Public gErrSQL As ADODB.Error

'RDO variables
'Public env As rdoEnvironment
'Public cnn As rdoConnection
'Public rst As ADODB.Recordset
'Public gErrSQL As rdoError

Public gMsg As String

'ReImport from Web and/or Marketron
Public sgReImportStatus As String

'VB variables
'Replaced with vbModal: Public Const Modal = 1
'Public Const vbDefault = 0
'Public Const vbHourglass = 11
Public Const Yes = 0
Public Const No = 1
'Public Const vbYes = 6
'Public Const vbNo = 7

'
Public Const SHIFTMASK = 1
Public Const CTRLMASK = 2
Public Const ALTMASK = 4
Public Const LEFTBUTTON = 1
Public Const RIGHTBUTTON = 2
Public Const KEYBACKSPACE = 8   'Back space key pressed
Public Const KEYDECPOINT = 46   'Decimal point pressed
Public Const KEY0 = 48          '0 key pressed
Public Const KEY9 = 57          '9 key pressed
Public Const KEYSLASH = 47      '/ Slash key
Public Const KEYLEFT = &H25
Public Const KEYUP = &H26
Public Const KEYRIGHT = &H27
Public Const KEYDOWN = &H28
Public Const KEYCOMMA = 44

'Line controls
Public sgBS As String * 1   'Backspace
Public sgTB As String * 1   'Tab
Public sgCR As String * 1   'Carriage Return
Public sgLF As String * 1   'Line Feed
Public sgCRLF As String * 2

'Blackout Pool Rotations
Public Const POOLROTATION = -1234
Public igLastPoolAdfCode As Integer
Public igPoolAdfCode() As Integer
Public igSortPoolAdfCode() As Integer

'Vehicle features
'UsingFeatures1
Public Const EXPORTINSERTION = &H1
Public Const EXPORTLOG = &H2
Public Const IMPORTINSERTION = &H4
Public Const IMPORTAFFILIATESPOTS = &H8
Public Const EXPORTISCIBYPLEDGE = &H10
'H20 left unused because old vehicles have a h20 (space) in this field
Public Const SUPPRESSWEBLOG = &H40
Public Const EXPORTPOSTEDTIMES = &H80

'Using Features2
Public Const XDSAPPLYMERGE = &H1


'Misc. variables
'Public IsDirty As Boolean
Public sgCurDir As String
Public IsRepDirty As Boolean      'Used in frmAffRep
Public IsYes As Boolean
Public IsNo As Boolean
Public iARIndex As Integer
Public iCPIndex As Integer
Public sFWeek As String
Public sLWeek As String
Public iCellColor As Integer
Public sContracts As String
Public sAdvtDates As String
Public RedText As Long
Public GreenText As Long
Public BlueText As Long
Public MagentaText As Long
Public sgClientName As String
Public gUsingUnivision As Boolean
Public gUsingWeb As Boolean
Public gWebAccessTestedOk As Boolean
Public sgExportISCI As String   'Not used. Retained with gISCIExport from Traffic spfFeatures8
Public sgRCSExportCart4 As String
Public sgRCSExportCart5 As String
Public sgUsingStationID As String
Public sgMissedMGBypass As String
Public sgUsingServiceAgreement As String
Public sgWebNumber As String    'Web Number (verion). TTP 9052
'8/1/14: Not used with v7.0
'Public sgMarketronCompliant As String 'A=Advertiser, P=Pledge
Public igRCSExportBy As Integer '4=4 Digit cart number format; 5=5 digit cart number format
Public sgShowByVehType As String
Public sgMDBPath As String
Public sgSDBPath As String
Public sgTDBPath As String
Public igRetrievalDB As Integer
Public igWebIniOptionsOK As Boolean
Public igSendAvails As Integer
Public sgDelNet As String   'using Delivery Links: Y=Yes; N=No

Public sgUserName As String
Public sgReportName As String
Public igUstCode As Integer
Public sgUstWin(0 To 14) As String * 1
Public sgUstPledge As String * 1
Public sgUstClear As String * 1
Public sgUstActivityLog As String * 1
Public sgUstDelete As String * 1
Public sgUstAllowCmmtChg As String * 1
Public sgUstAllowCmmtDelete As String * 1
Public sgExptSpotAlert As String
Public sgExptISCIAlert As String
Public sgTrafLogAlert As String
Public sgPhoneNo As String
Public igUstSSMnfCode As Integer
Public sgCity As String
Public sgEMail As String
Public lgEMailCefCode As Long
Public sgAllowedToBlock As String
Public sgChgExptPriority As String
Public sgExptSpec As String
Public sgChgRptPriority As String
Public sgCPRetStatus As String
Public sgErrorMsg As String

Public igUserLogButton As Integer   '1=Message; 2=Send Alert; 3=Send Ok message
Public lgUserLogUlfCode() As Long    'Traffic User Log to receive message or alert

Public sgActiveLogDate As String  'Last date the ulfActiveLogDate updated
Public sgActiveLogTime As String  'Last time the ulfActiveLogTime updated
Public lgActiveUlfCode As Long

Public igMergeType As Integer   '0=Market; 1=Format

Public sgBIAFileName As String
Public igBIARetStatus As Integer

Public sgSpfUseCartNo As String
Public sgSpfRemoteUsers As String

Public sgSpfUsingFeatures2 As String
Global Const SPLITCOPY = &H2
Global Const SPLITNETWORKS = &H4
Global Const STRONGPASSWORD = &H10

'Sport Info
Public sgSpfSportInfo As String
Global Const USINGSPORTS = &H1
Global Const PREEMPTREGPROG = &H2
Global Const USINGFEED = &H4
Global Const USINGLANG = &H8

Public sgSpfUsingFeatures5 As String
Global Const STATIONINTERFACE = &H20
Global Const RADAR = &H40

Public sgSpfUsingFeatures9 As String
Global Const AFFILIATECRM = &H1 'Affiliate CRM

'Using Features10
Global Const ADDADVTTOISCI = &H1 'X-Digital: Add Advertiser name to ISCI
Global Const MIDNIGHTBASEDHOUR = &H2     'X-Digital: Spot Insertion using Midnight Based time (12=0, 11,...11p=23)
Global Const PKGLNRATEONBR = &H4
Global Const WEGENERIPUMP = &H8
'9114
'Global Const UNITIDBYASTCODE = &H80 'X-Digital: substitue astCode for the Unit ID and generate xml table of ISCI
Global Const UNITIDBYASTCODEFORBREAK = &H80 'X-Digital: substitute astCode for the Unit ID for hb and hbp
Global Const UNITIDBYASTCODEFORISCI = &H2 'X-Digital: substitute astCode for the Unit ID for isci method
'Features1 stored into SAF
Global Const JELLIEXPORT = &H40
Global Const COMPENSATION = &H80
'Features2 stored into SAF
Global Const EMAILDISTRIBUTION = &H8

'Features3 stored into SAF
'Global Const SUPPRESSNETCOMM = &H1
'Global Const REQSTATIONPOSTING = &H2
Global Const SPLITCOPYSTATE = &HC0      'Split Copy Station State (0=Mailing; 1=License; 2=Physical)
Global Const SPLITCOPYLICENSE = &H4
Global Const SPLITCOPYPHYSICAL = &H8
'Global Const FREEZEDEFAULT = &H10       'Freeze calculation default
'Features5 stored into SAF
Global Const PROGRAMMATICALLOWED = &H1  'Programmatic Allowed
Global Const CSVAFFIDAVITIMPORT = &H80
'Features6 stored into SAF
Global Const OVERDUEEXPORT = &H80  'Affiliate Overdue Export

Public sgSplitState As String * 1       'M=Mailing; L=Licence; P=Physical

Public sgSpfUseProdSptScr As String

'Ini Values
Public sgStartupDirectory As String
Public sgIniPathFileName As String
Public sgDatabaseName As String
Public sgURL_Documentation As String
Public sgReportDirectory As String
Public sgExportDirectory As String
Public sgImportDirectory As String
Public sgExeDirectory As String
Public sgLogoDirectory As String
Public sgLogoName As String 'dan
Public sgContractPDFPath As String
Public sgISCIxRefExportPath As String 'TTP 10457 - ISCI Cross Reference Export
Public igWaitCount As Integer
Public igTimeOut As Integer 'ADO Query Timeout (-1, use default, in sec)
Public sgMsgDirectory As String
Public sgDBPath As String

Public igSQLSpec As Integer             '0=Pervasive 7; 1= Pervasive 2000 (default)
Public sgSQLDateForm As String          'Default: yyyy-mm-dd
Public sgSQLTimeForm As String          'Default: hh:mm:ss
Public sgShowDateForm As String         'Default m/d/yyyy
Public sgShowTimeWOSecForm As String    'Default h:mma/p
Public sgShowTimeWSecForm As String     'Default h:mm:ssa/p
Public sgCrystalDateForm As String
'Crystal Vars
'Public Appl As New CRAXDRT.Application 'dan commented out don't need with crxi

'Generic Storage areas for formulas
Public sgCrystlFormula1 As String
Public sgCrystlFormula2 As String
Public sgCrystlFormula3 As String
Public sgCrystlFormula4 As String
Public sgCrystlFormula5 As String
Public sgCrystlFormula6 As String   '7-12-04
Public sgCrystlFormula7 As String   '7-12-04
Public sgCrystlFormula8 As String   '7-12-04
Public sgCrystlFormula9 As String   '7-12-04
Public sgCrystlFormula10 As String   '7-12-04
Public sgCrystlFormula11 As String   '8-17-09
Public sgCrystlFormula12 As String   '5-21-12
Public sgCrystlFormula13 As String      '7-11-13

'Public sgRptName As String  'dan commented out don't need with cr2008

'Game selection variables
Public igGameVefCode As Integer
Public sgGameStartDate As String
Public sgGameEndDate As String
Public lgSelGameGsfCode As Long
Public igSelGameNo As Integer
Public sgSelGameDate As String
Public lgGameAttCode As Long


'Model parameters
Public igModelType As Integer   '1=From Radar(show vehicles with program schedules only)
Public igModelReturn As Integer  'True if item selected; False if cancelled or no item selected to model from
Public lgModelFromCode As Long  'Code of item to model from
Public sgResultFileName As String

'5/26/13: Report Queue
'Report Queue parameters
Public bgReportQueue As Boolean
Public igReportSource As Integer    '1=From button on Affiliate Directory screen; 2=From Report Queue
Public igReportReturn As Integer    '0=Cancelled; 1=Ok; 2=Error
Public sgRQReportName As String
Public lgReportRqtCode As Long
Public sgRQDescription As String
Public igRQReturnStatus As Integer
Public igReportModelessStatus As Integer    '0=Processing, 1=Completed

'Market Returned parameters
Public igMarketReturn As Integer    'True if market referenced, False if canceled
Public igMarketReturnCode As Integer

'Ownery Returned parameters
Public igOwnerReturn As Integer    'True if owner referenced, False if canceled
Public lgOwnerReturnCode As Long

'Format Returned parameters
Public igFormatReturn As Integer    'True if Format referenced, False if canceled
Public igFormatReturnCode As Integer
Public sgFormatCall As String   'M=Menu iten, N=Name, Name if from station and new not selected
Public sgFormatName As String

'Format Returned parameters
Public igGNMarketReturn As Integer    'True if Format referenced, False if canceled
Public igGNMarketReturnCode As Integer
Public sgGNMarketCall As String   'M=Menu iten, N=Name, Name if from station and new not selected
'Format Returned parameters
Public igTimeZoneReturn As Integer    'True if Format referenced, False if canceled
Public igTimeZoneReturnCode As Integer
Public sgTimeZoneCall As String   'M=Menu iten, N=Name, Name if from station and new not selected
'Format Returned parameters
Public igStateReturn As Integer    'True if Format referenced, False if canceled
Public igStateReturnCode As Integer
Public sgStateCall As String   'M=Menu iten, N=Name, Name if from station and new not selected
'Format Returned parameters
Public igVehicleReturn As Integer    'True if Format referenced, False if canceled
Public igVehicleReturnCode As Integer
Public sgVehicleCall As String   'M=Menu iten, N=Name, Name if from station and new not selected

Public igDepartmentReturn As Integer    'True if Department referenced, False if canceled
Public igDepartmentReturnCode As Integer
Public sgDepartmentName As String

'MultiName
Public sgMultiNameType As String * 1    'C=Copy; Y=County; A=Area; T=Territory; O=Operator; M=Moniker
Public sgMultiNameName As String
Public lgMultiNameCode As Long
Public igMultiNameReturn As Integer    'True if MultiName referenced, False if canceled

'Comment Source
Public sgCmmtSrcName As String
Public igCmmtSrcCode As Integer
Public igCmmtSrcReturn As Integer    'True if Comment Source referenced, False if canceled



'Alerts
Type AUF
    lCode As Long    'Internal code number for alert user file
    lEnteredDate As Long  'Entered Date
    lEnteredTime As Long  'Entered Time
    sStatus As String * 1   'Alert Status (C=Cleared; R=Requires Alert menu to be shown)
    sType As String * 1     'Alert Type (C=Contract; L=Log Reprinting required; F=Final Log Generated; R=Reprint Log Generated)
    sSubType As String * 1  'Alert Sub Type
                            'AufType = C:  C=Proposal changed to Complete
                            'AufType = L:  C=Copy Assigned; S=Spot Scheduled
                            'AufType = F:  I=ISCI Export; S=Spot Export
                            'AufType = R:  I=ISCI Export; S=Spot Export
    lChfCode As Long        'Contract Code: aufType = C
    iVefCode As Integer     'Vehicle Code: aufType = L
                            'When first created (selling vehicle will be used;
                            'When generating Logs or click on Menu Alert, the
                            'selling vehicle will be replaced with airing vehicle
    lMoWeekDate As Long  'Monday week date: aufType = L
    iCreateUrfCode As Integer   'User Code of person who created the Alert
    iCreateUstCode As Integer   'User Code of person who created the Alert
    iClearUrfCode As Integer    'User code of person who cleared the Alert on Traffic system
    iClearUstCode As Integer    'User code of person who cleared the Alert on Affiliate System
    sClearMethod As String * 1  'Clear method: M=Manually; A=Automatically
    lClearDate As Long  'Cleared Date
    lClearTime As Long  'Cleared Time
    lUlfCode As Long            ' User Log reference code
    lCefCode As Long            ' Notification Comment (aufType = N)
    iCountdown As Integer       ' Initial countdown value for shutdown
    sSpotCopyChg(0 To 6)  As String * 1      ' Spot and/or generic copy changed
                                             ' on Monday (Y/N). Test for Y.
    sRegionCopyChg(0 To 6) As String * 1     ' Region copy changed on Tuesday
                                             ' (Y/N). Test for Y
    sUnused As String * 10
End Type
    
Type AUFVIEW
    sKey As String * 20
    tAuf As AUF
End Type

Public tgAuf As AUF
Public igAlertTimer As Integer 'Number of minutes since last checked Alerts
Public igAlertFlash As Integer 'Number of times Alert as Flashed
Public igAlertInterval As Integer    'Obtained from spfGAlertInterval
Public rstAlert As ADODB.Recordset
Public rstAlertUlf As ADODB.Recordset
Public SQLAlertQuery As String
Public SQLAlertULF As String

Type CSF
    lCode As Long    'Internal code number for Comment-Contract
    iAdfCode As Integer 'Advertiser code number
    sType As String * 1 'S=Copy inventory scripts; R=Rotation comment
    'iStrLen As Integer  'String length (required by LVar)
    'sComment As String * 5002   'Last two bytes after the comment must be 0
    sComment As String * 5004   'Last bytes after the comment must be 0
End Type

Type LONGKEY0
    lCode As Long
End Type

'Cef record layout
Type CEF
    lCode As Long    'Internal code number for Comment
    sUnused As String * 1   'Fix length portion of record must be 5 bytes
    'iStrLen As Integer  'String length (required by LVar)
    'sComment As String * 1002   'Last two bytes after the comment must be 0
    sComment As String * 1004   'Last bytes after the comment must be 0
End Type

Type CPYROTCOM
    lCode As Long
    sComment As String
End Type

Public tgCopyRotInfo() As CPYROTCOM

Type GAMEINFO
    lgsfCode As Long
    sGameDate As String * 10
    sGameStartTime As String * 10
    sVisitTeamName As String * 20
    sVisitTeamAbbr As String * 10
    sHomeTeamName As String * 20
    sHomeTeamAbbr As String * 10
    sLanguageCode As String * 20
    sFeedSource As String * 1
    sEventCarried As String * 1
    lAttCode As Long
End Type

Public tgGameInfo() As GAMEINFO

Type STATIONINFO
    iCode As Integer
    sCallLetters As String * 40
    sMarket As String * 60
    iType As Integer
    lID As Long
    sZone As String * 3
    iMktCode As Integer
    sRank As String  'TTP 10051 - JW - 6/30/21 - for Station list loading Performance
    iMSAMktCode As Integer
    lOwnerCode As Long
    sOwner As String 'TTP 10051 - JW - 6/30/21 - for Station list loading Performance
    iFormatCode As Integer
    sFormat As String 'TTP 10051 - JW - 6/30/21 - for Station list loading Performance
    lMntCode As Long
    sTerritory As String * 20
    lAreaMntCode As Long
    iTztCode As Integer     'Time zone
    sPostalName As String * 2
    sUsedForATT As String * 1
    sUsedForXDigital As String * 1
    sUsedForWegener As String * 1
    sUsedForOLA As String * 1
    sUsedForPledgeVsAir As String * 1
    sSerialNo1 As String * 10
    sSerialNo2 As String * 10
    sPort As String * 1
    lPermStationID As Long
    iAckDaylight As Integer '0=Yes, 1=No
    sZip As String * 20
    sWebAddress As String * 90
    sWebPW As String * 10
    sFrequency As String * 6
    lMonikerMntCode As Long
    lMultiCastGroupID As Long
    lMarketClusterGroupID As Long
    sAgreementExist As String * 1
    sCommentExist As String * 1
    iMktRepUstCode As Integer
    iServRepUstCode As Integer
    lCityLicMntCode As Long
    lHistStartDate As Long
    sStationType As String  'C=Commercial; N=Non-Commercial
    lCountyLicMntCode As Long
    sMailAddress1 As String * 40
    sMailAddress2 As String * 40
    lMailCityMntCode As Long
    sMailState As String * 40
    sOnAir As String * 1    'Y=Yes; N=No
    lOperatorMntCode As Long
    lAudP12Plus As Long
    sPhone As String * 20
    sFax As String * 30
    sPhyAddress1 As String * 40
    sPhyAddress2 As String * 40
    lPhyCityMntCode As Long
    sPhyState As String * 40
    sPhyZip As String * 20
    sStateLic As String * 2
    sWebEMail As String * 1 '0=No; 1=Yes
    sEnterpriseID As String * 5
    lXDSStationID As Long
    lWatts As Long
    lClusterGrougID As Long
    sMasterCluster As String * 1
    '8418
    sWebNumber As String * 1
End Type

Public tgStationInfo() As STATIONINFO
Public tgStationInfoByCode() As STATIONINFO
Public tgShttSavedInfo(0) As STATIONINFO

Public sgShttTimeStamp As String

Public sgReplacementStamp As String

Type CPFINFO
    lCode              As Long            ' Internal code number for copy Produc
                                          ' t/Agency
    sName              As String * 35     ' Name
    sISCI              As String * 20     ' Agency ISCI code
    sCreative          As String * 30     ' Creative title
    sRotEndDate        As String * 10     ' Latest Rotation date using this inve
                                          ' ntory; Date Byte 0:Day, 1:Month, fol
                                          ' lowed by 2 byte year
    lsifCode           As Long            ' Short Title used with cmml sch & bul
                                          ' k feed
End Type
Public tgCpfInfo() As CPFINFO

Type MARKETINFO
    lCode As Long
    sName As String * 60
    iRank As Integer
    sBIA As String * 10
    sARB As String * 10
    sGroupName As String * 10
End Type

Public tgMarketInfo() As MARKETINFO
Public tgMSAMarketInfo() As MARKETINFO

Type MNTINFO
    lCode As Long
    sName As String * 40
    sState As String * 1
End Type


Public tgTerritoryInfo() As MNTINFO
Public tgCityInfo() As MNTINFO
Public tgCountyInfo() As MNTINFO
Public tgAreaInfo() As MNTINFO
Public tgMonikerInfo() As MNTINFO
Public tgOperatorInfo() As MNTINFO

Type REPINFO
    iUstCode As Integer
    sName As String * 20
    sLogInName As String * 20
    sReportName As String * 20
End Type

Public tgMarketRepInfo() As REPINFO
Public tgServiceRepInfo() As REPINFO


Type AFFAEINFO
    lCode As Long
    sFirstName As String * 20
    sLastName As String * 60
    sName As String * 81
    sEmail As String * 70
    iTntCode As Integer
End Type
Public tgAffAEInfo() As AFFAEINFO

Type ASTADDCODES
    lCode As Long
End Type

Type ASTDELETECODES
    lCode As Long
End Type


'6/29/06: change gGetAstInfo to use API call
'******************************************************************************
' ast Record Definition
'
'******************************************************************************
Type AST
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
    lCpfCode              As Long            ' Product/ISCI reference of spot
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
End Type

Type ASTKEY0
    lCode                 As Long
End Type

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


'7/15/11: ASTStatus field definitions
Public Const ASTAIR_LIVE = 0
Public Const ASTAIR_DELAY = 1
Public Const ASTAIR_NA_TECH = 2
Public Const ASTAIR_NA_BLACKOUT = 3
Public Const ASTAIR_NA_OTHER = 4
Public Const ASTAIR_NA_PRODUCT = 5
Public Const ASTAIR_OUTSIDE = 6
Public Const ASTAIR_NOTPLEDGED = 7
Public Const ASTAIR_NOTCARRIED = 8
Public Const ASTAIR_CMMLPRG = 9
Public Const ASTAIR_CMML = 10

Public Const ASTAIR_MISSED_MG_BYPASS = 14

Public Const ASTEXTENDED_MG = 11
Public Const ASTEXTENDED_BONUS = 12
Public Const ASTEXTENDED_REPLACEMENT = 13
Public Const ASTEXTENDED_MISSREASON = 100
Public Const ASTEXTENDED_ISCICHGD = 1000

Type VEHICLEINFO
    iCode As Integer
    sVehType As String * 1
    'sVehicle As String * 40
    sVehicle As String * 42 'Type: Name
    sVehicleName As String * 40
    sCodeStn As String * 5
    iNoDaysCycle As Integer
    sPrimaryZone As String * 3  'First Zone with Local Adjustment = 0
    'sPrimaryFeed As String * 3  'First Zone with Feed
    iNoZones As Integer     'Number of zones
    sZone(0 To 3) As String * 3 'Zone names
    iLocalAdj(0 To 3) As Integer    'Local adjustment corrected as relative to related *
                                    'because this value is used to adjust the LST
                                    'Use VehLocalAdj to correct times from traffic (avails and daypart)
                                    'i.e.
                                    '  as in VPG
                                    'Zone  Aff Adj    Fed
                                    'EST     0        *
                                    'CST    -1        E
                                    'MST    -2        P
                                    'PST    -3        *
                                    'In Vehicle Pop the Local Adsj and VehLocalAdj are set as follows
                                    'Zone Local Adj   Veh Local Adj
                                    'EST     0           0
                                    'CST     -1          -1
                                    'MST     +1          -2
                                    'PST     0           -3
                                    'LST will be created for EST and PST since both have *
    iVehLocalAdj(0 To 3) As Integer    'Vehicle Local adjustment from VPF
    sFed(0 To 3) As String * 1  '*=Feed Zone(lst created for this zone); Letter if zone that LST created
    iBaseZone(0 To 3) As Integer    'If sFed <> *, then this contains index of base zone
    'iESTEndTime(1 To 5) As Integer
    iESTEndTime(0 To 4) As Integer
    'iMSTEndTime(1 To 5) As Integer
    iMSTEndTime(0 To 4) As Integer
    'iCSTEndTime(1 To 5) As Integer
    iCSTEndTime(0 To 4) As Integer
    'iPSTEndTime(1 To 5) As Integer
    iPSTEndTime(0 To 4) As Integer
    lHd1CefCode As Long
    lLgNmCefCode As Long
    lFt1CefCode As Long
    lFt2CefCode As Long
    iVpfSAGroupNo As Integer
    sState As String * 1
    iProducerArfCode As Integer     'Produce ArfCode reference
    iProgProvArfCode As Integer   'Export Program Audio ArfCode Reference
    iCommProvArfCode As Integer   'Export Commercial Audio ArfCode reference
    sEmbeddedComm As String * 1   'Export Commercial Audio (Yes/No); Y = Export Comm Aud for Selected reference (not to stations); N=Export Comm Aud for Stations
    iMnfVehGp2 As Integer       'Subtotal Group Field
    iInterfaceID As Integer
    sWegenerExport As String * 1
    sOLAExport As String * 1
    iVefCode As Integer
    iOwnerMnfCode As Integer            'all vehicle grps added into public vehicleinfo array - 5-31-16
    iMnfVehGp3Mkt As Integer
    iMnfVehGp4Fmt As Integer
    iMnfVehGp5Rsch As Integer
    iMnfVehGp6Sub As Integer
End Type

Public tgVehicleInfo() As VEHICLEINFO

Type SELLINGVEHICLEINFO
    iCode As Integer
    sVehType As String * 1
    sVehicle As String * 40
    sCodeStn As String * 5
    iNoDaysCycle As Integer
    sPrimaryZone As String * 3  'First Zone with Local Adjustment = 0
    'sPrimaryFeed As String * 3  'First Zone with Feed
    iNoZones As Integer     'Number of zones
    sZone(0 To 3) As String * 3 'Zone names
    iLocalAdj(0 To 3) As Integer    'Local adjustment
    iVehLocalAdj(0 To 3) As Integer    'Vehicle Local adjustment from VPF
    sFed(0 To 3) As String * 1  '*=Feed Zone(lst created for this zone); Letter if zone that LST created
    iBaseZone(0 To 3) As Integer    'If sFed <> *, then this contains index of base zone
    'iESTEndTime(1 To 5) As Integer
    iESTEndTime(0 To 4) As Integer
    'iMSTEndTime(1 To 5) As Integer
    iMSTEndTime(0 To 4) As Integer
    'iCSTEndTime(1 To 5) As Integer
    iCSTEndTime(0 To 4) As Integer
    'iPSTEndTime(1 To 5) As Integer
    iPSTEndTime(0 To 4) As Integer
    lHd1CefCode As Long
    lLgNmCefCode As Long
    lFt1CefCode As Long
    lFt2CefCode As Long
    iVpfSAGroupNo As Integer
End Type

Public tgSellingVehicleInfo() As SELLINGVEHICLEINFO

Type VPFOPTIONS
    ivefKCode          As Integer         ' Internal code number for Vehicle Opt
    iSAGroupNo         As Integer
    iLNoDaysCycle      As Integer         ' Number of days in closing cycle (i.e 1, 7 ,14)
    sGZone1            As String * 3      ' Time zone
    iGLocalAdj1        As Integer         ' Local time adjustment
    sFedZ1             As String * 1      ' Transmit: N=No; Y=Yes
    sGZone2            As String * 3
    iGLocalAdj2        As Integer
    sFedZ2             As String * 1
    sGZone3            As String * 3
    iGLocalAdj3        As Integer
    sFedZ3             As String * 1
    sGZone4            As String * 3
    iGLocalAdj4        As Integer
    sFedZ4             As String * 1
    iESTEndTime1       As Integer         ' Daypart End Time of EST- in minutes
    iESTEndTime2       As Integer
    iESTEndTime3       As Integer
    iESTEndTime4       As Integer
    iESTEndTime5       As Integer
    iMSTEndTime1       As Integer         ' Daypart End Time of MST- in minutes
    iMSTEndTime2       As Integer
    iMSTEndTime3       As Integer
    iMSTEndTime4       As Integer
    iMSTEndTime5       As Integer
    iCSTEndTime1       As Integer         ' Daypart End Time of CST- in minutes
    iCSTEndTime2       As Integer
    iCSTEndTime3       As Integer
    iCSTEndTime4       As Integer
    iCSTEndTime5       As Integer
    iPSTEndTime1       As Integer         ' Daypart End Time of PST- in minutes
    iPSTEndTime2       As Integer
    iPSTEndTime3       As Integer
    iPSTEndTime4       As Integer
    iPSTEndTime5       As Integer
    lLgHd1CefCode      As Long            ' Header comment 1 from vehicle options
    lLgNmCefCode       As Long            ' Vehicle Log Name from vehicle options
    lLgFt1CefCode      As Long            ' Footer comment 1 from vehicle options
    lLgFt2CefCode      As Long            ' Footer comment 2 from vehicle options
    iProducerArfCode   As Integer         ' Produce ArfCode reference
    iProgProvArfCode   As Integer         ' Program Content Provider ArfCode Reference
    iCommProvArfCode   As Integer         ' Commercial Content Provider ArfCode reference
    sEmbeddedComm      As String * 1      ' Export Commercial Audio (Yes/No); Y = Export Comm Aud for Selected refer
                                          ' ence (not to stations); N=Export Comm Aud for Stations
    sAvailNameOnWeb     As String * 1     'show avail name on web
    sUsingFeatures1 As String * 1
    sWebLogSummary As String * 1    'Show web log spot summary (Y/N)
    sWebLogFeedTime As String * 1   'Show on web log spot feed time (Y/N)
    lEDASWindow As Long
    sStnFdXRef As String * 1
    sAllowSplitCopy As String * 1       'Y=Split Copy allowed
    iInterfaceID As Integer
    sWegenerExport As String * 1
    sOLAExport As String * 1
    sUsingFeatures2 As String * 1
    sEmbeddedOrROS As String * 1
    '11/26/17
    sLLD As String * 12
End Type

Public tgVpfOptions() As VPFOPTIONS

Type ADVTINFO
    iCode As Integer
    sAdvtName As String * 35
    sAdvtAbbr As String * 7
End Type
'6191
Type AGENCYINFO
    iCode As Integer
    sAgencyName As String * 40
    sAgencyAbbr As String * 5
End Type
Public tgAdvtInfo() As ADVTINFO
Public tgAgencyInfo() As AGENCYINFO
Type DAYPARTINFO
    iCode As Integer    'Daypart code
    sStartTime As String * 10  'Earliest start time obtained from the daypart
    sEndTime As String * 10    'Latest end time obtained from the daypart
End Type

Public tgDaypartInfo() As DAYPARTINFO

Type LSTINFO
    lstCode As Long
    lstLogVefCode As Integer
    lstLogDate As String * 10
    lstLogTime As String * 10
End Type

Public tgLstInfo() As LSTINFO

Type ATTINFO1
    attCode As Long
    attvefCode As Integer
    attShttCode As Integer
    attExportType As Integer
    attPledgeType As String * 1
End Type

Public tgAttInfo1() As ATTINFO1
Public lgAttCount As Long

Type CRFINFO1
    crfCode As Long
    crfCsfCode As Long
End Type

Public tgCrfInfo1() As CRFINFO1

Type CIFCPFINFO1
    cifCode As Long
    cifRotEndDate As String * 10
    cifAdfCode As Integer
    cifName As String * 5
    cifMcfCode As Integer
    cifReel As String * 10
    cifCpfCode As Long
    cpfISCI As String * 20
    cpfCreative As String * 30
    cpfName As String * 35
End Type

Public tgCifCpfInfo1() As CIFCPFINFO1

Type SHTTINFO1
    shttCode As Integer
    shttTimeZone As String * 3
End Type

Public tgShttInfo1() As SHTTINFO1
Public tgStationCount() As SHTTINFO1

Type CPTTINFO
    cpttCode As Long
    cpttatfCode As Long
    cpttStatus  As Integer
    cpttPostingStatus As Integer
    CpttStartDate As String * 10
End Type

Public tgCpttInfo() As CPTTINFO
Public lgCpttCount As Long

Type WEBSPOTSINFO
    lAstCode As Long    'AstCode
    iFlag As Integer    '0=Not Posted, 1=Posted but not submitted, 2=Posted and submitted
End Type

Type WEBINFO
    sFileName As String * 80
    sExeToRun As String * 30
    sTypeExpected As String * 30
    iStatus As Integer
    sCommand As String * 80
    'CommentStatus As String * 80
    'CommentTtls As String * 80
    'HeadersStatus As String * 40
    'HeadersTtls As String * 40
    'SpotsStatus As String * 40
    'SpotsTtl As String * 40
    'EmailStatus As String * 40
    'EmailTtl As String * 40
    'Reindex As String * 40
End Type


Public sgAufsKey() As String * 12

'Blackout Replacement for Split Networks
Type BOF
    lCode              As Long               ' AutoInc
    sType              As String * 1         ' S=Supression; R=Replacement
    iAdfCode           As Integer            ' Advertiser Code
    lsifCode           As Long               ' Short Title or Product
    iVefCode           As Integer            ' Vehicle Code (required for
                                             ' bofType = S only)
    lCifCode           As Long               ' Copy Inventory Code
    imnfComp1          As Integer            ' Product Protection Codes
                                             ' (Required for bofType = R)
    imnfComp2          As Integer
    sStartDate         As String * 10        ' Start date
    sEndDate           As String * 10        ' End Date (TFN allowed)
    sMo                As String * 1         ' Days (Y=Yes; N=No) Index 0 =
                                             ' Monday; 1= Tuedays;...
    sTu                As String * 1
    sWe                As String * 1
    sTh                As String * 1
    sFr                As String * 1
    sSa                As String * 1
    sSu                As String * 1
    sStartTime         As String * 11        ' Start Time
    sEndTime           As String * 11        ' End Time
    iurfCode           As Integer            ' User Code (Last to add or modify)
    lSChfCode          As Long               ' Suppress Contract Code
    iRAdfCode          As Integer            ' Replace Advertiser Code
    lRChfCode          As Long               ' Replace Contract Code
    iLen               As Integer
    sSource            As String * 1         ' N=Export NY; L=Log (Default)
    sUnused            As String * 9
End Type

Type BOFREC
    sKey As String * 20 'Random number
    tBof As BOF
    iLen As Integer
End Type

Public tgRBofRec() As BOFREC

Type DATRST
    lCode              As Long
    lAtfCode           As Long
    iShfCode           As Integer
    iVefCode           As Integer
    'iDACode            As Integer            ' 0=Daypart;1=Avail: Moved to Agreement attPledgeType (A, D or C)
    iFdMon             As Integer            ' 0=N; 1=Y
    iFdTue             As Integer
    iFdWed             As Integer
    iFdThu             As Integer
    iFdFri             As Integer
    iFdSat             As Integer
    iFdSun             As Integer
    sFdStTime          As String * 11
    sFdEdTime          As String * 11
    iFdStatus          As Integer            ' 0=Carried;1=Delay;2-5=Not;7=Specl
                                             ' ;8=Off air
    iPdMon             As Integer
    iPdTue             As Integer
    iPdWed             As Integer
    iPdThu             As Integer
    iPdFri             As Integer
    iPdSat             As Integer
    iPdSun             As Integer
    sPdStTime          As String * 11
    sPdEdTime          As String * 11
    sPdDayFed          As String * 1         ' Clarify when pledge day is prior to feed day if air date is Before (B) or After (A)
                                             ' Default is After (i.e. test for B)
    iAirPlayNo         As Integer            ' Air play number
    sEstimatedTime     As String * 1         ' Estimated Time allowed to be defined (Y or N). Test for Y.  This is only valid for Dayparts
    sEmbeddedOrROS     As String * 1         ' Delivery Embedded spots or ROS spots (E/R).  Test for E. Blank is the same as R.
    sUnused            As String * 15        ' Unused
End Type

Type LST
    lCode              As Long
    iType              As Integer            ' 0=Spot; 1=Avail
    lSdfCode           As Long               ' SdfCode
    lCntrNo            As Long               ' Contract Number
    iAdfCode           As Integer            ' AdfCode
    iAgfCode           As Integer            ' AgyCode
    sProd              As String * 35        ' Product Name
    iLineNo            As Integer            ' Line number
    iLnVefCode         As Integer            ' Line Vehicle vefcode
    sStartDate         As String * 10        ' Flight Start Date
    sEndDate           As String * 10        ' Flight End Date
    iMon               As Integer            ' Index 0:Monday,
                                             ' 1=Tuesday,...6=Sunday; 0=No;
                                             ' 1=Yes
    iTue               As Integer
    iWed               As Integer
    iThu               As Integer
    iFri               As Integer
    iSat               As Integer
    iSun               As Integer
    iSpotsWk           As Integer            ' Spots per week
    iPriceType         As Integer            ' Price type (0=True; 1=Bonus,...)
    lPrice             As Long               ' Price (xxxxx.xx)
    iSpotType          As Integer            ' Spot Type (0=Schd; 1=MG;
                                             ' 2=Filled; 3=Outside; 4=
                                             ' ;5=Added; 6=Open BB; 7=Close BB)
    iLogVefCode        As Integer            ' Log (airing) vehicle code
    sLogDate           As String * 10        ' Log date
    sLogTime           As String * 11        ' Log Time
    sDemo              As String * 6         ' Demo Name
    lAud               As Long               ' Audience for primary demo
    sISCI              As String * 20        ' ISCI code from copy
    iWkNo              As Integer
    iBreakNo           As Integer
    iPositionNo        As Integer
    iSeqNo             As Integer
    sZone              As String * 3
    sCart              As String * 7
    lCpfCode           As Long
    lCrfCsfCode        As Long
    iStatus            As Integer            ' 0=Carried; 1=Delay; 2-5=Not
                                             ' Carried; 7=Special; 8=Off Air
    iLen               As Integer            ' Spot length or unsold avail
                                             ' length
    iUnits             As Integer            ' Avail Units (Unsold)
    lCifCode           As Long               ' CifCode (required by odf)
    iAnfCode           As Integer            ' Avail name (required by odf to
                                             ' determine if avail starts with N)
    lEvtIDCefCode      As Long               ' Event ID
    sSplitNetwork      As String * 1         ' Split Network Flag: N=Not split
                                             ' network spot; P=Primary Split
                                             ' Network spot; S=Secondary Split
                                             ' Network Spot
    lRafCode           As Long               ' Split network Region (Rafcode)
                                             ' reference or zero if not a split
                                             ' network spot
    'sUnused            As String * 20        ' Unused
    '12/28/06: changed Unused to 11 and added 3 fields
    lFsfCode As Long    'Feed Spot (0 if contract spot)
    lgsfCode As Long    'Game schedule (0 if not game spot)
    sImportedSpot As String * 1     'M=MYL spot import; F=Feed Spot import (fsfCode will be zero and Contract number will be blank)
    lBkoutLstCode As Long   'Used to reference the lst blackout because of copy splits by this spot.
                            'This record create so that ast will show the create advertiser name.
    sLnStartTime          As String * 11        ' Line Start time.  Override time if defined, otherwise daypart time
    sLnEndTime            As String * 11        ' Line End time.  Override time if defined, otherwise daypart time
    sUnused               As String * 20     ' Unused
End Type

Type SPLITNETLASTFILL
    iBofIndex As Integer
    iFillLen As Integer
End Type

Public tgSplitNetLastFill() As SPLITNETLASTFILL

Public Const GRIDSCROLLWIDTH = 270

Type STATUSTYPES
    sName As String * 30
    iPledged As Integer '0=Live; 1=Delayed; 2=Not Carried; 3=No Pledge
    iStatus As Integer
End Type

Public tgStatusTypes(0 To 14) As STATUSTYPES

Type RNFINFO
    sName As String * 60
    iCode As Integer
    sRptExe As String * 12
End Type
Public tgRnfInfo() As RNFINFO

Type ISCIXREF
    sKey As String * 90     'CommProvArfCode | VehName | Call Letters
    iCommProvArfCode As Integer
    iEmbedded As Integer
    sVehName As String * 42
    sCallLetters As String * 40
    iShfCode As Integer
    iVefCode As Integer
End Type

Type ISCISENDINFO
    sKey As String * 120 'Compressed Format: CommProvArfCode | Call Letters | ISCI  ->Not Embedded
                         '                   !CommProvArfCode | Producer | ISCI  ->Not Embedded
                         'All Format:  CommProvArfCode | Vehicle Name | Call Letters | ISCI  -> Not Embedded
                         '             !CommProvArfCode | Vehicle | ISCI  -> Embedded
    iCommProvArfCode As Integer
    iEmbedded As Integer
    sVehName As String * 42
    sCallLetters As String * 41
    sISCI As String * 20
    sAdvtName As String * 30
    iShfCode As Integer
    iProducerArfCode As Integer
    lLatestFeedDate As Long
    lEarlestFeedDate As Long
    lEitCode As Long
    lDateSent As Long
    lLLDRef As Long
    lLLDSent As Long
    iUpdateDateSent As Integer
    iVefCode As Integer
    iACExistWithAtt As Integer  'Affidavit Name exist within Agreement
End Type


Type EMAILREF
    sKey As String * 90     'VehName | Call Letters | Date
    sVehName As String * 42
    sCallLetters As String * 40
    sWebEMail As String * 1000 'Collection of email addresses (1 to many)    TTP 10620 - JJB 2023-04-21
    sWebPW As String * 10
    iShfCode As Integer
    iVefCode As Integer
    lDate As Long
End Type

Type AVAILNAMESINFO      '10-24-05, needed to maintain max avails for a vehicle for all selected stations
                        'not all stations carry all avails
    iCode As Integer
    sName As String * 20
    sTrafToAff         As String * 1         ' Determine if spots within the
                                             ' avail will be sent from Traffic
                                             ' to the Affiliate system. Y/N.
                                             ' Test for N
    sISCIExport        As String * 1         ' Determine if spots within the
                                             ' Avail will be included as part of
                                             ' the ISCI Export. Y/N. Test for N
    sAudioExport       As String * 1         ' Determine if spots within the
                                             ' Avail will be included as part of
                                             ' the audio delivery system (XDS,
                                             ' Wegener, Starguide). Y/N Test for
                                             ' N
    sAutomationExport  As String * 1         ' Determine if spots within the
                                             ' avail will be included as part of
                                             ' the automation export. Y/N. Test
                                             ' for N
End Type

Public tgAvailNamesInfo() As AVAILNAMESINFO

Type MEDIACODESINFO
    iCode As Integer
    sName As String * 6
End Type

Public tgMediaCodesInfo() As MEDIACODESINFO

Type STATSBYSTATION     '10-25-05
    sKey As String * 5          'station code for sorting by string
    iShfCode As Long            'station code
    iAnfCode As Integer         'named avail
    lFeedTime As Long            'actual feed time
    iAired As Integer           'total count aired (by vehicle & station & named avail)
    iNotReported As Integer     'total count not reported (by veh & station & named avail)
    iNotAired As Integer        'total count not aired (by veh & station  & named avail)
End Type

Type AVAILSBYVEHICLE
    sKey As String * 5          'string feed time for sorting
    lFeedTime As Long           'avail time
    iAnfCode As Integer         'named avail code
    iAnfInx As Integer          'index of entry in list box to retrieve description
End Type

Type VPF
    ivefKCode          As Integer            ' Internal code number for Vehicle
                                             ' Option
    sGTime             As String * 11        ' Time (Byte 0:Hund sec; Byte 1:
                                             ' sec.; Byte 2: min.; Byte 3:hour)
    sGMedium           As String * 1         ' Medium (T=TV; R=Radio; C=Cable;
                                             ' N=Radio Net; V=TV Network; P=Podcast Spots; M=Podcast CPM
                                             ' S=Radio ROS Net)
    iurfGCode          As Integer            ' User code
    sAdvtSep           As String * 1         ' Advertiser separation by B=Break;
                                             ' T=Time
    sCPLogo            As String * 3         ' Certificate of Perforance Logo
                                             ' File Name (Gxx.Bmp)
    lLgHd1CefCode      As Long               ' Header comment 1 from vehicle
                                             ' options
    lLgNmCefCode       As Long               ' Vehicle Log Name from vehicle
                                             ' options
    sUnsoldBlank       As String * 1         ' If blackout replace is missing:
                                             ' Y=Show as Unsold Blank line;
                                             ' N=Remove Event (spot).  Test for
                                             ' N
    sUsingFeatures1    As String * 1
    iSAGroupNo         As Integer
    'sGPriceStat        As String * 1         ' Calculate spot price statistics
    '                                         ' on schedule lines (Y=Yes, N=No)
    sOwnership As String * 1   'A=Owned-Network; B=Owned-Station; C=Unowned-Network; D=Unowned-Station.  Default A or Blank. Test for B, C and D
    sGGridRes          As String * 1         ' Default grid resolution (F=Full
                                             ' hour, H=Half hour, Q=Quarter
                                             ' hour)
    sGScript           As String * 1         ' Using scripts (Y=Yes, N=No)
    iGLocalAdj1        As Integer            ' Local time adjustment
    iGLocalAdj2        As Integer
    iGLocalAdj3        As Integer
    iGLocalAdj4        As Integer
    iGLocalAdj5        As Integer
    iGFeedAdj1         As Integer            ' Feed time adjustment
    iGFeedAdj2         As Integer
    iGFeedAdj3         As Integer
    iGFeedAdj4         As Integer
    iGFeedAdj5         As Integer
    sGZone1            As String * 3         ' Time zone
    sGZone2            As String * 3
    sGZone3            As String * 3
    sGZone4            As String * 3
    sGZone5            As String * 3
    iGV2Z1             As Integer            ' Versions displacement in minutes
    iGV2Z2             As Integer
    iGV2Z3             As Integer
    iGV2Z4             As Integer
    iGV2Z5             As Integer
    iGV3Z1             As Integer            ' Versions displacement in minutes
    iGV3Z2             As Integer
    iGV3Z3             As Integer
    iGV3Z4             As Integer
    iGV3Z5             As Integer
    iGV4Z1             As Integer            ' Versions displacement in minutes
    iGV4Z2             As Integer
    iGV4Z3             As Integer
    iGV4Z4             As Integer
    iGV4Z5             As Integer
    sGCSVer1           As String * 1         ' Show on cmml schd
                                             ' report:O=Original only; A=All
                                             ' version
    sGCSVer2           As String * 1
    sGCSVer3           As String * 1
    sGCSVer4           As String * 1
    sGCSVer5           As String * 1
    iGmnfNCode1        As Integer            ' Feed codes (test this field if
                                             ' conventional to determine type of
                                             ' conventional (Index 1 only:
                                             ' with:field>0 or without
                                             ' delivery:field is zero)
    iGmnfNCode2        As Integer
    iGmnfNCode3        As Integer
    iGmnfNCode4        As Integer
    iGmnfNCode5        As Integer
    sGBus1             As String * 2         ' Bus code
    sGBus2             As String * 2
    sGBus3             As String * 2
    sGBus4             As String * 2
    sGBus5             As String * 2
    sGSked1            As String * 2         ' Schedule
    sGSked2            As String * 2
    sGSked3            As String * 2
    sGSked4            As String * 2
    sGSked5            As String * 2
    sSVarComm          As String * 1         ' Agency commission variable (Y=
                                             ' Yes; N=No)
    sSCompType         As String * 1         ' Separate competitives (T=Time;
                                             ' B=Break; N=not back to back)
    sSCompLen          As String * 11        ' Time (Byte 0:Hund sec; Byte 1:
                                             ' sec.; Byte 2: min.; Byte 3:hour)
    iSBBLen            As Integer            ' BB spot length
    sSSellout          As String * 1         ' Sellout defined (U=Units; B=Both;
                                             ' T= 30 second units; M=Matching
                                             ' Units)
    sSOverBook         As String * 1         ' Allow overbooking of avails
                                             ' (Y=Yes; N=No)
    sSForceMG          As String * 1         ' Set spots moved outside contract
                                             ' limits as MG's (W=Always; A=Ask)
    '7/15/14
    sEmbeddedOrROS      As String * 1         ' Default Delivery for agreements:E= Embedded or R = ROS.  Test for E. Blank should be treated same as R (replaced vpfSPlaceNet)
    'sSNetContr         As String * 1         ' Unused-'Allow net spots into
    '                                         ' contract avails(Y=Yes, N=No)
    'sSContrNet         As String * 1         ' Unused-'Allow contract spots into
    '                                         ' network avails (Y=Yes, N=No)
    sWegenerExport     As String * 1         'Wegener Export(Y=Yes, N=No). Test for Y
    sOLAExport         As String * 1         'OLA (OnLine Affidavit) Export (Y=Yes, N=No). Test for Y
    sAvailNameOnWeb    As String * 1         ' Show Avail Names on Web (Y = yes;
                                             ' N=No)
    'sSPTA              As String * 1         ' Unused-'Keep spots with programs,
    '                                         ' time period or Ask(P=Program,
    '                                         ' T=Time, A=Ask)
    sUsingFeatures2    As String * 1         'Bit Mat (Right to Left): Bit 0=XDS Apply "Merge" ProgID
    sSAvailOrder       As String * 1         ' Order spots within break
    iSLen1             As Integer            ' Valid spot lengths
    iSLen2             As Integer
    iSLen3             As Integer
    iSLen4             As Integer
    iSLen5             As Integer
    iSLen6             As Integer
    iSLen7             As Integer
    iSLen8             As Integer
    iSLen9             As Integer
    iSLen10            As Integer
    iSLenGroup1        As Integer            ' Group # for lengths
    iSLenGroup2        As Integer
    iSLenGroup3        As Integer
    iSLenGroup4        As Integer
    iSLenGroup5        As Integer
    iSLenGroup6        As Integer
    iSLenGroup7        As Integer
    iSLenGroup8        As Integer
    iSLenGroup9        As Integer
    iSLenGroup10       As Integer
    sSCommCalc         As String * 1         ' Method of calculating salesperson
                                             ' commission (B=on billing; C=on
                                             ' collections)
    lMPSA60            As Long               ' Index to last PSA 60 filler
                                             ' schedule line scheduled
    lMPSA30            As Long               ' Index to last PSA 30 filler
                                             ' schedule line scheduled
    lMPSA10            As Long               ' Index to last PSA 10 filler
                                             ' schedule line scheduled
    lMPromo60          As Long               ' Index to last Promo 60 filler
                                             ' schedule line scheduled
    lMPromo30          As Long               ' Index to last Promo 30 filler
                                             ' schedule line scheduled
    lMPromo10          As Long               ' Index to last Promo 10 filler
                                             ' schedule line scheduled
    iMMFPSA1           As Integer            ' M-F: Max PSAs allowed per hour
    iMMFPSA2           As Integer
    iMMFPSA3           As Integer
    iMMFPSA4           As Integer
    iMMFPSA5           As Integer
    iMMFPSA6           As Integer
    iMMFPSA7           As Integer
    iMMFPSA8           As Integer
    iMMFPSA9           As Integer
    iMMFPSA10          As Integer
    iMMFPSA11          As Integer
    iMMFPSA12          As Integer
    iMMFPSA13          As Integer
    iMMFPSA14          As Integer
    iMMFPSA15          As Integer
    iMMFPSA16          As Integer
    iMMFPSA17          As Integer
    iMMFPSA18          As Integer
    iMMFPSA19          As Integer
    iMMFPSA20          As Integer
    iMMFPSA21          As Integer
    iMMFPSA22          As Integer
    iMMFPSA23          As Integer
    iMMFPSA24          As Integer
    iMSaPSA1           As Integer            ' Sa: Max PSAs allowed per hour
    iMSaPSA2           As Integer
    iMSaPSA3           As Integer
    iMSaPSA4           As Integer
    iMSaPSA5           As Integer
    iMSaPSA6           As Integer
    iMSaPSA7           As Integer
    iMSaPSA8           As Integer
    iMSaPSA9           As Integer
    iMSaPSA10          As Integer
    iMSaPSA11          As Integer
    iMSaPSA12          As Integer
    iMSaPSA13          As Integer
    iMSaPSA14          As Integer
    iMSaPSA15          As Integer
    iMSaPSA16          As Integer
    iMSaPSA17          As Integer
    iMSaPSA18          As Integer
    iMSaPSA19          As Integer
    iMSaPSA20          As Integer
    iMSaPSA21          As Integer
    iMSaPSA22          As Integer
    iMSaPSA23          As Integer
    iMSaPSA24          As Integer
    iMSuPSA1           As Integer            ' Su: Max PSAs allowed per hour
    iMSuPSA2           As Integer
    iMSuPSA3           As Integer
    iMSuPSA4           As Integer
    iMSuPSA5           As Integer
    iMSuPSA6           As Integer
    iMSuPSA7           As Integer
    iMSuPSA8           As Integer
    iMSuPSA9           As Integer
    iMSuPSA10          As Integer
    iMSuPSA11          As Integer
    iMSuPSA12          As Integer
    iMSuPSA13          As Integer
    iMSuPSA14          As Integer
    iMSuPSA15          As Integer
    iMSuPSA16          As Integer
    iMSuPSA17          As Integer
    iMSuPSA18          As Integer
    iMSuPSA19          As Integer
    iMSuPSA20          As Integer
    iMSuPSA21          As Integer
    iMSuPSA22          As Integer
    iMSuPSA23          As Integer
    iMSuPSA24          As Integer
    iMMFPr1            As Integer            ' M-F: Max Promos allowed per hour
    iMMFPr2            As Integer
    iMMFPr3            As Integer
    iMMFPr4            As Integer
    iMMFPr5            As Integer
    iMMFPr6            As Integer
    iMMFPr7            As Integer
    iMMFPr8            As Integer
    iMMFPr9            As Integer
    iMMFPr10           As Integer
    iMMFPr11           As Integer
    iMMFPr12           As Integer
    iMMFPr13           As Integer
    iMMFPr14           As Integer
    iMMFPr15           As Integer
    iMMFPr16           As Integer
    iMMFPr17           As Integer
    iMMFPr18           As Integer
    iMMFPr19           As Integer
    iMMFPr20           As Integer
    iMMFPr21           As Integer
    iMMFPr22           As Integer
    iMMFPr23           As Integer
    iMMFPr24           As Integer
    iMSaPr1            As Integer            ' Sa: Max Promos allowed per hour
    iMSaPr2            As Integer
    iMSaPr3            As Integer
    iMSaPr4            As Integer
    iMSaPr5            As Integer
    iMSaPr6            As Integer
    iMSaPr7            As Integer
    iMSaPr8            As Integer
    iMSaPr9            As Integer
    iMSaPr10           As Integer
    iMSaPr11           As Integer
    iMSaPr12           As Integer
    iMSaPr13           As Integer
    iMSaPr14           As Integer
    iMSaPr15           As Integer
    iMSaPr16           As Integer
    iMSaPr17           As Integer
    iMSaPr18           As Integer
    iMSaPr19           As Integer
    iMSaPr20           As Integer
    iMSaPr21           As Integer
    iMSaPr22           As Integer
    iMSaPr23           As Integer
    iMSaPr24           As Integer
    iMSuPr1            As Integer            ' Su: Max Promos allowed per hour
    iMSuPr2            As Integer
    iMSuPr3            As Integer
    iMSuPr4            As Integer
    iMSuPr5            As Integer
    iMSuPr6            As Integer
    iMSuPr7            As Integer
    iMSuPr8            As Integer
    iMSuPr9            As Integer
    iMSuPr10           As Integer
    iMSuPr11           As Integer
    iMSuPr12           As Integer
    iMSuPr13           As Integer
    iMSuPr14           As Integer
    iMSuPr15           As Integer
    iMSuPr16           As Integer
    iMSuPr17           As Integer
    iMSuPr18           As Integer
    iMSuPr19           As Integer
    iMSuPr20           As Integer
    iMSuPr21           As Integer
    iMSuPr22           As Integer
    iMSuPr23           As Integer
    iMSuPr24           As Integer
    sLLD               As String * 10        ' Last Log Date Byte 0:Day,
                                             ' 1:Month, followed by 2 byte year
    sLPD               As String * 10        ' Last Preliminary Log Date Byte
                                             ' 0:Day, 1:Month, followed by 2
                                             ' byte year
    slTimeZone         As String * 1         ' Time zone (E=Eastern; C=Central;
                                             ' M=Mountain; P=Pacific)
    sLDaylight         As String * 1         ' Daylight savings (Y=Yes; N=No)
    sLTiming           As String * 1         ' Using Log Timing (Y=Yes; N=No)
    sLAvailLen         As String * 1         ' Show length on unsold avails
                                             ' (Y=Yes; N=No)
    iSDLen             As Integer            ' Default spot length
    iFTPArfCode        As Integer            ' FTP address stored into ARF
                                             ' (Audio Stored address so that
                                             ' affiliate can get the Audio)
    sLShowCut          As String * 1         ' Cut/instruction #'s (C=show cut
                                             ' only; I=show instruction only;
                                             ' B=show both; N=show neither)
    sLTimeFormat       As String * 1         ' Show time (A=AM/PM; M=Military)
    slZone             As String * 1         ' Log/C of P/Play List Zone (A=All;
                                             ' E=EST; C=CST; M=MST; P=PST)
    sCPTitle           As String * 1         ' Certificate of Performance Title
                                             ' required(Y/N)
    sPrtCPStation      As String * 1         ' Print Certificate of Performance
                                             ' by Station (Y/N)
    iRnfPlayCode       As Integer            ' Default Play List (rnf code)
    sLastCP            As String * 10        ' Last CP date (Affiliate)
    sStnFdCart         As String * 1         ' Show Carts on Station Feed (Y/N)
    sStnFdXRef         As String * 1         ' Include in Vehicle cross
                                             ' reference on Station Feed (Y/N)
    sGenLog            As String * 1         ' Generate Logs (Y=Print on
                                             ' separate pages; N=None; L=Live
                                             ' Posting; M=Merge Sport into
                                             ' Pre-empt Vehicle)
    sCopyOnAir         As String * 1         ' Allow copy definition for Airing
                                             ' Vehicles (Y/N)
    sBillSA            As String * 1         ' Bill Airing Spots (Y/N); This
                                             ' field is used for partial
                                             ' simulcast vehicles
    sExpVehNo          As String * 2         ' Export Vehicle Number
    sExpBkCpyCart      As String * 1         ' Y=Show Cart # on Bulk Feed Export
    sExpHiCmmlChg      As String * 1         ' Export Commercial Change:
                                             ' Y=Highlight vehicle; N=Don't
                                             ' Highlight but show
    lAPenny            As Long               ' Highest penny variance amount
                                             ' (xxx.xx)
    iGV1Z1             As Integer            ' Versions displacement in minutes
    iGV1Z2             As Integer
    iGV1Z3             As Integer
    iGV1Z4             As Integer
    iGV1Z5             As Integer
    sFedZ1             As String * 1         ' Transmit: N=No; Y=Yes
    sFedZ2             As String * 1
    sFedZ3             As String * 1
    sFedZ4             As String * 1
    sFedZ5             As String * 1
    sGGroupNo          As String * 1         ' Group number- Used to group
                                             ' Vehicles together in bulk feed
    sLLastDateCpyAsgn  As String * 10        ' Last Date Copy Assigned Byte
                                             ' 0:Day, 1:Month, followed by 2
                                             ' byte year
    iESTEndTime1       As Integer            ' Daypart End Time of EST- in
                                             ' minutes
    iESTEndTime2       As Integer
    iESTEndTime3       As Integer
    iESTEndTime4       As Integer
    iESTEndTime5       As Integer
    iCSTEndTime1       As Integer            ' Daypart End Time of CST- in
                                             ' minutes
    iCSTEndTime2       As Integer
    iCSTEndTime3       As Integer
    iCSTEndTime4       As Integer
    iCSTEndTime5       As Integer
    iMSTEndTime1       As Integer            ' Daypart End Time of MST- in
                                             ' minutes
    iMSTEndTime2       As Integer
    iMSTEndTime3       As Integer
    iMSTEndTime4       As Integer
    iMSTEndTime5       As Integer
    iPSTEndTime1       As Integer            ' Daypart End Time of PST- in
                                             ' minutes
    iPSTEndTime2       As Integer
    iPSTEndTime3       As Integer
    iPSTEndTime4       As Integer
    iPSTEndTime5       As Integer
    sMapZone1          As String * 3         ' Zone which program code is to be
                                             ' remapped into Daypart
    sMapZone2          As String * 3
    sMapZone3          As String * 3
    sMapZone4          As String * 3
    sMapProgCode1      As String * 5         ' Program code to be mapped into
                                             ' different daypart
    sMapProgCode2      As String * 5
    sMapProgCode3      As String * 5
    sMapProgCode4      As String * 5
    iMapDPNo1          As Integer            ' Daypart number that program code
                                             ' is to be mapped into
    iMapDPNo2          As Integer
    iMapDPNo3          As Integer
    iMapDPNo4          As Integer
    sExpHiClear        As String * 1         ' Export Clearance Spot:
                                             ' Y=Highlight vehicle; N=Don't
                                             ' Highlight
    sExpHiDallas       As String * 1         ' Export Dallas Feed: Y=Highlight
                                             ' vehicle; N=Don't Highlight but
                                             ' show
    sExpHiPhoenix      As String * 1         ' Export Phoenix Feed: Y=Highlight
                                             ' vehicle; N=Don't Highlight but
                                             ' show
    sExpHiNY           As String * 1         ' Export New York Feed: Y=Highlight
                                             ' vehicle; N=Don't Highlight but
                                             ' show
    sBulkXFer          As String * 1         ' Export Bulk Feed Cross Reference:
                                             ' Y=Include vehicle; N=bypass
                                             ' vehicle
    sClearAsSell       As String * 1         ' In Export Clearance Spots treat
                                             ' vehicle as Selling (Y or N)
    sClearChgTime      As String * 1         ' In Export Clearance Spots change
                                             ' 11:59:55pm to 12:00:01AM (Y or N)
    sMoveLLD           As String * 1         ' Y=Allowed to move spots between
                                             ' todays date and Last Log date
    irnfLogCode        As Integer            ' Default Log Name (rnf code)
    irnfCertCode       As Integer            ' Default Log Certification Name
                                             ' (rnf code)
    iLNoDaysCycle      As Integer            ' Number of days in closing cycle
                                             ' (i.e. 1, 7 ,14)
    iLLeadTime         As Integer            ' Closing lead time (number of days
                                             ' required before closing that log
                                             ' should be generated, excludes sa,
                                             ' su)
    sShowTime          As String * 1         ' On invoice show air time as
                                             ' D=Daypart; S=Spot Time; A=Avail
                                             ' Time
    sEDICallLetter     As String * 4         ' EDI Call Letters
    sAccruedRevenue    As String * 20        ' Accrued Revenue G/L #
    sAccruedTrade      As String * 20        ' Accrued Trade G/L #
    sBilledRevenue     As String * 20        ' Billed Revenue G/L #
    sBilledTrade       As String * 20        ' Billed Trade G/L #
    sLCmmlSmmyAvNm     As String * 1         ' Commercial Summary Aval Filter
                                             ' (first letter of avail name or
                                             ' blank for all)
    lEDASWindow        As Long               ' EDAS Time window in seconds (Used
                                             ' in StarGuide and KenCast export)
    sKCGenRot          As String * 1         ' KenCast:  Genertae Rotation
                                             ' (Y/N).  If no, then generate
                                             ' rotation but don't include it in
                                             ' the envelope export.
    sExportSQL         As String * 1         ' Allow export to SQL server (Y/N)
    sAllowSplitCopy    As String * 1         ' Allow split copy (Y/N).  Only if
                                             ' Site Set Using Split Copy
    sUnunsed1          As String * 1
    sLastLog           As String * 10        ' Last CP date (Affiliate)
    iRnfSvLogCode      As Integer            ' Save to File Log Report Code
    iRnfSvCertCode     As Integer            ' Save to File C.P. Report Code
    iRnfSvPlayCode     As Integer            ' Save to File Other Report Code
    lLgFt1CefCode      As Long               ' Footer comment 1 from vehicle
                                             ' options
    lLgFt2CefCode      As Long               ' Footer comment 2 from vehicle
                                             ' options
    sStnFdCode         As String * 2         ' Station feed code- this must be
                                             ' unique for each station
    iProducerArfCode   As Integer            ' Produce ArfCode reference
    iProgProvArfCode   As Integer            ' Program Content Provider ArfCode
                                             ' Reference
    iCommProvArfCode   As Integer            ' Commercial Content Provider
                                             ' ArfCode reference
    sEmbeddedComm      As String * 1         ' Export Commercial Audio (Yes/No);
                                             ' Y = Export Comm Aud for Selected
                                             ' reference (not to stations);
                                             ' N=Export Comm Aud for Stations
    sARBCode           As String * 6         ' Arbitron code for vehicle
    lEMailCefCode      As Long               ' Station E-Mail Address
    sShowRateOnInsert  As String * 1         ' Show Spot Rate on Insertion
                                             ' Orders (Y/N)
    iAutoExptArfCode   As Integer            ' Automation Export Drive\Path
    iAutoImptArfCode As Integer              'Automation Import Path
    sWebLogSummary As String * 1             'Show web log spot summary (Y/N)
    sWebLogFeedTime As String * 1            'Show on web log spot feed time (Y/N)
    'sUnused As String * 8                   ' Unused
    sRadarCode As String * 5                 'Radar Import codes
    sEDIBand As String * 1                   'Frequency Band (A or F)
    iInterfaceID As Integer
End Type

Public Const LOGSJOB = 6

Public Const CSI_MSG_NONE = 0    'No error
Public Const CSI_MSG_NOSHOW = 1  'Error- but don't show message as it was prevoiusly shown
Public Const CSI_MSG_PARSE = 2   'Parse error
Public Const CSI_MSG_POPREQ = 0  'List box populated OK
Public Const CSI_MSG_NOPOPREQ = 3   'List box didn't require population

Type TITLEINFO
    iCode As Integer
    sTitle As String * 40
End Type

Public tgTitleInfo() As TITLEINFO

Type OWNERINFO                      ' 2-23-06
    lCode As Long                   ' artt code
    sName As String * 60            ' owner name
    sPhone As String * 20           ' owner phone #
    sFax As String * 20             ' owner fax #
    sEmail As String * 70           ' owner email
End Type

Public tgOwnerInfo() As OWNERINFO

'Array for station information report
Type STATIONARRAY
    sAddress As String * 40         'address 1-4, city/state, zip or country
    sPersonInfo As String * 240     'phone #, fax, email
End Type

' Array for FMT (Station Format) information
Type FORMATINFO
    lCode      As Long              ' Internal code
    sName      As String * 60       ' Format Name
    sGroupName As String * 10
    iUstCode   As Integer           ' Pointer to the Ust table
End Type
Public tgFormatInfo() As FORMATINFO

'******************************************************************************
' VFF_Vehicle_Features Record Definition
'
'******************************************************************************
Type VEHICLEFEATURESINFO
    iVefCode              As Integer         ' Vehicle reference
    sGroupName            As String * 10     ' Group Name
    sWegenerExportID      As String * 10     ' Wegener export vehicle ID
    sOLAExportID          As String * 10     ' OLA export vehicle ID
    iLiveCompliantAdj     As Integer         ' Live pledge Plus/Minus window
                                             ' adjustment in minutes.
    sXDXMLForm            As String * 1      ' X-Digital XML Avail Cue Form: S= Avail ID by Hour and Break #, A=Avail ID by Hour #, Break # and Position; R=Replace Cue tag with Generic ISCI
    sXDISCIPrefix         As String * 6      ' HB or HBP: ISCI Prefix
    sXDSaveCF             As String * 1      ' X-Digital File Delivery Save on
                                             ' Compact Flash Drive (Y/N). Test
                                             ' for Y
    sXDSaveHDD            As String * 1      ' X-Digital File Delivery Save on
                                             ' Hard Drive (Y/N). Test for Y
    sXDSaveNAS            As String * 1      ' X-Digital File Delivery Save on
                                             ' Netwoork Attached Storage device
                                             ' (External) (Y/N). Test for Y
    sXDProgCodeID         As String * 8
    sPledgeVsAir          As String * 1      ' Include vehicle in the Pledge vs Air CSV Affiliate Export (Y/N).  Test for Y
    sMergeAffiliate       As String * 1      ' In Affiliate: M=Merge into Log
                                             ' vehicle; S=Separate from Log
                                             ' vehicles. Test for S.
    sMergeTraffic         As String * 1
    sMergeWeb             As String * 1      ' In Affiliate Affidavit: M=Merge
                                             ' into Log vehicle; S=Separate from
                                             ' Log vehicles; C=Choose (Only one
                                             ' airs).  Test for S or C.
    sWebName              As String * 40     ' Replacement name to appear on Web
                                             ' if defined for all vehicles. For
                                             ' Merged vehicles, show name within
                                             ' Log if defined.
    sPledgeByEvent        As String * 1      ' Pledge defined by Event (Y=Yes, N=No). Test for Y
    sIPumpEventTypeOV     As String * 2      ' iPump Event Type ID override of
                                             ' the Media Code Event Type. If
                                             ' blank, then use Media Code Event
                                             ' Type
    sExportIPump          As String * 1      ' Export Wegener-iPump(Y/N). Test
                                             ' for Y
    sXDSISCIPrefix        As String * 6      ' ISCI Form: ISCI Prefix
    sXDSSaveCF            As String * 1      ' X-Digital File Delivery Save on
                                             ' Compact Flash Drive (Y/N). Test
                                             ' for Y
    sXDSSaveHDD           As String * 1      ' X-Digital File Delivery Save on
                                             ' Hard Drive (Y/N). Test for Y
    sXDSSaveNAS           As String * 1      ' X-Digital File Delivery Save on
                                             ' Netwoork Attached Storage device
                                             ' (External) (Y/N). Test for Y
    sSentToXDS            As String * 1      ' Sent to XDS? Y = yes, N = No, M = Modified
    sExportJelli          As String * 1      ' Export Jelli (Y/N). Test for Y
    sMGsOnWeb             As String * 1      ' Allow MGs on Web (Y/N). Test for Y
    sReplacementOnWeb     As String * 1      ' Allow Replacement on Web (Y/N). Test for Y
    sStationComp          As String * 1      ' Station Compensation (Y/N). Test for Y
    sHonorZeroUnits       As String * 1      ' For airing vehicle honor zero unit when generating the Log and creating Pledge avails (Y/N). Test for Y
    sHideCommOnWeb        As String * 1      ' Hide comments on the web (Y/N). Test for Y
    '10933
    sXDEventZone          As String * 1      ' Cue Code by Zone Y or N(blank)
End Type

Public tgVffInfo() As VEHICLEFEATURESINFO


Type VEF
    iCode              As Integer            ' Internal code number for Vehicle
    sName              As String * 40        ' Name
    sAddr1             As String * 25        ' Address
    sAddr2             As String * 25
    sAddr3             As String * 25
    sPhone             As String * 14        ' Phone plus extension
    sFax               As String * 10        ' Fax number
    sUnused1           As String * 1         ' Unused was OwnRep: O=Owned,
                                             ' R=Repped
    sDialPos           As String * 5         ' Dial Position
    lPvfCode           As Long
    iReallDnfCode      As Integer            ' Reallocation Book Name Code
    sUpdateRvf1        As String * 1         ' Participant-Update Rvf(Y) or
                                             ' Phf(N) flag
    sUpdateRvf2        As String * 1
    sUpdateRvf3        As String * 1
    sUpdateRvf4        As String * 1
    sUpdateRvf5        As String * 1
    sUpdateRvf6        As String * 1
    sUpdateRvf7        As String * 1
    sUpdateRvf8        As String * 1
    iCombineVefCode    As Integer            ' Combine this vehicle (vefCode)
                                             ' with referenced vehicle
                                             ' (vefCombineVefCode) to create Log
    iMnfHubCode        As Integer            ' Multi-Name Hub Code (mnfType = U)
    iTrfCode           As Integer            ' Tax Rate reference
    sType              As String * 1         ' Type:C=Conventional;S=Selling;A=A
                                             ' iring;L=Log;V=Virtual;T=Simulcast
                                             ' ;R=Rep;N=NTR;P=Package;G=Sport; I=Import
                                             ' Note:Conve
                                             ' ntional w/Vpf.iGMnfNCode(1)>0 is
                                             ' included w/airing for creating
                                             ' delivery, TheLog(L) veh is used
                                             ' to combine conventional veh that
                                             ' ref Log veh into 1 LOG.
    sCodeStn           As String * 5         ' Station vehicle code
    iVefCode           As Integer            ' Combination vehicle code (this
                                             ' code number is stored into ODF as
                                             ' the vehicle code if not zero)
    iUnused2           As Integer            ' Unused
    iProdPct           As Integer            ' If sOwnRep = R, then producer's %
                                             ' (xx.xx)
    iProdPct2          As Integer
    iProdPct3          As Integer
    iProdPct4          As Integer
    iProdPct5          As Integer
    iProdPct6          As Integer
    iProdPct7          As Integer
    iProdPct8          As Integer
    sState             As String * 1         ' A=Active; D=Dormant
    imnfGroup          As Integer            ' Participant (Vehicle Group) code
                                             ' number
    imnfGroup2         As Integer
    imnfGroup3         As Integer
    imnfGroup4         As Integer
    imnfGroup5         As Integer
    imnfGroup6         As Integer
    imnfGroup7         As Integer
    imnfGroup8         As Integer
    iSort              As Integer            ' Sort code
    idnfCode           As Integer            ' Latest rating book code
    imnfDemo           As Integer            ' Prime Demo Code
    imnfSSCode1        As Integer            ' Sales Source code number
    imnfSSCode2        As Integer
    imnfSSCode3        As Integer
    imnfSSCode4        As Integer
    imnfSSCode5        As Integer
    imnfSSCode6        As Integer
    imnfSSCode7        As Integer
    imnfSSCode8        As Integer
    sExportRAB         As String * 1         ' Export RAB: Y=Yes; N=No. Test for Y
    lVsfCode           As Long               ' Virtual vehicle or Standard
                                             ' Package Vehicle
    lRateAud           As Long               ' Rating(xx.xx) or
                                             ' Audience(xxxxxx). Stored as
                                             ' xxxxxx.xx
    lCPPCPM            As Long               ' CPP(xxxxx) or (CPM(xx.xx). Stored
                                             ' as xxxxx.xx)
    lYearAvails        As Long               ' Avails for Year (xxxxx)
    iPctSellout        As Integer            ' % sellout (xxx)
    iMnfVehGp2         As Integer            ' Vehicle Group Set # 2
    iMnfVehGp3Mkt      As Integer            ' Market- Vehicle Group Set # 3
    iMnfVehGp4Fmt      As Integer            ' Format- Vehicle Group Set # 4
    iMnfVehGp5Rsch     As Integer            ' Research- Vehicle Group Set # 5
    iMnfVehGp6Sub      As Integer            ' SubCompany- Vehicle Group Set # 6
    iMnfVehGp7         As Integer            ' Unsed Vehicle Group Set # 7
    iSSMnfCode         As Integer            ' Sales Source from latest Participant table and from owner row (1st row of table)
    sStdPrice          As String * 1         ' Standard Package Price
                                             ' (Distribute price to hidden
                                             ' lines: R= Rate; A= Audience; P=
                                             ' Percent; S= Spot Count)
    sStdInvTime        As String * 1         ' Standard Package Invoice Generate
                                             ' Time (A=Real-use hidden line
                                             ' times; O=Virtual- use Package
                                             ' line times; E= Price for package
                                             ' line)
    sStdAlter          As String * 1         ' Standard Package Alter Hidder
                                             ' Flag (Y=Allow Hidden Lines to be
                                             ' Altered; N= Only Package Line can
                                             ' be Altered; C=N except comment)
    iStdIndex          As Integer            ' Standard Package Dollar index
    sStdAlterName      As String * 1         ' Standard Package: Allow Name to be altered (N=No; Y or Blank = Yes. Test for N)
    iRemoteID          As Integer            ' Remote ID (Note: Unique ID=Remote
                                             ' ID + AutoCode)
    iAutoCode          As Integer            ' Auto Incr Code (Note: Unique
                                             ' ID=Remote ID + AutoCode)
    sExtUpdateRvf1     As String * 1         ' Participant-External Update
                                             ' Rvf(Y) or Phf(N) flag
    sExtUpdateRvf2     As String * 1
    sExtUpdateRvf3     As String * 1
    sExtUpdateRvf4     As String * 1
    sExtUpdateRvf5     As String * 1
    sExtUpdateRvf6     As String * 1
    sExtUpdateRvf7     As String * 1
    sExtUpdateRvf8     As String * 1
    sStdSelCriteria    As String * 1         ' Standard package Vehicle
                                             ' selection: A=All; M=Matching
                                             ' dayparts
    sStdOverrideFlag   As String * 1         ' Standard Package override Flag:
                                             ' S=Hidden line as subset; A=Set
                                             ' All overrides event if not subset
    sContact           As String * 40        ' Contact name  9/10/02
End Type


Type MNF
    iCode              As Integer            ' Internal code number for
                                             ' multi-name
    sType              As String * 1         ' A=Announcer;B=BusCat;C=ProductPro
                                             ' tection;D=Demo;E=Genre;F=Research
                                             ' ;G=SalesRegion;H=VehGroup;I=NTR
                                             ' Type;J=Terms;K=
                                             ' Seg;L=Lang;M=MissedReason;N=FeedT
                                             ' ype;O=Compet;P=Pot;R=RevSet;S=Sal
                                             ' eSource;T=SaleTeam;V=InvSort;X=Ex
                                             ' clus.;Y=TransType;Z=SportTeam
    sName              As String * 20        ' Name
    sRPU               As String * 5    'Currency * 9       ' Rate as money (Pack Decimal
                                             ' Number, 2 places after dec point)
    sUnitType          As String * 6         ' Unit type with "per" removed
    sSSComm            As String * 4    'Type 5 Not Supported * 7    ' Commission (Pack Decimal
                                             ' Number, 4 places after dec point)
    iMerge             As Integer            ' Merge code number
    iGroupNo           As Integer            ' If "A": Group number or if "S":
                                             ' sales origin (1=Local;
                                             ' 2=Regional; 3=National)
    sCodeStn           As String * 5         ' "C" and "R" codes assigned by
                                             ' station
    iRemoteID          As Integer            ' Unique ID = Remote ID + AutoCode
    iAutoCode          As Integer            ' Unique ID = Remote ID + AutoCode
    sSyncDate          As String * 10        ' Sync Date (from Master Server or
                                             ' Remote System)
    sSyncTime          As String * 11        ' Sync Time (from Master Server or
                                             ' Remote System)
    sUnitsPer          As String * 15        ' NTR Type (I) = Units Per ------
    lCost              As Long               ' Type I (NTR): Acquisition Cost
                                             ' (xxxxxx.xx)
    sUnused            As String * 1
End Type

Type TEAMINFO
    iCode As Integer
    sName As String * 20
    sShortForm As String * 6
End Type

Public tgTeamInfo() As TEAMINFO

Type LANGINFO
    iCode As Integer
    sName As String * 20
    sEnglish As String * 1
End Type

Public tgLangInfo() As LANGINFO

Type TIMEZONEINFO
    iCode As Integer
    sName As String * 40
    sCSIName As String * 3
    sGroupName As String * 10
End Type

Public tgTimeZoneInfo() As TIMEZONEINFO

Type STATEINFO
    iCode As Integer
    sName As String * 40
    sPostalName As String * 2
    sGroupName As String * 10
End Type

Public tgStateInfo() As STATEINFO


Type SUBTOTALGROUPINFO
    iCode As Integer
    sName As String * 20
End Type

Public tgSubtotalGroupInfo() As SUBTOTALGROUPINFO

Type GHF
    lCode As Long            ' Auto Increment Code
    iVefCode As Integer         ' Vehicle Reference Code
    sSeasonName           As String * 20     ' Season Name
    sSeasonStartDate As String     ' Season Start Date
    sSeasonEndDate As String        ' Season End Date
    iNoGames As Integer         ' Number of Games
    sUnused As String * 10     ' Unused
End Type

Type GSF
    lCode              As Long               ' Auto Reference Code
    lGhfCode           As Long               ' Game Header Reference Code
    iVefCode           As Integer            ' Vehicle Reference Code
    iGameNo            As Integer            ' Game Number
    sFeedSource        As String * 1         ' Feed Source: Home (H) or Visiting
                                             ' (V)
    iLangMnfCode       As Integer            ' Language MultiName Reference Code
    iVisitMnfCode      As Integer            ' Visiting Team Name MultiName
                                             ' Reference Code
    iHomeMnfCode       As Integer            ' Home team Name MultiName
                                             ' Reference Code
    lLvfCode           As Long               ' Library Reference Code
    sAirDate           As String * 10        ' Air Date of Game
    sAirTime           As String * 11        ' Air Time of Game
    iAirVefCode        As Integer            ' Air Vehicle Reference Code (used
                                             ' to preempt spots from airing
                                             ' vehicle).
    sGameStatus        As String * 1         ' Game Date and Time status
                                             ' (T=Tentatiove; F=Firm;
                                             ' C=Cancelled)
    sLiveLogMerge      As String * 1         ' If Vehicle Option (vpfGenLog) set to A (Ask): L=Live Log+Preempt; M=Merge+Preempt
    sXDSProgCodeID     As String * 8         ' XDS Program Code ID to be used if vffXDProgCodeID is defined as Event.
    sBus As String * 20      'Engineering Export Bus value for Sports so that Engineering Links are not required.  Allow multi-bus to be separated with a ; as a separator
    iSubtotal1MnfCode     As Integer         ' Event Subtotal 1 MnfCode
                                             ' reference
    iSubtotal2MnfCode     As Integer         ' Event Subtotal 2 MnfCode
                                             ' reference
    sUnused               As String * 6      ' Unused
End Type

Type ADF
    iCode              As Integer            ' Internal code number for
                                             ' advertiser
    sName              As String * 30        ' Name
    sAbbr              As String * 7         ' Abbreviation
    sProd              As String * 35        ' Product name
    iSlfCode           As Integer            ' Salesperson code number
    iAgfCode           As Integer            ' Agency code number
    sBuyer             As String * 20        ' Buyers name
    sCodeRep           As String * 10        ' Rep advertiser Code
    sCodeAgy           As String * 10
    sCodeStn           As String * 10        ' Station advertiser Code
    imnfComp1          As Integer            ' Competitive code
    imnfComp2          As Integer
    imnfExcl1          As Integer            ' Program Exclusions code
    imnfExcl2          As Integer
    sCppCpm            As String * 1         ' P=CPP; M=CPM; N=N/A (was
                                             ' sPriceType)
    sDemo1             As String * 6         ' First-four Demo target
    sDemo2             As String * 6
    sDemo3             As String * 6
    sDemo4             As String * 6
    imnfDemo1          As Integer            ' Mnf Demo copy number
    imnfDemo2          As Integer
    imnfDemo3          As Integer
    imnfDemo4          As Integer
    lTarget1           As Long               ' CPP or CPM target (xxxxx.xx)
    lTarget2           As Long
    lTarget3           As Long
    lTarget4           As Long
    sCreditRestr       As String * 1         ' N=No Restrictions; L=Credit
                                             ' Limit; W=Cash in Advance each
                                             ' Week;M=Cash in Advance each
                                             ' Month; T=Cash in Advance
                                             ' Quarterly;P=Prohibit New Orders
    lCreditLimit       As Long               ' Credit limit amount (xxxxxxx.xx)
    sPaymRating        As String * 1         ' 0=Quick pay; 1=Normal pay; 2=
                                             ' Slow pay; 3= Difficult; 4=in
                                             ' Collection
    sISCI              As String * 1         ' Y=Yes; N=No
    imnfSort           As Integer            ' Invoice sorting code
    sBilAgyDir         As String * 1         ' A=Bill Agency; D=Bill Advertiser
                                             ' Directly
    sCntrAddr1         As String * 40        ' Contract address
    sCntrAddr2         As String * 40
    sCntrAddr3         As String * 40
    sBillAddr1         As String * 40        ' Billing address
    sBillAddr2         As String * 40
    sBillAddr3         As String * 40
    iarfLkCode         As Integer            ' Lock box code
    sPhone             As String * 14        ' Phone number and extension
    sFax               As String * 10        ' Fax number
    iarfContrCode      As Integer            ' Contract EDI service code number
    iarfInvCode        As Integer            ' Invoice EDI service code number
    sCntrPrtSz         As String * 1         ' W=Wide; N= Narrow
    iTrfCode           As Integer            ' Tax rate reference
    sCrdApp            As String * 1         ' Direct Only: Credit Approved
                                             ' (R=Requires Credit Check;
                                             ' A=Approved; D=Denied)
    sCrdRtg            As String * 5         ' Credit rating (AAA, B,..)
    ipnfBuyer          As Integer            ' Personnel Buyer Code Number
    ipnfPay            As Integer            ' Personnel Payable Clerk Code
                                             ' Number
    iPct90             As Integer            ' % Over 90 days (xxx)
    sCurrAR            As String * 6    'Currency * 11      ' Current A/R (xx,xxx,xxx.xx)
    sUnbilled          As String * 6    'Currency * 11      ' Unbilled amount (xx,xxx,xxx.xx);
                                             ' Computed in SetCrdit:; CrditRestr
                                             ' Value; P, N,
                                             ' L=UnbilledSpotsPlusPrjctdSpotsFor
                                             ' Spf.iRNoWks;W=UnbilledSpots+1wAct
                                             ' lSpts;M=UnbilledSpots+4wActlSpts;
                                             ' T=UnbilledSpots+13wActlSpts;Note:
                                             ' TodaysDatesIncludedW/Unbilled(pos
                                             ' tLog)
    sHiCredit          As String * 6    'Currency * 11      ' High credit amount
                                             ' (xx,xxx,xxx.xx)
    sTotalGross        As String * 6    'Currency * 11      ' Total gross (xx,xxx,xxx.xx)
    sDateEntrd         As String * 10        ' Date entered
    iNSFChks           As Integer            ' Number returned checks
    sDateLstInv        As String * 10        ' Date last invoiced
    sDateLstPaym       As String * 10        ' Date of last payment
    iAvgToPay          As Integer            ' Average number of days to pay
    iLstToPay          As Integer            ' Last number of days to pay
    iNoInvPd           As Integer            ' Number of invoices paid
    sNewBus            As String * 1         ' Y=Yes; N=No (when advertiser
                                             ' added- set to Y; When first
                                             ' contract reference this
                                             ' advertiser, this field is set to
                                             ' N
    sEndDate           As String * 10        ' Latest contract date
                                             ' referencinInvoice sorting codeg
                                             ' this advertiser
    iMerge             As Integer            ' Merge code number
    iurfCode           As Integer            ' Last user to modify record
    sState             As String * 1         ' A=Active; D=Dormant
    sCrdAppDate        As String * 10        ' Date credit approval changed
    sCrdAppTime        As String * 11        ' Time credit approval changed
    sPkInvShow         As String * 1         ' Package Invoice air times as
                                             ' D=Daypart; T=Time
    lGuar              As Long               ' %Guar(xxx) or Grimp
                                             ' Guarantee(xxxxxx) from contracts
    iLastYearNew          As Integer         ' Last Year advertiser set as New
    iLastMonthNew         As Integer         ' Last Month advertiser treated as New
    sBkoutPoolStatus      As String * 1      ' Blackout Pool Status (A=Active; D=Deactivated; N or Blank = Not part of Pool.  Test for A or D.
    sUnused2 As String * 5       'Not Used
    sRateOnInv         As String * 1         ' Show Rate on Invoice (Y/N).  Test
                                             ' for N
    iMnfBus            As Integer            ' Business Category from contract
    iUnused1           As Integer            ' Unused
    sAllowRepMG        As String * 1         ' Allow Rep Posted spots to be
                                             ' counted as MG (Y/N)
    sBonusOnInv        As String * 1         ' Show Bonus spots on invoices
                                             ' (Y/N)
    sRepInvGen         As String * 1         ' Rep Invoice Generated Internally
                                             ' (I) or Externally (E)
    iMnfInvTerms       As Integer            ' Terms to show on invoive (Stored
                                             ' with mnfType = "J" and
                                             ' mnfUnitType = "T")
    sPolitical         As String * 1         ' Airing Political spots (Y/N)
    sAddrID            As String * 9         ' Used for type ahead and sorting
                                             ' when multi-address exist for
                                             ' direct advertiser
End Type

Type CIF
    lCode              As Long               ' Internal code number for Copy
                                             ' Inventory
    imcfCode           As Integer            ' Media code number
    sName              As String * 5         ' Name
    sCut               As String * 1         ' Cut
    sReel              As String * 10        ' Reel
    iLen               As Integer            ' Copy length
    ietfCode           As Integer            ' Event Type code number
    ienfCode           As Integer            ' Event Name code number
    iAdfCode           As Integer            ' Advertiser code number
    lCpfCode           As Long               ' Product/Agency ID code number
    imnfComp1          As Integer            ' Competitive code
    imnfComp2          As Integer
    imnfAnn            As Integer            ' Announcer code
    sHouse             As String * 1         ' Tape in House Y/N
    sCleared           As String * 1         ' Tape cleared Y/N
    lCSFCode           As Long               ' Live script code
    iTimes             As Integer            ' Number of times aired
    sCDisp             As String * 1         ' Cart Disposition (N=N/A; S=Save;
                                             ' P=Purge; A=Ask after Expired)
    sTDisp             As String * 1         ' Tape Disposition (N=N/A;
                                             ' R=Return; D=Destroy; A=Ask after
                                             ' Expired)
    sPurged            As String * 1         ' Purge flag (A=Active; P=Purged;
                                             ' H=History)
    sPurgeDate         As String * 10        ' Purged Date; Date Byte 0:Day,
                                             ' 1:Month, followed by 2 byte year
    sEntryDate         As String * 10        ' Date entered; Date Byte 0:Day,
                                             ' 1:Month, followed by 2 byte year
    sUsedDate          As String * 10        ' Last Used Date; Date Byte 0:Day,
                                             ' 1:Month, followed by 2 byte year
    sRotStartDate      As String * 10        ' Earliest Rotation date using this
                                             ' inventory; Date Byte 0:Day,
                                             ' 1:Month, followed by 2 byte year
    sRotEndDate        As String * 10        ' Latest Rotation date using this
                                             ' inventory; Date Byte 0:Day,
                                             ' 1:Month, followed by 2 byte year
    iurfCode           As Integer            ' Last user who modified inventory
    sPrint             As String * 1         ' Printed: N=New or Not Printed;
                                             ' P=Printed
    iLangMnfCode       As Integer            ' Language Reference to MNF
    sInvSentDate       As String * 10        ' Inventory sent to vCreative date
    sUnused            As String * 4
End Type

Type CPF
    lCode              As Long               ' Internal code number for copy
                                             ' Product/Agency
    sName              As String * 35        ' Name
    sISCI              As String * 20        ' Agency ISCI code
    sCreative          As String * 30        ' Creative title
    sRotEndDate        As String * 10        ' Latest Rotation date using this
                                             ' inventory; Date Byte 0:Day,
                                             ' 1:Month, followed by 2 byte year
    lsifCode           As Long               ' Short Title used with cmml sch &
                                             ' bulk feed
End Type

Type CPTT
    lCode                 As Long
    lAtfCode              As Long            ' Agreement code (att above)
    iShfCode              As Integer         ' Station Code
    iVefCode              As Integer         ' Vehicle Code
    sCreateDate           As String * 10     ' Date record created date
    sStartDate            As String * 10     ' Date on log
    sReturnDate           As String * 10     ' Date CP Reurned (Todays date when
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

Type ANF
    iCode              As Integer            ' Internal code number for Avail
                                             ' Name
    sName              As String * 20        ' Identification name
    sSustain           As String * 1         ' Sustain allowed (Y/N)
    sState             As String * 1         ' A=Active; D=Dormant
    sSponsorship       As String * 1         ' Can the avail be sponsored (Y/N)
    iMerge             As Integer            ' Merge code number
    iRemoteID          As Integer            ' Unique ID = Remote ID + AutoCode
    iAutoCode          As Integer            ' Unique ID = Remote ID + AutoCode
    sBookLocalFeed     As String * 1         ' Book: L= Local Spots Only, F =
                                             ' Feed Spots Only; B = Both
    sRptDefault           As String * 1      ' Combo report default (Y/N)
    iSortCode             As Integer         ' Sort Code
    sTrafToAff            As String * 1      ' Determine if spots within the
                                             ' avail will be sent from Traffic
                                             ' to the Affiliate system. Y/N.
                                             ' Test for N
    sISCIExport           As String * 1      ' Determine if spots within the
                                             ' Avail will be included as part of
                                             ' the ISCI Export. Y/N. Test for N
    sAudioExport          As String * 1      ' Determine if spots within the
                                             ' Avail will be included as part of
                                             ' the audio delivery system (XDS,
                                             ' Wegener, Starguide). Y/N Test for
                                             ' N
    sAutomationExport     As String * 1      ' Determine if spots within the
                                             ' avail will be included as part of
                                             ' the automation export. Y/N. Test
                                             ' for N
    sFillRequired         As String * 1      ' Does the break need to be filled
                                             ' (Y or Blank=Yes; N=No; A=Only if
                                             ' Partially booked and treat
                                             ' Adjacent breaks a one break).
                                             ' Test for N.
    sUnused               As String * 1      ' Feed End Time
End Type

'Type AVAILNAMEINFO
'    iCode              As Integer            ' Internal code number for Avail
'                                             ' Name
'    sName              As String * 20        ' Identification name
'End Type

'Public tgAvailNameInfo() As AVAILNAMEINFO

Type LSTSPOT
    lLstCode As Long
    iAdfCode As Integer
    lAvailTime As Long
End Type

Type CLEASRIMPORTINFO
    iVefCode As Integer
    lClearDate As Long
    lgsfCode As Long
End Type

'Radar Information
Type RETINFO
    sProgCode          As String * 3         ' Program code (like 001, 002,...)
    lStartTime         As Long               ' Extraction Start Time
    lEndTime           As String * 11        ' Extraction End Time
    sDayType           As String * 2         ' Day Type (MF=M-F; Sa=Sa; Su=Su)
    '4/25/19
    iRepeatCount       As Integer            ' Used if output by program code
End Type

'Radr Export Info
Type RADAREXPORTINFO
    lSdfCode As Long
    iInitUnitCount As Integer
    iRepeatCount As Integer
End Type

'
' Added 08-22-03  J. Dutschke
' FTP Prototype information
' Added to support the Web Spots Application
'
Public Declare Function InternetOpen _
   Lib "wininet.dll" Alias "InternetOpenA" ( _
   ByVal sAgent As String, _
   ByVal nAccessType As Long, _
   ByVal sProxyName As String, _
   ByVal sProxyBypass As String, _
   ByVal nFlags As Long) As Long
   
Public Declare Function InternetConnect _
   Lib "wininet.dll" Alias "InternetConnectA" ( _
   ByVal hInternetSession As Long, _
   ByVal sServerName As String, _
   ByVal nServerPort As Integer, _
   ByVal sUserName As String, _
   ByVal sPassword As String, _
   ByVal nService As Long, _
   ByVal dwFlags As Long, _
   ByVal dwContext As Long) As Long

Public Declare Function FtpGetFile _
   Lib "wininet.dll" Alias "FtpGetFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszRemoteFile As String, _
   ByVal lpszNewFile As String, _
   ByVal fFailIfExists As Boolean, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal dwFlags As Long, _
   ByVal dwContext As Long) As Boolean

Public Declare Function FtpPutFile _
   Lib "wininet.dll" Alias "FtpPutFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszLocalFile As String, _
   ByVal lpszRemoteFile As String, _
   ByVal dwFlags As Long, _
   ByVal dwContext As Long) As Boolean

Public Declare Function FtpDeleteFile _
   Lib "wininet.dll" Alias "FtpDeleteFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszFileName As String) As Boolean

Public Declare Function FtpRenameFile _
   Lib "wininet.dll" Alias "FtpRenameFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszExisting As String, _
   ByVal lpszNewName As String) As Boolean

Public Declare Function FtpFindFirstFile _
   Lib "wininet.dll" Alias "FtpFindFirstFileA" ( _
   ByVal hFtpSession As Long, _
   ByVal lpszSearchFile As String, _
   ByRef lpFindFileData As WIN32_FIND_DATA, _
   ByVal dwFlags As Long, _
   ByVal dwContent As Long) As Long
   
Private Declare Function InternetFindNextFile _
   Lib "wininet.dll" Alias "InternetFindNextFileA" ( _
   ByVal hFind As Long, _
   ByRef lpvFindData As WIN32_FIND_DATA) As Long
   
Public Declare Function InternetCloseHandle _
   Lib "wininet.dll" (ByVal hInet As Long) As Integer

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * 260
   cAlternate As String * 14
End Type

Type EDF
    lCode              As Long
    lEsfCode           As Long               ' Link to ESF - Export Summary
    lAttCode           As Long               ' Link to Att - Agreement File
    iVefCode           As Integer            ' Link to Vef - Vehicle Code
    iShttCode          As Integer            ' Link to Shtt - Station File
    sExpTime           As String * 11        ' Time the spots were created for
                                             ' export
    sExpDate           As String * 10        ' Date the spots were created fort
                                             ' export
    lTtlAdd            As Long               ' Total number of Add records for
                                             ' that station/vehicle combination
    lTtlDel            As Long               ' Total number of Delete records
                                             ' for that station/vehicle
                                             ' combination
    lTtlRec            As Long               ' Total number of Add and Delete
                                             ' records for that station/vehicle
                                             ' combination
    sUser              As String * 10        ' The users name that they logged
                                             ' in with
    sUnused            As String * 40        ' Extra space
End Type


Type ESF
    lCode              As Long
    lTtlAdd            As Long               ' Total number of add records for
                                             ' the entire export
    lTtlDel            As Long               ' Total number of delete records
                                             ' for the entire export
    lTtlAddDel         As Long               ' Total number of add and delete
                                             ' records for the entire export
    lTtlEmails         As Long               ' Total number of emails sent for
                                             ' the entire export
    lTtlHdrs           As Long               ' Total number of headers sent out
                                             ' the entire export
    lTtlComments       As Long               ' Total number of spot comments
                                             ' sent out
    sStartTime         As String * 11        ' The time the export was started
    sStartDate         As String * 10        ' The date the export was started
    sEndTime           As String * 11        ' The time the export ended
    sEndDate           As String * 10        ' The date the export ended
    sElspdTime         As String * 11        ' The total elapsed time the export
                                             ' took start to finish
    sExpDate           As String * 10        ' The date the user entered for the
                                             ' start date of the export.
    iNumDays           As Integer            ' The number of days the user
                                             ' entered to export.
    sUser              As String * 10        ' The name of the user as they are
                                             ' logged in as
    sMachine           As String * 10        ' The name of the machine that was
                                             ' used to do the export
    sFileName          As String * 80        ' The name of the FIRST file ion
                                             ' the export.  An export can send
                                             ' many files all with different
                                             ' names
    sErrors            As String * 80        ' If an error occurs it gets logged
                                             ' here
    sExpSuccess        As String * 1         ' Export was successful or not.  Y
                                             ' = successful, N = not successful
    sUnused            As String * 40        ' Extra space
End Type

'***********Added for use with Crystal TTX (temp tables) RG 9-14-07
Type ATTEXPMON
    lCode                 As Long            ' Internal code number for Agreement
    sVehName              As String * 40        ' Name
    sCallLetters          As String * 40
    sMarket               As String * 60
End Type

Public tgAttExpMon() As ATTEXPMON
Public lgAttExpMonCount As Long

Public sgAttTimeStamp As String     'Date/time of file.  Used to know to repopulate arra

Type ESFCODE
    lCode              As Long
End Type
Type EDFCODE
    lAttCode           As Long               ' AttCode
    lEsfCode           As Long               ' Link to ESF - Export Summary

End Type

Type REPORTNAMES
    sRptName As String                      'report name in list box
    sCrystalName As String
    iRptIndex As Integer                    'report index
    sRptPicture As String                 'report .bmp, .jpg showing sample of report
    sRptDesc As String                    'report description
End Type

Type EMPLOYEEEMAILS
    Name As String
    eMail As String
End Type


Type AFFMESSAGES
    Name As String
    fileName As String
End Type

Type AST_INFO
    lSdfCode As Long
    lSdfCode2 As Long
    iPledgeStatus As Integer
    iPledgeStatus2 As Integer
    lCode As Long
    lCode2 As Long
    lAttCode As Long
    iStatus As Integer
    iStatus2 As Integer
    sAirDate As String * 10
    sAirTime As String * 11
    lFeedDate As Long
    lFeedTime As Long
    iFeedDay As Integer
    sAirTime2 As String * 11
End Type


Public tgReportNames() As REPORTNAMES
Declare Function CreateFieldDefFile Lib "p2smon.dll" (lpUnk As Object, ByVal fileName As String, ByVal bOverWriteExistingFile As Long) As Long

Public Const ERROR_NO_MORE_FILES = 18

Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

' These vars define whether we are importing or exporting files from the web server.
Public Const webImport = 1
Public Const webExport = 2

' Added 08-22-03  J. Dutschke
' Private INI file operations
'
Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefault$, ByVal lpReturnString$, ByVal nSize As Long, ByVal lpFileName$)
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpAppName$, ByVal lpKeyName As Any, ByVal lpWriteString$, ByVal lpFileName$)

Type SYSTEMTIME ' 16 Bytes
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName As String * 64
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName As String * 64
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type
Public Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long


Public Const MIB_IF_TYPE_ETHERNET                As Long = 6
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH      As Long = 128
Public Const MAX_ADAPTER_DESCRIPTION_LENGTH_p    As Long = MAX_ADAPTER_DESCRIPTION_LENGTH + 4
Public Const MAX_ADAPTER_NAME_LENGTH             As Long = 256
Public Const MAX_ADAPTER_NAME_LENGTH_p           As Long = MAX_ADAPTER_NAME_LENGTH + 4
Public Const MAX_ADAPTER_ADDRESS_LENGTH          As Long = 8
'dan
'Replaced with cdlOFNOverwritePrompt: Public Const cdlOFNOverwritePrompt = &H2&
'Replaced with cdlOFNHideReadOnly: Public Const cdlOFNHideReadOnly = &H4&
'Replaced with cdlOFNNoChangeDir: Public Const cdlOFNNoChangeDir = &H8&
'Replaced with cdlOFNPathMustExist: Public Const cdlOFNPathMustExist = &H800&
'Replaced with cdlOFNCreatePrompt: Public Const cdlOFNCreatePrompt = &H2000&
'Replaceed with CdlOFNNoReadOnlyReturn: Public Const CdlOFNNoReadOnlyReturn = &H8000&
'Public Const DLG_FILE_SAVE = 2

Type TIME_t
    aTime As Long
End Type

Type IP_ADDRESS_STRING
    IPadrString     As String * 16
End Type

Type IP_ADDR_STRING
    AdrNext         As Long
    IpAddress       As IP_ADDRESS_STRING
    IpMask          As IP_ADDRESS_STRING
    NTEcontext      As Long
End Type
' Information structure returned by GetIfEntry/GetIfTable
Type IP_ADAPTER_INFO
    Next As Long
    ComboIndex As Long
    AdapterName         As String * MAX_ADAPTER_NAME_LENGTH_p
    Description         As String * MAX_ADAPTER_DESCRIPTION_LENGTH_p
    MACadrLength        As Long
    MACaddress(0 To MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    AdapterIndex        As Long
    AdapterType         As Long             ' MSDN Docs say "UInt", but is 4 bytes
    DhcpEnabled         As Long             ' MSDN Docs say "UInt", but is 4 bytes
    CurrentIpAddress    As Long
    IpAddressList       As IP_ADDR_STRING
    GatewayList         As IP_ADDR_STRING
    DhcpServer          As IP_ADDR_STRING
    HaveWins            As Long             ' MSDN Docs say "Bool", but is 4 bytes
    PrimaryWinsServer   As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained       As TIME_t
    LeaseExpires        As TIME_t
End Type

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)
Public Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (ByRef pAdapterInfo As Any, ByRef pOutBufLen As Long) As Long

'******************************************************************************
' UAF_User_Activity Record Definition
'
'******************************************************************************
Type UAF
    lCode                 As Long            ' Auto Increment code
    sSystemType           As String * 1      ' System Type (T=Traffic;
                                             ' A=Affiliate)
    sSubType              As String * 1      ' Sub type (T=Task; R=Report)
    lUlfCode              As Long            ' User Log reference code
    iUserCode             As Integer         ' User Refernce code(urf if
                                             ' uafSystemType = T; or ust if
                                             ' uafSystemType = A)
    sName                 As String * 50     ' Task Name or Report name
    sStatus               As String * 1      ' Status (C=Complete; I=In
                                             ' Progress; A=Aborted)
    iStartDate(0 To 1)    As Integer         ' Start Date of Task or Report
    iStartTime(0 To 1)    As Integer         ' Start Time of Task or Report
    iEndDate(0 To 1)      As Integer         ' End date of Task or Report
                                             ' (12/31/2069 if not set)
    iEndTime(0 To 1)      As Integer         ' End Time of Task or Report
                                             ' (00:00:00 if not set)
    iCSIDate(0 To 1)      As Integer         ' CSI date
    sUnused               As String * 20     ' Unused
End Type

Type UAFKEY0
    lCode                 As Long
End Type

Type UAFKEY1
    sSystemType           As String * 1
    iUserCode             As Integer
End Type

Type UAFKEY2
    lUlfCode              As Long
    sStatus               As String * 1
    iStartDate(0 To 1)    As Integer
    iStartTime(0 To 1)    As Integer
End Type

Type UAFKEY3
    iStartDate(0 To 1)    As Integer
End Type

'******************************************************************************
' ULF_User_Log Record Definition
'
'******************************************************************************
Type ULF
    lCode                 As Long            ' Auto Increment code
    sDBType               As String * 1      ' Database Type: Production (P) or
                                             ' Test (T)
    sSystemType           As String * 1      ' System Type: Traffic(T) or
                                             ' Affiliate(A) or Engineering(E)
    iurfCode              As Integer         ' Traffic User Reference code
    iUstCode              As Integer         ' Affiliate User reference code
    iUieCode              As Integer         ' Engineering User reference code
    iSignOnDate(0 To 1)   As Integer         ' Sign On Date
    iSignOnTime(0 To 1)   As Integer         ' Sign On Time
    iSignOffDate(0 To 1)  As Integer         ' Sign Off Date
    iSignOffTime(0 To 1)  As Integer         ' Sign off time
    sPCName               As String * 255    ' Computer name
    sPCMACAddr            As String * 20     ' Network card address (Unique by
                                             ' PC)
    sTimeZone             As String * 1      ' Time zone: Eastern(E);
                                             ' Central(C); Mountain(M);
                                             ' Pacific(P)
    'sUnused               As String * 20     ' Unused
    iActiveLogDate(0 To 1)   As Integer         ' Updated hourly to be used to determine if user logged off incorrectly
    iActiveLogTime(0 To 1)   As Integer         ' Updated hourly to be used to determine if user logged off incorrectly
    iTrafJobNo            As Integer         'Traffic Job number
    iTrafListNo           As Integer         'Traffic List number
    iTrafRnfCode          As Integer         'Traffic report reference code
    iAffTaskNo            As Integer         'Affiliate Task number
    iAffSubtaskNo         As Integer         'Affiliate subtask number
    iAffRptNo             As Integer         'Affiliate report number
    'sUnused               As String * 12     ' Unused
End Type

Type ULFKEY0
    lCode                 As Long
End Type

Type ULFKEY1
    sDBType               As String * 1
    sSystemType           As String * 1
    iurfCode              As Integer
    iUstCode              As Integer
    iUieCode              As Integer
End Type

'Selling to Airing link information
Type SALINKINFO
    iSellCode As Integer
    iAirCode As Integer
    lSellTime As Long
    lAirTime As Long
    iBreak As Integer
    iPosition As Integer
End Type

' New functions to control backup.

'Using Features5
Global Const REMOTEEXPORT = &H1

'Using Features7
Global Const CSIBACKUP = &H1
Global Const XDIGITALISCIEXPORT = &H20
Global Const WEGENEREXPORT = &H40
Global Const OLAEXPORT = &H80

'Using Features8
Global Const ALLOWMSASPLITCOPY = &H4
Global Const XDIGITALBREAKEXPORT = &H10
Global Const ISCIEXPORT = &H20

Type CSISvr_Rsp_GetLastBackupDate
    sLastBackupDateTime As String * 20
End Type

Type CSISvr_Rsp_Answer
    iAnswer As Integer  ' 0=No, 1=Yes
End Type


Public bgUsingSockets As Boolean
Public gUsingCSIBackup As Boolean
Public gUsingXDigital As Boolean
Public gWegenerExport As Boolean
Public gOLAExport As Boolean
Public gUsingMSARegions As Boolean
Public gISCIExport As Boolean
Public bgRemoteExport As Boolean

Declare Function csiGetLastBackupDate Lib "CSI_CNT32.dll" (ByVal SDBPath$, ByRef SvrRsp As CSISvr_Rsp_GetLastBackupDate) As Integer
Declare Function csiGetLastCopyDate Lib "CSI_CNT32.dll" (ByVal SDBPath$, ByRef SvrRsp As CSISvr_Rsp_GetLastBackupDate) As Integer
Declare Function csiStartBackup Lib "CSI_CNT32.dll" (ByVal SDBPath$, ByVal sINIPathFileName$, ByVal BUType As Integer) As Integer
Declare Function csiIsBackupRunning Lib "CSI_CNT32.dll" (ByVal SDBPath$, ByRef SvrRsp As CSISvr_Rsp_Answer) As Integer
Declare Function csiCheckForFilesStuckInCntMode Lib "CSI_CNT32.dll" (ByVal SDBPath$, ByRef SvrRsp As CSISvr_Rsp_Answer) As Integer

' XML Web Service Functions
Type CSIRspGetXMLStatus
    sStatus As String * 2048
End Type
'6807
'Declare Function csiXMLStart Lib "CSI_Utils.dll" (ByVal slIni$, ByVal slSection$, ByVal slType$, ByVal slFileName$, ByVal slLineEndChar$) As Integer
Declare Function csiXMLStart Lib "CSI_Utils.dll" (ByVal slIni$, ByVal slSection$, ByVal slType$, ByVal slFileName$, ByVal slLineEndChar$, ByVal slXmlErrorResponseFile$) As Integer
Declare Function csiXMLSetMethod Lib "CSI_Utils.dll" (ByVal slMethodName$, ByVal slTableName$, ByVal slTransmissionID$, ByVal slSchemaName$) As Integer
Declare Function csiXMLData Lib "CSI_Utils.dll" (ByVal slType$, ByVal slTag$, ByVal slvalue$) As Integer
'Dan 6/27/14  FlushQueue. 1 is normal..send.  0 is queue and don't send
Declare Function csiXMLWrite Lib "CSI_Utils.dll" (ByVal FlushQueue As Long) As Integer
Declare Function csiXMLEnd Lib "CSI_Utils.dll" () As Integer
Declare Function csiXMLStatus Lib "CSI_Utils.dll" (ByRef sStatus As CSIRspGetXMLStatus) As Integer
'Dan M 10/26/10
'The function csiXMLStartRead cannot fail so it will always return true.The function csiXMLReadData will return true or false.
Declare Function csiXMLStartRead Lib "CSI_Utils.dll" (ByVal slIni$, ByVal slSection$, ByVal slExportPath$) As Integer
Declare Function csiXMLReadData Lib "CSI_Utils.dll" () As Integer
'6966 'FlushQueue is ignored. Send 1
Declare Function csiXMLResend Lib "CSI_Utils.dll" (ByVal FlushQueue As Long) As Integer

'General Message (AffGenMsg.  This was taken from Traffic GenMsg)
Public sgGenMsg As String       'Message to be shown
Public sgCMCTitle(0 To 3) As String    'Button Titles (set title to blank if not to be shown)
Public igDefCMC As Integer      'Default Button Number (0, 1 or 2)
Public igAnsCMC As Integer      'Button Selected (0, 1, 2 or 3)
Public igEditBox As Integer     '0=No; 1=Yes; 2=Radio buttons defined instead of edit box
Public sgEditValue As String
Public sgRadioTitle(0 To 3) As String 'Raio captions.

'Get Path
Public igGetPath As Integer '0=Done pressed; 1=Cancel pressed
Public sgGetPath As String
Public igPathType As Integer    '0=Any; 1=subfolder only
'Dan M 7/28/09  cr2008 and csiNetReporter2
Public Type RQF
    lCode                 As Long
    sPriority             As String * 1      ' L=Low, N=Normal, H=High
    iPrintCopies          As Integer         ' # of copies to print
    sReportName           As String * 20
    sRunType              As String * 1      ' N=Now,D=Daily,W=Weekly,F=First
                                             ' day after month end
    sReportSource         As String * 1      ' N=No pre-pass P=pre-pass
    sReportType           As String * 1      ' T=Traffic A=Affiliate
    sOutputType           As String * 1      ' D=Display P=Print S=Save to file
    iOutputSaveType       As Integer         ' 0=pdf 1=excel 2=word 3=text 4=csv
                                             ' 5=rtf
    sOutputFileName       As String * 200    ' can include path
    sRunMode              As String * 1      ' C=Client S=Server
    sRunTime              As String * 11
    iRunDay               As Integer         ' 1 = monday -> 7= Sunday
    sLastDateRun          As String * 10
    lPrePassDate          As Long
    lPrePassTime          As Long
    lEnteredDate          As Long
    lEnteredTime          As Long
    sUserName             As String * 20     ' User Name
    sDisposition          As String * 1      ' E=erase when done, R=retain when done
    sCompleted            As String * 1      ' N, Y, P (Processing) and E (Completed but had error)
    lConnection           As Long           '0 is api call, 1 is odbc
    lRqfCode              As Long            ' Used to obtain the Multi-Report
                                             ' RQF records.  Master (or Parent)
                                             ' RQF Code stored into each
                                             ' Multi-Report.
    iMultiReportSeqNo     As Integer         ' Sequence number of the
                                             ' multi-reports.
    sPCMACAddr            As String * 20     ' MAC Address
    sUnused               As String * 14
End Type


Public Type RQFKEY0 'VBC NR
    lCode                 As Long 'VBC NR
End Type 'VBC NR
' Dan M 10/21/09 these won't be used in traffic
Public Type RQFKEY1 'VBC NR
    sRunMode              As String * 1 'VBC NR
    sCompleted            As String * 1
    sPriority             As String * 1 'VBC NR
    lEnteredDate          As Long 'VBC NR
    lEnteredTime          As Long 'VBC NR
End Type 'VBC NR
Public Type RQFKEY2 'VBC NR
    lRqfCode              As Long
    iMultiReportSeqNo     As Integer
End Type 'VBC NR
Public Type RQFKEY3
    sUserName             As String * 20     ' User Name
    lEnteredDate          As Long 'VBC NR
    lEnteredTime          As Long 'VBC NR
End Type 'VBC NR

Public Type RFF
    lCode                 As Long
    lRqfCode              As Long
    iSequenceNumber       As Integer
    sFormulaName          As String * 40     ' used by traffic
    sType                 As String * 1      ' F=formula field, R= record selection, A= ADO, M=Multi-Report, P=Pre-pass report selection criteria
    sFormulaValue         As String * 255
    lRffCode              As Long            ' 0 unless is a child of extended value; parent code acts as foreign key
    lExtendExists         As Long            ' 0 = no 1=yes.  Only rff to be split (parent) gets a 1
    sUnused               As String * 20
End Type


Public Type RFFKEY0 'VBC NR
    lCode                 As Long 'VBC NR
End Type 'VBC NR

Public Type RFFKEY1 'VBC NR
    lRqfCode              As Long 'VBC NR
    sType                 As String * 1 'VBC NR
    iSequenceNumber       As Integer 'VBC NR
End Type 'VBC NR
Public tgRff() As RFF
Public tgRffExtended() As RFF
Public igJobRptNo As Integer

'Browser
Public igBrowseType As Integer '1=Station Update Data (Csv:comma delimited)
Public igBrowseReturn As Integer   '0=Cancelled; 1=Ok
Public sgBrowseFile As String  'Drive\Path\FileName if igBrowserReturn=1 of the file
Public sgBrowseMaskFile As String   '
Public sgBrowseTitle As String
Public sgStationImportTitles() As String

'Transfer Control
Public bgStationVisible As Boolean
Public bgAgreementVisible As Boolean
Public bgEMailVisible As Boolean
Public bgLogVisible As Boolean
Public bgAffidavitVisible As Boolean
Public bgPostBuyVisible As Boolean
Public bgManagementVisible As Boolean
Public bgExportVisible As Boolean
Public bgSiteVisible As Boolean
Public bgUserVisible As Boolean
Public bgRadarVisible As Boolean


Public sgStationCallSource As String * 1    'D=Directory; S=Search Station
Public sgAgreementCallSource As String * 1  'D=Directory; S=Search Station
Public igTCShttCode As Integer
Public sgTCCallLetters As String
Public lgTCAttCode As Long
Public sgStationSearchCallSource As String * 1  'M=Management; P=Post By

'Set Fields
Public sgSetFieldCallSource As String * 1   'M=Menu; S=Start of system

Type USTINFO
    iCode As Integer
    sName As String
    iDntCode As Integer
End Type
Public tgUstInfo() As USTINFO

Type DEPTINFO
    iCode As Integer
    sName As String
    lColor As Long
    sType As String
End Type
Public tgDeptInfo() As DEPTINFO

Type PETINFO
    lCode As Long
    lGhfCode As Long
    lgsfCode As Long
    sDeclaredStatus As String * 1
    sClearStatus As String * 1
    sChanged As String * 1
End Type
Public tgPetInfo() As PETINFO

Type SALESPPLINFO
    iSlfCode As Integer
    sFirstName As String * 20
    sLastName As String * 20
    sOffice As String * 20
    sSource As String * 20
    iSSMnfCode As Integer
End Type
Public tgSalesPeopleInfo() As SALESPPLINFO

'Monitor
Public Const MONITORTIMEINTERVAL As Long = 10   'In second
Public Const MONITORRUNINTERVAL As Long = 30    'In Seconds
Public Const MONITORELAPSEDINTERVAL As Long = 120    'In Seconds
Public Const MONITORELAPSEDINTERVALPERIOD As Long = 240   'In sec
Public Const MONITOREMAILINTERVAL As Long = 3600

Public sgTimeZone As String

'******************************************************************************
' TMF_Task_Monitor Record Definition
'
'******************************************************************************
Type TMF
    lCode                 As Long            ' Auto-increment
    sTaskCode             As String * 3      ' CSC=ContractSpotCreation;ASC=Affi
                                             ' liateSpotCreation;AEQ=AffiliateEx
                                             ' portQueue;ASI=AffidavitSpotImport
                                             ' ;ARQ=AffiliateReportQueue;ASQ=Ava
                                             ' ilSummaryGen;SC=SetCredit;PE=Proj
                                             ' ection;SFE=SalesForce;ME=Matrix;E
                                             ' PE=EfficioProj;ERE=EfficioRev;GPE
                                             ' =GetPaid;BD=BackupData  7967Dan WVI=WebVendorImport
    sTaskName             As String * 30     ' Task Name
    sService              As String * 1      ' C=CSI_Service; T=Task
                                             ' Schedule;N=None   7967Dan- W=Web Service
    sRunMode              As String * 1      ' P=Periodic; C=Continuous
    iRunningDate(0 To 1)  As Integer         ' Cintinuously Updated every X
                                             ' seconds. Used to determine if
                                             ' program is running
    iRunningTime(0 To 1)  As Integer         ' Cintinuously Updated every X
                                             ' seconds. Used to determine if
                                             ' program is running
    i1stStartRunDate(0 To 1) As Integer      ' 1st run start date for a day
    i1stStartRunTime(0 To 1) As Integer      ' 1st run time for a day
    i1stEndRunDate(0 To 1) As Integer        ' 1st run end date for a day
    i1stEndRunTime(0 To 1) As Integer        ' 1st run end time for a day
    iStartRunDate(0 To 1) As Integer         ' Last Start Date for a day task
                                             ' started to run. Periodic run
                                             ' mode. Used to determine if
                                             ' program is running
    iStartRunTime(0 To 1) As Integer         ' Last Start Time for a day task
                                             ' started to run. Periodic run
                                             ' mode. Used to determine if
                                             ' program is running
    iEndRunDate(0 To 1)   As Integer         ' Last End Date for a day task
                                             ' completed. Periodic run mode.
                                             ' Used to determine if program is
                                             ' running
    iEndRunTime(0 To 1)   As Integer         ' Last End Time for a day task
                                             ' completed. Periodic run mode.
                                             ' Used to determine if program is
                                             ' running
    sMo                   As String * 1      ' Run on Monday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sTu                   As String * 1      ' Run on Tuesday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sWe                   As String * 1      ' Run on Wednesday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sTh                   As String * 1      ' Run on Thursday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sFr                   As String * 1      ' Run on Friday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sSa                   As String * 1      ' Run on Saturday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sSu                   As String * 1      ' Run on Sunday: Y=Yes; N=No.
                                             ' Continuous Run Mode
    sMonthPeriod          As String * 2      ' SE=End Standard; EC=End Calendar;
                                             ' I=Invoice. Periodic Run Mode
    iDaysAfter            As Integer         ' Number of days after period that
                                             ' task is run. Periodic Run Mode
    iEMailReqDate(0 To 1) As Integer         ' Date detected that E-Mail needs
                                             ' to be sent. Wait W minutes before
                                             ' sending
    iEMailReqTime(0 To 1) As Integer         ' Time detected that E-Mail needs
                                             ' to be sent. Wait W minutes before
                                             ' sending
    iEMailSentDate(0 To 1) As Integer        ' E-Mail sent date
    sStatus               As String * 1      ' Status: S=Started; C=Completed; E=Error; A=Aborted
    sUnused               As String * 10     ' Unused
End Type

'Type TMFKEY0
'    lCode                 As Long
'End Type

Type TMFKEY1
    sTaskCode             As String * 3
End Type
'7458
'Type ENT
'    lCode              As Long               ' Export Number table reference
'    sType              As String * 1         ' S=Sent Unposted to Web;
'                                             ' R=Received Posted from
'                                             ' Web;E=Export Unposted to 3rd
'                                             ' party; I=Import Posted 3rd party;
'                                             ' P=Sent to Web 3rd party posted
'                                             ' spots; F=Forward from Web to 3rd
'                                             ' party;T=Imported to Web from 3rd
'                                             ' party
'    s3rdParty          As String * 1         ' Third Party: W=Web;
'                                             ' C=Wegener-Compel;
'                                             ' I=Wegener-iPump;
'                                             ' M=Marketron;Q=Cumulus; B=CBS;
'                                             ' J=Jelli; H=Clear Channel; D=IDC
'    lAttCode           As Long               ' Agreement internal reference code
'    iShttCode          As Integer            ' Station internal reference code
'    iVefCode           As Integer            ' Vehicle internal reference code
'    sFeedDate          As String * 10        ' Spot Feed date
'    lGsfCode           As Long               ' Event Internal reference code
'    iAstCount          As Integer            ' Affiliate Spot count including
'                                             ' Not Carried
'    iSpotCount         As Integer            ' Spots count. entType S: Spots
'                                             ' sent; entType R: Spots received
'    iMGCount           As Integer            ' entType R: MG spot count. entType
'                                             ' S: zero (0)
'    iReplaceCount      As Integer            ' entType R: Replacement count;
'                                             ' entType S: Zero (0)
'    iBonusCount        As Integer            ' entType R: Bonus count; entType
'                                             ' S: Zero (0)
'    iIngestedCount     As Integer            ' Affiliate spots ingested
'    sEnteredDate       As String * 10        ' Entered date (PC date)
'    sEnteredTime       As String * 11        ' Entered time (PC time)
'    sFileName          As String * 60        ' File Name transferred
'    sStatus            As String * 1         ' entType S: S=Successfully sent,
'                                             ' E=Error; entType R:S=Successfully
'                                             ' Imported, E=Error
'    iUstCode           As Integer            ' User internal reference code. 0
'                                             ' if Auto-Import
'  '  sUnused            As String * 20        ' Unused
'End Type

Type ENT
    lCode                 As Long            ' Export Number table reference
    sType                 As String * 1      ' S=Sent Unposted to Web;
                                             ' R=Received Posted from
                                             ' Web;E=Export Unposted to 3rd
                                             ' party; I=Import Posted 3rd party;
                                             ' P=Sent to Web 3rd party posted
                                             ' spots; F=Forward from Web to 3rd
                                             ' party;T=Imported to Web from 3rd
                                             ' party
    s3rdParty             As String * 1      ' Third Party: W=Web;
                                             ' C=Wegener-Compel;
                                             ' I=Wegener-iPump;
                                             ' M=Marketron;Q=Cumulus; B=CBS;
                                             ' J=Jelli; H=Clear Channel; D=IDC;
                                             ' X=X-Digital
    lAttCode              As Long            ' Agreement internal reference code
    iShttCode             As Integer         ' Station internal reference code
    iVefCode              As Integer         ' Vehicle internal reference code
    sFeedDate             As String * 10        ' Spot Feed date
    lgsfCode              As Long            ' Event Internal reference code
    iAstCount             As Integer         ' Affiliate Spot count including
                                             ' Not Carried
    iSpotCount            As Integer         ' Spots count. entType S: Spots
                                             ' sent; entType R: Spots received
    iMGCount              As Integer         ' entType R: MG spot count. entType
                                             ' S: zero (0)
    iReplaceCount         As Integer         ' entType R: Replacement count;
                                             ' entType S: Zero (0)
    iBonusCount           As Integer         ' entType R: Bonus count; entType
                                             ' S: Zero (0) entType E and 3rdParty X:1=all spots 2 = regional only
    iIngestedCount        As Integer         ' Affiliate spots ingested
    iDeleteCount          As Integer         ' Number of spots deleted. entType
                                             ' S: spots sent to be deleted;
                                             ' entType R: Spots ignored
    sEnteredDate          As String * 10        ' Entered date (PC date)
    sEnteredTime          As String * 11        ' Entered time (PC time)
    sFileName             As String * 60     ' File Name transferred
    sErrorMsg             As String * 40     ' Error message
    sStatus               As String * 1      ' entType S: S=Successfully sent,
                                             ' E=Error; I=Incomplete; N=Not
                                             ' sent; entType R:S=Successfully
                                             ' Imported, E=Error;
    iUstCode              As Integer         ' User internal reference code. 0
                                             ' if Auto-Import
    sUnused               As String * 20     ' Unused
End Type



Public Type PeekArrayType
    Ptr         As Long
    Reserved    As Currency
End Type
Public Declare Function PeekArray Lib "kernel32" Alias "RtlMoveMemory" (Arr() As Any, Optional ByVal LENGTH As Long = 4) As PeekArrayType
