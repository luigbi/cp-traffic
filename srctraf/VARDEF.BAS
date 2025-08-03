Attribute VB_Name = "VARDEF"
'******************************************************************************************
'***** VB Compress Pro 6.11.32 generated this copy of VARDEF.BAS on Wed 6/17/09 @ 12:56 PM
'***** Mode: AutoSelect Standard Mode (Internal References Only)***************************
'******************************************************************************************
'* Note: VBC id'd the following unreferenced items and handled them as described:         *
'*                                                                                        *
'* Public Variables (Removed)                                                             *
'*  sgManyCntCodeTagCB            sgItfTagRpt                   tgCgf                     *
'*  tgMsf                         tgChfImpt                     tgChfED                   *
'*                                                                                        *
'*                                                                                        *
'* Public Type Defs (Marked)                                                              *
'*  RQFKEY0                       RQFKEY1                       RFFKEY0                   *
'*  RFFKEY1                                                                               *
'******************************************************************************************
'
' Copyright 1993 Counterpoint Software, Inc. All rights reserved.
' Proprietary Software, Do not copy
'
' File Name: VarDef.BAS
'
' Release: 1.0
'
' Description:
'   This file contains the Public variables
'
'   Note: gEraseGlobalVar is used to Erase all Public Arrays
'         This routine is called from the Unload event within the
'         traffic form.
'
Option Explicit

Public igInvErrFlag As Integer  'd.s. 11/6/01Flag to notify there was an error in nvoicing
Public igSystemResources As Integer 'System resources window 3.0=475; 3.1=600
Public fgWinVersion As Single    'Window Version (3.1, 3.95,...)
Public igNoMinForms As Integer  'Number of min forms
Public igFromSignon As Integer  'Required to avoid loading a MODELESS (Basic10) prior
                                'to Signon totally unloaded when user types-head the
                                'user name and password (add 1 sec delay before loading basic10)
'Public lgDateCreditShown As Long    'Date that credit was shown for this user
Public igShowCreditShown As Integer
Public sgBrannerMsg As String   'Message in main form caption
Public sgDateBrannerMsg As String  'Branner moved to date
Public igShowPicture As Integer    'True = show picture instead of Jobs/Lists
                                    'False = show jobs/lists instead of picture
Public sgCSVersion As String    'Version message- set in gInitGlobalVar
Public sgDDFDateInfo As String
Public sgSystemDate As String   'System date when first start system
Public sgShowDateForm As String '6-16-11 for csi date control
Public igResetSystemDate As Integer 'True=reset system date when terminating traffic
Public igRemoteUser As Integer  'True=Remote user- Don't show picture
Public igBypassFileSize As Integer  'True=Bypass file size test
Public igDemo As Integer    'True=Demo system
Public igChkExe As Integer  'True=Check Date/Time of exe; rpt and lst
Public igGGFlag As Integer  '-1 = Not Set; 0= Bad Guy; 1= CSI; # of days
Public igRptGGFlag As Integer   '0=Disallow Report; 1=Allow reports. test only when igGGFlag = 0
Public igGGBDSignOn As Integer
Public igGGDBSignOnDays As Integer
Public igOnCitrix As Integer    'True or False
Public igBkgdProg As Integer    'See Gensubs gMsgBox for all of the definitions  0=No; 1=Scheduler; 2=Set Credit; 3=Export to SBS; 10=Initializing Background Mode (Log errors to TrafficErrors.txt)
Public igExportType As Integer  '0=Manual; 1=From Traffic, 2=Auto-Efficio Projection; 3=Auto-Efficio Revenue; 4=Auto-MatrixStd, 5=Auto-MatrixCal (N/A: matrix uses 4 for both std & cal), 5 = Auto-Tableau, 6 = auto-rab, 7 = Auto-CustomRevenueExport
Public sgExportIniSectionName As String  'reqd to find entries in exports.ini for matrix options

Public igPasswordOk As Integer  'Password status obtained from the form CSPWord
                                'This is a changing password that is stored into tmSpf
Public sgPasswordAddition As String 'Message to add to Previous Password
Public sgSpecialPassword As String  'Password part of sign on password (radioxxxx, save xxxx here if 4 characters)
                                    'Allow those same for characters as pasword to site and other task
Public tgPWUrf As URF
Public sgPWResult As String

Public igTestSystem As Integer  'True=Test System; False=Production System
Public igUpdateOk As Integer    'Pass update status to other modules (like Site -> Corp Cal)
Public sgCurDir As String       'Current Directly
Public igShowVersionNo As Integer   'Show version number status (0=Client; 1=internal not debug; 2=internal debug)
Public sgWallpaper As String    'Obtained from ini

Public sgUserNameToPassToAffiliate As String
Public sgUserPasswordToPassToAffiliate As String

Public lgCurrHRes As Long
Public lgCurrVRes As Long
Public lgCurrBPP As Long        '4=16 Colors; 8=256 colors; 16 High; 24 & 32 = True Color


Public igSiteSet As Integer 'True=Site previously defined; False=Site should be set
Public igVehicleDefined As Integer 'True=Vehicle defined; False=Vehicle not defined

Public sgApplyVehPreemptRule As String * 1
Public sgVehPreemptRule As String * 1

Public tgSite As SITE

Public sgDefaultTerms As String

Public sgCallAppName As String  'Calling application name-always first parameter in command line
Public igChildDone As Integer   'True = child done
                                'False = child still processing
Public igParentRestarted As Integer 'True=Parent has focus, unload
Public sgDoneMsg As String      'Done message
'Public igWaitForDDE As Integer  'True=Waiting for child to start DDE
                                'False=Receive first message from child or
                                'if child calling call time delay triggered

Public igStdAloneMode As Integer  'True=Program running in Debug as stand alone- bypass links)
                                    'False=Program running as EXE (shell initialled code- links required)

Public sgCommandStr As String   'Used to pass information to Modal module
Public sgUserName As String 'User name
Public sgUserDefVehicleName As String  'Users default Vehicle name
Public sgUserSalespersonName As String  'Salesperson name of the user
Public igUpdateAllowed As Integer    'User can update records
Public igRUDirectMode As Integer    '0=Remote user mode; 1=Direct Mode, set RemoteID to zero; 2=Standard Mode, leave RemoteID (Default), disallow Proposals
Public sgDBPath As String   'Data Base Path
Public sgMDBPath As String  'Master database path
Public sgSDBPath As String  'Slave Database Path
Public sgTDBPath As String  'Temporary Report File Database Path
'Public sgTDBNames(1 To 12) As String   'Set in initsubs.bas
Public igRetrievalDB As Integer         '0=Master, 1=Slave
Public sgRptPath As String              'Report path
Public sgRptSavePath As String          'Report File Save path
Public sgImportPath As String           'Report File Save path
Public sgExportPath As String           'Report File Save path
Public sgArchivePath As String          'Archive file save path
Public sgLogoPath As String             'Logo
Public sgHelpPath As String             'Help
Public sgExePath As String              'Path to exe files (used in Shell command)
Public sgDLExePath As String            'Path to Download exe Files
Public sgDLRptPath As String            'Path to Download exe Files
Public sgDDFPath As String              'Path to DDF files (used in report calls)
Public sgURL_Documentation As String    'Web Path for documentation 'Added 2024-01-09 JJB for dynamic documentation path
Public sgAttachment_Types As String     'Contract Attachment types that are allowed
Public sgLogGenExportPath As String        '2-11-15 Log export path (log generation files); default to sgRptSavePath if not defined
Public sgSaveFilePathToUse As String        '2-12-15  depending on task, this will be the default sgRptSavePath or the Log Gen export path if generating logs
                                            'Log Gen has option to save all log export files in a different path defined in traffic.ini
Public igDDEDelay As Integer
Public igUsingTimeStamp As Integer  'True=Obtain file Time Stamp; False=Ignore file Time Stamp
Public igSSFBaseLen As Integer
Public igAlignCharWidth As Integer
Public igTaskLoc As Integer     'Task bar location (0=top; 1=bottoom)
Public sgServerDatabase As String  'path as used on the server
Public igIsSignedOn As Integer 'False until the user is signed on
Public igTerminateAndUnload As Integer  'True= call mTerminate as part of unload command
Public bgRptLogoJpgExists As Boolean        '8-18-14
Public sgRptLogoExt As String * 3           '8-18-14

Public igUserLogButton As Integer   '1=Message; 2=Send Alert; 3=Send Ok message
Public lgUserLogUlfCode() As Long    'User Log to receive message or alert

Public sgProphetExportPath As String '1-6-05
Public sgProphetImportPath As String
Public igProphetImportPathFlag As Integer

Public sgWideOrbitExportPath As String '1-10-12
Public sgWideOrbitImportPath As String
'Public igWideOrbitImportPathFlag As Integer

Public sgJelliExportPath As String  '6-22-12
Public sgJelliImportPath As String

Public sgScottExportPath As String  '5-30-13
Public sgScottImportPath As String
Public sgZettaExportPath As String '1-7-16
Public sgZettaImportPath As String

Public sgStationInvoiceImportPath As String


Public sgRevenueExportPath As String    '9-1-05 Revenue projection export path

Public sgSalesBrochurePath As String

'Export
Public sgExportMenuItem As String
'ImportResearch
Public sgImportResearchMenuItem As String


'Browser
Public igBrowserType As Integer '0=Bmp (bit map); 1=Csv (comma delimited); 2= Txt (text); 3= *.Rec; 4= *.rt?; 5 = *.ct?; 6 = Tape*.*; 7 = Mask supplied
Public igBrowserReturn As Integer   '0=Cancelled; 1=Ok
Public sgBrowserFile As String  'Drive\Path\FileName if igBrowserReturn=1 of the file
Public sgBrowseMaskFile As String   'Used with igBrowserType = 7
Public sgBrowserTitle As String
Public sgBrowserDrivePath As String 'Used for PDF (8) and sgBrowserFile contains just the file names

'Report variables
Public igSwapBOBOption As Integer   '8-6-10 Changing from vehicle with subtotals & splits to slsp with subtotals & splits
Public igUsingCrystal As Integer
Public igShowCatOrNames As Integer  '0=Categories; 1=Names
Public sgPhoneImage As String
Public igJobRptNo As Integer
Public igZones As Integer                   'time zones  (0=all, 1=est, 2=cst, 3=mst, 4=pst)
Public igRnfCode As Integer
Public igInvoiceType As Integer     '0=invoice (no spots) - invport1.rpt, 1 = affidavit (no $) invaff.rpt
Public sgInvMonthYear As String     '11-16-16 month and year invoice genned, communication between invoice & report modules
Public igGenRpt As Integer     'True = Generating Report ignore user input
Public igOutput As Integer
Public igCopies As Integer
Public igFile As Integer
Public igOption As Integer
Public igOutputTo As Integer        '0 = display , 1 = print
Public sgInclRates As String           'Includes Rates on contract = Y/N
Public sgInclResearch As String        'Include Research on contract = Y/N
Public sgInclSplits As String           '2-13-04 show slsp comm splits Y/N
Public sgInclNTRBillSummary As String   '2-2-10 show NTR billing summary with air time summary Y/N
Public bgSuppressNTRDetails As Boolean   'TTP 10745 - NTR: add option to only show vehicle, billing date, and description on the contract report, and vehicle and description only on invoice reprint
Public igBRSumZer As Integer       '1-28-10 need to know if its monthly summary to bring in correct filters for crystal
Public igBRSum As Integer          '2-12-10 need to know its a research summary version to show NTR vehicle totals with the air time Research
Public sgShowNetOnProps As String       '2-3-10 show net & agy comm on proposals
Public sgShowProdProt As String         '8-25-15 Show prod protection on prop/orders
Public sgInclProof As String           'Include Hidden lines on contract = Y/N
Public igDetSumBoth As Integer         '0 = detail, 1 = summary, 2 = both
Public sgMissedText As String           'missed spots from curent week, current month, etc.
Public igUserCode As Integer           'user code id requesting snapshot
Public igVehCode As Integer            'vehicle code of snapshot requested
Public igStartOfWk(0 To 1) As Integer   'week of spot screen snapshot
Public igNowDate(0 To 1)  As Integer     'generation date of pre-pass file
Public igNowTime(0 To 1) As Integer       'generation time of pre-pass file
Public lgNowTime As Long                '10-20-01
Public sgTime As String
Public igODFGenDate(0 To 1) As Integer  '5-25-01
Public igODFGenTime(0 To 1) As Integer  '5-25-01
Public lgStartingCntrNo As Long
Public lgOrigCntrNo As Long
Public igYear As Integer                'budget year used for filtering
Public lgPrintedCnts() As Long             'table to maintain the contr pointers
Public igMonthOrQtr As Integer          'entered month or qtr
Public igPdStartDate(0 To 1) As Integer
Public sgPdType As String * 1
Public igRCSelectedIndex As Integer         'selected r/c index
Public igPeriods As Integer             '3-4-02 # periods requested for Billed & Booked
Public igSummaryID As Integer           'summary id # : (5-9)
Public igBSelectedIndex As Integer          'selected budget index
Public igBYear As Integer                   '11-30-16 budget year
Public igBFCSelectedIndex As Integer        'selected budget forecast index
Public igNoCodes As Integer
Public igcodes() As Integer
Public sgLogStartDate As String
Public sgLogNoDays As String
Public sgLogUserCode As String
Public sgLogStartTime As String
Public sgLogEndTime As String
Public sgRnfRptName As String * 3           'Report name from RNF file (L01, L02,... C01, C02,.. ..etc)
Public sgLogType As String * 1              'l=log, c=cp, o = other
Public tgVof As VOF                         'vehicle options for log/cps
Public igRARcfCode As Integer               '1-24-01
Public igBR_NTRDefined As Integer           '10-16-02
Public igBR_SchLinesExist As Integer        '12-13-02
Public igBR_SSLinesExist As Integer         '12-16-03
Public igBR_CPMDefined As Integer           '1-8-21
Public sgSelection As String                '7-13-16  save report selection (formula) for crystal
Public sgSelectionToAdd As String           '11-17-16  extra filtering on Invoice selection, (ignore edi or include edi)
Public bgSendPDF As Boolean                 '1-5-17 ; 6-22-17    OK to send pdf invoices to agency on finals or reprints (they may be send automatically (feature) or manually by user
Public bgSendSelevtivePDF As Boolean        'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist
Public bsSelectedEmailInvoices As String    'TTP 10826 / TTP 10813 - PDF invoice - selective invoice feature checklist - list of Invoices to not Display because they will be emailed
Public tgAirNameCode() As SORTCODE
Public sgAirNameCodeTag As String
Public tgAirNameCodeLg() As SORTCODE
Public tgAirNameCodePJ() As SORTCODE
Public tgCSVNameCode() As SORTCODE
Public sgCSVNameCodeTag As String
Public tgSellNameCode() As SORTCODE
Public sgSellNameCodeTag As String
Public tgSellNameCodeLg() As SORTCODE
Public tgSellNameCodeRA() As SORTCODE
Public sgSellNameCodeTagRA As String
Public tgSOCode() As SORTCODE
Public sgSOCodeTag As String
Public tgSOCodeCT() As SORTCODE
Public sgSOCodeTagCT As String
Public tgSOCodeAA() As SORTCODE
Public tgSOCodeCB() As SORTCODE
Public tgMultiCntrCode() As SORTCODE
Public sgMultiCntrCodeTag As String
Public tgMultiCntrCodeAD() As SORTCODE
Public sgMultiCntrCodeTagAD As String
Public tgMultiCntrCodeAP() As SORTCODE
Public sgMultiCntrCodeTagAP As String
Public tgMultiCntrCodeCB() As SORTCODE
Public sgMultiCntrCodeTagCB As String
Public tgMultiCntrCodeCT() As SORTCODE
Public sgMultiCntrCodeTagCT As String
Public tgMultiCntrCodeIA() As SORTCODE
Public sgMultiCntrCodeTagIA As String
Public tgManyCntCodeCB() As SORTCODE
Public tgManyCntCodeCT() As SORTCODE
Public sgManyCntCodeTagCT As String
Public tgMnfCodeCB() As SORTCODE
Public sgMnfCodeTagCB As String
Public tgMnfCodeCT() As SORTCODE
Public sgMNFCodeTagCT As String
Public sgMNFCodeTagLg As String
Public tgRptSelPjAgencyCode() As SORTCODE
Public tgRptSelPjSalespersonCode() As SORTCODE
Public tgRptSelPjAdvertiserCode() As SORTCODE
Public tgRptSelPjNameCode() As SORTCODE
Public tgRptSelPjBudgetCode() As SORTCODE
Public tgRptSelPjDemoCode() As SORTCODE
Public tgRptSelInAgencyCode() As SORTCODE
Public tgRptSelInSalespersonCode() As SORTCODE
Public tgRptSelInAdvertiserCode() As SORTCODE
Public tgRptSelInNameCode() As SORTCODE
Public tgRptSelInBudgetCode() As SORTCODE
Public tgRptSelInDemoCode() As SORTCODE
Public tgRptSelLgDemoCode() As SORTCODE
Public tgRptSelLgAgencyCode() As SORTCODE
Public tgRptSelLgSalespersonCode() As SORTCODE
Public tgRptSelLgAdvertiserCode() As SORTCODE
Public tgRptSelLgNameCode() As SORTCODE
Public tgRptSelLgBudgetCode() As SORTCODE
Public tgRptAgencyCode() As SORTCODE
Public sgRptAgencyCodeTag As String
Public tgRptSalespersonCode() As SORTCODE
Public sgRptSalespersonCodeTag As String
Public tgRptAdvertiserCode() As SORTCODE
Public sgRptAdvertiserCodeTag As String
Public tgRptNameCode() As SORTCODE
Public sgRptNameCodeTag As String
Public tgNamedAvail() As SORTCODE
Public tgMkMnf() As MNF 'Market
Public tgMNFCodeRpt() As SORTCODE
Public sgMNFCodeTagRpt As String
Public tgMcfCode() As SORTCODE          '4-27-07
Public sgMcfTagRpt As String
Public tgMCF() As MCF
Public tgItfCode() As SORTCODE

Public tgRegionCode() As SORTCODE       '7-2-10
Public sgRegionCodeTag As String

'Invoice
Public tgChfAdvtExt() As CHFADVTEXT
Public tgSortKey() As TYPESORTKEY
Public tgSortNoKey() As TYPESORTNOKEY
Public tgSelSort() As SELTYPESORT
Public tgRvf() As RVF
Public tgSSMnf() As MNF             'Sales Source
Public tgMRMnf() As MNF             'Missed Reason
Public tgNTRMnf() As MNF            '12-13-02 NTR invoicing
'Public tgMkMnf() As MNF             'Market
Public tgRvfByVef() As RVFVEF
Public tgInvVehCode() As VEHSORT     'Vehicle name and "\" code
Public tgUpdateAgf() As UPDATEADFAGF
Public tgUpdateAdf() As UPDATEADFAGF
Public tgPkLnGen() As PKLNGEN
Public tgSdfExtIn() As SORTSDFEXT
Public tgMGSdfExt() As SORTSDFEXT
Public tgLnSdfExtSort() As SDFEXTSORT
Public tgLnSdfExt() As SDFEXT
Public tgMktTotal() As MKTTOTAL
Public tgExportCount() As EXPORTCOUNT
Public bgUsingSockets As Boolean

Public tgSdfExtIn32() As SORTSDFEXT
Public tgEDIServiceInfo() As EDISERVICEINFO


'Remote User General info
Public sgRUExportPath As String  'Path to Export directory
Public sgRUImportPath As String   'Path to Import Directory
Public igRUStartExeCopy As Integer    'False=Don't start ExeCopy when Traffic terminated; True=Start ExeCopy when Traffic terminated
                                    'This is set to True only if Remote User performs a Remote User Export/Import operation
'Remote User Standalone Info
Public sgIniPath As String  'Traffic.Ini Path at Central System
Public igRUMode As Integer  'Remote User Mode of operations
                            '(0=Dedicated PC on Central System; 1= Direct updates- no PC on Central System; 2= Manual)
                            'Manual technique:
                            'Remote changes to Central
                            '   At Remote site- signon as remote user, change and/or add items, export
                            '                   Move export to central import
                            '   Central site- signon as superuser or planner and import
                            'Central changes to Remote
                            '   At Central site- signon as planner, change and/or add items
                            '                    signon as remote user, export
                            '                    move export to remote import directory
                            '   at remote site- signon as remote user and import file from central
'Remote User PC on Central System Info
Public lgRUCheckSec As Long   'Time in sec's to check for command file from RU (max is 60000)
Public lgRUTimeOut As Long   'Time in sec's to wait until indicating link not functioning
Public sgRUTimeOut As String    'Formated time
Public igRUXferWorking As Integer   'True=Transfer Working; False = Don't transfer working type proposals
Public igRUDialUp As Integer   'True=Dial-up connection; False = Direct connection (no dialing required)
Public sgRUCmmlPath As String       'Path to Command file directory
Public igDirectCall As Integer  '1=Direct Call processing; 0=Normal call mode

Public igMergeCallSource As Integer 'Merge call source
Public igMergeAdfCode As Integer    'AdfCode when Merging Products for an advertiser

Public igRptCallType As Integer  'Task number or List number of program calling the general report module
                                '-1=Unused, set this way for individual report modules
Public igRptType As Integer      'Which report within Calling type should be selected on initialization (used in RptSel only)
                                '-1=Unused, set this way for individual report modules
Public igRptReturn As Integer   'True=Show list; False = Exist
Public sgSUName As String
Public sgCPName As String

Public igGameReturn As Integer      'When returning from CGameSch:False or True
Public igGameInvReturn As Integer   'When returning from CGameInv:False=Cancelled or True=Changed
Public igLLFeedReturn As Integer   'When returning from LLFeed: True or False
Public sgLLFeedEndTime As String    'End time of current feed if igLLFeedReturn = True
Public sgLLFeedStartTime As String  'Start Time of Current feed is igLLFeedReturn = True
Public sgLLCurrentFeed As String    'Current Feed
Public igLLCurrentVefCode As Integer

Public igWinStatus(0 To 69) As Integer  'Composite of all Vehicle sWin status: used to enable/disable jobs/lists
                                        '0=Hide; 1=View;2=Input
Public igShowCredit As Integer      ' True = credit screen is showing
'Public igNoLJWinRes(0 To 69) As Integer 'Number of List/Job window resources (number of controls on form)
'Public igNoExeWinRes(0 To 30) As Integer   'Number of non List/Job window resources (number of controls on form)
Public igStopCancel As Integer  'Disallow return to signon (all records not saved)
Public igExitTraffic As Integer 'Set by Sigon when exit selected from the form
Public igManUnload As Integer   'Force unload even if file (records) not saved
Public igNoMenuDisabled As Integer  'Count of number of times menu disabled

Public igBasicOrAux As Integer  'Last screen up 1= basic or from signon, 2= aux
Public igJobShowing(0 To 13) As Integer     'Flag indicating if Job Task is visible
Public igBasic10Showing As Integer
Public igAuxWndShowing As Integer
Public igSportsSystem As Integer    'Limit function- 1=Sports: Only Enter Proposal, Enter Inventory and Reports; 2=Live Log; 3=Station Manually posting times
Public sgStationCallLetters As String
Public sgStationPW As String
Public lgMtfNoRecs As Long  'Number of MTF records, used to determine if MTF in uses
Public lgDpfNoRecs As Long  'Number of DPF records, used to determine if DPF in uses

'Find Match information in Combo or List Box
Public igFoundRow As Integer 'Row found with gFndFirst or gFndNext
Public cgList As String    'Control associated with igFoundRow

'Error messages
Public igBtrError As Integer    'Btrieve error return
Public sgErrLoc As String   'Location Description (1-34 in Schedule.Bas; 101-110 in Spots.Frm)

'General Message (GenMsg)
Public sgGenMsg As String       'Message to be shown
Public sgCMCTitle(0 To 3) As String    'Button Titles (set title to blank if not to be shown)
Public igDefCMC As Integer      'Default Button Number (0, 1 or 2)
Public igAnsCMC As Integer      'Button Selected (0, 1, 2 or 3)
Public igEditBox As Integer     '0=No; 1=Yes
Public sgEditValue As String
Public igEditBoxMaxCharacters As Integer
Public lgEditBoxMaxValue As Integer
'9459
Public bgUseRadioButtons As Boolean
'Calendar info
Public igCalActive As Integer   '0=no; 1=yes
Public fgCalLeft As Single      'Month calendar left position
Public fgCalTop As Single       'Month calendar top position
Public igCalMonth As Integer    'View month for month calendar
Public igCalYear As Integer     'View year (101-1999)
Public igCalMonthType As Integer    '0=Std;1=Cal;2=Julian+;3=Julian-;4=Corp
Public igCalByYear As Integer       'True=Show true year; False=Show one month
Public igCalAcross As Integer       'True=Show Jan, Feb, Mar, Apr acorss top
                                    'False = Show Jan, Feb, Mar down

'Calculator
Public igCalcActive As Integer   '0=no; 1=yes
Public fgCalcLeft As Single      'calculator left position
Public fgCalcTop As Single       'calculator top position

'Number pad input
Public fgPadDeltaY As Single
Public fgPadDeltaX As Single
Public fgPadMinY As Single
Public fgPadMinX As Single

'Get Path
Public igGetPath As Integer '0=Done pressed; 1=Cancel pressed
Public sgGetPath As String
Public igPathType As Integer    '0=Any; 1=subfolder only
Public lgCallLeft As Long
Public lgCallTop As Long
Public lgCallHeight As Long
Public lgCallWidth As Long

Public sgBS As String * 1   'Backspace
Public sgTB As String * 1   'Tab
Public sgLF As String * 1   'Line feed (new line)
Public sgCR As String * 1   'Carriage return
Public sgRUCrLf As String * 2 'CR/LF substitute string for cr/lf remote user cxf message transfer

Public sgCnCDelimiter As String * 1 'Export CnC: Delimiter (Tab or ,)
Public sgCnCNameBracket As String   'Export CnC: Enclose Names with this character (blank or ")

Public fgOffset As Single       'Distance control must be moved within outline box
Public fgBoxInsetX As Single    'Margin from box outline to text
Public fgBoxInsetY As Single    'Margin from box outline to text
Public fgBoxStH As Single       'Height of standard box
Public fgBoxAddH As Single      'Height of address box
Public fgBoxNmAddH As Single    'Height of name and address box
Public fgBoxGridH As Single     'Height of grid box
Public fgFlexGridRowH As Single 'Height of row within Flex Grid so that edit box will fit
Public fgBevelX As Single       'Width in X direction of bevel on a panel
Public fgBevelY As Single       'Width in Y direction of bevel on a panel
Public fgXBoxToBox As Single    'Distance from box to box along x direction
Public fgStDeltaY As Single     'Distance from standard control top to control top
Public fgAddDeltaY As Single    'Distance from address control to top of control
Public fgAddNmDeltaY As Single  'Distance from address/name control to top of control
Public fgBoxNoLabelH As Single  'Height of picture area without label (including bounders)
Public fgPanelAdj As Single     'Adjustment to modeless window (without caption) enclosing panel
Public fgListHtArial825 As Single   'Height of list item within arial 8.25 font
Public fgListHtSerif825 As Single   'Height of list item within MS San Serif 8.25 font
Public igScrollBarWidth As Integer  'Width of vertical scroll bar
Public igAvgCharWidth As Integer
Public lgPercentAdjH As Long        'Percentage of adjustment in for Height
Public lgPercentAdjW As Long        'Percentage of adjustment in for Widrg
Public hgDB As Integer  'Database handle
Public hgHlf As Integer 'Help handle
Public hgAuf As Integer 'Alert Handle
Public hgUlf As Integer
Public hgChf As Integer 'Alert contract check
Public hgVsf As Integer 'Alert contract check
Public hgVpf As Integer
Public igAlertTimer As Integer 'Number of minutes since last checked Alerts
Public igAlertFlash As Integer 'Number of times Alert as Flashed
Public igAlertCntrStatus As Integer
Public igDashboardCntrStatus As Integer 'TTP 10937 - Contract Dashboard: new screen to show contract header information
'  Value   Source    Screen       Definition
'    0                            Not from Alter or Spots
'    1     Alert     Proposal     Revise proposal (chfCntRevNo = 0)
'    2     Alert     Proposal     Revise hold or order (chfCntRevNo > 0)
'    3     Alert     Proposal     View
'    4     Alert     Order        Revise
'    5     Alert     Order        View
'    6     Spot      Order        Revise
'    7     Spot      Order        View
Public lgAlertCntrNo As Long
Public igOKToCallCntr As Integer
Public sgActiveLogDate As String  'Last date the ulfActiveLogDate updated
Public sgActiveLogTime As String  'Last time the ulfActiveLogTime updated
Public lgActiveUlfCode As Long
Public igAlertSpotStatus As Integer '0=Not from Alerts; 1=From Alerts
Public sgAlertSpotVehicle As String
Public sgAlertSpotMoDate As String
Public hgSite As Integer 'Affiliate Site

'List Sort
Public tgAdvertiser() As SORTCODE
Public sgAdvertiserTag As String
Public tgAgency() As SORTCODE
Public sgAgencyTag As String
Public tgAvailCode() As SORTCODE
Public sgAvailCodeTag As String
Public tgBookNameCode() As SORTCODE
Public sgBookNameCodeTag As String
Public tgBusCatCode() As SORTCODE
Public sgBusCatCodeTag As String
Public tgBuyerCode() As SORTCODE
Public sgBuyerCodeTag As String
Public tgCntrCode() As SORTCODE
Public sgCntrCodeTag As String
Public tgCommentCode() As SORTCODE
Public sgCommentCodeTag As String
Public tgCompCode() As SORTCODE
Public sgCompCodeTag As String
Public tgDemoCode() As SORTCODE
Public sgDemoCodeTag As String
Public tgCDemoCode() As SORTCODE
Public sgCDemoCodeTag As String
Public tgExclCode() As SORTCODE
Public sgExclCodeTag As String
Public tgManager() As SORTCODE
Public sgManagerTag As String
Public tgNameCode() As SORTCODE
Public sgNameCodeTag As String
Public tgPayableCode() As SORTCODE
Public sgPayableCodeTag As String
Public tgPlanner() As SORTCODE
Public sgPlannerTag As String
Public tgPotCode() As SORTCODE
Public sgPotCodeTag As String
Public sgSegCodeTag As String
Public tgSegCode() As SORTCODE
Public tgProdCode() As SORTCODE
Public sgProdCodeTag As String
Public tgRateCardCode() As SORTCODE
Public sgRateCardCodeTag As String
Public tgResvCode() As SORTCODE
Public sgResvCodeTag As String
Public tgRevSetCode() As SORTCODE       'Two Dimension
Public sgRevSetCodeTag As String
Public tgSalesperson() As SORTCODE
Public sgSalespersonTag As String
Public tgShareBudgetCode() As SORTCODE  'Two Dimension
Public sgShareBudgetCodeTag As String
Public tgSocEcoCode() As SORTCODE
Public sgSocEcoCodeTag As String
Public tgSPersonCombo() As SORTCODE
Public sgSPersonComboTag As String
Public tgShtTitleCode() As SORTCODE
Public sgShtTitleCodeTag As String
Public tgTmpSort() As SORTCODE
Public sgTmpSortTag As String
Public tgUser() As SORTCODE
Public sgUserTag As String
Public tgPopUrf() As URF
Public sgPopUrfTag As String
Public tgUserVehicle() As SORTCODE
Public sgUserVehicleTag As String
Public tgRCUserVehicle() As SORTCODE
Public sgRCUserVehicleTag As String
Public tgVehCombo() As SORTCODE
Public sgVehComboTag As String
Public tgVehicle() As SORTCODE
Public sgVehicleTag As String
Public tgLogVehicle() As SORTCODE
Public sgLogVehicleTag As String
Public tgMktCode() As SORTCODE
Public sgMktCode As String
Public tgUserSortCode() As SORTCODE     'ust & urf
Public sgUserSortCodeTag As String
Public igTrafficFirstActivate As Integer 'TTP 10937 - Contract Dashboard: new screen to show contract header information
'***** when adding to the above list- add names to gEraseGlobalVar within InitSubs


Public tgSpf As SPF 'Site Option file
Public sgSpfStamp As String 'Date/Time stamp of file
Public tgSpfx As SPFX 'Site Option Extension file

Public tgUrf() As URF 'User Option file
Public sgUrfStamp As String 'Date/Time stamp of file
Public igUserByVeh As Integer   'True=A User defined for vehicle; False=All user defined for all vehicles

Public tgVpf() As VPF 'Vehicle Option file
Public sgVpfStamp As String 'Date/Time stamp of file

Public tgVff() As VFF 'Vehicle feature file
Public sgVffStamp As String 'Date/Time stamp of file

Public tgJobHelp() As HLF   'Job one line help
Public tgListHelp() As HLF  'List one line help

Public lgVehComboCode() As Long  'List of allowed combo Vehicles as positive numbers
Public sgVsfStamp As String     'Time stamp of list

Public igVirtVefCode() As Integer  'List of Virtual Vehicles (code)
Public sgVVefStamp As String     'Time stamp of list

Public sgCntrForDateStamp As String

'Advertiser/Agency
Public igPopAdfAgfDormant As Integer 'True=Include dormant advt/agy in gPopAdvtBox and gPopAgyBox;
                                         'False=Exclude dormant when population advt/agy
Public igPopExternalAdvt As Integer 'True=Include external invoiced advertiser in gPopAdvtBox;
                                         'False=Exclude external invoiced advertiser when population advt

'Vehicle
Public tgMVef() As VEF
Public sgMVefStamp As String
Public igDormantVef() As Integer

'Corporate Calendar
Public tgMCof() As COF
Public sgMCofStamp As String

'Salesperson
Public tgMSlf() As SLF
Public sgMSlfStamp As String

'Tax Rate
Public tgTrf() As TRF
Public sgTrfTag As String

'Demo Book cross reference of Populations
Public tgPopDnfInfo() As POPDNFINFO
Public sgPopDnfInfoTag As String
'Books defined as using USA Population and site 'Allow Research Estimates is check
Public igEstByUSADnfCode() As Integer

'Rate Card
Public tgMRcf() As RCF  'Master file containing all records from RCF
Public tgMRif() As RIF  'Master file containing all records from RIF
Public tgMRdf() As RDF  'Master file containing all records from RDF
Public sgMRcfStamp As String
Public sgMRifStamp As String
Public sgMRdfStamp As String

'Contract data
Public sgCntrScreen As String 'Screen Title
Public tgChf As CHF            'CHF record image
Public tgClf() As CLFLIST      'CLF record image
Public tgCff() As CFFLIST      'CFF record image
Public tgPcf() As PCFLIST      'PCF record image
'Contract area
Public tgChfCntr As CHF            'CHF record image
Public tgClfCntr() As CLFLIST      'CLF record image
Public tgCffCntr() As CFFLIST      'CFF record image
Public tgCgfCntr() As CGFLIST      'CGF record image
Public tgMsfCntr() As MSFLIST      'MSF record image
Public tgMgfCntr() As MGFLIST      'MGF record image
Public tgLnAvailInfo() As LNAVAILINFO

Public sgMsfCntrCxf() As String
'Schedule area
Public lgOpenSchFileCount As Long   '5/2/19: Only open and close files when count is zero
Public tgChfSchd As CHF            'CHF record image
Public tgClfSchd() As CLFLIST      'CLF record image
Public tgCffSchd() As CFFLIST      'CFF record image
'Import area
Public tgClfImpt() As CLFLIST      'CLF record image
Public tgCffImpt() As CFFLIST      'CFF record image

Public tgChfCntrProj As CHF            'CHF record image
Public tgClfCntrProj() As CLFLIST      'CLF record image
Public tgCffCntrProj() As CFFLIST      'CFF record image
Public tgChfBud As CHF            'CHF record image
Public tgClfBud() As CLFLIST      'CLF record image
Public tgCffBud() As CFFLIST      'CFF record image
Public tgClfPostLog() As CLFLIST      'CLF record image
Public tgCffPostLog() As CFFLIST      'CFF record image
Public tgClfCopy() As CLFLIST      'CLF record image
Public tgCffCopy() As CFFLIST      'CFF record image
Public tgChfRC As CHF            'CHF record image
Public tgClfRC() As CLFLIST      'CLF record image
Public tgCffRC() As CFFLIST      'CFF record image
Public tgChfSpot As CHF            'CHF record image
Public tgClfSpot() As CLFLIST      'CLF record image
Public tgCffSpot() As CFFLIST      'CFF record image
Public tgChfInv As CHF            'CHF record image
Public tgClfInv() As CLFLIST      'CLF record image
Public tgCffInv() As CFFLIST      'CFF record image
Public tgChfRep As CHF            'CHF record image
Public tgClfRep() As CLFLIST      'CLF record image
Public tgCffRep() As CFFLIST      'CFF record image

Public bgBalanceContract As Boolean 'True=Compute Spot totals and Prices; False=Freeze the computing of spot totals and Prices

'For reports
Public tgChfAA As CHF            'CHF record image
Public tgClfAA() As CLFLIST      'CLF record image
Public tgCffAA() As CFFLIST      'CFF record image
Public tgChfAC As CHF            'CHF record image
Public tgClfAC() As CLFLIST      'CLF record image
Public tgCffAC() As CFFLIST      'CFF record image
Public tgChfAD As CHF            'CHF record image
Public tgClfAD() As CLFLIST      'CLF record image
Public tgCffAD() As CFFLIST      'CFF record image
Public tgChfAP As CHF            'CHF record image
Public tgClfAP() As CLFLIST      'CLF record image
Public tgCffAP() As CFFLIST      'CFF record image
Public tgChfAS As CHF            'CHF record image
Public tgClfAS() As CLFLIST      'CLF record image
Public tgCffAS() As CFFLIST      'CFF record image
Public tgChfAV As CHF            'CHF record image
Public tgClfAV() As CLFLIST      'CLF record image
Public tgCffAV() As CFFLIST      'CFF record image
Public tgChfCB As CHF            'CHF record image
Public tgClfCB() As CLFLIST      'CLF record image
Public tgCffCB() As CFFLIST      'CFF record image
Public tgChfCP As CHF            'CHF record image
Public tgClfCP() As CLFLIST      'CLF record image
Public tgCffCP() As CFFLIST      'CFF record image
Public tgChfCT As CHF            'CHF record image
Public tgClfCT() As CLFLIST      'CLF record image
Public tgCffCT() As CFFLIST      'CFF record image
Public tgClfED() As CLFLIST      'CLF record image
Public tgCffED() As CFFLIST      'CFF record image
Public tgChfID As CHF            'CHF record image
Public tgClfID() As CLFLIST      'CLF record image
Public tgCffID() As CFFLIST      'CFF record image
Public tgChfPA As CHF            'CHF record image
Public tgClfPA() As CLFLIST      'CLF record image
Public tgCffPA() As CFFLIST      'CFF record image
Public tgChfPC As CHF            'CHF record image
Public tgClfPC() As CLFLIST      'CLF record image
Public tgCffPC() As CFFLIST      'CFF record image
Public tgChfPS As CHF            'CHF record image
Public tgClfPS() As CLFLIST      'CLF record image
Public tgCffPS() As CFFLIST      'CFF record image

Public tgChfRI As CHF            'CHF record image  '2-22-02
Public tgClfRI() As CLFLIST      'CLF record image
Public tgCffRI() As CFFLIST      'CFF record image
Public tgChfRV As CHF            'CHF record image  '5-1-02
Public tgClfRV() As CLFLIST      'CLF record image
Public tgCffRV() As CFFLIST      'CFF record image
Public tgChfSP As CHF            'CHF record image
Public tgClfSP() As CLFLIST      'CLF record image
Public tgCffSP() As CFFLIST      'CFF record image
Public tgChfUS As CHF            'CHF record image
Public tgClfUS() As CLFLIST      'CLF record image
Public tgCffUS() As CFFLIST      'CFF record image
Public igTerminateReturn As Integer 'When returning from Terminate screen:0=Cancelled;1=Continue
'Public tgRcf As RCF     'Rcf record image (required so RateCard.frm and RCTerms.Frm can communicate)
'Public tgRpf() As RPF        'Rpf record image
'Public tgRif() As RIF
'Public tgRdf() As RDF
Public lgSchChfCode As Long    'Contract Code for schedule contract
Public igLineNo As Integer  'Line number currently selected- required by CAvail
Public igComp1Code As Integer   'Competitive code 1 from header
Public igComp2Code As Integer   'Competitive code 2 from header
Public tgFBSbf() As SBFLIST    'SBF record image of Fix bill
Public tgIBSbf() As SBFLIST    'SBF record image of Item billing
Public tgMBSbf() As SBFLIST    'SBF record image of Merchandising
Public tgPBSbf() As SBFLIST    'SBF record image of Promotion
Public sgFBBTotal As String 'Package billing amount
Public sgIBPTotal As String 'Item posted amount
Public sgIBBTotal As String 'Item billed amount
Public sgMBBTotal As String 'Merchandising billing amount
Public sgPBBTotal As String 'Promotion billing amount
Public igPkgChgd As Integer
Public igDirAdvt As Integer 'Communication from Contract to Bill: True=Direct Advt; False=Bill agency
Public igShiftDates As Integer  'True=Allow date shifting; False=Disallow date shifting
Public sgPassValue As String
Public igCloneLn As Integer
Public sgCloneDate As String
Public igCloneRcfCode As Integer
Public sgMktClusterDef As String * 1 'Any Market cluster defined (Y or N).  If True, then test if only same
                                       'market used within a contract
Public sgRepDef As String * 1 'Any rep vehicle defined (Y or N).  If True, then test if only
                                       'rep and clusters used within a contract
Public igMktVefCode() As Integer    'VefCode for each vehicle that has Market Cluster = Yes
Public igRepVefCode() As Integer    'VefCode for each vehicle that has Vehicle type = R
Public igPostType As Integer '0=Time; 1=Counts by Station; 2=Counts by Advt; 3=Received; 4=Cluster; 5=Date/Time by Network Name; 6=Date/Time by Vehicle
'Rate Card data
'Public tgRcfI As RCF     'Rcf record image (required so RateCard.frm and RCTerms.Frm can communicate)
'Public igRcfChg As Integer  'True=Rcf Changed; False=Not changed
'Public igRCMode As Integer  'When calling RCTerms:0=New; 1=Change
'Public igRCReturn As Integer 'When returning from RCTerms:0=Cancelled;1=Continue
'Public tgRpfI() As RPF        'Rpf record image
'Public sgReCalcSave(1 To 12) As String 'Values retained from recalc grid area
'Public igReCalcReturn As Integer    '0=Cancelled; 1= Recalculate
Public sgRCStamp As String

Public tgCommAgf() As AGFEXT
Public sgCommAgfStamp As String 'Date/Time stamp of file
Public tgCommAdf() As ADFEXT
Public sgCommAdfStamp As String 'Date/Time stamp of file
Public igInternalAdfCount As Integer
Public tgCompMnf() As MNFCOMPEXT
Public sgCompMnfStamp As String 'Date/Time stamp of file
Public tgRevSet() As MNF
Public sgRevSetStamp As String
Public tgPotMnf() As MNF
Public sgPotMnfStamp As String
Public tgBusCatMnf() As MNF
Public sgBusCatMnfStamp As String
Public tgExclMnf() As MNF
Public sgExclMnfStamp As String
Public tgShareBudget() As MNF
Public sgShareBudgetStamp As String
Public tgDemoMnf() As MNF
Public sgDemoMnfStamp As String
Public tgSocEcoMnf() As MNF
Public sgSocEcoMnfStamp As String
Public tgAvailAnf() As ANF
Public sgAvailAnfStamp As String
Public tgMktMnf() As MNF
Public sgMktMnfStamp As String

'
'Public tgStations() As SHTT
'Public sgStationsStamp As String
'Public tgMarkets() As MKT
'Public sgMarketsStamp As String
'Public tgOwners() As ARTT
'Public sgOwnersStamp As String
'
Type REGIONINTSORT
    iIntCode As Integer
    iShttCode As Integer
End Type
Type REGIONSTRSORT
    sStr As String * 40
    iShttCode As Integer
End Type

Public tgMktSort() As REGIONINTSORT
Public tgMSAMktSort() As REGIONINTSORT
'Public tgOwnerSort() As REGIONINTSORT
'Public tgZipSort() As REGIONSTRSORT
Public tgStateSort() As REGIONSTRSORT
Public tgFmtSort() As REGIONINTSORT
Public tgTztSort() As REGIONINTSORT

'Program names
Public igPrgNameVefCode As Integer
Public tgPaf() As PAF

'Sports
Public igGameSchdVefCode As Integer
Public lgSeasonGhfCode As Long
Public igGameLibReturn As Integer   'When returning from GameLib:False=Cancelled or True=Changed

'Get Games
Public igGetGameVefCode As Integer
Public igGetGameGhfCode As Integer
Public igGetGameIhfCode As Integer
Public igGetGameReturn As Integer   'Return from GetGames:False = Cancelled; True = Selections defined
Public igGetGameDefaultUnits As Integer

Public igGetCopyGridAdfCode As Integer
Public igGetCopyGridGnfCode As Integer
Public igGetCopyGridReturn As Integer   'Return from GetCopyGrid:False = Cancelled; True = Selections defined
Public bgGetCopyGridIncludeCopy As Boolean
Public bgAllowCopyGrid As Boolean
Public bgStationAreVehicles As Boolean

Public sgCntrVerifyType As String   'N=Not Verified; S=Sent to Agency; V=Verified

'5676
Public sgRootDrive As String
'8723
Public ogContractCreator As CContractCreation
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
'7967 turn on/off here  change to 15/16
'D.S. 03/21/18 16 to 17
Public tgTaskInfo(0 To 20) As TASKINFO      '1-23-20 expand from 17 to 20
Public lgDateMonitorChecked As Long
Public lgTimeMonitorChecked As Long


Type GETGAMERETURN
    iGameNo As Integer
    iNoUnits As Integer
    lRate As Long
    lCost As Long
    lIsfCode As Long
    sBilled As String * 1
    iNoUnitsOrdered As Integer
    iNoUnitsProp As Integer
End Type
Public tgGetGameReturn() As GETGAMERETURN 'Game numbers
'Conflict information
'Public igVcfUsed As Integer 'True=tgVcf contains conflict records
                            'False=no records within tgVcf
'Public tgVcf0() As VCF      'M-F Vcf records
'Public tgVcf5() As VCF      'Sa Vcf records
'Public tgVcf6() As VCF      'Su Vcf records

'DAF stamp
'Public sgDafStamp As String

'Spot form data
Public sgSpotCallType As String * 1 'T=Time; A=Advertiser
Public igSpotLineReturn As Integer    '0=Cancelled; >0 = Done (line #)

'Vehicle communication parameters
'Public sgVsfCallType As String * 1'F=Vehicle; S=Salesperson;
'Public igVsfCallSource As Integer   'Combo call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCE----=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
'Public sgVsfName As String          'Name from/to modules
'Public Combo_S As New Combo_V

'MultiName communication parameters
Public sgMnfCallType As String * 1 'I=Item Billing; A=Announcer; P=Sports; S=Sales Source
    'T=Sales Team; R=Revenue Sets; M=Missed Reason; C=Competitive; N=Network; V=Invoice sorting
    'X=Exclusions
Public igMNmCallSource As Integer   'MultiName calling source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCESALESOFFICE=who made call
                                    'CALLSOURCESALESPERSON=who made call
                                    'CALLSOURCEAGENCY=who made call
                                    'CALLSOURCEADVERTISER=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public sgMNmName As String          'Name from/to modules

'Vehicle communication parameters
Public igVehCallSource As Integer   'Vehicle call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCEENAME=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public sgVehName As String       'Name from Vehicle call

Public igPodTargetCallSource As Integer     'Podcast Target call source and return status
                                            'CALLNONE=secondary
                                            'CALLSOURCEENAME=who made call
                                            'CALLDONE=call completede
                                            'CALLCANCELLED=sales source cancelled
                                            'CALLTERMINATED=sales source terminated because of an error
Public sgPodTargetName As String       'Name from Podacst Target call

'Usere communication parameters
Public igUrfCallSource As Integer   'Vehicle call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCEENAME=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public sgUrfName As String       'Name from Vehicle call

'Lock communication parameters
Public igLockCallSource As Integer   'Lock call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCEENAME=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public sgLockName As String         'Name from Lock call-Vehicle code

'Unschedule communication parameters
Public igUnSchdCallSource As Integer   'Unschedule call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCEENAME=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error

'Site communication parameters
Public igSiteCallSource As Integer   'Vehicle call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCEENAME=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error

'Event type communication parameters
Public igETypeCallSource As Integer   'Event Type call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCEENAME=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public sgETypeName As String        'Name from event type call
Public igEtfRecsCreated As Integer  'True=Event records created; False=No records created

'Event name communication parameters
Public igENameCallSource As Integer   'Event Name call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCEENAME=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=program event cancelled
                                    'CALLTERMINATED=program event terminated because of an error
Public sgENameName As String        'Name from event name call

'Inventory type communication parameters
Public igInvTypeCallSource As Integer   'Inventory Type call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCEENAME=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public sgInvTypeName As String        'Name from event type call

'Inventory Item communication parameters
Public igInvItemCallSource As Integer   'Inventory Type call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCEENAME=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public sgInvItemName As String        'Name from event type call

'Avail names communication parameters
Public igANmCallSource As Integer   'Avail Name call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCE-----=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public sgANmName As String          'Name from avail call

Public igRNCallSource As Integer   'Avail Name call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCE-----=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public sgRNName As String          'Name from avail call
Public sgRNCallType As String           'R=Rep; N=Network

'Avail names communication parameters
Public igFdNmCallSource As Integer   'Feed Name call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCE-----=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public sgFdNmName As String          'Name from Feed call

'Media code communication parameters
Public igMcdCallSource As Integer   'Media Name call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCE-----=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public sgMcdName As String          'Name from media call

'Contract comment communication parameters
Public igCmmCallSource As Integer   'Contract comment call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCE-----=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=cancelled
                                    'CALLTERMINATED=terminated because of an error
Public sgCmmTitle As String          'Title from comment

'Sales office communication parameters
Public igSofCallSource As Integer   'Sales Office call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCESALESPERSON=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales office cancelled
                                    'CALLTERMINATED=sales office terminated because of an error
Public sgSofName As String          'Name from/to sales office call

'Lock box and EDI service communication parameters
'Name/Address data
Public sgArfCallType As String * 1 'L=Lock Box; A=Agency;
'Public NmAd_A As New NmAd_L
Public igArfCallSource As Integer   'Sales Office call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCEAGENCY=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales office cancelled
                                    'CALLTERMINATED=sales office terminated because of an error
Public sgArfName As String          'Name from/to sales office call

'Salesperson communication parameters
Public igSlfCallSource As Integer   'Sales Office call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCEAGENCY=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales office cancelled
                                    'CALLTERMINATED=sales office terminated because of an error
Public sgSlfName As String          'Name from/to sales office call

'Agency communication parameters
Public igAgyCallSource As Integer   'Agency call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCEADVERTISER=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales office cancelled
                                    'CALLTERMINATED=sales office terminated because of an error
Public sgAgyName As String          'Name from/to agency call

'Advertiser communication parameters
Public igAdvtCallSource As Integer   'Advertiser call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCECONTRACT=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales office cancelled
                                    'CALLTERMINATED=sales office terminated because of an error
Public sgAdvtName As String          'Name from/to advertiser call

'Advertiser Product communication parameters
Public igAdvtProdCallSource As Integer   'Advertiser product call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCECONTRACT=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales office cancelled
                                    'CALLTERMINATED=sales office terminated because of an error
Public sgAdvtProdName As String     'Name from/to advertiser product call
'Daypart call parameters
Public igDPCallSource As Integer    '0=Traffic; 1=RateCard
'Personnel communication parameters
Public igPersonnelCallSource As Integer   'Personnel call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCECONTRACT=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales office cancelled
                                    'CALLTERMINATED=sales office terminated because of an error
Public sgPersonnelName As String     'Name from/to personnel call
'Copy inventory communication parameters
Public igCopyInvCallSource As Integer   'Avail Name call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCE-----=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error
Public lgCopyInvCifCode As Long          'Advertiser code from call
Public igCopyInvAdfCode As Integer          'Advertiser code from call
Public igCopyInvEtfCode As Integer
Public igCopyInvEnfCode As Integer
Public igCopyInvRotLen As Integer   'Rotation spot length
Public sgCopyForBBs As String       'Y=Yes; N= No.  If Y only show lengths less then 30 sec.
Public sgUseCartNo As String
Public igCopyInvMcfCode As Integer
'Copy Assign communication parameters
Public igCopyAsgnCallSource As Integer   'Avail Name call source and return status
                                    'CALLNONE=secondary
                                    'CALLSOURCE-----=who made call
                                    'CALLDONE=call completede
                                    'CALLCANCELLED=sales source cancelled
                                    'CALLTERMINATED=sales source terminated because of an error

'Program event communication paramters
Public igLibType As Integer '0=Regulars; 1=Specials; 2= Sports
Public igViewType As Integer '0=On Air; 1=Alternate
Public tgPrg() As PRGDATE   'Times/Date that library event is to run
Public sgVefTypeViaPrg As String * 2 'Type: C=Conventional without Feed; S=Selling; A=Airing; CF = Conventional with Feed
                                    'Note: Conventional with Vpf.iGMnfNCode(1)>0 is included with airing for creating delivery
Public lgEarliestDateViaPrg As Long      'Earliest allowed date (0 if not defined)
Public igVehIndexViaPrg As Integer  'Index into selected vehicle from program to program events
                            'can't use cbcVeh.ListIndex as its value is not always set
Public igPrgDupl As Integer 'True=Duplicate prg; False=New or version
Public tgRPrg() As PRGDATE   'Times/Date that library event is to removed from
Public lgLibLength As Long  'Library length (or 0 if unknown) Or LibCode from Duplicate
Public sgRemLibName As String  'Library name being removed or Duplicated
Public igRemLibReturn As Integer    '0=Cancelled; 1=Deleted
'Vehicle option communication parameters
Public sgVehNameToVehOpt As String   'Name from Vehicle call
Public igVehNewToVehOpt As Integer  'Indicates if new vehicle
Public igVehOptCallSource As Integer   'Vehicle option call source
                                    'CALLNONE=secondary
                                    'CALLSOURCEVEHICLE=who made call

Public igManualPostVefCode As Integer
Public sgManualPostDate As String

Public igLegalTime(0 To 8)  As Integer
Public igLegalLength(0 To 9) As Integer

Public igSlfFirstNameFirst As Integer   'True = in list boxes showing salesperson names, show first name first
                                        'False = in list boxes showing salesperson names, show last name first

Public igMnfMissed As Integer   'Missed reason code or zero
Public igDefaultMnfMissed As Integer    'Missed default reason when igMnfMissed = 0 or sdfMnfMissed = 0
Public igPLAnfCode As Integer   'Post Log added avail record

Public igLogMktCode()  As Integer   'Market code used to filter vehicles in gPopUserVehicleByMkt
Public igCntrMktCode()  As Integer   'Market code used to filter vehicles in gPopVehViaRateCardBox
Public igSpotMktCode()  As Integer   'Market code used to filter vehicles in gPopUserVehicleByMkt
Public igMktCode() As Integer       'Market codes obtained from call to LogMkt form
Public sgDateRangeStartDate As String 'start date from daterange form
Public sgDateRangeEndDate As String   'end date from daterange form

Type BBCOMPARE                      '4-15-09 Table of unique sales sources w/vehicles for B & B comparison/budgers
    iSSCode As Integer
    iVefCode As Integer
End Type

Public tgBBCompare() As BBCOMPARE

'12/16/14: Add Hide/Show Hidden Lines
Public bgHideHiddenLines As Boolean
Public Type MAPLINEPAINT
    lPaintLnRowNo As Long
    lMapLnRowNo As Long
    sLnType As String * 1
End Type
Public tgMapLinePaint() As MAPLINEPAINT
'7496
Public sgAudioExtension As String

'Advertiser Split ID
Public igAdvtSplitID As Integer '0=Cancel pressed; 1=Done pressed
Public sgAVIndicatorID As String
Public sgXDSCue As String
'8032-moved to help compile
Public igStationXmlChoice As Integer
'7967
Public dgWvImportLast As Date
Public igWVImportElapsed As Integer
Public sgTDBNames(1 To 12) As String 'Set in initsubs.bas

'L.Bianchi
Public sgManualPostCpmFilter As String
'10910
Public sgContractAttachmentPath As String
