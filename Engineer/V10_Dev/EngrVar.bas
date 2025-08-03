Attribute VB_Name = "EngrVarSubs"
'
' Release: 1.0
'
' Description:
'   This file contains the declarations

Option Explicit
'VB variables
'Public Const Modal = 1
'Public Const vbDefault = 0
'Public Const vbHourglass = 11
'Public Const Yes = 0
'Public Const No = 1
'Public Const vbYes = 6
'Public Const vbNo = 7

Public igOperationMode As Integer '0=Standard mode (running Engineering), 1=Background Mode (running EngrService)
Public igRunningFrom As Integer '0=Server; 1=Client
Public sgSpecialPassword As String
Public sgPasswordAddition As String 'Message to add to Previous Password

Public sgClientFields As String     'A=ABC; W=Westwood One; Blank or N = None

Public igJobStatus(0 To 3) As Integer
Public igJobShowing(0 To 3) As Integer

Public igListStatus(0 To 17) As Integer

Public igExitEngrMain As Integer
Public igJobVisible As Integer
Public igListVisible As Integer

Public lgCurrHRes As Long
Public lgCurrVRes As Long
Public lgCurrBPP As Long        '4=16 Colors; 8=256 colors; 16 High; 24 & 32 = True Color


Public sgEngrVersion As String
Public Const MAX_PATH = 260
Public igManUnload As Integer
Public igStopCancel As Integer

Public sgUserName As String

'General Message (GenMsg)
Public sgGenMsg As String       'Message to be shown
Public sgCMCTitle(0 To 3) As String    'Button Titles (set title to blank if not to be shown)
Public igDefCMC As Integer      'Default Button Number (0, 1 or 2)
Public igAnsCMC As Integer      'Button Selected (0, 1, 2 or 3)
Public igEditBox As Integer     '0=No; 1=Yes
Public sgMsgEditValue As String 'Message to show above the Edit Box
Public sgEditValue As String

'Communication to other modules
Public igInitCallInfo As Integer    '0=Call form List screen; 1 or 2 or .. indicate information from calling module
                                    'Control Character call: 1=Audio; 2=Bus
                                    'TimeType: 1=Start; 2=End
                                    'Day Names: 1=Library; 2=Template
Public sgInitCallName As String     'New name or blank
Public sgTempDescription As String  'Default description for Air Info
Public igReturnCallStatus As Integer 'CALLDONE; CALLCANCELLED; CALLTERMINATED
Public sgReturnCallName As String   'Name of item last edited or accessed
Public igRptDest As Integer         '0 = display, 1 = print, 2 = save to file (return path & filename in sgReturnCallName)
Public igExportType As Integer      '3-28-06
Public sgReturnOption As String     '3-28-06
Public igLibCallType As Integer '0=New; 1=Change; 2=Model; 3=View; 4=Terminate
Public lgLibCallCode As Long    'Library Call Code

Public igTempCallType As Integer '0=New; 1=Change; 2=Model
Public lgTempCallCode As Long    'Template Call Code

Public igSchdCallType As Integer    '0=New; 1=Change
Public sgSchdDate As String         'If igSchdCallType = 1
Public sgAsAirCompareDate As String
Public sgAsAirLogDate As String

Public igGridIgnoreScroll As Integer

Public igAlertFlash As Integer

Public lgPurgeCount As Long

Public tgMie As MIE
Public lgLastServiceDate As Long
Public lgLastServiceTime As Long
Public igCountTimeNotChanged As Integer

Public lgSchTopRow As Long
Public lgLibTopRow As Long
Public lgTempTopRow As Long

'Communication to/from Import Schedule into Libraries and Templates
Public sgExtractType As String  'L=Library; T=Template
Public sgExtractName As String  'Name plus subname
Public tgExtract() As SCHDEXTRACT
Public sgExtractBusNames() As String
Public sgExtractHours As String * 24
Public sgExtractDays As String * 7
Public sgExtractAudios() As String
Public sgExtractStartTime As String
Public sgExtractEndTime As String
Public lgExtractStartTime As Long
Public lgExtractEndTime As Long
Public sgExtractOffsets() As String
Public lgExtractOffsetStart() As Long
Public lgExtractOffsetEnd() As Long

'Ini Values
Public sgStartupDirectory As String
Public sgIniPathFileName As String
Public sgDatabaseName As String
Public sgReportDirectory As String
Public sgExportDirectory As String
Public sgImportDirectory As String
Public sgLogoDirectory As String
Public sgExeDirectory As String
Public igWaitCount As Integer
Public igTimeOut As Integer 'RDO Query Timeout (-1, use default, in sec)
Public sgMsgDirectory As String
Public sgDBPath As String
Public sgServerDatabase As String
Public igBkgdProg As Integer
Public sgDSN As String
Public lgCartUnloadTime As Long 'in seconds

Public hgDB As Integer

Public sgNowDate As String

Public igSQLSpec As Integer             '0=Pervasive 7; 1= Pervasive 2000 (default)
Public sgSQLDateForm As String          'Default: yyyy-mm-dd
Public sgSQLTimeForm As String          'Default: hh:mm:ss
Public sgShowDateForm As String         'Default m/d/yyyy
Public sgShowTimeWOSecForm As String    'Default h:mma/p
Public sgShowTimeWSecForm As String     'Default h:mm:ssa/p

''RDO variables
'Public env As rdoEnvironment
'Public cnn As rdoConnection
'Public rst As rdoResultset

Public cnn As ADODB.Connection
Public rst As ADODB.Recordset
Public rst2 As ADODB.Recordset


Public gErrSQL As ADODB.Error  'rdoError
Public gMsg As String

Public sgCommand As String
Public sgClientName As String
Public igTimes As Integer
Public igTestSystem As Integer
Public igPasswordOk As Integer

'SQL varaibles
Public sgSQLQuery As String

'Comm Port
Public sgItemIDDate As String
Public tgItemIDChk() As ITEMIDCHK

'File Images


'Automation Contact
Public tgACE As ACE
Public sgCurrACEStamp As String
Public tgCurrACE() As ACE

'Automation Data Flags
Public tgADE As ADE
Public sgCurrADEStamp As String
Public tgCurrADE() As ADE

'Automation Equipment
Public tgAEE As AEE
Public sgCurrAEEStamp As String
Public tgCurrAEE() As AEE

'Automation Features
Public tgStartColAFE As AFE
Public tgNoCharAFE As AFE
Public sgCurrAFEStamp As String
Public tgCurrAFE() As AFE

'Automation Paths
Public tgAPE As APE
Public sgCurrAPEStamp As String
Public tgCurrAPE() As APE

'Audio Names
Public tgANE As ANE
Public sgCurrANEStamp As String
Public tgCurrANE() As ANE
Public tgCurrANE_Name() As NAMESORT
Public sgBothANEStamp As String
Public tgBothANE() As ANE
Public tgUsedANE() As ANE

'Advertiser
Public tgARE As ARE
Public sgCurrAREStamp As String
Public tgCurrARE() As ARE

'Audio Source
Public tgASE As ASE
Public sgCurrASEStamp As String
Public tgCurrASE() As ASE
Public sgBothASEStamp As String
Public tgBothASE() As ASE

'Audio Type
Public tgATE As ATE
Public sgCurrATEStamp As String
Public tgCurrATE() As ATE
Public sgBothATEStamp As String
Public tgBothATE() As ATE
Public tgUsedATE() As ATE


'Bus Definition
Public tgBDE As BDE
Public sgCurrBDEStamp As String
Public tgCurrBDE() As BDE
Public tgCurrBDE_Name() As NAMESORT
Public sgBothBDEStamp As String
Public tgBothBDE() As BDE
Public tgUsedBDE() As BDE

'Bus Groups
Public tgBGE As BGE
Public sgCurrBGEStamp As String
Public tgCurrBGE() As BGE


'Bus Selection Groups
Public tgBSE As BSE
Public sgCurrBSEStamp As String
Public tgCurrBSE() As BSE

'Control Character
Public tgCCE As CCE
Public sgCurrCCEStamp As String
Public tgCurrCCE() As CCE
Public sgCurrAudioCCEStamp As String
Public tgCurrAudioCCE() As CCE
Public tgUsedAudioCCE() As CCE
Public sgCurrBusCCEStamp As String
Public tgCurrBusCCE() As CCE
Public tgUsedBusCCE() As CCE

'Comment and Title
Public tgCTE As CTE
Public sgCurrCTEStamp As String
Public tgCurrCTE() As CTE
Public tgCurr2CTE_Name() As CTESORT
Public tgCurr1CTE_Name() As DEECTE
Public tgUsedT2CTE() As CTE
Public sgCurr1CTEStamp As String
Public tgCurr1CTE() As CTE

'Day Events
Public tgDEE As DEE
Public sgCurrDEEStamp As String
Public tgCurrDEE() As DEE
Public lgLibDheUsed() As Long

'Day Header
Public tgDHE As DHE
Public sgCurrDHEStamp As String
Public tgCurrDHE() As DHE
Public sgCurrLibDHEStamp As String
Public tgCurrLibDHE() As DHE
Public sgBothLibDHEStamp As String
Public tgBothLibDHE() As DHE
Public sgCurrTempDHEStamp As String
Public tgCurrTempDHE() As DHE

'Day Names
Public tgDNE As DNE
Public sgCurrDNEStamp As String
Public tgCurrDNE() As DNE
Public sgCurrLibDNEStamp As String
Public tgCurrLibDNE() As DNE
Public sgCurrTempDNEStamp As String
Public tgCurrTempDNE() As DNE

'Day Sub-Names
Public tgDSE As DSE
Public sgCurrDSEStamp As String
Public tgCurrDSE() As DSE

'Day Event Buses
Public tgEBE As EBE
Public sgCurrEBEStamp As String
Public tgCurrEBE() As EBE

'Event Properties
Public tgEPE As EPE
Public sgCurrEPEStamp As String
Public tgCurrEPE() As EPE
Public tgUsedSumEPE As EPE
Public tgManSumEPE As EPE
Public tgSchUsedSumEPE As EPE
Public tgSchManSumEPE As EPE

'Event Type
Public tgETE As ETE
Public sgCurrETEStamp As String
Public tgCurrETE() As ETE
Public tgUsedETE() As ETE

'Follow
Public tgFNE As FNE
Public sgCurrFNEStamp As String
Public tgCurrFNE() As FNE
Public tgUsedFNE() As FNE

'Item Test
Public tgITE As ITE
Public sgCurrITEStamp As String
Public tgCurrITE() As ITE

'Material Type
Public tgMTE As MTE
Public sgCurrMTEStamp As String
Public tgCurrMTE() As MTE
Public tgUsedMTE() As MTE

'Netcue Names
Public tgNNE As NNE
Public sgCurrNNEStamp As String
Public tgCurrNNE() As NNE
Public tgCurrNNE_Name() As NAMESORT
Public tgUsedNNE() As NNE

'Relay
Public tgRNE As RNE
Public sgCurrRNEStamp As String
Public tgCurrRNE() As RNE
Public tgCurrRNE_Name() As NAMESORT
Public tgUsedRNE() As RNE

'Silence Character
Public tgSCE As SCE
Public sgCurrSCEStamp As String
Public tgCurrSCE() As SCE
Public tgUsedSCE() As SCE

'Schedule Events
'Day Events
Public tgSEE As SEE
Public sgCurrSEEStamp As String
Public tgCurrSEE() As SEE
Public tgSpotCurrSEE() As SEE

'Site_Gen_Schd
Public tgSGE As SGE
Public sgCurrSGEStamp As String
Public tgCurrSGE() As SGE

'Schedule Header
Public tgCurrDateSHE As SHE

'Site_Option
Public tgSOE As SOE
Public sgCurrSOEStamp As String
Public tgCurrSOE() As SOE

'Site_Path
Public tgSPE As SPE
Public sgCurrSPEStamp As String
Public tgCurrSPE() As SPE
Public tgCurrSSE As SSE

'Template Schedule
Public sgCurrTSEStamp As String
Public tgCurrTSE() As TSE
Public tgAirInfoTSE() As TSE

'Time Type
Public tgTTE As TTE
Public sgCurrTTEStamp As String
Public tgCurrTTE() As TTE
Public sgCurrStartTTEStamp As String
Public tgCurrStartTTE() As TTE
Public tgUsedStartTTE() As TTE
Public sgCurrEndTTEStamp As String
Public tgCurrEndTTE() As TTE
Public tgUsedEndTTE() As TTE

'Task_Names
Public sgCurrTNEStamp As String
Public tgCurrTNE() As TNE

'User_Info
Public tgUIE As UIE
Public sgCurrUIEStamp As String
Public tgCurrUIE() As UIE

'User_Tasks
Public tgUTE As UTE         'Sign On User record image
Public sgCurrUTEStamp As String
Public tgCurrUTE() As UTE

'Active_Info
Public tgAIE As AIE

Public tgYNMatchList() As MATCHLIST

Public tgT1MatchList() As MATCHLIST
Public tgT2MatchList() As MATCHLIST

Public sgBS As String  'Backspace
Public sgTB As String  'Tab
Public sgLF As String  'Line Feed (New Line)
Public sgCR As String  'Carriage Return
Public sgCRLF As String

Public fgPanelAdj As Single     'Adjustment to modeless window (without caption) enclosing panel

Public tgJobTaskNames() As TNE
Public tgListTaskNames() As TNE
Public tgExtraTaskNames() As TNE
Public tgAlertTaskNames() As TNE
Public tgNoticeTaskNames() As TNE

Public tgDDFFileNames() As DDFFILENAMES     'list of filenames from DDF for conversion of btrieve to odbc drivers for reporting

Public tgReportNames() As REPORTNAMES


'Filters
Public tgFilterValues() As FILTERVALUES
Public tgFilterFields() As FIELDSELECTION
Public igAnsFilter As Integer      'CALLDONE; CALLCANCELLED;

Public tgSchdReplaceValues() As SCHDREPLACEVALUES
Public tgLibReplaceValues() As LIBREPLACEVALUES
Public tgReplaceFields() As FIELDSELECTION
Public igAnsReplace As Integer      'CALLDONE; CALLCANCELLED;
Public igReplaceCallInfo As Integer    '0=Call form Schedule screen; 1=Call from Library definition screen; 2= Call from Temoplate screen; 3=call from Library Names screen
Public sgReplaceDefaultHours As String
Public bgApplyToEventType(0 To 2) As Boolean    '0=Program; 1=Avail; 2=Spot

Public sgFileAttachment As String           '2-16-05 Message feature
Public sgFileAttachmentName As String




Public Sub gInitVar()
    Dim ilLoop As Integer
    
    igExitEngrMain = False
    
    igGridIgnoreScroll = False
    
    sgEngrVersion = "Version 1.0" ' created 10/02/01 at 12:00PM"
    
    sgBS = Chr$(8)  'Backspace
    sgTB = Chr$(9)  'Tab
    sgLF = Chr$(10) 'Line Feed (New Line)
    sgCR = Chr$(13) 'Carriage Return
    sgCRLF = sgCR + sgLF
    
    fgPanelAdj = 90    'Adjustment to window around panel
    
End Sub

Public Sub gEraseVar()

End Sub
